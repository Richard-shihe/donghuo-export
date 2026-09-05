#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
阶段1 · IEC 质保书 PDF 下载 → 上传「原件」飞书文件夹
======================================================

流程：
  S0  IEC 登录（ibaosteel_client）+ 选组织
  S1  分页查询质保书列表（queryCertificate，交货期月份过滤 + 质保书日期倒序）
  S2  逐行 decryptByCertificateInfo → 新 RSA + customerId
  S3  Playwright 一体化下载（IEC SSO → ICSC Vue → 点下载图标 → save）
  S4  原件 PDF 上传到「待处理」文件夹（同名跳过，幂等）
  S5  机器人汇报（cert_common.notify）

交期可选（三选一，或交互式）：
  直接运行（不带交期参数）    本地终端弹出交互菜单：当天/近7天/当月/近3月/自定义
  --preset today    当天（接口按交货期月份过滤，即当月）
  --preset week     近 7 天
  --preset month    当月
  --preset month3   近 3 个月
  --days N          近 N 天（按月覆盖）
  --date-from YYYYMM --date-to YYYYMM   精确月份区间
  CI/管道等非交互环境不带参数时，自动按近 3 个月执行

用法：
  python stage1_certs_download.py                       # 近3个月，最多200条
  python stage1_certs_download.py --preset today        # 交期=当月
  python stage1_certs_download.py --days 7 --max 50     # 近7天，最多50条
  python stage1_certs_download.py --max 100 --workers 6  # 提高并发下载速度（默认4）
  python stage1_certs_download.py --date-from 202606 --date-to 202608
  python stage1_certs_download.py --dry-run             # 登录+查询+decrypt，不下载不上传

环境变量：
  IBAO_USERNAME / IBAO_PASSWORD        IEC 账密
  FEISHU_APP_ID   / FEISHU_APP_SECRET  飞书自建应用
  CERT_RAW_FOLDER_TOKEN                待处理文件夹（默认 YIrbf0NzlloFNKdYAjvcL6gXnhe）
  CERT_NOTIFY_UNION_IDS                汇报人员 union_id（逗号分隔；回退 FEISHU_UNION_IDS）
"""
from __future__ import annotations

import argparse
import asyncio
import datetime
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Optional, Tuple
from urllib.parse import unquote

import requests
from dotenv import load_dotenv
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

load_dotenv()  # 兼容从仓库根目录运行时读取 .env

_CERT_DIR = Path(__file__).resolve().parent
_REPO_ROOT = _CERT_DIR.parent
for _p in (str(_CERT_DIR), str(_REPO_ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from ibaosteel_client import IEC  # noqa: E402
from cert_common import (  # noqa: E402
    env, log, notify, tenant_access_token, list_folder_files,
    upload_to_folder, folder_url,
)

# ============================================================
# 常量
# ============================================================
BASE_HTTPS = "https://www.ibaosteel.com"
IECS = f"{BASE_HTTPS}/iecs"
# BSUrl 与页面 showRsaCertificate 保持一致
BSURL = "https://ecommerce.ibaosteel.com/icsc/TLfqmAction/tLfqmActionCasNewRedirectEncrypt?"
CERT_PAGE_URL = f"{IECS}/freight/productCertificate/productCertificate/initLoads"

# showRsaCertificate 9 参数固定值（渠道销售 · 钢贸三部）
FORM_FIELDS = {
    "uuCode": "U41634",
    "userNum": "062122",
    "saleNetWork": "E",
    "companyNo": "QE000000",
    "userNo": "QE000000",
    "system": "ES",
    "smartSegNo": "QE000000",
}

DEFAULT_RAW_FOLDER = "YIrbf0NzlloFNKdYAjvcL6gXnhe"   # 待处理（原件）
DEFAULT_PROCESSED_FOLDER = "HfYCfYMZnlhVpFdKI82cTHQ5npb"  # 已处理（stage2 原件归档，增量去重依据）

WORK_DIR = _CERT_DIR / "_cert_work"
WORK_DIR.mkdir(parents=True, exist_ok=True)   # CI 全新 checkout 无此目录，必须先建
PDF_DIR = WORK_DIR / "pdfs"
PDF_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# 日期
# ============================================================
def _parse_yymm(yyyymm: str) -> str:
    if not yyyymm:
        return ""
    s = re.sub(r"\D", "", yyyymm)
    if len(s) != 6:
        raise argparse.ArgumentTypeError(f"日期格式应为 YYYYMM（如 202607），实际={yyyymm}")
    return s


def month_range_from_days(days: int) -> Tuple[str, str]:
    """从近 N 天反推交货期月份起/止（YYYYMM）。"""
    today = datetime.date.today()
    start = today - datetime.timedelta(days=max(days - 1, 0))
    return f"{start.year:04d}{start.month:02d}", f"{today.year:04d}{today.month:02d}"


def resolve_date_range(args: argparse.Namespace, interactive: bool = True) -> Tuple[str, str]:
    """交期解析优先级：显式 --date-from/--date-to > --days N > --preset > 交互菜单 > 默认近3月。
    交互菜单仅在无任何交期参数且 stdin 为终端（本地运行）时出现；
    CI/管道等非交互环境自动落到默认近 3 个月。"""
    today = datetime.date.today()
    cur_ym = f"{today.year:04d}{today.month:02d}"
    if args.date_from or args.date_to:
        return args.date_from or cur_ym, args.date_to or cur_ym
    if args.days:
        return month_range_from_days(args.days)
    if args.preset:
        if args.preset == "today":      # 接口按交货期月份过滤，"当天"即当月
            return cur_ym, cur_ym
        if args.preset == "week":
            return month_range_from_days(7)
        if args.preset == "month":
            return cur_ym, cur_ym
        return month_range_from_days(90)  # month3
    if interactive:
        try:
            if sys.stdin.isatty():
                return prompt_date_range()
        except (EOFError, KeyboardInterrupt):
            print()
            log("  交互中断，使用默认近 3 个月")
    return month_range_from_days(90)


def prompt_date_range() -> Tuple[str, str]:
    """交互式选择交货期范围。"""
    print()
    print("请选择交货期范围：")
    print("  1. 当天（当月）")
    print("  2. 近 7 天")
    print("  3. 当月")
    print("  4. 近 3 个月（默认，直接回车）")
    print("  5. 自定义月份区间")
    print("  6. 自定义天数")
    choice = input("输入序号 [4]: ").strip()
    if choice in ("", "4"):
        return month_range_from_days(90)
    if choice == "1" or choice == "3":
        today = datetime.date.today()
        ym = f"{today.year:04d}{today.month:02d}"
        return ym, ym
    if choice == "2":
        return month_range_from_days(7)
    if choice == "5":
        lo = _parse_yymm(input("  交期起 YYYYMM（如 202606）: ").strip())
        hi = _parse_yymm(input("  交期止 YYYYMM（直接回车=当月）: ").strip())
        if not hi:
            today = datetime.date.today()
            hi = f"{today.year:04d}{today.month:02d}"
        return lo, hi
    if choice == "6":
        n = int(input("  近多少天（如 45）: ").strip() or "90")
        return month_range_from_days(n)
    raise ValueError(f"无效选择: {choice!r}")


# ============================================================
# S0: 登录
# ============================================================
def build_session() -> requests.Session:
    s = requests.Session()
    retry = Retry(total=3, backoff_factor=0.5, status_forcelist=[429, 500, 502, 503, 504],
                  allowed_methods=["GET", "POST"])
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=40)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    s.trust_env = False
    try:
        from requests.packages.urllib3.exceptions import InsecureRequestWarning
        import urllib3
        urllib3.disable_warnings(InsecureRequestWarning)
    except Exception:
        pass
    return s


def login_and_session() -> Tuple[IEC, requests.Session, str]:
    iec = IEC(username=env("IBAO_USERNAME"), password=env("IBAO_PASSWORD"), retries=5)
    iec.login()
    s = iec.session
    token = iec.token
    # 选组织（渠道销售），与页面人工流程一致
    s.get(f"{BASE_HTTPS}/ibaosteel/bizIntelli?access_token={token}", timeout=20)
    s.get(f"{IECS}/index?token={token}", timeout=20, allow_redirects=True)
    return iec, s, token


# ============================================================
# S1: 分页查询质保书列表
# ============================================================
def query_certificate_page(s: requests.Session, page_num: int, page_size: int,
                           from_ym: str, to_ym: str) -> list[dict]:
    body = {
        "delivaryDateFrom": from_ym, "delivaryDateTo": to_ym,
        "segNo": "QE000000", "system": "ES", "userNum": "062122",
        "uuCode": "U41634", "userNo": "QE000000", "companyNo": "QE000000", "saleNetwork": "E",
        "contractNum": "", "factoryOrderNum": "",
        "packNum": "", "tcNumTc": "",
        "inDateFrom": "", "inDateTo": "", "sortField": "inDate", "sortRule": "desc",
        "pageDomain": {"pageNum": str(page_num), "pageSize": str(page_size)},
    }
    r = s.post(
        f"{IECS}/freight/productCertificate/productCertificate/queryCertificate",
        json=body,
        headers={
            "Accept": "text/html, */*; q=0.01",
            "Content-Type": "application/json; charset=utf-8",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": CERT_PAGE_URL,
        },
        timeout=30,
    )
    html = r.text
    rows = []
    for m in re.finditer(r'<input[^>]*name="ck_box"[^>]*>', html, re.I):
        tag = m.group(0)
        def _a(n):
            am = re.search(r"\b" + n + r'="([^"]*)"', tag, re.I)
            return unquote(am.group(1)) if am else ""
        rows.append({
            "tcNumTc": _a("tcNumTc"),
            "tcNumTcRsa": _a("tcNumTcRsa"),
            "factoryOrderNumRsa": _a("factoryOrderNumRsa"),
            "contractNum": _a("contractNum"),
            "boardPlank": _a("boardPlank") or " ",
        })
    return rows


def stage_collect_rows(s: requests.Session, *, from_ym: str, to_ym: str,
                       page_size: int = 10, max_rows: int = 30,
                       log_file: Optional[Path] = None) -> list[dict]:
    """分页查询 checkbox，按 tcNumTc 去重取前 max_rows 条。"""
    all_rows: list[dict] = []
    seen: set[str] = set()
    for page_num in range(1, 100):
        page_rows = query_certificate_page(s, page_num, page_size, from_ym, to_ym)
        log(f"  分页 {page_num}: {len(page_rows)} 条 checkbox", log_file)
        if not page_rows:
            break
        for r in page_rows:
            if r["tcNumTc"] and r["tcNumTc"] not in seen:
                seen.add(r["tcNumTc"])
                all_rows.append(r)
                if len(all_rows) >= max_rows:
                    break
        if len(all_rows) >= max_rows or len(page_rows) < page_size:
            break
    log(f"  合计（去重）{len(all_rows)} 条，取前 {max_rows} 条", log_file)
    return all_rows


# ============================================================
# S2: 单行 decrypt → 新 RSA + customerId
# ============================================================
def stage_decrypt_one(s: requests.Session, row: dict) -> Tuple[str, str]:
    body = [{
        "tcNumTc": row["tcNumTcRsa"],
        "boardPlank": row["boardPlank"],
        "factoryOrderNum": row["factoryOrderNumRsa"],
        "customerId": "",
    }]
    r = s.post(
        f"{IECS}/freight/productCertificate/productCertificate/decryptByCertificateInfo",
        json=body,
        headers={
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Content-Type": "application/json; charset=utf-8",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": CERT_PAGE_URL,
        },
        timeout=30,
    )
    d = r.json()
    if str(d.get("code", "")) != "0":
        raise RuntimeError(f"decrypt failed code={d.get('code')} msg={d.get('msg','')}")
    data = d.get("data") or {}
    if isinstance(data, list):
        data = data[0] if data else {}
    new_rsa = str(data.get("tcNumTcRsa") or "")
    if not new_rsa:
        raise RuntimeError("decrypt 返回无 tcNumTcRsa")
    return new_rsa, str(data.get("customerId") or "")


# ============================================================
# S3: Playwright 一体化下载
# ============================================================
def _is_pdf(b: bytes) -> bool: return b[:4] == b"%PDF"
def _is_zip(b: bytes) -> bool: return b[:2] == b"PK"

WAIT_VUE_JS = r"""() => {
    function fv(el){if(!el)return null;for(const k of Object.keys(el))if(k.startsWith('__vue')||k.startsWith('__VUE'))return el[k];return null;}
    function w(r,c,d){if(!r||d>40)return;c(r);if(r.$children)for(const x of r.$children)w(x,c,d+1);if(r.$refs)for(const k in r.$refs)w(r.$refs[k],c,d+1);}
    const v0=fv(document.getElementById('app')||document.body);
    let rows=0;
    if(v0){try{w(v0,vm=>{try{const td=vm.tableData||(vm.$data&&vm.$data.tableData);if(Array.isArray(td)&&td.length>rows)rows=td.length;}catch(e){}})}catch(e){}}
    return {hasVue:!!v0,rows,title:document.title,url:location.href};
}"""


# 页面表格第一行点击下载图标（与页面人工点击行为一致）
CLICK_DOWNLOAD_JS = r'''() => {
    try {
        const rows = document.querySelectorAll(".vxe-table--body-wrapper tbody tr, .el-table__body-wrapper tbody tr, table tbody tr");
        if (!rows.length) return "no_tbody";
        const first = rows[0];
        const cands = first.querySelectorAll(
            ".iconxiazai, .icondayin, .iconpdf, .iconfont.xiazai, .iconfont.dayin,"
            +"i[class*=xiazai], i[class*=dayin], i[class*=down], i[class*=view], i[class*=pdf],"
            +"[title*=下载], [title*=预览], [title*=打印],"
            +".view-pdf, .download-pdf"
        );
        if (cands.length) { cands[0].click(); return "icon_clicked:"+cands.length; }
        const btns = first.querySelectorAll("button, a, i.item, span.item");
        if (btns.length) { btns[btns.length-1].click(); return "lastBtn_clicked"; }
        return "no_candidates";
    } catch(e) { return "err:"+String(e.message||e).slice(0,120); }
}'''

SEED_TOKEN_JS = '''(token) => { try {
    localStorage.setItem("BAOSTEEL_AT", token);
    sessionStorage.setItem("BAOSTEEL_AT", token);
    localStorage.setItem("token", token);
} catch(e){} }'''


async def _dump_fail(page, tc: str) -> None:
    """失败现场：截图 + HTML 落盘（尽力而为）。"""
    try:
        png = WORK_DIR / f"_pwdl_fail_{tc}.png"
        await page.screenshot(path=str(png), full_page=True)
        (WORK_DIR / f"_pwdl_fail_{tc}.html").write_text((await page.content()) or "", encoding="utf-8")
    except Exception:
        pass


async def _pw_download_one(ctx, sem: asyncio.Semaphore, token: str,
                           gi: int, total: int, r: dict,
                           log_file: Optional[Path]) -> Optional[dict]:
    """单条：种 token → POST 跳转 → 等 Vue rows≥1 → 点下载 → save PDF。"""
    async with sem:
        tc = r["tcNumTc"]
        new_rsa = r["_newTcNumTcRsa"]
        cust = r.get("_orderOwner") or ""
        form = {**FORM_FIELDS, "tcNumTc": new_rsa, "orderOwner": cust}
        jump_url = f"{BSURL}access_token={token}"

        page = await ctx.new_page()
        try:
            await page.evaluate(SEED_TOKEN_JS, token)
            try:
                resp = await page.request.post(jump_url, form=form, timeout=60000, max_redirects=10)
                target_url = resp.url
                if target_url.startswith("http://ecommerce"):
                    target_url = "https://" + target_url[len("http://"):]
                await page.goto(target_url, wait_until='domcontentloaded', timeout=60000)
            except Exception as e:
                log(f"  [{gi}/{total}] navigate 异常: {e}", log_file)

            # 等 Vue 初始化 + tableData.rows=1（300ms 快轮询，总预算 ~15s）
            info_vue = {}
            for _ in range(50):
                try:
                    info_vue = await page.evaluate(WAIT_VUE_JS)
                except Exception as e:
                    info_vue = {"err": str(e)}
                if info_vue.get("rows", 0) > 0:
                    break
                await page.wait_for_timeout(300)
            log(
                f"  [{gi}/{total}] tc={tc} Vue={info_vue.get('hasVue')} rows={info_vue.get('rows')}",
                log_file,
            )
            if info_vue.get("rows", 0) < 1:
                await _dump_fail(page, tc)
                return None

            # 监听 response.body 兜底
            holder = {"pdf_bytes": None}

            async def _on_response(response):
                if holder["pdf_bytes"]:
                    return
                try:
                    ct = (response.headers.get("content-type") or "").lower()
                    url = response.url.lower()
                    if ('pdf' in ct or 'octet-stream' in ct or 'download' in ct
                            or 'muldownload' in url or 'vieweplatpdf' in url
                            or 'views3pdf' in url or 'downloadsignpdf' in url):
                        body = await response.body()
                        if len(body) > 1024 and (_is_pdf(body) or _is_zip(body)):
                            holder["pdf_bytes"] = body
                except Exception:
                    pass
            page.on('response', _on_response)

            # 点下载图标，优先 expect_download，失败用 response.body 兜底
            saved_path: Optional[Path] = None
            try:
                async with page.expect_download(timeout=20000) as dl_info:
                    await page.evaluate(CLICK_DOWNLOAD_JS)
                dl = await dl_info.value
                ext = '.zip' if (dl.suggested_filename or '').lower().endswith('.zip') else '.pdf'
                saved_path = PDF_DIR / f'{tc}{ext}'
                await dl.save_as(str(saved_path))
            except Exception:
                wait_until = time.time() + 20
                while time.time() < wait_until and holder["pdf_bytes"] is None:
                    await page.wait_for_timeout(500)
                if holder["pdf_bytes"]:
                    body = holder["pdf_bytes"]
                    ext = '.pdf' if _is_pdf(body) else '.zip'
                    saved_path = PDF_DIR / f'{tc}{ext}'
                    saved_path.write_bytes(body)

            if saved_path and saved_path.exists() and saved_path.stat().st_size > 1024:
                size = saved_path.stat().st_size
                log(f"  [{gi}/{total}] ✅ {tc} → {saved_path.name} ({size:,} bytes)", log_file)
                return {
                    "tc": tc,
                    "contractNum": r.get("contractNum") or "",
                    "raw_pdf": str(saved_path.absolute()),
                    "size": size,
                }
            log(f"  [{gi}/{total}] ❌ 下载失败 tc={tc}", log_file)
            await _dump_fail(page, tc)
            return None
        except Exception as e:
            log(f"  [{gi}/{total}] ❌ 下载异常 tc={tc}: {e}", log_file)
            try:
                await _dump_fail(page, tc)
            except Exception:
                pass
            return None
        finally:
            try:
                await page.close()
            except Exception:
                pass


def stage_pw_download_all_in_one(token: str, decrypted_rows: list[dict],
                                 *, log_file: Optional[Path] = None,
                                 workers: int = 4) -> list[dict]:
    """并发下载：同一浏览器 context 内开 N 个页面并行执行 SSO → Vue → 点下载。
    各页面共享 cookie/localStorage 登录态，总耗时约为串行版的 1/N。"""
    return asyncio.run(_pw_download_all(token, decrypted_rows,
                                        log_file=log_file, workers=max(1, workers)))


async def _pw_download_all(token: str, decrypted_rows: list[dict],
                           *, log_file: Optional[Path] = None,
                           workers: int = 4) -> list[dict]:
    from playwright.async_api import async_playwright

    results: list[dict] = []
    total = len(decrypted_rows)
    sem = asyncio.Semaphore(workers)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, timeout=60000)
        ctx = await browser.new_context(viewport={'width': 1920, 'height': 1080},
                                        accept_downloads=True, ignore_https_errors=True)
        ctx.set_default_timeout(60000)

        # Warmup：种 BAOSTEEL_AT + 过 SSO 跳转（一次性，后续所有页面共享 context 登录态）
        warmup = await ctx.new_page()
        await warmup.evaluate(SEED_TOKEN_JS, token)
        try:
            await warmup.goto(f"{BASE_HTTPS}/ibaosteel/bizIntelli?access_token={token}",
                              wait_until='domcontentloaded', timeout=60000)
        except Exception as e:
            log(f"  warmup bizIntelli 警告: {e}", log_file)
        await warmup.wait_for_timeout(3000)
        try:
            await warmup.goto(f"{IECS}/index?token={token}", wait_until='domcontentloaded',
                              timeout=60000)
        except Exception as e:
            log(f"  warmup iecs/index 警告: {e}", log_file)
        await warmup.wait_for_timeout(5000)
        try:
            await warmup.close()
        except Exception:
            pass

        async def _run(gi: int, r: dict):
            res = await _pw_download_one(ctx, sem, token, gi, total, r, log_file)
            if res:
                results.append(res)

        await asyncio.gather(*(_run(i, r) for i, r in enumerate(decrypted_rows, 1)))

        try:
            await browser.close()
        except Exception:
            pass
    return results


# ============================================================
# S4: 原件上传（同名跳过）
# ============================================================
def stage_upload_raw(token: str, raw_folder: str, downloaded: list[dict],
                     *, log_file: Optional[Path] = None) -> list[dict]:
    existing = set()
    try:
        existing = {f.get("name") or "" for f in list_folder_files(token, raw_folder)}
    except Exception as e:
        log(f"  ⚠️ 列待处理文件夹失败（不去重直接传）: {e}", log_file)
    oks = skipped = failed = 0
    for i, d in enumerate(downloaded, 1):
        fp = Path(d["raw_pdf"])
        name = fp.name
        if name in existing:
            skipped += 1
            d["upload_skipped"] = True
            log(f"  [{i}/{len(downloaded)}] ⏭ 文件夹已有同名，跳过: {name}", log_file)
            continue
        try:
            ft = upload_to_folder(token, raw_folder, name, fp.read_bytes())
            d["feishu_file_token"] = ft
            d["feishu_folder"] = raw_folder
            oks += 1
            log(f"  [{i}/{len(downloaded)}] ✅ 上传原件 {name} → {ft}", log_file)
        except Exception as e:
            failed += 1
            log(f"  [{i}/{len(downloaded)}] ❌ 上传失败 {name}: {e}", log_file)
        time.sleep(0.1)
    log(f"S4 完成：上传 {oks}，同名跳过 {skipped}，失败 {failed}", log_file)
    return downloaded


# ============================================================
# Main
# ============================================================
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="阶段1 · IEC 质保书 PDF 下载 → 上传「原件」飞书文件夹",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--preset", choices=["today", "week", "month", "month3"], default=None,
                   help="交期快捷档：today 当天(当月) / week 近7天 / month 当月 / month3 近3个月；"
                        "不传任何交期参数时，本地运行弹出交互菜单（CI 默认近3个月）")
    p.add_argument("--days", type=int, default=0,
                   help="交期近 N 天（按月覆盖），优先级高于 --preset")
    p.add_argument("--date-from", type=_parse_yymm, default="",
                   help="交期起 YYYYMM（优先级最高）")
    p.add_argument("--date-to", type=_parse_yymm, default="",
                   help="交期止 YYYYMM（默认当月）")
    p.add_argument("--max", type=int, default=200,
                   help="最多处理多少条质保书（按 tcNumTc 去重，默认 200）")
    p.add_argument("--page-size", type=int, default=10,
                   help="queryCertificate 每页大小（默认 10）")
    p.add_argument("--workers", type=int, default=4,
                   help="并发下载页面数（同一浏览器内并行，默认 4；服务端限流时可调小）")
    p.add_argument("--raw-folder", default=env("CERT_RAW_FOLDER_TOKEN", DEFAULT_RAW_FOLDER),
                   help=f"原件（待处理）文件夹 token（默认 {DEFAULT_RAW_FOLDER}）")
    p.add_argument("--processed-folder", default=env("CERT_PROCESSED_FOLDER_TOKEN", DEFAULT_PROCESSED_FOLDER),
                   help=f"已处理文件夹 token（stage2 原件归档，增量去重依据，默认 {DEFAULT_PROCESSED_FOLDER}）")
    p.add_argument("--no-incremental", action="store_true",
                   help="关闭增量去重（强制重新下载全部，默认按质保书号跳过已下载的）")
    p.add_argument("--dry-run", action="store_true",
                   help="登录+查询+decrypt，不下载不上传不汇报")
    p.add_argument("--no-notify", action="store_true", help="跳过机器人汇报")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    t0 = time.time()
    log_file = WORK_DIR / f'stage1_{time.strftime("%Y%m%d_%H%M%S")}.log'
    date_from, date_to = resolve_date_range(args)

    log("========== 【阶段1】IEC 质保书下载 → 原件文件夹 ==========", log_file)
    log(f"  交期范围: {date_from} ~ {date_to} (YYYYMM, preset={args.preset})", log_file)
    log(f"  最多条数: {args.max}   每页: {args.page_size}   并发: {args.workers}", log_file)
    log(f"  待处理文件夹: {args.raw_folder}", log_file)
    log(f"  dry_run: {args.dry_run}", log_file)

    # S0 登录
    log("\n[S0] IEC 登录 + 选组织", log_file)
    try:
        iec, session, token = login_and_session()
    except Exception as e:
        log(f"  ❌ 登录失败: {e}", log_file)
        return 1
    log(f"  ✅ 登录成功 token[:24]={token[:24]}…", log_file)

    # S1 列表
    log("\n[S1] 分页查询质保书列表（queryCertificate）", log_file)
    rows = stage_collect_rows(session, from_ym=date_from, to_ym=date_to,
                              page_size=args.page_size, max_rows=args.max, log_file=log_file)
    if not rows:
        log("  ❌ 0 条，结束", log_file)
        return 1

    # S2 decrypt（线程池并发）
    log("\n[S2] 并发 decrypt → 新 RSA + customerId", log_file)

    def _decrypt_one(r: dict) -> Tuple[dict, str, str, Optional[Exception]]:
        try:
            new_rsa, cust = stage_decrypt_one(session, r)
            return r, new_rsa, cust, None
        except Exception as e:
            return r, "", "", e

    decrypted: list[dict] = []
    with ThreadPoolExecutor(max_workers=min(6, max(1, len(rows)))) as ex:
        for i, (r, new_rsa, cust, err) in enumerate(ex.map(_decrypt_one, rows), 1):
            if err is not None:
                log(f"  [{i}/{len(rows)}] decrypt 失败 tc={r['tcNumTc']}: {err}", log_file)
            else:
                r["_newTcNumTcRsa"] = new_rsa
                r["_orderOwner"] = cust
                decrypted.append(r)
    log(f"  ✅ decrypt 通过 {len(decrypted)}/{len(rows)}", log_file)
    if not decrypted:
        log("  ❌ decrypt 全失败，结束", log_file)
        return 2

    # S2.5 增量去重：已在「待处理/已处理」文件夹的质保书（按质保书号文件名）不再下载
    # 说明：stage2 跑完后待处理会清空、原件进「已处理」，所以两个文件夹都要查，
    # 否则每天都会把交期窗口内的同一批质保书重复下载一遍。
    n_exist = 0
    n_decrypted_total = len(decrypted)
    if args.no_incremental:
        log("\n[S2.5] 增量去重已关闭（--no-incremental），将下载全部", log_file)
    else:
        log("\n[S2.5] 增量去重：扫描 待处理 + 已处理 文件夹", log_file)
        already_names: set[str] = set()
        try:
            ftok = tenant_access_token()
            for fld, label in ((args.raw_folder, "待处理"),
                               (args.processed_folder, "已处理")):
                if not fld:
                    continue
                try:
                    names = {f.get("name") or "" for f in list_folder_files(ftok, fld)}
                    log(f"  {label}文件夹已有 {len(names)} 个文件", log_file)
                    already_names |= names
                except Exception as e:
                    log(f"  ⚠️ 列{label}文件夹失败（该文件夹不参与去重）: {e}", log_file)
            if already_names:
                kept = [r for r in decrypted
                        if f'{r["tcNumTc"]}.pdf' not in already_names
                        and f'{r["tcNumTc"]}.zip' not in already_names]
                n_exist = len(decrypted) - len(kept)
                decrypted = kept
                log(f"  增量结果：{n_exist} 条已下载过 → 跳过；本次新增 {len(decrypted)} 条", log_file)
        except Exception as e:
            log(f"  ⚠️ 取飞书 token 失败，本次不做增量去重: {e}", log_file)

    if args.dry_run:
        log(f"\n[dry-run] 停在下载前：列表 {len(rows)} | decrypt {n_decrypted_total} | "
            f"新增待下载 {len(decrypted)} | 已存在跳过 {n_exist}", log_file)
        return 0

    if not decrypted:
        elapsed = int(time.time() - t0)
        mins, secs = divmod(elapsed, 60)
        log("  ✅ 无新增质保书（窗口内全部已下载过），结束", log_file)
        msg = (f"【质保书·①IEC下载】完成 ✅\n"
               f"交期范围: {date_from} ~ {date_to}\n"
               f"列表 {len(rows)} 条 | decrypt {n_decrypted_total}\n"
               f"新增下载: 0（{n_exist} 条均已下载过，跳过）\n"
               f"耗时: {mins}分{secs}秒")
        print(msg)
        if not args.no_notify:
            notify(msg)
        return 0

    # S3 Playwright 一体化下载（并发）
    log(f"\n[S3] Playwright 并发下载 PDF（workers={args.workers}）", log_file)
    downloaded = stage_pw_download_all_in_one(token, decrypted, log_file=log_file,
                                              workers=args.workers)
    if not downloaded:
        log("  ❌ 0 份下载成功", log_file)
        return 4

    # S4 原件上传
    log("\n[S4] 原件上传 → 待处理文件夹", log_file)
    try:
        ftok = tenant_access_token()
    except Exception as e:
        log(f"  ❌ 取飞书 token 失败: {e}", log_file)
        return 3
    downloaded = stage_upload_raw(ftok, args.raw_folder, downloaded, log_file=log_file)

    # S5 汇报
    elapsed = int(time.time() - t0)
    mins, secs = divmod(elapsed, 60)
    n_up = sum(1 for d in downloaded if d.get("feishu_file_token"))
    n_skip = sum(1 for d in downloaded if d.get("upload_skipped"))
    msg = (f"【质保书·①IEC下载】完成 ✅\n"
           f"交期范围: {date_from} ~ {date_to}\n"
           f"列表 {len(rows)} 条 | decrypt {n_decrypted_total} | 新增 {len(downloaded)}"
           f"（已存在跳过 {n_exist}）\n"
           f"上传原件: {n_up}" + (f"（同名跳过 {n_skip}）" if n_skip else "") + "\n"
           f"文件夹: {folder_url(args.raw_folder)}\n"
           f"耗时: {mins}分{secs}秒")
    print(msg)
    if not args.no_notify:
        notify(msg)

    result_path = WORK_DIR / f'stage1_result_{time.strftime("%Y%m%d_%H%M%S")}.json'
    with open(result_path, "w", encoding="utf-8") as f:
        json.dump({
            "args": {**vars(args), "date_from": date_from, "date_to": date_to},
            "elapsed_sec": elapsed,
            "rows": len(rows),
            "decrypted": len(decrypted),
            "downloaded": len(downloaded),
            "uploaded": n_up,
            "upload_skipped": n_skip,
            "raw_folder": args.raw_folder,
            "items": downloaded,
        }, f, ensure_ascii=False, indent=2)
    log(f"📄 结果快照: {result_path}", log_file)
    return 0


if __name__ == "__main__":
    sys.exit(main())
