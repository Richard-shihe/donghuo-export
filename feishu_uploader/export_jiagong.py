#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火钢城系统 - 加工单导出 → 飞书多维表格 / 本地 CSV

数据来源：加工管理 > 加工单 > 详情页 (v_jigondan)
流程：
  1. 自动登录 https://erpa.donghuo.vip（ddddocr 识别验证码）
  2. 调 getlist 获取加工单号列表
  3. 逐个访问详情页 HTML，解析加工单级字段 + 明细行
  4. 每行明细合并加工单级字段 = 一条记录
  5. 按"明细行ID"去重，追加到飞书多维表格

环境变量（必填）：
  DH_USERNAME           erpa 登录账号
  DH_PASSWORD           erpa 登录密码

环境变量（交付方式，默认 bitable）：
  DELIVERY_MODE         bitable(默认) / local

环境变量（飞书多维表格交付，DELIVERY_MODE=bitable 时必填）：
  FEISHU_APP_ID         飞书自建应用 App ID
  FEISHU_APP_SECRET     飞书自建应用 App Secret
  BITABLE_APP_TOKEN     多维表格 app_token
  BITABLE_TABLE_ID      目标 table id
（可选）
  BITABLE_DEDUP         按"明细行ID"去重: 1(默认)=跳过重复 0=即使重复也追加

环境变量（可选筛选）：
  FILTER_ZHUANTAI       状态筛选：待加工完成 / 加工完成 / 空(全部)
  EXPORT_DAYS           导出最近 N 天，0=全部
  JIAGONG_MAX           最多导出多少个加工单（调试用，0=全部）
"""

import os
import io
import csv
import re
import html as html_mod
import time
import json
import datetime
import traceback

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_URL = "https://erpa.donghuo.vip"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"

# 加工单列表 API
JIAGONG_LIST_API = f"{BASE_URL}/model/admin/caigou/m_jiagon/getlist"
# 加工单详情页
JIAGONG_DETAIL_PAGE = f"{BASE_URL}/view/admin/jiagon/v_jigondan"
# 加工单框架页（预热用）
JIAGONG_FRAME_PAGE = f"{BASE_URL}/view/admin/jiagon/v_ifram_jg"
JIAGONG_INDEX_PAGE = f"{BASE_URL}/view/admin/jiagon/v_index"


# ---------------- 通用 ----------------

def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def create_session() -> requests.Session:
    s = requests.Session()
    retry = Retry(total=5, backoff_factor=1,
                  status_forcelist=[500, 502, 503, 504, 429],
                  allowed_methods=["POST", "GET"])
    adapter = HTTPAdapter(max_retries=retry)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    s.headers.update({
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Accept-Language": "zh-CN,zh;q=0.9",
    })
    return s


# ---------------- 登录 ----------------

def recognize_captcha(image_bytes: bytes) -> str:
    try:
        import ddddocr
        return ddddocr.DdddOcr(show_ad=False).classification(image_bytes).strip().replace(" ", "")
    except ImportError:
        print("[警告] 未安装 ddddocr")
        return ""
    except Exception as exc:
        print(f"[错误] 验证码识别异常: {exc}")
        return ""


def login(session: requests.Session, username: str, password: str,
          max_attempts: int = 10) -> bool:
    login_url = f"{BASE_URL}/controller/admin/c_longin/index"
    captcha_url = f"{BASE_URL}/common/captcha"

    for attempt in range(1, max_attempts + 1):
        print(f"[登录] 尝试 {attempt}/{max_attempts} ...")
        img_resp = session.get(captcha_url, timeout=15)
        if img_resp.status_code != 200:
            time.sleep(2)
            continue

        captcha_code = recognize_captcha(img_resp.content)
        if not captcha_code:
            captcha_code = "1234"
        print(f"  识别结果: {captcha_code}")

        resp = session.post(login_url,
                            data={"u_name": username, "u_pass": password,
                                  "captcha": captcha_code},
                            timeout=15)
        text = resp.text.strip()
        try:
            result = json.loads(text)
            if isinstance(result, dict) and str(result.get("code")) == "200":
                print("[登录] 成功")
                return True
            msg = result.get("msg") or result.get("message") or text[:100]
            print(f"  失败: {msg}")
        except json.JSONDecodeError:
            print(f"  非JSON响应(前200): {text[:200]}")
        time.sleep(2)

    print(f"[登录] 已达最大尝试次数 {max_attempts}，登录失败")
    return False


# ---------------- 加工单列表 ----------------

def get_jiagong_danhao_list(session: requests.Session,
                            zhuantai: str = "",
                            max_count: int = 0) -> list[str]:
    """调 getlist 获取加工单号列表（仅返回加工单号）"""
    # 预热
    try:
        session.get(JIAGONG_FRAME_PAGE, timeout=15)
        session.get(JIAGONG_INDEX_PAGE, timeout=15)
    except Exception:
        pass

    all_danhaos: list[str] = []
    page = 1
    page_size = 50
    total_pages = 1

    while page <= total_pages:
        data = {"page": page, "limit": page_size}
        if zhuantai:
            data["zhuantai"] = zhuantai

        resp = session.post(JIAGONG_LIST_API, data=data, timeout=30,
                            headers={"X-Requested-With": "XMLHttpRequest"})
        try:
            parsed = json.loads(resp.text)
        except json.JSONDecodeError:
            print(f"[错误] 第 {page} 页解析失败: {resp.text[:200]}")
            break

        if page == 1:
            total_pages = int(parsed.get("pgtotal") or 1)
            total_count = int(parsed.get("rtotal") or 0)
            print(f"[加工单列表] 共 {total_count} 条, {total_pages} 页")

        rows = parsed.get("root") or []
        for r in rows:
            danhao = str(r.get("加工单号") or "").strip()
            if danhao:
                all_danhaos.append(danhao)

        print(f"  第 {page}/{total_pages} 页: {len(rows)} 条, 累计 {len(all_danhaos)} 个加工单号")

        if max_count > 0 and len(all_danhaos) >= max_count:
            all_danhaos = all_danhaos[:max_count]
            print(f"  已达到 JIAGONG_MAX={max_count} 限制，停止")
            break

        page += 1
        time.sleep(0.3)

    return all_danhaos


# ---------------- 详情页 HTML 解析 ----------------

def parse_jigondan_html(html_text: str, danhao: str = "") -> list[dict]:
    """
    从详情页 HTML 解析出记录列表。
    每行明细 + 加工单级字段 = 一条记录。
    返回 list[dict]，key 为飞书表字段名。
    """
    # === 加工单级字段 ===
    jg_danhao = danhao
    if not jg_danhao:
        m = re.search(r'加工单号[：:]\s*([A-Z]\d{4}-\d+)', html_text)
        if m:
            jg_danhao = m.group(1)

    # input 字段
    def get_input_value(field_id: str) -> str:
        m = re.search(rf'<input[^>]*?id=["\']{field_id}["\'][^>]*?value=["\']([^"\']*)["\']',
                      html_text)
        return m.group(1).strip() if m else ""

    jiner = get_input_value("jiner")
    beizhu = get_input_value("beizhu")

    # select 字段
    def get_select_value(sel_id: str) -> str:
        m = re.search(rf'<select[^>]*?id=["\']{sel_id}["\'][^>]*?>(.*?)</select>',
                      html_text, re.S)
        if not m:
            return ""
        sel_html = m.group(1)
        sel_m = re.search(
            r'<option[^>]*?value=["\']([^"\']*)["\'][^>]*?selected[^>]*?>([^<]*)</option>',
            sel_html, re.S)
        if not sel_m:
            sel_m = re.search(
                r'<option[^>]*?selected[^>]*?value=["\']([^"\']*)["\'][^>]*?>([^<]*)</option>',
                sel_html, re.S)
        if sel_m:
            val = sel_m.group(1).strip()
            txt = sel_m.group(2).strip()
            if txt and txt != "请选择" and val:
                return val
            if val and val != "0":
                return val
        return ""

    fyname = get_select_value("fyname")
    funame = get_select_value("funame")
    fcompany = get_select_value("fcompany")
    jsdanwei = get_select_value("jsdanwei")
    shuilv = get_select_value("shuilv")

    # 加工备注
    jg_beizhu = ""
    m = re.search(r'id=["\']jg_beizhu["\'][^>]*?>(.*?)</div>', html_text, re.S)
    if not m:
        m = re.search(r'id=["\']jg_beizhu["\'][^>]*?>(.*?)</textarea>', html_text, re.S)
    if m:
        jg_beizhu = re.sub(r'<[^>]+>', '', m.group(1)).strip()

    # === 明细行 ===
    # 直接搜索含 id="zid" 的 tr（不依赖 table 嵌套匹配）
    records: list[dict] = []
    for row_m in re.finditer(r'<tr[^>]*>(.*?)</tr>', html_text, re.S):
        row_html = row_m.group(1)
        zid_m = re.search(
            r'<input[^>]*?id=["\']zid["\'][^>]*?value=["\']([^"\']*)["\']',
            row_html)
        if not zid_m:
            continue

        zid = zid_m.group(1).strip()
        # 解析 td 文本
        tds = re.findall(r'<td[^>]*>(.*?)</td>', row_html, re.S)
        td_texts: list[str] = []
        for td in tds:
            inp_m = re.search(r'<input[^>]*?value=["\']([^"\']*)["\']', td)
            if inp_m:
                td_texts.append(inp_m.group(1).strip())
            else:
                clean = re.sub(r'<[^>]+>', '', td).strip()
                clean = html_mod.unescape(clean).replace('&nbsp;', '').strip()
                td_texts.append(clean)

        # 结构：[序号, 捆包号, 品名, 规格, 产地, 材质, 件(张)数, 重量(吨), 库位号, 操作]
        fields_order = ["捆包号", "品名", "规格", "产地", "材质",
                        "件(张)数", "重量(吨)", "库位号"]
        detail: dict[str, str] = {
            "加工单号": jg_danhao,
            "服务商名称": fyname,
            "服务商联系人": funame,
            "服务商公司": fcompany,
            "结算单位": jsdanwei,
            "税率": shuilv,
            "金额": jiner,
            "备注": beizhu,
            "加工备注": jg_beizhu,
            "明细行ID": zid,
            "入库单价": "",  # 原料表格中无此列
        }
        for i, fname in enumerate(fields_order):
            idx = i + 1  # 跳过序号列
            detail[fname] = td_texts[idx] if idx < len(td_texts) else ""
        records.append(detail)

    # 如果没有明细行，仍返回一条加工单级记录
    if not records:
        records.append({
            "加工单号": jg_danhao,
            "服务商名称": fyname, "服务商联系人": funame,
            "服务商公司": fcompany, "结算单位": jsdanwei,
            "税率": shuilv, "金额": jiner, "备注": beizhu,
            "加工备注": jg_beizhu, "明细行ID": "",
            "捆包号": "", "品名": "", "规格": "", "产地": "",
            "材质": "", "件(张)数": "", "重量(吨)": "",
            "库位号": "", "入库单价": "",
        })

    return records


def fetch_all_records(session: requests.Session,
                      danhao_list: list[str]) -> list[dict]:
    """逐个访问详情页，解析所有记录"""
    all_records: list[dict] = []
    total = len(danhao_list)

    for i, danhao in enumerate(danhao_list, 1):
        print(f"[详情] {i}/{total} 解析 {danhao} ...", end=" ", flush=True)
        try:
            resp = session.get(JIAGONG_DETAIL_PAGE,
                               params={"jg_danhao": danhao},
                               timeout=15)
            if resp.status_code != 200 or len(resp.text) < 500:
                print(f"HTTP {resp.status_code}, len={len(resp.text)} [跳过]")
                continue

            records = parse_jigondan_html(resp.text, danhao)
            print(f"{len(records)} 行明细")
            all_records.extend(records)
        except Exception as exc:
            print(f"异常: {exc}")

        time.sleep(0.3)  # 友好限速

    print(f"\n[汇总] 共解析 {total} 个加工单, {len(all_records)} 条记录")
    return all_records


# ---------------- CSV ----------------

CSV_FIELDS = [
    "加工单号", "服务商名称", "服务商联系人", "服务商公司",
    "结算单位", "税率", "金额", "备注", "加工备注",
    "明细行ID", "捆包号", "品名", "规格", "产地",
    "材质", "件(张)数", "重量(吨)", "库位号", "入库单价",
]


def records_to_csv_bytes(records: list[dict]) -> bytes:
    if not records:
        return b""
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=CSV_FIELDS, extrasaction="ignore")
    writer.writeheader()
    for r in records:
        row = {f: str(r.get(f, "") or "") for f in CSV_FIELDS}
        writer.writerow(row)
    return buf.getvalue().encode("utf-8-sig")


# ---------------- 飞书多维表格 ----------------

def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": app_id, "app_secret": app_secret},
                      timeout=15, proxies={"http": None, "https": None})
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换取 tenant_access_token 失败: {data}")
    token = str(data.get("tenant_access_token") or "")
    if not token:
        raise RuntimeError(f"响应中无 tenant_access_token: {data}")
    print(f"[飞书] 获取 tenant_access_token 成功 (len={len(token)})")
    return token


def bitable_find_existing_zids(token: str, app_token: str,
                                table_id: str,
                                zid_list: list[str]) -> set[str]:
    """按"明细行ID"用条件查询检查哪些已存在"""
    if not zid_list:
        return set()

    unique_zids = list({z.strip() for z in zid_list if z and z.strip()})
    if not unique_zids:
        return set()

    no_proxy = {"http": None, "https": None}
    existing: set[str] = set()
    page_token = ""
    seen_tokens: set[str] = set()
    max_pages = 20

    conditions = [
        {"field_name": "明细行ID", "operator": "is", "value": [z]}
        for z in unique_zids
    ]

    page_count = 0
    while True:
        page_count += 1
        if page_count > max_pages:
            break
        if page_token and page_token in seen_tokens:
            break
        if page_token:
            seen_tokens.add(page_token)

        url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
               f"/tables/{table_id}/records/search")
        body: dict = {
            "page_size": 500,
            "filter": {
                "conjunction": "or",
                "conditions": conditions,
            },
        }
        if page_token:
            body["page_token"] = page_token
        r = requests.post(url, headers={"Authorization": f"Bearer {token}"},
                          json=body, timeout=15, proxies=no_proxy)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"bitable 条件查询失败: {data}")
        d = data.get("data") or {}
        items = d.get("items") or []
        for it in items:
            f = it.get("fields") or {}
            v = f.get("明细行ID") or ""
            if isinstance(v, list):
                for seg in v:
                    if isinstance(seg, dict) and seg.get("text"):
                        existing.add(str(seg["text"]).strip())
                        break
            elif isinstance(v, str):
                if v.strip():
                    existing.add(v.strip())
        has_more = d.get("has_more", False)
        page_token = d.get("page_token") or ""
        if not has_more or not page_token:
            break

    print(f"[多维表格] 待查 {len(unique_zids)} 个明细行ID，已存在 {len(existing)} 个")
    return existing


def bitable_batch_create(token: str, app_token: str, table_id: str,
                         records: list[dict]) -> list[dict]:
    if not records:
        return []
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/records/batch_create")
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"},
                      json={"records": records}, timeout=15,
                      proxies={"http": None, "https": None})
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"bitable 批量创建失败: {json.dumps(data, ensure_ascii=False)[:1200]}")
    return (data.get("data") or {}).get("records") or []


def bitable_append_records(token: str, app_token: str, table_id: str,
                           rows: list[dict],
                           dedup: bool = True) -> dict:
    """
    将 rows（解析出的记录列表）追加到多维表格。
    去重逻辑：按"明细行ID"条件查询，存在的跳过。
    """
    existing_zids: set[str] = set()
    if dedup:
        zid_list = [str(r.get("明细行ID") or "").strip() for r in rows]
        try:
            existing_zids = bitable_find_existing_zids(
                token, app_token, table_id, zid_list)
        except Exception as exc:
            print(f"[警告] 按明细行ID条件查询失败（将不做去重）: {exc}")

    # "加工单"表所有字段都是 type=1 (Text)，全部按文本写入
    records_payload: list[dict] = []
    skipped = 0
    for row in rows:
        zid = str(row.get("明细行ID") or "").strip()
        if dedup and zid and zid in existing_zids:
            skipped += 1
            continue
        fields: dict = {}
        for col_name in CSV_FIELDS:
            v = str(row.get(col_name, "") or "").strip()
            if v:
                fields[col_name] = v
        records_payload.append({"fields": fields})

    print(f"[多维表格] 本次将追加 {len(records_payload)} 条 "
          f"（跳过重复 {skipped} 条）")

    created_ids: list[str] = []
    for i in range(0, len(records_payload), 500):
        batch = records_payload[i:i + 500]
        created = bitable_batch_create(token, app_token, table_id, batch)
        created_ids.extend([(c or {}).get("record_id", "") for c in created])
        print(f"  批次 {i // 500 + 1}: 成功写入 {len(created)} 条")
        if i + 500 < len(records_payload):
            time.sleep(0.5)

    return {"created": len(created_ids), "skipped": skipped,
            "record_ids": created_ids}


# ---------------- 飞书机器人通知 ----------------

def _feishu_sign(secret: str, timestamp: str) -> str:
    import hmac
    import hashlib
    import base64 as _b64
    string_to_sign = f"{timestamp}\n{secret}"
    h = hmac.new(secret.encode("utf-8"),
                 string_to_sign.encode("utf-8"),
                 digestmod=hashlib.sha256)
    return _b64.b64encode(h.digest()).decode("utf-8")


def feishu_send_bot_text(webhook_url: str, secret: str, text: str) -> None:
    if not webhook_url:
        return
    body: dict = {"msg_type": "text", "content": {"text": text}}
    if secret:
        ts = str(int(time.time()))
        body["timestamp"] = ts
        body["sign"] = _feishu_sign(secret, ts)
    r = requests.post(webhook_url, json=body, timeout=15,
                      proxies={"http": None, "https": None})
    data = r.json()
    if data.get("code") != 0:
        print(f"[飞书机器人] 发送失败: {data}")
    else:
        print(f"[飞书机器人] 通知发送成功")


# ---------------- 主流程 ----------------

def main() -> int:
    username = env("DH_USERNAME")
    password = env("DH_PASSWORD")
    delivery_mode = env("DELIVERY_MODE", "bitable").lower()

    fs_app_id = env("FEISHU_APP_ID")
    fs_app_secret = env("FEISHU_APP_SECRET")
    fs_webhook_url = env("FEISHU_WEBHOOK_URL")
    fs_webhook_secret = env("FEISHU_WEBHOOK_SECRET")

    bt_app_token = env("BITABLE_APP_TOKEN")
    bt_table_id = env("BITABLE_TABLE_ID")
    bt_dedup = env("BITABLE_DEDUP", "1") != "0"

    zhuantai = env("FILTER_ZHUANTAI")
    max_count = int(env("JIAGONG_MAX", "0") or "0")

    if not username or not password:
        print("[错误] 缺少 DH_USERNAME / DH_PASSWORD")
        return 2

    # 1) 登录
    session = create_session()
    if not login(session, username, password):
        return 1

    # 2) 获取加工单号列表
    print(f"\n[步骤1] 获取加工单号列表 (状态={zhuantai or '全部'})")
    danhao_list = get_jiagong_danhao_list(session, zhuantai=zhuantai,
                                           max_count=max_count)
    if not danhao_list:
        print("[警告] 没有获取到任何加工单号")
        return 0

    # 3) 逐个解析详情页
    print(f"\n[步骤2] 逐个解析 {len(danhao_list)} 个加工单详情页")
    records = fetch_all_records(session, danhao_list)
    session.close()

    if not records:
        print("[警告] 没有解析到任何记录")
        return 0

    # 4) 生成 CSV
    csv_bytes = records_to_csv_bytes(records)
    now = datetime.datetime.now()
    filename = f"jiagongdan_{now.strftime('%Y%m%d_%H%M')}.csv"

    summary_lines = [
        f"导出时间: {now.strftime('%Y-%m-%d %H:%M:%S')}",
        f"加工单数: {len(danhao_list)} 个",
        f"明细记录: {len(records)} 条",
        f"文件名: {filename}",
        f"文件大小: {len(csv_bytes)/1024:.1f} KB",
    ]
    print("\n[汇总]")
    for line in summary_lines:
        print(f"  {line}")

    # ========== 交付：本地 ==========
    if delivery_mode == "local":
        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        with open(out_path, "wb") as fp:
            fp.write(csv_bytes)
        print(f"\n[本地保存] 已保存到 {out_path}")
        return 0

    # ========== 交付：飞书多维表格 ==========
    if delivery_mode == "bitable":
        if not (fs_app_id and fs_app_secret and bt_app_token and bt_table_id):
            print("[错误] 交付 bitable 缺少参数: FEISHU_APP_ID / FEISHU_APP_SECRET / "
                  "BITABLE_APP_TOKEN / BITABLE_TABLE_ID")
            out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            with open(out_path, "wb") as fp:
                fp.write(csv_bytes)
            print(f"[本地保存] 已保存到 {out_path}")
            return 0

        try:
            token = feishu_tenant_access_token(fs_app_id, fs_app_secret)
            result = bitable_append_records(
                token=token,
                app_token=bt_app_token,
                table_id=bt_table_id,
                rows=records,
                dedup=bt_dedup,
            )

            notify_lines = ["✅ 加工单 → 飞书多维表格导出完成"] + summary_lines
            notify_lines.append(f"目标表: app_token={bt_app_token} table_id={bt_table_id}")
            notify_lines.append(f"去重: {'开' if bt_dedup else '关'}")
            notify_lines.append(f"新增 {result['created']} 条 / 跳过重复 {result['skipped']} 条")
            notify_text = "\n".join(notify_lines)

            if fs_webhook_url:
                try:
                    feishu_send_bot_text(fs_webhook_url, fs_webhook_secret, notify_text)
                except Exception as exc:
                    print(f"[警告] 通知发送失败，但写入已完成: {exc}")
            else:
                print("\n[通知预览]")
                for line in notify_lines:
                    print(f"  {line}")

            print(f"\n[完成] 新增 {result['created']} 条 / 跳过重复 {result['skipped']} 条")
        except Exception as exc:
            print(f"[错误] 飞书多维表格写入失败: {exc}")
            traceback.print_exc()
            out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            with open(out_path, "wb") as fp:
                fp.write(csv_bytes)
            print(f"[本地保存] 已保存到 {out_path}")
            return 4
        return 0

    print(f"[错误] 未知 DELIVERY_MODE={delivery_mode}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
