#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火「销售订单 → 订单明细汇总」→ 飞书多维表格「数据汇总（2026）/订单明细」自动更新工作流。

流程：
  1. ddddocr + Playwright 自动登录懂火钢城系统
  2. 直达「订单明细汇总」页 v_xmxhz（该页是「销售订单」模块里"订单明细汇总"Tab 的
     内容页，支持像出库记录 v_xjlall 一样独立直达，已实测验证）
  3. 点"筛选"，把「出库状态」设为"全部"，点确认按钮
  4. 点系统自带"导出"按钮 → 下载 HTML 格式 .xls（不扒接口）
  5. pandas 解析 → 转 CSV（UTF-8-SIG）落本地备份（上传成功后自动删除）
  6. 清空飞书多维表 tblcEZoQatk7lCAO 现有全部记录
  7. 按多维表字段格式批量写入新数据（字段类型通过飞书 API 动态探测，无需硬编码）
  8. 发送飞书卡片通知给洪（更新条数、耗时；失败时发红色告警卡片）

与 sync_chuku_to_bitable.py 的关系：
  同属懂火同步项目，登录方式与 chuku 完全一致（浏览器内联登录，原因见 chuku
  脚本内「登录分工说明」）；两脚本保持独立、互不影响，故意不抽公共库。
  差异点：导航路径、筛选条件（出库状态=全部 vs 日期区间）、目标表、字段映射方式
  （本表字段通过 API 动态探测，chuku 表是硬编码）。

使用：python sync_dingdan_to_bitable.py [--headless] [--skip-download] [--dry-run] [--no-notify]
凭据：仓库根目录 .env 里的 DH_USERNAME / DH_PASSWORD + 系统环境变量 FEISHU_APP_ID / FEISHU_APP_SECRET
"""
import sys, os, json, time, datetime, argparse, traceback, math, requests
from pathlib import Path

# ===== stdout/stderr 编码双保险（Windows subprocess 里 print emoji 会崩）=====
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

from dotenv import load_dotenv
load_dotenv(Path(__file__).resolve().parent.parent / ".env", override=False)  # .env 统一放仓库根目录

# ===== 固定配置 =====
BITABLE_APP_TOKEN = "VahHb3YDBaBTwTsCjeAcaAhhnHc"   # 数据汇总（2026）
BITABLE_TABLE_ID  = "tblcEZoQatk7lCAO"               # 订单明细
DONGHUO_BASE      = "https://erpa.donghuo.vip"
DONGHUO_LOGIN_URL = f"{DONGHUO_BASE}/view/admin/v_login"
# 「订单明细汇总」内容页：懂火里它挂在「销售订单」模块的 Tab 上（容器页 v_ifram_dd，
# 内容 iframe 原始 src=v_xmxhz），与出库记录 v_xjlall 一样支持登录后独立直达（2026-09-05 实测）
DONGHUO_ORDER_SUMMARY_URL = f"{DONGHUO_BASE}/view/admin/xiaoshou/v_xmxhz"
FILTER_LABEL = "出库状态"    # 筛选面板里的字段名
FILTER_VALUE = "全部"        # 筛选面板里要选的值

DOWNLOAD_DIR = Path(__file__).parent / "downloads" / "dingdan"
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
CSV_DIR = Path(__file__).parent / "csv_backup" / "dingdan"
CSV_DIR.mkdir(parents=True, exist_ok=True)

FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
BATCH_SIZE = 500    # 飞书 bitable batch_create/batch_delete 上限

# 同步完成后飞书通知（默认发给 洪 on_b09bcbf3e74f5d423900aa9b2f00eb63）
FEISHU_NOTIFY_UNION_ID = "on_b09bcbf3e74f5d423900aa9b2f00eb63"

# ===== 飞书多维表字段类型码 → 写入策略（2026-09-05 API 规范）=====
# 1=Text 2=Number 3=SingleSelect 4=MultiSelect 5=DateTime 7=Checkbox 13=Phone
# 其余（11=User 15=Url 17=Attachment 18=Link 19=Lookup 20=Formula 1001+=自动字段）
# 均不可直接写文本值，跳过不写。
FIELD_TYPE_CODE_MAP = {
    1: "text", 2: "number", 3: "text", 4: "multiselect",
    5: "datetime", 7: "checkbox", 13: "text",
}

# 懂火导出列名 → 多维表字段名 别名映射（2026-09-05 实测表结构）：
#   表里文本字段叫「订单 号」（中间带空格），另有一个「订单号」是公式字段不可写；
#   导出的「备注」列在表里叫「注释」。
COLUMN_ALIASES = {
    "订单号": "订单 号",
    "备注": "注释",
}


# ============ 工具函数 ============

def log(msg: str):
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def env(name: str) -> str:
    return os.environ.get(name, "").strip()


def feishu_token() -> str:
    """获取飞书 tenant_access_token（每次现取，2 小时有效期）"""
    app_id = os.environ.get("FEISHU_APP_ID") or os.environ.get("FEISHU_NOTIFY_APP_ID", "").strip()
    app_secret = os.environ.get("FEISHU_APP_SECRET") or os.environ.get("FEISHU_NOTIFY_APP_SECRET", "").strip()
    if not app_id or not app_secret:
        raise RuntimeError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET（或 FEISHU_NOTIFY_APP_ID / NOTIFY_APP_SECRET）环境变量")
    r = requests.post(
        f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret},
        timeout=15,
    )
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换 tenant_access_token 失败: {data}")
    return data["tenant_access_token"]


# ===== 通知卡片模板 =====
# 同步数据来源/目标描述（通知卡片里展示）
SOURCE_DESC = "懂火「销售订单 → 订单明细汇总」（出库状态=全部）"
TARGET_DESC = "飞书多维表「数据汇总（2026）/订单明细」"

# 当前执行环节（失败通知卡片里定位用）
CURRENT_STEP = "初始化"


def feishu_send_card(union_id: str, card: dict, token: str):
    """通过飞书机器人给指定用户发卡片消息"""
    url = f"{FEISHU_OPEN_BASE}/im/v1/messages?receive_id_type=union_id"
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "receive_id": union_id,
        "msg_type": "interactive",
        "content": json.dumps(card, ensure_ascii=False),
    }
    r = requests.post(url, headers=h, json=body, timeout=15)
    data = r.json()
    if data.get("code") != 0:
        log(f"[飞书通知] ❌ 发送失败: code={data.get('code')} msg={data.get('msg')}")
    else:
        log(f"[飞书通知] ✅ 卡片已发送给 {union_id}")


def build_success_card(written: int, cleared, elapsed_s: float) -> dict:
    """同步成功通知卡片（绿色）"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    fields = [
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**更新记录**\n{written} 条"}},
        {"is_short": True, "text": {"tag": "lark_md",
            "content": f"**清空旧记录**\n{cleared} 条" if cleared is not None else "**清空旧记录**\n跳过"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**耗时**\n{elapsed_s:.1f} 秒"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**数据来源**\n{SOURCE_DESC}"}},
    ]
    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "green", "title": {"tag": "plain_text", "content": "✅ 懂火订单明细同步成功"}},
        "elements": [
            {"tag": "div", "fields": fields},
            {"tag": "hr"},
            {"tag": "note", "elements": [
                {"tag": "plain_text", "content": f"{TARGET_DESC} ｜ 同步时间 {now}"}
            ]},
        ],
    }


def build_failure_card(step: str, error: str) -> dict:
    """同步失败通知卡片（红色）"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "red", "title": {"tag": "plain_text", "content": "❌ 懂火订单明细同步失败"}},
        "elements": [
            {"tag": "div", "fields": [
                {"is_short": True, "text": {"tag": "lark_md", "content": f"**失败环节**\n{step}"}},
                {"is_short": True, "text": {"tag": "lark_md", "content": f"**数据来源**\n{SOURCE_DESC}"}},
            ]},
            {"tag": "div", "text": {"tag": "lark_md", "content": f"**错误信息**\n{str(error)[:600]}"}},
            {"tag": "hr"},
            {"tag": "note", "elements": [{"tag": "plain_text", "content": f"同步时间 {now}"}]},
        ],
    }


# ============ 懂火 → 下载 xls ============
#
# 登录方式与 sync_chuku_to_bitable.py 完全一致：浏览器内联登录（ddddocr 识别验证码
# → 填表单 → 点 #laysubmit，重试 10 次）。为什么不用 requests 版 donghuo_login.py，
# 见 chuku 脚本内「登录分工说明」。账号密码同样走 .env 的 DH_USERNAME / DH_PASSWORD。

def find_owner_for_text(page, text: str, tries: int = 10, interval_ms: int = 1000):
    """
    在主页面与所有 iframe 中查找精确文本元素，返回承载它的 Page 或 Frame。
    懂火是 layui admin，点导航打开的内容可能在 iframe 里；直接 goto URL 则在主页面。
    找不到返回 None。
    """
    for _ in range(tries):
        try:
            if page.get_by_text(text, exact=True).count() > 0:
                return page
        except Exception:
            pass
        for f in page.frames:
            try:
                if f.get_by_text(text, exact=True).count() > 0:
                    return f
            except Exception:
                pass
        page.wait_for_timeout(interval_ms)
    return None


JS_SET_SELECT = """
([labelText, wantText]) => {
    // 把筛选面板里 labelText 对应的 select 设为 wantText 选项
    const selects = Array.from(document.querySelectorAll('select'));
    const hasOpt = (s, t) => Array.from(s.options).some(o => o.textContent.trim() === t);
    const report = selects.map(s => ({
        id: s.id, name: s.name,
        opts: Array.from(s.options).map(o => o.textContent.trim()).slice(0, 10),
    }));
    let hit = null;
    // 规则1：通过 layui form-item 的 label 文本关联
    for (const s of selects) {
        if (!hasOpt(s, wantText)) continue;
        const item = s.closest('.layui-form-item') || s.parentElement;
        const lb = item ? item.querySelector('.layui-form-label, label, .layui-inline') : null;
        if (lb && lb.textContent.trim().includes(labelText)) { hit = s; break; }
    }
    // 规则2：兜底——全页唯一一个含该选项的 select
    if (!hit) {
        const cands = selects.filter(s => hasOpt(s, wantText));
        if (cands.length === 1) hit = cands[0];
    }
    if (!hit) return {ok: false, report};
    const target = Array.from(hit.options).find(o => o.textContent.trim() === wantText);
    hit.value = target.value;
    hit.dispatchEvent(new Event('input', {bubbles: true}));
    hit.dispatchEvent(new Event('change', {bubbles: true}));
    try { if (window.layui && layui.form) layui.form.render('select'); } catch (e) {}
    return {
        ok: true, id: hit.id, name: hit.name, value: hit.value,
        text: hit.selectedOptions[0] ? hit.selectedOptions[0].textContent.trim() : '',
    };
}
"""

JS_PICK_CONFIRM = """
() => {
    // 在筛选弹窗里挑确认按钮：优先文本含 查询/搜索/确定/确认/提交 的可见 layui-btn，
    // 其中优先 y>500（弹窗底部），再按 DOM 靠后者（弹窗层后插入）；找不到才退回旧启发式。
    const all = Array.from(document.querySelectorAll('button.layui-btn, a.layui-btn'))
      .map((b, i) => {
        const r = b.getBoundingClientRect();
        return {idx: i, cls: b.className, txt: b.textContent.trim(),
                y: Math.round(r.top), w: Math.round(r.width), h: Math.round(r.height)};
      })
      .filter(b => b.w > 0 && b.h > 0 && b.y > 0);
    const want = ['查询', '搜索', '确定', '确认', '提交'];
    let cands = all.filter(b => b.cls && !b.cls.includes('close') && want.some(t => b.txt.includes(t)));
    if (cands.length) {
        cands.sort((a, b) => ((b.y > 500) - (a.y > 500)) || (b.idx - a.idx));
        return cands[0];
    }
    const popup = all.filter(b => b.y > 500 && b.cls && !b.cls.includes('close'));
    const picked = popup[0] || all[0] || null;
    return picked;
}
"""

JS_CLICK_VISIBLE_OPTION = """
(wantText) => {
    // layui 渲染层下拉：点当前可见的 dd 选项（文本精确等于 wantText）
    const dds = Array.from(document.querySelectorAll('.layui-form-select dl dd'));
    const visible = dds.filter(d => d.offsetParent !== null);
    const t = visible.find(d => d.textContent.trim() === wantText);
    if (t) { t.click(); return true; }
    return false;
}
"""


def download_dingdan_xls(headless: bool = False) -> Path:
    """
    用 Playwright 自动登录懂火 → 直达订单明细汇总页 → 筛选出库状态=全部
    → 点系统导出按钮 → 下载 xls。返回下载的文件路径。
    """
    import ddddocr
    import playwright.sync_api as pw

    username = env("DH_USERNAME")
    password = env("DH_PASSWORD")
    if not username or not password:
        raise RuntimeError("缺少 DH_USERNAME / DH_PASSWORD")

    ocr = ddddocr.DdddOcr(show_ad=False)
    log(f"[懂火] 账号: {username}")
    global CURRENT_STEP
    CURRENT_STEP = "登录懂火"

    # 清下载目录
    for f in DOWNLOAD_DIR.iterdir():
        if f.is_file():
            f.unlink()

    with pw.sync_playwright() as p:
        browser = p.chromium.launch(
            channel="chrome",
            headless=headless,
            args=["--disable-blink-features=AutomationControlled"],
        )
        ctx = browser.new_context(accept_downloads=True)
        page = ctx.new_page()

        # --- 登录（重试最多 10 次，与 chuku 完全一致）---
        logged_in = False
        for attempt in range(1, 11):
            log(f"[懂火] 登录尝试 {attempt}/10 ...")
            page.goto(DONGHUO_LOGIN_URL, wait_until="domcontentloaded", timeout=20000)
            page.wait_for_timeout(1200)

            user_input = page.locator("input.layui-input:not(#captcha):not([type='password'])").first
            pwd_input = page.locator("#u_pass")
            captcha_input = page.locator("#captcha")
            user_input.fill(username)
            pwd_input.fill(password)

            captcha_img = page.locator("img[src*='captcha']").first
            code = ocr.classification(captcha_img.screenshot()).strip()
            log(f"       验证码识别: '{code}'")
            captcha_input.fill(code)
            page.wait_for_timeout(200)
            page.locator("#laysubmit").click()
            page.wait_for_timeout(1200)

            if "/v_login" not in page.url:
                log("[懂火] ✅ 登录成功")
                logged_in = True
                break

        if not logged_in:
            browser.close()
            raise RuntimeError("懂火登录失败，已达最大重试次数")

        # --- 直达「订单明细汇总」页 ---
        CURRENT_STEP = "打开订单明细汇总页"
        page.goto(DONGHUO_ORDER_SUMMARY_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(4000)

        owner = find_owner_for_text(page, "筛选", tries=6)
        if owner is None:
            browser.close()
            raise RuntimeError(f"未找到「筛选」元素（{DONGHUO_ORDER_SUMMARY_URL} 页面结构可能已变化）")
        log(f"[懂火] ✅ 已到达订单明细汇总页: {page.url}（载体: {'主页面' if owner is page else 'iframe'}）")

        # --- 点"筛选"按钮 ---
        CURRENT_STEP = "设置筛选条件（出库状态=全部）"
        owner.get_by_text("筛选", exact=True).first.click()
        page.wait_for_timeout(1500)
        log("[懂火] ✅ 筛选面板已打开")

        # --- 把「出库状态」设为"全部" ---
        res = owner.evaluate(JS_SET_SELECT, [FILTER_LABEL, FILTER_VALUE])
        if res.get("ok"):
            log(f"[懂火] ✅ {FILTER_LABEL} 已设为 '{res.get('text')}'（select id={res.get('id')} name={res.get('name')}）")
        else:
            log(f"[懂火] ⚠️ 原生 select 未命中，页面 select 概况: {json.dumps(res.get('report'), ensure_ascii=False)}")
            # 兜底：layui 渲染层下拉（点开下拉再点"全部"选项）
            try:
                label_loc = owner.get_by_text(FILTER_LABEL, exact=True).first
                parent = label_loc.locator("xpath=..")
                dd = parent.locator(".layui-form-select").first
                if dd.count() == 0:
                    dd = owner.locator(".layui-form-select").last
                dd.locator("input.layui-input, .layui-select-title, .layui-edge").first.click()
                page.wait_for_timeout(600)
                clicked = owner.evaluate(JS_CLICK_VISIBLE_OPTION, FILTER_VALUE)
                page.wait_for_timeout(300)
                if clicked:
                    log(f"[懂火] ✅ 已通过 layui 渲染层把 {FILTER_LABEL} 点选为 '{FILTER_VALUE}'")
                else:
                    raise RuntimeError("渲染层未见可见的'全部'选项")
            except Exception as e:
                browser.close()
                raise RuntimeError(f"设置 {FILTER_LABEL}='{FILTER_VALUE}' 失败: {e}；请人工确认筛选面板结构") from e

        # --- 点筛选面板的确认按钮 ---
        btn = owner.evaluate(JS_PICK_CONFIRM)
        if btn is None:
            browser.close()
            raise RuntimeError("筛选面板未找到确认按钮")
        log(f"[懂火] 点击筛选确认按钮: '{btn['txt']}' (y={btn['y']})")
        owner.locator("button.layui-btn, a.layui-btn").nth(btn["idx"]).click()
        page.wait_for_timeout(2500)
        log("[懂火] ✅ 筛选已应用")

        # --- 点"导出"按钮，等待下载 ---
        CURRENT_STEP = "点击系统导出按钮下载"
        log("[懂火] 开始导出（可能几秒到几十秒）...")
        export_loc = owner.get_by_text("导出", exact=True).first
        with page.expect_download(timeout=180000) as dl_info:
            export_loc.click()
        dl = dl_info.value
        target = DOWNLOAD_DIR / dl.suggested_filename
        dl.save_as(str(target))
        size_kb = target.stat().st_size / 1024
        log(f"[懂火] ✅ 下载完成: {target.name} ({size_kb:.1f} KB)")

        browser.close()
        return target


# ============ xls → DataFrame → CSV ============

def xls_to_clean_df(xls_path: Path):
    """
    解析懂火导出的 HTML 格式 xls → 返回 DataFrame（第 0 行为表头已跳过，列名已正确设置）。
    同时落一份 UTF-8-SIG CSV 到 csv_backup/dingdan/（上传成功后由调用方删除）。
    返回 (DataFrame, csv_path)。
    """
    import pandas as pd
    df_raw = pd.read_html(str(xls_path))[0]
    # 懂火导出的 xls 第 0 行就是字段名，第 1 行起是数据
    header = list(df_raw.iloc[0])
    df = df_raw.iloc[1:].reset_index(drop=True)
    df.columns = header
    # 去掉可能的完全空行
    df = df.dropna(how="all")
    log(f"[解析] {xls_path.name}: {len(df)} 行 × {len(df.columns)} 列")

    # 落 CSV 备份
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = CSV_DIR / f"dingdan_export_{now}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    log(f"[备份] CSV 已保存: {csv_path}")
    return df, csv_path


# ============ 飞书多维表：字段探测 + 清空 + 写入 ============

def bitable_get_field_types(token: str) -> dict[str, str]:
    """通过 API 动态探测多维表字段类型，返回 {字段名: 写入策略}"""
    h = {"Authorization": f"Bearer {token}"}
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/fields"
    out = {}
    page_token = None
    while True:
        qs = "page_size=100" + (f"&page_token={page_token}" if page_token else "")
        r = requests.get(f"{url}?{qs}", headers=h, timeout=30)
        d = r.json()
        if d.get("code") != 0:
            raise RuntimeError(f"获取字段列表失败: {d}")
        dd = d.get("data") or {}
        for f in dd.get("items") or []:
            out[f["field_name"]] = FIELD_TYPE_CODE_MAP.get(f["type"], "unsupported")
        if not dd.get("has_more"):
            break
        page_token = dd.get("page_token")
        if not page_token:
            break
    return out


def bitable_list_all_records(token: str) -> list[str]:
    """返回多维表内所有 record_id（⚠️ search 接口的 page_token 会永远不推进，必须用 GET list 接口 + query string）"""
    h = {"Authorization": f"Bearer {token}"}
    base_url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records"
    all_ids = []
    seen_ids = set()
    page_token = None
    while True:
        qs = f"page_size={BATCH_SIZE}"
        if page_token:
            qs += f"&page_token={page_token}"
        r = requests.get(f"{base_url}?{qs}", headers=h, timeout=30)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"list records 失败: {data}")
        d = data.get("data") or {}
        items = d.get("items") or []
        if not items:
            break
        dup = 0
        for it in items:
            rid = it.get("record_id")
            if rid and rid not in seen_ids:
                all_ids.append(rid)
                seen_ids.add(rid)
            else:
                dup += 1
        if dup == len(items):
            # 全重复 → page_token 没推进，硬停
            log(f"[飞书] ⚠️ page_token 未推进，停止；累计 {len(all_ids)} 条")
            break
        if len(all_ids) % 5000 == 0:
            log(f"[飞书] list: 累计 {len(all_ids)}")
        if not d.get("has_more"):
            break
        page_token = d.get("page_token")
        if not page_token:
            break
    log(f"[飞书] list 完成: 共 {len(all_ids)} 条（去重后）")
    return all_ids


def bitable_batch_delete(token: str, record_ids: list[str]):
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    # ignore_consistency_check=true 加速（牺牲强一致换吞吐）
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records/batch_delete?ignore_consistency_check=true"
    total = len(record_ids)
    for i in range(0, total, BATCH_SIZE):
        batch = record_ids[i:i + BATCH_SIZE]
        r = requests.post(url, headers=h, json={"records": batch}, timeout=30)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_delete 失败: batch {i // BATCH_SIZE + 1}, err={data.get('msg')}")
        if (i // BATCH_SIZE + 1) % 10 == 0 or i + BATCH_SIZE >= total:
            log(f"[飞书] delete: {len(batch)} 条 (批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE})")


def _convert_value(raw, ftype: str):
    """把 DataFrame 单元格值转成飞书 API 接受的字段值"""
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return None
    if ftype == "text":
        s = str(raw).strip()
        return s if s else None
    if ftype == "number":
        s = str(raw).strip()
        if not s:
            return None
        try:
            f = float(s.replace(",", ""))
            return int(f) if f == int(f) else f
        except ValueError:
            return None
    if ftype == "datetime":
        if isinstance(raw, (int, float)) and not isinstance(raw, bool):
            return int(raw) if raw > 10**12 else int(raw * 1000)  # 秒/毫秒时间戳兜底
        s = str(raw).strip()
        if not s:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                    "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
            try:
                dt = datetime.datetime.strptime(s, fmt)
                return int(dt.timestamp() * 1000)
            except ValueError:
                continue
        return None
    if ftype == "multiselect":
        s = str(raw).strip()
        if not s:
            return None
        parts = [p.strip() for p in s.replace("，", ",").replace("、", ",").replace(";", ",").replace("；", ",").split(",")]
        vals = [p for p in parts if p]
        return vals or None
    if ftype == "checkbox":
        s = str(raw).strip().lower()
        if s in ("是", "true", "1", "y", "yes", "√", "已"):
            return True
        if s in ("否", "false", "0", "n", "no", ""):
            return False
        return None
    return None


def df_to_bitable_records(df, field_types: dict) -> list[dict]:
    """把 DataFrame 转成飞书 batch_create 需要的 [{"fields": {...}}, ...] 列表（按动态探测的字段类型）"""
    # 表字段名 → 导出列名：同名优先；COLUMN_ALIASES 兜底（如「订单 号」← 订单号）
    field_to_col = {f: f for f in field_types}
    for csv_col, field_name in COLUMN_ALIASES.items():
        if field_name in field_types and field_name not in df.columns and csv_col in df.columns:
            field_to_col[field_name] = csv_col
    # 提示一次：完全对不上的列 / 不可写字段
    for col in df.columns:
        if col not in field_types and col not in COLUMN_ALIASES:
            log(f"[飞书] ⚠️ 列 '{col}' 在多维表里不存在，跳过")
    for field_name, ftype in field_types.items():
        src = field_to_col.get(field_name)
        if ftype == "unsupported" and src and src in df.columns:
            log(f"[飞书] ⚠️ 字段 '{field_name}' 类型不可直接写入（公式/查找/人员等），跳过")
    records = []
    for _, row in df.iterrows():
        fields = {}
        for field_name, ftype in field_types.items():
            if ftype == "unsupported":
                continue
            src = field_to_col.get(field_name)
            if src is None or src not in df.columns:
                continue
            val = _convert_value(row.get(src), ftype)
            if val is not None:
                fields[field_name] = val
        records.append({"fields": fields})
    return records


def bitable_batch_create(token: str, records: list[dict]):
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records/batch_create?ignore_consistency_check=true"
    total = len(records)
    created = 0
    for i in range(0, total, BATCH_SIZE):
        batch = records[i:i + BATCH_SIZE]
        r = requests.post(url, headers=h, json={"records": batch}, timeout=60)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_create 失败 batch {i // BATCH_SIZE + 1}: code={data.get('code')} msg={data.get('msg')} sample={str(data)[:500]}")
        created += len((data.get("data") or {}).get("records") or [])
        log(f"[飞书] create: {len(batch)} 条 (批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE})，累计已建 {created}")
    return created


# ============ 主流程 ============

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--headless", action="store_true", help="headless Chrome（默认 headed，便于人工观察）")
    ap.add_argument("--skip-download", action="store_true", help="跳过下载，直接用 downloads/dingdan/ 下最新的 xls")
    ap.add_argument("--skip-clear", action="store_true", help="跳过清空旧表（仅追加，不推荐，仅调试用）")
    ap.add_argument("--skip-upload", action="store_true", help="跳过写入多维表（只下载+转 CSV）")
    ap.add_argument("--download-only", action="store_true", help="等价 --skip-upload --skip-clear")
    ap.add_argument("--dry-run", action="store_true", help="下载+转 CSV+打印前几行，不碰飞书")
    ap.add_argument("--no-notify", action="store_true", help="不发送飞书通知（默认同步完成会通知洪）")
    args = ap.parse_args()

    global CURRENT_STEP
    t0 = time.time()
    log("==== 懂火订单明细汇总 → 飞书多维表 同步工作流 启动 ====")

    # ---- Step 1: 下载 ----
    if args.skip_download:
        CURRENT_STEP = "读取本地导出文件"
        xls_files = sorted(DOWNLOAD_DIR.glob("*.xls"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not xls_files:
            raise RuntimeError("--skip-download 但 downloads/dingdan/ 下没有 xls 文件")
        xls_path = xls_files[0]
        log(f"[跳过下载] 使用已有文件: {xls_path}")
    else:
        xls_path = download_dingdan_xls(headless=args.headless)

    # ---- Step 2: 解析 ----
    CURRENT_STEP = "解析导出文件"
    df, csv_path = xls_to_clean_df(xls_path)
    if len(df) == 0:
        raise RuntimeError("导出文件为空，中止")

    if args.dry_run:
        log("[DRY-RUN] 不操作飞书")
        log(f"  列名: {list(df.columns)}")
        log(f"  前 3 行: {df.head(3).to_dict(orient='records')}")
        elapsed = time.time() - t0
        log(f"==== 完成（DRY-RUN），耗时 {elapsed:.1f}s ====")
        return 0

    if args.download_only:
        elapsed = time.time() - t0
        log(f"==== 完成（仅下载+解析），耗时 {elapsed:.1f}s ====")
        return 0

    # ---- Step 3: 飞书 ----
    CURRENT_STEP = "获取飞书凭证与字段"
    token = feishu_token()
    log("[飞书] ✅ tenant_access_token 已获取")

    field_types = bitable_get_field_types(token)
    writable = [k for k, v in field_types.items() if v != "unsupported"]
    log(f"[飞书] ✅ 字段探测完成: 共 {len(field_types)} 个字段，可写 {len(writable)} 个")

    cleared_count = None
    if not args.skip_clear:
        CURRENT_STEP = "清空多维表旧记录"
        log("[飞书] Step A: 清空现有记录 ...")
        ids = bitable_list_all_records(token)
        log(f"[飞书] 现有记录数: {len(ids)}")
        if ids:
            bitable_batch_delete(token, ids)
            log("[飞书] ✅ 已清空")
        else:
            log("[飞书] 无需清空（表已空）")
        cleared_count = len(ids)
    else:
        log("[飞书] 跳过清空（--skip-clear）")

    if not args.skip_upload:
        CURRENT_STEP = "预检写入（前 5 条）"
        log(f"[飞书] Step B: 写入 {len(df)} 条新记录 ...")
        records = df_to_bitable_records(df, field_types)
        # 先小批量试写 5 条，验证字段类型映射没问题
        test_batch = records[:min(5, len(records))]
        h = {"Authorization": f"Bearer {token}"}
        url_test = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records/batch_create"
        r = requests.post(url_test, headers=h, json={"records": test_batch}, timeout=60)
        d = r.json()
        if d.get("code") != 0:
            log(f"[飞书] ❌ 预检写入失败: code={d.get('code')} msg={d.get('msg')}")
            log(f"  样本记录: {json.dumps(test_batch[0], ensure_ascii=False)[:800]}")
            raise RuntimeError(f"预检写入失败: code={d.get('code')} msg={d.get('msg')}")
        log(f"[飞书] ✅ 预检写入 {len(test_batch)} 条成功，继续写剩余 {len(records) - len(test_batch)} 条")
        CURRENT_STEP = "批量写入多维表"
        remaining = records[len(test_batch):]
        if remaining:
            bitable_batch_create(token, remaining)
        total_written = len(test_batch) + len(remaining)
        log(f"[飞书] ✅ 全部写入完成，共 {total_written} 条")
        # 上传成功后删除 CSV 备份
        if csv_path.exists():
            csv_path.unlink()
            log(f"[清理] CSV 已删除: {csv_path.name}")
    else:
        log("[飞书] 跳过写入（--skip-upload）")

    elapsed = time.time() - t0
    log(f"==== 完成，耗时 {elapsed:.1f}s ====")

    # 飞书通知（上传成功后才通知，--no-notify 可跳过）
    if not args.no_notify and not args.skip_upload and not args.dry_run and not args.download_only:
        try:
            CURRENT_STEP = "发送飞书通知"
            token = feishu_token()
            card = build_success_card(total_written, cleared_count, elapsed)
            feishu_send_card(FEISHU_NOTIFY_UNION_ID, card, token)
        except Exception as e:
            log(f"[飞书通知] ⚠️ 发送异常: {e}")

    return 0


if __name__ == "__main__":
    try:
        rc = main()
    except Exception as e:
        log(f"❌ 未捕获异常: {e}")
        traceback.print_exc()
        # 失败通知卡片（尽力而为，发送失败也不影响退出码）
        try:
            _t = feishu_token()
            feishu_send_card(FEISHU_NOTIFY_UNION_ID, build_failure_card(CURRENT_STEP, str(e)), _t)
        except Exception as ne:
            log(f"[飞书通知] ⚠️ 失败通知发送异常: {ne}")
        rc = 99
    sys.exit(rc)
