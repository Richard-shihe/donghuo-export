#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火「应收管理/应收结算」→ 飞书多维表格「数据汇总（2026）/tblpjne9dIuif5HD」自动更新工作流。

流程：
  1. ddddocr + Playwright 自动登录懂火钢城系统
  2. 打开 应收管理/应收结算 页（caiwu/v_x_jiesuan）
  3. 点"筛选" → 清空开始日期（结束日期保持今天）→ 点"查询"应用
  4. 点系统自带"导出"按钮（POST excelbiao/x_jiesuan）→ 下载 HTML 格式 .xls
  5. pandas 解析 → 转 CSV（UTF-8-SIG）落本地备份
  6. 清空飞书多维表 tblpjne9dIuif5HD 现有全部记录
  7. 按原字段格式批量写入新数据
  8. 上传成功后删除 CSV 备份
  9. 发送飞书通知卡片给洪（更新条数、耗时等摘要）

使用：python sync_yingshou_to_bitable.py [--headless] [--skip-download] [--dry-run] [--no-notify] ...
凭据：仓库根目录 .env 里的 DH_USERNAME / DH_PASSWORD + FEISHU_APP_ID / FEISHU_APP_SECRET
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
BITABLE_TABLE_ID  = "tblpjne9dIuif5HD"               # 应收管理（应收结算）
DONGHUO_LOGIN_URL = "https://erpa.donghuo.vip/view/admin/v_login"
DONGHUO_YS_URL    = "https://erpa.donghuo.vip/view/admin/caiwu/v_x_jiesuan"  # 应收管理→应收结算
EXPORT_START_DATE = ""      # 清空开始日期（用户要求：不设起始日期，导出全量）
DOWNLOAD_DIR = Path(__file__).parent / "downloads"
DOWNLOAD_DIR.mkdir(exist_ok=True)
CSV_DIR = Path(__file__).parent / "csv_backup"
CSV_DIR.mkdir(exist_ok=True)

FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
BATCH_SIZE = 500    # 飞书 bitable batch_create/batch_delete 上限

# 同步完成后飞书通知（默认发给 洪 on_b09bcbf3e74f5d423900aa9b2f00eb63）
FEISHU_NOTIFY_UNION_ID = "on_b09bcbf3e74f5d423900aa9b2f00eb63"

# ===== 懂火导出列名 → 多维表字段名 别名映射 =====
# 1) 导出列是「订单号」，多维表里文本字段叫「订单 号」（带空格）；「订单号」是公式列，禁止写入
# 2) 导出列是「销售状态」，多维表字段叫「发货状态」
COLUMN_ALIASES = {
    "订单号": "订单 号",
    "销售状态": "发货状态",
}

# ===== 多维表字段类型映射（2026-09-06 API 探测 tblpjne9dIuif5HD）=====
# 参与(人员)、订单号(公式)、字段1~7 不由本脚本写入
BITABLE_FIELD_TYPES = {
    "订单 号":     "text",
    "日期":         "datetime",
    "发货状态":     "text",
    "所属公司":     "text",
    "销售人":       "text",
    "客户名称":     "text",
    "实发重量":     "number",
    "实发金额":     "number",
    "销售费用":     "number",
    "其它款项":     "number",
    "已结金额":     "number",
    "未结金额":     "number",
}


# ============ 工具函数 ============

def log(msg: str):
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


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


def env(name: str) -> str:
    return os.environ.get(name, "").strip()


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


# ===== 通知卡片模板 =====
SOURCE_DESC = "懂火「应收管理/应收结算」（清空开始日期，全量）"
TARGET_DESC = "飞书多维表「数据汇总（2026）/应收管理 tblpjne9dIuif5HD」"

# 当前执行环节（失败通知卡片里定位用）
CURRENT_STEP = "初始化"


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
        "header": {"template": "green", "title": {"tag": "plain_text", "content": "✅ 懂火应收结算同步成功"}},
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
        "header": {"template": "red", "title": {"tag": "plain_text", "content": "❌ 懂火应收结算同步失败"}},
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
# 【登录分工说明】与 sync_chuku_to_bitable.py 相同：
#   导出按钮是纯前端交互，必须用 Playwright 真实浏览器点击；
#   因此在浏览器里内联重写登录（ddddocr 识别验证码 → 填表单 → 点 #laysubmit），
#   账号密码仍从 .env 的 DH_USERNAME / DH_PASSWORD 读取。

def download_yingshou_xls(headless: bool = False) -> Path:
    """
    用 Playwright 自动登录懂火 → 应收结算页 → 筛选里清空开始日期 → 查询
    → 点系统导出按钮（POST excelbiao/x_jiesuan）→ 下载 xls。返回文件路径。
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

        # --- 登录（重试最多 10 次）---
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

        # --- 打开应收结算页（iframe 网格页可直接访问，与出库记录同模式）---
        CURRENT_STEP = "打开应收结算页"
        page.goto(DONGHUO_YS_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(4000)
        log(f"[懂火] ✅ 已到达应收结算页: {page.url}")

        # --- 点"筛选"按钮 ---
        CURRENT_STEP = "打开筛选面板并清空开始日期"
        page.get_by_text("筛选", exact=True).first.click()
        page.wait_for_timeout(1500)
        log("[懂火] ✅ 筛选面板已打开")

        # --- JS 清空开始日期（结束日期保持系统默认值不变）---
        result = page.evaluate("""() => {
            const startEl = document.getElementById('start_time');
            if (startEl) {
                startEl.removeAttribute('readonly');
                startEl.value = '';
                startEl.setAttribute('readonly', 'readonly');
                startEl.dispatchEvent(new Event('change', {bubbles: true}));
                startEl.dispatchEvent(new Event('blur', {bubbles: true}));
            }
            return {start: startEl ? startEl.value : '(无此输入框)'};
        }""")
        log(f"[懂火] ✅ 开始日期已清空: {result}")

        # --- 点筛选弹窗底部的"查询"按钮应用筛选 ---
        clicked = False
        try:
            page.get_by_text("查询", exact=True).first.click(timeout=5000)
            clicked = True
        except Exception:
            pass
        if not clicked:
            # 兜底：按 chuku 脚本的位置启发式（弹窗底部第一个非 close 按钮）
            btns = page.evaluate("""
            () => Array.from(document.querySelectorAll('button.layui-btn'))
              .map((b, i) => {
                const r = b.getBoundingClientRect();
                return {idx: i, cls: b.className, y: Math.round(r.top),
                        rect: [Math.round(r.left), Math.round(r.top), Math.round(r.width), Math.round(r.height)]};
              })
              .filter(b => b.rect[2] > 0 && b.rect[3] > 0)
            """)
            popup_btns = [b for b in btns if b['y'] > 500]
            confirm = next((b for b in popup_btns if 'close' not in b['cls']), None)
            if confirm is None:
                confirm = popup_btns[0] if popup_btns else (btns[0] if btns else None)
            if confirm is None:
                raise RuntimeError("找不到筛选弹窗的查询按钮")
            page.locator("button.layui-btn").nth(confirm['idx']).click()
        page.wait_for_timeout(2500)
        log("[懂火] ✅ 筛选已应用（开始日期为空=不限起始）")

        # --- 点"导出"按钮，等待下载 ---
        CURRENT_STEP = "点击系统导出按钮下载"
        log("[懂火] 开始导出（全量数据，可能十几秒到几十秒）...")
        export_btn = page.get_by_text("导出", exact=True).first
        with page.expect_download(timeout=300000) as dl_info:
            export_btn.click()
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
    同时落一份 UTF-8-SIG CSV 到 csv_backup/（上传成功后由调用方删除）。
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
    # 列名别名（懂火导出名 → 多维表字段名）
    df = df.rename(columns=COLUMN_ALIASES)
    log(f"[解析] {xls_path.name}: {len(df)} 行 × {len(df.columns)} 列")

    # 落 CSV 备份
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = CSV_DIR / f"yingshou_export_{now}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    log(f"[备份] CSV 已保存: {csv_path}")
    return df, csv_path


# ============ 飞书多维表：清空 + 写入 ============

def bitable_list_all_records(token: str) -> list:
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


def bitable_batch_delete(token: str, record_ids: list):
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
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
            f = float(s)
            return int(f) if f == int(f) else f
        except ValueError:
            return None
    if ftype == "datetime":
        s = str(raw).strip()
        if not s:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
            try:
                dt = datetime.datetime.strptime(s, fmt)
                return int(dt.timestamp() * 1000)
            except ValueError:
                continue
        return None
    return None


def df_to_bitable_records(df) -> list:
    """把 DataFrame 转成飞书 batch_create 需要的 [{"fields": {...}}, ...] 列表"""
    records = []
    for _, row in df.iterrows():
        fields = {}
        for col, ftype in BITABLE_FIELD_TYPES.items():
            if col not in df.columns:
                continue
            val = _convert_value(row.get(col), ftype)
            if val is not None:
                fields[col] = val
        records.append({"fields": fields})
    return records


def bitable_batch_create(token: str, records: list):
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
    ap.add_argument("--skip-download", action="store_true", help="跳过下载，直接用 downloads/ 下最新的 xls")
    ap.add_argument("--skip-clear", action="store_true", help="跳过清空旧表（仅追加，不推荐，仅调试用）")
    ap.add_argument("--skip-upload", action="store_true", help="跳过写入多维表（只下载+转 CSV）")
    ap.add_argument("--download-only", action="store_true", help="等价 --skip-upload --skip-clear")
    ap.add_argument("--dry-run", action="store_true", help="下载+转 CSV+打印前几行，不碰飞书")
    ap.add_argument("--no-notify", action="store_true", help="不发送飞书通知（默认同步完成会通知洪）")
    args = ap.parse_args()

    global CURRENT_STEP
    t0 = time.time()
    log("==== 懂火应收结算 → 飞书多维表 同步工作流 启动 ====")

    # ---- Step 1: 下载 ----
    if args.skip_download:
        CURRENT_STEP = "读取本地导出文件"
        xls_files = sorted(DOWNLOAD_DIR.glob("*.xls"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not xls_files:
            raise RuntimeError("--skip-download 但 downloads/ 下没有 xls 文件")
        xls_path = xls_files[0]
        log(f"[跳过下载] 使用已有文件: {xls_path}")
    else:
        xls_path = download_yingshou_xls(headless=args.headless)

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
    CURRENT_STEP = "获取飞书凭证"
    token = feishu_token()
    log("[飞书] ✅ tenant_access_token 已获取")

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
        records = df_to_bitable_records(df)
        # 先小批量试写，验证字段类型/名称没问题
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
        total_written = 0

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
