#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火 5 合 1 数据汇总同步工作流（一次登录 · 五部分合并执行 · 一张汇总通知卡）

把 懂火出库同步/ 下 5 个独立脚本的流程合并成一次运行：

  部分     模块               取数方式                        目标表(数据汇总 2026)   同步语义
  ─────────────────────────────────────────────────────────────────────────────────────────
  ① 出库   出库记录           Playwright UI 导出(2026-01-01起)  tblolnj06JZkYNiU      全量替换
  ② 订单   订单明细汇总        Playwright UI 导出(出库状态=全部)  tblcEZoQatk7lCAO      全量替换
  ③ 应收   应收结算           Playwright UI 导出(清空开始日期)   tblpjne9dIuif5HD      全量替换
  ④ 往来   收付款流水          getlist API(确认状态=全部)        tblbS1dPaDVL3GY8      全量替换
  ⑤ 客户   客户管理 CRM        getlist API(筛选全空)            tblCE7zIWs804RR5      增量+已删除标记

【一次登录的实现】
  - Playwright 真 Chrome 登录一次（ddddocr 识别验证码）
  - 登录成功后立即从 browser context 提取 cookies（登录态固定，后续导航不影响）
  → 同一浏览器依次完成 ①②③ 三个 UI 导出（导出按钮是纯前端交互，requests 点不了）
  → cookies 注入 requests.Session 调 ④⑤ getlist 接口（与页面表格同一数据源，非扒网页）
  - 若 cookie 注入不被接受或未走浏览器（--skip-download），④⑤ 自动退回 donghuo_login.py
  - 客户模块「导出」按钮被部署方禁用（khdown 首行 return false + 服务端返回"没有权限"，
    2026-09-06 实测），往来模块无独立导出页，两者均走 getlist 接口

【失败隔离】
  每个部分独立 try/except：一部分失败不阻断其他部分。
  成功部分上传后自动删 CSV；失败部分 CSV 保留在 csv_backup/ 备查。
  结束发一张汇总卡片给洪：全成功绿 / 部分失败橙 / 全失败红，附失败详情与客户删除名单。

使用：python sync_all_to_bitable.py [--headless] [--skip-download] [--dry-run] [--no-notify]
                                    [--only chuku,dingdan,yingshou,wanglai,kehu]
凭据：仓库根目录 .env 里的 DH_USERNAME / DH_PASSWORD + FEISHU_APP_ID / FEISHU_APP_SECRET
（各模块的字段映射/筛选逻辑与 5 个单脚本完全一致，单脚本仍可独立运行）
"""
import sys, os, json, time, re, math, datetime, argparse, traceback, requests
from pathlib import Path
from collections import defaultdict

# ===== stdout/stderr 编码双保险（Windows subprocess 里 print emoji 会崩）=====
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

from dotenv import load_dotenv
load_dotenv(Path(__file__).resolve().parent.parent / ".env", override=False)  # .env 统一放仓库根目录

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from donghuo_login import login_donghuo  # noqa: E402  cookie 兜底登录（--skip-download 时）

# ===== 固定配置 =====
DONGHUO_BASE   = "https://erpa.donghuo.vip"
DONGHUO_LOGIN  = f"{DONGHUO_BASE}/view/admin/v_login"
DONGHUO_URLS = {
    "chuku":    f"{DONGHUO_BASE}/view/admin/xiaoshou/v_xjlall",     # 出库记录
    "dingdan":  f"{DONGHUO_BASE}/view/admin/xiaoshou/v_xmxhz",      # 订单明细汇总
    "yingshou": f"{DONGHUO_BASE}/view/admin/caiwu/v_x_jiesuan",     # 应收结算
    "wanglai":  f"{DONGHUO_BASE}/model/admin/caiwu/m_liushui/getlist",  # 往来流水 API
    "kehu":     f"{DONGHUO_BASE}/model/admin/crm/m_kehu/getlist",   # 客户管理 API
}

BITABLE_APP_TOKEN = "VahHb3YDBaBTwTsCjeAcaAhhnHc"   # 数据汇总（2026）
TABLES = {
    "chuku":    "tblolnj06JZkYNiU",
    "dingdan":  "tblcEZoQatk7lCAO",
    "yingshou": "tblpjne9dIuif5HD",
    "wanglai":  "tblbS1dPaDVL3GY8",
    "kehu":     "tblCE7zIWs804RR5",
}
PART_NAMES = {
    "chuku": "① 出库记录", "dingdan": "② 订单明细", "yingshou": "③ 应收结算",
    "wanglai": "④ 往来流水", "kehu": "⑤ 客户管理",
}

EXPORT_START_DATE = "2026-01-01"   # ① 出库筛选起始日期
PAGE_SIZE = 300                    # ④⑤ getlist 单页上限
MAX_PAGES = 50                     # ④⑤ 分页保险丝
BATCH_SIZE = 500                   # 飞书 bitable batch_* 上限
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
FEISHU_NOTIFY_UNION_ID = "on_b09bcbf3e74f5d423900aa9b2f00eb63"   # 洪

DOWNLOAD_DIR = Path(__file__).parent / "downloads" / "all"   # ①②③ xls 统一落这里
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
CSV_DIR = Path(__file__).parent / "csv_backup"
CSV_DIR.mkdir(exist_ok=True)

# ===== 各部分字段映射（与单脚本一致）=====

# ① 出库：懂火导出列名 == 多维表字段名（2026-09-05 硬编码映射）
CHUKU_FIELD_TYPES = {
    "所属公司": "text", "出库日期": "datetime", "销售人": "text", "客户名称": "text",
    "订单号": "text", "品名": "text", "规格": "text", "材质": "text", "产地": "text",
    "等级": "text", "件(张)数": "number", "采购重量(吨)": "number", "重量(吨)": "number",
    "挂牌价": "number", "销售单价": "number", "销售税率": "number", "销售金额": "number",
    "未开发票": "number", "供应商": "text", "采购单价": "number", "采购税率": "number",
    "采购金额": "number", "费用金额": "number", "利润": "number", "市场盈利": "number",
    "仓库": "text", "库位号": "text", "捆包号": "text", "合同号": "text", "车船号": "text",
    "提单号": "text", "备注": "text",
}

# ② 订单：动态探测字段 + 列名别名（表字段「订单 号」带空格是公式不可写，可写文本叫「订单 号」）
DINGDAN_ALIASES = {"订单号": "订单 号", "备注": "注释"}
FIELD_TYPE_CODE_MAP = {   # 飞书字段类型码 → 写入策略；其余（User/Lookup/Formula/Url/Attachment/自动）不可写
    1: "text", 2: "number", 3: "text", 4: "multiselect", 5: "datetime", 7: "checkbox", 13: "text",
}

# ③ 应收：硬编码映射 + df 层面列名 rename（2026-09-06 探测）
YINGSHOU_FIELD_TYPES = {
    "订单 号": "text", "日期": "datetime", "发货状态": "text", "所属公司": "text",
    "销售人": "text", "客户名称": "text", "实发重量": "number", "实发金额": "number",
    "销售费用": "number", "其它款项": "number", "已结金额": "number", "未结金额": "number",
}
YINGSHOU_ALIASES = {"订单号": "订单 号", "销售状态": "发货状态"}

# ④ 往来：硬编码映射（2026-09-06 探测）
WANGLAI_FIELD_TYPES = {
    "所属公司": "text", "日期": "datetime", "我方帐户": "text", "交易对方": "text",
    "交易类型": "text", "科目名称": "text", "结算方式": "text", "金额": "number",
    "状态": "text", "结算对方": "text", "订单号": "text", "销售人": "text",
    "备注说明": "text", "提交人": "text", "确认人": "text",
}

# ⑤ 客户：主键=客户名称；跟踪字段；仅新增写创建时间；已删除标记
KEHU_TRACKED_FIELDS = [
    "客户类型", "所属人", "联系人", "联系人职位", "固定电话", "移动电话",
    "邮箱地址", "所属省份", "联系地址", "主营产品", "采购产品", "备注",
]
DELETED_MARK = "已删除"
_JUNK_RE = re.compile(r"^\d{1,2}$")   # '1'/'0'/'00' 等占位垃圾值


# ============ 工具函数 ============

def log(msg: str):
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def env(name: str) -> str:
    return os.environ.get(name, "").strip()


def norm_text(v) -> str:
    """文本归一化：None/空白/全角空格 → ''（⑤ 客户用）"""
    if v is None:
        return ""
    return str(v).replace("\u3000", " ").strip()


def is_meaningful(v) -> bool:
    """值是否有意义：非空且非 1~2 位纯数字占位垃圾（⑤ 客户用）"""
    s = norm_text(v)
    return bool(s) and not _JUNK_RE.match(s)


def to_datetime_ms(raw):
    """'2026-09-02 13:18:43' → 毫秒时间戳（⑤ 客户创建时间用）"""
    s = norm_text(raw)
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
        try:
            return int(datetime.datetime.strptime(s, fmt).timestamp() * 1000)
        except ValueError:
            continue
    return None


def feishu_token() -> str:
    """获取飞书 tenant_access_token（每次现取，2 小时有效期）"""
    app_id = os.environ.get("FEISHU_APP_ID") or os.environ.get("FEISHU_NOTIFY_APP_ID", "").strip()
    app_secret = os.environ.get("FEISHU_APP_SECRET") or os.environ.get("FEISHU_NOTIFY_APP_SECRET", "").strip()
    if not app_id or not app_secret:
        raise RuntimeError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET 环境变量")
    r = requests.post(
        f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret}, timeout=15,
    )
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换 tenant_access_token 失败: {data}")
    return data["tenant_access_token"]


def feishu_send_card(union_id: str, card: dict, token: str):
    """通过飞书机器人给指定用户发卡片消息"""
    url = f"{FEISHU_OPEN_BASE}/im/v1/messages?receive_id_type=union_id"
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"receive_id": union_id, "msg_type": "interactive",
            "content": json.dumps(card, ensure_ascii=False)}
    r = requests.post(url, headers=h, json=body, timeout=15)
    data = r.json()
    if data.get("code") != 0:
        log(f"[飞书通知] ❌ 发送失败: code={data.get('code')} msg={data.get('msg')}")
    else:
        log(f"[飞书通知] ✅ 汇总卡片已发送给 {union_id}")


# ============ ①②③ Playwright：登录一次 + 三个 UI 导出 ============

def browser_login(page, ocr, username: str, password: str):
    """在给定 page 上完成懂火登录（ddddocr 识别验证码，最多 10 次）。失败 raise。"""
    for attempt in range(1, 11):
        log(f"[登录] 尝试 {attempt}/10 ...")
        page.goto(DONGHUO_LOGIN, wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(1200)
        page.locator("input.layui-input:not(#captcha):not([type='password'])").first.fill(username)
        page.locator("#u_pass").fill(password)
        captcha_img = page.locator("img[src*='captcha']").first
        code = ocr.classification(captcha_img.screenshot()).strip()
        log(f"[登录] 验证码识别: '{code}'")
        page.locator("#captcha").fill(code)
        page.wait_for_timeout(200)
        page.locator("#laysubmit").click()
        page.wait_for_timeout(1200)
        if "/v_login" not in page.url:
            log("[登录] ✅ 登录成功")
            return
    raise RuntimeError("懂火登录失败，已达最大重试次数（10 次）")


def session_from_browser(ctx) -> requests.Session:
    """从浏览器 context 提取 cookies → 构造带同款 UA 的 requests.Session（供 ④⑤ API 用）"""
    s = requests.Session()
    s.headers.update({
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Accept-Language": "zh-CN,zh;q=0.9",
        "X-Requested-With": "XMLHttpRequest",
    })
    for c in ctx.cookies():
        s.cookies.set(c["name"], c["value"],
                      domain=c.get("domain", "").lstrip("."), path=c.get("path", "/"))
    return s


def _export_via_download(page, owner, target_dir: Path, prefix: str, timeout_ms: int) -> Path:
    """点「导出」→ 等 download → 存为 {prefix}__原名.xls"""
    log(f"[导出] 点击导出按钮（可能几秒到几十秒）...")
    export_loc = owner.get_by_text("导出", exact=True).first
    with page.expect_download(timeout=timeout_ms) as dl_info:
        export_loc.click()
    dl = dl_info.value
    target = target_dir / f"{prefix}__{dl.suggested_filename}"
    dl.save_as(str(target))
    log(f"[导出] ✅ 下载完成: {target.name} ({target.stat().st_size / 1024:.1f} KB)")
    return target


def export_chuku(page) -> Path:
    """① 出库记录：筛选 2026-01-01 起 → 系统导出按钮（与单脚本逻辑一致）"""
    page.goto(DONGHUO_URLS["chuku"], wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(4000)
    page.get_by_text("筛选", exact=True).first.click()
    page.wait_for_timeout(1500)
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    result = page.evaluate("""([startVal, todayStr]) => {
        const startEl = document.getElementById('start_time');
        const endEl = document.getElementById('end_time');
        if (startEl) {
            startEl.removeAttribute('readonly');
            startEl.value = startVal;
            startEl.setAttribute('readonly', 'readonly');
            startEl.dispatchEvent(new Event('change', {bubbles: true}));
            startEl.dispatchEvent(new Event('blur', {bubbles: true}));
        }
        if (endEl) {
            endEl.removeAttribute('readonly');
            endEl.value = todayStr;
            endEl.setAttribute('readonly', 'readonly');
            endEl.dispatchEvent(new Event('change', {bubbles: true}));
            endEl.dispatchEvent(new Event('blur', {bubbles: true}));
        }
        return {start: startEl?.value, end: endEl?.value};
    }""", [EXPORT_START_DATE, today_str])
    log(f"[① 出库] ✅ 日期筛选已设: {result}")
    # 弹窗底部第一个非 close 的 layui-btn = 查询（原脚本启发式）
    btns = page.evaluate("""
    () => Array.from(document.querySelectorAll('button.layui-btn'))
      .map((b, i) => {
        const r = b.getBoundingClientRect();
        return {idx: i, cls: b.className, y: Math.round(r.top),
                rect: [Math.round(r.left), Math.round(r.top), Math.round(r.width), Math.round(r.height)]};
      })
      .filter(b => b.rect[2] > 0 && b.rect[3] > 0)
    """)
    popup_btns = [b for b in btns if b["y"] > 500]
    confirm = next((b for b in popup_btns if "close" not in b["cls"]), None)
    if confirm is None:
        confirm = popup_btns[0] if popup_btns else (btns[0] if btns else None)
    if confirm is None:
        raise RuntimeError("① 出库：筛选弹窗未找到查询按钮")
    page.locator("button.layui-btn").nth(confirm["idx"]).click()
    page.wait_for_timeout(2500)
    return _export_via_download(page, page, DOWNLOAD_DIR, "chuku", 120000)


# ② 订单需要的 JS 工具（与单脚本一致）
JS_SET_SELECT = """
([labelText, wantText]) => {
    const selects = Array.from(document.querySelectorAll('select'));
    const hasOpt = (s, t) => Array.from(s.options).some(o => o.textContent.trim() === t);
    const report = selects.map(s => ({
        id: s.id, name: s.name,
        opts: Array.from(s.options).map(o => o.textContent.trim()).slice(0, 10),
    }));
    let hit = null;
    for (const s of selects) {
        if (!hasOpt(s, wantText)) continue;
        const item = s.closest('.layui-form-item') || s.parentElement;
        const lb = item ? item.querySelector('.layui-form-label, label, .layui-inline') : null;
        if (lb && lb.textContent.trim().includes(labelText)) { hit = s; break; }
    }
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
    return {ok: true, id: hit.id, name: hit.name, value: hit.value,
            text: hit.selectedOptions[0] ? hit.selectedOptions[0].textContent.trim() : ''};
}
"""

JS_PICK_CONFIRM = """
() => {
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
        cands.sort((a, b) => ((b.y > 500) - (a.y > 500)) || (a.idx - b.idx));
        return cands[0];
    }
    const popup = all.filter(b => b.y > 500 && b.cls && !b.cls.includes('close'));
    return popup[0] || all[0] || null;
}
"""

JS_CLICK_VISIBLE_OPTION = """
(wantText) => {
    const dds = Array.from(document.querySelectorAll('.layui-form-select dl dd'));
    const visible = dds.filter(d => d.offsetParent !== null);
    const t = visible.find(d => d.textContent.trim() === wantText);
    if (t) { t.click(); return true; }
    return false;
}
"""


def _find_owner_for_text(page, text: str, tries: int = 6):
    """在主页面与所有 iframe 中查找精确文本元素，返回承载它的 Page/Frame（② 订单用）"""
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
        page.wait_for_timeout(1000)
    return None


def export_dingdan(page) -> Path:
    """② 订单明细汇总：筛选 出库状态=全部 → 系统导出按钮（与单脚本逻辑一致）"""
    page.goto(DONGHUO_URLS["dingdan"], wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(4000)
    owner = _find_owner_for_text(page, "筛选")
    if owner is None:
        raise RuntimeError("② 订单：页面未找到「筛选」元素（页面结构可能变化）")
    owner.get_by_text("筛选", exact=True).first.click()
    page.wait_for_timeout(1500)
    res = owner.evaluate(JS_SET_SELECT, ["出库状态", "全部"])
    if res.get("ok"):
        log(f"[② 订单] ✅ 出库状态已设为 '{res.get('text')}'")
    else:
        log(f"[② 订单] ⚠️ 原生 select 未命中，尝试 layui 渲染层兜底")
        try:
            label_loc = owner.get_by_text("出库状态", exact=True).first
            parent = label_loc.locator("xpath=..")
            dd = parent.locator(".layui-form-select").first
            if dd.count() == 0:
                dd = owner.locator(".layui-form-select").last
            dd.locator("input.layui-input, .layui-select-title, .layui-edge").first.click()
            page.wait_for_timeout(600)
            clicked = owner.evaluate(JS_CLICK_VISIBLE_OPTION, "全部")
            page.wait_for_timeout(300)
            if not clicked:
                raise RuntimeError("渲染层未见可见的'全部'选项")
            log("[② 订单] ✅ 已通过 layui 渲染层点选 '全部'")
        except Exception as e:
            raise RuntimeError(f"② 订单：设置 出库状态='全部' 失败: {e}") from e
    btn = owner.evaluate(JS_PICK_CONFIRM)
    if btn is None:
        raise RuntimeError("② 订单：筛选面板未找到确认按钮")
    owner.locator("button.layui-btn, a.layui-btn").nth(btn["idx"]).click()
    page.wait_for_timeout(2500)
    return _export_via_download(page, owner, DOWNLOAD_DIR, "dingdan", 180000)


def export_yingshou(page) -> Path:
    """③ 应收结算：筛选清空开始日期 → 系统导出按钮（与单脚本逻辑一致）"""
    page.goto(DONGHUO_URLS["yingshou"], wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(4000)
    page.get_by_text("筛选", exact=True).first.click()
    page.wait_for_timeout(1500)
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
    log(f"[③ 应收] ✅ 开始日期已清空: {result}")
    clicked = False
    try:
        page.get_by_text("查询", exact=True).first.click(timeout=5000)
        clicked = True
    except Exception:
        pass
    if not clicked:
        btns = page.evaluate("""
        () => Array.from(document.querySelectorAll('button.layui-btn'))
          .map((b, i) => {
            const r = b.getBoundingClientRect();
            return {idx: i, cls: b.className, y: Math.round(r.top),
                    rect: [Math.round(r.left), Math.round(r.top), Math.round(r.width), Math.round(r.height)]};
          })
          .filter(b => b.rect[2] > 0 && b.rect[3] > 0)
        """)
        popup_btns = [b for b in btns if b["y"] > 500]
        confirm = next((b for b in popup_btns if "close" not in b["cls"]), None)
        if confirm is None:
            confirm = popup_btns[0] if popup_btns else (btns[0] if btns else None)
        if confirm is None:
            raise RuntimeError("③ 应收：筛选弹窗未找到查询按钮")
        page.locator("button.layui-btn").nth(confirm["idx"]).click()
    page.wait_for_timeout(2500)
    return _export_via_download(page, page, DOWNLOAD_DIR, "yingshou", 300000)


# ============ ④⑤ requests：getlist 分页拉取（共用同一 session）============

def _getlist_paged(session, url: str, referer: str, tag: str, max_pages: int = MAX_PAGES) -> list:
    """getlist 分页拉全量：root=行列表、rtotal=总数；不带过滤参数=全量"""
    headers = {"X-Requested-With": "XMLHttpRequest", "Referer": referer}
    all_rows, rtotal = [], None
    for page_no in range(1, max_pages + 1):
        r = session.post(url, data={"page": page_no, "limit": PAGE_SIZE},
                         headers=headers, timeout=30)
        try:
            data = r.json()
        except ValueError:
            raise RuntimeError(f"{tag}：getlist 第 {page_no} 页返回非 JSON（登录态可能失效）: {r.text[:150]}")
        root = data.get("root") or []
        if rtotal is None:
            rtotal = int(data.get("rtotal") or 0)
            log(f"[{tag}] 总数 rtotal={rtotal}，每页 {PAGE_SIZE}，预计 {math.ceil(rtotal / PAGE_SIZE)} 页")
        all_rows.extend(root)
        if page_no % 5 == 0 or not root or (rtotal and len(all_rows) >= rtotal):
            log(f"[{tag}] 第 {page_no} 页: 累计 {len(all_rows)}/{rtotal}")
        if not root or (rtotal and len(all_rows) >= rtotal):
            break
        time.sleep(0.3)
    if rtotal and len(all_rows) < rtotal:
        log(f"[{tag}] ⚠️ 仅拉到 {len(all_rows)}/{rtotal} 条")
    return all_rows


def fetch_wanglai(session) -> list:
    """④ 往来流水：确认状态=全部（不传 zhuantai）"""
    rows = _getlist_paged(session, DONGHUO_URLS["wanglai"],
                          f"{DONGHUO_BASE}/view/admin/caiwu/v_x_jiesuan", "④ 往来")
    log(f"[④ 往来] ✅ 拉取完成，共 {len(rows)} 条")
    return rows


def fetch_kehu(session) -> list:
    """⑤ 客户管理：筛选清空所有条件（不带任何过滤参数）"""
    rows = _getlist_paged(session, DONGHUO_URLS["kehu"],
                          f"{DONGHUO_BASE}/view/admin/crm/v_kehu", "⑤ 客户")
    # 重名预警
    seen = defaultdict(int)
    for row in rows:
        n = norm_text(row.get("客户名称"))
        if n:
            seen[n] += 1
    dups = {n: c for n, c in seen.items() if c > 1}
    if dups:
        log(f"[⑤ 客户] ⚠️ 名称重名 {len(dups)} 组（增量合并时按数据丰富度处理）")
    log(f"[⑤ 客户] ✅ 拉取完成，共 {len(rows)} 行 / {len(seen)} 个唯一名称")
    return rows


# ============ 解析 / CSV 备份 ============

def xls_to_df(xls_path: Path, csv_prefix: str):
    """懂火 HTML 伪 xls → DataFrame + UTF-8-SIG CSV 备份。返回 (df, csv_path)。
    ③ 应收在 CSV 阶段就完成列名 rename（与单脚本一致）。"""
    import pandas as pd
    df_raw = pd.read_html(str(xls_path))[0]
    header = list(df_raw.iloc[0])
    df = df_raw.iloc[1:].reset_index(drop=True)
    df.columns = header
    df = df.dropna(how="all")
    if csv_prefix == "yingshou":
        df = df.rename(columns=YINGSHOU_ALIASES)
    log(f"[解析] {xls_path.name}: {len(df)} 行 × {len(df.columns)} 列")
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = CSV_DIR / f"{csv_prefix}_export_{now}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    log(f"[备份] CSV 已保存: {csv_path.name}")
    return df, csv_path


def rows_to_csv(rows: list, csv_prefix: str) -> Path:
    """④⑤ API 原始 dict 列表 → CSV 备份"""
    import pandas as pd
    df = pd.DataFrame(rows).dropna(how="all")
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = CSV_DIR / f"{csv_prefix}_export_{now}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    log(f"[备份] CSV 已保存: {csv_path.name}（{len(df)} 行 × {len(df.columns)} 列）")
    return csv_path


# ============ 飞书多维表通用（table_id 参数化）============

def _records_url(table_id: str) -> str:
    return f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{table_id}/records"


def bitable_list_records(token: str, table_id: str) -> list:
    """读取全部记录 [{"record_id", "fields"}, ...]（⚠️ search 接口 page_token 不推进，必须 GET list）"""
    h = {"Authorization": f"Bearer {token}"}
    out, seen, page_token = [], set(), None
    while True:
        qs = f"page_size={BATCH_SIZE}" + (f"&page_token={page_token}" if page_token else "")
        r = requests.get(f"{_records_url(table_id)}?{qs}", headers=h, timeout=30)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"list records 失败({table_id}): {data}")
        d = data.get("data") or {}
        items = d.get("items") or []
        if not items:
            break
        dup = 0
        for it in items:
            rid = it.get("record_id")
            if rid and rid not in seen:
                out.append({"record_id": rid, "fields": it.get("fields") or {}})
                seen.add(rid)
            else:
                dup += 1
        if dup == len(items):
            log(f"[飞书] ⚠️ page_token 未推进，停止；累计 {len(out)} 条")
            break
        if not d.get("has_more") or not d.get("page_token"):
            break
        page_token = d.get("page_token")
    return out


def bitable_batch_delete(token: str, table_id: str, record_ids: list):
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{_records_url(table_id)}/batch_delete?ignore_consistency_check=true"
    total = len(record_ids)
    for i in range(0, total, BATCH_SIZE):
        batch = record_ids[i:i + BATCH_SIZE]
        r = requests.post(url, headers=h, json={"records": batch}, timeout=30)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_delete 失败: {data.get('msg')}")
        if (i // BATCH_SIZE + 1) % 10 == 0 or i + BATCH_SIZE >= total:
            log(f"[飞书] delete: 批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE}")


def bitable_batch_create(token: str, table_id: str, records: list) -> int:
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{_records_url(table_id)}/batch_create?ignore_consistency_check=true"
    total, created = len(records), 0
    for i in range(0, total, BATCH_SIZE):
        batch = records[i:i + BATCH_SIZE]
        r = requests.post(url, headers=h, json={"records": batch}, timeout=60)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_create 失败 批 {i // BATCH_SIZE + 1}: {data.get('code')} {data.get('msg')} 样本={str(data)[:300]}")
        created += len((data.get("data") or {}).get("records") or [])
        log(f"[飞书] create: {len(batch)} 条 (批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE})，累计 {created}")
    return created


def bitable_batch_update(token: str, table_id: str, updates: list) -> int:
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{_records_url(table_id)}/batch_update?ignore_consistency_check=true"
    total, done = len(updates), 0
    for i in range(0, total, BATCH_SIZE):
        batch = updates[i:i + BATCH_SIZE]
        r = requests.post(url, headers=h, json={"records": batch}, timeout=60)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_update 失败 批 {i // BATCH_SIZE + 1}: {data.get('code')} {data.get('msg')}")
        done += len((data.get("data") or {}).get("records") or [])
        if (i // BATCH_SIZE + 1) % 10 == 0 or i + BATCH_SIZE >= total:
            log(f"[飞书] update: 批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE}，累计 {done}")
    return done


def bitable_get_field_types(token: str, table_id: str) -> dict:
    """动态探测字段类型（② 订单表用），返回 {字段名: 写入策略}"""
    h = {"Authorization": f"Bearer {token}"}
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{table_id}/fields"
    out, page_token = {}, None
    while True:
        qs = "page_size=100" + (f"&page_token={page_token}" if page_token else "")
        r = requests.get(f"{url}?{qs}", headers=h, timeout=30)
        d = r.json()
        if d.get("code") != 0:
            raise RuntimeError(f"获取字段列表失败: {d}")
        dd = d.get("data") or {}
        for f in dd.get("items") or []:
            out[f["field_name"]] = FIELD_TYPE_CODE_MAP.get(f["type"], "unsupported")
        if not dd.get("has_more") or not dd.get("page_token"):
            break
        page_token = dd.get("page_token")
    return out


def convert_value(raw, ftype: str):
    """DataFrame 单元格值 → 飞书字段值（② 订单增强版：千分位/多选/复选/时间戳兜底，兼容其余部分）"""
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
            return int(raw) if raw > 10**12 else int(raw * 1000)
        s = str(raw).strip()
        if not s:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                    "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
            try:
                return int(datetime.datetime.strptime(s, fmt).timestamp() * 1000)
            except ValueError:
                continue
        return None
    if ftype == "multiselect":
        s = str(raw).strip()
        if not s:
            return None
        parts = [p.strip() for p in re.split(r"[,，、;；]", s)]
        return [p for p in parts if p] or None
    if ftype == "checkbox":
        s = str(raw).strip().lower()
        if s in ("是", "true", "1", "y", "yes", "√", "已"):
            return True
        if s in ("否", "false", "0", "n", "no", ""):
            return False
        return None
    return None


def df_to_records(df, field_types: dict, aliases: dict = None) -> list:
    """DataFrame → batch_create 记录列表。aliases: 导出列名→表字段名（仅 ② 订单需要）"""
    aliases = aliases or {}
    field_to_col = {f: f for f in field_types}
    for csv_col, field_name in aliases.items():
        if field_name in field_types and field_name not in df.columns and csv_col in df.columns:
            field_to_col[field_name] = csv_col
    for col in df.columns:
        if col not in field_types and col not in aliases:
            log(f"[飞书] ⚠️ 列 '{col}' 在表字段里不存在，跳过")
    records = []
    for _, row in df.iterrows():
        fields = {}
        for field_name, ftype in field_types.items():
            if ftype == "unsupported":
                continue
            src = field_to_col.get(field_name)
            if src is None or src not in df.columns:
                continue
            val = convert_value(row.get(src), ftype)
            if val is not None:
                fields[field_name] = val
        records.append({"fields": fields})
    return records


def rows_to_records(rows: list, field_types: dict) -> list:
    """API 原始 dict 列表 → batch_create 记录列表（④ 往来）"""
    records = []
    for row in rows:
        fields = {}
        for col, ftype in field_types.items():
            val = convert_value(row.get(col), ftype)
            if val is not None:
                fields[col] = val
        records.append({"fields": fields})
    return records


# ============ ⑤ 客户：增量计划（与单脚本一致）============

def merge_duplicate_rows(rows: list) -> dict:
    """懂火同名客户多行 → 合并：基础行=有效值最多行；空字段用其他行有效值补；新增时间取最早"""
    groups = defaultdict(list)
    for row in rows:
        name = norm_text(row.get("客户名称"))
        if name:
            groups[name].append(row)
    merged = {}
    for name, grp in groups.items():
        if len(grp) == 1:
            merged[name] = grp[0]
            continue

        def score(r):
            return sum(1 for c in KEHU_TRACKED_FIELDS if is_meaningful(r.get(c)))

        base = max(grp, key=score)
        m = dict(base)
        for c in KEHU_TRACKED_FIELDS:
            if not is_meaningful(m.get(c)):
                for r in grp:
                    if r is not base and is_meaningful(r.get(c)):
                        m[c] = r[c]
                        break
        ts_list = [t for t in (to_datetime_ms(r.get("新增时间")) for r in grp) if t]
        if ts_list:
            m["新增时间"] = datetime.datetime.fromtimestamp(min(ts_list) / 1000).strftime("%Y-%m-%d %H:%M:%S")
        log(f"[⑤ 客户] 重名合并: {name}（{len(grp)} 行 → 基础行 id={base.get('id')}）")
        merged[name] = m
    return merged


def kehu_build_new_fields(row: dict) -> dict:
    """懂火客户行 → 新增记录 fields（垃圾值不写入）"""
    fields = {"客户名称": norm_text(row.get("客户名称"))}
    for col in KEHU_TRACKED_FIELDS:
        v = norm_text(row.get(col))
        if is_meaningful(v):
            fields[col] = v
    ts = to_datetime_ms(row.get("新增时间"))
    if ts:
        fields["创建时间"] = ts
    return fields


def kehu_diff(row: dict, feishu_fields: dict) -> dict:
    """比对跟踪字段，返回需要更新的 fields（只写有意义且不同的值）"""
    changed = {}
    for col in KEHU_TRACKED_FIELDS:
        new_v = norm_text(row.get(col))
        old_v = norm_text(feishu_fields.get(col))
        if is_meaningful(new_v) and new_v != old_v:
            changed[col] = new_v
    return changed


def kehu_plan(donghuo_rows: list, feishu_records: list) -> dict:
    """以客户名称为主键生成增量计划（新增/更新/恢复/标记删除/无变化）"""
    dh_map = merge_duplicate_rows(donghuo_rows)
    fs_map = defaultdict(list)
    no_name = 0
    for rec in feishu_records:
        name = norm_text((rec["fields"] or {}).get("客户名称"))
        if not name:
            no_name += 1
            continue
        fs_map[name].append(rec)
    if no_name:
        log(f"[⑤ 客户] ⚠️ 飞书 {no_name} 条记录无客户名称，跳过比对")

    to_create, to_update, to_restore = [], [], []
    unchanged = 0
    for name, row in dh_map.items():
        matches = fs_map.get(name)
        if not matches:
            to_create.append(kehu_build_new_fields(row))
            continue
        for rec in matches:
            changed = kehu_diff(row, rec["fields"] or {})
            was_deleted = bool(norm_text((rec["fields"] or {}).get("AI 提示")))
            if was_deleted:
                fields = dict(changed)
                fields["AI 提示"] = ""
                to_restore.append({"record_id": rec["record_id"], "fields": fields})
            elif changed:
                to_update.append({"record_id": rec["record_id"], "fields": changed})
            else:
                unchanged += 1

    to_mark_deleted = []
    for name, recs in fs_map.items():
        if name in dh_map:
            continue
        for rec in recs:
            if norm_text((rec["fields"] or {}).get("AI 提示")):
                continue
            to_mark_deleted.append({"record_id": rec["record_id"], "name": name})

    return {"to_create": to_create, "to_update": to_update, "to_restore": to_restore,
            "to_mark_deleted": to_mark_deleted, "unchanged": unchanged,
            "total_donghuo": len(dh_map), "total_feishu": len(feishu_records)}


# ============ 各部分飞书写入（返回 stats dict；失败 raise）============

def run_full_replace(token: str, part: str, payload: dict) -> dict:
    """全量替换：清空 → 写入（①②③④ 共用）"""
    table_id = TABLES[part]
    log(f"[{PART_NAMES[part]}] 飞书写入开始 ...")
    existing = bitable_list_records(token, table_id)
    ids = [r["record_id"] for r in existing]
    if ids:
        bitable_batch_delete(token, table_id, ids)
    log(f"[{PART_NAMES[part]}] 已清空旧记录 {len(ids)} 条")

    if part == "dingdan":
        field_types = bitable_get_field_types(token, table_id)
        writable = [k for k, v in field_types.items() if v != "unsupported"]
        log(f"[{PART_NAMES[part]}] 动态字段探测: {len(field_types)} 个字段，可写 {len(writable)} 个")
        records = df_to_records(payload["df"], field_types, DINGDAN_ALIASES)
    elif part == "chuku":
        records = df_to_records(payload["df"], CHUKU_FIELD_TYPES)
    elif part == "yingshou":
        records = df_to_records(payload["df"], YINGSHOU_FIELD_TYPES)
    else:  # wanglai
        records = rows_to_records(payload["rows"], WANGLAI_FIELD_TYPES)

    # 预检：先写 5 条验证字段映射
    test_batch = records[:min(5, len(records))]
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    r = requests.post(f"{_records_url(table_id)}/batch_create", headers=h,
                      json={"records": test_batch}, timeout=60)
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"预检写入失败: {d.get('code')} {d.get('msg')} 样本={json.dumps(test_batch[0], ensure_ascii=False)[:400]}")
    remaining = records[len(test_batch):]
    written = len(test_batch)
    if remaining:
        written += bitable_batch_create(token, table_id, remaining)
    log(f"[{PART_NAMES[part]}] ✅ 写入完成 {written} 条")
    return {"cleared": len(ids), "written": written}


def run_kehu_incremental(token: str, payload: dict) -> dict:
    """⑤ 客户：增量写入 + 已删除标记（与单脚本一致）"""
    table_id = TABLES["kehu"]
    feishu_records = bitable_list_records(token, table_id)
    plan = kehu_plan(payload["rows"], feishu_records)
    log(f"[⑤ 客户] 计划: 新增 {len(plan['to_create'])} · 更新 {len(plan['to_update'])} · "
        f"标记删除 {len(plan['to_mark_deleted'])} · 恢复 {len(plan['to_restore'])} · 无变化 {plan['unchanged']}")

    created = updated = restored = marked = 0
    if plan["to_create"]:
        created = bitable_batch_create(token, table_id, [{"fields": f} for f in plan["to_create"]])
    if plan["to_update"]:
        updated = bitable_batch_update(token, table_id, plan["to_update"])
    if plan["to_restore"]:
        restored = bitable_batch_update(token, table_id, plan["to_restore"])
        log(f"[⑤ 客户] ♻️ {restored} 条重新出现，已取消「已删除」标记")
    if plan["to_mark_deleted"]:
        marked = bitable_batch_update(token, table_id, [
            {"record_id": r["record_id"], "fields": {"AI 提示": DELETED_MARK}}
            for r in plan["to_mark_deleted"]])
        log(f"[⑤ 客户] ⚠️ {marked} 条在懂火已不存在，已标记「已删除」（记录保留）")
    return {"created": created, "updated": updated, "restored": restored,
            "marked_deleted": marked, "unchanged": plan["unchanged"],
            "deleted_names": [r["name"] for r in plan["to_mark_deleted"]],
            "total_donghuo": plan["total_donghuo"], "total_feishu": plan["total_feishu"]}


def run_kehu_fix_owner(token: str) -> int:
    """
    ⑤ 客户后处理：补参与 + 标记请复检。
    条件：参与为空 AND 来自-收付登记（Lookup）有值 AND AI 提示 != "已删除"。
    动作：取 Lookup 里第一个 user 的 open_id → 写入参与 User 字段；AI 提示写 "请复检"。
    返回处理条数。
    """
    table_id = TABLES["kehu"]
    CURRENT_STEP = "客户后处理：补参与 + 标记请复检"
    log("[⑤ 客户] 后处理：扫描参与空白但有来自-收付登记的记录 ...")
    records = bitable_list_records(token, table_id)
    updates = []
    for rec in records:
        f = rec["fields"] or {}
        # 跳过 AI 提示已标记 "已删除" 的客户
        ai_hint = norm_text(f.get("AI 提示"))
        if ai_hint == DELETED_MARK:
            continue
        # 参与为空 = 没有 User 对象
        owner = f.get("参与")
        owner_empty = not owner or (isinstance(owner, list) and len(owner) == 0) or (isinstance(owner, dict) and not owner)
        if not owner_empty:
            continue
        # 来自-收付登记是 Lookup dict，结构 {"users": [{"id": "ou_xxx", "name": "..."}]}
        lookup = f.get("来自-收付登记")
        if not isinstance(lookup, dict):
            continue
        users = lookup.get("users") or []
        if not users or not isinstance(users, list):
            continue
        first_id = users[0].get("id")
        if not first_id or not first_id.startswith("ou_"):
            continue
        # 还要排除 AI 提示已经是 "请复检" 的（避免重复写相同值）
        if ai_hint == "请复检":
            continue
        updates.append({
            "record_id": rec["record_id"],
            "fields": {
                "参与": [{"id": first_id}],
                "AI 提示": "请复检",
            },
        })
    if not updates:
        log("[⑤ 客户] ✅ 无需补参与的记录（0 条）")
        return 0
    log(f"[⑤ 客户] ⚙️ 补参与 + 标记请复检: {len(updates)} 条")
    done = bitable_batch_update(token, table_id, updates)
    log(f"[⑤ 客户] ✅ 已处理 {done} 条")
    return done


# ============ 汇总通知卡片 ============

def _part_line_md(key: str, res: dict) -> str:
    """单部分一行的 markdown 文本"""
    name = PART_NAMES[key]
    if not res.get("ok"):
        err = (res.get("error") or "未知错误").replace("\n", " ")
        return f"**{name}** ❌ 失败\n{err[:200]}"
    s = res.get("stats") or {}
    if key == "kehu":
        extra = f" · 补参与+请复检 {s.get('fixed_owner', 0)}" if s.get("fixed_owner") else ""
        return (f"**{name}** ✅ 新增 {s.get('created', 0)} · 更新 {s.get('updated', 0)} · "
                f"标记删除 {s.get('marked_deleted', 0)} · 恢复 {s.get('restored', 0)} · "
                f"无变化 {s.get('unchanged', 0)}{extra}")
    return (f"**{name}** ✅ 清空 {s.get('cleared', 0)} 条 · 写入 {s.get('written', 0)} 条")


def build_summary_card(results: dict, elapsed_s: float) -> dict:
    """5 部分汇总卡片：全成功绿 / 部分失败橙 / 全失败红"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total = len(results)
    ok_cnt = sum(1 for r in results.values() if r.get("ok"))
    if ok_cnt == total:
        template, emoji, title = "green", "✅", f"懂火数据汇总同步完成（{ok_cnt}/{total} 部分）"
    elif ok_cnt == 0:
        template, emoji, title = "red", "❌", f"懂火数据汇总同步全部失败（0/{total}）"
    else:
        template, emoji, title = "orange", "⚠️", f"懂火数据汇总同步部分完成（{ok_cnt}/{total} 部分）"

    elements = [
        {"tag": "div", "fields": [
            {"is_short": True, "text": {"tag": "lark_md", "content": f"**总耗时**\n{elapsed_s:.1f} 秒"}},
            {"is_short": True, "text": {"tag": "lark_md", "content": f"**完成时间**\n{now}"}},
        ]},
        {"tag": "hr"},
    ]
    for key in ["chuku", "dingdan", "yingshou", "wanglai", "kehu"]:
        if key in results:
            elements.append({"tag": "div", "text": {"tag": "lark_md", "content": _part_line_md(key, results[key])}})

    # 客户部分被标记删除的名单
    kehu_res = results.get("kehu") or {}
    if kehu_res.get("ok"):
        names = (kehu_res.get("stats") or {}).get("deleted_names") or []
        if names:
            show = names[:20]
            lines = "\n".join(f"· {n}" for n in show)
            if len(names) > 20:
                lines += f"\n……等共 {len(names)} 家"
            elements.append({"tag": "hr"})
            elements.append({"tag": "div", "text": {"tag": "lark_md",
                "content": f"**⚠️ 已在表内标记「已删除」的客户（懂火中已不存在，记录保留）：**\n{lines}"}})

    elements.append({"tag": "hr"})
    elements.append({"tag": "note", "elements": [{
        "tag": "plain_text",
        "content": f"飞书多维表「数据汇总（2026）」 ｜ 一次登录 · 5 部分合并同步 ｜ {now}"
    }]})
    return {"config": {"wide_screen_mode": True},
            "header": {"template": template,
                       "title": {"tag": "plain_text", "content": f"{emoji} {title}"}},
            "elements": elements}


# ============ 主流程 ============

def main():
    ap = argparse.ArgumentParser(description="懂火 5 合 1 同步工作流")
    ap.add_argument("--headless", action="store_true", help="headless Chrome（默认 headed；--ci 自动开启）")
    ap.add_argument("--ci", action="store_true",
                    help="GitHub Actions 环境：用 playwright chromium 替代真 Chrome，自动 --headless")
    ap.add_argument("--skip-download", action="store_true",
                    help="①②③ 复用 downloads/all/ 下各部分最新 xls（不启动浏览器；④⑤ 自动退回 requests 登录）")
    ap.add_argument("--dry-run", action="store_true", help="取数+解析+CSV+打印计划，不写飞书")
    ap.add_argument("--no-notify", action="store_true", help="不发飞书通知")
    ap.add_argument("--only", default="", help="只跑部分：逗号分隔 chuku,dingdan,yingshou,wanglai,kehu")
    args = ap.parse_args()

    t0 = time.time()
    only = {s.strip() for s in args.only.split(",") if s.strip()} or set(PART_NAMES)
    invalid = only - set(PART_NAMES)
    if invalid:
        raise SystemExit(f"--only 含未知部分: {invalid}（可选: {list(PART_NAMES)}）")
    ui_parts = [p for p in ["chuku", "dingdan", "yingshou"] if p in only]
    api_parts = [p for p in ["wanglai", "kehu"] if p in only]
    results = {k: {"ok": False, "error": None, "stats": None} for k in
               ["chuku", "dingdan", "yingshou", "wanglai", "kehu"] if k in only}

    log(f"==== 懂火 5 合 1 同步工作流启动（部分: {sorted(only)}）====")
    username, password = env("DH_USERNAME"), env("DH_PASSWORD")
    if not username or not password:
        raise SystemExit("缺少 DH_USERNAME / DH_PASSWORD（.env）")

    # ---- Phase 1: 浏览器登录一次 + ①②③ UI 导出 ----
    xls_paths = {}
    api_session = None
    if ui_parts and not args.skip_download:
        import ddddocr
        import playwright.sync_api as pw
        ocr = ddddocr.DdddOcr(show_ad=False)
        for f in DOWNLOAD_DIR.iterdir():
            if f.is_file():
                f.unlink()
        log("[浏览器] 启动 Chrome 并登录（全程仅此一次）...")
        browser = None
        try:
            with pw.sync_playwright() as p:
                # CI 环境用 playwright chromium（Ubuntu runner 无 Chrome），本地用真 Chrome
                launch_kwargs = {"headless": args.headless or args.ci,
                                 "args": ["--disable-blink-features=AutomationControlled"]}
                if not args.ci:
                    launch_kwargs["channel"] = "chrome"
                browser = p.chromium.launch(**launch_kwargs)
                ctx = browser.new_context(accept_downloads=True)
                page = ctx.new_page()
                browser_login(page, ocr, username, password)
                # 登录态固定，立即提取 cookies 给 ④⑤ API 用（防浏览器中途挂掉）
                api_session = session_from_browser(ctx)
                log("[浏览器] ✅ cookies 已提取（④⑤ API 复用此登录态）")
                exporters = {"chuku": export_chuku, "dingdan": export_dingdan,
                             "yingshou": export_yingshou}
                for part in ui_parts:
                    try:
                        xls_paths[part] = exporters[part](page)
                    except Exception as e:
                        results[part]["error"] = f"UI 导出失败: {e}"
                        log(f"[{PART_NAMES[part]}] ❌ {e}")
                        if page.is_closed():
                            for rest in ui_parts[ui_parts.index(part) + 1:]:
                                results[rest]["error"] = "浏览器已关闭，导出中断"
                            break
        except Exception as e:
            log(f"[浏览器] ❌ 登录/浏览器失败: {e}")
            for part in ui_parts:
                if not results[part]["error"]:
                    results[part]["error"] = f"浏览器登录失败: {e}"
        finally:
            if browser:
                try:
                    browser.close()
                except Exception:
                    pass
    elif ui_parts:
        # --skip-download：复用已有 xls（按 {part}__ 前缀找最新）
        for part in ui_parts:
            files = sorted(DOWNLOAD_DIR.glob(f"{part}__*.xls"),
                           key=lambda x: x.stat().st_mtime, reverse=True)
            if files:
                xls_paths[part] = files[0]
                log(f"[{PART_NAMES[part]}] [跳过下载] 复用: {files[0].name}")
            else:
                results[part]["error"] = f"--skip-download 但 downloads/all/ 下没有 {part}__*.xls"

    # ---- Phase 2: ④⑤ API 拉取（cookie session 或退回 donghuo_login）----
    rows_data = {}
    if api_parts:
        if api_session is None:
            log("[API] 未走浏览器（--skip-download 或浏览器失败），退回 donghuo_login.py 登录...")
            api_session = login_donghuo()
        if api_session is None:
            for part in api_parts:
                results[part]["error"] = "懂火 requests 登录失败（④⑤ 均无法拉取）"
        else:
            fetchers = {"wanglai": fetch_wanglai, "kehu": fetch_kehu}
            for part in api_parts:
                try:
                    rows_data[part] = fetchers[part](api_session)
                except Exception as e:
                    results[part]["error"] = f"API 拉取失败: {e}"
                    log(f"[{PART_NAMES[part]}] ❌ {e}")

    # ---- Phase 3: 解析 + CSV 备份 ----
    payload = {}   # key -> {"df"? , "rows"?, "csv"}
    for part in ["chuku", "dingdan", "yingshou"]:
        if part not in xls_paths:
            continue
        try:
            df, csv_path = xls_to_df(xls_paths[part], part)
            if len(df) == 0:
                raise RuntimeError("导出文件解析后为空")
            payload[part] = {"df": df, "csv": csv_path}
        except Exception as e:
            results[part]["error"] = results[part]["error"] or f"解析失败: {e}"
            log(f"[{PART_NAMES[part]}] ❌ 解析失败: {e}")
    for part in ["wanglai", "kehu"]:
        if part not in rows_data:
            continue
        try:
            if not rows_data[part]:
                raise RuntimeError("接口返回空")
            csv_path = rows_to_csv(rows_data[part], part)
            payload[part] = {"rows": rows_data[part], "csv": csv_path}
        except Exception as e:
            results[part]["error"] = results[part]["error"] or f"CSV 备份失败: {e}"
            log(f"[{PART_NAMES[part]}] ❌ {e}")

    # ---- dry-run：打印各部分统计 + ⑤ 客户增量计划，到此为止 ----
    if args.dry_run:
        log("==== DRY-RUN 结果（不写飞书，CSV 全部保留）====")
        for part in ["chuku", "dingdan", "yingshou"]:
            if part not in only:
                continue
            if part in payload:
                df = payload[part]["df"]
                log(f"  {PART_NAMES[part]}: {len(df)} 行，列={list(df.columns)[:8]}...")
            else:
                log(f"  {PART_NAMES[part]}: ❌ {results[part]['error']}")
        for part in ["wanglai", "kehu"]:
            if part not in only:
                continue
            if part in payload:
                log(f"  {PART_NAMES[part]}: {len(payload[part]['rows'])} 行")
            else:
                log(f"  {PART_NAMES[part]}: ❌ {results[part]['error']}")
        if "kehu" in payload:
            try:
                token = feishu_token()
                fs = bitable_list_records(token, TABLES["kehu"])
                plan = kehu_plan(payload["kehu"]["rows"], fs)
                log(f"  ⑤ 客户增量计划: 新增 {len(plan['to_create'])} · 更新 {len(plan['to_update'])} · "
                    f"标记删除 {len(plan['to_mark_deleted'])} · 恢复 {len(plan['to_restore'])} · "
                    f"无变化 {plan['unchanged']}")
                if plan["to_mark_deleted"]:
                    log(f"    将标记: {[r['name'] for r in plan['to_mark_deleted'][:10]]}")
            except Exception as e:
                log(f"  ⑤ 客户计划读取失败: {e}")
        log(f"==== DRY-RUN 完成，耗时 {time.time() - t0:.1f}s ====")
        return 0

    # ---- Phase 4: 飞书写入（每部分独立 try，互不阻断）----
    if payload:
        log("[飞书] 获取 tenant_access_token ...")
        token = feishu_token()
        for part in ["chuku", "dingdan", "yingshou", "wanglai", "kehu"]:
            if part not in payload:
                continue
            try:
                if part == "kehu":
                    stats = run_kehu_incremental(token, payload[part])
                    # 后处理：补参与 + 标记请复检（在增量同步成功后跑；失败则跳过）
                    if args.dry_run:
                        log("[⑤ 客户] [DRY-RUN] 跳过补参与后处理")
                    else:
                        try:
                            fixed = run_kehu_fix_owner(token)
                            stats["fixed_owner"] = fixed
                        except Exception as e2:
                            log(f"[⑤ 客户] ⚠️ 补参与后处理异常（不阻断主流程）: {e2}")
                            stats["fixed_owner"] = 0
                else:
                    stats = run_full_replace(token, part, payload[part])
                results[part].update(ok=True, stats=stats)
                # 该部分全部成功 → 删它的 CSV
                csv_p = payload[part]["csv"]
                if csv_p.exists():
                    csv_p.unlink()
                    log(f"[{PART_NAMES[part]}] [清理] CSV 已删除: {csv_p.name}")
            except Exception as e:
                results[part]["error"] = f"飞书写入失败: {e}"
                log(f"[{PART_NAMES[part]}] ❌ 写入失败: {e}")
                log(f"    {traceback.format_exc(limit=3)}")

    elapsed = time.time() - t0
    ok_cnt = sum(1 for r in results.values() if r.get("ok"))
    log(f"==== 全部结束: {ok_cnt}/{len(results)} 部分成功，总耗时 {elapsed:.1f}s ====")

    # ---- Phase 5: 汇总通知（一张卡）----
    if not args.no_notify:
        try:
            notify_token = feishu_token()
            feishu_send_card(FEISHU_NOTIFY_UNION_ID,
                             build_summary_card(results, elapsed), notify_token)
        except Exception as e:
            log(f"[飞书通知] ⚠️ 发送异常: {e}")

    return 0 if ok_cnt == len(results) else (2 if ok_cnt == 0 else 1)


if __name__ == "__main__":
    sys.exit(main())
