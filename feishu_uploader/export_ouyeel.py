#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
欧冶平台 - 产能预售明细导出 → 飞书云盘

数据来源：https://www.ouyeel.com/search-ng/queryWip/wipMain?pageIndex=0&pageSize=50
流程：
  1. Playwright 启动浏览器（绕过瑞数反爬）
  2. 加载 storage_state（cookie 复用，避开 SSO 登录流程不确定性）
  3. 拦截 XHR 响应拿 JSON（兜底：DOM evaluate 抽卡片）
  4. 分页遍历所有页
  5. 生成 CSV（UTF-8-SIG）
  6. 上传到飞书云盘指定文件夹
  7. 可选：飞书机器人 Webhook 通知

环境变量（必填）：
  OUYEEL_STORAGE_STATE     base64 编码的 storage_state.json
  FEISHU_APP_ID            飞书自建应用 App ID
  FEISHU_APP_SECRET        飞书自建应用 App Secret
  FEISHU_FOLDER_TOKEN      欧冶数据目标文件夹 token (fldcn...)

环境变量（可选）：
  OUYEEL_PAGE_SIZE         每页条数，默认 50
  OUYEEL_MAX_PAGES         最大页数防失控，默认 10
  FEISHU_WEBHOOK_URL       飞书机器人 Webhook（用于告警/通知）
  FEISHU_WEBHOOK_SECRET    机器人签名密钥
  DELIVERY_MODE            交付方式: feishu(默认) / local(仅本地保存)

本地刷新 storage_state（cookie 过期后执行）：
  python feishu_uploader/export_ouyeel.py --dump-state
  → 启动浏览器，手动登录欧冶后回车，生成 .ouyeel_state.json
"""

import argparse
import base64
import csv
import datetime
import io
import json
import os
import sys
import time
import traceback
from typing import Any

import requests


OUYEEL_BASE = "https://www.ouyeel.com"
WIP_URL_TEMPLATE = f"{OUYEEL_BASE}/search-ng/queryWip/wipMain"
SSO_LOGIN_HOST = "login-ng.ouyeel.com"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"

STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".ouyeel_state.json")

# 真实浏览器 UA（避开 headless 检测）
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

# 反检测脚本：伪造浏览器指纹，避开瑞数信息（RS-Anti-Bot）自动化检测
# 覆盖 navigator.webdriver / plugins / languages / platform / window.chrome / permissions 等
STEALTH_JS = r"""
// 1. navigator.webdriver = undefined
try {
    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
} catch (e) {}

// 2. navigator.languages
try {
    Object.defineProperty(navigator, 'languages', {get: () => ['zh-CN', 'zh', 'en']});
} catch (e) {}

// 3. navigator.plugins（空数组会被识别为 headless）
try {
    Object.defineProperty(navigator, 'plugins', {
        get: () => [
            {name: 'Chrome PDF Plugin', filename: 'internal-pdf-viewer', description: 'Portable Document Format'},
            {name: 'Chrome PDF Viewer', filename: 'internal-pdf-viewer', description: ''},
            {name: 'Native Client', filename: 'internal-pdf-viewer', description: ''}
        ]
    });
} catch (e) {}

// 4. window.chrome 对象
try {
    window.chrome = window.chrome || {};
    window.chrome.runtime = window.chrome.runtime || {};
    window.chrome.loadTimes = function() { return {}; };
    window.chrome.csi = function() { return {}; };
} catch (e) {}

// 5. navigator.permissions.query（避免 Notification permission 异常）
try {
    const origQuery = window.navigator.permissions.query;
    window.navigator.permissions.query = (parameters) =>
        parameters.name === 'notifications'
            ? Promise.resolve({state: Notification.permission})
            : origQuery(parameters);
} catch (e) {}

// 6. WebGL renderer/vendor（避开 headless 的 SwiftShader 标记）
try {
    const getParameter = WebGLRenderingContext.prototype.getParameter;
    WebGLRenderingContext.prototype.getParameter = function(parameter) {
        if (parameter === 37445) return 'Intel Inc.';
        if (parameter === 37446) return 'Intel Iris OpenGL Engine';
        return getParameter.apply(this, arguments);
    };
} catch (e) {}

// 7. navigator.platform
try {
    Object.defineProperty(navigator, 'platform', {get: () => 'Win32'});
} catch (e) {}

// 8. navigator.hardwareConcurrency（headless 常返回 1-2）
try {
    Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});
} catch (e) {}

// 9. navigator.deviceMemory
try {
    Object.defineProperty(navigator, 'deviceMemory', {get: () => 8});
} catch (e) {}

// 10. 隐藏 Playwright 自动化痕迹
try {
    delete window.navigator.__proto__.webdriver;
} catch (e) {}
"""


def apply_stealth(context) -> None:
    """给 Playwright context 注入反检测脚本"""
    context.add_init_script(STEALTH_JS)


class CookieExpiredError(RuntimeError):
    """storage_state cookie 过期，被重定向到登录页"""


# ---------------- 通用 ----------------

def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def log(msg: str) -> None:
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")


# ---------------- 飞书云盘（复用 lindiao 模式，自带一份） ----------------

def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": app_id, "app_secret": app_secret}, timeout=15)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换取 tenant_access_token 失败: {data}")
    token = str(data.get("tenant_access_token") or "")
    if not token:
        raise RuntimeError(f"响应中无 tenant_access_token: {data}")
    log(f"[飞书] 获取 tenant_access_token 成功 (len={len(token)})")
    return token


def feishu_upload_to_folder(token: str, folder_token: str,
                            file_bytes: bytes, filename: str,
                            max_size_mb: int = 20) -> dict:
    size = len(file_bytes)
    if size == 0:
        raise ValueError("上传文件为空")
    if size > max_size_mb * 1024 * 1024:
        raise ValueError(f"文件 {size/1024/1024:.1f}MB 超过 upload_all 上限 {max_size_mb}MB")
    if not folder_token:
        raise ValueError("缺少 FEISHU_FOLDER_TOKEN")

    files = {"file": (filename, file_bytes, "application/octet-stream")}
    data = {
        "file_name": filename,
        "parent_type": "explorer",
        "parent_node": folder_token,
        "size": str(size),
    }
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{FEISHU_OPEN_BASE}/drive/v1/files/upload_all"

    r = requests.post(url, data=data, files=files, headers=headers, timeout=120)
    resp = r.json()
    if resp.get("code") != 0:
        raise RuntimeError(f"上传飞书云盘失败: code={resp.get('code')}, "
                           f"msg={resp.get('msg')}, raw={resp}")
    file_info = resp.get("data") or {}
    file_token = file_info.get("file_token") or file_info.get("token") or ""
    log(f"[飞书云盘] 上传成功, file_token={file_token}, name={file_info.get('name')}")
    return file_info


def _feishu_sign(secret: str, timestamp: str) -> str:
    import hmac
    import hashlib
    import base64 as _b64
    string_to_sign = f"{timestamp}\n{secret}"
    h = hmac.new(secret.encode("utf-8"),
                 string_to_sign.encode("utf-8"),
                 hashlib.sha256)
    return _b64.b64encode(h.digest()).decode("utf-8")


def feishu_send_bot_text(webhook_url: str, secret: str, text: str) -> None:
    body = {"msg_type": "text", "content": {"text": text}}
    if secret:
        timestamp = str(int(time.time()))
        body["timestamp"] = timestamp
        body["sign"] = _feishu_sign(secret, timestamp)
    r = requests.post(webhook_url, json=body, timeout=15)
    resp = r.json()
    if resp.get("code") not in (0, None) and resp.get("StatusCode") != 0:
        raise RuntimeError(f"飞书机器人通知失败: {resp}")
    log("[飞书通知] 发送成功")


def notify(webhook_url: str, webhook_secret: str, text: str) -> None:
    """安全通知：失败只警告不抛"""
    if not webhook_url:
        print("\n[通知内容]")
        print(text)
        return
    try:
        feishu_send_bot_text(webhook_url, webhook_secret, text)
    except Exception as exc:
        log(f"[警告] 通知发送失败: {exc}")


# ---------------- storage_state 处理 ----------------

def load_storage_state_bytes() -> bytes:
    """从环境变量 OUYEEL_STORAGE_STATE (base64) 或本地 .ouyeel_state.json 读取"""
    b64 = env("OUYEEL_STORAGE_STATE")
    if b64:
        try:
            return base64.b64decode(b64)
        except Exception as exc:
            raise RuntimeError(f"OUYEEL_STORAGE_STATE base64 解码失败: {exc}")
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "rb") as fp:
            return fp.read()
    raise RuntimeError(
        "未找到 storage_state：请设置 OUYEEL_STORAGE_STATE 环境变量，"
        "或本地先跑 `python feishu_uploader/export_ouyeel.py --dump-state` 生成 .ouyeel_state.json"
    )


def save_state_for_runtime() -> str:
    """把 base64 环境变量解码落盘为 .ouyeel_state.json，返回路径"""
    b64 = env("OUYEEL_STORAGE_STATE")
    if b64:
        try:
            raw = base64.b64decode(b64)
            with open(STATE_FILE, "wb") as fp:
                fp.write(raw)
            log(f"[state] 已从 OUYEEL_STORAGE_STATE 解码到 {STATE_FILE}")
        except Exception as exc:
            raise RuntimeError(f"OUYEEL_STORAGE_STATE 解码失败: {exc}")
    elif not os.path.exists(STATE_FILE):
        raise RuntimeError(
            "缺少 storage_state：请设置 OUYEEL_STORAGE_STATE 或本地跑 --dump-state"
        )
    return STATE_FILE


# ---------------- Playwright 浏览器 ----------------

def create_browser_and_context(headless: bool = True, storage_state_path: str | None = None,
                                 use_real_chrome: bool = False):
    """启动浏览器，建立带 storage_state 的 context（含完整反检测配置）"""
    from playwright.sync_api import sync_playwright

    pw = sync_playwright().start()
    launch_args = [
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-features=IsolateOrigins,site-per-process",
        "--disable-site-isolation-trials",
    ]
    # 关键：排除 Playwright 默认加的 --enable-automation 标志（瑞数会检测）
    launch_kwargs = {
        "headless": headless,
        "args": launch_args,
        "ignore_default_args": ["--enable-automation"],
    }
    # 本地 dump-state 模式可选：用本机真实 Chrome（指纹更真）
    if use_real_chrome:
        launch_kwargs["channel"] = "chrome"
    try:
        browser = pw.chromium.launch(**launch_kwargs)
    except Exception as exc:
        # channel="chrome" 失败（本机没装 Chrome）→ 回退默认 chromium
        if use_real_chrome:
            log(f"[警告] 本机未检测到 Chrome，回退到 chromium: {exc}")
            launch_kwargs.pop("channel", None)
            browser = pw.chromium.launch(**launch_kwargs)
        else:
            raise
    context_kwargs = {
        "user_agent": USER_AGENT,
        "viewport": {"width": 1366, "height": 900},
        "locale": "zh-CN",
        "timezone_id": "Asia/Shanghai",
    }
    if storage_state_path and os.path.exists(storage_state_path):
        context_kwargs["storage_state"] = storage_state_path
    context = browser.new_context(**context_kwargs)
    apply_stealth(context)
    return pw, browser, context


def check_login_redirect(page) -> bool:
    """检测是否被重定向到 SSO 登录页（cookie 过期）"""
    cur = page.url or ""
    return SSO_LOGIN_HOST in cur or "/sso/login" in cur


# ---------------- 数据抽取：XHR 拦截 ----------------

def register_xhr_capture(page, captured: list, log_all: bool = False) -> None:
    """注册 response 处理器，捕获含 wip/query 关键字的 JSON 响应。
    log_all=True 时打印所有 XHR URL（用于首跑诊断）"""
    def on_response(response):
        try:
            url = response.url
            url_lower = url.lower()
            # 诊断模式：打印所有 XHR/fetch 响应
            if log_all:
                try:
                    rt = response.request.resource_type
                    if rt in ("xhr", "fetch"):
                        ct = response.headers.get("content-type", "")
                        log(f"  [XHR-ALL] {response.status} [{rt}] {ct[:30]} {url[:120]}")
                except Exception:
                    pass
            # 宽过滤：URL 含 wip / query / search 等
            if not any(k in url_lower for k in ("wip", "query", "search", "list")):
                return
            ct = response.headers.get("content-type", "")
            if "json" not in ct.lower():
                return
            try:
                data = response.json()
            except Exception:
                return
            captured.append({"url": url, "data": data})
            preview = json.dumps(data, ensure_ascii=False)[:200]
            log(f"  [XHR] {url}")
            log(f"         预览: {preview}")
        except Exception:
            pass
    page.on("response", on_response)


def extract_list_from_json(data: Any) -> list:
    """递归从 JSON 找出最像数据列表的数组（长度 >= 5 的数组优先）"""
    candidates: list[tuple[int, list]] = []

    def walk(node):
        if isinstance(node, list):
            if len(node) > 0 and isinstance(node[0], (dict, str)):
                candidates.append((len(node), node))
        elif isinstance(node, dict):
            for v in node.values():
                walk(v)

    walk(data)
    if not candidates:
        return []
    # 取最长的数组
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


# ---------------- 数据抽取：DOM 兜底 ----------------

def extract_from_dom(page) -> list[dict]:
    """DOM 兜底：抽所有含数据特征的卡片节点"""
    js = """
    () => {
        const cards = document.querySelectorAll(
            '[class*="card"], [class*="item"], [class*="product"], ' +
            '[class*="goods"], [class*="list"] > div'
        );
        const results = [];
        for (const c of cards) {
            const text = (c.innerText || '').trim();
            if (text.length > 30 && (
                text.includes('准发') || text.includes('可供') ||
                text.includes('/吨') || text.includes('签约截止')
            )) {
                results.push({
                    text: text,
                    className: c.className || '',
                    tagName: c.tagName
                });
            }
        }
        return results;
    }
    """
    try:
        items = page.evaluate(js)
    except Exception as exc:
        log(f"  [DOM] evaluate 失败: {exc}")
        return []

    records = []
    for item in items:
        text = item.get("text", "")
        if not text:
            continue
        records.append({"raw_text": text, "source": "dom"})
    return records


# ---------------- 分页抓取 ----------------

def fetch_one_page(context, page_index: int, page_size: int,
                   captured: list, timeout_ms: int = 30000,
                   debug: bool = False) -> list[dict]:
    """抓单页：goto URL → 等 XHR → 兜底 DOM。debug=True 打印详细诊断"""
    page = context.new_page()
    try:
        # 首页 + debug 模式：打印所有 XHR URL 用于诊断接口名
        register_xhr_capture(page, captured, log_all=(debug and page_index == 0))

        url = f"{WIP_URL_TEMPLATE}?pageIndex={page_index}&pageSize={page_size}"
        log(f"[抓取] 第 {page_index + 1} 页: {url}")

        try:
            page.goto(url, wait_until="networkidle", timeout=timeout_ms)
        except Exception as exc:
            log(f"  [警告] 页面加载异常: {exc}，尝试继续")

        if check_login_redirect(page):
            raise CookieExpiredError("cookie 过期，被重定向到 SSO 登录页")

        # 给 SPA + 瑞数挑战 JS 一点时间渲染
        page.wait_for_timeout(3000)

        # 诊断：打印页面状态（首跑必看，判断是否被瑞数拦截）
        if debug and page_index == 0:
            try:
                cur_url = page.url
                title = page.title()
                log(f"  [诊断] 最终 URL: {cur_url}")
                log(f"  [诊断] 页面标题: {title}")
                # 页面文本前 300 字符
                body_text = page.evaluate(
                    "() => document.body ? document.body.innerText.substring(0, 300) : ''"
                )
                if body_text and len(body_text.strip()) > 5:
                    log(f"  [诊断] 页面文本前300: {body_text[:300]}")
                else:
                    log(f"  [诊断] ⚠️ 页面可视文本为空！瑞数可能在拦截。")
                # HTML 前 500 字符（看是否是瑞数挑战页 $_ts=window['$_ts']）
                html = page.content()
                log(f"  [诊断] HTML 前500字符:")
                log(f"  {html[:500]}")
                # 保存截图
                try:
                    shot_path = os.path.join(os.path.dirname(STATE_FILE), f".ouyeel_page{page_index}_dbg.png")
                    page.screenshot(path=shot_path, full_page=False)
                    log(f"  [诊断] 截图已保存: {shot_path}")
                except Exception:
                    pass
            except Exception as exc:
                log(f"  [诊断] 状态读取失败: {exc}")

        # 优先用 XHR 拦截的 JSON
        records: list[dict] = []
        for cap in captured:
            lst = extract_list_from_json(cap["data"])
            if len(lst) > len(records):
                records = lst
                log(f"  [XHR] 选用 {cap['url']}，{len(lst)} 条记录")

        if not records:
            log(f"  [XHR] 未拦到数据 JSON，回退 DOM 抽取")
            records = extract_from_dom(page)
            log(f"  [DOM] 抽到 {len(records)} 条卡片")

        if not records:
            log(f"  [警告] 第 {page_index + 1} 页未取到数据")

        return records
    finally:
        try:
            page.close()
        except Exception:
            pass


def fetch_all_pages(context, page_size: int, max_pages: int,
                    debug: bool = False) -> tuple[list[dict], int]:
    """遍历所有页，返回 (合并记录列表, 估算总页数)"""
    all_records: list[dict] = []
    total_pages = 1

    for page_index in range(max_pages):
        captured: list = []
        try:
            records = fetch_one_page(context, page_index, page_size, captured,
                                      debug=debug)
        except CookieExpiredError:
            raise
        except Exception as exc:
            log(f"  [警告] 第 {page_index + 1} 页抓取失败: {exc}")
            break

        if not records:
            break

        # 首页推算总页数
        if page_index == 0:
            # 优先从 captured 的 JSON 找 total 字段
            total = find_total_count(captured)
            if total and total > 0:
                total_pages = (total + page_size - 1) // page_size
                log(f"[分页] 首页返回 total={total}，预计 {total_pages} 页")
            # 否则按"行数等于 page_size 则继续"的策略

        all_records.extend(records)

        # 提前终止：当前页行数 < page_size（最后一页）
        if len(records) < page_size:
            log(f"[分页] 第 {page_index + 1} 页仅 {len(records)} 条 < {page_size}，已是末页")
            break

        # 避免请求过快
        time.sleep(1)

    return all_records, total_pages


def find_total_count(captured: list) -> int | None:
    """从 XHR 捕获的 JSON 找 total / totalcount / count 字段"""
    for cap in captured:
        data = cap.get("data")
        if not isinstance(data, dict):
            continue
        for key in ("total", "totalCount", "totalcount", "count", "totalNum", "totalRows"):
            v = data.get(key)
            if isinstance(v, (int, float)) and v > 0:
                return int(v)
            # 嵌套一层
            for sub in data.values():
                if isinstance(sub, dict):
                    v2 = sub.get(key)
                    if isinstance(v2, (int, float)) and v2 > 0:
                        return int(v2)
    return None


# ---------------- CSV 生成 ----------------

def records_to_csv_bytes(records: list[dict]) -> tuple[bytes, int, int]:
    """把记录列表转为 CSV 字节（UTF-8-SIG）"""
    if not records:
        raise RuntimeError("无数据可写 CSV")

    # 统一所有 key 作为列
    all_keys: list[str] = []
    seen = set()
    for r in records:
        for k in r.keys():
            if k not in seen:
                seen.add(k)
                all_keys.append(k)

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(all_keys)
    for r in records:
        writer.writerow([_flatten_value(r.get(k, "")) for k in all_keys])
    csv_bytes = buf.getvalue().encode("utf-8-sig")
    return csv_bytes, len(records), len(all_keys)


def _flatten_value(v: Any) -> str:
    """把 dict/list 值扁平化为字符串"""
    if v is None:
        return ""
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return str(v).strip()


# ---------------- --dump-state 模式 ----------------

def _attach_diagnostics(page, log_prefix: str = "[事件]") -> None:
    """给 page 挂事件级诊断（requestfailed/pageerror/console/关键响应）。
    经验：空白页常见原因是 JS 报错、资源请求失败、或 Console 有明确提示。
    加上这三项就能直接定位。"""
    # 记录前 15 条关键响应（首页 HTML / 瑞数 JS bundle / 登录接口）
    response_seen: list[str] = []

    def on_response(resp):
        try:
            url = resp.url
            # 过滤图标/无关静态资源，只记录关键响应
            if any(ext in url.lower() for ext in (".png", ".jpg", ".jpeg", ".gif", ".ico", ".svg", ".css")):
                return
            ct = resp.headers.get("content-type", "")
            ct_lower = ct.lower()
            is_key = (
                "html" in ct_lower
                or "javascript" in ct_lower
                or "json" in ct_lower
                or resp.request.method.upper() == "POST"
                or resp.request.resource_type in ("document", "script", "xhr", "fetch")
            )
            if is_key and len(response_seen) < 20:
                status = resp.status
                size = resp.headers.get("content-length", "?")
                resource = resp.request.resource_type
                response_seen.append(url)
                log(f"{log_prefix} 响应 {status} [{resource}] {size}B {url[:120]}")
                # 如果是 4xx/5xx，记录下
                if 400 <= status < 600:
                    try:
                        body = resp.text()
                        log(f"{log_prefix}   响应异常 {status}，内容前500: {body[:500]}")
                    except Exception:
                        pass
        except Exception:
            pass

    def on_requestfailed(req):
        try:
            err = req.failure
            err_text = ""
            if err is not None and callable(err):
                try:
                    err_text = str(err())
                except Exception:
                    err_text = str(err)
            log(f"{log_prefix} 请求失败 [{req.method}] {req.url[:120]}")
            if err_text:
                log(f"{log_prefix}   原因: {err_text}")
        except Exception:
            pass

    def on_pageerror(err):
        log(f"{log_prefix} ⚠️ JS 报错: {str(err)[:300]}")

    def on_console(msg):
        # 只记录 error/warning，普通 info 会刷屏
        try:
            mt = msg.type
            if mt in ("error", "warning"):
                text = msg.text[:300]
                log(f"{log_prefix} 控制台 {mt.upper()}: {text}")
                if mt == "error":
                    # 附带位置信息（如果有）
                    try:
                        loc = msg.location
                        if loc:
                            log(f"{log_prefix}   位置: {loc.get('url','')}:{loc.get('lineNumber','')}")
                    except Exception:
                        pass
        except Exception:
            pass

    page.on("response", on_response)
    page.on("requestfailed", on_requestfailed)
    page.on("pageerror", on_pageerror)
    page.on("console", on_console)


def dump_state_interactive() -> int:
    """启动浏览器让用户手动登录，登录后保存 storage_state（含完整事件级诊断）"""
    import tempfile

    log("启动浏览器，请在打开的窗口中手动登录欧冶...")
    log(f"登录页: https://{SSO_LOGIN_HOST}/sso/login?service={OUYEEL_BASE}")
    log("[提示] 优先用本机真实 Chrome（指纹更不易被识别）；未装则用 chromium。")

    pw = browser = context = None
    try:
        # use_real_chrome=True：本地模式优先用真实 Chrome
        pw, browser, context = create_browser_and_context(
            headless=False, use_real_chrome=True
        )
        # 经验755694：先取已有第1页（context 启动可能自带 about:blank），没有才新建
        existing_pages = context.pages
        if existing_pages:
            page = existing_pages[0]
            # 关掉多余的页（Chrome恢复标签页会残留其他tab，导致用户看到的是旧空白页）
            for p in existing_pages[1:]:
                try: p.close()
                except Exception: pass
            log(f"[诊断] 复用已有页面: {page.url}")
        else:
            page = context.new_page()
            log(f"[诊断] 新建页面")

        # 强制切到前台（用户肉眼看到的就是这个页，不会被其他tab干扰）
        try:
            page.bring_to_front()
        except Exception:
            pass

        # ===== 诊断 1：挂事件监听器 =====
        _attach_diagnostics(page, log_prefix="[诊断]")

        login_url = f"https://{SSO_LOGIN_HOST}/sso/login?service={OUYEEL_BASE}"
        log(f"[加载] {login_url}")

        # ===== 诊断 2：把 goto 包成 try/except，失败时保存关键证据 =====
        resp = None
        goto_exception = None
        try:
            resp = page.goto(
                login_url,
                wait_until="networkidle",
                timeout=60000,
            )
        except Exception as exc:
            goto_exception = exc
            log(f"[诊断] goto 抛异常: {type(exc).__name__}: {exc}")

        # 给瑞数挑战 JS 额外执行时间
        log("[诊断] 等待瑞数挑战脚本执行 (5 秒)...")
        page.wait_for_timeout(5000)

        # 再次 bring_to_front 防被抢走焦点
        try: page.bring_to_front()
        except Exception: pass

        # ===== 诊断 3：加载完后保存截图 + HTML 片段 =====
        try:
            shot_path = os.path.join(os.path.dirname(STATE_FILE), ".ouyeel_login_dbg.png")
            page.screenshot(path=shot_path, full_page=False)
            log(f"[诊断] 页面截图已保存: {shot_path}（请打开看看，肉眼可见是空白还是被拦截）")
        except Exception as exc:
            log(f"[诊断] 截图失败: {exc}")

        # 状态读取
        try:
            cur_url = page.url
            status_code = getattr(resp, "status", None) if resp is not None else None
            title = ""
            try: title = page.title()
            except Exception: title = ""
            log(f"[状态] 最终 URL: {cur_url}")
            log(f"[状态] 响应状态码: {status_code}")
            log(f"[状态] 页面标题: {title}")

            # HTML 前 2000 字符（看是否被重定向到拦截页）
            try:
                html = page.content()
                html_snippet_path = os.path.join(os.path.dirname(STATE_FILE), ".ouyeel_login_html.txt")
                with open(html_snippet_path, "w", encoding="utf-8") as fp:
                    fp.write(html)
                log(f"[诊断] 页面 HTML 已保存到: {html_snippet_path}（完整HTML）")
                log(f"[诊断] HTML 前500字符:\n{html[:500]}")
            except Exception as exc:
                log(f"[诊断] HTML 读取失败: {exc}")

            # Body 文本
            try:
                body_text = page.evaluate(
                    "() => document.body ? document.body.innerText.substring(0, 500) : ''"
                )
                if body_text and len(body_text.strip()) > 10:
                    log(f"[状态] 页面文本预览: {body_text[:300]}")
                else:
                    log("[警告] 页面可视文本为空或过短！")
                    log("[提示] 请打开上面的 .ouyeel_login_dbg.png 截图，看一下你肉眼看到的页面是什么样。")
            except Exception as exc:
                log(f"[诊断] Body文本读取失败: {exc}")

            # document.readyState + 是否有 script 标签
            try:
                info = page.evaluate(
                    """() => {
                        const scripts = Array.from(document.scripts).map(s => s.src).filter(Boolean).slice(0, 10);
                        return {
                            readyState: document.readyState,
                            scripts: scripts,
                            origin: location.origin,
                            errors: window.__playwright_errors ? window.__playwright_errors.length : 0
                        };
                    }"""
                )
                log(f"[诊断] readyState={info.get('readyState')}, origin={info.get('origin')}")
                scripts = info.get("scripts") or []
                log(f"[诊断] 页面脚本源 (前{len(scripts)}个):")
                for src in scripts:
                    log(f"        - {src[:140]}")
            except Exception as exc:
                log(f"[诊断] 页面结构读取失败: {exc}")
        except Exception as exc:
            log(f"[警告] 无法读取页面状态: {exc}")

        # 如果 goto 抛了异常，把截图 / HTML 保存完后再退出（不要直接继续等用户登录）
        if goto_exception is not None:
            log(f"[错误] 导航失败，无法继续。请检查上述 [诊断] 日志与截图文件，把结果贴给开发者。")
            return 1

        print()
        log("如果浏览器已正常显示登录表单（看截图/肉眼），请在浏览器中完成登录。")
        log("登录成功后回到此终端按 Enter 继续...")
        input()

        # 保存 storage_state
        context.storage_state(path=STATE_FILE)
        log(f"[OK] storage_state 已保存到 {STATE_FILE}")

        # 顺便验证一下能不能访问数据页
        log("正在验证能否访问产能预售页...")
        try:
            page.goto(f"{WIP_URL_TEMPLATE}?pageIndex=0&pageSize=5", wait_until="networkidle", timeout=30000)
            if check_login_redirect(page):
                log("[警告] 访问数据页时被重定向到登录页，登录可能未完成")
            else:
                log(f"[OK] 数据页可访问，当前 URL: {page.url}")
                log(f"      页面标题: {page.title()}")
        except Exception as exc:
            log(f"[警告] 验证数据页失败: {exc}")
    finally:
        try:
            if browser: browser.close()
        except Exception:
            pass
        if pw:
            try: pw.stop()
            except Exception: pass

    # 输出 base64（方便用户更新 GitHub Secret）
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "rb") as fp:
            b64 = base64.b64encode(fp.read()).decode("ascii")
        b64_file = STATE_FILE + ".b64"
        with open(b64_file, "w", encoding="ascii") as fp:
            fp.write(b64)
        log(f"[OK] base64 已保存到 {b64_file}")
        log(f"      长度 {len(b64)} 字符，可直接粘贴到 GitHub Secret OUYEEL_STORAGE_STATE")
    else:
        log("[错误] storage_state 文件未生成，登录流程未完成")
        return 1
    return 0


# ---------------- 主入口 ----------------

def main() -> int:
    parser = argparse.ArgumentParser(description="欧冶产能预售明细导出")
    parser.add_argument("--dump-state", action="store_true",
                        help="本地登录模式：启动浏览器手动登录后导出 storage_state")
    args = parser.parse_args()

    if args.dump_state:
        return dump_state_interactive()

    # 飞书配置
    fs_app_id = env("FEISHU_APP_ID")
    fs_app_secret = env("FEISHU_APP_SECRET")
    fs_folder_token = env("FEISHU_FOLDER_TOKEN") or env("OUYEEL_FOLDER_TOKEN")
    fs_webhook_url = env("FEISHU_WEBHOOK_URL")
    fs_webhook_secret = env("FEISHU_WEBHOOK_SECRET")

    delivery_mode = (env("DELIVERY_MODE") or "feishu").lower()
    page_size = int(env("OUYEEL_PAGE_SIZE") or "50")
    max_pages = int(env("OUYEEL_MAX_PAGES") or "10")
    # 本地调试用：OUYEEL_HEADLESS=false 切换有头模式（headless 更易被瑞数检测）
    headless = env("OUYEEL_HEADLESS", "true").lower() not in ("false", "0", "no")
    # OUYEEL_DEBUG=true 打印页面 URL/标题/HTML/截图/所有 XHR（首跑必开）
    debug = env("OUYEEL_DEBUG", "false").lower() in ("true", "1", "yes")

    now = datetime.datetime.now()

    # 1) 还原 storage_state
    try:
        state_path = save_state_for_runtime()
    except Exception as exc:
        log(f"[错误] {exc}")
        return 2

    # 2) 启动浏览器抓数据
    mode_label = "headless" if headless else "有头（可见）"
    log(f"[启动] Playwright 浏览器（{mode_label}）" + (" + 诊断模式" if debug else ""))
    pw = browser = context = None
    try:
        pw, browser, context = create_browser_and_context(
            headless=headless, storage_state_path=state_path
        )
    except Exception as exc:
        log(f"[错误] 浏览器启动失败: {exc}")
        traceback.print_exc()
        return 1

    exit_code = 0
    records: list[dict] = []
    total_pages = 1
    try:
        records, total_pages = fetch_all_pages(context, page_size, max_pages,
                                                debug=debug)
    except CookieExpiredError as exc:
        log(f"[错误] {exc}")
        notify(fs_webhook_url, fs_webhook_secret,
               f"❌ 欧冶产能预售导出失败\n原因: {exc}\n请本地跑 "
               f"`python feishu_uploader/export_ouyeel.py --dump-state` 刷新登录态，"
               f"并更新 GitHub Secret OUYEEL_STORAGE_STATE")
        exit_code = 1
    except Exception as exc:
        log(f"[错误] 数据抓取失败: {exc}")
        traceback.print_exc()
        notify(fs_webhook_url, fs_webhook_secret,
               f"❌ 欧冶产能预售导出失败\n原因: {exc}")
        exit_code = 3
    finally:
        try:
            if browser: browser.close()
        except Exception:
            pass
        if pw:
            try: pw.stop()
            except Exception: pass

    if exit_code != 0:
        return exit_code

    if not records:
        log("[错误] 未抓到任何数据")
        notify(fs_webhook_url, fs_webhook_secret,
               "❌ 欧冶产能预售导出失败\n原因: 未抓到任何数据，请检查页面结构或登录态")
        return 3

    # 3) 生成 CSV
    try:
        csv_bytes, row_count, col_count = records_to_csv_bytes(records)
    except Exception as exc:
        log(f"[错误] CSV 生成失败: {exc}")
        traceback.print_exc()
        return 3

    log(f"[解析] CSV 生成: {row_count} 行, {col_count} 列, "
        f"{len(csv_bytes)/1024:.1f} KB")

    filename = f"ouyeel_{now.strftime('%Y%m%d_%H%M')}.csv"
    summary_lines = [
        f"导出时间: {now.strftime('%Y-%m-%d %H:%M:%S')}",
        f"数据来源: {WIP_URL_TEMPLATE}",
        f"数据量: {row_count} 条 ({col_count} 列)",
        f"分页: 共抓 {total_pages} 页 (pageSize={page_size})",
        f"文件名: {filename}",
        f"文件大小: {len(csv_bytes)/1024:.1f} KB",
    ]
    print("\n[汇总]")
    for line in summary_lines:
        print(f"  {line}")

    # 4) 交付
    if delivery_mode == "local":
        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        with open(out_path, "wb") as fp:
            fp.write(csv_bytes)
        print(f"\n[本地保存] 已保存到 {out_path}")
        return 0

    if delivery_mode == "feishu":
        if not (fs_app_id and fs_app_secret and fs_folder_token):
            print("[错误] 未配置完整飞书参数（FEISHU_APP_ID / FEISHU_APP_SECRET / FEISHU_FOLDER_TOKEN）")
            print("       可设 DELIVERY_MODE=local 仅本地保存")
            return 4

        try:
            token = feishu_tenant_access_token(fs_app_id, fs_app_secret)
            file_info = feishu_upload_to_folder(token, fs_folder_token,
                                                csv_bytes, filename)
            file_token = file_info.get("file_token") or file_info.get("token") or ""

            notify_lines = [f"✅ 欧冶产能预售导出完成"] + summary_lines
            notify_lines.append(f"文件位置: 飞书云盘指定文件夹")
            if file_token:
                notify_lines.append(f"file_token: {file_token}")
            notify(fs_webhook_url, fs_webhook_secret, "\n".join(notify_lines))
        except Exception as exc:
            log(f"[错误] 飞书交付失败: {exc}")
            traceback.print_exc()
            notify(fs_webhook_url, fs_webhook_secret,
                   f"❌ 欧冶产能预售导出失败\n原因: 飞书交付失败 - {exc}")
            return 4
        return 0

    print(f"[错误] 未识别的 DELIVERY_MODE: {delivery_mode}")
    return 5


if __name__ == "__main__":
    sys.exit(main())
