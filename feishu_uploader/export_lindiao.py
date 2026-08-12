#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火钢城系统 - 临调库存导出 → 飞书云盘

数据来源：库存查询 > 临调库存 页面（v_kucun_ld）的"导出"按钮
流程：
  1. 自动登录 https://erpa.donghuo.vip（ddddocr 识别验证码）
  2. 调用导出接口 /view/admin/excelbiao/kucunld，获取服务端返回的 .xls
     （本质是带 mso-number-format 样式的 HTML 表格，Excel 可直接打开）
  3. 解析 HTML 表格为 CSV（UTF-8-SIG 避免 Excel 乱码）
  4. 上传 CSV 到飞书云盘指定文件夹
  5. 可选：飞书机器人 Webhook 通知

环境变量（必填）：
  DH_USERNAME           erpa 登录账号
  DH_PASSWORD           erpa 登录密码
  FEISHU_APP_ID         飞书自建应用 App ID
  FEISHU_APP_SECRET     飞书自建应用 App Secret
  FEISHU_FOLDER_TOKEN   目标文件夹 token (fldcn...)

环境变量（可选）：
  FILTER_SXZHUANTAI     状态筛选: 空(全部) / 已锁 / 未锁
  FILTER_HUOQUAN        货权筛选: 空(全部) / 拥有 / 待赎
  FILTER_CANKU          仓库筛选（精确匹配，如 "仲鼎库"）
  FILTER_PINMIN         品名筛选
  FEISHU_WEBHOOK_URL    飞书机器人 Webhook
  FEISHU_WEBHOOK_SECRET 机器人签名密钥
  DELIVERY_MODE         交付方式: feishu(默认) / local(仅本地保存)
"""

import os
import re
import io
import csv
import time
import json
import datetime
import traceback

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_URL = "https://erpa.donghuo.vip"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"

# 库存查询外层 iframe 框架页（含 4 个 tab：公司库存/临调库存/期货库存/我的锁定）
LINDIAO_FRAME_PAGE = f"{BASE_URL}/view/admin/xiaoshou/v_ifram_kc"
# 临调库存子页面（导出按钮所在页）
LINDIAO_PAGE = f"{BASE_URL}/view/admin/xiaoshou/v_kucun_ld"
# 导出按钮对应接口（POST form1 表单 → 返回 .xls 即 HTML 表格）
EXPORT_API = f"{BASE_URL}/view/admin/excelbiao/kucunld"


# ---------------- 通用 ----------------

def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def create_session() -> requests.Session:
    """带重试机制的 requests Session"""
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
    """用 ddddocr 识别图形验证码"""
    try:
        import ddddocr  # type: ignore
        ocr = ddddocr.DdddOcr(show_ad=False)
        return ocr.classification(image_bytes).strip().replace(" ", "")
    except ImportError:
        print("[警告] 未安装 ddddocr，无法识别验证码")
        return ""
    except Exception as exc:
        print(f"[错误] 验证码识别异常: {exc}")
        return ""


def login(session: requests.Session, username: str, password: str,
          max_attempts: int = 10) -> bool:
    """登录 erpa 系统，自动重试验证码"""
    login_url = f"{BASE_URL}/controller/admin/c_longin/index"
    captcha_url = f"{BASE_URL}/common/captcha"

    for attempt in range(1, max_attempts + 1):
        print(f"[登录] 尝试 {attempt}/{max_attempts} ...")
        img_resp = session.get(captcha_url, timeout=15)
        if img_resp.status_code != 200:
            print(f"  获取验证码失败: HTTP {img_resp.status_code}")
            time.sleep(2)
            continue

        captcha_code = recognize_captcha(img_resp.content)
        if not captcha_code:
            captcha_code = "1234"  # 占位触发刷新
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


# ---------------- 临调库存导出 ----------------

def export_lindiao(session: requests.Session,
                  filters: dict | None = None,
                  timeout: int = 120) -> tuple[bytes, dict]:
    """
    调用"导出按钮"对应接口，返回 (HTML 表格字节, 元信息)。

    接口：POST /view/admin/excelbiao/kucunld
    返回：Content-Type=Application/x-msexcel，本质是 HTML 表格伪装为 .xls

    filters: 可选筛选 dict（对应 form1 表单字段），如:
        {"sxzhuantai": "已锁", "huoquan": "拥有", "canku": "仲鼎库"}
    """
    # 先访问外层框架页与内层页面，建立 Referer / 会话状态
    try:
        session.get(LINDIAO_FRAME_PAGE, timeout=15)
        session.get(LINDIAO_PAGE, timeout=15)
    except Exception as exc:
        print(f"  [警告] 访问页面预热线: {exc}")

    data = filters or {}
    print(f"[导出] POST {EXPORT_API} (筛选: {data or '无'})")
    r = session.post(EXPORT_API, data=data, timeout=timeout, allow_redirects=True)
    if r.status_code != 200:
        raise RuntimeError(f"导出接口返回 HTTP {r.status_code}")

    info = {
        "content_type": r.headers.get("Content-Type", ""),
        "content_disposition": r.headers.get("Content-Disposition", ""),
        "size_bytes": len(r.content),
        "size_kb": round(len(r.content) / 1024, 1),
    }
    print(f"  HTTP {r.status_code}, Content-Type={info['content_type']}, "
          f"size={info['size_kb']} KB")
    return r.content, info


def html_table_to_csv_bytes(html_bytes: bytes) -> tuple[bytes, int, int]:
    """
    将服务端返回的 HTML 表格（伪装 .xls）解析为 CSV 字节。

    返回：(csv 字节 UTF-8-SIG, 行数含表头, 列数)
    """
    text = html_bytes.decode("utf-8", errors="replace")

    # 找所有 <tr>...</tr>
    rows_html = re.findall(r'<tr[^>]*>(.*?)</tr>', text, flags=re.S | re.I)
    if not rows_html:
        raise RuntimeError("HTML 中未找到 <tr> 表格行")

    rows_data: list[list[str]] = []
    for tr in rows_html:
        cells = re.findall(r'<td[^>]*>(.*?)</td>', tr, flags=re.S | re.I)
        cells_clean = []
        for c in cells:
            # 去 HTML 标签和 HTML 实体
            v = re.sub(r'<[^>]+>', '', c)
            v = (v.replace('&nbsp;', ' ')
                  .replace('&amp;', '&')
                  .replace('&lt;', '<')
                  .replace('&gt;', '>')
                  .replace('&quot;', '"')
                  .replace('&#39;', "'"))
            cells_clean.append(v.strip())
        if cells_clean:
            rows_data.append(cells_clean)

    if not rows_data:
        raise RuntimeError("HTML 表格解析后无数据")

    # 写 CSV (UTF-8-SIG)
    buf = io.StringIO()
    writer = csv.writer(buf)
    for row in rows_data:
        writer.writerow(row)
    csv_bytes = buf.getvalue().encode("utf-8-sig")

    return csv_bytes, len(rows_data), len(rows_data[0])


# ---------------- 飞书云盘 ----------------

def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    """用自建应用 App ID/Secret 换 tenant_access_token"""
    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": app_id, "app_secret": app_secret},
                      timeout=15)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换取 tenant_access_token 失败: {data}")
    token = str(data.get("tenant_access_token") or "")
    if not token:
        raise RuntimeError(f"响应中无 tenant_access_token: {data}")
    print(f"[飞书] 获取 tenant_access_token 成功 (len={len(token)})")
    return token


def feishu_upload_to_folder(token: str, folder_token: str,
                            file_bytes: bytes, filename: str,
                            max_size_mb: int = 20) -> dict:
    """通过 drive/v1/files/upload_all 上传文件到飞书云盘指定文件夹（≤20MB）"""
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
    print(f"[飞书云盘] 上传成功, file_token={file_token}, name={file_info.get('name')}")
    return file_info


# ---------------- 飞书机器人通知（可选） ----------------

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
    print("[飞书通知] 发送成功")


# ---------------- 主入口 ----------------

def build_filters_from_env() -> dict:
    """从环境变量构建可选筛选 dict（对应 form1 表单字段）"""
    filters: dict[str, str] = {}
    # 状态：已锁 / 未锁
    v = env("FILTER_SXZHUANTAI")
    if v:
        filters["sxzhuantai"] = v
    # 货权：拥有 / 待赎
    v = env("FILTER_HUOQUAN")
    if v:
        filters["huoquan"] = v
    # 仓库
    v = env("FILTER_CANKU")
    if v:
        filters["canku"] = v
    # 品名
    v = env("FILTER_PINMIN")
    if v:
        filters["pinmin"] = v
    return filters


def main() -> int:
    username = env("DH_USERNAME")
    password = env("DH_PASSWORD")

    # 飞书相关
    fs_app_id = env("FEISHU_APP_ID")
    fs_app_secret = env("FEISHU_APP_SECRET")
    fs_folder_token = env("FEISHU_FOLDER_TOKEN")
    fs_webhook_url = env("FEISHU_WEBHOOK_URL")
    fs_webhook_secret = env("FEISHU_WEBHOOK_SECRET")

    delivery_mode = (env("DELIVERY_MODE") or "feishu").lower()

    if not username or not password:
        print("[错误] 缺少 DH_USERNAME / DH_PASSWORD")
        return 2

    # 1) 登录
    session = create_session()
    if not login(session, username, password):
        return 1

    # 2) 调用导出接口（点"导出按钮"）
    filters = build_filters_from_env()
    try:
        html_bytes, info = export_lindiao(session, filters=filters)
    except Exception as exc:
        print(f"[错误] 调用导出接口失败: {exc}")
        traceback.print_exc()
        return 3

    # 3) 解析 HTML 表格为 CSV
    try:
        csv_bytes, row_count, col_count = html_table_to_csv_bytes(html_bytes)
    except Exception as exc:
        print(f"[错误] HTML 表格解析失败: {exc}")
        traceback.print_exc()
        return 3

    print(f"[解析] CSV 生成: {row_count} 行 (含表头), {col_count} 列, "
          f"{len(csv_bytes)/1024:.1f} KB")

    now = datetime.datetime.now()
    filename = f"lindiao_{now.strftime('%Y%m%d_%H%M')}.csv"

    filter_desc = ""
    if filters:
        filter_desc = " | 筛选: " + ", ".join(f"{k}={v}" for k, v in filters.items())
    else:
        filter_desc = " | 筛选: 无(全部)"

    summary_lines = [
        f"导出时间: {now.strftime('%Y-%m-%d %H:%M:%S')}",
        f"数据来源: {LINDIAO_PAGE}{filter_desc}",
        f"数据量: {row_count - 1} 条 (表头 {col_count} 列)",
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

            # 通知
            notify_lines = [f"✅ 临调库存导出完成"] + summary_lines
            notify_lines.append(f"文件位置: 飞书云盘指定文件夹")
            if file_token:
                notify_lines.append(f"file_token: {file_token}")
            notify_text = "\n".join(notify_lines)

            if fs_webhook_url:
                try:
                    feishu_send_bot_text(fs_webhook_url, fs_webhook_secret, notify_text)
                except Exception as exc:
                    print(f"[警告] 通知发送失败，但上传已完成: {exc}")
            else:
                print("\n[通知内容]")
                print(notify_text)
        except Exception as exc:
            print(f"[错误] 飞书交付失败: {exc}")
            traceback.print_exc()
            return 4
        return 0

    print(f"[错误] 未识别的 DELIVERY_MODE: {delivery_mode}")
    return 5


if __name__ == "__main__":
    exit(main())
