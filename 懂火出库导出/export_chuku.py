#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火钢城系统 - 出库记录自动导出脚本
功能：
  1. 自动登录 https://erpa.donghuo.vip （识别验证码）
  2. 导出出库记录数据为 CSV
  3. 通过邮件发送 CSV 附件到指定邮箱
"""

import os
import re
import io
import csv
import time
import json
import base64
import smtplib
import datetime
import traceback
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_URL = "https://erpa.donghuo.vip"


def create_session() -> requests.Session:
    """创建带重试机制的 requests Session"""
    session = requests.Session()
    retry = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[500, 502, 503, 504, 429],
        allowed_methods=["POST", "GET"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "zh-CN,zh;q=0.9",
    })
    return session


def recognize_captcha(image_bytes: bytes) -> str:
    """
    识别图形验证码。
    优先使用 ddddocr（需 pip install ddddocr），失败或未安装时返回空字符串。
    """
    try:
        import ddddocr  # type: ignore

        ocr = ddddocr.DdddOcr(show_ad=False)
        result = ocr.classification(image_bytes)
        code = result.strip().replace(" ", "")
        print(f"[验证码识别] 结果: {code}")
        return code
    except ImportError:
        print("[警告] 未安装 ddddocr，跳过验证码自动识别")
        return ""
    except Exception as exc:
        print(f"[错误] 验证码识别失败: {exc}")
        return ""


def login(session: requests.Session, username: str, password: str,
          max_attempts: int = 10) -> bool:
    """
    登录懂火钢城系统，自动重试验证码。
    登录接口: POST /controller/admin/c_longin/index
    """
    login_url = f"{BASE_URL}/controller/admin/c_longin/index"
    captcha_url = f"{BASE_URL}/common/captcha"

    for attempt in range(1, max_attempts + 1):
        print(f"[登录] 尝试 {attempt}/{max_attempts} ...")

        # 1) 拉取验证码图片（建立/刷新 Session Cookie）
        img_resp = session.get(captcha_url, timeout=15)
        if img_resp.status_code != 200:
            print(f"  获取验证码图片失败: HTTP {img_resp.status_code}")
            time.sleep(2)
            continue
        captcha_image_bytes = img_resp.content

        # 2) 识别验证码
        captcha_code = recognize_captcha(captcha_image_bytes)
        if not captcha_code:
            print("  验证码识别失败，将在登录时再尝试刷新")
            captcha_code = "1234"  # 占位，必然失败，但可触发失败后刷新

        # 3) 提交登录
        data = {
            "u_name": username,
            "u_pass": password,
            "captcha": captcha_code,
        }
        resp = session.post(login_url, data=data, timeout=15)
        text = resp.text.strip()
        print(f"  登录响应 (前300字符): {text[:300]}")

        try:
            result = json.loads(text)
            if isinstance(result, dict) and str(result.get("code")) == "200":
                print("[登录] 成功")
                return True
            else:
                msg = result.get("msg") or result.get("message") or text[:100]
                print(f"  登录失败: {msg}")
        except json.JSONDecodeError:
            if "验证码" in text:
                print("  登录失败：验证码错误")
            else:
                print(f"  登录失败（非JSON响应）: {text[:200]}")

        time.sleep(2)

    print(f"[登录] 已达到最大尝试次数 {max_attempts}，登录失败")
    return False


# 出库记录数据接口（已通过实际抓包确认）
OUTBOUND_API = f"{BASE_URL}/model/admin/xiaoshou/m_xiaoshou/xjilulist"
# 出库记录页面（先访问一次建立 Referer / 会话状态）
OUTBOUND_PAGE = f"{BASE_URL}/view/admin/xiaoshou/v_xjlall"


def _parse_date(text: str) -> str:
    """规范化日期字符串为 YYYY-MM-DD，无法解析时返回空串"""
    if not text:
        return ""
    text = str(text).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text


def _row_in_range(row: dict, since_date: str) -> bool:
    """根据"出库日期"判断是否 >= since_date"""
    if not since_date:
        return True
    d = _parse_date(str(row.get("出库日期", "")))
    if not d:
        return True  # 无法解析日期时保留，避免误删
    return d >= since_date


def get_outbound_records(session: requests.Session,
                         export_days: int = 30,
                         page_size: int = 50) -> list[dict]:
    """
    调用真实接口拉取出库记录。

    接口：POST /model/admin/xiaoshou/m_xiaoshou/xjilulist
    参数：page=页码, limit=每页条数（默认30，接口上限未明，50 稳妥）
    响应：JSON，root=数据数组，pgtotal=总页数，rtotal=总条数

    export_days: 导出最近 N 天数据；0 表示全部历史
    """
    print(f"[出库记录] 开始拉取 (export_days={export_days})")

    # 先访问一次出库记录页面，建立 Referer / 会话状态
    try:
        session.get(OUTBOUND_PAGE, timeout=15)
    except Exception as exc:
        print(f"  访问出库记录页面警告: {exc}")

    # 计算起始日期（用于本地过滤；接口返回按"出库日期"降序）
    since_date = ""
    if export_days > 0:
        since = datetime.date.today() - datetime.timedelta(days=export_days)
        since_date = since.strftime("%Y-%m-%d")
        print(f"  仅保留 出库日期 >= {since_date} 的记录")

    # 拉取第一页，拿到总页数
    all_rows: list[dict] = []
    page = 1
    resp = session.post(OUTBOUND_API, data={"page": page, "limit": page_size},
                        timeout=30,
                        headers={"X-Requested-With": "XMLHttpRequest"})
    parsed = _try_parse_json(resp.text)
    if not parsed or not isinstance(parsed, dict) or "root" not in parsed:
        print(f"[错误] 接口返回非预期格式: {resp.text[:200]}")
        return []

    total_pages = int(parsed.get("pgtotal") or 1)
    total_count = int(parsed.get("rtotal") or 0)
    print(f"  接口返回: 共 {total_count} 条, {total_pages} 页 (每页 {page_size})")

    rows = parsed.get("root") or []
    kept = [r for r in rows if _row_in_range(r, since_date)]
    all_rows.extend(kept)
    print(f"  第 {page}/{total_pages} 页: 取 {len(rows)} 行, 命中 {len(kept)} 行, 累计 {len(all_rows)} 行")

    # 由于数据按"出库日期"降序，一旦某页全部早于 since_date，可提前终止
    stop_early = False
    if since_date and rows:
        last_row_date = _parse_date(str(rows[-1].get("出库日期", "")))
        if last_row_date and last_row_date < since_date:
            stop_early = True

    # 翻页
    while not stop_early and page < total_pages:
        page += 1
        try:
            resp = session.post(OUTBOUND_API,
                                data={"page": page, "limit": page_size},
                                timeout=30,
                                headers={"X-Requested-With": "XMLHttpRequest"})
            parsed = _try_parse_json(resp.text)
            if not parsed or not isinstance(parsed, dict):
                print(f"  第 {page} 页解析失败，停止")
                break
            rows = parsed.get("root") or []
            if not rows:
                break
            kept = [r for r in rows if _row_in_range(r, since_date)]
            all_rows.extend(kept)
            print(f"  第 {page}/{total_pages} 页: 取 {len(rows)} 行, 命中 {len(kept)} 行, 累计 {len(all_rows)} 行")

            if since_date and rows:
                last_row_date = _parse_date(str(rows[-1].get("出库日期", "")))
                if last_row_date and last_row_date < since_date:
                    print(f"  已到达 {since_date} 之前的数据，提前终止")
                    stop_early = True

            time.sleep(0.3)  # 友好限速
        except Exception as exc:
            print(f"  抓取第 {page} 页失败: {exc}")
            break

    print(f"[出库记录] 完成, 共获取 {len(all_rows)} 行")
    return all_rows


def _try_parse_json(text: str):
    try:
        return json.loads(text)
    except Exception:
        return None


# ---------- CSV / 邮件 ----------

def rows_to_csv_bytes(rows: list[dict]) -> bytes:
    """将 list[dict] 写入 CSV（UTF-8-SIG 避免 Excel 乱码），返回字节"""
    if not rows:
        return b""
    fields: list[str] = []
    for r in rows:
        for k in r.keys():
            if k not in fields:
                fields.append(k)
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=fields, extrasaction="ignore")
    writer.writeheader()
    for r in rows:
        writer.writerow({k: ("" if v is None else str(v)) for k, v in r.items()})
    return buf.getvalue().encode("utf-8-sig")


def send_email_with_attachment(subject: str, body: str, to_email: str,
                                smtp_host: str, smtp_port: int,
                                smtp_user: str, smtp_password: str,
                                attachment_bytes: bytes,
                                attachment_filename: str,
                                use_ssl: bool = True,
                                from_email: str | None = None) -> None:
    """发送带附件的邮件（作为备用交付方式）"""
    sender = from_email or smtp_user

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment_bytes)
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename=\"{attachment_filename}\"",
    )
    msg.attach(part)

    if use_ssl:
        server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30)
    else:
        server = smtplib.SMTP(smtp_host, smtp_port, timeout=30)
        server.starttls()
    try:
        server.login(smtp_user, smtp_password)
        server.sendmail(sender, [to_email], msg.as_string())
        print(f"[邮件] 已成功发送到 {to_email}")
    finally:
        try:
            server.quit()
        except Exception:
            pass


# ---------- 飞书云盘 + 通知 ----------

FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"


def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    """用自建应用 App ID/Secret 换 tenant_access_token，有效期 2 小时"""
    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url,
                      json={"app_id": app_id, "app_secret": app_secret},
                      timeout=15)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换 tenant_access_token 失败: {data}")
    token = str(data.get("tenant_access_token") or "")
    if not token:
        raise RuntimeError(f"响应中无 tenant_access_token: {data}")
    print(f"[飞书] 获取 tenant_access_token 成功 (len={len(token)})")
    return token


def feishu_upload_to_folder(token: str, folder_token: str,
                            file_bytes: bytes, filename: str,
                            max_size_mb: int = 20) -> dict:
    """
    通过 drive/v1/files/upload_all 把文件传到飞书云空间指定文件夹。
    返回接口 data（含 file_token / name 等）。

    doc: https://open.feishu.cn/document/server-docs/docs/drive-v1/file/upload_all
      - parent_type = "explorer"
      - parent_node = folder_token (形如 "fldcnXXXXXX")
      - size <= 20MB (单个请求上限)
    """
    if len(file_bytes) == 0:
        raise ValueError("上传的文件为空字节")
    size = len(file_bytes)
    if size > max_size_mb * 1024 * 1024:
        raise ValueError(
            f"文件大小 {size/1024/1024:.1f}MB 超过 upload_all 上限 {max_size_mb}MB"
        )
    if not folder_token:
        raise ValueError("缺少 FEISHU_FOLDER_TOKEN (目标文件夹 token)")

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
        raise RuntimeError(f"上传到飞书云盘失败: code={resp.get('code')}, "
                           f"msg={resp.get('msg')}, raw={resp}")
    file_info = resp.get("data") or {}
    file_token = file_info.get("file_token") or file_info.get("token") or ""
    print(f"[飞书云盘] 上传成功, file_token={file_token}, name={file_info.get('name')}")
    return file_info


def _feishu_sign(secret: str, timestamp: str) -> str:
    """
    飞书自定义机器人 Webhook 签名。
    string_to_sign = f"{timestamp}\n{secret}"
    HMAC-SHA256(key=secret, msg=string_to_sign) → Base64
    """
    import hmac
    import hashlib
    import base64 as _b64
    string_to_sign = f"{timestamp}\n{secret}"
    h = hmac.new(secret.encode("utf-8"),
                 string_to_sign.encode("utf-8"),
                 hashlib.sha256)
    return _b64.b64encode(h.digest()).decode("utf-8")


def feishu_send_bot_text(webhook_url: str, secret: str, text: str) -> None:
    """
    通过飞书自定义机器人 Webhook 发送文本消息。
    若 secret 为空则跳过签名。
    """
    if not webhook_url:
        raise ValueError("缺少 FEISHU_WEBHOOK_URL")

    body = {
        "msg_type": "text",
        "content": {"text": text},
    }
    if secret:
        timestamp = str(int(time.time()))
        sign = _feishu_sign(secret, timestamp)
        body["timestamp"] = timestamp
        body["sign"] = sign

    r = requests.post(webhook_url, json=body, timeout=15)
    resp = r.json()
    # 自定义机器人 code=0 为成功
    if resp.get("code") != 0 and resp.get("StatusCode") != 0:
        raise RuntimeError(f"飞书机器人通知失败: {resp}")
    print("[飞书通知] 发送成功")


# ---------- 主入口 ----------

def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def main() -> int:
    username = env("DH_USERNAME")
    password = env("DH_PASSWORD")

    # 导出最近 N 天数据；0 = 全部历史；默认 30 天
    export_days = int(env("EXPORT_DAYS", "30") or "30")

    # 交付方式: "feishu" (默认) 或 "mail"
    delivery_mode = (env("DELIVERY_MODE") or "feishu").lower()

    # ===== 邮件相关（备用） =====
    to_email = env("MAIL_TO")
    smtp_host = env("SMTP_HOST")
    smtp_port = int(env("SMTP_PORT", "465") or "465")
    smtp_user = env("SMTP_USER")
    smtp_password = env("SMTP_PASSWORD")
    mail_from = env("MAIL_FROM") or smtp_user
    use_ssl = env("SMTP_USE_SSL", "1") != "0"

    # ===== 飞书相关 =====
    fs_app_id = env("FEISHU_APP_ID")
    fs_app_secret = env("FEISHU_APP_SECRET")
    fs_folder_token = env("FEISHU_FOLDER_TOKEN")  # 目标文件夹 token (fldcn...)
    fs_webhook_url = env("FEISHU_WEBHOOK_URL")     # 通知机器人 Webhook
    fs_webhook_secret = env("FEISHU_WEBHOOK_SECRET")  # 机器人签名密钥（可选）

    if not username or not password:
        print("[错误] 缺少 DH_USERNAME / DH_PASSWORD 环境变量")
        return 2

    session = create_session()

    if not login(session, username, password, max_attempts=10):
        print("[退出] 登录失败")
        return 1

    records = get_outbound_records(session, export_days=export_days)
    if not records:
        print("[警告] 未取到任何出库记录数据")
        records = []

    csv_bytes = rows_to_csv_bytes(records)

    now = datetime.datetime.now()
    filename = f"chuku_{now.strftime('%Y%m%d_%H%M')}.csv"
    summary_subject = f"出库记录导出 - {now.strftime('%Y-%m-%d %H:%M')} (共{len(records)}条)"
    summary_body_lines = [
        f"导出时间: {now.strftime('%Y-%m-%d %H:%M:%S')}",
        f"数据来源: {BASE_URL}",
        f"数据量: {len(records)} 条",
        f"导出范围: 最近 {export_days} 天" if export_days > 0 else "导出范围: 全部历史",
        f"文件名: {filename}",
    ]

    # ========== 交付：飞书云盘 + 通知 ==========
    if delivery_mode == "feishu":
        need_fs = fs_app_id and fs_app_secret and fs_folder_token
        if not need_fs:
            print("[跳过] 未配置完整飞书参数（FEISHU_APP_ID / FEISHU_APP_SECRET / FEISHU_FOLDER_TOKEN），"
                  "退化为本地保存")
            out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            with open(out_path, "wb") as fp:
                fp.write(csv_bytes)
            print(f"[本地保存] 已保存到 {out_path}")
            return 0

        try:
            token = feishu_tenant_access_token(fs_app_id, fs_app_secret)
            file_info = feishu_upload_to_folder(token, fs_folder_token,
                                                 csv_bytes, filename)
            file_token = (file_info.get("file_token")
                          or file_info.get("token") or "")
            # 上传成功后发通知
            notify_lines = list(summary_body_lines)
            notify_lines.insert(0, f"✅ {summary_subject}")
            notify_lines.append(f"文件位置: 飞书云盘指定文件夹")
            if file_token:
                notify_lines.append(f"file_token: {file_token}")
            notify_text = "\n".join(notify_lines)

            if fs_webhook_url:
                try:
                    feishu_send_bot_text(fs_webhook_url, fs_webhook_secret, notify_text)
                except Exception as exc:
                    print(f"[警告] 飞书通知发送失败，但云盘上传已完成: {exc}")
            else:
                print("[跳过] 未配置 FEISHU_WEBHOOK_URL，跳过通知")
                print(notify_text)

        except Exception as exc:
            print(f"[错误] 飞书交付失败: {exc}")
            traceback.print_exc()
            return 4
        return 0

    # ========== 交付：邮件（备用） ==========
    if delivery_mode == "mail":
        body = "\n".join(summary_body_lines) + "\n(本邮件由自动任务发送，请勿直接回复)"
        if to_email and smtp_host and smtp_user and smtp_password:
            try:
                send_email_with_attachment(
                    subject=summary_subject,
                    body=body,
                    to_email=to_email,
                    smtp_host=smtp_host,
                    smtp_port=smtp_port,
                    smtp_user=smtp_user,
                    smtp_password=smtp_password,
                    attachment_bytes=csv_bytes or b"",
                    attachment_filename=filename,
                    use_ssl=use_ssl,
                    from_email=mail_from,
                )
            except Exception as exc:
                print(f"[错误] 发送邮件失败: {exc}")
                traceback.print_exc()
                return 3
        else:
            print("[跳过] 未配置完整邮件参数（MAIL_TO / SMTP_HOST / SMTP_USER / SMTP_PASSWORD），"
                  "退化为本地保存")
            out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            with open(out_path, "wb") as fp:
                fp.write(csv_bytes)
            print(f"[本地保存] 已保存到 {out_path}")
        return 0

    print(f"[错误] 未识别的 DELIVERY_MODE: {delivery_mode}")
    return 5


if __name__ == "__main__":
    exit(main())
