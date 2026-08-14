#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火钢城系统 - 开票申请单（待确认）导出 → 飞书云盘 / 本地 / 邮件

数据来源：开票申请单 页面（销项发票申请单 v_xfpdan）
接口：POST /model/admin/caiwu/m_fapiao/getlist  (分页)
筛选：zhuantai = "待确认"

流程：
  1. 自动登录 https://erpa.donghuo.vip（ddddocr 识别验证码）
  2. 分页调用接口拉取所有 "待确认" 状态的开票申请单
  3. 生成 CSV (UTF-8-SIG，避免 Excel 乱码)
  4. 上传 CSV 到飞书云盘指定文件夹 / 或本地保存 / 或邮件发送

环境变量（必填）：
  DH_USERNAME           erpa 登录账号
  DH_PASSWORD           erpa 登录密码

环境变量（交付方式，四选一，默认 feishu）：
  DELIVERY_MODE         feishu(默认) / local / mail / bitable

环境变量（飞书交付，DELIVERY_MODE=feishu 时必填）：
  FEISHU_APP_ID         飞书自建应用 App ID
  FEISHU_APP_SECRET     飞书自建应用 App Secret
  FEISHU_FOLDER_TOKEN   目标文件夹 token (fldcn...)
（可选飞书通知）
  FEISHU_WEBHOOK_URL    飞书机器人 Webhook
  FEISHU_WEBHOOK_SECRET 机器人签名密钥

环境变量（飞书多维表格交付，DELIVERY_MODE=bitable 时必填）：
  FEISHU_APP_ID         同上
  FEISHU_APP_SECRET     同上
  BITABLE_APP_TOKEN     多维表格 app_token (Mw62...)
  BITABLE_TABLE_ID      目标 table id (tbl...)，默认 tblfqRNfFg3NUcTo(滚动表)
  BITABLE_VIEW_ID       目标 view id (可选，vew9cZGRa2=放入视图)
（可选）
  BITABLE_DEDUP         按"申请单号"去重: 1(默认)=跳过重复 0=即使重复也追加

环境变量（邮件交付，DELIVERY_MODE=mail 时必填）：
  MAIL_TO               收件人邮箱
  SMTP_HOST             SMTP 服务器
  SMTP_PORT             SMTP 端口 (默认 465)
  SMTP_USER             SMTP 账号
  SMTP_PASSWORD         SMTP 密码
（可选邮件）
  MAIL_FROM             发件人显示邮箱 (默认 SMTP_USER)
  SMTP_USE_SSL          是否 SSL (默认 1=是，0=STARTTLS)

环境变量（可选筛选）：
  FILTER_ZHUANTAI       状态筛选：待确认(默认) / 已确认 / 空(全部)
  FILTER_SCOMPANY       所属公司筛选（精确匹配）
  FILTER_JSDANWEI       结算对方筛选（模糊匹配，取决于后端）
  FILTER_START_TIME     开始日期 (YYYY-MM-DD)
  FILTER_END_TIME       结束日期 (YYYY-MM-DD)
  EXPORT_DAYS           导出最近 N 天，0=全部；仅当未设置 START/END 时生效
"""

import os
import re
import io
import csv
import time
import json
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
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"

# 开票申请单列表 API (分页 POST)
KAIPIAO_API = f"{BASE_URL}/model/admin/caiwu/m_fapiao/getlist"
# 开票申请单页面（先访问预热 Referer / 会话状态）
KAIPIAO_PAGE = f"{BASE_URL}/view/admin/xiaoshou/v_xfpdan"


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


# ---------------- 开票申请单数据拉取 ----------------

def _try_parse_json(text: str):
    try:
        return json.loads(text)
    except Exception:
        return None


def _parse_date(text: str) -> str:
    """规范化日期字符串为 YYYY-MM-DD，无法解析时返回空串"""
    if not text:
        return ""
    text = str(text).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d %H:%M"):
        try:
            return datetime.datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text


def _row_in_range(row: dict, since_date: str, until_date: str) -> bool:
    """根据"开票日期"判断是否在 [since_date, until_date] 范围内"""
    if not since_date and not until_date:
        return True
    d = _parse_date(str(row.get("开票日期", "")))
    if not d:
        return True
    if since_date and d < since_date:
        return False
    if until_date and d > until_date:
        return False
    return True


def build_filters_from_env() -> dict:
    """从环境变量构建接口筛选 dict

    zhuantai 规则（避免 env() 默认空串无法区分"未设置"与"明确设空"）：
      - 未设置 FILTER_ZHUANTAI         → 默认 "待确认"（用户本任务要求）
      - 明确设置 FILTER_ZHUANTAI=""    → 不加筛选(全部)
      - 设置为其他值(如"已确认")        → 用该值筛选
    """
    filters: dict[str, str] = {}

    # 状态：使用 os.environ.get 区分 None(未设) 与 ""(明确空串) 与 具体值
    zhuantai_raw = os.environ.get("FILTER_ZHUANTAI")
    if zhuantai_raw is None:
        filters["zhuantai"] = "待确认"
    elif zhuantai_raw == "":
        pass  # 明确设空 => 不加入筛选 => 全部
    else:
        zhuantai_val = zhuantai_raw.strip()
        if zhuantai_val:
            filters["zhuantai"] = zhuantai_val

    # 所属公司
    v = env("FILTER_SCOMPANY")
    if v:
        filters["scompany"] = v

    # 结算对方
    v = env("FILTER_JSDANWEI")
    if v:
        filters["jsdanwei"] = v

    # 日期（START/END 优先级高于 EXPORT_DAYS）
    start = env("FILTER_START_TIME")
    end = env("FILTER_END_TIME")
    if start:
        filters["start_time"] = start
    if end:
        filters["end_time"] = end

    return filters


def get_kaipiao_records(session: requests.Session,
                        filters: dict | None = None,
                        page_size: int = 50) -> list[dict]:
    """
    调用真实接口分页拉取开票申请单数据。

    接口：POST /model/admin/caiwu/m_fapiao/getlist
    参数：page=页码, limit=每页条数, zhuantai=状态 等筛选字段
    响应：JSON，root=数据数组，pgtotal=总页数，rtotal=总条数
    """
    filters = filters or {}

    # 决定本地日期过滤（当没设 START/END 但设了 EXPORT_DAYS 时）
    local_since = ""
    local_until = ""
    if not filters.get("start_time") and not filters.get("end_time"):
        export_days = int(env("EXPORT_DAYS", "0") or "0")
        if export_days > 0:
            since = datetime.date.today() - datetime.timedelta(days=export_days)
            local_since = since.strftime("%Y-%m-%d")
            print(f"[本地日期过滤] 仅保留 开票日期 >= {local_since} 的记录 "
                  f"(EXPORT_DAYS={export_days})")

    print(f"[开票申请单] 开始拉取 (筛选={filters})")

    # 先访问页面预热 Referer / 会话状态
    try:
        session.get(KAIPIAO_PAGE, timeout=15)
    except Exception as exc:
        print(f"  [警告] 访问页面预热线: {exc}")

    # POST 数据基础
    def build_post_data(page: int) -> dict:
        data = {"page": page, "limit": page_size}
        for k, v in filters.items():
            if v:
                data[k] = v
        return data

    all_rows: list[dict] = []

    # 拉取第一页
    page = 1
    data1 = build_post_data(page)
    resp = session.post(KAIPIAO_API, data=data1, timeout=30,
                        headers={"X-Requested-With": "XMLHttpRequest"})
    parsed = _try_parse_json(resp.text)
    if not parsed or not isinstance(parsed, dict) or "root" not in parsed:
        print(f"[错误] 接口返回非预期格式 (第1页): {resp.text[:300]}")
        return []

    total_pages = int(parsed.get("pgtotal") or 1)
    total_count = int(parsed.get("rtotal") or 0)
    print(f"  接口返回: 共 {total_count} 条, {total_pages} 页 (每页 {page_size})")

    rows = parsed.get("root") or []
    kept = [r for r in rows if _row_in_range(r, local_since, local_until)]
    all_rows.extend(kept)
    print(f"  第 {page}/{total_pages} 页: 取 {len(rows)} 行, 命中 {len(kept)} 行, "
          f"累计 {len(all_rows)} 行")

    # 当按开票日期降序时，如果本页最后一条都早于 since_date，可以提前终止
    stop_early = False
    if local_since and rows:
        last_date = _parse_date(str(rows[-1].get("开票日期", "")))
        if last_date and last_date < local_since:
            stop_early = True

    # 翻页
    while not stop_early and page < total_pages:
        page += 1
        try:
            data = build_post_data(page)
            resp = session.post(KAIPIAO_API, data=data, timeout=30,
                                headers={"X-Requested-With": "XMLHttpRequest"})
            parsed = _try_parse_json(resp.text)
            if not parsed or not isinstance(parsed, dict):
                print(f"  第 {page} 页解析失败，停止")
                break
            rows = parsed.get("root") or []
            if not rows:
                break
            kept = [r for r in rows if _row_in_range(r, local_since, local_until)]
            all_rows.extend(kept)
            print(f"  第 {page}/{total_pages} 页: 取 {len(rows)} 行, "
                  f"命中 {len(kept)} 行, 累计 {len(all_rows)} 行")

            if local_since and rows:
                last_date = _parse_date(str(rows[-1].get("开票日期", "")))
                if last_date and last_date < local_since:
                    print(f"  已到达 {local_since} 之前的数据，提前终止")
                    stop_early = True

            time.sleep(0.3)  # 友好限速
        except Exception as exc:
            print(f"  抓取第 {page} 页失败: {exc}")
            break

    print(f"[开票申请单] 完成，共获取 {len(all_rows)} 条记录")
    return all_rows


# ---------------- CSV ----------------

# CSV输出字段：(原始字段名 → 输出字段名)，未列出的字段会被丢弃
# 删除: id, 发票日期, 我方名称, 发票类型
# 重命名: 发票对方→发票抬头, 发票数量→重量, 发票金额→金额, 经办人→提交人, 新增时间→申请时间
CSV_OUTPUT_FIELDS: list[tuple[str, str]] = [
    ("状态", "状态"),
    ("所属公司", "所属公司"),
    ("申请单号", "申请单号"),
    ("开票日期", "开票日期"),
    ("结算对方", "结算对方"),
    ("发票对方", "发票抬头"),
    ("发票数量", "重量"),
    ("发票金额", "金额"),
    ("发票号码", "发票号码"),
    ("备注", "备注"),
    ("业务人员", "业务人员"),
    ("经办人", "提交人"),
    ("新增时间", "申请时间"),
]


def rows_to_csv_bytes(rows: list[dict]) -> bytes:
    """将 list[dict] 写入 CSV（UTF-8-SIG 避免 Excel 乱码），返回字节

    字段清洗：只保留 CSV_OUTPUT_FIELDS 中定义的13列，并按飞书字段名重命名
    """
    if not rows:
        return b""
    out_field_names = [new for _, new in CSV_OUTPUT_FIELDS]
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=out_field_names, extrasaction="ignore")
    writer.writeheader()
    for r in rows:
        out_row = {}
        for orig, new in CSV_OUTPUT_FIELDS:
            v = r.get(orig, "")
            out_row[new] = "" if v is None else str(v)
        writer.writerow(out_row)
    return buf.getvalue().encode("utf-8-sig")


# ---------------- 飞书云盘 ----------------

def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    """用自建应用 App ID/Secret 换 tenant_access_token"""
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


def feishu_upload_to_folder(token: str, folder_token: str,
                            file_bytes: bytes, filename: str,
                            max_size_mb: int = 20) -> dict:
    """通过 drive/v1/files/upload_all 上传文件到飞书云盘指定文件夹（≤20MB）"""
    size = len(file_bytes)
    if size == 0:
        raise ValueError("上传文件为空")
    if size > max_size_mb * 1024 * 1024:
        raise ValueError(f"文件 {size/1024/1024:.1f}MB 超过 upload_all 上限 "
                         f"{max_size_mb}MB")
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
    print(f"[飞书云盘] 上传成功, file_token={file_token}, "
          f"name={file_info.get('name')}")
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
    print("[飞书通知] 发送成功")


# ---------------- 飞书多维表格（追加记录） ----------------

# CSV字段名 → (飞书字段名, 字段类型)  — field_name 来自 GET /fields 接口确认
BITABLE_FIELD_MAPPING: dict[str, tuple[str, str]] = {
    "id":          ("序号", "number"),
    "状态":         ("状态", "text"),
    "申请单号":      ("申请单号", "text"),
    "开票日期":      ("开票日期", "text"),
    "结算对方":      ("结算对方", "text"),
    "发票对方":      ("发票抬头", "text"),
    "发票数量":      ("重量", "number"),
    "发票金额":      ("金额", "number"),
    "发票号码":      ("发票号码", "text"),
    "备注":         ("备注", "text"),
    "业务人员":      ("业务人员", "text"),
    "经办人":        ("提交人", "text"),
    "所属公司":      ("所属公司", "text"),
    "新增时间":      ("排票日期", "datetime"),
    # 字段删除：id、发票日期、我方名称、发票类型（CSV清洗时已排除，但原始API数据中id/发票日期/我方名称仍有，在此跳过）
    # 发票日期 没有对应字段，跳过
    # 我方名称 没有对应字段，跳过
    # 发票类型 没有对应字段，跳过
}


def _btable_to_text(val) -> str | None:
    s = str(val).strip() if val is not None else ""
    if not s:
        return None
    return s


def _btable_to_number(val) -> int | float | None:
    s = str(val).strip() if val is not None else ""
    if not s:
        return None
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except ValueError:
        return None


def _btable_to_datetime(val) -> int | None:
    """支持 YYYY-MM-DD, YYYY-MM-DD HH:MM:SS 等 → 毫秒时间戳"""
    s = str(val).strip() if val is not None else ""
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                "%Y/%m/%d %H:%M:%S", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return int(datetime.datetime.strptime(s, fmt).timestamp() * 1000)
        except ValueError:
            continue
    return None


def _btable_convert(csv_val, ftype: str):
    if ftype == "text":
        return _btable_to_text(csv_val)
    if ftype == "number":
        return _btable_to_number(csv_val)
    if ftype == "datetime":
        return _btable_to_datetime(csv_val)
    return None


def bitable_find_existing_apply_ids(token: str, app_token: str,
                                    table_id: str,
                                    apply_ids: list[str]) -> set[str]:
    """按指定的申请单号列表，用条件查询检查哪些已存在（1次API调用即可）
    使用 POST /records/search + filter condition 按申请单号精确匹配。
    """
    if not apply_ids:
        return set()

    # 去重+去空
    unique_ids = list({aid.strip() for aid in apply_ids if aid and aid.strip()})
    if not unique_ids:
        return set()

    no_proxy = {"http": None, "https": None}
    existing: set[str] = set()
    page_token = ""
    seen_tokens: set[str] = set()
    max_pages = 10

    # 飞书 filter: 对每个申请单号用 "is" 操作符，conjunction=or
    # value 格式: ["XFP2026-1293"]
    conditions = [
        {"field_name": "申请单号", "operator": "is", "value": [aid]}
        for aid in unique_ids
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
            apply_no = f.get("申请单号") or ""
            if isinstance(apply_no, list):
                for seg in apply_no:
                    if isinstance(seg, dict) and seg.get("text"):
                        existing.add(str(seg["text"]).strip())
                        break
            elif isinstance(apply_no, str):
                if apply_no.strip():
                    existing.add(apply_no.strip())
        has_more = d.get("has_more", False)
        page_token = d.get("page_token") or ""
        if not has_more or not page_token:
            break

    print(f"[多维表格] 待查 {len(unique_ids)} 个申请单号，已存在 {len(existing)} 个")
    return existing


def bitable_batch_create(token: str, app_token: str, table_id: str,
                         records: list[dict]) -> list[dict]:
    """批量创建记录（records = [{"fields": {...}}]，单批≤500）"""
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


def bitable_append_rows(token: str, app_token: str, table_id: str,
                        view_id: str, rows: list[dict],
                        dedup: bool = True) -> dict:
    """
    将 rows（懂火返回的 list[dict]）按映射追加到多维表格
    去重逻辑：先从 rows 提取申请单号，用条件查询检查哪些已存在，存在的跳过。
    返回 {"created": n, "skipped": n, "record_ids": [...]}
    """
    existing_ids: set[str] = set()
    if dedup:
        apply_ids = [str(r.get("申请单号") or "").strip() for r in rows]
        try:
            existing_ids = bitable_find_existing_apply_ids(
                token, app_token, table_id, apply_ids)
        except Exception as exc:
            print(f"[警告] 按申请单号条件查询失败（将不做去重）: {exc}")

    records_payload: list[dict] = []
    skipped = 0
    for row in rows:
        apply_no = str(row.get("申请单号") or "").strip()
        if dedup and apply_no and apply_no in existing_ids:
            skipped += 1
            continue
        fields: dict = {}
        for csv_col, (bt_name, ftype) in BITABLE_FIELD_MAPPING.items():
            v = _btable_convert(row.get(csv_col, ""), ftype)
            if v is not None:
                fields[bt_name] = v
        records_payload.append({"fields": fields})

    print(f"[多维表格] 本次将追加 {len(records_payload)} 条 "
          f"（跳过重复 {skipped} 条）")

    created_ids: list[str] = []
    # 飞书 batch_create 单次上限 500 条
    for i in range(0, len(records_payload), 500):
        batch = records_payload[i:i + 500]
        created = bitable_batch_create(token, app_token, table_id, batch)
        created_ids.extend([(c or {}).get("record_id", "") for c in created])
        print(f"  批次 {i // 500 + 1}: 成功写入 {len(created)} 条")
        if i + 500 < len(records_payload):
            time.sleep(0.5)

    return {"created": len(created_ids), "skipped": skipped,
            "record_ids": created_ids}


# ---------------- 飞书多维表格：打铃通知 ----------------

BELL_TABLE_ID = "tbl9S5hzVd4IDB6y"  # '打铃通知！！！'表
BELL_RECORD_ID = "recv2rrwvSvzWk"   # 描述='更新完毕' 的那条固定记录


def bitable_update_bell_time(token: str, app_token: str) -> str:
    """更新'打铃通知！！！'表中固定记录的'最后打铃时间'为当下时间"""
    now_ms = int(datetime.datetime.now().timestamp() * 1000)
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{BELL_TABLE_ID}/records/{BELL_RECORD_ID}")
    body = {"fields": {"最后打铃时间": now_ms}}
    r = requests.put(url, headers={"Authorization": f"Bearer {token}"},
                     json=body, timeout=10,
                     proxies={"http": None, "https": None})
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"打铃通知更新失败: {json.dumps(data, ensure_ascii=False)[:800]}")
    print(f"[打铃通知] 已更新'最后打铃时间'={now_str}, record_id={BELL_RECORD_ID}")
    return BELL_RECORD_ID


# ---------------- 邮件（备用交付方式） ----------------

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


# ---------------- 主入口 ----------------

def main() -> int:
    username = env("DH_USERNAME")
    password = env("DH_PASSWORD")

    # 交付方式: "feishu" (默认) / "local" / "mail"
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
    fs_folder_token = env("FEISHU_FOLDER_TOKEN")
    fs_webhook_url = env("FEISHU_WEBHOOK_URL")
    fs_webhook_secret = env("FEISHU_WEBHOOK_SECRET")

    # ===== 飞书多维表格（bitable 模式） =====
    bt_app_token = env("BITABLE_APP_TOKEN")
    bt_table_id = env("BITABLE_TABLE_ID") or "tblfqRNfFg3NUcTo"
    bt_view_id = env("BITABLE_VIEW_ID") or "vew9cZGRa2"
    bt_dedup = env("BITABLE_DEDUP", "1") != "0"

    if not username or not password:
        print("[错误] 缺少 DH_USERNAME / DH_PASSWORD")
        return 2

    # 1) 登录
    session = create_session()
    if not login(session, username, password):
        return 1

    # 2) 构建筛选 & 拉取数据
    filters = build_filters_from_env()
    try:
        records = get_kaipiao_records(session, filters=filters)
    except Exception as exc:
        print(f"[错误] 拉取开票申请单数据失败: {exc}")
        traceback.print_exc()
        return 3
    finally:
        session.close()  # 关闭懂火 Session，避免连接池干扰飞书 API 调用

    # 3) 生成 CSV
    csv_bytes = rows_to_csv_bytes(records)

    now = datetime.datetime.now()
    filename = f"kaipiao_daiqueren_{now.strftime('%Y%m%d_%H%M')}.csv"

    filter_desc_parts = []
    zt = filters.get("zhuantai") or "(未筛选)"
    filter_desc_parts.append(f"状态={zt}")
    for k in ("scompany", "jsdanwei"):
        if filters.get(k):
            filter_desc_parts.append(f"{k}={filters[k]}")
    for k in ("start_time", "end_time"):
        if filters.get(k):
            filter_desc_parts.append(f"{k}={filters[k]}")
    ed = int(env("EXPORT_DAYS", "0") or "0")
    if ed > 0 and not filters.get("start_time") and not filters.get("end_time"):
        filter_desc_parts.append(f"最近{ed}天")
    filter_desc = " | 筛选: " + ", ".join(filter_desc_parts)

    summary_lines = [
        f"导出时间: {now.strftime('%Y-%m-%d %H:%M:%S')}",
        f"数据来源: {KAIPIAO_PAGE}{filter_desc}",
        f"数据量: {len(records)} 条",
        f"文件名: {filename}",
        f"文件大小: {len(csv_bytes)/1024:.1f} KB",
    ]
    print("\n[汇总]")
    for line in summary_lines:
        print(f"  {line}")

    if not records:
        print("[警告] 没有取到任何符合条件的开票申请单记录")

    # ========== 交付：本地 ==========
    if delivery_mode == "local":
        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        with open(out_path, "wb") as fp:
            fp.write(csv_bytes)
        print(f"\n[本地保存] 已保存到 {out_path}")
        return 0

    # ========== 交付：飞书云盘 ==========
    if delivery_mode == "feishu":
        need_fs = fs_app_id and fs_app_secret and fs_folder_token
        if not need_fs:
            print("[跳过] 未配置完整飞书参数（FEISHU_APP_ID / FEISHU_APP_SECRET / "
                  "FEISHU_FOLDER_TOKEN），退化为本地保存")
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

            notify_lines = [f"✅ 开票申请单(待确认)导出完成"] + summary_lines
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

    # ========== 交付：邮件（备用） ==========
    if delivery_mode == "mail":
        subject = f"开票申请单(待确认)导出 - {now.strftime('%Y-%m-%d %H:%M')} (共{len(records)}条)"
        body = ("\n".join(summary_lines)
                + "\n(本邮件由自动任务发送，请勿直接回复)")
        if to_email and smtp_host and smtp_user and smtp_password:
            try:
                send_email_with_attachment(
                    subject=subject,
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
            print("[跳过] 未配置完整邮件参数（MAIL_TO / SMTP_HOST / SMTP_USER / "
                  "SMTP_PASSWORD），退化为本地保存")
            out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            with open(out_path, "wb") as fp:
                fp.write(csv_bytes)
            print(f"[本地保存] 已保存到 {out_path}")
        return 0

    # ========== 交付：飞书多维表格（按行追加 + 去重） ==========
    if delivery_mode == "bitable":
        if not (fs_app_id and fs_app_secret and bt_app_token):
            print("[错误] 交付 bitable 缺少参数: FEISHU_APP_ID / FEISHU_APP_SECRET / "
                  "BITABLE_APP_TOKEN")
            # 退化为本地保存
            out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            with open(out_path, "wb") as fp:
                fp.write(csv_bytes)
            print(f"[本地保存] 已保存到 {out_path}")
            return 0
        try:
            token = feishu_tenant_access_token(fs_app_id, fs_app_secret)
            result = bitable_append_rows(
                token=token,
                app_token=bt_app_token,
                table_id=bt_table_id,
                view_id=bt_view_id,
                rows=records,
                dedup=bt_dedup,
            )

            # 更新打铃通知的'最后打铃时间'为当下时间
            try:
                bitable_update_bell_time(token=token, app_token=bt_app_token)
            except Exception as exc:
                print(f"[警告] 打铃通知更新失败（不影响主流程）: {exc}")

            notify_lines = [f"✅ 开票申请单 → 飞书多维表格追加完成"] + summary_lines
            notify_lines.append(f"目标表: app_token={bt_app_token} table_id={bt_table_id}")
            if bt_view_id:
                notify_lines.append(f"视图: {bt_view_id} (仅用于去重时参考视图筛选)")
            notify_lines.append(f"去重: {'开' if bt_dedup else '关'}")
            notify_lines.append(f"新增 {result['created']} 条 / 跳过重复 {result['skipped']} 条")
            notify_text = "\n".join(notify_lines)

            if fs_webhook_url:
                try:
                    feishu_send_bot_text(fs_webhook_url, fs_webhook_secret, notify_text)
                except Exception as exc:
                    print(f"[警告] 通知发送失败，但写入已完成: {exc}")
            else:
                print("\n[完成通知]")
                print(notify_text)
        except Exception as exc:
            print(f"[错误] 多维表格写入失败: {exc}")
            traceback.print_exc()
            return 4
        return 0

    print(f"[错误] 未识别的 DELIVERY_MODE: {delivery_mode}")
    return 5


if __name__ == "__main__":
    exit(main())
