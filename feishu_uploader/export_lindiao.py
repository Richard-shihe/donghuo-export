#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火钢城系统 - 临调库存导出 → 飞书云盘 / 飞书多维表格

数据来源：库存查询 > 临调库存 页面（v_kucun_ld）的"导出"按钮
流程：
  1. 自动登录 https://erpa.donghuo.vip（ddddocr 识别验证码）
  2. 调用导出接口 /view/admin/excelbiao/kucunld，获取服务端返回的 .xls
     （本质是带 mso-number-format 样式的 HTML 表格，Excel 可直接打开）
  3. 解析 HTML 表格为 CSV（UTF-8-SIG 避免 Excel 乱码）
  4. 按需交付：
     - feishu（默认）：上传 CSV 到飞书云盘指定文件夹
     - bitable：将数据追加写入飞书多维表格（按字段类型转换，不去重直接 append）
     - local：仅本地保存 CSV
  5. 可选：飞书机器人 Webhook 通知

环境变量（必填）：
  DH_USERNAME           erpa 登录账号
  DH_PASSWORD           erpa 登录密码
  FEISHU_APP_ID         飞书自建应用 App ID
  FEISHU_APP_SECRET     飞书自建应用 App Secret

环境变量（按交付方式）：
  DELIVERY_MODE=feishu（默认）需要 FEISHU_FOLDER_TOKEN（目标文件夹 token fldcn...）
  DELIVERY_MODE=bitable 需要 BITABLE_APP_TOKEN（多维表格 app_token）
                              BITABLE_TABLE_ID（目标 table id tbl...）
                       注：飞书应用需对该多维表格有「可编辑」协作者权限
  DELIVERY_MODE=local   无需额外变量

环境变量（可选）：
  FILTER_SXZHUANTAI     状态筛选: 空(全部) / 已锁 / 未锁
  FILTER_HUOQUAN        货权筛选: 空(全部) / 拥有 / 待赎
  FILTER_CANKU          仓库筛选（精确匹配，如 "仲鼎库"）
  FILTER_PINMIN         品名筛选
  FEISHU_WEBHOOK_URL    飞书机器人 Webhook
  FEISHU_WEBHOOK_SECRET 机器人签名密钥
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

# 临调库存页面 JSON 列表分页接口（前端 grid 展示用）
# 格式: {"pgtotal":N, "rtotal":M, "root":[{row dict}, ...]}
# 比 Excel 导出接口更稳定，境外 IP 不会被过滤成 0 条
LINDIAO_LIST_API = f"{BASE_URL}/model/admin/xiaoshou/m_kucun/ld_kucun"

# 与 EXCEL 导出接口一致的 23 列 CSV 列头顺序（CSV 按此顺序写，保持向下兼容）
# 说明：
#   - list API 的 JSON 字段名是中文，和 Excel 导出的 <td> 表头基本一致
#   - list API 有 40+ 个字段，CSV 只保留和旧 Excel 相同的 23 列
#   - list API 里叫 "重量(吨)"，Excel 导出版叫 "重量" → 下面映射处理
#   - list API 缺少 "销售单价" 列（字段可售件数/可售重量存在但销售单价值为 0），保持空
CSV_HEADER_ORDER = [
    "所属公司", "货权", "品名", "规格", "材质", "产地", "等级", "锌层",
    "涂料", "结构", "颜色", "件(张)数", "米数", "重量",
    "销售单价", "仓库", "库位号", "捆包号", "合同号", "车船号",
    "提单号", "备注", "入库日期",
]

# list API 字段名 → CSV 字段名映射（只处理不一致的）
_LIST2CSV_ALIAS = {
    "重量(吨)": "重量",
}

# bitable 写入时用的字段（和飞书多维表格实际字段对齐），从 rows 保留这些 key 即可
# 说明：CSV_HEADER_ORDER 23 列正好对应多维表格里的 23 个字段
BITABLE_WRITE_FIELDS = CSV_HEADER_ORDER


def _preheat_pages(session: requests.Session) -> None:
    """访问框架页和子页，预热会话状态（与 Excel 导出保持一致行为）"""
    try:
        session.get(LINDIAO_FRAME_PAGE, timeout=15)
        session.get(LINDIAO_PAGE, timeout=15)
    except Exception as exc:
        print(f"  [警告] 访问页面预热线: {exc}")


def export_lindiao(session: requests.Session,
                  filters: dict | None = None,
                  timeout: int = 120) -> tuple[bytes, dict]:
    """
    调用"导出按钮"对应接口（Excel 导出），返回 (HTML 表格字节, 元信息)。

    接口：POST /view/admin/excelbiao/kucunld
    返回：Content-Type=Application/x-msexcel，本质是 HTML 表格伪装为 .xls

    注意：境外 IP（如 GitHub Actions Azure 美国机房）访问时，懂火服务端可能返回
    "只有表头、0 条数据"的空结果。此场景应优先使用 export_lindiao_listapi()。
    """
    _preheat_pages(session)

    data = filters or {}
    print(f"[导出-Excel] POST {EXPORT_API} (筛选: {data or '无'})")
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


def export_lindiao_listapi(session: requests.Session,
                           filters: dict | None = None,
                           timeout: int = 120,
                           ) -> tuple[list[dict], bytes, int, int]:
    """
    通过列表 JSON API 拉取临调库存（替代 Excel 导出接口的优先方案）。

    接口：POST /model/admin/xiaoshou/m_kucun/ld_kucun
    必须带上的 body：page, limit, sxzhuantai, huoquan, canku, pinmin
      （PHP 代码直接引用 $sxzhuantai 等变量，不传会 Fatal error）

    返回：
      (rows, csv_bytes, row_count_with_header, col_count)
        rows                 : list[dict]，每个 dict 的 key 是 CSV_HEADER_ORDER 中的字段名
                               （与 Excel 导出 html_table_to_csv_bytes 的 rows 格式一致）
        csv_bytes            : 按 CSV_HEADER_ORDER 写的 CSV (UTF-8-SIG)
        row_count_with_header: len(rows)+1（与 html_table_to_csv_bytes 的返回值对齐）
        col_count            : len(CSV_HEADER_ORDER)（与 Excel 导出 23 列一致）
    """
    _preheat_pages(session)

    user_filters = filters or {}
    # 列表 API PHP 代码要求必须存在这 4 个筛选字段（空值 = 不过滤）
    payload = {
        "page": 1,
        "limit": 500,
        "sxzhuantai": user_filters.get("sxzhuantai", "") or "",
        "huoquan":    user_filters.get("huoquan", "")    or "",
        "canku":      user_filters.get("canku", "")      or "",
        "pinmin":     user_filters.get("pinmin", "")     or "",
    }
    print(f"[导出-ListAPI] POST {LINDIAO_LIST_API} (筛选: {user_filters or '无'})")

    all_rows: list[dict] = []
    page = 1
    while True:
        payload["page"] = page
        r = session.post(LINDIAO_LIST_API, data=payload, timeout=timeout)
        if r.status_code != 200:
            raise RuntimeError(f"列表 API 返回 HTTP {r.status_code}")
        # 懂火返回的 Content-Type=text/html，但内容其实是 JSON
        try:
            data = json.loads(r.text)
        except json.JSONDecodeError as exc:
            # 不是 JSON → PHP error / login redirect / 被风控 → 抛异常让上层回退 Excel 导出
            raise RuntimeError(
                f"列表 API 返回非 JSON: {exc}; resp前300={r.text[:300]!r}"
            ) from exc

        root = data.get("root") or []
        if not isinstance(root, list):
            raise RuntimeError(f"列表 API root 不是 list: {type(root).__name__}")
        rtotal = data.get("rtotal")
        pgtotal = data.get("pgtotal")
        all_rows.extend(root)
        print(f"  page {page}: {len(root)} 条, 累计 {len(all_rows)} 条, rtotal={rtotal}, pgtotal={pgtotal}")

        if len(root) < payload["limit"]:
            break
        if pgtotal and page >= pgtotal:
            break
        page += 1
        time.sleep(0.15)

    # 把 list API 字段名（可能含 "重量(吨)" 等）→ 对齐 CSV_HEADER_ORDER 的字段名
    rows_aligned: list[dict] = []
    for raw in all_rows:
        if not isinstance(raw, dict):
            continue
        # 先做 alias 映射
        mapped = {}
        for k, v in raw.items():
            new_k = _LIST2CSV_ALIAS.get(k, k)
            # list API 数字字段有些是 int/float（如 重量(吨)=27.11），CSV 统一转字符串
            mapped[new_k] = "" if v is None else str(v)
        # 按 CSV_HEADER_ORDER 取字段，缺失的填空（保证 23 列和老 CSV 一致）
        row_out = {h: mapped.get(h, "") for h in CSV_HEADER_ORDER}
        rows_aligned.append(row_out)

    # 生成 CSV (UTF-8-SIG，和 Excel 导出保持一致)
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(CSV_HEADER_ORDER)
    for row in rows_aligned:
        writer.writerow([row[h] for h in CSV_HEADER_ORDER])
    csv_bytes = buf.getvalue().encode("utf-8-sig")

    row_count = len(rows_aligned) + 1  # +1 表头
    col_count = len(CSV_HEADER_ORDER)
    print(f"[解析] ListAPI: {len(rows_aligned)} 条数据")
    return rows_aligned, csv_bytes, row_count, col_count


def html_table_to_csv_bytes(html_bytes: bytes) -> tuple[bytes, int, int, list[dict]]:
    """
    将服务端返回的 HTML 表格（伪装 .xls）解析为 CSV 字节。

    返回：(csv 字节 UTF-8-SIG, 行数含表头, 列数, rows list[dict])
      rows: 第一行作为表头，后续每行 dict[str, str]，供 bitable 写入使用
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

    # 第一行作为表头，后续行转 dict
    headers = rows_data[0]
    rows: list[dict] = []
    for raw in rows_data[1:]:
        row_dict = {}
        for i, h in enumerate(headers):
            if i < len(raw):
                row_dict[h] = raw[i]
            else:
                row_dict[h] = ""
        rows.append(row_dict)

    return csv_bytes, len(rows_data), len(rows_data[0]), rows


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


# ---------------- 飞书多维表格（追加） ----------------

# 临调库存 CSV 字段名 → 字段类型（与多维表格字段同名直接映射）
# 文本字段：直接传字符串
# 数字字段：转 float/int（飞书 bitable 数字字段需要 number 类型）
# 日期字段：转毫秒时间戳
# 表里其他字段（公式 / lookup / auto_fill / 系统字段）不在此映射中，自动跳过
BITABLE_FIELD_TYPES: dict[str, str] = {
    "所属公司": "text", "货权": "text", "品名": "text", "规格": "text",
    "材质": "text", "产地": "text", "等级": "text", "锌层": "text",
    "涂料": "text", "结构": "text", "颜色": "text", "米数": "text",
    "仓库": "text", "库位号": "text", "捆包号": "text", "合同号": "text",
    "车船号": "text", "提单号": "text", "备注": "text",
    "件(张)数": "number", "重量": "number", "销售单价": "number",
    "入库日期": "datetime",
}

# batch_create 单次最多 500 条
BITABLE_BATCH_SIZE = 500


def _to_bitable_text(val) -> str:
    s = str(val).strip() if val is not None else ""
    return s


def _to_bitable_number(val):
    s = str(val).strip() if val is not None else ""
    if not s:
        return None
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return None


def _to_bitable_datetime(val):
    """日期字符串 → 毫秒时间戳；支持常见格式"""
    s = str(val).strip() if val is not None else ""
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return int(dt.timestamp() * 1000)
        except ValueError:
            continue
    return None


def _convert_field(val, ftype: str):
    if ftype == "text":
        return _to_bitable_text(val)
    if ftype == "number":
        return _to_bitable_number(val)
    if ftype == "datetime":
        return _to_bitable_datetime(val)
    return None


def bitable_batch_create(token: str, app_token: str, table_id: str,
                          records: list[dict]) -> list[dict]:
    """调 batch_create API，单次最多 500 条，自动分批"""
    # === DEBUG: 记录每次调用 ===
    import os as _os
    debug_file = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)),
                                "_debug_batch_create.log")
    with open(debug_file, "a", encoding="utf-8") as _fp:
        import datetime as _dt
        _fp.write(f"[{_dt.datetime.now()}] batch_create called with {len(records)} records\n")
        if records:
            sample_bundle = records[0].get("fields", {}).get("捆包号", "")
            _fp.write(f"  first record 捆包号: {sample_bundle!r}\n")
    print(f"  [DEBUG] batch_create called with {len(records)} records", flush=True)

    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/records/batch_create")
    headers = {"Authorization": f"Bearer {token}"}
    created: list[dict] = []
    total = len(records)
    for start in range(0, total, BITABLE_BATCH_SIZE):
        batch = records[start:start + BITABLE_BATCH_SIZE]
        r = requests.post(url, headers=headers,
                          json={"records": batch}, timeout=60)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(
                f"bitable 批量创建失败: code={data.get('code')}, "
                f"msg={data.get('msg')}, "
                f"raw={json.dumps(data, ensure_ascii=False)[:1200]}"
            )
        items = (data.get("data") or {}).get("records") or []
        created.extend(items)
        print(f"  [bitable] 写入 {start + len(batch)}/{total} (本次 {len(items)} 条)", flush=True)
    return created


def bitable_append_records(token: str, app_token: str, table_id: str,
                            rows: list[dict]) -> dict:
    """
    将临调库存 rows（解析出的 dict 列表）追加到多维表格。
    字段映射见 BITABLE_FIELD_TYPES；不在此映射的字段自动跳过。
    不去重，直接 append。
    """
    if not rows:
        print("[bitable] 无数据可写入")
        return {"created": 0}

    records_payload: list[dict] = []
    skipped_field_counts: dict[str, int] = {}
    for row in rows:
        fields: dict = {}
        for col_name, ftype in BITABLE_FIELD_TYPES.items():
            v = _convert_field(row.get(col_name, ""), ftype)
            # 文本字段空字符串、数字/日期字段 None 都跳过（不写入空值）
            if v is None or (isinstance(v, str) and not v):
                continue
            fields[col_name] = v
        # 记录哪些 CSV 字段未在映射中（仅第一次出现时打印）
        for col_name in row.keys():
            if col_name not in BITABLE_FIELD_TYPES:
                skipped_field_counts[col_name] = skipped_field_counts.get(col_name, 0) + 1
        records_payload.append({"fields": fields})

    if skipped_field_counts:
        print(f"[bitable] 跳过的未映射 CSV 字段: {skipped_field_counts}")

    print(f"[bitable] 准备写入 {len(records_payload)} 条记录到表 {table_id}")
    created = bitable_batch_create(token, app_token, table_id, records_payload)
    print(f"[bitable] 成功写入 {len(created)} 条记录")
    return {"created": len(created), "records": created}


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

    # 多维表格相关（DELIVERY_MODE=bitable 时必填）
    bitable_app_token = env("BITABLE_APP_TOKEN")
    bitable_table_id = env("BITABLE_TABLE_ID")

    delivery_mode = (env("DELIVERY_MODE") or "feishu").lower()

    if not username or not password:
        print("[错误] 缺少 DH_USERNAME / DH_PASSWORD")
        return 2

    # 1) 登录
    session = create_session()
    if not login(session, username, password):
        return 1

    # 2) 导出数据（优先 ListAPI，失败或异常才回退 Excel 导出）
    filters = build_filters_from_env()
    rows: list[dict] = []
    csv_bytes: bytes = b""
    row_count: int = 0
    col_count: int = 0
    data_source = ""
    MIN_ROWS = 20  # 阈值：临调库存不会少于 20 条；低于此值视为可能被 IP 风控过滤 → 回退 Excel

    print("\n========== 导出阶段（先 ListAPI，失败回退 Excel） ==========")
    try:
        rows, csv_bytes, row_count, col_count = export_lindiao_listapi(
            session, filters=filters)
        data_count = len(rows)
        print(f"[ListAPI] 返回 {data_count} 条")
        if data_count < MIN_ROWS:
            print(f"[ListAPI] 数据异常少 (<{MIN_ROWS})，判断可能被 IP 过滤，自动回退 Excel 导出接口")
            raise RuntimeError(f"ListAPI only returned {data_count} rows (< {MIN_ROWS})")
        data_source = "ListAPI"
    except Exception as exc_list:
        print(f"[ListAPI] 失败: {exc_list}")
        print("         → 回退 Excel 导出接口 (kucunld)")
        try:
            html_bytes, info = export_lindiao(session, filters=filters)
            csv_bytes, row_count, col_count, rows = html_table_to_csv_bytes(html_bytes)
            data_source = "ExcelExport"
        except Exception as exc_excel:
            print(f"[错误] ListAPI 和 Excel 导出都失败")
            print(f"  ListAPI error: {exc_list}")
            print(f"  Excel err:   {exc_excel}")
            traceback.print_exc()
            return 3

    print(f"[解析] 数据源: {data_source}; CSV: {row_count} 行 (含表头), {col_count} 列, "
          f"{len(csv_bytes)/1024:.1f} KB")

    now = datetime.datetime.now()
    filename = f"lindiao_{now.strftime('%Y%m%d_%H%M%S')}.csv"

    filter_desc = ""
    if filters:
        filter_desc = " | 筛选: " + ", ".join(f"{k}={v}" for k, v in filters.items())
    else:
        filter_desc = " | 筛选: 无(全部)"

    summary_lines = [
        f"导出时间: {now.strftime('%Y-%m-%d %H:%M:%S')}",
        f"数据来源: {LINDIAO_PAGE}{filter_desc}（接口: {data_source}）",
        f"数据量: {row_count - 1} 条 (表头 {col_count} 列)",
        f"文件名: {filename}",
        f"文件大小: {len(csv_bytes)/1024:.1f} KB",
    ]
    print("\n[汇总]")
    for line in summary_lines:
        print(f"  {line}")

    # 4) 交付
    # 支持的模式：
    #   local  : 仅本地保存 CSV
    #   feishu : 上传 CSV 到飞书云盘
    #   bitable: 追加写入飞书多维表格
    #   both   : 先上传云盘，再写多维表格
    actions: list[str] = []
    if delivery_mode == "local":
        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        with open(out_path, "wb") as fp:
            fp.write(csv_bytes)
        print(f"\n[本地保存] 已保存到 {out_path}")
        return 0
    elif delivery_mode == "feishu":
        actions = ["feishu"]
    elif delivery_mode == "bitable":
        actions = ["bitable"]
    elif delivery_mode == "both":
        actions = ["feishu", "bitable"]
    else:
        print(f"[错误] 未识别的 DELIVERY_MODE: {delivery_mode}")
        return 5

    # 飞书 token（feishu / bitable / both 都需要）
    try:
        fs_token = feishu_tenant_access_token(fs_app_id, fs_app_secret)
    except Exception as exc:
        print(f"[错误] 获取飞书 tenant_access_token 失败: {exc}")
        return 4

    rc = 0
    file_token = ""
    bitable_created = 0
    for action in actions:
        print(f"\n========== 交付动作: {action} ==========")
        if action == "feishu":
            if not (fs_app_id and fs_app_secret and fs_folder_token):
                print("[错误] feishu 缺少 FEISHU_APP_ID / FEISHU_APP_SECRET / FEISHU_FOLDER_TOKEN")
                rc = 4
                continue
            try:
                file_info = feishu_upload_to_folder(fs_token, fs_folder_token,
                                                    csv_bytes, filename)
                file_token = file_info.get("file_token") or file_info.get("token") or ""
                print(f"[feishu] 已上传到云盘 file_token={file_token}")
            except Exception as exc:
                print(f"[错误] 飞书云盘上传失败: {exc}")
                traceback.print_exc()
                rc = 4
                continue

        elif action == "bitable":
            if not (bitable_app_token and bitable_table_id):
                print("[错误] bitable 缺少 BITABLE_APP_TOKEN / BITABLE_TABLE_ID")
                rc = 4
                continue
            try:
                result = bitable_append_records(
                    fs_token, bitable_app_token, bitable_table_id, rows)
                bitable_created = result.get("created", 0)
                print(f"[bitable] 写入 {bitable_created} 条")
            except Exception as exc:
                print(f"[错误] bitable 写入失败: {exc}")
                traceback.print_exc()
                rc = 4
                continue

    # 汇总通知
    if rc == 0:
        ok_parts = []
        if "feishu" in actions:
            ok_parts.append("已上传云盘")
        if "bitable" in actions:
            ok_parts.append(f"已写入多维表格 {bitable_created} 条")
        notify_lines = [f"✅ 临调库存导出完成（{' + '.join(ok_parts)}）"] + summary_lines
        if "feishu" in actions and file_token:
            notify_lines.append(f"file_token: {file_token}")
        if "bitable" in actions:
            notify_lines.append(f"app_token: {bitable_app_token}")
            notify_lines.append(f"table_id: {bitable_table_id}")
        notify_text = "\n".join(notify_lines)
        if fs_webhook_url:
            try:
                feishu_send_bot_text(fs_webhook_url, fs_webhook_secret, notify_text)
            except Exception as exc:
                print(f"[警告] 通知发送失败: {exc}")
        else:
            print("\n[通知内容]")
            print(notify_text)

    return rc


if __name__ == "__main__":
    exit(main())
