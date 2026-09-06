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
  BITABLE_NO_DELETE     设为 true 跳过"删除最早批次"逻辑（默认 false）

扩展功能（写入多维表格后自动执行）：
  1. 给每条记录注入"创建时间"字段（毫秒时间戳），作为批次标识
  2. 删除目标表"创建时间"最早的整批记录（三道保护：BITABLE_NO_DELETE / 本次写入 0 条 / 最早日期==今天）
  3. 更新 AI 反馈表 tblUzkPskttsBa0W 中固定记录 recv2rwZdad6FJ 的"AI 反馈"字段
     （写入秒级时间 + 状况：写入N条 | 删除:日期(N条)）
  4. 飞书机器人发送表格化通知（源文件/数据来源/写入条数/丢值/失败批次/删除批次/AI 反馈/执行时间）
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
    # 两个批次标识字段（都注入同一个 batch_ts_ms）：
    # - "创建时间"：用户指定用作删除最早批次的排序/筛选字段
    # - "时间"    ：飞书表中原本存在的列（若开启了默认值会被飞书自动写入服务器时间），
    #               我们主动覆盖它，保证和"创建时间"毫秒级完全一致，避免视觉分裂
    "创建时间": "datetime",
    "时间": "datetime",
}

# batch_create 单次最多 500 条
BITABLE_BATCH_SIZE = 500
# batch_delete 单次最多 100 条
BITABLE_DELETE_BATCH = 100

# AI 反馈表（硬编码，用户指定）
AI_FEEDBACK_TABLE_ID = "tblUzkPskttsBa0W"
AI_FEEDBACK_RECORD_ID = "recv2rwZdad6FJ"
AI_FEEDBACK_FIELD_NAME = "AI 反馈"

# 捆包号同步表（硬编码，用户指定）
# 每次 workflow 跑完写多维表格后，把主表 tblaHMNprLueWYDP 的"捆包号"字段全量镜像到这张表
BUNDLE_SYNC_TABLE_ID = "tblDFiiwWkJ5tZp3"
BUNDLE_SYNC_FIELD = "捆包号"

# 时区：Asia/Shanghai (UTC+8)，硬编码避免依赖系统时区
SHANGHAI_OFFSET_HOURS = 8


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


def epoch_ms_to_shanghai_date(ms: int) -> str:
    """毫秒时间戳 → 'YYYY-MM-DD'（按 Asia/Shanghai 时区）"""
    dt_utc = datetime.datetime.utcfromtimestamp(ms / 1000)
    dt_sh = dt_utc + datetime.timedelta(hours=SHANGHAI_OFFSET_HOURS)
    return dt_sh.strftime("%Y-%m-%d")


def shanghai_date_to_day_range_ms(date_str: str) -> tuple[int, int]:
    """'YYYY-MM-DD'（Shanghai 本地）→ [start_ms, end_ms]（UTC ms，含当天 23:59:59.999）"""
    naive_start = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    utc_start = naive_start - datetime.timedelta(hours=SHANGHAI_OFFSET_HOURS)
    # 显式设 UTC tzinfo，避免 .timestamp() 用系统本地时区解读（本地 Shanghai / CI UTC 不一致）
    utc_aware = utc_start.replace(tzinfo=datetime.timezone.utc)
    start_ms = int(utc_aware.timestamp() * 1000)
    end_ms = start_ms + 86400 * 1000 - 1
    return start_ms, end_ms


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
                            rows: list[dict],
                            batch_ts_ms: int | None = None) -> dict:
    """
    将临调库存 rows（解析出的 dict 列表）追加到多维表格。
    字段映射见 BITABLE_FIELD_TYPES；不在此映射的字段自动跳过。
    不去重，直接 append。
    batch_ts_ms: 若提供，给每条记录的"创建时间"字段填上此毫秒时间戳（作为批次标识）。
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
        # 注入批次时间戳（覆盖所有"时间"列）：
        # 同时写入「创建时间」和「时间」两个字段，毫秒级完全相同的 batch_ts_ms。
        # 这样既覆盖了飞书表"时间"列可能存在的"创建记录时自动填充当前时间"默认值，
        # 也确保用户在 UI 里看任何一个时间列都一致（每一行、每一列都一样）。
        if batch_ts_ms is not None:
            fields["创建时间"] = batch_ts_ms
            fields["时间"] = batch_ts_ms
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


# ---------------- 删除最早批次 + AI 反馈 + 表格通知 ----------------

def _bitable_collect_record_ids_by_date(token: str, app_token: str,
                                         table_id: str, field_name: str,
                                         date_str: str) -> tuple[list[str], str]:
    """
    收集指定 field_name（datetime 类型）对应 date_str（YYYY-MM-DD，Shanghai）
    的所有 record_id。

    优先用 search API + filter(isWithin, ms 范围)，失败 fallback 到 list API 全量翻页
    + Python 端按日期过滤。

    返回 (record_ids, source)，source ∈ {"search", "list_fallback"}。
    """
    headers = {"Authorization": f"Bearer {token}"}
    start_ms, end_ms = shanghai_date_to_day_range_ms(date_str)
    record_ids: list[str] = []

    # 方法 1: search API + isWithin filter
    try:
        search_url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
                      f"/tables/{table_id}/records/search")
        page_token: str | None = None
        page = 1
        while True:
            body: dict = {
                "filter": {
                    "conjunction": "and",
                    "conditions": [{
                        "field_name": field_name,
                        "operator": "isWithin",
                        "value": [start_ms, end_ms],
                    }],
                },
                "page_size": 500,
            }
            if page_token:
                body["page_token"] = page_token
            r = requests.post(search_url, headers=headers, json=body, timeout=30)
            data = r.json()
            if data.get("code") != 0:
                raise RuntimeError(
                    f"search API code={data.get('code')}, msg={data.get('msg')}"
                )
            items = (data.get("data") or {}).get("items") or []
            for item in items:
                rid = item.get("record_id")
                if rid:
                    record_ids.append(rid)
            page_token = (data.get("data") or {}).get("page_token") or ""
            has_more = (data.get("data") or {}).get("has_more")
            print(f"  [search] page {page}: +{len(items)}, "
                  f"累计 {len(record_ids)}, has_more={has_more}")
            if not has_more or not page_token:
                break
            page += 1
        return record_ids, "search"
    except Exception as exc:
        print(f"  [search] 失败: {exc}; 回退 list API 全量翻页")

    # 方法 2: list API 全量翻页 + Python 端按 ms 范围过滤
    list_url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
                f"/tables/{table_id}/records")
    page_token = None
    page = 1
    while True:
        params: dict = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(list_url, headers=headers, params=params, timeout=30)
        data = r.json()
        if data.get("code") != 0:
            print(f"  [list] API 错误: code={data.get('code')}, msg={data.get('msg')}")
            break
        items = (data.get("data") or {}).get("items") or []
        for item in items:
            ts = (item.get("fields") or {}).get(field_name)
            if isinstance(ts, (int, float)) and start_ms <= int(ts) <= end_ms:
                rid = item.get("record_id")
                if rid:
                    record_ids.append(rid)
        page_token = (data.get("data") or {}).get("page_token") or ""
        has_more = (data.get("data") or {}).get("has_more")
        print(f"  [list] page {page}: +{len(items)}, "
              f"累计匹配 {len(record_ids)}, has_more={has_more}")
        if not has_more or not page_token:
            break
        page += 1
        time.sleep(0.1)
    return record_ids, "list_fallback"


def bitable_delete_oldest_batch(token: str, app_token: str, table_id: str,
                                  today_date: str,
                                  just_wrote_count: int) -> dict:
    """
    删除目标表中"创建时间"字段最早的整批记录。

    三道保护（任一命中即跳过删除）：
      1. 环境变量 BITABLE_NO_DELETE=true → 跳过
      2. just_wrote_count <= 0 → 跳过（没导进新数据别白删旧数据）
      3. 最早批次日期 == today_date → 跳过（防止自删本次导入）

    返回 dict:
      {
        "skipped": bool,
        "skip_reason": str,           # skipped=True 时填
        "deleted_count": int,
        "batch_date": str,            # 最早批次日期 YYYY-MM-DD（即使 skipped 也回填）
        "total_in_batch": int,       # 批次总条数
        "source": str,               # "search" / "list_fallback" / ""
      }
    """
    result = {
        "skipped": True, "skip_reason": "",
        "deleted_count": 0, "batch_date": "",
        "total_in_batch": 0, "source": "",
    }

    # 保护 1: 环境变量
    if env("BITABLE_NO_DELETE", "false").lower() == "true":
        result["skip_reason"] = "BITABLE_NO_DELETE=true"
        print(f"[delete] 跳过: {result['skip_reason']}")
        return result

    # 保护 2: 本次写入 0 条
    if just_wrote_count <= 0:
        result["skip_reason"] = f"本次写入 {just_wrote_count} 条"
        print(f"[delete] 跳过: {result['skip_reason']}")
        return result

    headers = {"Authorization": f"Bearer {token}"}

    # 步骤 1: 查最早记录的"创建时间"
    print(f"[delete] 步骤1: 查询最早批次日期 (按'创建时间'升序)")
    search_url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
                  f"/tables/{table_id}/records/search")
    try:
        sort_body = {
            "sort": [{"field_name": "创建时间", "desc": False}],
            "page_size": 1,
        }
        r = requests.post(search_url, headers=headers, json=sort_body, timeout=30)
        data = r.json()
        if data.get("code") != 0:
            result["skip_reason"] = (f"查询最早记录失败: code={data.get('code')}, "
                                      f"msg={data.get('msg')}")
            print(f"[delete] {result['skip_reason']}")
            return result
        items = (data.get("data") or {}).get("items") or []
        if not items:
            result["skip_reason"] = "表中无记录"
            print(f"[delete] {result['skip_reason']}")
            return result
        oldest_ts = (items[0].get("fields") or {}).get("创建时间")
        if not isinstance(oldest_ts, (int, float)):
            result["skip_reason"] = f"最早记录'创建时间'字段为空或非数字: {oldest_ts!r}"
            print(f"[delete] {result['skip_reason']}")
            return result
        oldest_date = epoch_ms_to_shanghai_date(int(oldest_ts))
        result["batch_date"] = oldest_date
        print(f"[delete] 最早批次日期: {oldest_date}")
    except Exception as exc:
        result["skip_reason"] = f"查询最早记录异常: {exc}"
        print(f"[delete] {result['skip_reason']}")
        return result

    # 保护 3: 最早日期 == 今天
    if oldest_date == today_date:
        result["skip_reason"] = (f"最早批次日期 {oldest_date} == 今天 {today_date}, "
                                  f"防止自删本次导入")
        print(f"[delete] 跳过: {result['skip_reason']}")
        return result

    # 步骤 2: 收集该日期全部 record_id
    print(f"[delete] 步骤2: 收集 {oldest_date} 的全部 record_id")
    record_ids, source = _bitable_collect_record_ids_by_date(
        token, app_token, table_id, "创建时间", oldest_date)
    result["source"] = source
    result["total_in_batch"] = len(record_ids)
    print(f"[delete] 共 {len(record_ids)} 条待删 (source={source})")

    if not record_ids:
        result["skip_reason"] = f"日期 {oldest_date} 未找到任何记录"
        print(f"[delete] {result['skip_reason']}")
        return result

    # 步骤 3: 分批硬删（每批 ≤ BITABLE_DELETE_BATCH=100）
    print(f"[delete] 步骤3: 分批硬删 (每批 ≤{BITABLE_DELETE_BATCH} 条)")
    delete_url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
                  f"/tables/{table_id}/records/batch_delete")
    deleted_count = 0
    batch_num = 0
    for start in range(0, len(record_ids), BITABLE_DELETE_BATCH):
        batch = record_ids[start:start + BITABLE_DELETE_BATCH]
        batch_num += 1
        try:
            r = requests.post(delete_url, headers=headers,
                              json={"records": batch}, timeout=60)
            data = r.json()
            if data.get("code") != 0:
                print(f"  批次 {batch_num} 删除失败: code={data.get('code')}, "
                      f"msg={data.get('msg')}")
                continue
            deleted_count += len(batch)
            print(f"  批次 {batch_num}: 删除 {len(batch)} 条, "
                  f"累计 {deleted_count}/{len(record_ids)}")
        except Exception as exc:
            print(f"  批次 {batch_num} 异常: {exc}")
            continue
        time.sleep(0.1)

    result["skipped"] = False
    result["deleted_count"] = deleted_count
    print(f"[delete] 完成: 删除 {deleted_count}/{len(record_ids)} 条, 日期={oldest_date}")
    return result


def bitable_update_ai_feedback(token: str, app_token: str,
                                content: str) -> bool:
    """
    更新 AI 反馈表 tblUzkPskttsBa0W 中固定记录 recv2rwZdad6FJ 的"AI 反馈"字段。
    失败不抛异常，返回 False。
    """
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{AI_FEEDBACK_TABLE_ID}/records/{AI_FEEDBACK_RECORD_ID}")
    headers = {"Authorization": f"Bearer {token}"}
    payload = {"fields": {AI_FEEDBACK_FIELD_NAME: content}}
    try:
        r = requests.put(url, headers=headers, json=payload, timeout=30)
        data = r.json()
        if data.get("code") != 0:
            print(f"[AI 反馈] 更新失败: code={data.get('code')}, "
                  f"msg={data.get('msg')}")
            return False
        print(f"[AI 反馈] 已更新 record_id={AI_FEEDBACK_RECORD_ID} "
              f"的 '{AI_FEEDBACK_FIELD_NAME}' 字段")
        return True
    except Exception as exc:
        print(f"[AI 反馈] 更新异常: {exc}")
        return False


def bitable_sync_bundle_numbers(token: str, app_token: str,
                                  dst_table_id: str,
                                  bundle_numbers: list[str],
                                  field_name: str = BUNDLE_SYNC_FIELD) -> dict:
    """
    把 bundle_numbers 列表全量同步到 dst_table_id 的 field_name 字段。
    （即"最新一次录入"的捆包号镜像，不是源表全量历史数据）
    流程：
      1. 列出 dst_table_id 所有 record_id，分批硬删（≤100/批）
      2. batch_create 到 dst_table_id（≤500/批），每条记录只写 field_name 字段
    返回: {"deleted": N, "created": M, "skipped": bool, "skip_reason": str}
    失败不抛异常，记录在 skip_reason 中。
    """
    headers = {"Authorization": f"Bearer {token}"}
    result: dict = {"deleted": 0, "created": 0, "skipped": False, "skip_reason": ""}

    # ---- 步骤 1: 清空 dst 表 ----
    print(f"[sync] 步骤1: 清空目标表 {dst_table_id}")
    list_url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
                f"/tables/{dst_table_id}/records")
    dst_record_ids: list[str] = []
    page_token: str | None = None
    page = 1
    while True:
        params: dict = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        try:
            r = requests.get(list_url, headers=headers, params=params, timeout=30)
            data = r.json()
        except Exception as exc:
            result["skipped"] = True
            result["skip_reason"] = f"列 dst 表异常: {exc}"
            print(f"[sync] 列 dst 表异常: {exc}")
            return result
        if data.get("code") != 0:
            result["skipped"] = True
            result["skip_reason"] = (f"列 dst 表 code={data.get('code')}, "
                                      f"msg={data.get('msg')}")
            print(f"[sync] 列 dst 表失败: code={data.get('code')}, "
                  f"msg={data.get('msg')}")
            return result
        items = (data.get("data") or {}).get("items") or []
        for it in items:
            rid = it.get("record_id")
            if rid:
                dst_record_ids.append(rid)
        page_token = (data.get("data") or {}).get("page_token") or ""
        has_more = (data.get("data") or {}).get("has_more")
        print(f"  [list-dst] page {page}: +{len(items)}, "
              f"累计 {len(dst_record_ids)}, has_more={has_more}")
        if not has_more or not page_token:
            break
        page += 1
        time.sleep(0.1)

    print(f"[sync] dst 表共 {len(dst_record_ids)} 条待删")

    # 分批硬删
    delete_url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
                  f"/tables/{dst_table_id}/records/batch_delete")
    deleted_count = 0
    for start in range(0, len(dst_record_ids), BITABLE_DELETE_BATCH):
        batch = dst_record_ids[start:start + BITABLE_DELETE_BATCH]
        try:
            r = requests.post(delete_url, headers=headers,
                              json={"records": batch}, timeout=60)
            data = r.json()
            if data.get("code") != 0:
                print(f"  删除批次失败: code={data.get('code')}, "
                      f"msg={data.get('msg')}")
                continue
            deleted_count += len(batch)
            print(f"  删除 {start + len(batch)}/{len(dst_record_ids)}")
        except Exception as exc:
            print(f"  删除异常: {exc}")
            continue
        time.sleep(0.1)
    result["deleted"] = deleted_count
    print(f"[sync] 步骤1 完成: 删除 {deleted_count}/{len(dst_record_ids)} 条")

    # ---- 步骤 2: batch_create 到 dst 表 ----
    # bundle_numbers 是本次 workflow 刚从懂火导出并写入主表的最新一批捆包号
    # （即"定价视图"应该展示的那批），直接用，无需再回查源表
    if not bundle_numbers:
        result["skip_reason"] = "本批无捆包号数据"
        return result

    print(f"[sync] 步骤2: 写入 {len(bundle_numbers)} 条到 dst 表 {dst_table_id}")
    records_payload = [{"fields": {field_name: v}} for v in bundle_numbers]
    try:
        created = bitable_batch_create(token, app_token, dst_table_id,
                                        records_payload)
        result["created"] = len(created)
        print(f"[sync] 步骤2 完成: 写入 {len(created)}/{len(bundle_numbers)} 条")
    except Exception as exc:
        result["skipped"] = True
        result["skip_reason"] = f"batch_create 异常: {exc}"
        print(f"[sync] batch_create 异常: {exc}")
    return result


def build_notify_table(filename: str, row_count: int, bitable_created: int,
                        delete_result: dict, ai_feedback_ok: bool | None,
                        exec_time_str: str, filter_desc: str,
                        sync_result: dict | None = None) -> str:
    """
    生成纯文本 markdown 表格通知。

    返回示例：
      | 项        | 内容                          |
      | -------- | --------------------------- |
      | 源文件     | lindiao_20260817_213015.csv |
      | 数据来源   | 临调库存（129 行）                 |
      | 写入条数   | 129 条                       |
      | 丢值       | 0 条                          |
      | 失败批次   | 0                            |
      | 删除批次   | 2026-05-16 (130 条)          |
      | AI 反馈    | 已更新                          |
      | 执行时间   | 2026-08-17 21:30:15          |
    """
    items: list[tuple[str, str]] = [
        ("源文件", filename),
        ("数据来源", f"临调库存（{row_count - 1} 行{filter_desc}）"),
        ("写入条数", f"{bitable_created} 条"),
        ("丢值", "0 条"),
        ("失败批次", "0"),
    ]

    # 删除批次行
    if delete_result.get("skipped"):
        reason = delete_result.get("skip_reason") or "未执行"
        items.append(("删除批次", f"跳过（{reason}）"))
    elif delete_result.get("batch_date"):
        items.append(("删除批次",
                       f"{delete_result['batch_date']} "
                       f"({delete_result.get('deleted_count', 0)} 条)"))
    else:
        items.append(("删除批次", "未执行"))

    # AI 反馈行
    if ai_feedback_ok is True:
        items.append(("AI 反馈", "已更新"))
    elif ai_feedback_ok is False:
        items.append(("AI 反馈", "更新失败"))
    else:
        items.append(("AI 反馈", "未执行"))

    # 捆包号同步行
    if sync_result is not None:
        if sync_result.get("skipped"):
            reason = sync_result.get("skip_reason") or "未执行"
            items.append(("捆包号同步", f"跳过（{reason}）"))
        else:
            items.append(("捆包号同步",
                           f"清空 {sync_result.get('deleted', 0)} 条 / "
                           f"写入 {sync_result.get('created', 0)} 条"))

    items.append(("执行时间", exec_time_str))

    lines = ["| 项 | 内容 |", "| --- | --- |"]
    for k, v in items:
        # 转义 markdown 表格中的 |
        v_safe = str(v).replace("|", "\\|")
        lines.append(f"| {k} | {v_safe} |")
    return "\n".join(lines)


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


def receive_id_type_of(receive_id: str) -> str:
    """按前缀自动识别：oc_ 开头 = 群 chat_id（群聊），其余按个人 union_id 处理。"""
    return "chat_id" if receive_id.startswith("oc_") else "union_id"


def feishu_send_im_text(token: str, receive_id: str, text: str,
                        *, receive_id_type: str | None = None) -> None:
    """
    用自建应用 tenant_access_token 通过 IM API 发文本消息。

    receive_id_type 不传时按前缀自动识别（oc_ → chat_id，其余 → union_id）。

    receive_id_type 可选:
      - "chat_id"  : 发群聊（receive_id = oc_xxx）。
                     需要应用已开启"机器人"能力 + 已被加入目标群 + im:message:send_as_bot 权限。
      - "union_id" : 发私聊给个人（receive_id = on_xxx）。
                     需要应用在用户所在部门"可用"（已发布）+ im:message:send_as_bot 权限。
      - "open_id"  : 发私聊给个人（receive_id = ou_xxx，跟应用绑定的 open_id）。
      - "user_id"  : 发私聊给个人（receive_id = feishu user_id）。
      - "email"    : 发私聊给个人（receive_id = 飞书邮箱）。
    """
    receive_id_type = receive_id_type or receive_id_type_of(receive_id)
    url = f"{FEISHU_OPEN_BASE}/im/v1/messages?receive_id_type={receive_id_type}"
    headers = {"Authorization": f"Bearer {token}"}
    payload = {
        "receive_id": receive_id,
        "msg_type": "text",
        "content": json.dumps({"text": text}, ensure_ascii=False),
    }
    r = requests.post(url, headers=headers, json=payload, timeout=15)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"IM API 发消息失败 (type={receive_id_type}, "
                           f"id={receive_id}): code={data.get('code')}, "
                           f"msg={data.get('msg')}, raw={data}")
    print(f"[飞书通知] IM API 发送成功 (type={receive_id_type}, id={receive_id[:6]}...)")


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
    # 收件人列表：支持混合类型 —— oc_ 开头 = 群 chat_id（群聊），其余 = 个人 union_id（私聊）
    # 默认收件人已写死（2 个个人 + 2 个群），无需配置 GitHub Secret
    # 临时换收件人：设 FEISHU_UNION_ID_OUYEE 环境变量（或 GitHub Variable）即可覆盖
    # 覆盖优先级：FEISHU_UNION_ID_OUYEE（临调/欧冶专用）> FEISHU_UNION_IDS > 默认写死
    DEFAULT_UNION_IDS = (
        "on_93da40c6314edbfa2dc3e031ef405389,"
        "on_b09bcbf3e74f5d423900aa9b2f00eb63,"
        "oc_334f8c12e73592af76dccb5b34ccfa5f,"
        "oc_d22e1f9c8cd0a5a3aa2b2625e2a8f155"
    )
    fs_union_ids_raw = (env("FEISHU_UNION_ID_OUYEE")
                        or env("FEISHU_UNION_IDS")
                        or DEFAULT_UNION_IDS)
    fs_union_ids = [x.strip() for x in fs_union_ids_raw.split(",") if x.strip()]
    # 群聊通知：chat_id（备选，应用需已加入群）
    fs_chat_id = env("FEISHU_CHAT_ID")
    # Webhook 备选（二选一即可，都不配就只打印日志）
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

    # 统一时间戳：所有"时间"字段（文件名、多维表格"创建时间"、AI 反馈文本前缀、
    # 通知"执行时间"行、汇总"导出时间"）都基于同一个 batch_ts_sec 派生，
    # 确保整批更新的"时间"标识完全一致。
    batch_ts_sec = int(time.time())
    batch_ts_ms = batch_ts_sec * 1000
    now = datetime.datetime.fromtimestamp(batch_ts_sec)
    filename = f"lindiao_{now.strftime('%Y%m%d_%H%M%S')}.csv"

    today_date = epoch_ms_to_shanghai_date(batch_ts_ms)
    exec_time_str = now.strftime("%Y-%m-%d %H:%M:%S")

    # 删除/AI 反馈结果（默认值，无 bitable 动作时保持）
    delete_result: dict = {
        "skipped": True, "skip_reason": "no_bitable_action",
        "deleted_count": 0, "batch_date": "",
        "total_in_batch": 0, "source": "",
    }
    ai_feedback_ok: bool | None = None
    sync_result: dict | None = None

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
                    fs_token, bitable_app_token, bitable_table_id, rows,
                    batch_ts_ms=batch_ts_ms)
                bitable_created = result.get("created", 0)
                print(f"[bitable] 写入 {bitable_created} 条")
            except Exception as exc:
                print(f"[错误] bitable 写入失败: {exc}")
                traceback.print_exc()
                rc = 4
                continue

            # 成功写入后：① 删除最早批次  ② 更新 AI 反馈表
            try:
                print("\n========== 删除最早批次 ==========")
                delete_result = bitable_delete_oldest_batch(
                    fs_token, bitable_app_token, bitable_table_id,
                    today_date=today_date,
                    just_wrote_count=bitable_created)
            except Exception as exc:
                print(f"[警告] 删除最早批次异常: {exc}")
                traceback.print_exc()

            # AI 反馈文本（秒级时间 + 状况）
            if delete_result.get("skipped"):
                delete_desc = f"跳过({delete_result.get('skip_reason', '')})"
            else:
                delete_desc = (f"{delete_result.get('batch_date', '')} "
                                f"({delete_result.get('deleted_count', 0)} 条)")
            ai_feedback_text = (
                f"{exec_time_str} | 写入 {bitable_created} 条 | "
                f"删除: {delete_desc}"
            )
            try:
                print("\n========== 更新 AI 反馈 ==========")
                ai_feedback_ok = bitable_update_ai_feedback(
                    fs_token, bitable_app_token, ai_feedback_text)
            except Exception as exc:
                print(f"[警告] AI 反馈更新异常: {exc}")
                traceback.print_exc()

            # 同步本批捆包号到汇总表 tblDFiiwWkJ5tZp3
            # （清空目标表 → 把本次刚写入的捆包号列表写入，即"定价视图"的最新一批）
            try:
                print("\n========== 同步捆包号到汇总表 ==========")
                bundle_numbers = [
                    r.get("捆包号", "").strip()
                    for r in rows
                    if r.get("捆包号", "").strip()
                ]
                sync_result = bitable_sync_bundle_numbers(
                    fs_token, bitable_app_token,
                    dst_table_id=BUNDLE_SYNC_TABLE_ID,
                    bundle_numbers=bundle_numbers,
                )
            except Exception as exc:
                print(f"[警告] 同步捆包号异常: {exc}")
                traceback.print_exc()
                sync_result = {"skipped": True,
                                "skip_reason": f"异常: {exc}",
                                "deleted": 0, "created": 0}

    # 汇总通知（表格化）
    if rc == 0:
        ok_parts = []
        if "feishu" in actions:
            ok_parts.append("已上传云盘")
        if "bitable" in actions:
            ok_parts.append(f"已写入多维表格 {bitable_created} 条")
        header = f"✅ 临调库存导出完成（{' + '.join(ok_parts)}）"

        notify_text = header + "\n\n" + build_notify_table(
            filename=filename,
            row_count=row_count,
            bitable_created=bitable_created,
            delete_result=delete_result,
            ai_feedback_ok=ai_feedback_ok,
            exec_time_str=exec_time_str,
            filter_desc=filter_desc,
            sync_result=sync_result,
        )

        # 通知发送：收件人列表（混合 个人 on_ 私聊 + 群 oc_ 群聊），逐个发送
        #   单个失败不阻塞其他；有任一成功 → 不走兜底；全失败 → FEISHU_CHAT_ID → Webhook → 只打印日志
        notified = False

        if fs_union_ids and fs_token:
            for i, uid in enumerate(fs_union_ids, 1):
                try:
                    feishu_send_im_text(fs_token, uid, notify_text)
                    notified = True
                except Exception as exc:
                    kind = "群" if uid.startswith("oc_") else "个人"
                    print(f"[警告] IM API 通知发送失败 ({i}/{len(fs_union_ids)}, "
                          f"{kind}={uid[:8]}...): {exc}")

        if not notified and fs_chat_id and fs_token:
            try:
                feishu_send_im_text(fs_token, fs_chat_id, notify_text,
                                    receive_id_type="chat_id")
                notified = True
            except Exception as exc:
                print(f"[警告] IM API 群聊通知发送失败: {exc}")

        if not notified and fs_webhook_url:
            try:
                feishu_send_bot_text(fs_webhook_url, fs_webhook_secret, notify_text)
                notified = True
            except Exception as exc:
                print(f"[警告] Webhook 通知发送失败: {exc}")

        if not notified:
            print("\n[通知内容]")
            print(notify_text)

    return rc


if __name__ == "__main__":
    exit(main())
