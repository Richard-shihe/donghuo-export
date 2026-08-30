#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IEC —— 入库管理：出厂码单 → 码单捆包下载 → 飞书云盘
====================================================

任务说明（来自用户）：
  "封装模板登录 IEC，获取 '出厂码单' 页面里 '码单捆包下载'，
   放入飞书多维文件夹：https://s2v31ke6sl.feishu.cn/drive/folder/DfQdfSxl2ld25wdx6Rxcub9hnDf
   文件命名规则：码单_YYMMDD_HHMMSS"

页面定位（实测）：
  菜单路径  : 货物管理 → 出厂码单
  页面 URL  : /iecs/freight/weightMemo/weightMemo/initLoads
  页面 title: 码单查询

下载接口（来自页面 JS，已实测 219 行 × 21 列 xlsx）：
  POST /iecs/common/iec/exportExcel
    apiBean   = "com.baosight.iecs.freight.weightMemo.api.IMemoStackDetailService"
    methodName= "queryALLMemoStackDownload"        （页面上点击"码单捆包下载"全量下载时调用）
    settleUserNum = "062122"
    筛选字段（日期格式 YYYYMMDD）:
      putoutStackingBeginTime / putoutStackingEndTime    出厂日期起/止
      putoutDateBeginTime     / putoutDateEndTime        （同上值，下载按钮里额外赋值）
      stackingRecNum          / putoutStackingRecNum     码单号（未填即空）
      contractNum / factoryOrderNum / shopsign / prodCode / deliveryMonthStart / End …
      instockFlag = "ALL"      （全量下载时写死）
      offset=0  limit=1000
  下载 URL:
    GET  /iecs/common/download?fileName=<URLencoded>&delete=true

输出：
  文件名 : 码单_YYMMDD_HHMMSS.xlsx  (2 位年份，精确到秒，防同名覆盖)
  上传   : 飞书云盘 folder_token = "DfQdfSxl2ld25wdx6Rxcub9hnDf"
           （对应 URL /folder/DfQdfSxl2ld25wdx6Rxcub9hnDf）

默认筛选（与网页『加载更多到 3 天左右』操作习惯对齐）：
  出厂日期: 近 5 天 ~ 今天（= 今天前 4 天到今天，共 5 天）
  可用 --days N 覆盖；也可直接用 --date-start YYYYMMDD --date-end YYYYMMDD 指定具体日期。

环境变量（必须配置在 GitHub Secrets 或本地 .env / 系统 env）：
  IBAO_USERNAME / IBAO_PASSWORD        IEC 账号密码
  FEISHU_APP_ID   / FEISHU_APP_SECRET  飞书自建应用（需 drive:drive 权限 + 目标文件夹可编辑协作者）
  FEISHU_FOLDER_TOKEN                  默认 DfQdfSxl2ld25wdx6Rxcub9hnDf，可覆盖
  DELIVERY_MODE                        目前仅 'feishu'，即上传云盘（默认）

命令行（建议从仓库根目录执行）：
  python RUKU/export_iec_madan.py                     # 出厂日期近 5 天（含今日），自动上传
  python RUKU/export_iec_madan.py --days 30          # 近 30 天
  python RUKU/export_iec_madan.py --date-start 20260801 --date-end 20260821
  python RUKU/export_iec_madan.py --dry-run          # 只下载不上传
  python RUKU/export_iec_madan.py --no-upload        # 同上（等价 DRY_RUN=1）

注：脚本现在位于 RUKU/ 子目录，但以下文件仍在仓库根目录（统一路径管理）：
  - ibaosteel_client.py（IEC 登录封装）
  - .env（环境变量）
  - iecc.json（IEC 登录 token 缓存，会写入 CWD 方便本地复用）
  - 码单_*.xlsx 输出文件（写入 CWD）
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import hmac
import json
import os
import re
import sys
import time
import urllib.parse
from pathlib import Path
from typing import Optional, Tuple

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ============================================================
# 路径定位：脚本在 RUKU/ 子目录，但 ibaosteel_client / .env 等在仓库根
#   无论从哪里执行，都确保能 import 到 ibaosteel_client，也能找到根目录 .env
# ============================================================
_HERE = Path(__file__).resolve().parent             # RUKU/
_REPO_ROOT = _HERE.parent                            # 仓库根
# 把仓库根 + 脚本所在目录都加入 sys.path（兼顾根目录/子目录两种运行方式）
for _p in (str(_REPO_ROOT), str(_HERE)):
    if _p and _p not in sys.path:
        sys.path.insert(0, _p)
# dotenv：优先从仓库根目录找 .env（找不到再搜当前工作目录）
try:
    from dotenv import load_dotenv as _load_dotenv
    _env_paths = [_REPO_ROOT / ".env", Path.cwd() / ".env"]
    for _ep in _env_paths:
        if _ep.is_file():
            _load_dotenv(_ep, override=False)
            break
except ImportError:
    pass

try:
    import pandas as pd
    _HAS_PANDAS = True
except ImportError:
    _HAS_PANDAS = False

from ibaosteel_client import IEC


# ============================================================
# 常量（从探测脚本 + 实测得来，勿随意改动）
# ============================================================
BASE = "https://www.ibaosteel.com"
IECS_INDEX = f"{BASE}/iecs/index"
WEIGHTMEMO_PAGE = f"{BASE}/iecs/freight/weightMemo/weightMemo/initLoads"
EXPORT_API = f"{BASE}/iecs/common/iec/exportExcel"
DOWNLOAD_API = f"{BASE}/iecs/common/download"
SETTLE_USER_NUM = "062122"

# 码单捆包下载（全量）  实测 2026-08-21 → 219 行 xlsx
DEFAULT_API_BEAN = "com.baosight.iecs.freight.weightMemo.api.IMemoStackDetailService"
DEFAULT_METHOD_NAME = "queryALLMemoStackDownload"

# 飞书
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}
# 用户明确指定的文件夹 URL 中提取：/folder/DfQdfSxl2ld25wdx6Rxcub9hnDf
DEFAULT_FOLDER_TOKEN = "DfQdfSxl2ld25wdx6Rxcub9hnDf"

# 本地/CI 默认文件名：码单_YYMMDD_HHMMSS.xlsx（2 位年份 + 秒级）
FILE_PREFIX = "码单"

# 通知：默认推 2 个 Union ID
DEFAULT_NOTIFY_UNION_IDS = [
    "on_b09bcbf3e74f5d423900aa9b2f00eb63",   # 洪
    "on_287d6d65c2f47f75c4379dd3f77a106a",   # 王阳
]
DEFAULT_NOTIFY_APP_ID = "cli_aaf0ce1e9ef89d27"


# ============================================================
# 工具
# ============================================================
def env(name: str, default: str = "") -> str:
    return (os.environ.get(name) or "").strip()


def create_session() -> requests.Session:
    s = requests.Session()
    retry = Retry(total=5, backoff_factor=1,
                  status_forcelist=[500, 502, 503, 504, 429],
                  allowed_methods=["POST", "GET"])
    adapter = HTTPAdapter(max_retries=retry)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    return s


def _parse_yyyymmdd(s: str) -> str:
    s = str(s).strip()
    if s == "":
        return ""
    if len(s) == 8 and s.isdigit():
        y, m, d = int(s[:4]), int(s[4:6]), int(s[6:8])
        datetime.date(y, m, d)  # 校验
        return s
    raise argparse.ArgumentTypeError(f"日期格式错误: {s!r}（应为 YYYYMMDD，如 20260821）")


def _default_date_range(days: int) -> Tuple[str, str]:
    """返回 (start, end) 的 YYYYMMDD 字符串，end=今天, start=end 往前 days-1 天。

    days=3 → [today - 2 days, today] 共 3 天（与网页上『加载更多到 3 天左右』对齐）。
    """
    today = datetime.date.today()
    start = today - datetime.timedelta(days=days - 1)
    return start.strftime("%Y%m%d"), today.strftime("%Y%m%d")


def _auto_filename() -> str:
    """码单_YYMMDD_HHMMSS.xlsx"""
    now = datetime.datetime.now()
    return f"{FILE_PREFIX}_{now:%y%m%d_%H%M%S}.xlsx"


def _detect_ext(head: bytes) -> str:
    """探测文件扩展名。xlsx = PK\\x03\\x04 zip 头；xls = OLE Compound (D0CF11E0)."""
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return ".xls"
    # 绝大多数为 zip(xlsx)；兜底给 xlsx
    return ".xlsx"


# ============================================================
# Stage 1: IEC 登录 → exportExcel → 下载 xlsx
# ============================================================
def iec_login_and_enter() -> Tuple[IEC, requests.Session, str]:
    """登录 IEC → 打开 iecs/index → 进入出厂码单页面。返回 (iec, session, token)."""
    iec = IEC()
    if not iec.login():
        raise RuntimeError("IEC 登录失败（检查 IBAO_USERNAME / IBAO_PASSWORD 或验证码）")
    iec.save("iecc.json")
    s = iec.session
    token = iec.token
    ref = f"{IECS_INDEX}?token={token}"
    r0 = s.get(ref, timeout=20, headers={"Referer": ref})
    if r0.status_code != 200:
        raise RuntimeError(f"打开 iecs/index 失败 status={r0.status_code}")
    r1 = s.get(WEIGHTMEMO_PAGE, timeout=20, headers={"Referer": ref})
    if r1.status_code != 200 or len(r1.text) < 5000:
        raise RuntimeError(f"打开出厂码单页面失败 status={r1.status_code} len={len(r1.text)}")
    print(f"[IEC] 登录 OK → 出厂码单页 size={len(r1.text)}")
    return iec, s, token


# 查询列表 API（与页面点「查询/加载更多」用的接口相同）
QUERY_LIST_API = f"{BASE}/iecs/freight/weightMemo/weightMemo/queryWeightMemoSellMoreThanOneOrder"
# 列表 API 的 apiBean
QUERY_LIST_API_BEAN = "com.baosight.iecs.freight.weightMemo.api.IWeightMemoService"
QUERY_LIST_METHOD = "queryWeightMemoSellMoreThanOneOrder"


def _query_memo_list(
    *,
    date_start: str,
    date_end: str,
    memo_code: str = "",
    contract_num: str = "",
    factory_order_num: str = "",
    shopsign: str = "",
    prod_code: str = "",
    s: Optional[requests.Session] = None,
    verbose: bool = True,
) -> list:
    """第一步：查询码单列表，获取所有记录的 ID 三元组。

    对应页面点「查询」时的 queryWeightMemo() 逻辑。
    用 putoutStackingBeginTime/EndTime（码单日期）筛选，今天的记录立即可见。
    返回 [{stackingRecNum, factoryOrderNum, orderNum}, ...]

    注意：列表 API 返回的是纯 HTML（text/html），不是 JSON。
    每行有 3 个 hidden input：
      <input name="stackingRecNum" type="hidden" value="F076040059">
      <input name="factoryOrderNum" type="hidden" value="X6E0013470">
      <input name="orderNum" type="hidden" value="QE267022440003">
    """
    if not s:
        raise ValueError("必须传已建立的 session")

    all_records = []
    page = 1
    page_size = 100  # 与页面一致

    while True:
        param = {
            "putoutStackingRecNum": memo_code,
            "contractNum": contract_num,
            "factoryOrderNum": factory_order_num,
            "prodCode": prod_code,
            "prodCode1": "",
            "shopsign": shopsign,
            "putoutStackingBeginTime": date_start,
            "putoutStackingEndTime": date_end,
            "deliveryMonthStart": "",
            "deliveryMonthEnd": "",
            "settleUserNum": SETTLE_USER_NUM,
            "pageDomain": {
                "pageNum": page,
                "pageSize": page_size,
                "total": 0,
            },
        }
        r = s.post(QUERY_LIST_API, data=json.dumps(param), timeout=60,
                   headers={
                       "Content-Type": "application/json; charset=utf-8",
                       "Accept": "text/html, application/json, */*",
                       "X-Requested-With": "XMLHttpRequest",
                       "Referer": WEIGHTMEMO_PAGE,
                   })
        r.raise_for_status()
        # ★ 注意：此 API 返回纯 HTML，不是 JSON
        html_content = r.text

        # 提取 hidden input 里的 stackingRecNum / factoryOrderNum / orderNum
        # 格式：<input name="stackingRecNum" type="hidden" value="F076040059">
        stacking_nums = re.findall(r'name="stackingRecNum"\s+type="hidden"\s+value="([^"]*)"', html_content)
        factory_nums = re.findall(r'name="factoryOrderNum"\s+type="hidden"\s+value="([^"]*)"', html_content)
        order_nums = re.findall(r'name="orderNum"\s+type="hidden"\s+value="([^"]*)"', html_content)

        count = len(stacking_nums)
        if count == 0:
            break

        for i in range(count):
            all_records.append({
                "stackingRecNum": stacking_nums[i] if i < len(stacking_nums) else "",
                "factoryOrderNum": factory_nums[i] if i < len(factory_nums) else "",
                "orderNum": order_nums[i] if i < len(order_nums) else "",
            })

        # 从 hidden input 读 total
        total_match = re.search(r'<input[^>]*id="total"[^>]*value="(\d+)"', html_content)
        total = int(total_match.group(1)) if total_match else 0
        fetched = len(all_records)

        if verbose:
            print(f"    第 {page} 页: {count} 条, 累计 {fetched}/{total}")

        if fetched >= total or count < page_size:
            break

        page += 1

    return all_records


def download_bundle(
    out_path: str,
    *,
    date_start: str,
    date_end: str,
    memo_code: str = "",           # putoutStackingRecNum 码单号（可选）
    contract_num: str = "",       # 销售合同号（可选）
    factory_order_num: str = "",  # 钢厂订单号（可选）
    shopsign: str = "",           # 牌号（可选）
    prod_code: str = "",          # 品种（可选）
    delivery_month_start: str = "",
    delivery_month_end: str = "",
    instock_flag: str = "ALL",
    s: Optional[requests.Session] = None,
    token: str = "",
    verbose: bool = True,
) -> Tuple[str, int]:
    """调用 IEC 码单捆包下载（两步法，与网页点击行为一致）。

    第一步：查列表 API 获取记录 ID 三元组（用码单日期筛选，今天立即可见）
    第二步：把 ID 列表作为 memoList 传给 exportExcel 下载

    返回 (abs_path, row_count)。row_count 在没装 pandas 时返回 -1。
    """
    if not s or not token:
        raise ValueError("必须传已建立的 session 与 token")

    # ========== 第一步：查询列表 ==========
    if verbose:
        print(f"[Step 1] 查询码单列表  码单日期 {date_start}~{date_end}")
    memo_list = _query_memo_list(
        date_start=date_start,
        date_end=date_end,
        memo_code=memo_code,
        contract_num=contract_num,
        factory_order_num=factory_order_num,
        shopsign=shopsign,
        prod_code=prod_code,
        s=s,
        verbose=verbose,
    )
    if not memo_list:
        raise RuntimeError(f"日期范围 {date_start}~{date_end} 内没有码单记录")
    if verbose:
        print(f"       找到 {len(memo_list)} 条记录")

    # ========== 第二步：用 memoList 调用 exportExcel ==========
    param = {
        "putoutStackingRecNum": memo_code,
        "contractNum": contract_num,
        "factoryOrderNum": factory_order_num,
        "prodCode": prod_code,
        "prodCode1": "",
        "shopsign": shopsign,
        "putoutStackingBeginTime": date_start,
        "putoutStackingEndTime": date_end,
        "deliveryMonthStart": delivery_month_start,
        "deliveryMonthEnd": delivery_month_end,
        "memoList": memo_list,  # ★ 关键：选中记录的 ID 列表
        "apiBean": DEFAULT_API_BEAN,
        "methodName": DEFAULT_METHOD_NAME,
        "settleUserNum": SETTLE_USER_NUM,
    }
    if verbose:
        print(f"[Step 2] POST exportExcel  "
              f"memoList={len(memo_list)} 条  "
              f"apiBean={DEFAULT_API_BEAN.split('.')[-1]}  "
              f"methodName={DEFAULT_METHOD_NAME}")
    t0 = time.time()
    # 15~30 天范围码单可达 60+ 条，IEC 同步生成 xlsx 可能超过 5 分钟，读超时放宽到 10 分钟
    r2 = s.post(EXPORT_API, data=json.dumps(param), timeout=600,
                headers={
                    "Content-Type": "application/json; charset=utf-8",
                    "Accept": "application/json, text/javascript, */*; q=0.01",
                    "X-Requested-With": "XMLHttpRequest",
                    "Referer": WEIGHTMEMO_PAGE,
                })
    r2.raise_for_status()
    try:
        result = r2.json()
    except Exception as e:
        raise RuntimeError(f"exportExcel 返回不是 JSON: {e}  body[:400]={r2.text[:400]}")
    if result.get("code") != 0:
        raise RuntimeError(f"exportExcel 失败 code={result.get('code')} "
                           f"msg={result.get('msg')}")
    server_file = str(result.get("msg") or "").strip()
    if not server_file:
        raise RuntimeError(f"exportExcel 响应中 msg 为空: {result}")
    if verbose:
        print(f"       server_file = {server_file}  ({time.time()-t0:.1f}s)")

    dl_url = f"{DOWNLOAD_API}?fileName={urllib.parse.quote(server_file)}&delete=true"
    r3 = s.get(dl_url, timeout=600, stream=True,
               headers={"Referer": WEIGHTMEMO_PAGE})
    r3.raise_for_status()

    # 根据文件头决定扩展名（99% 是 .xlsx，但兼容服务器偶发 .xls）
    head = r3.content[:16]
    ext = _detect_ext(head)
    final_path = out_path
    if not final_path.lower().endswith(ext):
        if final_path.lower().endswith((".xlsx", ".xls")):
            final_path = final_path.rsplit(".", 1)[0] + ext
        else:
            final_path = final_path + ext

    with open(final_path, "wb") as f:
        for chunk in r3.iter_content(chunk_size=65536):
            if chunk:
                f.write(chunk)
    size = os.path.getsize(final_path)
    rows = -1
    if _HAS_PANDAS and size > 0:
        try:
            if ext == ".xlsx":
                df = pd.read_excel(final_path, engine="openpyxl")
            else:
                df = pd.read_excel(final_path)
            rows = len(df)
        except Exception as e:
            print(f"       ⚠️ pandas 读 xlsx 失败: {e}，跳过行数统计")
    if verbose:
        print(f"       已保存: {final_path}  ({size:,} bytes)  rows≈{rows}  ext={ext}")
    return final_path, rows


# ============================================================
# Stage 2: 飞书云盘上传
# ============================================================
def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")
    r = requests.post(
        f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret},
        timeout=15, proxies=NO_PROXY)
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"tenant_access_token 失败: {d}")
    tok = str(d.get("tenant_access_token") or "")
    if not tok:
        raise RuntimeError(f"响应里没有 tenant_access_token: {d}")
    return tok


def feishu_upload_file(tok: str, folder_token: str, file_path: str,
                       max_size_mb: int = 50) -> dict:
    if not os.path.exists(file_path):
        raise FileNotFoundError(file_path)
    size = os.path.getsize(file_path)
    if size == 0:
        raise ValueError("上传文件为空")
    if size > max_size_mb * 1024 * 1024:
        raise ValueError(f"文件过大 {size/1024/1024:.1f}MB > 上限 {max_size_mb}MB，"
                         f"请改用分片上传（暂未实现）")
    if not folder_token:
        raise ValueError("缺少 folder_token")

    filename = os.path.basename(file_path)
    with open(file_path, "rb") as fh:
        raw = fh.read()
    data = {
        "file_name": filename,
        "parent_type": "explorer",
        "parent_node": folder_token,
        "size": str(size),
    }
    files = {"file": (filename, raw, "application/octet-stream")}
    r = requests.post(
        f"{FEISHU_OPEN_BASE}/drive/v1/files/upload_all",
        data=data, files=files,
        headers={"Authorization": f"Bearer {tok}"},
        timeout=180, proxies=NO_PROXY,
    )
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"上传飞书失败 code={d.get('code')} msg={d.get('msg')}  data={d}")
    info = d.get("data") or {}
    ft = info.get("file_token") or info.get("token") or ""
    name = info.get("name") or filename
    print(f"[飞书云盘] ✅ 上传成功  name={name}  file_token={ft}  "
          f"url=https://s2v31ke6sl.feishu.cn/drive/file/{ft}")
    return info


def _feishu_list_recent_madan(tok: str, folder_token: str, *,
                              within_hours: int = 2) -> list[dict]:
    """列出云盘 folder 内最近 within_hours 小时内创建的 码单_*.xlsx 文件。
    返回 [{name, token, created_time, modified_time}, ...]，按创建时间倒序。
    """
    import re as _re
    name_re = _re.compile(rf"^{_re.escape(FILE_PREFIX)}_(\d{{6}})_(\d{{6}})\.xlsx$")
    items: list[dict] = []
    page_token = ""
    now_ts = time.time()
    cutoff = now_ts - within_hours * 3600
    while True:
        params: dict = {"folder_token": folder_token, "page_size": 50,
                        "order_by": "EditedTime", "direction": "DESC"}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(f"{FEISHU_OPEN_BASE}/drive/v1/files",
                         headers={"Authorization": f"Bearer {tok}"},
                         params=params, timeout=20, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") != 0:
            break
        files = (d.get("data") or {}).get("files") or []
        for f in files:
            name = f.get("name", "")
            if not name_re.match(name):
                continue
            cts = int(f.get("created_time") or 0)
            # Feishu created_time 有时是秒有时是毫秒：兼容一下
            if cts > 10_000_000_000:
                cts = cts // 1000
            if cts and cts < cutoff:
                continue  # 超过时间窗口，不再看（列表是倒序，所以后面只会更旧）
            tok_ = f.get("token") or f.get("file_token") or ""
            if not tok_:
                continue
            items.append({"name": name, "token": tok_, "created_time": cts})
        if not (d.get("data") or {}).get("has_more"):
            break
        page_token = (d.get("data") or {}).get("page_token") or ""
        if not page_token:
            break
    items.sort(key=lambda x: x["created_time"], reverse=True)
    return items


def _feishu_file_bundle_ids(tok: str, file_token: str) -> Optional[set[str]]:
    """下载云盘 xlsx，提取'捆包号'集合。失败返回 None。"""
    try:
        r = requests.get(f"{FEISHU_OPEN_BASE}/drive/v1/files/{file_token}/download",
                         headers={"Authorization": f"Bearer {tok}"},
                         timeout=120, proxies=NO_PROXY)
        if r.status_code != 200:
            return None
        import tempfile, io
        if not _HAS_PANDAS:
            return None
        data = io.BytesIO(r.content)
        ext = _detect_ext(r.content[:16])
        if ext == ".xlsx":
            df = pd.read_excel(data, engine="openpyxl", dtype=str).fillna("")
        else:
            df = pd.read_excel(data, dtype=str).fillna("")
        if "捆包号" not in df.columns:
            return None
        return set(df["捆包号"].astype(str).str.strip())
    except Exception:
        return None


def _local_file_bundle_ids(file_path: str) -> Optional[set[str]]:
    """读取本地 xlsx，提取'捆包号'集合。失败返回 None。"""
    if not _HAS_PANDAS:
        return None
    try:
        if file_path.lower().endswith(".xlsx"):
            df = pd.read_excel(file_path, engine="openpyxl", dtype=str).fillna("")
        else:
            df = pd.read_excel(file_path, dtype=str).fillna("")
        if "捆包号" not in df.columns:
            return None
        return set(df["捆包号"].astype(str).str.strip())
    except Exception:
        return None


def stage_upload(file_path: str, folder_token: str) -> Optional[dict]:
    app_id = env("FEISHU_APP_ID")
    app_secret = env("FEISHU_APP_SECRET")
    if not app_id or not app_secret:
        print("[飞书云盘] ⚠️  未配置 FEISHU_APP_ID / FEISHU_APP_SECRET，跳过上传")
        return None
    if not folder_token:
        print("[飞书云盘] ⚠️  未配置 FEISHU_FOLDER_TOKEN，跳过上传")
        return None
    try:
        tok = feishu_tenant_access_token(app_id, app_secret)

        # ========= 防重复：近 2 小时内云盘若已存在"捆包号完全一致"的码单文件则跳过 =========
        local_ids = _local_file_bundle_ids(file_path)
        if local_ids:
            recent = _feishu_list_recent_madan(tok, folder_token, within_hours=2)
            if recent:
                print(f"[飞书云盘] 🔍 云盘近 2 小时已上传 {len(recent)} 个码单文件，比对捆包号集合…")
                for f in recent:
                    remote_ids = _feishu_file_bundle_ids(tok, f["token"])
                    if remote_ids and remote_ids == local_ids:
                        print(f"[飞书云盘] ⚠️  跳过重复上传：云盘已有 {f['name']} "
                              f"（捆包号 {len(local_ids)} 个完全一致，本地上传文件名={os.path.basename(file_path)}）")
                        # 返回一个伪结果，方便 main() 走"已上传"逻辑
                        return {"_duplicate": True, "name": os.path.basename(file_path),
                                "existing_file": f["name"],
                                "token": f["token"]}
                print(f"[飞书云盘]    未发现重复捆包集合（对比了 {len(recent)} 个文件），继续上传…")
        # ================================================================================

        return feishu_upload_file(tok, folder_token, file_path)
    except Exception as e:
        print(f"[飞书云盘] ❌ 上传失败: {e}")
        return None


# ============================================================
# Stage 3: 飞书机器人通知（Webhook + 私信 Union ID 双通道）
# ============================================================
def _feishu_sign(secret: str, ts: str) -> str:
    import base64 as _b64
    h = hmac.new(secret.encode("utf-8"),
                 f"{ts}\n{secret}".encode("utf-8"), hashlib.sha256)
    return _b64.b64encode(h.digest()).decode("utf-8")


def stage_notify_webhook(message: str) -> None:
    url = env("FEISHU_WEBHOOK_URL")
    if not url:
        return
    body = {"msg_type": "text", "content": {"text": message}}
    secret = env("FEISHU_WEBHOOK_SECRET")
    if secret:
        ts = str(int(time.time()))
        body["timestamp"] = ts
        body["sign"] = _feishu_sign(secret, ts)
    try:
        r = requests.post(url, json=body, timeout=10, proxies=NO_PROXY)
        d = r.json() if r.content else {}
        code = d.get("code") if isinstance(d, dict) else None
        ok = (code == 0) or d.get("StatusCode") == 0 or r.status_code == 200
        print(f"[通知·Webhook] {'OK' if ok else 'FAIL'}  code={code} status={r.status_code} "
              f"msg={(d.get('msg') if isinstance(d, dict) else '') or r.text[:100]}")
    except Exception as e:
        print(f"[通知·Webhook] 异常: {e}")


def _id_type(oid: str) -> str:
    if "@" in oid:
        return "email"
    if oid.startswith("on_"):
        return "union_id"
    if oid.startswith("ou_"):
        return "user_id"
    if oid.startswith("oc_"):
        return "open_id"
    if oid.isdigit() and 8 <= len(oid) <= 20:
        return "mobile"
    return "open_id"


def feishu_send_to_user(tenant_tok: str, open_id: str, text: str,
                        *, receive_id_type: Optional[str] = None) -> Tuple[bool, str]:
    if not tenant_tok or not open_id:
        return False, "missing token/open_id"
    id_type = receive_id_type or _id_type(open_id)
    try:
        r = requests.post(
            f"{FEISHU_OPEN_BASE}/im/v1/messages",
            headers={"Authorization": f"Bearer {tenant_tok}",
                     "Content-Type": "application/json; charset=utf-8"},
            params={"receive_id_type": id_type},
            json={"receive_id": open_id, "msg_type": "text",
                  "content": json.dumps({"text": text}, ensure_ascii=False)},
            timeout=15, proxies=NO_PROXY,
        )
        d = r.json()
        if d.get("code") == 0:
            return True, f"type={id_type} message_id={(d.get('data') or {}).get('message_id', '')}"
        return False, f"type={id_type} code={d.get('code')} msg={d.get('msg', '')}"
    except Exception as e:
        return False, f"type={id_type} exception={e}"


def stage_notify_dm(message: str, union_ids: list[str],
                    notify_app_id: str, notify_app_secret: str,
                    default_app_secret: str) -> None:
    """向若干 Union ID 发飞书私信。若 NOTIFY_APP_ID/SECRET 都没配就回退到 Bitable 应用。"""
    if not union_ids:
        return
    # 1) 取 token：优先专用通知应用（如果配了 SECRET），否则 Bitable 应用
    tok = ""
    hint_used = ""
    if notify_app_id and notify_app_secret:
        try:
            tok = feishu_tenant_access_token(notify_app_id, notify_app_secret)
            hint_used = f"通知专用 App {notify_app_id}"
        except Exception as e:
            print(f"[通知·私信] 专用 App token 失败: {e}  → 回退 Bitable 应用")
    if not tok and notify_app_id and default_app_secret:
        try:
            tok = feishu_tenant_access_token(notify_app_id, default_app_secret)
            hint_used = f"通知专用 App {notify_app_id} (Secret 回退 FEISHU_APP_SECRET)"
        except Exception as e:
            print(f"[通知·私信] 专用 App(回退 Bitable secret) 失败: {e}")
    if not tok:
        print("[通知·私信] 无可用 token，SKIP")
        return
    print(f"[通知·私信] 使用: {hint_used}")
    NAME = {
        "on_93da40c6314edbfa2dc3e031ef405389": "洪",
        "on_b09bcbf3e74f5d423900aa9b2f00eb63": "王阳",
    }
    for oid in union_ids:
        ok, info = feishu_send_to_user(tok, oid, message)
        name = NAME.get(oid, oid)
        print(f"[通知·私信] {'OK  ' if ok else 'FAIL'} → {name}({oid})  {info}")
        time.sleep(0.15)


def _parse_notify_users(s: str) -> list[str]:
    if not s:
        return []
    parts = re.split(r"[\s,，;；]+", s.strip())
    return [p for p in parts if p]


# ============================================================
# Main
# ============================================================
def parse_args():
    p = argparse.ArgumentParser(
        description="IEC 出厂码单 → 码单捆包下载 → 飞书云盘文件夹 DfQdfSxl2ld25wdx6Rxcub9hnDf",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--days", type=int, default=5,
                   help="出厂日期范围：近 N 天（含今日，默认 5）。"
                        "被 --date-start/end 覆盖。")
    p.add_argument("--date-start", type=_parse_yyyymmdd, default="",
                   help="出厂日期起 YYYYMMDD（默认按 --days 计算）")
    p.add_argument("--date-end", type=_parse_yyyymmdd, default="",
                   help="出厂日期止 YYYYMMDD（默认今天）")
    # 可选筛选
    p.add_argument("--memo-code", default="", help="按 码单号 精确筛选（可选）")
    p.add_argument("--contract", default="", help="按 销售合同号 筛选（可选）")
    p.add_argument("--factory-order", default="", help="按 钢厂订单号 筛选（可选）")
    p.add_argument("--shopsign", default="", help="按 牌号 筛选（可选）")
    p.add_argument("--prod-code", default="", help="按 品种代码 筛选（可选）")
    p.add_argument("--delivery-start", type=_parse_yyyymmdd, default="",
                   help="交货月起 YYYYMMDD（可选，传空即不筛）")
    p.add_argument("--delivery-end", type=_parse_yyyymmdd, default="",
                   help="交货月止 YYYYMMDD（可选）")

    p.add_argument("-o", "--output", default="",
                   help="输出文件名（默认自动命名：码单_YYMMDD_HHMMSS.xlsx）")
    p.add_argument("--folder-token", default="",
                   help=f"飞书文件夹 token（默认 {DEFAULT_FOLDER_TOKEN}）")

    p.add_argument("--dry-run", action="store_true",
                   help="只下载不上传（等价 --no-upload / DRY_RUN=1）")
    p.add_argument("--no-upload", action="store_true", help="同 --dry-run")

    # 通知
    p.add_argument("--notify-user", dest="notify_users", action="append",
                   default=None, metavar="ID",
                   help="飞书私信 ID（可重复传），支持 on_/oc_/ou_/email/手机号，自动识别类型")
    p.add_argument("--no-notify-default", action="store_true",
                   help="关闭默认 2 人推送（等价 NOTIFY_DISABLE_DEFAULT=1）")
    p.add_argument("--notify-app-id", default="",
                   help="私信专用 App ID（默认 cli_aaf0ce1e9ef89d27）")
    p.add_argument("--notify-app-secret", default="",
                   help="私信专用 App Secret（默认读 FEISHU_NOTIFY_APP_SECRET，"
                        "没配则用 FEISHU_APP_SECRET）")
    return p.parse_args()


def main():
    args = parse_args()
    dry_run = (args.dry_run or args.no_upload or env("DRY_RUN") == "1")

    # 日期
    ds = args.date_start or env("EXPORT_DATE_START")
    de = args.date_end or env("EXPORT_DATE_END")
    if not ds or not de:
        ds, de = _default_date_range(max(1, args.days))

    # 文件名
    out_path = args.output or _auto_filename()

    # 云盘
    folder_token = args.folder_token or env("FEISHU_FOLDER_TOKEN") or DEFAULT_FOLDER_TOKEN

    # 通知名单
    disable_default = args.no_notify_default or env("NOTIFY_DISABLE_DEFAULT") == "1"
    notify_ids: list[str] = list(_parse_notify_users(env("NOTIFY_USERS")))
    if args.notify_users:
        for x in args.notify_users:
            notify_ids.extend(_parse_notify_users(x))
    if not notify_ids and not disable_default:
        notify_ids = list(DEFAULT_NOTIFY_UNION_IDS)
    _seen: set[str] = set()
    notify_ids = [o for o in notify_ids if not (o in _seen or _seen.add(o))]

    notify_app_id = (args.notify_app_id or env("FEISHU_NOTIFY_APP_ID")
                     or DEFAULT_NOTIFY_APP_ID)
    notify_app_secret = (args.notify_app_secret or env("FEISHU_NOTIFY_APP_SECRET")
                         or env("FEISHU_APP_SECRET") or "")

    print("=" * 66)
    print("IEC 入库管理 → 出厂码单·码单捆包下载 → 飞书云盘")
    print(f"  出厂日期: {ds} ~ {de}   "
          f"({'DRY-RUN  只下载不上传' if dry_run else '上传飞书'})")
    print(f"  输出文件: {out_path}")
    print(f"  云盘 folder_token = {folder_token}")
    print(f"  私信通知: {notify_ids if notify_ids else '(off)'}")
    print(f"  通知 App: {notify_app_id}")
    print("=" * 66)

    t0 = time.time()
    file_path: Optional[str] = None
    rows = -1
    upload_info: Optional[dict] = None
    rc = 0

    try:
        # Stage 1: IEC 下载
        print("\n[Stage 1] 登录 IEC 并下载码单捆包 xlsx ...")
        try:
            iec, s, tok = iec_login_and_enter()
            file_path, rows = download_bundle(
                out_path,
                date_start=ds, date_end=de,
                memo_code=args.memo_code,
                contract_num=args.contract,
                factory_order_num=args.factory_order,
                shopsign=args.shopsign,
                prod_code=args.prod_code,
                delivery_month_start=args.delivery_start,
                delivery_month_end=args.delivery_end,
                s=s, token=tok,
            )
        except Exception as e:
            print(f"[Stage 1] ❌ 下载失败: {e}")
            import traceback; traceback.print_exc()
            rc = 2

        # Stage 2: 上传
        print("\n[Stage 2] 飞书云盘上传 ...")
        if rc == 0 and file_path:
            if dry_run:
                print("  [DRY-RUN] 跳过上传")
            else:
                upload_info = stage_upload(file_path, folder_token)

        # Stage 3: 通知
        elapsed = time.time() - t0
        if upload_info:
            if upload_info.get("_duplicate"):
                ft = upload_info.get("token") or ""
                upload_line = (f"⏭️  跳过重复：已有 {upload_info.get('existing_file','')} "
                               f"→ https://s2v31ke6sl.feishu.cn/drive/file/{ft}")
            else:
                ft = upload_info.get("file_token") or upload_info.get("token") or ""
                upload_line = f"已上传 → https://s2v31ke6sl.feishu.cn/drive/file/{ft}"
        else:
            upload_line = ("跳过(DRY-RUN)" if dry_run else
                           ("失败/未配置" if rc == 0 else "下载失败，未上传"))

        message = (
            f"【IEC 码单捆包→飞书】{'DRY-RUN ' if dry_run else ''}"
            f"{time.strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"出厂日期: {ds} ~ {de}\n"
            f"数据: {rows} 行  文件: {os.path.basename(file_path) if file_path else '(未生成)'}\n"
            f"云盘({folder_token[:10]}…): {upload_line}\n"
            f"耗时: {elapsed:.1f}s"
        )
        print("\n" + "=" * 66)
        print(message)
        print("=" * 66)

        # 通知暂时关闭（用户要求）
        # if rc == 0:
        #     stage_notify_webhook(message)
        #     stage_notify_dm(message, notify_ids,
        #                     notify_app_id=notify_app_id,
        #                     notify_app_secret=env("FEISHU_NOTIFY_APP_SECRET") or "",
        #                     default_app_secret=env("FEISHU_APP_SECRET") or "")
    finally:
        # 确保 iec token 持久化到 iecc.json，避免下次再跑验证码
        try:
            iec.save("iecc.json")
        except Exception:
            pass

    return rc


if __name__ == "__main__":
    raise SystemExit(main())
