#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IEC「生产跟踪」查询 → 飞书多维表 tblvugnoJPS8GrpX 按"资源号"批量标记"结案"
============================================================================

数据来源：IEC 生产跟踪页面 queryProduction 接口
    url   : POST /iecs/freight/goodsProduction/goodsProduction/queryProduction
    filter: deliveryDateChrStart / End (交货期 YYYYMM，默认前4月~后1月)
    解析  : HTML tbody 片段 → BeautifulSoup 找所有 <tr>

业务规则（按用户要求）：
    1. 从 IEC 拉「生产跟踪」数据（交期范围默认前4月~后1月）
    2. 本地筛 "生产状态"(td[11]) == "已完成"  的记录
    3. 取其 "钢厂订单号"(td[4]，格式 X6E0008431 / L4E0006973 等) 作为"资源号"
    4. 全量拉取 bitable tblvugnoJPS8GrpX 记录，构建 (资源号列表) → record_id 映射
    5. 资源号命中 → 且当前「进度」≠ "结案" →  batch_update 进度为 "结案"

关联关系：
    IEC  钢厂订单号(td[4])    ←→   飞书 「资源号」字段(Lookup，值形如 ["L4E0006973"])
    IEC  生产状态(td[11])=已完成  →  飞书 「进度」字段(SingleSelect) = "结案"

飞书进度单选选项：炼钢 / 热轧 / 酸洗 / 冷轧 / 准发 / 结案 / 撤单

环境变量（必填）：
    IBAO_USERNAME / IBAO_PASSWORD           IEC 登录
    FEISHU_APP_ID / FEISHU_APP_SECRET       飞书自建应用
    BITABLE_APP_TOKEN                       Tz0XbQVzkaZuJasBwb8cRjkfnoe (可默认)
    BITABLE_TABLE_ID                        tblvugnoJPS8GrpX (可默认)

环境变量（可选）：
    EXPORT_START / EXPORT_END    交期 YYYYMM
    DRY_RUN=1                    只打印不实际 UPDATE
    UPDATE_LIMIT=0               最大 UPDATE 条数（0=不限）
    BITABLE_PAGE_SIZE=500        多维表分页大小（默认 500）
    IEC_PAGE_SIZE=100            IEC 分页大小（默认 100）
    VERBOSE=1                    详细日志（打印每个资源号匹配/更新明细）
    LOG_MATCH_FILE=xxx.csv       把匹配过程按行写到 CSV（调试用，留空不写）

运行:
    python iec_jiean_to_bitable.py --dry-run                  # 预览
    python iec_jiean_to_bitable.py --verbose                  # 详细日志
    python iec_jiean_to_bitable.py --log-match match.csv      # 输出匹配日志到 CSV
"""
from __future__ import annotations

import argparse
import csv
import datetime
import hashlib
import hmac
import io
import json
import os
import re
import sys
import time
import urllib.parse
from collections import Counter
from typing import Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from bs4 import BeautifulSoup
    _HAS_BS4 = True
except ImportError:
    _HAS_BS4 = False

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ibaosteel_client import IEC


# ---------------- 常量 ----------------
BASE = "https://www.ibaosteel.com"
IECS_INDEX = f"{BASE}/iecs/index"
PROD_PAGE = f"{BASE}/iecs/freight/goodsProduction/goodsProduction/initLoads"
QUERY_API = f"{BASE}/iecs/freight/goodsProduction/goodsProduction/queryProduction"

FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}

# Bitable 默认值
DEFAULT_APP_TOKEN = "Tz0XbQVzkaZuJasBwb8cRjkfnoe"
DEFAULT_TABLE_ID = "tblvugnoJPS8GrpX"
# 字段名
FIELD_ZIYUANHAO = "资源号"   # Lookup，值为 ["L4E0006973"]
FIELD_JINDU = "进度"     # SingleSelect，写 "结案"
JINDU_TARGET = "结案"

# 进度单选项（已知存在），但写入时直接传字符串即可，Feishu 会匹配选项名
SETTLE_USER_NUM = "062122"
USER_NUM = "U41634"
U_CODE = "U41634"
SALE_NETWORK = "E"
SYSTEM = "ES"

# td 位置（按生产跟踪 <th> 表头）：
# td[0] checkbox/隐藏字段
# td[1] 销售订单号         td[2] 销售订单子项号      td[3] 订单状态(labelcode)
# td[4] 钢厂订单号(=资源号) td[5] 品种               td[6] 规格描述
# td[7] 牌号               td[8] 订货量              td[9] 已生产
# td[10] 客户母材编号      td[11] 生产状态(已完成→结案)  td[12] 生产时间
# td[13] 交货期            td[14] 订货用户(需方)      td[15] 用户需求号
# td[16] 最终用户
COL_ZIYUANHAO = 4
COL_SHENGCHAN_STATUS = 11
SHENGCHAN_DONE = "已完成"


# ---------------- 全局日志控制 ----------------
class MatchLog:
    """统一的详细匹配/更新日志器：既能 stdout 打印（--verbose），也能写 CSV（--log-match）。

    用法:
        ml = MatchLog(verbose=True, csv_path="match.csv")
        ml.add(stage="IEC", row=1, ziyuan="X6E0008431", sc_status="已完成", ...)
        ml.add(stage="BITABLE_INDEX", row=3, ziyuan="X6E0008431", ...)
        ml.add(stage="MATCH", ...)
        ml.add(stage="UPDATE", ...)
        ml.print_summary()
        ml.close()
    """

    CSV_COLUMNS = [
        "ts",                 # 秒级时间戳 YYYY-MM-DD HH:MM:SS
        "stage",              # IEC / BITABLE_INDEX / MATCH / UPDATE
        "seq",                # 序号（每阶段自增）
        "资源号",
        "record_id",          # 飞书 record_id（IEC/BITABLE_INDEX 阶段可能为空）
        "销售订单号",         # IEC 销售订单号（可选）
        "子项号",             # IEC 子项号（可选）
        "IEC_生产状态",        # IEC 端 生产状态
        "IEC_订单状态",        # IEC 端 订单状态
        "Bitable_当前进度",    # 飞书进度字段 当前值
        "匹配结果",           # 命中/未命中/重复/覆盖/已结案跳过/计划更新/更新成功/更新失败/被LIMIT截断/更新跳过
        "细节",               # 额外说明（如被哪个 record_id 覆盖、失败 reason）
    ]

    def __init__(self, verbose: bool = False, csv_path: str = ""):
        self.verbose = verbose
        self.csv_path = csv_path.strip()
        self._csv_fh = None
        self._csv_w = None
        self._seq: dict[str, int] = {}
        if self.csv_path:
            self._csv_fh = open(self.csv_path, "w", encoding="utf-8-sig", newline="")
            self._csv_w = csv.writer(self._csv_fh)
            self._csv_w.writerow(self.CSV_COLUMNS)
            if self.verbose:
                print(f"[日志] 匹配/更新明细 CSV: {os.path.abspath(self.csv_path)}")

    @staticmethod
    def _ts() -> str:
        return time.strftime("%Y-%m-%d %H:%M:%S")

    def add(self, stage: str, *, ziyuan: str = "", record_id: str = "",
            row: Optional[int] = None,
            销售订单号: str = "", 子项号: str = "",
            IEC_生产状态: str = "", IEC_订单状态: str = "",
            Bitable_当前进度: str = "",
            匹配结果: str = "",
            细节: str = "") -> None:
        seq = self._seq.get(stage, 0) + 1
        self._seq[stage] = seq
        seq_display = row if row is not None else seq
        row_data = [
            self._ts(), stage, seq_display, ziyuan, record_id,
            销售订单号, 子项号,
            IEC_生产状态, IEC_订单状态,
            Bitable_当前进度, 匹配结果, 细节,
        ]
        if self._csv_w is not None:
            self._csv_w.writerow(row_data)
        if self.verbose:
            # 构造单行紧凑文本（太长就截断细节）
            msg = (f"[LOG:{stage}#{seq_display}] 资源号={ziyuan or '-'}")
            if record_id:
                msg += f" rid={record_id}"
            if IEC_生产状态:
                msg += f" IEC:生产={IEC_生产状态} 订单={IEC_订单状态}"
            if Bitable_当前进度 != "":
                msg += f" 进度={Bitable_当前进度!r}"
            if 匹配结果:
                msg += f" → {匹配结果}"
            if 细节:
                d = 细节 if len(细节) < 120 else 细节[:117] + "..."
                msg += f"  细节:{d}"
            print(msg)

    def print_summary(self):
        if not self._seq:
            return
        print(f"[日志汇总] 各阶段记录数: {dict(self._seq)}")
        if self.csv_path and self._csv_fh:
            size = os.path.getsize(self.csv_path)
            print(f"[日志汇总] CSV 已写入: {os.path.abspath(self.csv_path)}  ({size:,} bytes)")

    def close(self):
        if self._csv_fh is not None:
            try:
                self._csv_fh.close()
            except Exception:
                pass
            self._csv_fh = None


# ---------------- 通用 ----------------
def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def create_session() -> requests.Session:
    s = requests.Session()
    retry = Retry(total=5, backoff_factor=1,
                  status_forcelist=[500, 502, 503, 504, 429],
                  allowed_methods=["POST", "GET"])
    adapter = HTTPAdapter(max_retries=retry)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    return s


def _default_range() -> tuple[str, str]:
    """默认交货期范围：当前月往前 4 个月到往后 1 个月 (用户指定)"""
    today = datetime.date.today()
    sm, sy = today.month - 4, today.year
    if sm <= 0:
        sm += 12
        sy -= 1
    em, ey = today.month + 1, today.year
    if em > 12:
        em -= 12
        ey += 1
    return f"{sy:04d}{sm:02d}", f"{ey:04d}{em:02d}"


def _parse_month(s: str) -> str:
    s = str(s).strip()
    if s == "":
        return ""  # 空表示没传，由 main 里再走 EXPORT_* / default range
    if len(s) == 6 and s.isdigit():
        y, m = int(s[:4]), int(s[4:6])
        if 2000 <= y <= 2100 and 1 <= m <= 12:
            return s
    raise argparse.ArgumentTypeError(f"月份格式错误: {s!r}（应为 YYYYMM）")


# ============================================================
# Stage 1: 登录 IEC + 分页拉 queryProduction → 解析成 dict 列表
# ============================================================
def iec_login_and_enter() -> tuple[IEC, requests.Session, str]:
    """登录 IEC 并进入生产跟踪页面。返回 (iec, session, access_token)"""
    if not _HAS_BS4:
        raise ImportError("需要 beautifulsoup4：pip install beautifulsoup4")

    iec = IEC.load("iecc.json")
    ok = iec.ok or iec.login()
    if not ok:
        # 重新登录
        iec = IEC()
        if not iec.login():
            raise RuntimeError("IEC 登录失败")
    iec.save("iecc.json")
    s = iec.session
    token = iec.token
    # 注意：IEC 后端只认 access_token 参数（不是 token），传 token= 会直接 401
    ref = f"{IECS_INDEX}?access_token={token}"
    s.get(ref, timeout=20, headers={"Referer": ref}).raise_for_status()
    s.get(PROD_PAGE, timeout=20, headers={"Referer": ref}).raise_for_status()
    return iec, s, token


def _query_one_page(s: requests.Session, token: str, start: str, end: str,
                    page_num: int, page_size: int) -> tuple[int, list[dict]]:
    """查询一页。返回 (total_rows, rows_of_dict)，rows 可能为空。"""
    select_form = {
        "contractNum": "", "factoryProductId": "", "materialCode": "",
        "prodCode": "", "contractStatus": "", "tradeCode": "",
        "deliveryDateChrStart": start, "deliveryDateChrEnd": end,
        "custOrdNum": "", "shopsigns": "", "sizeDesc": "",
        "finUserNum": "", "finUserName": "", "memberCodeFlag": "",
        "saleVarCode": "",
        "system": SYSTEM, "userNum": USER_NUM, "uCode": U_CODE,
        "saleNetwork": SALE_NETWORK, "settleUserNum": SETTLE_USER_NUM,
        "access_token": token,
        "pageDomain": {"pageNum": page_num, "pageSize": page_size},
        "orderByColumn": "", "isAsc": "",
        "pageNum": page_num, "pageSize": page_size,
    }
    r = s.post(QUERY_API, data=json.dumps(select_form), timeout=60,
               headers={"Content-Type": "application/json; charset=utf-8",
                        "X-Requested-With": "XMLHttpRequest",
                        "Accept": "application/json, text/javascript, */*; q=0.01",
                        "Referer": PROD_PAGE})
    r.raise_for_status()
    body = r.text
    m = re.search(r'<input[^>]+id=["\']total["\'][^>]*value=["\'](\d+)["\']', body, re.I)
    total = int(m.group(1)) if m else 0
    # 解析 <tr>
    rows: list[dict] = []
    if not _HAS_BS4:
        return total, rows
    soup = BeautifulSoup(body, "html.parser")
    for tr in soup.find_all("tr"):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 12:
            tds = tr.find_all("td")
        if len(tds) < 12:
            continue
        cells = [td.get_text(strip=True) for td in tds]
        ziyuan = cells[COL_ZIYUANHAO].strip() if len(cells) > COL_ZIYUANHAO else ""
        sc_status = cells[COL_SHENGCHAN_STATUS].strip() if len(cells) > COL_SHENGCHAN_STATUS else ""
        if not ziyuan:
            continue
        rows.append({
            "资源号": ziyuan,
            "生产状态": sc_status,
            "销售订单号": cells[1].strip() if len(cells) > 1 else "",
            "子项号": cells[2].strip() if len(cells) > 2 else "",
            "订单状态": tds[3].get_text(strip=True) if len(tds) > 3 else "",
            "品种": cells[5].strip() if len(cells) > 5 else "",
            "规格": cells[6].strip() if len(cells) > 6 else "",
            "牌号": cells[7].strip() if len(cells) > 7 else "",
            "订货量": cells[8].strip() if len(cells) > 8 else "",
            "已生产": cells[9].strip() if len(cells) > 9 else "",
            "交货期": cells[13].strip() if len(cells) > 13 else "",
            "订货用户": cells[14].strip() if len(cells) > 14 else "",
            "最终用户": cells[16].strip() if len(cells) > 16 else "",
        })
    return total, rows


def iec_pull_all(start: str, end: str, page_size: int = 100,
                 dry_run: bool = False,
                 limit: int = 0) -> list[dict]:
    """分页拉取 IEC 生产跟踪全部记录。返回完整 rows 列表。"""
    iec, s, token = iec_login_and_enter()
    print(f"[IEC] 登录 OK，查询交期 {start}~{end}  page_size={page_size}")
    # 先查 1 页拿 total
    total, rows = _query_one_page(s, token, start, end, 1, page_size)
    print(f"[IEC] 总条数: {total}    第 1 页: {len(rows)} 条")
    all_rows = list(rows)
    if total > len(rows) and len(rows) == page_size:
        import math
        total_pages = math.ceil(total / page_size)
        for pn in range(2, total_pages + 1):
            _, r2 = _query_one_page(s, token, start, end, pn, page_size)
            all_rows.extend(r2)
            if pn % 5 == 0 or pn == total_pages:
                print(f"[IEC]   已拉 {pn}/{total_pages} 页   累计 {len(all_rows)} 条")
            if limit and len(all_rows) >= limit:
                all_rows = all_rows[:limit]
                break
            if not r2:
                break
            time.sleep(0.15)
    print(f"[IEC] 合计 {len(all_rows)} 行")
    iec.save("iecc.json")
    return all_rows


def iec_filter_done(rows: list[dict], *, log: Optional[MatchLog] = None) -> set[str]:
    """从 IEC 数据里筛出 生产状态="已完成" 的资源号集合。启用 MatchLog 时逐行记录。"""
    done: set[str] = set()
    sc_counter: dict[str, int] = {}
    for idx, r in enumerate(rows, 1):
        k = r["生产状态"] or "(空)"
        sc_counter[k] = sc_counter.get(k, 0) + 1
        is_done = (r["生产状态"] == SHENGCHAN_DONE)
        if is_done:
            done.add(r["资源号"])
        if log is not None:
            log.add(
                stage="IEC", row=idx,
                ziyuan=r["资源号"],
                销售订单号=r.get("销售订单号", ""),
                子项号=r.get("子项号", ""),
                IEC_生产状态=r.get("生产状态", ""),
                IEC_订单状态=r.get("订单状态", ""),
                匹配结果=("保留(生产=已完成)" if is_done else "过滤(生产≠已完成)"),
                细节=f"品种={r.get('品种','')}  牌号={r.get('牌号','')}  交货期={r.get('交货期','')}",
            )
    print(f"[IEC] 生产状态分布: {sc_counter}")
    print(f"[IEC] → 生产状态='{SHENGCHAN_DONE}' 共 {len(done)} 个唯一资源号")
    if log is not None:
        # 再在 summary 里加一条分布
        print(f"[IEC] （已按 VERBOSE / CSV 记录 {len(rows)} 条 IEC 行）")
    return done


# ============================================================
# Stage 2: 飞书多维表格
# ============================================================
def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET 环境变量")
    r = requests.post(f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
                      json={"app_id": app_id, "app_secret": app_secret},
                      timeout=15, proxies=NO_PROXY)
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"tenant_access_token 失败: {d}")
    return str(d.get("tenant_access_token") or "")


def bitable_list_all_records(token: str, app_token: str, table_id: str,
                             page_size: int = 500) -> list[dict]:
    """分页拉多维表全部记录（因为 Lookup 字段不能 filter，内存匹配更稳）。"""
    out: list[dict] = []
    page_token = None
    page = 0
    while True:
        page += 1
        params: dict = {"page_size": page_size}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(
            f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records",
            headers={"Authorization": f"Bearer {token}"},
            params=params, timeout=60, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") != 0:
            raise RuntimeError(f"list records 失败: {json.dumps(d, ensure_ascii=False)[:500]}")
        data = d.get("data") or {}
        items = data.get("items") or []
        out.extend(items)
        has_more = data.get("has_more")
        page_token = data.get("page_token")
        if page % 4 == 0 or not has_more:
            print(f"[飞书] 拉记录 第 {page} 页  累计 {len(out)} 条  has_more={has_more}")
        if not has_more:
            break
    return out


def _extract_ziyuanhao(fields: dict) -> list[str]:
    """从记录的 fields 中解析「资源号」字段的值。
    Lookup 字段返回值格式多样：
      - list[str] (示例中是 ['L4E0006973'])
      - list[dict] 也有可能出现，text 字段内拿
    返回一个 list（通常只有 1 个，但兼容多值）。
    """
    val = fields.get(FIELD_ZIYUANHAO)
    if val is None:
        return []
    if isinstance(val, str):
        v = val.strip()
        return [v] if v else []
    if isinstance(val, list):
        out = []
        for x in val:
            if isinstance(x, str):
                if x.strip():
                    out.append(x.strip())
            elif isinstance(x, dict):
                t = str(x.get("text") or x.get("value") or "").strip()
                if t:
                    out.append(t)
            else:
                t = str(x).strip()
                if t:
                    out.append(t)
        return out
    t = str(val).strip()
    return [t] if t else []


def build_index(records: list[dict], *, log: Optional[MatchLog] = None,
                verbose: bool = False) -> tuple[dict[str, str], int, int]:
    """构建 资源号 → record_id 映射（1 资源号对应 1 条最新记录，若重复覆盖）。
    返回 (ziyuan_to_record_id, 记录总数, 无资源号数)。
    启用 MatchLog 时每条 Bitable 记录都会逐行写入 BITABLE_INDEX。
    """
    ziyuan_to_rid: dict[str, str] = {}
    missing = 0
    dup = 0
    for row_idx, rec in enumerate(records, 1):
        rid = rec.get("record_id") or ""
        fields = rec.get("fields") or {}
        zys = _extract_ziyuanhao(fields)
        jd = fields.get(FIELD_JINDU)
        if isinstance(jd, dict):
            jd = str(jd.get("text") or jd.get("name") or "")
        jd = str(jd or "").strip()
        if not zys:
            missing += 1
            if log is not None:
                log.add(
                    stage="BITABLE_INDEX",
                    row=row_idx,
                    record_id=rid,
                    Bitable_当前进度=jd,
                    匹配结果="无资源号(跳过)",
                    细节="资源号字段为空；无法参与匹配",
                )
            continue
        for zy in zys:
            detail = ""
            if zy in ziyuan_to_rid:
                dup += 1
                prev = ziyuan_to_rid[zy]
                detail = f"重复覆盖：旧 rid={prev} → 新 rid={rid}"
                result = "重复(后写入覆盖)"
            else:
                result = "索引(首次加入)"
            ziyuan_to_rid[zy] = rid
            if log is not None:
                log.add(
                    stage="BITABLE_INDEX",
                    row=row_idx,
                    ziyuan=zy,
                    record_id=rid,
                    Bitable_当前进度=jd,
                    匹配结果=result,
                    细节=detail,
                )
    print(f"[索引] 记录总数 {len(records)}  "
          f"→ 唯一资源号 {len(ziyuan_to_rid)}  "
          f"无资源号 {missing}  重复资源号覆盖 {dup}")
    if verbose or (log is not None and (log.verbose or log.csv_path)):
        # 另外打印几个重复资源号的例子（最多 5 个，便于快速排查）
        # 做法：再次遍历 records，按资源号收集 rid 列表，长度>1 即为重复
        ziyuan_all_rids: dict[str, list[str]] = {}
        for rec in records:
            rid = rec.get("record_id") or ""
            fields = rec.get("fields") or {}
            for zy in _extract_ziyuanhao(fields):
                ziyuan_all_rids.setdefault(zy, []).append(rid)
        dups = [(zy, rids) for zy, rids in ziyuan_all_rids.items() if len(rids) > 1]
        if dups:
            print(f"[索引·样例] 发现 {len(dups)} 个重复资源号，前 5 个：")
            for zy, rids in dups[:5]:
                kept = rids[-1]  # 最终保留的（最后覆盖）
                print(f"    资源号={zy}  出现 {len(rids)} 条 → 保留 rid={kept}  "
                      f"其他 {rids[:-1]} 将不会匹配到 IEC 数据")
    return ziyuan_to_rid, len(records), missing


def _batch_update(token: str, app_token: str, table_id: str,
                  records: list[dict]) -> dict:
    """批量 UPDATE：records = [{"record_id": xx, "fields": {进度:结案}}, ...]。"""
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_update"
    r = requests.post(url,
                      headers={"Authorization": f"Bearer {token}",
                               "Content-Type": "application/json; charset=utf-8"},
                      json={"records": records},
                      timeout=60, proxies=NO_PROXY)
    d = r.json()
    if d.get("code") != 0:
        print(f"  ⚠️  batch_update 失败: {json.dumps(d, ensure_ascii=False)[:500]}")
    return d


def bitable_update_jindu(token: str, app_token: str, table_id: str,
                         records: list[dict], done_set: set[str],
                         dry_run: bool = False, update_limit: int = 0,
                         *, log: Optional[MatchLog] = None,
                         verbose: bool = False) -> tuple[int, int, int]:
    """按资源号匹配 → 把进度更新为"结案"。
    返回 (计划条数, 实际成功条数, 失败条数)。
    启用 MatchLog 时：
      - 每个 IEC 已完成资源号会写 1 行 MATCH 阶段记录（命中/未命中/已结案/计划更新/截断）
      - 每条 UPDATE 会写 1 行 UPDATE 阶段记录（成功/失败）
    """
    ziyuan_to_rid, total_records, _ = build_index(records, log=log, verbose=verbose)

    # 交集：IEC 已完成资源号 ∩ Bitable 资源号
    intersect = done_set & set(ziyuan_to_rid.keys())
    not_in_bitable = done_set - intersect  # IEC 有但 Bitable 找不到的
    print(f"[匹配] IEC 已完成资源号 {len(done_set)} 个  ∩  Bitable 资源号 {len(ziyuan_to_rid)} 个  "
          f"= 命中 {len(intersect)} 个  |  IEC 有但 Bitable 无 = {len(not_in_bitable)} 个")

    # 构建待 UPDATE 列表（进度≠结案的才入）
    # 先建 record_id→当前进度 映射加速
    rid_to_jindu: dict[str, str] = {}
    for rec in records:
        rid = rec.get("record_id") or ""
        fields = rec.get("fields") or {}
        jd = fields.get(FIELD_JINDU)
        if isinstance(jd, dict):
            # 有些单选字段返回的是对象
            jd = str(jd.get("text") or jd.get("name") or "")
        rid_to_jindu[rid] = str(jd or "").strip()

    planned: list[dict] = []  # [{record_id, ziyuan, old_jindu}]
    skipped_done = 0
    cutoff_count = 0
    # 按 done_set 全量打日志（这样 MATCH 阶段涵盖每个 IEC 资源号）
    for zy in sorted(done_set):
        if zy not in ziyuan_to_rid:
            # 未命中 Bitable
            if log is not None:
                log.add(stage="MATCH", ziyuan=zy,
                        匹配结果="未命中(Bitable 无此资源号)",
                        细节=f"Bitable 唯一资源号 {len(ziyuan_to_rid)} 中未找到；"
                             f"若确认资源号存在，检查 Lookup 字段显示格式")
            continue
        rid = ziyuan_to_rid[zy]
        old = rid_to_jindu.get(rid, "")
        if old == JINDU_TARGET:
            skipped_done += 1
            if log is not None:
                log.add(stage="MATCH", ziyuan=zy, record_id=rid,
                        Bitable_当前进度=old,
                        匹配结果="已结案(跳过)",
                        细节=f"当前进度已是 {JINDU_TARGET!r}，不做 UPDATE")
            continue
        # 命中 + 需要更新
        if update_limit and len(planned) >= update_limit:
            cutoff_count += 1
            if log is not None:
                log.add(stage="MATCH", ziyuan=zy, record_id=rid,
                        Bitable_当前进度=old,
                        匹配结果=f"被LIMIT截断(>={update_limit})",
                        细节=f"--update-limit={update_limit} 已满，本条未入计划")
            continue
        planned.append({"record_id": rid, "资源号": zy, "old_jindu": old})
        if log is not None:
            log.add(stage="MATCH", ziyuan=zy, record_id=rid,
                    Bitable_当前进度=old,
                    匹配结果="计划更新",
                    细节=f"进度 {old!r} → {JINDU_TARGET!r}  已排队，共 {len(planned)} 条")

    # 额外的概览样例（无论 verbose 与否，对小量数据直接打印；量大时只打印前若干）
    print(f"[匹配·分类] 未命中Bitable={len(not_in_bitable)}  "
          f"命中+已是结案(跳过)={skipped_done}  "
          f"命中+需更新(计划)={len(planned)}  "
          f"被 update_limit 截断={cutoff_count}")
    if not_in_bitable and (verbose or log is not None):
        sample = sorted(not_in_bitable)[:10]
        print(f"  未命中 Bitable 样例({len(sample)}/{len(not_in_bitable)}): {sample}")

    match_stats = {
        "intersect": len(intersect),                # 命中 Bitable 资源号
        "not_in_bitable": len(not_in_bitable),     # IEC 有但 Bitable 无
        "skipped_done": skipped_done,              # 命中但进度已是结案
        "cutoff": cutoff_count,                     # 被 update_limit 截断
        "bitable_ziyuan_unique": len(ziyuan_to_rid),
    }

    print(f"[UPDATE] 计划 {len(planned)} 条（排除进度已是结案的 {skipped_done} 条）")
    if not planned:
        print("  无需要更新的记录")
        return 0, 0, 0, match_stats

    # 预览前 N 条（详细到每条，不只 5 条）
    preview_count = len(planned) if (verbose or len(planned) <= 20) else 5
    print(f"  前 {preview_count} 条计划更新:")
    for i, p in enumerate(planned[:preview_count], 1):
        print(f"    [{i}] 资源号={p['资源号']}  rid={p['record_id']}  "
              f"当前进度={p['old_jindu']!r} → {JINDU_TARGET!r}")

    if dry_run:
        print(f"\n[DRY-RUN] 仅预览，不实际 UPDATE。共 {len(planned)} 条计划更新")
        if log is not None:
            # 即便 dry-run 也写 UPDATE 行标注 DRY-RUN
            for p in planned:
                log.add(stage="UPDATE", ziyuan=p["资源号"], record_id=p["record_id"],
                        Bitable_当前进度=p["old_jindu"],
                        匹配结果="DRY-RUN(未执行)",
                        细节=f"计划更新 {p['old_jindu']!r} → {JINDU_TARGET!r}")
        return len(planned), 0, 0, match_stats

    # 每批 500 条批量更新；逐批打印每批的 资源号 列表，便于核对
    BATCH = 500
    ok_count = 0
    fail_count = 0
    for bi in range(0, len(planned), BATCH):
        batch = planned[bi:bi + BATCH]
        payload = []
        for p in batch:
            payload.append({
                "record_id": p["record_id"],
                "fields": {FIELD_JINDU: JINDU_TARGET},
            })
        print(f"\n  === 批 {bi // BATCH + 1}/{(len(planned) - 1) // BATCH + 1} ===")
        print(f"     提交 {len(payload)} 条: "
              f"{[ (p['资源号'] + '→' + p['old_jindu'] + '→结案') for p in batch ]}")
        d = _batch_update(token, app_token, table_id, payload)
        if d.get("code") == 0:
            records_resp = (d.get("data") or {}).get("records") or []
            # 飞书 batch_update 返回的 records 按提交顺序一一对应
            n = len(records_resp)
            ok_count += n
            print(f"     成功 {n} 条")
            if log is not None:
                for i, p in enumerate(batch):
                    if i < len(records_resp):
                        rid_out = (records_resp[i] or {}).get("record_id") or p["record_id"]
                        log.add(stage="UPDATE", ziyuan=p["资源号"], record_id=rid_out,
                                Bitable_当前进度=p["old_jindu"],
                                匹配结果="更新成功",
                                细节=f"{p['old_jindu']!r} → {JINDU_TARGET!r}  "
                                     f"批#{bi // BATCH + 1} 位置#{i + 1}")
                    else:
                        # 返回条数少于提交（罕见）
                        log.add(stage="UPDATE", ziyuan=p["资源号"], record_id=p["record_id"],
                                Bitable_当前进度=p["old_jindu"],
                                匹配结果="更新失败(缺返回)",
                                细节=f"响应条数 {len(records_resp)} < 提交 {len(batch)}；本条未在响应内")
                        fail_count += 1
        else:
            code, msg = d.get("code"), d.get("msg", "")
            fail_count += len(batch)
            print(f"     失败 {len(batch)}  code={code}  msg={msg}")
            if log is not None:
                for p in batch:
                    log.add(stage="UPDATE", ziyuan=p["资源号"], record_id=p["record_id"],
                            Bitable_当前进度=p["old_jindu"],
                            匹配结果="更新失败(请求错误)",
                            细节=f"code={code}  msg={msg}")
        time.sleep(0.2)
    print(f"\n[UPDATE] 完成。成功 {ok_count}  失败 {fail_count}  "
          f"跳过（已是结案）= {skipped_done}  未命中= {len(not_in_bitable)}  "
          f"被LIMIT截断= {cutoff_count}")
    return len(planned), ok_count, fail_count, match_stats


# ============================================================
# Stage 3: 飞书机器人通知
#   方式 1：webhook 群机器人（FEISHU_WEBHOOK_URL，可选 SECRET 加签）
#   方式 2：自建应用 im/v1/messages 向指定 open_id 单发消息（不走机器人）
# ============================================================
def _feishu_sign(secret: str, timestamp: str) -> str:
    import base64 as _b64
    string_to_sign = f"{timestamp}\n{secret}"
    h = hmac.new(secret.encode("utf-8"),
                 string_to_sign.encode("utf-8"), hashlib.sha256)
    return _b64.b64encode(h.digest()).decode("utf-8")


def stage_notify_webhook(message: str):
    """向 FEISHU_WEBHOOK_URL 发纯文本消息（群机器人）。"""
    webhook_url = env("FEISHU_WEBHOOK_URL")
    if not webhook_url:
        return
    secret = env("FEISHU_WEBHOOK_SECRET")
    body = {"msg_type": "text", "content": {"text": message}}
    if secret:
        ts = str(int(time.time()))
        body["timestamp"] = ts
        body["sign"] = _feishu_sign(secret, ts)
    try:
        r = requests.post(webhook_url, json=body, timeout=10, proxies=NO_PROXY)
        d = r.json()
        code = d.get("code") if isinstance(d, dict) else None
        ok = (code == 0) or d.get("StatusCode") == 0 or r.status_code == 200
        print(f"[通知·Webhook] {'OK' if ok else 'FAIL'}  code={code}  status={r.status_code}  "
              f"msg={(d.get('msg') if isinstance(d,dict) else '') or d}")
    except Exception as e:
        print(f"[通知·Webhook] 异常: {e}")


def _id_type(oid: str) -> str:
    """根据 ID 前缀自动判断飞书 id 的类型。
    oc_ → open_id      on_ → union_id
    ou_ → user_id      含@ → email
    纯数字+长度 8~20  → 可能是手机号（不过手机号请直接走 receive_id_type=mobile，这里仅兜底）
    其他 → 回退为 open_id
    """
    if not oid:
        return "open_id"
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


def feishu_send_to_user(tenant_token: str, open_id: str, text: str,
                        *, receive_id_type: Optional[str] = None) -> tuple[bool, str]:
    """用自建应用的 tenant_access_token 给指定 ID 发一条纯文本消息。

    文档：POST /open-apis/im/v1/messages?receive_id_type=xxx
    body: { receive_id, msg_type:"text", content: json.dumps({"text":...}) }

    receive_id_type 不指定时会按 ID 前缀自动猜（oc_→open_id / on_→union_id /
    @→email / ou_→user_id / 数字→mobile），避免了 open_id 跨应用不能复用的坑。
    返回 (成功?, 描述)
    """
    if not tenant_token or not open_id:
        return False, "missing token/open_id"
    id_type = receive_id_type or _id_type(open_id)
    url = f"{FEISHU_OPEN_BASE}/im/v1/messages"
    headers = {"Authorization": f"Bearer {tenant_token}",
               "Content-Type": "application/json; charset=utf-8"}
    params = {"receive_id_type": id_type}
    payload = {
        "receive_id": open_id,
        "msg_type": "text",
        "content": json.dumps({"text": text}, ensure_ascii=False),
    }
    try:
        r = requests.post(url, headers=headers, params=params, json=payload,
                          timeout=15, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") == 0:
            return True, f"type={id_type} message_id={(d.get('data') or {}).get('message_id', '')}"
        return False, f"type={id_type} code={d.get('code')} msg={d.get('msg','')}"
    except Exception as e:
        return False, f"type={id_type} exception={e}"


def stage_notify(message: str, *, tenant_token: str = "",
                 user_open_ids: Optional[list[str]] = None,
                 notify_app_id: str = "", notify_app_secret: str = ""):
    """统一通知入口：先 webhook 群机器人，再向指定用户单发。

    若传入了 notify_app_id / notify_app_secret（即 FEISHU_NOTIFY_APP_ID/SECRET），
    则私信改用该独立应用的 tenant_access_token 来发消息 —— 因为飞书 open_id 是
    每个应用独立签发的，不能跨应用复用。如果未传则回退到 tenant_token。

    每个 ID 的 receive_id_type 会按前缀自动判断（on_→union_id / oc_→open_id 等）。
    """
    stage_notify_webhook(message)
    if not user_open_ids:
        return
    # 选择用于私信的 token
    im_token = tenant_token
    used_app_hint = "Bitable 默认自建应用"
    if notify_app_id and notify_app_secret:
        try:
            im_token = feishu_tenant_access_token(notify_app_id, notify_app_secret)
            used_app_hint = f"通知专用 App {notify_app_id}"
            print(f"[通知·私信] 获取通知应用 token OK (len={len(im_token)})  app={notify_app_id}")
        except Exception as e:
            print(f"[通知·私信] 获取通知应用 token FAIL: {e}；回退使用 Bitable 默认应用 token")
    if not im_token:
        print("[通知·私信] SKIP: 无可用的飞书 token")
        return
    # 预定义 ID → 备注名，用于日志（open_id + union_id 都有）
    NAMES = {
        "oc_5c3d3676b3a2ab87a55a39d37cc52589": "洪",
        "oc_d22e1f9c8cd0a5a3aa2b2625e2a8f155": "王阳",
        "on_93da40c6314edbfa2dc3e031ef405389": "洪",
        "on_b09bcbf3e74f5d423900aa9b2f00eb63": "王阳",
    }
    print(f"[通知·私信] 推送渠道: {used_app_hint}")
    for oid in user_open_ids:
        ok, info = feishu_send_to_user(im_token, oid, message)
        name = NAMES.get(oid, oid)
        print(f"[通知·私信] {'OK  ' if ok else 'FAIL'} → {name}({oid})  {info}")
        time.sleep(0.15)


# ============================================================
# Main
# ============================================================
# 默认要通知的用户（用 Union ID，跨应用通用；放前面，避免再踩到 open_id cross app）
DEFAULT_NOTIFY_OPEN_IDS = [
    "on_93da40c6314edbfa2dc3e031ef405389",   # 洪
    "on_b09bcbf3e74f5d423900aa9b2f00eb63",   # 王阳
]

# 通知专用应用 App ID（不是机密，可硬编码；真实 Secret 走环境变量 FEISHU_NOTIFY_APP_SECRET，
# 不进仓库以避免 GitHub Push Protection 拦截）
DEFAULT_NOTIFY_APP_ID = "cli_aaf0ce1e9ef89d27"


def _parse_notify_users(s: str) -> list[str]:
    """把逗号/空格分隔的 open_id 列表解析成 list（保留非空）。"""
    if not s:
        return []
    parts = re.split(r"[\s,，;；]+", s.strip())
    return [p for p in parts if p]


def parse_args():
    p = argparse.ArgumentParser(
        description="IEC 生产跟踪 → 飞书 tblvugnoJPS8GrpX：资源号命中→进度=结案")
    p.add_argument("--start", type=_parse_month, default="",
                   help="交期起始 YYYYMM（默认前4月）")
    p.add_argument("--end", type=_parse_month, default="",
                   help="交期结束 YYYYMM（默认后1月）")
    p.add_argument("--dry-run", action="store_true",
                   help="预览模式（不实际 UPDATE）")
    p.add_argument("--update-limit", type=int, default=0,
                   help="最大 UPDATE 条数（0=不限）")
    p.add_argument("--iec-page-size", type=int,
                   default=int(env("IEC_PAGE_SIZE") or "100"),
                   help="IEC 分页大小（默认 100）")
    p.add_argument("--bitable-page-size", type=int,
                   default=int(env("BITABLE_PAGE_SIZE") or "500"),
                   help="飞书拉记录分页大小（默认 500）")
    p.add_argument("-v", "--verbose", action="store_true",
                   help="详细日志：逐行打印每个资源号的匹配/更新过程（等价于 VERBOSE=1）")
    p.add_argument("--log-match", default="", metavar="FILE.csv",
                   help="把每阶段的匹配与更新过程按行写入 CSV（支持 Excel 打开的 UTF-8-sig），等价 LOG_MATCH_FILE=xxx.csv")
    p.add_argument("--notify-user", dest="notify_users", action="append",
                   default=None, metavar="ID",
                   help="向指定 ID 单发汇总（可多次指定），支持 on_ (union_id) / oc_ (open_id) / "
                        "ou_ (user_id) / email / 手机号（按前缀自动识别 receive_id_type）。"
                        "若未设置任何 --notify-user 且 NOTIFY_USERS 为空，"
                        "默认推送给内置 2 人：洪 + 王阳（Union ID）")
    p.add_argument("--no-notify-default", action="store_true",
                   help="关闭默认推送名单（等价 NOTIFY_DISABLE_DEFAULT=1），只推 --notify-user 指定的")
    p.add_argument("--notify-app-id", default="",
                   help="用于发私信的自建应用 App ID（与 open_id 对应，默认读 FEISHU_NOTIFY_APP_ID）")
    p.add_argument("--notify-app-secret", default="",
                   help="用于发私信的自建应用 App Secret（默认读 FEISHU_NOTIFY_APP_SECRET）")
    return p.parse_args()


def main():
    args = parse_args()
    dry_run = args.dry_run or env("DRY_RUN") == "1"
    verbose = args.verbose or env("VERBOSE") == "1"
    log_match_file = args.log_match or env("LOG_MATCH_FILE")
    start = args.start or env("EXPORT_START")
    end = args.end or env("EXPORT_END")
    if not start or not end:
        start, end = _default_range()

    app_token = env("BITABLE_APP_TOKEN") or DEFAULT_APP_TOKEN
    table_id = env("BITABLE_TABLE_ID") or DEFAULT_TABLE_ID

    # 通知专用应用（如果没有配置，就用默认 cli_aaf0ce1e9ef89d27；Secret 必须走环境变量）
    notify_app_id = args.notify_app_id or env("FEISHU_NOTIFY_APP_ID") or DEFAULT_NOTIFY_APP_ID
    notify_app_secret = args.notify_app_secret or env("FEISHU_NOTIFY_APP_SECRET") or env("FEISHU_APP_SECRET")

    # 构造 notify list:
    #   1. 先收 NOTIFY_USERS 环境变量（逗号/空格分隔）
    #   2. 再追加所有 --notify-user
    #   3. 若最终仍空 & 未打开 --no-notify-default / NOTIFY_DISABLE_DEFAULT=1 → 追加 DEFAULT 2 人
    disable_default = args.no_notify_default or env("NOTIFY_DISABLE_DEFAULT") == "1"
    notify_ids: list[str] = list(_parse_notify_users(env("NOTIFY_USERS")))
    if args.notify_users:
        for x in args.notify_users:
            notify_ids.extend(_parse_notify_users(x))
    if not notify_ids and not disable_default:
        notify_ids = list(DEFAULT_NOTIFY_OPEN_IDS)
    # 去重（保持原顺序）
    _seen: set[str] = set()
    notify_ids = [o for o in notify_ids if not (o in _seen or _seen.add(o))]

    print("=" * 60)
    print("IEC 结案 → 飞书 tblvugnoJPS8GrpX 进度 UPDATE")
    print(f"  交期: {start}~{end}   {'DRY-RUN' if dry_run else '实际执行'}")
    print(f"  APP_TOKEN={app_token}  TABLE_ID={table_id}")
    print(f"  Bitable App: FEISHU_APP_ID={env('FEISHU_APP_ID') or '(未设置)'}  "
          f"通知 App: NOTIFY_APP_ID={notify_app_id or '-'}")
    print(f"  VERBOSE={'ON' if verbose else 'off'}  "
          f"LOG_MATCH={log_match_file or '-'}")
    if args.update_limit:
        print(f"  UPDATE_LIMIT={args.update_limit}")
    name_lookup = {
        "oc_5c3d3676b3a2ab87a55a39d37cc52589": "洪",
        "oc_d22e1f9c8cd0a5a3aa2b2625e2a8f155": "王阳",
    }
    if notify_ids:
        named = [f"{name_lookup.get(o, '')}<{o}>" for o in notify_ids]
        print(f"  私信通知: {', '.join(named)}")
    else:
        print("  私信通知: OFF (不发任何私信)")
    print("=" * 60)

    # 初始化日志器（如果 verbose 或指定了 csv，则启用）
    log: Optional[MatchLog] = None
    if verbose or log_match_file:
        log = MatchLog(verbose=verbose, csv_path=log_match_file)

    t0 = time.time()
    rc = 0

    try:
        # ---- Stage 1: IEC 拉数据 ----
        print("\n[Stage 1] IEC 生产跟踪 拉取...")
        try:
            all_rows = iec_pull_all(start, end,
                                    page_size=args.iec_page_size,
                                    dry_run=dry_run,
                                    limit=0)
        except Exception as e:
            print(f"[Stage 1] ❌ 失败: {e}")
            import traceback; traceback.print_exc()
            rc = 2
            return rc
        done_set = iec_filter_done(all_rows, log=log)

        # ---- Stage 2: 飞书 Bitable ----
        print("\n[Stage 2] 飞书 Bitable 拉记录 + 更新进度...")
        fs_token = ""
        try:
            fs_token = feishu_tenant_access_token(env("FEISHU_APP_ID"), env("FEISHU_APP_SECRET"))
            print(f"  tenant_access_token OK (len={len(fs_token)})")
        except Exception as e:
            print(f"[Stage 2] ❌ 飞书登录失败: {e}")
            rc = 3
            return rc

        try:
            records = bitable_list_all_records(fs_token, app_token, table_id,
                                               page_size=args.bitable_page_size)
            print(f"  Bitable 记录总数: {len(records)}")
        except Exception as e:
            print(f"[Stage 2] ❌ 拉记录失败: {e}")
            import traceback; traceback.print_exc()
            rc = 4
            return rc

        try:
            planned, ok_count, fail_count, match_stats = bitable_update_jindu(
                fs_token, app_token, table_id, records, done_set,
                dry_run=dry_run, update_limit=args.update_limit,
                log=log, verbose=verbose)
        except Exception as e:
            print(f"[Stage 2] ❌ UPDATE 失败: {e}")
            import traceback; traceback.print_exc()
            rc = 5
            return rc

        # ---- Stage 3: 通知 ----
        elapsed = time.time() - t0
        # 报告格式：去掉"已是结案跳过"汇总行和 dry-run 差异说明，保留耗时时间
        update_line = (f"计划 UPDATE: {planned}  实际成功: {ok_count}  失败: {fail_count}"
                       if not dry_run else f"计划 UPDATE: {planned}（DRY-RUN 未执行）")
        result_str = (
            f"【IEC结案→飞书】{'DRY-RUN ' if dry_run else ''}{time.strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"交期: {start}~{end}\n"
            f"IEC 总记录: {len(all_rows)}  已完成资源号: {len(done_set)}\n"
            f"飞书记录: {len(records)}  匹配命中: {match_stats['intersect']}  未命中: {match_stats['not_in_bitable']}\n"
            f"{update_line}\n"
            f"耗时时间: {elapsed:.1f}s"
            + (f"\n详细匹配日志: {log_match_file}" if log_match_file else "")
        )
        print("\n" + "=" * 60)
        print(result_str)
        print("=" * 60)
        if log is not None:
            log.print_summary()
        stage_notify(result_str, tenant_token=fs_token, user_open_ids=notify_ids or None,
                     notify_app_id=notify_app_id, notify_app_secret=notify_app_secret)
        return 0
    finally:
        if log is not None:
            log.close()


if __name__ == "__main__":
    raise SystemExit(main())
