#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
飞书中间表 → 懂火采购订单入库（路线1）
======================================

业务路径：飞书多维表「入库管理」→ 懂火采购订单 → 采购明细 → 提交审核 → （人工审核）→ 入库

两阶段执行：
  阶段1 (create, 默认)：创建采购订单 + 采购明细 + 提交审核
  阶段2 (ruku)        ：人工审核通过后，按捆包号入库

运行方式（建议从仓库根目录执行）:
  # 预览模式（DRY-RUN）
  python RUKU/import_to_donghuo.py --phase create --dry-run

  # 阶段1：创建订单+明细+提交审核
  python RUKU/import_to_donghuo.py --phase create

  # 阶段2：入库（需人工审核通过后运行）
  python RUKU/import_to_donghuo.py --phase ruku

  # 限制处理条数（调试用）
  python RUKU/import_to_donghuo.py --phase create --limit 5

依赖环境变量:
  DH_USERNAME / DH_PASSWORD           懂火登录
  FEISHU_APP_ID / FEISHU_APP_SECRET   飞书自建应用
  BITABLE_APP_TOKEN                    多维表 app_token（默认 OSuobf2ZkaWtUAsXEE9c3aBTnwh）
  BITABLE_TABLE_ID                     多维表 table_id（默认 tblrhqzHuTsAprU3）
  FEISHU_WEBHOOK_URL / FEISHU_WEBHOOK_SECRET  飞书机器人通知（可选）
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import hmac
import json
import os
import sys
import time
from collections import defaultdict

import requests

# 脚本位于 RUKU/ 子目录，donghuo_login.py 在仓库根目录
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_REPO_ROOT, os.path.dirname(os.path.abspath(__file__))):
    if _p and _p not in sys.path:
        sys.path.insert(0, _p)
from donghuo_login import login_donghuo


# ============================================================
# 常量
# ============================================================
BASE_URL = "https://erpa.donghuo.vip"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}

BITABLE_APP_TOKEN_DEFAULT = "OSuobf2ZkaWtUAsXEE9c3aBTnwh"
BITABLE_TABLE_ID_DEFAULT = "tblrhqzHuTsAprU3"

SHANGHAI_OFFSET_HOURS = 8

DEFAULT_COMPANY = "上海士禾实业有限公司"
DEFAULT_CAIGOUREN = "机器人"
DEFAULT_SHUILV = "0.13"

# 飞书表新增状态字段
STATUS_FIELDS: list[tuple[str, int]] = [
    ("采购订单号", 1),    # TYPE_TEXT=1
    ("采购明细ID", 1),
    ("处理状态", 1),
    ("写入结果", 1),
    ("写入时间", 5),      # TYPE_DATETIME=5
]

# 懂火接口
API_ORDER_LIST = f"{BASE_URL}/model/admin/caigou/m_dindan/getlist"
API_ITEM_LIST = f"{BASE_URL}/model/admin/caigou/m_dindan/mxlist"
API_ITEM_LOAD = f"{BASE_URL}/model/admin/caigou/m_dindan/loadmx"
API_ORDER_ADD = f"{BASE_URL}/controller/admin/caigou/c_cdindan/addc"
API_ITEM_ADD = f"{BASE_URL}/controller/admin/caigou/c_cdindan/addcmx"
API_AUDIT_SUBMIT = f"{BASE_URL}/controller/admin/caigou/c_cdindan/upshenhe"
API_RUKU_LIST = f"{BASE_URL}/model/admin/caigou/m_kucun/caigou_list"
API_RUKU_ADD = f"{BASE_URL}/controller/admin/caigou/c_kucun/addruku"

AJAX_HEADERS = {"X-Requested-With": "XMLHttpRequest"}


# ============================================================
# 工具
# ============================================================
def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def now_shanghai_str() -> str:
    utc_now = datetime.datetime.utcnow()
    sh_now = utc_now + datetime.timedelta(hours=SHANGHAI_OFFSET_HOURS)
    return sh_now.strftime("%Y-%m-%d %H:%M:%S")


def now_shanghai_ts_ms() -> int:
    utc_now = datetime.datetime.utcnow()
    sh_now = utc_now + datetime.timedelta(hours=SHANGHAI_OFFSET_HOURS)
    return int(sh_now.timestamp() * 1000)


def today_str() -> str:
    utc_now = datetime.datetime.utcnow()
    sh_now = utc_now + datetime.timedelta(hours=SHANGHAI_OFFSET_HOURS)
    return sh_now.strftime("%Y-%m-%d")


def _flatten(v) -> str:
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, (int, float, bool)):
        return str(v).strip()
    if isinstance(v, list):
        parts = []
        for seg in v:
            if isinstance(seg, dict):
                if "text" in seg:
                    parts.append(str(seg["text"]).strip())
                elif "name" in seg:
                    parts.append(str(seg["name"]).strip())
            elif isinstance(seg, str):
                parts.append(seg.strip())
        return "".join(p for p in parts if p).strip()
    return str(v).strip()


def _to_float(v, default=0.0) -> float:
    s = str(v).strip() if v is not None else ""
    if not s:
        return default
    try:
        return float(s)
    except (ValueError, TypeError):
        return default


def _fmt_float(v, decimals=3) -> str:
    f = _to_float(v, 0)
    return f"{f:.{decimals}f}"


# 字段别名：飞书表可能用不同命名（手动建表 vs import_lindiao_to_bitable.py 自动建表）
# 取值时按顺序尝试，命中即返回
FIELD_ALIASES: dict[str, list[str]] = {
    "重量": ["重量", "重量(吨)", "净重"],
    "结构": ["结构", "涂层结构"],
    "件数": ["件(张)数", "件数", "张数"],
    "采购税率": ["采购税率", "税率"],
}


def _get_field(row: dict, canonical: str) -> str:
    """按规范字段名取值，兼容别名。返回 stripped 字符串。"""
    names = FIELD_ALIASES.get(canonical, [canonical])
    for n in names:
        v = row.get(n)
        if v is not None and str(v).strip():
            return str(v).strip()
    # fallback：原规范名
    return str(row.get(canonical) or "").strip()


def _get_field_float(row: dict, canonical: str, default=0.0) -> float:
    return _to_float(_get_field(row, canonical), default)


# ============================================================
# 飞书
# ============================================================
def feishu_token() -> str:
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={
        "app_id": env("FEISHU_APP_ID", ""),
        "app_secret": env("FEISHU_APP_SECRET", "")
    }, timeout=15, proxies=NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换取 tenant_access_token 失败: {data}")
    return data["tenant_access_token"]


def bitable_list_fields(tok, app_token, table_id) -> list:
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/fields"
    r = requests.get(url, headers={"Authorization": f"Bearer {tok}"},
                     timeout=15, proxies=NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        print(f"[bitable] list_fields 失败: code={data.get('code')}, msg={data.get('msg')}")
        return []
    return (data.get("data") or {}).get("items") or []


def bitable_create_field(tok, app_token, table_id, name, type_code) -> dict:
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/fields"
    r = requests.post(url, headers={"Authorization": f"Bearer {tok}"},
                      json={"field_name": name, "type": type_code},
                      timeout=15, proxies=NO_PROXY)
    return r.json()


def bitable_ensure_status_fields(tok, app_token, table_id):
    existing = bitable_list_fields(tok, app_token, table_id)
    existing_names = {f.get("field_name", "") for f in existing}
    created = 0
    for name, type_code in STATUS_FIELDS:
        if name in existing_names:
            continue
        print(f"[bitable] 创建状态字段 {name!r}...")
        try:
            bitable_create_field(tok, app_token, table_id, name, type_code)
            created += 1
            time.sleep(0.3)
        except Exception as e:
            print(f"  ⚠️  创建字段 {name!r} 失败: {e}")
    print(f"[bitable] 状态字段就绪：新建 {created} 个")


def bitable_read_all(tok, app_token, table_id) -> list:
    rows = []
    page_token = ""
    for _ in range(200):
        url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(url, params=params,
                         headers={"Authorization": f"Bearer {tok}"},
                         timeout=30, proxies=NO_PROXY)
        d = r.json()
        items = (d.get("data") or {}).get("items") or []
        for it in items:
            rec = {"_record_id": it.get("record_id", "")}
            for k, v in (it.get("fields") or {}).items():
                rec[k] = _flatten(v)
            rows.append(rec)
        d_data = d.get("data") or {}
        if not d_data.get("has_more") or not d_data.get("page_token"):
            break
        page_token = d_data["page_token"]
    return rows


def bitable_update(tok, app_token, table_id, record_id, fields):
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/records/{record_id}")
    r = requests.put(url, json={"fields": fields},
                     headers={"Authorization": f"Bearer {tok}"},
                     timeout=15, proxies=NO_PROXY)
    return r.json()


def feishu_notify(msg: str):
    webhook_url = env("FEISHU_WEBHOOK_URL")
    if not webhook_url:
        return
    secret = env("FEISHU_WEBHOOK_SECRET")
    body = {"msg_type": "text", "content": {"text": msg}}
    headers = {}
    if secret:
        ts = str(int(time.time()))
        sign = _feishu_sign(secret, ts)
        body["timestamp"] = ts
        body["sign"] = sign
    try:
        requests.post(webhook_url, json=body, headers=headers,
                       timeout=10, proxies=NO_PROXY)
    except Exception:
        pass


def _feishu_sign(secret, timestamp):
    import base64 as _b64
    string_to_sign = f"{timestamp}\n{secret}"
    h = hmac.new(secret.encode("utf-8"),
                 string_to_sign.encode("utf-8"),
                 hashlib.sha256)
    return _b64.b64encode(h.digest()).decode("utf-8")


# ============================================================
# 懂火接口
# ============================================================
def donghuo_query_orders(session, contract_no: str) -> list:
    """按采购合同号查询已有采购订单列表"""
    all_rows = []
    page = 1
    while True:
        r = session.post(API_ORDER_LIST,
                         data={"page": page, "limit": 100,
                               "c_hetonhao": contract_no},
                         headers=AJAX_HEADERS, timeout=15)
        try:
            d = r.json()
        except Exception:
            break
        rows = d.get("root") or []
        if not rows:
            break
        all_rows.extend(rows)
        page += 1
        time.sleep(0.2)
    return all_rows


def donghuo_query_items(session, c_danhao: str) -> list:
    """查询某采购订单的明细列表"""
    all_rows = []
    page = 1
    while True:
        r = session.post(API_ITEM_LIST,
                         data={"page": page, "limit": 100,
                               "sc_danhao": c_danhao},
                         headers=AJAX_HEADERS, timeout=15)
        try:
            d = r.json()
        except Exception:
            break
        rows = d.get("root") or []
        if not rows:
            break
        all_rows.extend(rows)
        page += 1
        time.sleep(0.2)
    return all_rows


def donghuo_query_ruku(session, kunbaohao: str) -> list:
    """按捆包号查已入库记录"""
    all_rows = []
    page = 1
    while True:
        r = session.post(API_RUKU_LIST,
                         data={"page": page, "limit": 100,
                               "c_kunbaohao": kunbaohao},
                         headers=AJAX_HEADERS, timeout=15)
        try:
            d = r.json()
        except Exception:
            break
        rows = d.get("root") or []
        if not rows:
            break
        all_rows.extend(rows)
        page += 1
        time.sleep(0.2)
    return all_rows


def donghuo_load_item(session, htmx_id: str, verbose: bool = False) -> dict:
    """加载采购明细详情（loadmx），返回明细字段 dict。

    入库前调用，读取明细级字段（品名/规格/采购单价/税率等）的权威值，
    确保入库表单跟懂火系统里的明细一致，避免飞书表数据被改错导致
    入库数据与明细不一致。

    Parameters
    ----------
    htmx_id : str
        采购明细行 ID（zid）
    verbose : bool
        打印返回内容（调试用）

    Returns
    -------
    dict
        明细字段字典（pinmin/guige/caizhi/c_danjia/c_shuilv 等）。
        失败返回空 dict。
    """
    r = session.post(API_ITEM_LOAD,
                     data={"zid": htmx_id},
                     headers=AJAX_HEADERS, timeout=15)
    try:
        d = r.json()
    except Exception:
        if verbose:
            print(f"  [loadmx] 非 JSON 响应: {r.text[:200]}")
        return {}

    if verbose:
        print(f"  [loadmx] 返回: {str(d)[:300]}")

    # loadmx 直接返回明细字段 dict（无 code/msg 包装）
    # 检测有效数据：有 id 或 pinmin 字段即为成功
    if isinstance(d, dict) and (d.get("id") or d.get("pinmin")):
        return d
    # 兼容 {code:200, root:[{...}]} 结构
    if str(d.get("code", "")) == "200":
        root = d.get("root")
        if isinstance(root, list) and root:
            return root[0] or {}
        data = d.get("data")
        if isinstance(data, dict):
            return data
    return {}


def donghuo_create_order(session, company: str, caigoudanwei: str,
                         c_hetonhao: str, caigouren: str) -> dict:
    """创建采购订单，返回 {code, c_danhao, msg}"""
    data = {
        "company": company,
        "c_time": today_str(),
        "caigoudanwei": caigoudanwei,
        "c_hetonhao": c_hetonhao,
        "caigouren": caigouren,
    }
    r = session.post(API_ORDER_ADD, data=data,
                     headers=AJAX_HEADERS, timeout=15)
    try:
        d = r.json()
        return {"code": str(d.get("code", "")),
                "c_danhao": d.get("c_danhao") or d.get("danhao") or "",
                "msg": d.get("msg", r.text[:200])}
    except Exception:
        return {"code": "-1", "c_danhao": "",
                "msg": r.text[:300]}


def donghuo_create_item(session, c_danhao: str, fields: dict) -> dict:
    """创建采购明细，返回 {code, htmx_id, msg}"""
    data = dict(fields)
    data["c_danhao"] = c_danhao
    r = session.post(API_ITEM_ADD, data=data,
                     headers=AJAX_HEADERS, timeout=15)
    try:
        d = r.json()
        return {"code": str(d.get("code", "")),
                "htmx_id": str(d.get("id") or d.get("htmx_id") or ""),
                "msg": d.get("msg", r.text[:200])}
    except Exception:
        return {"code": "-1", "htmx_id": "",
                "msg": r.text[:300]}


def donghuo_submit_audit(session, c_danhao: str) -> dict:
    """提交审核"""
    r = session.post(API_AUDIT_SUBMIT,
                     data={"c_danhao": c_danhao},
                     headers=AJAX_HEADERS, timeout=15)
    try:
        d = r.json()
        return {"code": str(d.get("code", "")),
                "msg": d.get("msg", r.text[:200])}
    except Exception:
        return {"code": "-1", "msg": r.text[:300]}


def donghuo_ruku(session, c_danhao: str, htmx_id: str,
                 fields: dict) -> dict:
    """入库提交"""
    data = dict(fields)
    data["c_danhao"] = c_danhao
    data["htmx_id"] = htmx_id
    r = session.post(API_RUKU_ADD, data=data,
                     headers=AJAX_HEADERS, timeout=15)
    try:
        d = r.json()
        return {"code": str(d.get("code", "")),
                "msg": d.get("msg", r.text[:200])}
    except Exception:
        return {"code": "-1", "msg": r.text[:300]}


# ============================================================
# 阶段1：创建订单 + 明细 + 提交审核
# ============================================================
def phase_create(session, tok, app_token, table_id,
                 rows: list, dry_run: bool, limit: int) -> dict:
    """处理创建阶段"""
    stats = {"orders_created": 0, "orders_skipped": 0,
             "items_created": 0, "items_skipped": 0,
             "audit_submitted": 0, "failed": 0}

    # 按合同号分组
    by_contract = defaultdict(list)
    for r in rows:
        contract = str(r.get("合同号") or "").strip()
        if contract:
            by_contract[contract].append(r)

    print(f"\n[阶段1] 共 {len(by_contract)} 个合同号，{len(rows)} 条记录")

    contracts = sorted(by_contract.keys())
    if limit > 0:
        contracts = contracts[:limit]
        print(f"  ⚠️  limit={limit}，只处理前 {limit} 个合同号")

    for idx, contract in enumerate(contracts, 1):
        group = by_contract[contract]
        print(f"\n--- [{idx}/{len(contracts)}] 合同号: {contract} "
              f"({len(group)} 条捆包) ---")

        # 1. 查已有订单
        existing_orders = donghuo_query_orders(session, contract)
        c_danhao = ""
        if existing_orders:
            c_danhao = str(existing_orders[0].get("订单号") or "").strip()
            stats["orders_skipped"] += 1
            print(f"  订单已存在: {c_danhao}，跳过创建")
        else:
            supplier = str(group[0].get("供应商") or "").strip()
            if not supplier:
                supplier = "未知供应商"

            if dry_run:
                print(f"  [DRY-RUN] 将创建订单: company={DEFAULT_COMPANY} "
                      f"supplier={supplier} contract={contract}")
                c_danhao = "[DRY-RUN]"
            else:
                result = donghuo_create_order(session,
                                              DEFAULT_COMPANY,
                                              supplier,
                                              contract,
                                              DEFAULT_CAIGOUREN)
                if result["code"] == "200":
                    c_danhao = result.get("c_danhao", "")
                    stats["orders_created"] += 1
                    print(f"  ✅ 订单创建成功: {c_danhao}")
                else:
                    stats["failed"] += 1
                    print(f"  ❌ 订单创建失败: {result['msg']}")
                    # 回写失败状态
                    _write_back(tok, app_token, table_id, group,
                                "失败", f"订单创建失败: {result['msg'][:200]}",
                                dry_run)
                    continue
            time.sleep(0.3)

        # 回写订单号到飞书表
        if c_danhao and c_danhao != "[DRY-RUN]":
            _write_back(tok, app_token, table_id, group,
                        "已创建订单" if not dry_run else "待处理",
                        f"订单号: {c_danhao}",
                        dry_run, c_danhao=c_danhao)

        # 2. 按规格分组创建明细
        by_spec = defaultdict(list)
        for r in group:
            spec_key = (
                str(r.get("品名") or "").strip(),
                str(r.get("规格") or "").strip(),
                str(r.get("材质") or "").strip(),
                str(r.get("产地") or "").strip(),
            )
            by_spec[spec_key].append(r)

        # 查已有明细
        existing_items = []
        if c_danhao and c_danhao != "[DRY-RUN]":
            existing_items = donghuo_query_items(session, c_danhao)

        for spec_key, spec_group in by_spec.items():
            pinmin, guige, caizhi, chandi = spec_key
            spec_label = f"{pinmin}/{guige}/{caizhi}/{chandi}"

            # 查重：已有同规格明细
            existing_item = None
            for it in existing_items:
                if (str(it.get("品名") or "").strip() == pinmin and
                    str(it.get("规格") or "").strip() == guige and
                    str(it.get("材质") or "").strip() == caizhi):
                    existing_item = it
                    break

            if existing_item:
                htmx_id = str(existing_item.get("id") or "").strip()
                stats["items_skipped"] += 1
                print(f"  明细已存在: {spec_label} (id={htmx_id})，跳过")
            else:
                # 计算同规格重量合计
                total_weight = sum(_get_field_float(r, "重量") for r in spec_group)
                item_fields = {
                    "pinmin": pinmin,
                    "guige": guige,
                    "caizhi": caizhi,
                    "chandi": chandi,
                    "denji": str(spec_group[0].get("等级") or "").strip(),
                    "xincen": str(spec_group[0].get("锌层") or "").strip(),
                    "tuliao": str(spec_group[0].get("涂料") or "").strip(),
                    "jiegou": _get_field(spec_group[0], "结构"),
                    "yanse": str(spec_group[0].get("颜色") or "").strip(),
                    "c_zhonlian": _fmt_float(total_weight),
                    "c_danjia": str(spec_group[0].get("采购单价") or "0").strip(),
                    "c_jiner": _fmt_float(total_weight * _to_float(spec_group[0].get("采购单价")), 2),
                    "c_shuilv": _get_field(spec_group[0], "采购税率") or DEFAULT_SHUILV,
                    "c_tidanhao": str(spec_group[0].get("提单号") or "").strip(),
                    "beizhu": str(spec_group[0].get("备注") or "").strip(),
                }

                if dry_run:
                    print(f"  [DRY-RUN] 将创建明细: {spec_label} "
                          f"weight={_fmt_float(total_weight)}")
                    htmx_id = "[DRY-RUN]"
                else:
                    result = donghuo_create_item(session, c_danhao, item_fields)
                    if result["code"] == "200":
                        htmx_id = result.get("htmx_id", "")
                        stats["items_created"] += 1
                        print(f"  ✅ 明细创建成功: {spec_label} (id={htmx_id})")
                    else:
                        stats["failed"] += 1
                        print(f"  ❌ 明细创建失败: {result['msg']}")
                        continue
                time.sleep(0.3)

            # 回写明细ID到飞书表
            if htmx_id and htmx_id != "[DRY-RUN]":
                _write_back(tok, app_token, table_id, spec_group,
                            "已创建订单", f"明细ID: {htmx_id}",
                            dry_run, c_danhao=c_danhao, htmx_id=htmx_id)

        # 3. 提交审核
        if c_danhao and c_danhao != "[DRY-RUN]":
            if dry_run:
                print(f"  [DRY-RUN] 将提交审核: {c_danhao}")
            else:
                result = donghuo_submit_audit(session, c_danhao)
                if result["code"] == "200":
                    stats["audit_submitted"] += 1
                    print(f"  ✅ 提交审核成功: {c_danhao}")
                    _write_back(tok, app_token, table_id, group,
                                "已提交审核", "等待人工审核通过",
                                dry_run, c_danhao=c_danhao)
                else:
                    stats["failed"] += 1
                    print(f"  ❌ 提交审核失败: {result['msg']}")
                    _write_back(tok, app_token, table_id, group,
                                "失败", f"审核提交失败: {result['msg'][:200]}",
                                dry_run, c_danhao=c_danhao)
                time.sleep(0.3)

    return stats


# ============================================================
# 阶段2：入库
# ============================================================
def phase_ruku(session, tok, app_token, table_id,
               rows: list, dry_run: bool, limit: int) -> dict:
    """处理入库阶段"""
    stats = {"ruku_done": 0, "ruku_skipped": 0, "failed": 0}

    # 只处理有采购明细ID且状态为"已提交审核"的记录
    pending = [r for r in rows
               if str(r.get("采购明细ID") or "").strip()
               and str(r.get("处理状态") or "").strip() == "已提交审核"]

    print(f"\n[阶段2] 共 {len(pending)} 条待入库记录")
    if limit > 0:
        pending = pending[:limit]
        print(f"  ⚠️  limit={limit}，只处理前 {limit} 条")

    for idx, r in enumerate(pending, 1):
        kb = str(r.get("捆包号") or "").strip()
        c_danhao = str(r.get("采购订单号") or "").strip()
        htmx_id = str(r.get("采购明细ID") or "").strip()
        print(f"\n--- [{idx}/{len(pending)}] 捆包号: {kb} 订单: {c_danhao} ---")

        # 1. 查已入库
        existing = donghuo_query_ruku(session, kb)
        if existing:
            stats["ruku_skipped"] += 1
            print(f"  已入库，跳过")
            _write_back_single(tok, app_token, table_id, r,
                               "已入库", "已入库（跳过）", dry_run)
            continue

        # 2. 加载采购明细（loadmx）—— 明细级字段权威值
        #    入库表单的明细级字段（品名/规格/采购单价/税率等）从 loadmx 读，
        #    确保跟懂火系统里的明细一致；loadmx 为空的字段用飞书表兜底
        item = donghuo_load_item(session, htmx_id, verbose=dry_run)
        if not item:
            if dry_run:
                print(f"  [DRY-RUN] loadmx 返回空，将用飞书表字段兜底")
            else:
                stats["failed"] += 1
                print(f"  ❌ 加载明细失败（loadmx 返回空），跳过入库")
                _write_back_single(tok, app_token, table_id, r,
                                    "失败",
                                    f"loadmx 加载明细失败 htmx_id={htmx_id}",
                                    dry_run)
                continue
        else:
            print(f"  ✅ 明细加载: {_flatten(item.get('pinmin'))}"
                  f"/{_flatten(item.get('guige'))}")

        # 3. 构造入库表单
        #    明细级字段（品名/规格/材质/产地/等级/锌层/涂料/结构/颜色/
        #                采购单价/采购税率/提单号/合同号）→ 从 loadmx 读，权威值
        #    入库级字段（件数/重量/仓库/库位号/捆包号/车船号/米数/备注/
        #               销售单价/upid）→ 从飞书表读
        def _item_str(key):
            """从 loadmx 返回取明细字段值，stripped。空则返回空串。"""
            return _flatten(item.get(key) or "")

        ruku_fields = {
            # 明细级字段（来自 loadmx，权威值；loadmx 为空用飞书表兜底）
            "pinmin": _item_str("pinmin") or str(r.get("品名") or "").strip(),
            "guige": _item_str("guige") or str(r.get("规格") or "").strip(),
            "caizhi": _item_str("caizhi") or str(r.get("材质") or "").strip(),
            "chandi": _item_str("chandi") or str(r.get("产地") or "").strip(),
            "denji": _item_str("denji") or str(r.get("等级") or "").strip(),
            "xincen": _item_str("xincen") or str(r.get("锌层") or "").strip(),
            "tuliao": _item_str("tuliao") or str(r.get("涂料") or "").strip(),
            "jiegou": _item_str("jiegou") or _get_field(r, "结构"),
            "yanse": _item_str("yanse") or str(r.get("颜色") or "").strip(),
            "c_danjia": _item_str("c_danjia") or str(r.get("采购单价") or "0").strip(),
            "c_shuilv": _item_str("c_shuilv") or _get_field(r, "采购税率") or DEFAULT_SHUILV,
            "c_tidanhao": _item_str("c_tidanhao") or str(r.get("提单号") or "").strip(),
            "c_hetonhao": _item_str("c_hetonhao") or str(r.get("合同号") or "").strip(),
            # 入库级字段（来自飞书表）
            "c_jianshu": _get_field(r, "件数") or "1",
            "c_zhonlian": _fmt_float(_get_field_float(r, "重量")),
            "x_danjia": "0.00",
            "canku": str(r.get("仓库") or "").strip(),
            "kuweihao": str(r.get("库位号") or "").strip(),
            "c_kunbaohao": kb,
            "c_cchuanhao": str(r.get("车船号") or "").strip(),
            "c_guihao": str(r.get("米数") or "").strip(),
            "beizhu": str(r.get("备注") or "").strip(),
            "upid": htmx_id,
        }

        if dry_run:
            print(f"  [DRY-RUN] 将入库: kb={kb} weight={ruku_fields['c_zhonlian']}")
            stats["ruku_skipped"] += 1
        else:
            result = donghuo_ruku(session, c_danhao, htmx_id, ruku_fields)
            if result["code"] == "200":
                stats["ruku_done"] += 1
                print(f"  ✅ 入库成功")
                _write_back_single(tok, app_token, table_id, r,
                                    "已入库", f"入库成功 {now_shanghai_str()}",
                                    dry_run)
            else:
                stats["failed"] += 1
                print(f"  ❌ 入库失败: {result['msg']}")
                _write_back_single(tok, app_token, table_id, r,
                                    "失败", f"入库失败: {result['msg'][:200]}",
                                    dry_run)
        time.sleep(0.5)

    return stats


# ============================================================
# 飞书回写
# ============================================================
def _write_back(tok, app_token, table_id, rows: list,
                status: str, result: str, dry_run: bool,
                c_danhao: str = "", htmx_id: str = ""):
    """批量回写飞书表状态"""
    if dry_run:
        return
    now_ms = now_shanghai_ts_ms()
    for r in rows:
        rid = r.get("_record_id", "")
        if not rid:
            continue
        fields = {
            "处理状态": status,
            "写入结果": result[:500],
            "写入时间": now_ms,
        }
        if c_danhao:
            fields["采购订单号"] = c_danhao
        if htmx_id:
            fields["采购明细ID"] = htmx_id
        try:
            bitable_update(tok, app_token, table_id, rid, fields)
        except Exception as e:
            print(f"  ⚠️  回写飞书失败 {rid}: {e}")


def _write_back_single(tok, app_token, table_id, row: dict,
                        status: str, result: str, dry_run: bool):
    """单条回写"""
    if dry_run:
        return
    rid = row.get("_record_id", "")
    if not rid:
        return
    fields = {
        "处理状态": status,
        "写入结果": result[:500],
        "写入时间": now_shanghai_ts_ms(),
    }
    try:
        bitable_update(tok, app_token, table_id, rid, fields)
    except Exception as e:
        print(f"  ⚠️  回写飞书失败 {rid}: {e}")


# ============================================================
# 主流程
# ============================================================
def parse_args():
    p = argparse.ArgumentParser(
        description="飞书中间表 → 懂火采购订单入库",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python RUKU/import_to_donghuo.py --phase create --dry-run
  python RUKU/import_to_donghuo.py --phase create --limit 5
  python RUKU/import_to_donghuo.py --phase ruku --limit 5
""")
    p.add_argument("--phase", choices=["create", "ruku", "all"],
                   default="create", help="执行阶段")
    p.add_argument("--dry-run", action="store_true", help="预览模式")
    p.add_argument("--limit", type=int, default=0,
                   help="限制处理条数（调试用）")
    return p.parse_args()


def main():
    args = parse_args()
    dry_run = args.dry_run or env("DRY_RUN") == "1"
    limit = args.limit or int(env("IMPORT_LIMIT", "0") or "0")
    phase = args.phase

    app_token = env("BITABLE_APP_TOKEN", BITABLE_APP_TOKEN_DEFAULT)
    table_id = env("BITABLE_TABLE_ID", BITABLE_TABLE_ID_DEFAULT)

    print(f"====== 飞书中间表 → 懂火采购订单入库 ======")
    print(f"  阶段: {phase}")
    print(f"  模式: {'DRY-RUN' if dry_run else '实际执行'}")
    print(f"  limit: {limit if limit else '全部'}")
    print(f"  飞书表: {app_token} / {table_id}")

    t0 = time.time()

    # 1. 飞书 token
    print("\n[步骤] 获取飞书 token...")
    tok = feishu_token()
    print("  ✅ 飞书 token 获取成功")

    # 2. 确保状态字段
    if phase in ("create", "all"):
        print("[步骤] 确保飞书表状态字段...")
        bitable_ensure_status_fields(tok, app_token, table_id)

    # 3. 读取飞书表
    print("[步骤] 读取飞书中间表...")
    all_rows = bitable_read_all(tok, app_token, table_id)
    print(f"  共 {len(all_rows)} 条记录")

    if phase in ("create", "all"):
        # 待处理：处理状态为空 或 "待处理"
        pending = [r for r in all_rows
                   if not str(r.get("处理状态") or "").strip()
                   or str(r.get("处理状态") or "").strip() == "待处理"]
        print(f"  待创建订单: {len(pending)} 条")

        # 4. 登录懂火
        print("\n[步骤] 登录懂火...")
        session = login_donghuo()
        if session is None:
            print("登录失败")
            return 1

        # 5. 执行创建阶段
        stats_create = phase_create(session, tok, app_token, table_id,
                                    pending, dry_run, limit)
        session.close()

    if phase in ("ruku", "all"):
        # 入库阶段：处理状态="已提交审核" 且有明细ID
        pending_ruku = [r for r in all_rows
                        if str(r.get("处理状态") or "").strip() == "已提交审核"
                        and str(r.get("采购明细ID") or "").strip()]
        print(f"\n  待入库: {len(pending_ruku)} 条")

        if pending_ruku:
            if phase == "all":
                print("\n⚠️  phase=all 模式下入库阶段需要人工审核通过，"
                      "如果尚未审核，入库会失败。")
                print("  建议先运行 --phase create，审核通过后再运行 --phase ruku")
                if not dry_run:
                    print("  跳过入库阶段")
                    stats_ruku = {"ruku_done": 0, "ruku_skipped": 0, "failed": 0}
                else:
                    session = login_donghuo()
                    if session is None:
                        print("登录失败")
                        return 1
                    stats_ruku = phase_ruku(session, tok, app_token, table_id,
                                             pending_ruku, dry_run, limit)
                    session.close()
            else:
                print("\n[步骤] 登录懂火...")
                session = login_donghuo()
                if session is None:
                    print("登录失败")
                    return 1
                stats_ruku = phase_ruku(session, tok, app_token, table_id,
                                         pending_ruku, dry_run, limit)
                session.close()
        else:
            stats_ruku = {"ruku_done": 0, "ruku_skipped": 0, "failed": 0}

    # 6. 汇总
    elapsed = time.time() - t0
    print(f"\n{'='*60}")
    print(f"📊 汇总（耗时 {elapsed:.1f}s）{' DRY-RUN' if dry_run else ''}")

    if phase in ("create", "all"):
        s = stats_create
        print(f"  [创建阶段]")
        print(f"    订单: 新建 {s['orders_created']}, 跳过 {s['orders_skipped']}")
        print(f"    明细: 新建 {s['items_created']}, 跳过 {s['items_skipped']}")
        print(f"    审核: 提交 {s['audit_submitted']}, 失败 {s['failed']}")

    if phase in ("ruku", "all"):
        s = stats_ruku
        print(f"  [入库阶段]")
        print(f"    入库: 成功 {s['ruku_done']}, 跳过 {s['ruku_skipped']}, "
              f"失败 {s['failed']}")

    # 7. 飞书通知
    msg_lines = [
        f"【懂火采购订单入库】{'DRY-RUN ' if dry_run else ''}{now_shanghai_str()}",
        f"阶段: {phase}  耗时: {elapsed:.1f}s",
    ]
    if phase in ("create", "all"):
        s = stats_create
        msg_lines.append(
            f"创建: 订单 {s['orders_created']}新建/{s['orders_skipped']}跳过, "
            f"明细 {s['items_created']}新建/{s['items_skipped']}跳过, "
            f"审核 {s['audit_submitted']}, 失败 {s['failed']}"
        )
    if phase in ("ruku", "all"):
        s = stats_ruku
        msg_lines.append(
            f"入库: 成功 {s['ruku_done']}, 跳过 {s['ruku_skipped']}, 失败 {s['failed']}"
        )
    feishu_notify("\n".join(msg_lines))

    print(f"{'='*60}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
