"""
宝钢 IEC 准发数据 → 飞书多维表格"进度"表更新
================================================

功能：
  1. 读取 download_bundle.py 导出的 xlsx（含"钢厂订单号""准发量""出厂量"）
  2. 在"合约"表中按"资源号"=钢厂订单号 批量查询 record_id
     （进度表的"资源号"是 Lookup 不能直接 filter；但合约表和进度表的 record_id 一一对应）
  3. 用 batch_update 批量更新进度表对应 record 的"准发量"和"出厂量"

写入规则：
  - 目标表："进度"
  - 唯一号：xlsx"钢厂订单号" ↔ 合约表"资源号" → 同 record_id 定位到进度表
  - 写入字段：进度表的"准发量"、"出厂量"

环境变量（必填）：
  FEISHU_APP_ID        飞书自建应用 App ID（cli_ 开头）
  FEISHU_APP_SECRET    飞书自建应用 App Secret

环境变量（可选）：
  BITABLE_APP_TOKEN    多维表格 app_token（默认 Tz0XbQVzkaZuJasBwb8cRjkfnoe）

用法：
  # 1) 配置飞书凭据（PowerShell）
  $env:FEISHU_APP_ID = "cli_xxxxxxxxxxxxxx"
  $env:FEISHU_APP_SECRET = "xxxxxxxxxxxxxxxxxxxxxxxx"

  # 2) 跑脚本（默认找最新的 准发下载_*.xlsx，自动更新"进度"表）
  python update_progress_delivery.py

  # 3) 指定 xlsx 文件
  python update_progress_delivery.py --xlsx 准发下载_260816_1044.xlsx

  # 4) 先干跑预览，不实际更新
  python update_progress_delivery.py --dry-run

前置条件：
  - 飞书自建应用需开通 bitable:app 权限
  - 应用需被添加为目标多维表格的"可编辑"协作者
  - 目标字段（"准发量"、"出厂量"、合约表"资源号"）需在字段权限管理中对应用授权
"""
from __future__ import annotations

import argparse
import glob
import os
import sys
import time
from typing import Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import pandas as pd
    _HAS_PANDAS = True
except ImportError:
    _HAS_PANDAS = False


# ============================================================
# 常量
# ============================================================
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"

# 从分享 URL https://s2v31ke6sl.feishu.cn/base/Tz0XbQVzkaZuJasBwb8cRjkfnoe 提取
DEFAULT_APP_TOKEN = "Tz0XbQVzkaZuJasBwb8cRjkfnoe"
DEFAULT_TABLE_PROGRESS = "进度"          # 写入目标表
DEFAULT_TABLE_CONTRACT = "合约"          # 查 record_id 的源表（资源号是 Text，可 filter）
DEFAULT_FIELD_RESOURCE = "资源号"        # 合约表里用来匹配的字段（Text 类型）
DEFAULT_FIELD_QUASI = "准发量"           # 进度表写入字段
DEFAULT_FIELD_SHIP = "出厂量"            # 进度表写入字段

# xlsx 中的列名
XLSX_COL_ORDER = "钢厂订单号"
XLSX_COL_QUASI = "准发量"
XLSX_COL_SHIP = "出厂量"

# 默认查找最新的"准发下载_*.xlsx"
DEFAULT_XLSX_GLOB = "准发下载_*.xlsx"

_NO_PROXY = {"http": None, "https": None}


# ============================================================
# 通用
# ============================================================
def env(name: str, default: str = "") -> str:
    return (os.environ.get(name) or "").strip() or default


def find_latest_xlsx(pattern: str = DEFAULT_XLSX_GLOB) -> str:
    """在当前目录找最新的匹配文件（按修改时间）"""
    files = glob.glob(pattern)
    if not files:
        return ""
    files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return files[0]


# ============================================================
# 飞书认证
# ============================================================
def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    """用自建应用 App ID/Secret 换 tenant_access_token"""
    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET 环境变量")
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": app_id, "app_secret": app_secret},
                      timeout=15, proxies=_NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换取 tenant_access_token 失败: {data}")
    token = str(data.get("tenant_access_token") or "")
    if not token:
        raise RuntimeError(f"响应中无 tenant_access_token: {data}")
    print(f"[飞书] 获取 tenant_access_token 成功 (len={len(token)})")
    return token


# ============================================================
# 飞书多维表格：列出表 & 字段
# ============================================================
def list_tables(token: str, app_token: str) -> list[dict]:
    """GET /bitable/v1/apps/{app_token}/tables，返回 [{table_id, name, ...}]"""
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                     timeout=15, proxies=_NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"列出多维表格失败: {data}")
    items = (data.get("data") or {}).get("items") or []
    return items


def find_table_id(token: str, app_token: str, table_name: str) -> str:
    """根据表名查找 table_id"""
    tables = list_tables(token, app_token)
    print(f"[飞书] 多维表格共 {len(tables)} 张表:")
    for t in tables:
        tid = t.get("table_id", "")
        nm = t.get("name", "")
        print(f"  - {tid}: {nm}")
        if nm == table_name:
            return tid
    # 模糊匹配
    for t in tables:
        if table_name in (t.get("name") or ""):
            tid = t.get("table_id", "")
            print(f"  [模糊匹配] 用 {tid}: {t.get('name')}")
            return tid
    raise RuntimeError(f"未找到名为 '{table_name}' 的表，请检查表名")


def list_fields(token: str, app_token: str, table_id: str) -> list[dict]:
    """GET /bitable/v1/apps/{app_token}/tables/{table_id}/fields"""
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/fields")
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                     timeout=15, proxies=_NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"列出字段失败: {data}")
    return (data.get("data") or {}).get("items") or []


# ============================================================
# 飞书多维表格：按"资源号"批量查询记录
# ============================================================
def extract_text_field(val) -> str:
    """从飞书字段值中提取纯文本"""
    if val is None:
        return ""
    if isinstance(val, str):
        return val.strip()
    if isinstance(val, list):
        for seg in val:
            if isinstance(seg, dict):
                t = str(seg.get("text") or seg.get("name") or "").strip()
                if t:
                    return t
        return ""
    return str(val).strip()


def search_records_by_resource(
    token: str, app_token: str, table_id: str,
    resource_field: str, resource_values: list[str],
    link_field: str = "编号",
    page_size: int = 500,
    batch_size: int = 50,
) -> dict[str, list[str]]:
    """在合约表按"资源号"查记录，返回 {资源号值: [编号值, ...]}。

    飞书 filter conditions 单次最多 50 个，超过自动分批查询。
    返回的 dict 的 value 是"编号"字段的文本值（用于关联进度表）。
    """
    if not resource_values:
        return {}

    unique_vals = list({v.strip() for v in resource_values if v and v.strip()})
    if not unique_vals:
        return {}

    result: dict[str, list[str]] = {}
    total_batches = (len(unique_vals) + batch_size - 1) // batch_size
    print(f"[飞书] 合约表分批查询：{len(unique_vals)} 个资源号，每批 {batch_size}，共 {total_batches} 批")

    for bi in range(0, len(unique_vals), batch_size):
        batch_vals = unique_vals[bi:bi + batch_size]
        conditions = [
            {"field_name": resource_field, "operator": "is", "value": [v]}
            for v in batch_vals
        ]
        page_token = ""
        seen_tokens: set[str] = set()
        page_count = 0
        batch_no = bi // batch_size + 1

        while True:
            page_count += 1
            if page_count > 20:
                print(f"[警告] 批 {batch_no} 已达最大翻页 20，停止")
                break
            if page_token and page_token in seen_tokens:
                break
            if page_token:
                seen_tokens.add(page_token)

            url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
                   f"/tables/{table_id}/records/search")
            body: dict = {
                "page_size": page_size,
                "filter": {
                    "conjunction": "or",
                    "conditions": conditions,
                },
            }
            if page_token:
                body["page_token"] = page_token

            r = requests.post(url, headers={"Authorization": f"Bearer {token}"},
                              json=body, timeout=30, proxies=_NO_PROXY)
            data = r.json()
            if data.get("code") != 0:
                raise RuntimeError(f"按资源号查询失败 (批 {batch_no}): {data}")
            d = data.get("data") or {}
            items = d.get("items") or []
            for it in items:
                fields = it.get("fields") or {}
                res_val = extract_text_field(fields.get(resource_field))
                link_val = extract_text_field(fields.get(link_field))
                if res_val and link_val:
                    result.setdefault(res_val, []).append(link_val)

            has_more = d.get("has_more", False)
            page_token = d.get("page_token") or ""
            if not has_more or not page_token:
                break
            time.sleep(0.2)
        time.sleep(0.15)

    print(f"[飞书] 合约表查询到 {len(result)} 个唯一资源号有匹配记录 "
          f"(待查 {len(unique_vals)} 个)")
    return result


def search_progress_by_codes(
    token: str, app_token: str, table_id: str,
    code_field: str, code_values: list[str],
    page_size: int = 500,
    batch_size: int = 50,
) -> dict[str, list[str]]:
    """在进度表按"编号"查记录，返回 {编号值: [record_id, ...]}。

    飞书 filter conditions 单次最多 50 个，超过自动分批查询。
    """
    if not code_values:
        return {}

    unique_vals = list({v.strip() for v in code_values if v and v.strip()})
    if not unique_vals:
        return {}

    result: dict[str, list[str]] = {}
    total_batches = (len(unique_vals) + batch_size - 1) // batch_size
    print(f"[飞书] 进度表分批查询：{len(unique_vals)} 个编号，每批 {batch_size}，共 {total_batches} 批")

    for bi in range(0, len(unique_vals), batch_size):
        batch_vals = unique_vals[bi:bi + batch_size]
        conditions = [
            {"field_name": code_field, "operator": "is", "value": [v]}
            for v in batch_vals
        ]
        page_token = ""
        seen_tokens: set[str] = set()
        page_count = 0
        batch_no = bi // batch_size + 1

        while True:
            page_count += 1
            if page_count > 20:
                print(f"[警告] 批 {batch_no} 已达最大翻页 20，停止")
                break
            if page_token and page_token in seen_tokens:
                break
            if page_token:
                seen_tokens.add(page_token)

            url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
                   f"/tables/{table_id}/records/search")
            body: dict = {
                "page_size": page_size,
                "filter": {
                    "conjunction": "or",
                    "conditions": conditions,
                },
            }
            if page_token:
                body["page_token"] = page_token

            r = requests.post(url, headers={"Authorization": f"Bearer {token}"},
                              json=body, timeout=30, proxies=_NO_PROXY)
            data = r.json()
            if data.get("code") != 0:
                raise RuntimeError(f"进度表按编号查询失败 (批 {batch_no}): {data}")
            d = data.get("data") or {}
            items = d.get("items") or []
            for it in items:
                rid = it.get("record_id", "")
                fields = it.get("fields") or {}
                code_val = extract_text_field(fields.get(code_field))
                if code_val and rid:
                    result.setdefault(code_val, []).append(rid)

            has_more = d.get("has_more", False)
            page_token = d.get("page_token") or ""
            if not has_more or not page_token:
                break
            time.sleep(0.2)
        time.sleep(0.15)

    print(f"[飞书] 进度表查询到 {len(result)} 个唯一编号有匹配记录 "
          f"(待查 {len(unique_vals)} 个)")
    return result


# ============================================================
# 飞书多维表格：批量更新
# ============================================================
def batch_update_records(
    token: str, app_token: str, table_id: str,
    records: list[dict],
) -> int:
    """批量更新记录（records=[{"record_id":"xxx","fields":{...}}]，单批≤500）"""
    if not records:
        return 0
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/records/batch_update")
    updated = 0
    for i in range(0, len(records), 500):
        batch = records[i:i + 500]
        r = requests.post(url, headers={"Authorization": f"Bearer {token}"},
                          json={"records": batch}, timeout=30,
                          proxies=_NO_PROXY)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"批量更新失败 (批次 {i // 500 + 1}): {data}")
        n = len((data.get("data") or {}).get("records") or [])
        updated += n
        print(f"  批次 {i // 500 + 1}: 成功更新 {n} 条")
        if i + 500 < len(records):
            time.sleep(0.4)
    return updated


def _list_all_record_ids(
    token: str, app_token: str, table_id: str,
    page_size: int = 500,
) -> set[str]:
    """遍历整张表，返回所有 record_id 的集合"""
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/records")
    rids: set[str] = set()
    page_token = ""
    page = 0
    while True:
        page += 1
        params = {"page_size": page_size}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(url, params=params,
                         headers={"Authorization": f"Bearer {token}"},
                         timeout=30, proxies=_NO_PROXY)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"拉取 record_id 失败: {data}")
        d = data.get("data") or {}
        for it in d.get("items") or []:
            rid = it.get("record_id")
            if rid:
                rids.add(rid)
        if not d.get("has_more"):
            break
        page_token = d.get("page_token") or ""
        if not page_token:
            break
        time.sleep(0.2)
    return rids


# ============================================================
# 读取 xlsx
# ============================================================
def read_xlsx(xlsx_path: str) -> dict[str, dict[str, float]]:
    """读取 xlsx，按"钢厂订单号"提取"准发量"和"出厂量"。

    Returns
    -------
    dict
        {钢厂订单号: {"准发量": x, "出厂量": y}}
    """
    if not _HAS_PANDAS:
        raise ImportError("需要 pandas 和 openpyxl：pip install pandas openpyxl")
    df = pd.read_excel(xlsx_path)
    for col in (XLSX_COL_ORDER, XLSX_COL_QUASI, XLSX_COL_SHIP):
        if col not in df.columns:
            raise RuntimeError(f"xlsx 缺少 '{col}' 列，现有列: {list(df.columns)}")

    df[XLSX_COL_QUASI] = pd.to_numeric(df[XLSX_COL_QUASI], errors="coerce").fillna(0)
    df[XLSX_COL_SHIP] = pd.to_numeric(df[XLSX_COL_SHIP], errors="coerce").fillna(0)
    # 同一钢厂订单号可能有多行（订单子项号不同），按订单号汇总求和
    grouped = df.groupby(XLSX_COL_ORDER).agg(
        quasi=(XLSX_COL_QUASI, "sum"),
        ship=(XLSX_COL_SHIP, "sum"),
    )
    result: dict[str, dict[str, float]] = {}
    for k, row in grouped.iterrows():
        k_str = str(k).strip()
        if k_str:
            result[k_str] = {
                XLSX_COL_QUASI: round(float(row["quasi"]), 3),
                XLSX_COL_SHIP: round(float(row["ship"]), 3),
            }
    print(f"[xlsx] {xlsx_path}: {len(df)} 行 → 汇总为 {len(result)} 个唯一钢厂订单号")
    return result


# ============================================================
# 主流程
# ============================================================
def update_progress(
    xlsx_path: str,
    app_token: Optional[str] = None,
    progress_table_name: str = DEFAULT_TABLE_PROGRESS,
    contract_table_name: str = DEFAULT_TABLE_CONTRACT,
    resource_field: str = DEFAULT_FIELD_RESOURCE,
    quasi_field: str = DEFAULT_FIELD_QUASI,
    ship_field: str = DEFAULT_FIELD_SHIP,
    dry_run: bool = False,
) -> dict:
    """主流程：xlsx → 汇总 → 查合约表拿 record_id → 批量更新进度表

    Returns
    -------
    dict
        {"xlsx_orders": N, "matched": M, "updated": K, "missing": [...]}
    """
    app_token = app_token or env("BITABLE_APP_TOKEN", DEFAULT_APP_TOKEN)

    # 1. 读取 xlsx
    print(f"\n=== 步骤 1/5：读取 xlsx 并按钢厂订单号汇总 ===")
    order_data = read_xlsx(xlsx_path)
    if not order_data:
        print("[警告] xlsx 无有效数据，结束")
        return {"xlsx_orders": 0, "matched": 0, "updated": 0, "missing": []}
    sample = list(order_data.items())[:5]
    print("  预览（前 5 个）:")
    for k, v in sample:
        print(f"    {k} → 准发量 {v[XLSX_COL_QUASI]}, 出厂量 {v[XLSX_COL_SHIP]}")

    # 2. 飞书认证 + 找两张表
    print(f"\n=== 步骤 2/5：飞书认证 + 找 '{contract_table_name}' 和 '{progress_table_name}' 表 ===")
    app_id = env("FEISHU_APP_ID")
    app_secret = env("FEISHU_APP_SECRET")
    token = feishu_tenant_access_token(app_id, app_secret)

    contract_id = find_table_id(token, app_token, contract_table_name)
    progress_id = find_table_id(token, app_token, progress_table_name)
    print(f"[飞书] 合约表: {contract_id}（查 record_id）")
    print(f"[飞书] 进度表: {progress_id}（写 准发量/出厂量）")

    # 校验合约表的"资源号"字段
    contract_fields = list_fields(token, app_token, contract_id)
    cfield_meta = next((f for f in contract_fields if f.get("field_name") == resource_field), None)
    if not cfield_meta:
        cnames = [f.get("field_name", "") for f in contract_fields]
        raise RuntimeError(f"合约表无 '{resource_field}' 字段，现有字段: {cnames}")
    print(f"[飞书] 合约表 '{resource_field}' 字段类型: {cfield_meta.get('type')} ({cfield_meta.get('ui_type')})")

    # 校验进度表的"准发量""出厂量"字段
    progress_fields = list_fields(token, app_token, progress_id)
    for fn in (quasi_field, ship_field):
        meta = next((f for f in progress_fields if f.get("field_name") == fn), None)
        if not meta:
            pnames = [f.get("field_name", "") for f in progress_fields]
            raise RuntimeError(
                f"进度表无 '{fn}' 字段（可能未授权字段权限），现有字段名: {pnames}"
            )
        print(f"[飞书] 进度表 '{fn}' 字段类型: {meta.get('type')} ({meta.get('ui_type')})")

    # 3. 在合约表按"资源号"批量查询 → 拿到"编号"字段值
    #    （进度表"资源号"是 Lookup，通过"编号"字段关联合约表）
    print(f"\n=== 步骤 3/5：在合约表按 '{resource_field}' 查询，读取'编号' ===")
    order_list = list(order_data.keys())
    matched = search_records_by_resource(
        token, app_token, contract_id,
        resource_field, order_list,
        link_field="编号",
    )

    xlsx_set = set(order_data.keys())
    matched_set = set(matched.keys())
    missing = xlsx_set - matched_set
    print(f"[飞书] 合约表匹配结果:")
    print(f"  xlsx 中订单数: {len(xlsx_set)}")
    print(f"  合约表匹配数:  {len(matched_set)}")
    print(f"  未匹配数:       {len(missing)}")
    if missing:
        print(f"  未匹配订单号示例: {list(missing)[:10]}")

    # 4. 收集所有"编号"值，在进度表按"编号"查 record_id
    all_codes: list[str] = []
    for codes in matched.values():
        all_codes.extend(codes)
    print(f"\n=== 步骤 4/5：在进度表按'编号'查询 record_id ===")
    code_to_rids = search_progress_by_codes(
        token, app_token, progress_id,
        "编号", all_codes,
    )

    # 5. 组合：钢厂订单号 → 编号 → 进度表 record_id → 更新
    print(f"\n=== 步骤 5/5：在进度表批量更新 '{quasi_field}' 和 '{ship_field}' ===")
    update_records: list[dict] = []
    skipped_no_progress = 0
    multi_count = 0
    for order_no, vals in order_data.items():
        codes = matched.get(order_no)
        if not codes:
            continue
        found_any = False
        for code in codes:
            rids = code_to_rids.get(code, [])
            for rid in rids:
                update_records.append({
                    "record_id": rid,
                    "fields": {
                        quasi_field: vals[XLSX_COL_QUASI],
                        ship_field: vals[XLSX_COL_SHIP],
                    },
                })
                found_any = True
            if len(rids) > 1:
                multi_count += 1
        if not found_any:
            skipped_no_progress += 1

    if skipped_no_progress:
        print(f"  [提示] {skipped_no_progress} 个订单在合约表有但进度表无对应记录，跳过")
    if multi_count:
        print(f"  [提示] 有 {multi_count} 个编号在进度表中有多条匹配记录")

    print(f"  实际待更新记录数: {len(update_records)}")

    if dry_run:
        print("\n[DRY-RUN] 仅预览，不实际写入。前 10 条:")
        for rec in update_records[:10]:
            print(f"  {rec['record_id']} → "
                  f"{quasi_field}={rec['fields'][quasi_field]}, "
                  f"{ship_field}={rec['fields'][ship_field]}")
        return {
            "xlsx_orders": len(xlsx_set),
            "matched": len(matched_set),
            "updated": 0,
            "missing": list(missing),
            "dry_run": True,
        }

    if not update_records:
        print("[完成] 没有需要更新的记录")
        return {
            "xlsx_orders": len(xlsx_set),
            "matched": len(matched_set),
            "updated": 0,
            "missing": list(missing),
        }

    updated = batch_update_records(token, app_token, progress_id, update_records)
    print(f"\n✅ 完成：实际更新 {updated} 条记录")
    return {
        "xlsx_orders": len(xlsx_set),
        "matched": len(matched_set),
        "updated": updated,
        "missing": list(missing),
    }


# ============================================================
# 命令行入口
# ============================================================
if __name__ == "__main__":
    p = argparse.ArgumentParser(
        description="把准发下载 xlsx 的'准发量'和'出厂量'按钢厂订单号更新到飞书'进度'表",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例:
  # 先配置飞书凭据（PowerShell）
  $env:FEISHU_APP_ID = "cli_xxxxx"
  $env:FEISHU_APP_SECRET = "xxxxx"

  # 自动找最新的 准发下载_*.xlsx
  python update_progress_delivery.py

  # 指定 xlsx + 干跑预览
  python update_progress_delivery.py --xlsx 准发下载_260816_1044.xlsx --dry-run
""")
    p.add_argument("--xlsx", default="",
                   help=f"输入 xlsx 路径（默认自动找最新的 {DEFAULT_XLSX_GLOB}）")
    p.add_argument("--app-token", default=None,
                   help=f"多维表格 app_token（默认 {DEFAULT_APP_TOKEN}）")
    p.add_argument("--progress-table", default=DEFAULT_TABLE_PROGRESS,
                   help=f"写入目标表名（默认 '{DEFAULT_TABLE_PROGRESS}'）")
    p.add_argument("--contract-table", default=DEFAULT_TABLE_CONTRACT,
                   help=f"查 record_id 的源表名（默认 '{DEFAULT_TABLE_CONTRACT}'）")
    p.add_argument("--resource-field", default=DEFAULT_FIELD_RESOURCE,
                   help=f"'资源号'字段名（默认 '{DEFAULT_FIELD_RESOURCE}'）")
    p.add_argument("--quasi-field", default=DEFAULT_FIELD_QUASI,
                   help=f"'准发量'字段名（默认 '{DEFAULT_FIELD_QUASI}'）")
    p.add_argument("--ship-field", default=DEFAULT_FIELD_SHIP,
                   help=f"'出厂量'字段名（默认 '{DEFAULT_FIELD_SHIP}'）")
    p.add_argument("--dry-run", action="store_true",
                   help="只预览不实际写入")
    args = p.parse_args()

    # 自动找最新的 xlsx
    xlsx_path = args.xlsx or find_latest_xlsx()
    if not xlsx_path or not os.path.exists(xlsx_path):
        print(f"❌ 找不到 xlsx 文件：{args.xlsx or DEFAULT_XLSX_GLOB}", file=sys.stderr)
        print(f"   请先运行 python download_bundle.py 生成数据", file=sys.stderr)
        sys.exit(1)
    print(f"使用 xlsx: {xlsx_path}")

    try:
        result = update_progress(
            xlsx_path=xlsx_path,
            app_token=args.app_token,
            progress_table_name=args.progress_table,
            contract_table_name=args.contract_table,
            resource_field=args.resource_field,
            quasi_field=args.quasi_field,
            ship_field=args.ship_field,
            dry_run=args.dry_run,
        )
        print("\n=== 最终结果 ===")
        print(f"  xlsx 中唯一订单数: {result['xlsx_orders']}")
        print(f"  合约表匹配数:      {result['matched']}")
        print(f"  实际更新:          {result['updated']}")
        print(f"  未匹配订单:        {len(result['missing'])} 个")
        if result['missing']:
            print(f"  未匹配订单号: {result['missing'][:20]}")
    except Exception as e:
        print(f"\n❌ 失败：{e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)
