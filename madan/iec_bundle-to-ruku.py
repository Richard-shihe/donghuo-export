#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IEC 码单 → 飞书 Bitable 入库（阶段2：增量捆包写入采购订单明细汇总）
====================================================================

职责：
  把 IEC 码单比对出的"新增捆包"按字段名映射表写入飞书 Bitable
  目标表：采购订单明细汇总 (tblrhqzHuTsAprU3) @ App OSuobf2ZkaWtUAsXEE9c3aBTnwh

上游：
  由比对脚本（_compare_madan.py 或同等逻辑）生成 CSV，列结构同 IEC 码单导出 xlsx：
    码单号, 销售订单号, 销售订单子项号, 钢厂订单号, 捆包号, 净重, 毛重,
    实际长度, 规格, 品种, 牌号, 车船号, 船名, 炉号, 张数/根数/支数,
    收货单位名称, 交货地点名称, 代运公司, 出厂日期, 最终用户名称, 渠道规格, 比对结果

映射规则（来自 tblXIj688zJBjCNP 字段名映射2，已固化为本脚本常量；如映射表调整请同步改 MAPPING）：
  A. Excel 列直接映射：品种→品名 / 规格→规格 / 钢厂订单号→结构 / 净重→重量 /
     捆包号→捆包号 / 销售订单号→合同号 / 车船号→车船号
  B. 固定值：所属公司=上海士禾实业有限公司 / 等级=正品 /
     供应商=上海宝钢钢材贸易有限公司 / 处理状态=已提交审核
  C. 按"钢厂订单号 = 结构"匹配现有记录取值：
     材质 / 产地 / 锌层 / 采购单价 / 采购明细ID / 采购订单号 / 批次时间
  D. 主动空白（按映射表）：颜色 / 件(张)数 / 米数 / 成本单价 / 销售单价
  E. 入库日期 = Excel"出厂日期"（码单时间）
  F. 脚本写入：写入结果="从 IEC 码单导入 {时间}" / 写入时间=当前时间戳
  G. 三列全空留空：货权 / 涂料 / 仓库 / 库位号 / 提单号 / 备注
  H. 只读/系统自动：采购订单(Formula) / 创建时间(CreatedTime)
  链接：父记录(SingleLink) = 匹配记录的 record_id

命令行（建议从仓库根目录执行）：
  python madan/iec_bundle-to-ruku.py --csv 比对_最新多出记录.csv          # 干跑（默认）
  python madan/iec_bundle-to-ruku.py --csv 比对_最新多出记录.csv --commit # 实际写入

环境变量（来自仓库根目录 .env）：
  FEISHU_APP_ID / FEISHU_APP_SECRET  飞书自建应用（需 bitable:app 权限 + 目标表可编辑协作者）
"""
from __future__ import annotations

import argparse
import csv
import datetime
import json
import os
import sys
from pathlib import Path

import requests

# ============================================================
# 路径定位：脚本位于 madan/，.env 在仓库根
# ============================================================
_HERE = Path(__file__).resolve().parent
_REPO_ROOT = _HERE.parent
for _p in (str(_REPO_ROOT), str(_HERE)):
    if _p and _p not in sys.path:
        sys.path.insert(0, _p)
try:
    from dotenv import load_dotenv as _load_dotenv
    for _ep in (_REPO_ROOT / ".env", Path.cwd() / ".env"):
        if _ep.is_file():
            _load_dotenv(_ep, override=False)
            break
except ImportError:
    pass

# ============================================================
# 常量
# ============================================================
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}
APP_TOKEN = "OSuobf2ZkaWtUAsXEE9c3aBTnwh"
TARGET_TABLE_ID = "tblrhqzHuTsAprU3"        # 采购订单明细汇总
MAPPING_TABLE_ID = "tblXIj688zJBjCNP"       # 字段名映射2（仅作参考源，规则已固化到本脚本）
DEFAULT_CSV = "比对_最新多出记录.csv"

# A. Excel 列 → Bitable 字段（来自映射表"对应字段名"列）
EXCEL_TO_BITABLE = {
    "品种":       "品名",
    "规格":       "规格",
    "钢厂订单号":  "结构",
    "净重":       "重量",
    "捆包号":     "捆包号",
    "销售订单号":  "合同号",
    "车船号":     "车船号",
}

# B. 固定值（来自映射表"固定字段内容"列）
FIXED_VALUES = {
    "所属公司": "上海士禾实业有限公司",
    "等级":     "正品",
    "供应商":   "上海宝钢钢材贸易有限公司",
    "处理状态": "已提交审核",
}

# C. 按"钢厂订单号 = 结构"匹配现有记录取值（含材质）
LOOKUP_FIELDS = {"材质", "产地", "锌层", "采购单价", "采购明细ID", "采购订单号", "批次时间"}
LOOKUP_LINK_FIELD = "父记录"   # SingleLink，同样按结构匹配取 record_id

# D. 主动空白
BLANK_FIELDS = {"颜色", "件(张)数", "米数", "成本单价", "销售单价"}

# H. 只读/系统自动
READONLY_FIELDS = {"采购订单", "创建时间"}


# ============================================================
# 飞书 API
# ============================================================
def get_token() -> str:
    app_id = os.environ.get("FEISHU_APP_ID", "").strip()
    app_secret = os.environ.get("FEISHU_APP_SECRET", "").strip()
    if not app_id or not app_secret:
        print("❌ 缺少 FEISHU_APP_ID / FEISHU_APP_SECRET"); sys.exit(1)
    r = requests.post(f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
                      json={"app_id": app_id, "app_secret": app_secret},
                      timeout=15, proxies=NO_PROXY)
    d = r.json()
    if d.get("code") != 0:
        print(f"❌ token 失败: {d}"); sys.exit(1)
    return d["tenant_access_token"]


def fetch_field_types(tok: str) -> dict:
    """返回 {字段名: type}，type 来自飞书 Bitable 字段定义"""
    r = requests.get(f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TARGET_TABLE_ID}/fields",
                     headers={"Authorization": f"Bearer {tok}"},
                     params={"page_size": 100}, timeout=15, proxies=NO_PROXY)
    d = r.json()
    if d.get("code") != 0:
        print(f"❌ 拉字段失败: {json.dumps(d, ensure_ascii=False)[:300]}"); sys.exit(1)
    return {f.get("field_name"): f.get("type") for f in (d.get("data") or {}).get("items") or []}


def fetch_match_index(tok: str) -> dict:
    """遍历目标表全部记录，按"结构"字段建索引 {结构值: (record_id, fields_dict)}"""
    print("\n[建匹配索引] 遍历目标表，按'结构'字段聚合...")
    idx = {}
    page_token = None
    total = 0
    while True:
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TARGET_TABLE_ID}/records",
                         headers={"Authorization": f"Bearer {tok}"},
                         params=params, timeout=20, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") != 0:
            print(f"  ❌ 拉记录失败: {json.dumps(d, ensure_ascii=False)[:200]}")
            return idx
        data = d.get("data") or {}
        items = data.get("items") or []
        total += len(items)
        for rec in items:
            fd = rec.get("fields") or {}
            jiegou = fd.get("结构")
            if jiegou and jiegou not in idx:
                idx[jiegou] = (rec.get("record_id"), fd)
        if not data.get("has_more"):
            break
        page_token = data.get("page_token")
    print(f"  ✅ 共 {total} 条记录，聚合 {len(idx)} 个结构值")
    return idx


# ============================================================
# 字段值转换工具
# ============================================================
def to_number(val):
    """把飞书返回的任意值转成 float（Number 字段写入需要）"""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        return float(val)
    if isinstance(val, list) and val:
        first = val[0]
        if isinstance(first, dict):
            return float(first.get("text") or first.get("name") or 0)
        return float(first)
    if isinstance(val, dict):
        return float(val.get("text") or val.get("name") or 0)
    raise ValueError(f"无法转 Number: {val!r}")


# ============================================================
# 行映射
# ============================================================
def row_to_fields(row: dict, field_types: dict, match_idx: dict, batch_ts: int) -> tuple[dict, list]:
    """把 CSV 一行按映射规则转成 Bitable fields。返回 (fields, issues)"""
    fields = {}
    issues = []

    # A. Excel 列直接映射
    for excel_col, bitable_field in EXCEL_TO_BITABLE.items():
        raw = (row.get(excel_col) or "").strip()
        ftype = field_types.get(bitable_field)
        if ftype == 2:  # Number
            try:
                fields[bitable_field] = float(raw) if raw else None
            except ValueError:
                fields[bitable_field] = None
                issues.append(f"{bitable_field}: Number 类型但值 '{raw}' 非数字")
        else:
            fields[bitable_field] = raw if raw else None

    # B. 固定值
    for fname, fval in FIXED_VALUES.items():
        fields[fname] = fval

    # C. 按结构匹配现有记录取值
    jiegou = (row.get("钢厂订单号") or "").strip()
    matched = match_idx.get(jiegou)
    if matched:
        rec_id, matched_fields = matched
        for fname in LOOKUP_FIELDS:
            val = matched_fields.get(fname)
            if val is None:
                issues.append(f"{fname}: 匹配记录 结构='{jiegou}'，但该字段值为空")
            else:
                # Number 字段强制转换
                if field_types.get(fname) == 2:
                    try:
                        val = to_number(val)
                    except (ValueError, TypeError):
                        issues.append(f"{fname}: Number 转换失败，原值={val!r}")
                        val = None
            fields[fname] = val
        # 父记录 SingleLink: [record_id]
        fields[LOOKUP_LINK_FIELD] = [rec_id]
    else:
        for fname in LOOKUP_FIELDS:
            fields[fname] = None
        fields[LOOKUP_LINK_FIELD] = None
        issues.append(f"结构='{jiegou}' 在现有表未匹配到记录")

    # E. 入库日期 = Excel 出厂日期
    churi = (row.get("出厂日期") or "").strip()
    if churi:
        try:
            dt = datetime.datetime.strptime(churi, "%Y-%m-%d %H:%M:%S")
            fields["入库日期"] = int(dt.timestamp() * 1000)
        except ValueError:
            issues.append(f"入库日期: Excel '出厂日期' 格式异常 '{churi}'")
    else:
        issues.append("入库日期: Excel '出厂日期' 为空")

    # F. 脚本写入字段
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    fields["写入结果"] = f"从 IEC 码单导入 {now_str}"
    fields["写入时间"] = batch_ts

    return fields, issues


# ============================================================
# 主流程
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="IEC 码单增量捆包 → 飞书 Bitable 入库")
    parser.add_argument("--csv", default=DEFAULT_CSV, help=f"比对结果 CSV 路径（默认 {DEFAULT_CSV}）")
    parser.add_argument("--commit", action="store_true", help="实际写入（默认 dry-run 仅打印）")
    args = parser.parse_args()

    mode = "实际写入" if args.commit else "干跑(仅打印)"
    print("=" * 72)
    print(f"{mode}: CSV → {TARGET_TABLE_ID} (采购订单明细汇总)")
    print(f"  App: {APP_TOKEN}")
    print(f"  CSV: {args.csv}")
    print("=" * 72)

    # 1) 读 CSV
    if not os.path.exists(args.csv):
        print(f"❌ CSV 不存在: {args.csv}"); sys.exit(1)
    with open(args.csv, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    print(f"\n✅ 读取 {len(rows)} 条 CSV 记录")

    # 2) 拉目标表字段类型 + 建匹配索引
    tok = get_token()
    field_types = fetch_field_types(tok)
    print(f"✅ 目标表字段类型: {len(field_types)} 个字段")
    match_idx = fetch_match_index(tok)

    # 3) 逐行映射
    batch_ts = int(datetime.datetime.now().timestamp() * 1000)
    all_fields = []
    all_issues = []
    for row in rows:
        fields, issues = row_to_fields(row, field_types, match_idx, batch_ts)
        all_fields.append(fields)
        all_issues.append(issues)

    # 4) 匹配统计
    matched_count = sum(1 for row in rows if (row.get("钢厂订单号") or "").strip() in match_idx)
    print(f"\n[匹配] {matched_count}/{len(rows)} 条匹配成功")
    unmatched = [(i + 1, (row.get("钢厂订单号") or "").strip())
                 for i, row in enumerate(rows)
                 if (row.get("钢厂订单号") or "").strip() not in match_idx]
    if unmatched:
        print(f"  ❌ 未匹配 {len(unmatched)} 条:")
        for i, jg in unmatched:
            print(f"     记录{i}: 钢厂订单号={jg}")

    # 5) 打印前 3 条映射结果（预览）
    print("\n[映射预览] 前 3 条:")
    for i, (row, fields, issues) in enumerate(zip(rows, all_fields, all_issues), 1):
        if i > 3: break
        print(f"\n--- 记录 {i} | 捆包号={row.get('捆包号')} | 钢厂订单号={row.get('钢厂订单号')} ---")
        for k, v in fields.items():
            if isinstance(v, (list, dict)):
                vs = json.dumps(v, ensure_ascii=False)
            else:
                vs = str(v)
            print(f"   {k:>14} = {vs}")
        if issues:
            print(f"   ⚠️ 问题:")
            for iss in issues:
                print(f"       - {iss}")

    # 6) 写入或终止
    if not args.commit:
        print("\n" + "=" * 72)
        print("✅ 干跑完成 — 未实际写入")
        print("   确认无误后加 --commit 实际写入")
        print("=" * 72)
        return

    # 实际写入
    print("\n" + "=" * 72)
    print(f"⚠️ 实际写入: 提交 {len(all_fields)} 条到飞书 Bitable")
    print("=" * 72)

    payload = []
    for fields in all_fields:
        clean = {k: v for k, v in fields.items() if v is not None}
        payload.append({"fields": clean})

    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TARGET_TABLE_ID}/records/batch_create"
    r = requests.post(url,
                      headers={"Authorization": f"Bearer {tok}", "Content-Type": "application/json"},
                      json={"records": payload}, timeout=30, proxies=NO_PROXY)
    d = r.json()
    if d.get("code") != 0:
        print(f"\n❌ 写入失败: code={d.get('code')} msg={d.get('msg')}")
        print(f"  完整响应: {json.dumps(d, ensure_ascii=False)[:2000]}")
        sys.exit(1)

    records = (d.get("data") or {}).get("records") or []
    print(f"\n✅ 写入成功！返回 {len(records)} 条记录")
    print(f"\n{'#':>3}  {'record_id':<20}  {'捆包号':<14}  {'钢厂订单号':<14}  {'采购订单号'}")
    print("-" * 80)
    for i, (rec, row) in enumerate(zip(records, rows), 1):
        rid = rec.get("record_id", "")
        fd = rec.get("fields") or {}
        kbbh = fd.get("捆包号", "")
        jg = (row.get("钢厂订单号") or "").strip()
        order_no = fd.get("采购订单号", "")
        print(f"{i:>3}  {rid:<20}  {kbbh:<14}  {jg:<14}  {order_no}")
    print(f"\n飞书表格: https://s2v31ke6sl.feishu.cn/base/{APP_TOKEN}/table/{TARGET_TABLE_ID}")


if __name__ == "__main__":
    main()
