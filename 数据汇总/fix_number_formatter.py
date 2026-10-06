# -*- coding: utf-8 -*-
"""一次性修复：把 ⑦采购订单 / ⑧采购明细 / ⑨库存 里 Number 字段的显示精度改对。

## 问题
这三张表是 Update_Data.py 用 `--ensure-tables` 自动建的，建表代码只传了
`{"field_name": f, "type": 2}`、没带 property，飞书就给了默认 formatter `"0.0"`。
后果：屏幕把 3 位小数四舍五入到 1 位 —— ERP 的 2.190 显示成 2.2、0.370 显示成 0.4、
4.136 显示成 4.1。**存储值是对的**（逐条比过 1119 行，飞书 ↔ ERP 导出 ↔ ERP 接口三方一致），
被改掉的只有屏幕显示。

①②③④⑥ 是人工建的表，formatter 本来就是 0.000(重量)/0.00(金额)/0(计数)，不受影响。

## 这次改什么
只改字段的 property.formatter，**不碰 field_name、不碰 type、不碰任何记录**。
精度口径按用户 2026-10-04 定的：重量 3 位（与 ①②⑥ 一致）、金额/单价/税率 2 位、计数 0 位。

## 用法
    python 数据汇总/fix_number_formatter.py --dry-run      # 只列清单，不写
    python 数据汇总/fix_number_formatter.py --one          # 只改 ⑨库存 的「重量(吨)」一个（验证用）
    python 数据汇总/fix_number_formatter.py --apply        # 改全部 29 个
"""
from __future__ import annotations

import argparse
import importlib.util
import io
import os
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent
os.chdir(ROOT)
_spec = importlib.util.spec_from_file_location("ud", ROOT / "数据汇总" / "Update_Data.py")
ud = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ud)

# ---- 要改的清单：表 -> {字段名: 目标 formatter} ----
# 依据（全部来自本地 10-03 CSV 实测的小数位数）：
#   ⑨ 重量(吨) 最多3位 / 可售重量 最多3位 / 采购单价 2 / 成本单价 2 / 采购金额 2 /
#      费用金额 2 / 采购税率 2(0.13) / 销售单价、最低售价 全整数 / 件(张)数、可售件数、库龄 全整数
#   ⑦ 订单重量、入库重量 最多4位 → 按用户定的 3 位口径 / 金额类 2 位
#   ⑧ 同上
PLAN = {
    "kucun": {
        "重量(吨)": "0.000",
        "可售重量": "0.000",
        "采购单价": "0.00",
        "成本单价": "0.00",
        "销售单价": "0.00",
        "最低售价": "0.00",
        "采购金额": "0.00",
        "费用金额": "0.00",
        "采购税率": "0.00",
        "件(张)数": "0",
        "可售件数": "0",
        "库龄": "0",
    },
    "cgdd": {
        "订单重量": "0.000",
        "入库重量": "0.000",
        "订单金额": "0.00",
        "入库金额": "0.00",
        "采购费用": "0.00",
        "其它款项": "0.00",
        "合同定金": "0.00",
        "应结金额": "0.00",
        "已结金额": "0.00",
        "未结金额": "0.00",
        "合同未结": "0.00",
    },
    "cgmx": {
        "订单重量": "0.000",
        "入库重量": "0.000",
        "采购单价": "0.00",
        "采购金额": "0.00",
        "入库金额": "0.00",
        "采购税率": "0.00",
    },
}

# --one 只动这一个（先验证 API 行为和存储值不受影响）
ONE = ("kucun", "重量(吨)")


def list_fields(tok: str, tid: str) -> dict:
    """返回 {字段名: 字段原始 dict}"""
    r = ud._fs_session().get(
        f"{ud.FEISHU_OPEN_BASE}/bitable/v1/apps/{ud.BITABLE_APP_TOKEN}/tables/{tid}/fields",
        headers={"Authorization": f"Bearer {tok}"}, params={"page_size": 200}, timeout=30)
    j = r.json()
    if j.get("code") != 0:
        raise RuntimeError(f"读字段失败: {j}")
    return {it["field_name"]: it for it in j["data"]["items"]}


def update_formatter(tok: str, tid: str, field: dict, target: str) -> tuple:
    r = ud._fs_session().put(
        f"{ud.FEISHU_OPEN_BASE}/bitable/v1/apps/{ud.BITABLE_APP_TOKEN}/tables/{tid}/fields/{field['field_id']}",
        headers={"Authorization": f"Bearer {tok}", "Content-Type": "application/json; charset=utf-8"},
        json={"field_name": field["field_name"], "type": field["type"],
              "property": {"formatter": target}},
        timeout=30)
    j = r.json()
    return j.get("code"), j.get("msg")


def main() -> int:
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--dry-run", action="store_true")
    g.add_argument("--one", action="store_true")
    g.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    tok = ud.feishu_token()
    buf = io.StringIO()

    todo = []          # (part, name, fid, type, old_formatter, target, tid)
    problems = []
    for part, want in PLAN.items():
        tid = ud.table_id_of(part)
        have = list_fields(tok, tid)
        for name, target in want.items():
            it = have.get(name)
            if it is None:
                problems.append(f"{part} 表里找不到字段「{name}」")
                continue
            if it["type"] != 2:
                problems.append(f"{part}.{name} 不是 Number（type={it['type']}），跳过")
                continue
            old = (it.get("property") or {}).get("formatter")
            todo.append((part, name, it["field_id"], it["type"], old, target, tid))

    if args.one:
        todo = [t for t in todo if (t[0], t[1]) == ONE]

    buf.write("=== 计划（%s）===\n" % ("dry-run" if args.dry_run else ("只改一个" if args.one else "全改")))
    buf.write("要改 %d 个字段，覆盖 %d 张表\n\n" % (len(todo), len({t[6] for t in todo})))
    for part, name, fid, typ, old, target, _tid in todo:
        mark = "  (已是目标值，跳过)" if old == target else ""
        buf.write("  %-6s %-12s %-14s  %-8s -> %-8s%s\n" % (part, name, fid, repr(old), repr(target), mark))
    if problems:
        buf.write("\n⚠️ 问题:\n")
        for p in problems:
            buf.write("  %s\n" % p)

    if args.dry_run:
        buf.write("\n(--dry-run，未写任何东西)\n")
        out = ROOT / "数据汇总" / "fix_formatter_dryrun.txt"
        out.write_text(buf.getvalue(), encoding="utf-8")
        print(buf.getvalue())
        return 0

    # 改前先记下一条记录的存储值，改完比对，证明只动了显示格式
    before = {}
    chk = "kucun"
    chk_tid = ud.table_id_of(chk)
    recs = ud.bitable_list_records(tok, chk_tid)
    for r in recs[:3]:
        before[r["record_id"]] = {k: v for k, v in (r.get("fields") or {}).items()}

    done, failed = 0, []
    for part, name, fid, typ, old, target, tid in todo:
        if old == target:
            continue
        code, msg = update_formatter(tok, tid, {"field_id": fid, "field_name": name, "type": typ}, target)
        if code == 0:
            done += 1
            buf.write("  ✅ %-6s %-12s %s -> %s\n" % (part, name, old, target))
        else:
            failed.append((part, name, code, msg))
            buf.write("  ❌ %-6s %-12s code=%s %s\n" % (part, name, code, msg))

    buf.write("\n成功 %d / 失败 %d\n" % (done, len(failed)))

    # 复核 1：字段元数据真的变了
    buf.write("\n=== 复核：重新读字段元数据 ===\n")
    want_all = {}
    for part, want in PLAN.items():
        for n, t in want.items():
            want_all.setdefault(part, {})[n] = t
    for part in PLAN:
        have = list_fields(tok, ud.table_id_of(part))
        bad = []
        for n, t in want_all[part].items():
            got = (have.get(n, {}).get("property") or {}).get("formatter")
            if got != t:
                bad.append(f"{n}: 期望{t} 实际{got}")
        buf.write("  %-6s %s\n" % (part, "✅ 全部到位" if not bad else "❌ " + "; ".join(bad)))

    # 复核 2：存储值一个字没动
    buf.write("\n=== 复核：记录存储值是否被改动 ===\n")
    recs2 = ud.bitable_list_records(tok, chk_tid)
    after = {r["record_id"]: (r.get("fields") or {}) for r in recs2}
    diff = 0
    for rid, f0 in before.items():
        f1 = after.get(rid, {})
        for k, v0 in f0.items():
            if str(f1.get(k)) != str(v0):
                diff += 1
                if diff <= 5:
                    buf.write("  ❌ %s.%s 改前=%r 改后=%r\n" % (rid[:8], k, v0, f1.get(k)))
    buf.write("  %s\n" % ("✅ 前 3 条记录全部字段逐值未变" if diff == 0 else "❌ 有 %d 处变化" % diff))

    out = ROOT / "数据汇总" / "fix_formatter_result.txt"
    out.write_text(buf.getvalue(), encoding="utf-8")
    print(buf.getvalue())
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
