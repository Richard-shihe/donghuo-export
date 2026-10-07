# -*- coding: utf-8 -*-
"""一次性修复：把旧基地「费用管理」表(tblnlyACUGBUXfUS)里 Number 字段的显示精度改对。

## 为什么
这张表是 `sync_feiyong_to_bitable.py` 建的（也是自动建表、没带 property），
飞书给了默认 formatter `"0.0"`，屏幕把小数四舍五入到 1 位：
  · 税率  0.13 显示成 0.1   · -0.21 显示成 -0.2
  · 费用金额 1234.56 显示成 1234.6
  · 计价重量 13.135 显示成 13.1（3082 行是 3 位小数）
**存储值是对的，坏的只有屏幕显示** —— 和 2026-10-04 发现 ⑦⑧⑨ 那次同一个病。

## 精度口径（按 11158 行全量实测的小数位数定，不靠猜）
    ├ 税率 / 费用单价 / 费用金额 / 计算金额：实测 0~2 位        → "0.00"
    ├ 计价重量：实测 0~4 位（3 位 3082 行、4 位 1 行）          → "0.000"（用户 2026-10-05 定的重量 3 位口径）
    └ 米数：11158 行**全空**，实测不到任何小数位              → "0.00"（无数据可依的中性取值，可随时改）

## 只改什么
只改字段的 property.formatter，**不碰 field_name、不碰 type、不碰任何记录**。

## 用法（一律从仓库根目录运行）
    python 费用录入/fix_old_feiyong_formatter.py --dry-run    # 只列清单
    python 费用录入/fix_old_feiyong_formatter.py --one        # 只改「计价重量」一个（验证 API 行为）
    python 费用录入/fix_old_feiyong_formatter.py --apply      # 改全部 6 个
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

# 旧基地「费用管理」表 —— 九合一写的是另一个 base（VahHb3YDBaBTwTsCjeAcaAhhnHc），别搞混
APP = "WJ2IbCPUgax1Phsogk0cC0eNnCd"
TID = "tblnlyACUGBUXfUS"

PLAN = {
    "米数": "0.00",
    "税率": "0.00",
    "费用单价": "0.00",
    "费用金额": "0.00",
    "计价重量": "0.000",
    "计算金额": "0.00",
}

# --one 只动这一个（先验证：改完存储值必须一个字不变）
ONE = "计价重量"


def list_fields(tok: str) -> dict:
    """返回 {字段名: 字段原始 dict}"""
    r = ud._fs_session().get(
        f"{ud.FEISHU_OPEN_BASE}/bitable/v1/apps/{APP}/tables/{TID}/fields",
        headers={"Authorization": f"Bearer {tok}"}, params={"page_size": 200}, timeout=30)
    j = r.json()
    if j.get("code") != 0:
        raise RuntimeError(f"读字段失败: {j}")
    return {it["field_name"]: it for it in j["data"]["items"]}


def update_formatter(tok: str, field: dict, target: str) -> tuple:
    r = ud._fs_session().put(
        f"{ud.FEISHU_OPEN_BASE}/bitable/v1/apps/{APP}/tables/{TID}/fields/{field['field_id']}",
        headers={"Authorization": f"Bearer {tok}", "Content-Type": "application/json; charset=utf-8"},
        json={"field_name": field["field_name"], "type": field["type"],
              "property": {"formatter": target}},
        timeout=30)
    j = r.json()
    return j.get("code"), j.get("msg")


def read_records(tok: str, n: int = 3) -> dict:
    """读前 n 条记录的存储值，用来证明只动显示、不动数据"""
    r = ud._fs_session().get(
        f"{ud.FEISHU_OPEN_BASE}/bitable/v1/apps/{APP}/tables/{TID}/records",
        headers={"Authorization": f"Bearer {tok}"},
        params={"page_size": n}, timeout=60)
    j = r.json()
    if j.get("code") != 0:
        raise RuntimeError(f"读记录失败: {j}")
    return {it["record_id"]: (it.get("fields") or {}) for it in j["data"]["items"]}


def main() -> int:
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--dry-run", action="store_true")
    g.add_argument("--one", action="store_true")
    g.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    tok = ud.feishu_token()
    buf = io.StringIO()

    have = list_fields(tok)
    todo, problems = [], []
    for name, target in PLAN.items():
        it = have.get(name)
        if it is None:
            problems.append(f"表里找不到字段「{name}」")
            continue
        if it["type"] != 2:
            problems.append(f"「{name}」不是 Number（type={it['type']}），跳过")
            continue
        old = (it.get("property") or {}).get("formatter")
        todo.append((name, it["field_id"], it["type"], old, target))

    if args.one:
        todo = [t for t in todo if t[0] == ONE]

    buf.write("=== 旧「费用管理」表 formatter 修复（%s）===\n"
              % ("dry-run" if args.dry_run else ("只改一个" if args.one else "全改")))
    buf.write("base=%s  table=%s\n\n" % (APP, TID))
    buf.write("要改 %d 个字段：\n" % len(todo))
    for name, fid, typ, old, target in todo:
        mark = "  (已是目标值，跳过)" if old == target else ""
        buf.write("  %-10s %-14s  %-8s -> %-8s%s\n" % (name, fid, repr(old), repr(target), mark))
    if problems:
        buf.write("\n⚠️ 问题:\n")
        for p in problems:
            buf.write("  %s\n" % p)

    if args.dry_run:
        buf.write("\n(--dry-run，未写任何东西)\n")
        out = ROOT / "费用管理" / "fix_old_formatter_dryrun.txt"
        out.write_text(buf.getvalue(), encoding="utf-8")
        print(buf.getvalue())
        return 0

    # 改前记下 3 条记录的存储值
    before = read_records(tok)

    done, failed = 0, []
    for name, fid, typ, old, target in todo:
        if old == target:
            continue
        code, msg = update_formatter(tok, {"field_id": fid, "field_name": name, "type": typ}, target)
        if code == 0:
            done += 1
            buf.write("  ✅ %-10s %s -> %s\n" % (name, old, target))
        else:
            failed.append((name, code, msg))
            buf.write("  ❌ %-10s code=%s %s\n" % (name, code, msg))

    buf.write("\n成功 %d / 失败 %d\n" % (done, len(failed)))

    # 复核 1：字段元数据真的变了
    buf.write("\n=== 复核 1：重新读字段元数据 ===\n")
    have2 = list_fields(tok)
    bad = []
    for name, target in PLAN.items():
        got = (have2.get(name, {}).get("property") or {}).get("formatter")
        if got != target:
            bad.append(f"{name}: 期望{target} 实际{got}")
    buf.write("  %s\n" % ("✅ 6 个字段全部到位" if not bad else "❌ " + "; ".join(bad)))

    # 复核 2：存储值一个字没动
    buf.write("\n=== 复核 2：记录存储值是否被改动 ===\n")
    after = read_records(tok)
    diff = 0
    for rid, f0 in before.items():
        f1 = after.get(rid, {})
        for k, v0 in f0.items():
            if str(f1.get(k)) != str(v0):
                diff += 1
                if diff <= 5:
                    buf.write("  ❌ %s.%s 改前=%r 改后=%r\n" % (rid[:8], k, v0, f1.get(k)))
    buf.write("  %s\n" % ("✅ 前 %d 条记录全部字段逐值未变" % len(before)
                          if diff == 0 else "❌ 有 %d 处变化" % diff))

    out = ROOT / "费用管理" / "fix_old_formatter_result.txt"
    out.write_text(buf.getvalue(), encoding="utf-8")
    print(buf.getvalue())
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
