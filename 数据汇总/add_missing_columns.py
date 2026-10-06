# -*- coding: utf-8 -*-
"""一次性补列：把「接口有键、飞书表里没列」的字段补齐到 ①③④⑤⑥。

## 背景
十一合一按「飞书列名 == 懂火接口键名」同名直连。历史上这几张表是人工建的，
有 27 个接口键在表里没有对应列，脚本只能记日志跳过（`--check-fields` 里显示为
「➖ 接口有键、表里无列 → 不入表」）。2026-10-05 用户拍板：**全部补上，跟 ERP 同步**。

## 补什么（30 列）
  ①出库   17 列：状态·销售日期·入库日期·锌层·涂料·结构·颜色·最低价·锁定发票·
                销售其它·销售费用·合计成本·采购单号·米数·目的港·捆包号N·新增时间
  ②销售明细 3 列：销售人·销售状态·出库操作
  ③应收    2 列：期初应收·期初应开发票
  ④往来    2 列：提交日期·米数
  ⑤客户    3 列：所属公司·最后更新·新增时间
  ⑥销售订单 3 列：销售日期·发货状态·应结金额

  ② 的 3 列是 2026-10-05 跑 `--check-fields` 时发现的漏网之鱼（首轮清单只覆盖
  ①③④⑤⑥）。② 是**动态探测表**（写入时实时读飞书现有字段建映射），所以列一建好、
  下一次同步就自动灌值，脚本侧一个字都不用改。实测值：销售人 9349/9352 非空（17 值）、
  销售状态 6612/9352 非空（已完成/已审核/请提交审核/待审核/锁定中）、出库操作 9352/9352
  非空（已足量/未完成）——都不是空列，全按文本建。

## 类型口径（按接口实测值定，不猜）
  · 米数 → **文本**。①出库的「米数」有 4955 行非空，值全是交期/仓库/往来方简称这类
    备注（形如「挂-××」「XXX-×.××」「×.××离港」），不是量 —— ERP 那个列名是历史
    包袱，装的根本不是数字；⑨库存里它已经是文本。文本能装数字、数字装不下文字，
    安全方向也只有这一个。（真实值不入库，免得把往来方写进公开仓库。）
  · 其余数字列 → 实测全是 2 位小数 → formatter "0.00"（走 Update_Data 的 _number_formatter）
  · 日期列 → {"auto_fill": false, "date_formatter": "yyyy/MM/dd"}（与 ①②⑨ 现有列一致）
  · 其余文本列 → 文本

## 安全性
**只新增字段，不碰任何已有字段、不碰任何记录。** 已有数据的列一个不动；
新列建出来是空的，下一次十一合一跑起来就把值灌进去。

## 事后：5 列被摘掉（2026-10-05 用户拍板）
补完列、拉了一次全量值后逐列看实测分布，这 5 列**没有任何信息量**，用户决定不要：
  ①状态      12854 行**全部**是「已出库」    （不同值只有 1 个）
  ①最低价    12854 行**全部**是 0.00         （不同值只有 1 个）
  ①销售其它  12850 行是 0.00，只有 4 行有真值
  ③期初应收 / ③期初应开发票   各 577 行**全部**是 0.00（不同值只有 1 个）
已在 `Update_Data.py` 的 `EXCLUDED_API_KEYS` 里登记 —— 既不入表，也不误报成「接口有键表里无列」。
**飞书表里这 5 个字段本身还在（会一直是空的）**，要删得去飞书手动删。
（同批的「捆包号N」是另一回事：它是每次请求都变的随机哈希，也已从飞书 ① 表删掉了。）

## 用法（一律从仓库根目录运行）
    python 数据汇总/add_missing_columns.py --dry-run    # 只列清单
    python 数据汇总/add_missing_columns.py --apply      # 真加
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

DATE_PROP = {"auto_fill": False, "date_formatter": "yyyy/MM/dd"}

# 表 id → {要补的字段: 类型}
PLAN = {
    "①出库数据": ("tblolnj06JZkYNiU", {
        "状态": "text", "销售日期": "datetime", "入库日期": "datetime",
        "锌层": "text", "涂料": "text", "结构": "text", "颜色": "text",
        "最低价": "number", "锁定发票": "number", "销售其它": "number",
        "销售费用": "number", "合计成本": "number", "采购单号": "text",
        "米数": "text", "目的港": "text", "捆包号N": "text", "新增时间": "datetime",
    }),
    "②销售明细": ("tblcEZoQatk7lCAO", {
        "销售人": "text", "销售状态": "text", "出库操作": "text",
    }),
    "③应收汇总": ("tblpjne9dIuif5HD", {
        "期初应收": "number", "期初应开发票": "number",
    }),
    "④往来": ("tblbS1dPaDVL3GY8", {
        "提交日期": "datetime", "米数": "text",
    }),
    "⑤客户": ("tblCE7zIWs804RR5", {
        "所属公司": "text", "最后更新": "datetime", "新增时间": "datetime",
    }),
    "⑥销售订单": ("tblJZIyNXoe8PZer", {
        "销售日期": "datetime", "发货状态": "text", "应结金额": "number",
    }),
}


def build_field(name: str, typ: str) -> dict:
    """按类型构造建字段请求体；number 必须显式带 formatter（否则飞书默认 "0.0"）"""
    fd = {"field_name": name, "type": ud._STRATEGY_TO_CREATE_CODE[typ]}
    if typ == "number":
        fd["property"] = {"formatter": ud._number_formatter(name)}
    elif typ == "datetime":
        fd["property"] = DATE_PROP
    return fd


def list_fields(tok: str, tid: str) -> dict:
    r = ud._fs_session().get(
        f"{ud.FEISHU_OPEN_BASE}/bitable/v1/apps/{ud.BITABLE_APP_TOKEN}/tables/{tid}/fields",
        headers={"Authorization": f"Bearer {tok}"}, params={"page_size": 200}, timeout=30)
    j = r.json()
    if j.get("code") != 0:
        raise RuntimeError(f"读字段失败: {j}")
    return {it["field_name"]: it for it in j["data"]["items"]}


def create_field(tok: str, tid: str, body: dict) -> tuple:
    r = ud._fs_session().post(
        f"{ud.FEISHU_OPEN_BASE}/bitable/v1/apps/{ud.BITABLE_APP_TOKEN}/tables/{tid}/fields",
        headers={"Authorization": f"Bearer {tok}", "Content-Type": "application/json; charset=utf-8"},
        json=body, timeout=30)
    j = r.json()
    return j.get("code"), j.get("msg")


def main() -> int:
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--dry-run", action="store_true")
    g.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    tok = ud.feishu_token()
    buf = io.StringIO()
    buf.write("=== 补列（%s）base=%s ===\n\n"
              % ("dry-run" if args.dry_run else "真加", ud.BITABLE_APP_TOKEN))

    total_planned = total_added = total_skip = 0
    failures = []

    for tname, (tid, want) in PLAN.items():
        have = list_fields(tok, tid)
        todo = [(n, t) for n, t in want.items() if n not in have]
        skip = [n for n in want if n in have]
        buf.write("【%s】现有 %d 字段，要补 %d 个%s\n"
                  % (tname, len(have), len(todo),
                     ("，已存在跳过: " + ", ".join(skip)) if skip else ""))
        total_planned += len(want)
        total_skip += len(skip)

        for n, t in todo:
            body = build_field(n, t)
            buf.write("   + %-12s %-9s %s\n"
                      % (n, t, body.get("property", "")))
            if args.dry_run:
                continue
            code, msg = create_field(tok, tid, body)
            if code == 0:
                total_added += 1
                buf.write("       ✅ 已建\n")
            else:
                failures.append((tname, n, code, msg))
                buf.write("       ❌ code=%s %s\n" % (code, msg))
        buf.write("\n")

    buf.write("计划 %d 个，其中已存在跳过 %d 个，实建 %d 个，失败 %d 个\n"
              % (total_planned, total_skip, total_added, len(failures)))

    if args.dry_run:
        buf.write("\n(--dry-run，未写任何东西)\n")
    else:
        # 复核：重新读一遍元数据 + 检查每个新 Number 字段的 formatter
        buf.write("\n=== 复核：重新读字段元数据 ===\n")
        for tname, (tid, want) in PLAN.items():
            have = list_fields(tok, tid)
            bad = []
            for n, t in want.items():
                it = have.get(n)
                if it is None:
                    bad.append(f"{n}: 没建出来")
                    continue
                if it["type"] != ud._STRATEGY_TO_CREATE_CODE[t]:
                    bad.append(f"{n}: type={it['type']} 期望 {ud._STRATEGY_TO_CREATE_CODE[t]}")
                if t == "number":
                    f = (it.get("property") or {}).get("formatter")
                    if f != ud._number_formatter(n):
                        bad.append(f"{n}: formatter={f} 期望 {ud._number_formatter(n)}")
            buf.write("  %-10s %s\n" % (tname, "✅ 全部到位" if not bad else "❌ " + "; ".join(bad)))

    out = ROOT / "数据汇总" / ("add_columns_result.txt" if not args.dry_run else "add_columns_dryrun.txt")
    out.write_text(buf.getvalue(), encoding="utf-8")
    print(buf.getvalue())
    return 0 if not failures else 1


if __name__ == "__main__":
    raise SystemExit(main())
