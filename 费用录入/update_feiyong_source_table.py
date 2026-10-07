#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
飞书「费用录入」源表改造（幂等）

三个方向的费用写入方式不同，源表需要能区分并定位：

  方向    定位依据                      ERP 写接口
  销售    销售单号 + 可选 牌号/捆包号      c_feiyon/save       (danhao + xids + leixin='销售')
  采购    采购单号                       c_feiyon/save       (danhao + leixin='采购')
  库存    捆包号                         c_feiyon/save_kcfy  (zid = 库存行 id)

本次改动：
  1. 主字段「订单号」→ 更名「关联单号」（销售单号 X… / 采购单号 C…）
  2. 新增「费用类型」单选：销售 / 采购 / 库存
  3. 新增「捆包号」文本（库存费用必填；销售费用想只加某一捆时也填）
  4. 新增「牌号」文本（销售费用想按材质细分时填，对应 ERP 出库行的「材质」）
  5. 新增「懂火费用编号」文本（ERP 里的 F 开头编号，人找单用；
     「懂火记录号」存的是 ERP 内部 id，程序删除/定位用）

用法（从仓库根目录运行）：
  python 费用录入/update_feiyong_source_table.py --dry-run
  python 费用录入/update_feiyong_source_table.py

凭据：仓库根目录 .env 的 FEISHU_APP_ID / FEISHU_APP_SECRET
"""
import os
import sys
import argparse
import pathlib

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

import requests
from dotenv import load_dotenv

REPO_ROOT = pathlib.Path(__file__).resolve().parent.parent
load_dotenv(REPO_ROOT / ".env", override=False)

BITABLE_APP_TOKEN = "WJ2IbCPUgax1Phsogk0cC0eNnCd"
TABLE_NAME = "费用录入"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"

TEXT, SINGLE_SELECT = 1, 3
OLD_PRIMARY, NEW_PRIMARY = "订单号", "关联单号"
LEIXIN_OPTIONS = ["销售", "采购", "库存"]


def log(m):
    print(m, flush=True)


def token_and_session():
    app_id = os.environ.get("FEISHU_APP_ID", "").strip()
    app_secret = os.environ.get("FEISHU_APP_SECRET", "").strip()
    if not app_id or not app_secret:
        raise SystemExit("[FAIL] 缺少 FEISHU_APP_ID / FEISHU_APP_SECRET（请补进仓库根 .env）")
    s = requests.Session()
    d = s.post(f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
               json={"app_id": app_id, "app_secret": app_secret}, timeout=20).json()
    if d.get("code") != 0:
        raise SystemExit(f"[FAIL] 取 token 失败: {d.get('code')} {d.get('msg')}")
    return s, d["tenant_access_token"]


def find_table(s, h) -> str:
    d = s.get(f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables",
              headers=h, params={"page_size": 100}, timeout=30).json()
    if d.get("code") != 0:
        raise SystemExit(f"[FAIL] 读表列表失败: {d}")
    for t in d["data"].get("items", []):
        if t["name"] == TABLE_NAME:
            return t["table_id"]
    raise SystemExit(f"[FAIL] base 里没有「{TABLE_NAME}」表")


def list_fields(s, h, table_id) -> list:
    d = s.get(f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{table_id}/fields",
              headers=h, params={"page_size": 200}, timeout=30).json()
    if d.get("code") != 0:
        raise SystemExit(f"[FAIL] 读字段失败: {d}")
    return d["data"].get("items", [])


def main() -> int:
    p = argparse.ArgumentParser(description="改造费用录入源表")
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args()

    s, tok = token_and_session()
    h = {"Authorization": f"Bearer {tok}", "Content-Type": "application/json"}
    table_id = find_table(s, h)
    log(f"[源表] 「{TABLE_NAME}」 → {table_id}")

    fields = list_fields(s, h, table_id)
    by_name = {f["field_name"]: f for f in fields}
    log(f"[源表] 现有 {len(fields)} 列: {', '.join(by_name)}")

    base_url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{table_id}/fields"
    plan = []

    if NEW_PRIMARY in by_name:
        log(f"[改名] 「{NEW_PRIMARY}」已存在，跳过")
    elif OLD_PRIMARY in by_name:
        plan.append(("rename", by_name[OLD_PRIMARY]["field_id"], OLD_PRIMARY, NEW_PRIMARY))
    else:
        raise SystemExit(f"[FAIL] 既没有「{OLD_PRIMARY}」也没有「{NEW_PRIMARY}」，无法确定主字段")

    if "费用类型" in by_name:
        log("[新增] 「费用类型」已存在，跳过")
    else:
        plan.append(("add", None, "费用类型", SINGLE_SELECT))
    for name in ("捆包号", "牌号"):
        if name in by_name:
            log(f"[新增] 「{name}」已存在，跳过")
        else:
            plan.append(("add", None, name, TEXT))
    for name in ("懂火费用编号",):
        if name in by_name:
            log(f"[新增] 「{name}」已存在，跳过")
        else:
            plan.append(("add", None, name, TEXT))

    if not plan:
        log("[完成] 无需改动")
        return 0

    log("\n[计划]")
    for kind, fid, old, new in plan:
        log(f"  {kind:<7} {old} → {new}" if kind == "rename" else f"  {kind:<7} {old} (type={new})")
    if args.dry_run:
        log("\n[dry-run] 未实际改动")
        return 0

    for kind, fid, old, new in plan:
        if kind == "rename":
            r = s.put(f"{base_url}/{fid}", headers=h,
                      json={"field_name": new, "type": TEXT}, timeout=30).json()
        else:
            body = {"field_name": old, "type": new}
            if new == SINGLE_SELECT:
                body["property"] = {"options": [{"name": n} for n in LEIXIN_OPTIONS]}
            r = s.post(base_url, headers=h, json=body, timeout=30).json()
        if r.get("code") != 0:
            raise SystemExit(f"[FAIL] {kind}「{old}」失败: {r.get('code')} {r.get('msg')}")
        log(f"  [OK] {kind} {old} → {new}")

    got = list_fields(s, h, table_id)
    log(f"\n[回读校验] {len(got)} 列: {', '.join(f['field_name'] for f in got)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
