#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
重建飞书「费用录入」源表（按理想列顺序）

飞书 API 不支持调整列顺序，所以要改顺序只能删表重建。
⚠️ 安全闸：表里只要有 1 条记录就拒绝删除并退出，绝不会拿已有数据冒险。

列顺序（18 列）：
  懂火费用编号* | 关联单号 费用类型 捆包号 牌号 | 结算日期 费用名称 服务商 税率 计价方式
  计价重量 单价 金额 备注说明 | 写入状态 写入时间 懂火记录号 写入结果
  （* 为主字段，必须是文本；竖线只是分组示意）

⚠️ 主字段 = 第一列 = 「懂火费用编号」。新录入的行在写入懂火之前这一列是空的
   （编号要写入后才由 ERP 生成），这是主字段的必然结果。

「懂火记录号」= ERP 内部 id（程序定位/删除用）；
「懂火费用编号」= ERP 里的 F 开头单号（人查单用）。

下拉选项全部来自懂火主数据：
  费用名称 ← /model/admin/m_load/getlist?leixin=费用名称
  服务商   ← /model/admin/m_load/list_fuwu
  税率 / 计价方式 / 费用类型 / 写入状态 ← 懂火表单里写死的枚举

用法（从仓库根目录运行）：
  python 费用录入/rebuild_feiyong_source_table.py --dry-run
  python 费用录入/rebuild_feiyong_source_table.py

凭据：仓库根目录 .env 的 FEISHU_APP_ID / FEISHU_APP_SECRET
      DH_USERNAME / DH_PASSWORD（拉费用名称、服务商选项）
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
sys.path.insert(0, str(REPO_ROOT))
from donghuo_login import login_donghuo, BASE_URL  # noqa: E402

BITABLE_APP_TOKEN = "WJ2IbCPUgax1Phsogk0cC0eNnCd"
TABLE_NAME = "费用录入"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
AJAX = {"X-Requested-With": "XMLHttpRequest"}

TEXT, NUMBER, SINGLE_SELECT, DATETIME = 1, 2, 3, 5

FEE_TYPE_OPTIONS = ["销售", "采购", "库存"]
SHUILV_OPTIONS = ["0", "0.03", "0.06", "0.09", "0.11", "0.13"]
JIJIA_OPTIONS = ["重量", "整车", "整单"]
STATUS_OPTIONS = ["待写入", "写入中", "已写入", "失败"]

APP_ID = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}"


def log(m):
    print(m, flush=True)


def token(s):
    aid = os.environ.get("FEISHU_APP_ID", "").strip()
    sec = os.environ.get("FEISHU_APP_SECRET", "").strip()
    if not aid or not sec:
        raise SystemExit("[FAIL] 缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")
    d = s.post(f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
               json={"app_id": aid, "app_secret": sec}, timeout=20).json()
    if d.get("code") != 0:
        raise SystemExit(f"[FAIL] token: {d}")
    return d["tenant_access_token"]


def _load(session, path, data):
    items = session.post(f"{BASE_URL}{path}", data=data, headers=AJAX, timeout=30).json()
    names = [str(it.get("value") or it.get("key")).strip()
             for it in items if isinstance(it, dict)]
    return [n for n in dict.fromkeys(names) if n]


def fetch_options():
    """一次登录，拉「费用名称」和「服务商」两套主数据（顺序与 ERP 下拉一致）"""
    sess = login_donghuo()
    if sess is None:
        raise SystemExit("[FAIL] 懂火登录失败，无法拉主数据选项")
    fee_names = _load(sess, "/model/admin/m_load/getlist", {"leixin": "费用名称"})
    fuwu = _load(sess, "/model/admin/m_load/list_fuwu", {})
    return fee_names, fuwu


def build_fields(fee_names, fuwu):
    return [
        {"field_name": "懂火费用编号", "type": TEXT},
        {"field_name": "关联单号", "type": TEXT},
        {"field_name": "费用类型", "type": SINGLE_SELECT,
         "property": {"options": [{"name": n} for n in FEE_TYPE_OPTIONS]}},
        {"field_name": "捆包号", "type": TEXT},
        {"field_name": "牌号", "type": TEXT},
        {"field_name": "结算日期", "type": DATETIME},
        {"field_name": "费用名称", "type": SINGLE_SELECT,
         "property": {"options": [{"name": n} for n in fee_names]}},
        {"field_name": "服务商", "type": SINGLE_SELECT,
         "property": {"options": [{"name": n} for n in fuwu]}},
        {"field_name": "税率", "type": SINGLE_SELECT,
         "property": {"options": [{"name": n} for n in SHUILV_OPTIONS]}},
        {"field_name": "计价方式", "type": SINGLE_SELECT,
         "property": {"options": [{"name": n} for n in JIJIA_OPTIONS]}},
        {"field_name": "计价重量", "type": NUMBER},
        {"field_name": "单价", "type": NUMBER},
        {"field_name": "金额", "type": NUMBER},
        {"field_name": "备注说明", "type": TEXT},
        {"field_name": "写入状态", "type": SINGLE_SELECT,
         "property": {"options": [{"name": n} for n in STATUS_OPTIONS]}},
        {"field_name": "写入时间", "type": DATETIME},
        {"field_name": "懂火记录号", "type": TEXT},
        {"field_name": "写入结果", "type": TEXT},
    ]


def main():
    p = argparse.ArgumentParser(description="重建费用录入源表")
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args()

    s = requests.Session()
    tok = token(s)
    h = {"Authorization": f"Bearer {tok}", "Content-Type": "application/json"}
    log("[飞书] [OK] 已取得 tenant_access_token")

    d = s.get(f"{APP_ID}/tables", headers=h, params={"page_size": 100}, timeout=30).json()
    if d.get("code") != 0:
        raise SystemExit(f"[FAIL] 读表列表失败: {d}")
    tables = {t["name"]: t["table_id"] for t in d["data"].get("items", [])}
    log(f"[源表] 当前 base 下的表: {', '.join(tables) if tables else '（无）'}")
    old_id = tables.get(TABLE_NAME)

    if old_id:
        rd = s.get(f"{APP_ID}/tables/{old_id}/records",
                   headers=h, params={"page_size": 1}, timeout=30).json()
        if rd.get("code") != 0:
            raise SystemExit(f"[FAIL] 读记录失败: {rd}")
        total = rd["data"].get("total")
        log(f"[源表] 旧表「{TABLE_NAME}」({old_id}) 现有 {total} 条记录")
        if total:
            raise SystemExit(f"[安全闸] 旧表里有 {total} 条记录，拒绝删除。"
                             f"请先自行备份/清空，或改用增量加列的方式。")

    fee_names, fuwu = fetch_options()
    log(f"[懂火] [OK] 费用名称选项 {len(fee_names)} 项；服务商选项 {len(fuwu)} 项")
    fields = build_fields(fee_names, fuwu)
    log(f"\n[重建] 新表 {len(fields)} 列，顺序：")
    for i, f in enumerate(fields, 1):
        opts = (f.get("property") or {}).get("options")
        if opts and len(opts) > 8:
            tail = f"  选项 {len(opts)} 项（{', '.join(o['name'] for o in opts[:6])} …）"
        elif opts:
            tail = f"  选项={', '.join(o['name'] for o in opts)}"
        else:
            tail = ""
        log(f"  {i:>2}. {f['field_name']:<8} type={f['type']}{tail}")

    if args.dry_run:
        log("\n[dry-run] 未实际改动")
        return 0

    if old_id:
        dd = s.delete(f"{APP_ID}/tables/{old_id}", headers=h, timeout=30).json()
        if dd.get("code") != 0:
            raise SystemExit(f"[FAIL] 删除旧表失败: {dd.get('code')} {dd.get('msg')}")
        log(f"[重建] 已删除空表 {old_id}")

    body = {"table": {"name": TABLE_NAME, "default_view_name": "表格", "fields": fields}}
    cd = s.post(f"{APP_ID}/tables", headers=h, json=body, timeout=60).json()
    if cd.get("code") != 0:
        raise SystemExit(f"[FAIL] 建表失败: {cd.get('code')} {cd.get('msg')}")
    new_id = cd["data"]["table_id"]
    log(f"[重建] [OK] 新表 {new_id}")

    got = s.get(f"{APP_ID}/tables/{new_id}/fields", headers=h,
                params={"page_size": 200}, timeout=30).json()["data"]["items"]
    order = [f["field_name"] for f in got]
    log(f"[回读校验] {len(order)} 列: {', '.join(order)}")
    missing = [f["field_name"] for f in fields if f["field_name"] not in order]
    if missing:
        log(f"[WARN] 缺少: {missing}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
