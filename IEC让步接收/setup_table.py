#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IEC 让步接收 · 建表脚本（幂等）
- 在 IEC 业务台账 base 建「让步接收单」表（含全部字段），已存在则跳过
- 飞书应用身份（.env: FEISHU_APP_ID / FEISHU_APP_SECRET）
"""
from __future__ import annotations

import json
import os
import sys

import requests

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
NO_PROXY = {"http": None, "https": None}

APP_TOKEN = "Tz0XbQVzkaZuJasBwb8cRjkfnoe"   # 综合整理——环月（新）
TABLE_NAME = "让步接收单"
OPEN = "https://open.feishu.cn/open-apis"

FIELDS = [
    {"field_name": "让步申请单号", "type": 1},
    {"field_name": "钢厂订单号", "type": 1},
    {"field_name": "变更原因", "type": 1},
    {"field_name": "变更描述", "type": 1},
    {"field_name": "规格", "type": 1},
    {"field_name": "牌号", "type": 1},
    {"field_name": "品种", "type": 1},
    {"field_name": "制造单元", "type": 1},
    {"field_name": "交货月", "type": 1},
    {"field_name": "客户回复", "type": 3,
     "property": {"options": [{"name": "是"}, {"name": "否"}]}},
    {"field_name": "地区公司回复", "type": 3,
     "property": {"options": [{"name": "是"}, {"name": "否"}]}},
    {"field_name": "业务部门回复", "type": 3,
     "property": {"options": [{"name": "是"}, {"name": "否"}]}},
    {"field_name": "最终用户", "type": 1},
    {"field_name": "首次抓取时间", "type": 5},
    {"field_name": "最后同步时间", "type": 5},
]


def env(name: str) -> str:
    p = os.path.join(ROOT, ".env")
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                if k.strip() == name:
                    return v.strip()
    return os.environ.get(name, "").strip()


def token() -> str:
    # SH Robot（本机 .env: FEISHU_NOTIFY_APP_*，CI: FEISHU_APP_ID/SECRET 同一应用）
    r = requests.post(f"{OPEN}/auth/v3/tenant_access_token/internal",
                      json={"app_id": env("FEISHU_NOTIFY_APP_ID") or env("FEISHU_APP_ID"),
                            "app_secret": env("FEISHU_NOTIFY_APP_SECRET") or env("FEISHU_APP_SECRET")},
                      timeout=20, proxies=NO_PROXY)
    d = r.json()
    if not d.get("tenant_access_token"):
        raise RuntimeError(f"取 token 失败: {d}")
    return d["tenant_access_token"]


def main() -> int:
    tk = token()
    H = {"Authorization": f"Bearer {tk}"}

    # 1) 查现有表
    r = requests.get(f"{OPEN}/bitable/v1/apps/{APP_TOKEN}/tables",
                     params={"page_size": 100}, headers=H, timeout=20, proxies=NO_PROXY)
    d = r.json()
    print(f"[查表] code={d.get('code')} msg={d.get('msg')} 表数={len((d.get('data') or {}).get('items') or [])}")
    if d.get("code") != 0:
        print("[FAIL] 读表清单失败 —— 应用可能没有该 base 的权限")
        print("       需要把应用 cli_aa2e9605e8389ce9 加为该多维表格的可编辑协作者")
        return 2
    for t in (d["data"]["items"] or []):
        if t["name"] == TABLE_NAME:
            print(f"[跳过] 表已存在: {TABLE_NAME} ({t['table_id']})")
            return 0

    # 2) 建表
    r = requests.post(f"{OPEN}/bitable/v1/apps/{APP_TOKEN}/tables", headers=H, timeout=30,
                      json={"table": {"name": TABLE_NAME, "default_view_name": "全部", "fields": FIELDS}},
                      proxies=NO_PROXY)
    d = r.json()
    print(f"[建表] code={d.get('code')} msg={d.get('msg')}")
    if d.get("code") != 0:
        print(json.dumps(d, ensure_ascii=False)[:500])
        return 3
    tid = d["data"]["table_id"]
    print(f"[OK] 已建表 {TABLE_NAME} → {tid}")

    # 3) 回读验证字段
    r = requests.get(f"{OPEN}/bitable/v1/apps/{APP_TOKEN}/tables/{tid}/fields",
                     params={"page_size": 100}, headers=H, timeout=20, proxies=NO_PROXY)
    d = r.json()
    print(f"[回读] 字段数={len((d.get('data') or {}).get('items') or [])}")
    for f in (d.get("data") or {}).get("items") or []:
        props = f.get("property") or {}
        extra = ""
        if f.get("type") == 3:
            extra = " 选项=" + "/".join(o["name"] for o in props.get("options", []))
        print(f"   {f['field_name']:<14} type={f['type']}{extra}")

    p = os.path.join(HERE, "_work", "table_id.txt")
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as fh:
        fh.write(tid)
    print(f"[OK] table_id 已存 {p}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
