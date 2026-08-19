#!/usr/bin/env python3
"""Inspect Feishu Bitable tblvugnoJPS8GrpX fields + sample records."""
import os, json
import requests

APP_TOKEN = "Tz0XbQVzkaZuJasBwb8cRjkfnoe"
TABLE_ID = "tblvugnoJPS8GrpX"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}

def env(name, default=""):
    return os.environ.get(name, default).strip()

def get_token():
    r = requests.post(f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
                      json={"app_id": env("FEISHU_APP_ID"), "app_secret": env("FEISHU_APP_SECRET")},
                      timeout=15, proxies=NO_PROXY)
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"token failed: {d}")
    return d["tenant_access_token"]

t = get_token()
print(f"token OK len={len(t)}")

# 1) list fields
r = requests.get(
    f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/fields",
    headers={"Authorization": f"Bearer {t}"},
    params={"page_size": 100},
    timeout=15, proxies=NO_PROXY)
print(f"\n[fields] status={r.status_code}")
d = r.json()
if d.get("code") != 0:
    print(json.dumps(d, ensure_ascii=False, indent=2)[:2000])
    raise SystemExit(1)
fields = (d.get("data") or {}).get("items") or []
print(f"共 {len(fields)} 个字段:")
for i, f in enumerate(fields, 1):
    fid = f.get("field_id")
    fname = f.get("field_name")
    ftype = f.get("type")
    ftyp_name = f.get("ui_type") or ftype
    # 单选/多选 选项
    prop = f.get("property") or {}
    opts = (prop.get("options") if isinstance(prop, dict) else []) or []
    opt_str = ""
    if opts:
        vals = [o.get("name") for o in opts]
        opt_str = f"  options={vals}"
    print(f"  [{i:2d}] id={fid}  name={fname!r}  type={ftyp_name} ({ftype}){opt_str}")

# 2) list 5 records preview (see values)
r2 = requests.get(
    f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records",
    headers={"Authorization": f"Bearer {t}"},
    params={"page_size": 5},
    timeout=15, proxies=NO_PROXY)
d2 = r2.json()
print(f"\n[records preview] code={d2.get('code')}  msg={d2.get('msg','')}")
items = (d2.get("data") or {}).get("items") or []
print(f"共 {len(items)} 条（预览 5）")
for i, rec in enumerate(items, 1):
    rid = rec.get("record_id")
    fields0 = rec.get("fields") or {}
    print(f"\n  rec[{i}] id={rid}")
    for k, v in fields0.items():
        # 简化值展示
        if isinstance(v, list) and len(v) > 0 and isinstance(v[0], dict):
            # link/lookup 等对象列表
            first = v[0]
            if "text" in first:
                vv = [x.get("text","") for x in v]
            else:
                vv = f"<link/map len={len(v)}>"
        else:
            vv = v
        print(f"    {k!r}: {vv!r}")
