#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""查询'滚动'表的视图列表，并用 view_id 获取'放入'视图的字段顺序"""

import os
import csv
import requests

APP_TOKEN = "Mw62b7vijaFW8VsVAwXcncNEnbc"
TABLE_ID = "tblfqRNfFg3NUcTo"  # '滚动'表
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"


def env(name, default=""):
    return os.environ.get(name, default).strip()


def get_token():
    app_id = env("FEISHU_APP_ID")
    app_secret = env("FEISHU_APP_SECRET")
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": app_id, "app_secret": app_secret}, timeout=15)
    data = r.json()
    return data["tenant_access_token"]


def list_views(token):
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/views"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=15)
    data = r.json()
    if data.get("code") != 0:
        print(f"[错误] 获取视图列表失败: {data}")
        return []
    items = (data.get("data") or {}).get("items") or []
    print(f"\n=== '滚动'表视图列表 (共 {len(items)} 个) ===")
    for v in items:
        print(f"  view_id={v.get('view_id')}  name={v.get('view_name')!r}  type={v.get('view_type')}")
    return items


def list_fields_with_view(token, view_id=None):
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/fields"
    params = {}
    if view_id:
        params["view_id"] = view_id
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params, timeout=15)
    data = r.json()
    if data.get("code") != 0:
        print(f"[错误] 获取字段失败: {data}")
        return []
    return (data.get("data") or {}).get("items") or []


def main():
    token = get_token()

    # 1. 获取视图列表
    views = list_views(token)

    # 2. 找"放入"视图
    target_view = None
    for v in views:
        if "放入" in (v.get("view_name") or ""):
            target_view = v
            break

    if not target_view:
        print("\n[错误] 未找到名称含'放入'的视图")
        return 1

    view_id = target_view.get("view_id")
    print(f"\n[找到] '放入'视图: view_id={view_id}")

    # 3. 用 view_id 获取该视图的字段顺序
    fields = list_fields_with_view(token, view_id)
    print(f"\n=== '放入'视图字段顺序 (共 {len(fields)} 个) ===")
    for i, f in enumerate(fields):
        print(f"  [{i}] field_id={f.get('field_id')}  name={f.get('field_name')!r}  type={f.get('type')}  ui_type={f.get('ui_type')!r}")

    # 4. 找"序号"字段的位置
    xuhao_idx = None
    for i, f in enumerate(fields):
        if f.get("field_name") == "序号":
            xuhao_idx = i
            break

    if xuhao_idx is None:
        print("\n[错误] '放入'视图中未找到'序号'字段")
        return 1

    print(f"\n[定位] '序号'字段在'放入'视图中的位置: [{xuhao_idx}]")

    # 5. 从"序号"开始往右的字段，按顺序对应 CSV 的17个字段
    csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "kaipiao_daiqueren_20260814_0040.csv")
    with open(csv_path, encoding="utf-8-sig") as fp:
        csv_fields = next(csv.reader(fp))

    print(f"\n=== 映射关系 (从'序号'开始往右 → CSV字段) ===")
    view_fields_from_xuhao = fields[xuhao_idx:]
    print(f"  视图从'序号'开始有 {len(view_fields_from_xuhao)} 个字段")
    print(f"  CSV 有 {len(csv_fields)} 个字段")
    print()
    max_len = max(len(view_fields_from_xuhao), len(csv_fields))
    for i in range(max_len):
        vf = view_fields_from_xuhao[i] if i < len(view_fields_from_xuhao) else None
        cf = csv_fields[i] if i < len(csv_fields) else None
        vf_str = f"{vf.get('field_name')!r} (id={vf.get('field_id')})" if vf else "(无)"
        cf_str = f"{cf!r}" if cf else "(无)"
        print(f"  [{i}] 视图字段: {vf_str}  ←  CSV字段: {cf_str}")

    return 0


if __name__ == "__main__":
    exit(main())
