#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""将CSV中8条待确认开票申请单追加到飞书多维表格'滚动'表"""

import os
import csv
import json
import time
from datetime import datetime
import requests

APP_TOKEN = "Mw62b7vijaFW8VsVAwXcncNEnbc"
TABLE_ID = "tblfqRNfFg3NUcTo"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"

# CSV字段 → (飞书字段名, 字段类型)
FIELD_MAPPING = {
    "id":          ("序号", "number"),       # 序号
    "状态":         ("2", "text"),           # 字段名"2"
    "申请单号":      ("申请单号", "text"),     # 申请单号
    "开票日期":      ("3", "text"),           # 字段名"3"
    "发票日期":      ("申请时间", "datetime"), # 申请时间
    "我方名称":      ("1", "text"),           # 字段名"1"
    "结算对方":      ("4", "text"),           # 字段名"4"
    "发票对方":      ("发票抬头", "text"),     # 发票抬头
    "发票数量":      ("重量", "number"),       # 重量
    "发票金额":      ("金额", "number"),       # 金额
    "发票号码":      ("5", "text"),           # 字段名"5"
    "备注":         ("备注", "text"),         # 备注
    "业务人员":      ("6", "text"),           # 字段名"6"
    "经办人":        ("提交人", "text"),       # 提交人
    "所属公司":      ("所属公司", "text"),     # 所属公司
    "新增时间":      ("排票日期", "datetime"), # 排票日期
}


def env(name, default=""):
    return os.environ.get(name, default).strip()


def get_token():
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={
        "app_id": env("FEISHU_APP_ID"),
        "app_secret": env("FEISHU_APP_SECRET"),
    }, timeout=15)
    return r.json()["tenant_access_token"]


def to_text(val):
    """转飞书Text格式（直接传字符串）"""
    s = str(val).strip() if val else ""
    if not s:
        return None
    return s


def to_number(val):
    """转数字"""
    s = str(val).strip() if val else ""
    if not s:
        return None
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except ValueError:
        return None


def to_datetime(val):
    """转毫秒时间戳"""
    s = str(val).strip() if val else ""
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
        try:
            dt = datetime.strptime(s, fmt)
            return int(dt.timestamp() * 1000)
        except ValueError:
            continue
    return None


def convert_value(csv_val, field_type):
    if field_type == "text":
        return to_text(csv_val)
    elif field_type == "number":
        return to_number(csv_val)
    elif field_type == "datetime":
        return to_datetime(csv_val)
    return None


def batch_create(token, records):
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records/batch_create"
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"},
                      json={"records": records}, timeout=30)
    data = r.json()
    if data.get("code") != 0:
        print(f"[错误] 批量创建失败: {json.dumps(data, ensure_ascii=False)[:1000]}")
        return None
    return data.get("data") or {}


def main():
    csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "kaipiao_daiqueren_20260814_0040.csv")
    with open(csv_path, encoding="utf-8-sig") as fp:
        rows = list(csv.DictReader(fp))

    print(f"[读取CSV] {len(rows)} 条记录")

    token = get_token()

    records = []
    for i, row in enumerate(rows):
        fields = {}
        for csv_col, (field_name, ftype) in FIELD_MAPPING.items():
            val = convert_value(row.get(csv_col, ""), ftype)
            if val is not None:
                fields[field_name] = val
        records.append({"fields": fields})
        print(f"  行{i+1}: 申请单号={row.get('申请单号','')} → {len(fields)}个字段")

    print(f"\n[写入] 开始批量创建 {len(records)} 条记录...")
    result = batch_create(token, records)
    if result:
        created = result.get("records") or []
        print(f"[完成] 成功创建 {len(created)} 条记录")
        for r in created:
            rid = r.get("record_id", "")
            print(f"  record_id={rid}")
    else:
        print("[失败] 写入失败")
        return 1

    return 0


if __name__ == "__main__":
    exit(main())
