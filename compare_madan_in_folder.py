#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IEC 码单云盘比对 → 增量 CSV（阶段2 前置：2a 下载 + 2b 比对）
=============================================================

职责：
  从飞书云盘文件夹下载最新两份"码单_YYMMDD_HHMMSS.xlsx"，
  以"捆包号"为唯一键比对，输出新增捆包 CSV 供 iec_bundle-to-ruku.py 写入 Bitable。

输入：
  飞书云盘文件夹（默认 DfQdfSxl2ld25wdx6Rxcub9hnDf，可 --folder-token 覆盖）
  文件名规则：码单_YYMMDD_HHMMSS.xlsx （2 位年份 + 秒级时间戳，由 export_iec_bundle.py 产出）

输出：
  CSV（默认 比对_最新多出记录.csv，可 --csv 覆盖）
  列结构同 xlsx 原列名 + 末列"比对结果"="新增"
  编码 UTF-8-SIG（Excel 直接打开不乱码）

环境变量：
  FEISHU_APP_ID / FEISHU_APP_SECRET  飞书自建应用（需 drive:drive + 目标文件夹可编辑协作者）

命令行：
  python compare_madan_in_folder.py                        # 默认文件夹 + 默认 CSV 名
  python compare_madan_in_folder.py --csv out.csv         # 指定输出 CSV
  python compare_madan_in_folder.py --folder-token XXX    # 指定云盘文件夹
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
import requests

try:
    import pandas as pd
except ImportError:
    print("❌ 缺少 pandas，请先 pip install pandas openpyxl"); sys.exit(1)


FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}
DEFAULT_FOLDER_TOKEN = "DfQdfSxl2ld25wdx6Rxcub9hnDf"
DEFAULT_CSV = "比对_最新多出记录.csv"
FILE_PREFIX = "码单_"
FILE_SUFFIX = ".xlsx"

# 文件名时间戳排序 key：码单_YYMMDD_HHMMSS.xlsx → (YYMMDD, HHMMSS)
NAME_TS_RE = re.compile(rf"{re.escape(FILE_PREFIX)}(\d{{6}})_(\d{{6}}){re.escape(FILE_SUFFIX)}$")


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


def list_folder_xlsx(tok: str, folder_token: str) -> list[dict]:
    """列出文件夹内所有 码单_*.xlsx 文件，按文件名时间戳倒序返回"""
    items = []
    page_token = None
    while True:
        params = {"folder_token": folder_token, "page_size": 50}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(f"{FEISHU_OPEN_BASE}/drive/v1/files",
                         headers={"Authorization": f"Bearer {tok}"},
                         params=params, timeout=20, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") != 0:
            print(f"❌ 列文件失败: {json.dumps(d, ensure_ascii=False)[:300]}"); sys.exit(1)
        files = (d.get("data") or {}).get("files") or []
        for f in files:
            name = f.get("name", "")
            if name.startswith(FILE_PREFIX) and name.endswith(FILE_SUFFIX):
                items.append(f)
        if not (d.get("data") or {}).get("has_more"):
            break
        page_token = (d.get("data") or {}).get("page_token")
        if not page_token:
            break

    # 按文件名内嵌时间戳倒序排序（最新在前）
    def sort_key(f):
        m = NAME_TS_RE.match(f.get("name", ""))
        if m:
            return (m.group(1), m.group(2))
        return ("0", "0")
    items.sort(key=sort_key, reverse=True)
    return items


def download_file(tok: str, file_token: str, save_path: str) -> int:
    """下载飞书云盘文件到本地，返回字节数"""
    url = f"{FEISHU_OPEN_BASE}/drive/v1/files/{file_token}/download"
    r = requests.get(url,
                     headers={"Authorization": f"Bearer {tok}"},
                     timeout=120, proxies=NO_PROXY)
    if r.status_code != 200:
        print(f"❌ 下载失败 HTTP {r.status_code}: {r.text[:200]}"); sys.exit(1)
    content = r.content
    with open(save_path, "wb") as f:
        f.write(content)
    return len(content)


def compare_and_write_csv(latest_path: str, prev_path: str, csv_path: str) -> int:
    """以'捆包号'为唯一键，找出 latest 有但 prev 没有的记录，写 CSV"""
    df_new = pd.read_excel(latest_path, dtype=str).fillna("")
    df_old = pd.read_excel(prev_path, dtype=str).fillna("")

    if "捆包号" not in df_new.columns or "捆包号" not in df_old.columns:
        print(f"❌ xlsx 缺少'捆包号'列。最新列: {list(df_new.columns)}")
        sys.exit(1)

    new_ids = set(df_new["捆包号"].astype(str).str.strip())
    old_ids = set(df_old["捆包号"].astype(str).str.strip())
    added_ids = new_ids - old_ids

    df_added = df_new[df_new["捆包号"].astype(str).str.strip().isin(added_ids)].copy()
    df_added["比对结果"] = "新增"

    df_added.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"\n✅ 比对完成: 最新 {len(new_ids)} 条 / 次新 {len(old_ids)} 条 / 新增 {len(added_ids)} 条")
    print(f"   CSV: {csv_path}  ({os.path.getsize(csv_path)} 字节)")
    return len(added_ids)


def main():
    parser = argparse.ArgumentParser(description="IEC 码单云盘比对 → 增量 CSV")
    parser.add_argument("--folder-token", default=DEFAULT_FOLDER_TOKEN,
                        help=f"飞书云盘文件夹 token（默认 {DEFAULT_FOLDER_TOKEN}）")
    parser.add_argument("--csv", default=DEFAULT_CSV,
                        help=f"输出 CSV 路径（默认 {DEFAULT_CSV}）")
    args = parser.parse_args()

    print("=" * 72)
    print(f"码单云盘比对 → CSV")
    print(f"  文件夹: {args.folder_token}")
    print(f"  输出 CSV: {args.csv}")
    print("=" * 72)

    tok = get_token()
    print("✅ 飞书 token 获取成功")

    files = list_folder_xlsx(tok, args.folder_token)
    print(f"\n[列文件] 找到 {len(files)} 个 码单_*.xlsx")
    for i, f in enumerate(files[:5], 1):
        print(f"  {i}. {f.get('name')}  token={f.get('token', '')[:10]}...")
    if len(files) > 5:
        print(f"  ... (共 {len(files)} 个)")

    if len(files) < 2:
        print(f"\n❌ 需要至少 2 个码单文件来比对，当前只有 {len(files)} 个")
        print("   请先运行 export_iec_bundle.py 至少两次（产生两次码单）")
        sys.exit(1)

    latest = files[0]
    prev = files[1]
    print(f"\n[选定]")
    print(f"  最新: {latest.get('name')}")
    print(f"  次新: {prev.get('name')}")

    tmp_dir = "_madan_tmp"
    os.makedirs(tmp_dir, exist_ok=True)
    latest_path = os.path.join(tmp_dir, latest["name"])
    prev_path = os.path.join(tmp_dir, prev["name"])

    print(f"\n[下载]")
    size1 = download_file(tok, latest["token"], latest_path)
    print(f"  最新: {latest['name']}  ({size1:,} 字节)")
    size2 = download_file(tok, prev["token"], prev_path)
    print(f"  次新: {prev['name']}  ({size2:,} 字节)")

    print(f"\n[比对]")
    added = compare_and_write_csv(latest_path, prev_path, args.csv)

    # 清理临时文件
    for p in (latest_path, prev_path):
        try: os.remove(p)
        except OSError: pass
    try: os.rmdir(tmp_dir)
    except OSError: pass

    if added == 0:
        print("\n⚠️ 本次无新增捆包，CSV 为空表头")
    print("\n下一步: python iec_bundle-to-ruku.py --csv {} --commit".format(args.csv))


if __name__ == "__main__":
    main()
