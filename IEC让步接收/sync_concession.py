#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IEC 让步接收单 → 飞书多维表格 同步
=====================================
数据来源：IEC 订单管理 → 让步接收单
    URL   : POST /iecs/order/concession/contractConcessionReceive/queryConcessionReceiveList
    返回  : HTML tbody 片段（<input id="total"> + <tr> 行）
    行结构: td[0]让步申请单号 td[1]钢厂订单号 td[2]变更原因 td[3]变更描述
            td[4]规格 td[5]牌号 td[6]品种 td[7]制造单元 td[8]交货月
            td[9]客户回复标记 td[10]地区公司回复标记 td[11]业务部门回复标记
            td[12]最终用户 td[13]操作(详情按钮)

目标表：IEC 业务台账 base（综合整理——环月（新））→「让步接收单」表
    以「让步申请单号」为唯一键做差分：
      · IEC 有、表里没有        → 新增记录（首次抓取时间 = 现在）
      · IEC 有、表里有、回复标记有变化 → 更新（最后同步时间 = 现在）
      · 其余不动

通知：由飞书多维表格「自动化」负责（记录新增时触发），本脚本不发消息。

环境变量（.env）：
    IBAO_USERNAME / IBAO_PASSWORD   IEC 登录
    FEISHU_APP_ID / FEISHU_APP_SECRET  飞书自建应用（需为本表的可编辑协作者）

参数：
    --dry-run        只打印差异，不写飞书
    --from YYYYMM    IEC 创建时间下限（默认 202001）
    --to   YYYYMM    IEC 创建时间上限（默认 203012）
    --verbose        打印每行明细
"""
from __future__ import annotations

import argparse
import html as html_mod
import os
import re
import sys
import time
from typing import Optional

import requests

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
from ibaosteel_client import IEC  # noqa: E402

NO_PROXY = {"http": None, "https": None}
BASE = "https://www.ibaosteel.com"
IECS = f"{BASE}/iecs"
PAGE = f"{IECS}/order/concession/contractConcessionReceive/initLoads"
QUERY = f"{IECS}/order/concession/contractConcessionReceive/queryConcessionReceiveList"

OPEN = "https://open.feishu.cn/open-apis"
BITABLE_APP = "Tz0XbQVzkaZuJasBwb8cRjkfnoe"      # 综合整理——环月（新）
BITABLE_TABLE_NAME = "让步接收单"

# IEC 列 → 飞书字段
COLS = ["让步申请单号", "钢厂订单号", "变更原因", "变更描述", "规格", "牌号", "品种",
        "制造单元", "交货月", "客户回复", "地区公司回复", "业务部门回复", "最终用户"]
IDX = {name: i for i, name in enumerate(COLS)}
REPLY_FIELDS = ("客户回复", "地区公司回复", "业务部门回复")


def env(name: str, default: str = "") -> str:
    v = os.environ.get(name)
    if v:
        return v.strip()
    p = os.path.join(ROOT, ".env")
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, val = line.split("=", 1)
                if k.strip() == name:
                    return val.strip()
    return default


# ============================================================
# IEC 侧
# ============================================================
def iec_login() -> requests.Session:
    iec = IEC(username=env("IBAO_USERNAME"), password=env("IBAO_PASSWORD"), retries=10)
    if not iec.login():
        raise RuntimeError(f"IEC 登录失败: {iec.last_error}")
    s = iec.session
    token = iec.token
    s.get(f"{BASE}/ibaosteel/bizIntelli?access_token={token}", timeout=20)
    s.get(f"{IECS}/index?token={token}", timeout=20, allow_redirects=True)
    s.get(f"{PAGE}?token={token}", timeout=30, allow_redirects=True)
    return s


def fetch_concession(s: requests.Session, ymin: str, ymax: str,
                     page_size: int = 200) -> list[dict]:
    """拉全部让步接收单（自动翻页）。返回 list[dict(COLS)]"""
    rows: list[dict] = []
    page = 1
    total = None
    while True:
        body = {
            "lldNum": "", "contractNum": "", "factoryOrderNum": "",
            "thickTbthDimMin": "", "thickTbthDimMax": "", "widthMin": "", "widthMax": "",
            "shopsign": "", "prodCodeCname": "", "prodCode": "",
            "recCreateTimeMin": ymin, "recCreateTimeMax": ymax,
            "settleUserNum": "062122", "saleNetwork": "E", "custFlag": " ",
            "pageDomain": {"pageNum": str(page), "pageSize": str(page_size)},
        }
        r = s.post(QUERY, json=body, timeout=60, headers={
            "Accept": "text/html, */*; q=0.01",
            "Content-Type": "application/json; charset=utf-8",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": PAGE,
        })
        h = r.text
        m = re.search(r'id="total"\s+value="(\d+)"', h)
        if m and total is None:
            total = int(m.group(1))
        page_rows = []
        for tr in re.findall(r"<tr\b[^>]*>(.*?)</tr>", h, re.S):
            tds = []
            for td in re.finditer(r"<td\b[^>]*>(.*?)</td>", tr, re.S):
                t = re.sub(r"<[^>]+>", " ", td.group(1))
                tds.append(html_mod.unescape(re.sub(r"\s+", " ", t)).strip())
            if len(tds) < 13:
                continue
            page_rows.append({name: tds[i] for name, i in IDX.items()})
        rows.extend(page_rows)
        if not page_rows or (total is not None and len(rows) >= total):
            break
        page += 1
        time.sleep(0.3)
    print(f"[IEC] 共 {len(rows)} 条（服务端 total={total}，{page} 页）")
    return rows


def aggregate_by_lld(rows: list[dict]) -> list[dict]:
    """按让步申请单号聚合（一个单号一行，2026-10-11 实测：回复标记在单号内一致）

    - 单值列（变更原因/描述、品种、制造单元、交货月、三个回复标记）→ 原样
    - 多值列（钢厂订单号/规格/牌号/最终用户）→ 去重保序后用 \\n 连接
    """
    order: list[str] = []
    grouped: dict[str, dict[str, list[str]]] = {}
    for r0 in rows:
        key = r0["让步申请单号"]
        if key not in grouped:
            grouped[key] = {}
            order.append(key)
        g = grouped[key]
        for col in COLS[1:]:
            v = (r0.get(col) or "").strip()
            if v and v not in g.setdefault(col, []):
                g[col].append(v)
    out = []
    for key in order:
        row = {"让步申请单号": key}
        for col in COLS[1:]:
            row[col] = "\n".join(grouped[key].get(col, []))
        out.append(row)
    return out


# ============================================================
# 飞书侧
# ============================================================
def feishu_token() -> str:
    # SH Robot（本机 .env: FEISHU_NOTIFY_APP_*，CI: FEISHU_APP_ID/SECRET 同一应用）
    r = requests.post(f"{OPEN}/auth/v3/tenant_access_token/internal", timeout=20, proxies=NO_PROXY,
                      json={"app_id": env("FEISHU_NOTIFY_APP_ID") or env("FEISHU_APP_ID"),
                            "app_secret": env("FEISHU_NOTIFY_APP_SECRET") or env("FEISHU_APP_SECRET")})
    d = r.json()
    if not d.get("tenant_access_token"):
        raise RuntimeError(f"取飞书 token 失败: {d}")
    return d["tenant_access_token"]


def find_table(tk: str) -> Optional[str]:
    H = {"Authorization": f"Bearer {tk}"}
    r = requests.get(f"{OPEN}/bitable/v1/apps/{BITABLE_APP}/tables",
                     params={"page_size": 100}, headers=H, timeout=20, proxies=NO_PROXY)
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"读表清单失败: {d.get('code')} {d.get('msg')}")
    for t in (d.get("data") or {}).get("items") or []:
        if t["name"] == BITABLE_TABLE_NAME:
            return t["table_id"]
    return None


def load_existing(tk: str, tid: str) -> dict[str, dict]:
    """{让步申请单号: {record_id, fields}}"""
    H = {"Authorization": f"Bearer {tk}"}
    out: dict[str, dict] = {}
    pt = None
    while True:
        params = {"page_size": 500}
        if pt:
            params["page_token"] = pt
        r = requests.get(f"{OPEN}/bitable/v1/apps/{BITABLE_APP}/tables/{tid}/records",
                         params=params, headers=H, timeout=30, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") != 0:
            raise RuntimeError(f"读记录失败: {d.get('code')} {d.get('msg')}")
        data = d.get("data") or {}
        for rec in data.get("items") or []:
            f = rec.get("fields") or {}
            key = f.get("让步申请单号")
            if isinstance(key, list):   # 文本字段回读可能是 [{"text": ...}]
                key = "".join(x.get("text", "") for x in key if isinstance(x, dict))
            key = str(key or "").strip()
            if key:
                out[key] = {"record_id": rec["record_id"], "fields": f}
        if not data.get("has_more"):
            break
        pt = data.get("page_token")
    return out


def ftext(v) -> str:
    """飞书回读的文本归一化"""
    if v is None:
        return ""
    if isinstance(v, list):
        return "".join(ftext(x) for x in v)
    if isinstance(v, dict):
        return str(v.get("text") or v.get("name") or v.get("value") or "")
    return str(v).strip()


def now_ms() -> int:
    return int(time.time() * 1000)


# ============================================================
# 主流程
# ============================================================
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--from", dest="ymin", default="202001")
    ap.add_argument("--to", dest="ymax", default="203012")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    print("[1/4] 登录 IEC ...", flush=True)
    s = iec_login()
    rows = fetch_concession(s, args.ymin, args.ymax)
    if not rows:
        print("[WARN] IEC 返回 0 条 —— 检查时间窗口 / 组织切换")
        return 1

    # 按单号聚合（一个单号一行；一个单号可挂多个钢厂订单）
    rows = aggregate_by_lld(rows)
    print(f"[IEC] 聚合后 {len(rows)} 张让步申请单")

    print("[2/4] 连接飞书 ...", flush=True)
    tk = feishu_token()
    tid = find_table(tk)
    if not tid:
        if args.dry_run:
            print(f"[WARN] 表「{BITABLE_TABLE_NAME}」不存在（dry-run 继续，按空表算）")
            tid, existing = "", {}
        else:
            print(f"[FAIL] 表「{BITABLE_TABLE_NAME}」不存在，先跑 setup_table.py")
            return 2
    else:
        print(f"[OK] 目标表 {BITABLE_TABLE_NAME} ({tid})")
        existing = load_existing(tk, tid)
    print(f"[飞书] 现有记录 {len(existing)} 条")

    print("[3/4] 计算差异 ...", flush=True)
    to_create, to_update = [], []
    for r0 in rows:
        key = r0["让步申请单号"]
        cur = existing.get(key)
        if cur is None:
            f = {k: r0[k] for k in COLS}
            f["首次抓取时间"] = now_ms()
            f["最后同步时间"] = now_ms()
            to_create.append({"fields": f})
        else:
            changes = {}
            for fld in REPLY_FIELDS:
                old = ftext(cur["fields"].get(fld))
                new = r0[fld]
                if new and old != new:
                    changes[fld] = new
            if changes:
                changes["最后同步时间"] = now_ms()
                to_update.append({"record_id": cur["record_id"], "fields": changes})
                if args.verbose:
                    print(f"   [更新] {key}: {changes}")
    print(f"[差异] 新增 {len(to_create)} 条，更新 {len(to_update)} 条")

    if args.dry_run:
        print("[dry-run] 不写飞书。新增样例:")
        for c in to_create[:3]:
            print("   ", {k: v for k, v in c["fields"].items() if k != "首次抓取时间"})
        for u in to_update[:3]:
            print("   [更新样例]", u)
        return 0

    print("[4/4] 写入飞书 ...", flush=True)
    H = {"Authorization": f"Bearer {tk}", "Content-Type": "application/json; charset=utf-8"}
    ok_c = ok_u = 0
    for i in range(0, len(to_create), 400):
        chunk = to_create[i:i + 400]
        r = requests.post(f"{OPEN}/bitable/v1/apps/{BITABLE_APP}/tables/{tid}/records/batch_create",
                          headers=H, timeout=60, proxies=NO_PROXY, json={"records": chunk})
        d = r.json()
        if d.get("code") != 0:
            print(f"[FAIL] batch_create 失败: {d.get('code')} {d.get('msg')}")
            return 3
        ok_c += len(d["data"]["records"])
        time.sleep(0.3)
    for i in range(0, len(to_update), 400):
        chunk = to_update[i:i + 400]
        r = requests.post(f"{OPEN}/bitable/v1/apps/{BITABLE_APP}/tables/{tid}/records/batch_update",
                          headers=H, timeout=60, proxies=NO_PROXY, json={"records": chunk})
        d = r.json()
        if d.get("code") != 0:
            print(f"[FAIL] batch_update 失败: {d.get('code')} {d.get('msg')}")
            return 4
        ok_u += len(d["data"]["records"])
        time.sleep(0.3)
    print(f"[OK] 新增 {ok_c} 条，更新 {ok_u} 条")
    return 0


if __name__ == "__main__":
    sys.exit(main())
