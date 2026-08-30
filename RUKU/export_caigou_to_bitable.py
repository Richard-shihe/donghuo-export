#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火采购订单 → 飞书多维表（路线B）
================================

业务流程：
  1. 登录懂火，拉采购订单列表 + 每个订单的明细列表
  2. 每条明细 → 一条飞书记录（明细ID去重，已存在跳过）
  3. 入库专属字段（仓库/库位号/捆包号/件数/重量等）留空，人工补充
  4. 人工补充后把"处理状态"改为"已提交审核"
  5. 跑 import_to_donghuo.py --phase ruku 入库

明细里已包含订单级信息（供应商/合同号/所属公司），无需分别查。

运行（建议从仓库根目录执行）:
  # 预览模式
  python RUKU/export_caigou_to_bitable.py --dry-run

  # 实际执行
  python RUKU/export_caigou_to_bitable.py

  # 只拉最近 N 个订单（调试用）
  python RUKU/export_caigou_to_bitable.py --limit 5

  # 跳过已入库的明细（入库操作=已足量）
  python RUKU/export_caigou_to_bitable.py --skip-ruku-done

环境变量:
  DH_USERNAME / DH_PASSWORD           懂火登录
  FEISHU_APP_ID / FEISHU_APP_SECRET   飞书自建应用
  BITABLE_APP_TOKEN                    多维表 app_token
  BITABLE_TABLE_ID                     多维表 table_id
  FEISHU_WEBHOOK_URL / FEISHU_WEBHOOK_SECRET  机器人通知（可选）
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import hmac
import os
import sys
import time
from collections import defaultdict

import requests

# 脚本位于 RUKU/ 子目录，donghuo_login.py 在仓库根目录
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_REPO_ROOT, os.path.dirname(os.path.abspath(__file__))):
    if _p and _p not in sys.path:
        sys.path.insert(0, _p)
from donghuo_login import login_donghuo


# ============================================================
# 常量
# ============================================================
BASE_URL = "https://erpa.donghuo.vip"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}
SHANGHAI_OFFSET_HOURS = 8

BITABLE_APP_TOKEN_DEFAULT = "OSuobf2ZkaWtUAsXEE9c3aBTnwh"
BITABLE_TABLE_ID_DEFAULT = "tblrhqzHuTsAprU3"

# 懂火接口
API_ORDER_LIST = f"{BASE_URL}/model/admin/caigou/m_dindan/getlist"
API_ITEM_LIST = f"{BASE_URL}/model/admin/caigou/m_dindan/mxlist"
AJAX_HEADERS = {"X-Requested-With": "XMLHttpRequest"}


# ============================================================
# 工具
# ============================================================
def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def now_shanghai_str() -> str:
    utc_now = datetime.datetime.utcnow()
    sh = utc_now + datetime.timedelta(hours=SHANGHAI_OFFSET_HOURS)
    return sh.strftime("%Y-%m-%d %H:%M:%S")


def now_shanghai_ts_ms() -> int:
    utc_now = datetime.datetime.utcnow()
    sh = utc_now + datetime.timedelta(hours=SHANGHAI_OFFSET_HOURS)
    return int(sh.timestamp() * 1000)


def _flatten(v) -> str:
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, (int, float, bool)):
        return str(v).strip()
    if isinstance(v, list):
        parts = []
        for seg in v:
            if isinstance(seg, dict):
                if "text" in seg:
                    parts.append(str(seg["text"]).strip())
                elif "name" in seg:
                    parts.append(str(seg["name"]).strip())
            elif isinstance(seg, str):
                parts.append(seg.strip())
        return "".join(p for p in parts if p).strip()
    return str(v).strip()


def _to_float(v, default=0.0) -> float:
    s = str(v).strip() if v is not None else ""
    if not s:
        return default
    try:
        return float(s)
    except (ValueError, TypeError):
        return default


def _date_to_ts(date_str: str) -> int:
    """日期字符串 '2026-08-19' → 毫秒时间戳"""
    s = str(date_str or "").strip()
    if not s:
        return 0
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return int(dt.timestamp() * 1000)
        except ValueError:
            continue
    return 0


# ============================================================
# 飞书
# ============================================================
def feishu_token() -> str:
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={
        "app_id": env("FEISHU_APP_ID", ""),
        "app_secret": env("FEISHU_APP_SECRET", ""),
    }, timeout=15, proxies=NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换 tenant_access_token 失败: {data}")
    return data["tenant_access_token"]


def bitable_read_all(tok, app_token, table_id) -> list:
    rows = []
    page_token = ""
    empty_streak = 0  # 连续空页计数（防止 has_more=true 但 items=[] 死循环）
    for page_idx in range(200):
        url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(url, params=params,
                         headers={"Authorization": f"Bearer {tok}"},
                         timeout=30, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") != 0:
            print(f"  ⚠️  飞书表查询返回错误 code={d.get('code')} msg={d.get('msg')}")
            break
        items = (d.get("data") or {}).get("items") or []
        if not items:
            empty_streak += 1
            if empty_streak >= 3:
                print(f"  ⚠️  连续 {empty_streak} 页返回空 items，强制中断分页"
                      f"（has_more={((d.get('data') or {}).get('has_more'))}）")
                break
        else:
            empty_streak = 0
        for it in items:
            rec = {"_record_id": it.get("record_id", "")}
            for k, v in (it.get("fields") or {}).items():
                rec[k] = _flatten(v)
            rows.append(rec)
        d_data = d.get("data") or {}
        if page_idx % 5 == 0 or not items:
            print(f"  [分页 {page_idx}] 本页 {len(items)} 条，累计 {len(rows)} 条，"
                  f"has_more={d_data.get('has_more')}")
        if not d_data.get("has_more") or not d_data.get("page_token"):
            break
        page_token = d_data["page_token"]
        time.sleep(0.2)
    return rows


def bitable_batch_create(tok, app_token, table_id, records: list,
                         batch_size: int = 500) -> dict:
    """批量创建记录，返回 {ok, fail, errors}"""
    ok = 0
    fail = 0
    errors = []
    total_batches = (len(records) + batch_size - 1) // batch_size
    for batch_idx, i in enumerate(range(0, len(records), batch_size), 1):
        batch = records[i:i + batch_size]
        url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
               f"/tables/{table_id}/records/batch_create")
        r = requests.post(url, json={"records": batch},
                          headers={"Authorization": f"Bearer {tok}"},
                          timeout=60, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") == 0:
            ok += len(batch)
        else:
            fail += len(batch)
            errors.append(f"batch {i}: code={d.get('code')} msg={d.get('msg')}")
        if batch_idx % 5 == 0 or batch_idx == total_batches or fail:
            print(f"    [写入批次 {batch_idx}/{total_batches}]"
                  f" 成功 {ok} / 失败 {fail}", flush=True)
        time.sleep(0.5)
    return {"ok": ok, "fail": fail, "errors": errors}


def feishu_notify(msg: str):
    webhook_url = env("FEISHU_WEBHOOK_URL")
    if not webhook_url:
        return
    secret = env("FEISHU_WEBHOOK_SECRET")
    body = {"msg_type": "text", "content": {"text": msg}}
    if secret:
        ts = str(int(time.time()))
        body["timestamp"] = ts
        body["sign"] = _feishu_sign(secret, ts)
    try:
        requests.post(webhook_url, json=body, timeout=10, proxies=NO_PROXY)
    except Exception:
        pass


def _feishu_sign(secret, timestamp):
    import base64 as _b64
    string_to_sign = f"{timestamp}\n{secret}"
    h = hmac.new(secret.encode("utf-8"),
                 string_to_sign.encode("utf-8"),
                 hashlib.sha256)
    return _b64.b64encode(h.digest()).decode("utf-8")


# ============================================================
# 懂火
# ============================================================
def donghuo_query_orders(session, limit: int = 0) -> list:
    """拉采购订单列表，返回订单号列表（含空循环保护 + 进度日志）"""
    all_rows = []
    page = 1
    empty_streak = 0
    MAX_PAGES = 500  # 绝对上限，防止死循环
    while page <= MAX_PAGES:
        r = session.post(API_ORDER_LIST,
                         data={"page": page, "limit": 100},
                         headers=AJAX_HEADERS, timeout=15)
        try:
            d = r.json()
        except Exception as exc:
            print(f"  ⚠️  订单列表第 {page} 页解析失败: {exc}，中断", flush=True)
            break
        rows = d.get("root") or []
        if not rows:
            empty_streak += 1
            if empty_streak >= 3:
                print(f"  ⚠️  订单列表连续 {empty_streak} 页为空，中断", flush=True)
                break
        else:
            empty_streak = 0
        all_rows.extend(rows)
        if limit and len(all_rows) >= limit:
            all_rows = all_rows[:limit]
            break
        if page % 10 == 0 or not rows:
            print(f"  [订单列表 第 {page} 页] 本页 {len(rows)} 条，累计 {len(all_rows)} 条",
                  flush=True)
        page += 1
        time.sleep(0.2)
    return all_rows


def donghuo_query_items_all(session, page_size: int = 500,
                            max_pages: int = 50) -> list:
    """全量拉取所有采购明细（不传 sc_danhao，分页拉完整个"订单明细汇总"）。

    相比逐订单查明细，请求次数从 N(订单数) 降到 ceil(总数/page_size)，
    2500+ 明细只需 3~5 页即可拉完。
    """
    all_rows = []
    page = 1
    empty_streak = 0
    while page <= max_pages:
        r = session.post(API_ITEM_LIST,
                         data={"page": page, "limit": page_size},
                         headers=AJAX_HEADERS, timeout=30)
        try:
            d = r.json()
        except Exception as exc:
            print(f"  ⚠️  明细列表第 {page} 页解析失败: {exc}，中断", flush=True)
            break
        rows = d.get("root") or []
        if not rows:
            empty_streak += 1
            if empty_streak >= 2:
                break
        else:
            empty_streak = 0
        all_rows.extend(rows)
        # 首页打印服务端返回的总条数/总页数（如有）
        if page == 1:
            total_count = d.get("rtotal")
            total_pages = d.get("pgtotal")
            print(f"  明细接口返回: rtotal={total_count} pgtotal={total_pages}"
                  f" (本页 {len(rows)} 条)", flush=True)
        print(f"  [明细 第 {page} 页] 本页 {len(rows)} 条，累计 {len(all_rows)} 条",
              flush=True)
        if not rows:
            break
        page += 1
        time.sleep(0.3)
    return all_rows


# ============================================================
# 明细 → 飞书记录
# ============================================================
def item_to_feishu_record(item: dict) -> dict:
    """懂火明细 → 飞书多维表 fields dict

    明细列表已包含订单级信息（供应商/合同号/所属公司）。
    入库专属字段（仓库/库位号/捆包号/件数等）留空，人工补充。
    """
    item_id = str(item.get("id") or "").strip()
    c_danhao = str(item.get("订单号") or "").strip()
    ruku_status = str(item.get("入库操作") or "").strip()

    fields = {
        # 订单级（从明细直接取）
        "采购订单号": c_danhao,
        "采购明细ID": item_id,
        "所属公司": str(item.get("所属公司") or "").strip(),
        "供应商": str(item.get("供应商") or "").strip(),
        "合同号": str(item.get("采购合同号") or "").strip(),
        # 明细级
        "品名": str(item.get("品名") or "").strip(),
        "规格": str(item.get("规格") or "").strip(),
        "材质": str(item.get("材质") or "").strip(),
        "产地": str(item.get("产地") or "").strip(),
        "等级": str(item.get("等级") or "").strip(),
        "锌层": str(item.get("锌层") or "").strip(),
        "涂料": str(item.get("涂料") or "").strip(),
        "结构": str(item.get("结构") or "").strip(),
        "颜色": str(item.get("颜色") or "").strip(),
        "提单号": str(item.get("提单号") or "").strip(),
        "备注": str(item.get("备注") or "").strip(),
        # 处理状态：已入库→跳过，否则→待入库
        "处理状态": "已入库" if "已足量" in ruku_status else "待入库",
        "写入结果": f"从懂火导入 {now_shanghai_str()}",
        "写入时间": now_shanghai_ts_ms(),
        "批次时间": now_shanghai_ts_ms(),
    }

    # 数字字段（空值不传，飞书表留空）
    weight = _to_float(item.get("订单重量"), 0)
    if weight:
        fields["重量"] = weight

    price = _to_float(item.get("采购单价"), 0)
    if price:
        fields["采购单价"] = price

    # 入库日期 → 毫秒时间戳
    date_ts = _date_to_ts(str(item.get("日期") or "").strip())
    if date_ts:
        fields["入库日期"] = date_ts

    return fields


# ============================================================
# 主流程
# ============================================================
def parse_args():
    p = argparse.ArgumentParser(
        description="懂火采购订单 → 飞书多维表（路线B）",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--dry-run", action="store_true", help="预览模式")
    p.add_argument("--limit", type=int, default=0,
                   help="只拉前 N 个订单（调试用）")
    p.add_argument("--skip-ruku-done", action="store_true",
                   help="跳过已入库的明细（入库操作=已足量）")
    return p.parse_args()


def main():
    args = parse_args()
    dry_run = args.dry_run or env("DRY_RUN") == "1"

    app_token = env("BITABLE_APP_TOKEN", BITABLE_APP_TOKEN_DEFAULT)
    table_id = env("BITABLE_TABLE_ID", BITABLE_TABLE_ID_DEFAULT)

    print(f"====== 懂火采购订单 → 飞书多维表 ======", flush=True)
    print(f"  模式: {'DRY-RUN' if dry_run else '实际执行'}", flush=True)
    print(f"  limit: {args.limit if args.limit else '全部'}", flush=True)
    print(f"  跳过已入库: {args.skip_ruku_done}", flush=True)
    print(f"  飞书表: {app_token} / {table_id}", flush=True)

    t0 = time.time()

    # 1. 飞书 token
    print("\n[步骤 1/4] 获取飞书 token...", flush=True)
    tok = feishu_token()
    print("  ✅ 飞书 token OK", flush=True)

    # 2. 读飞书表已有的明细ID（去重用）
    print("\n[步骤 2/4] 读飞书表已有记录（去重用）...", flush=True)
    existing_rows = bitable_read_all(tok, app_token, table_id)
    existing_item_ids = set()
    for r in existing_rows:
        iid = str(r.get("采购明细ID") or "").strip()
        if iid:
            existing_item_ids.add(iid)
    print(f"  飞书表已有 {len(existing_rows)} 条，其中 {len(existing_item_ids)} 条有明细ID",
          flush=True)

    # 3. 登录懂火
    print("\n[步骤 3/4] 登录懂火...", flush=True)
    session = login_donghuo()
    if session is None:
        print("登录失败", flush=True)
        return 1

    # 4. 全量拉采购明细 + 过滤 C2026- 开头
    print("\n[步骤 4/4] 全量拉采购明细（mxlist 不传 sc_danhao）...", flush=True)
    all_items_raw = donghuo_query_items_all(session)
    print(f"  明细接口共返回 {len(all_items_raw)} 条", flush=True)

    ORDER_PREFIX = "C2026-"
    total_before = len(all_items_raw)
    all_items = [it for it in all_items_raw
                 if str(it.get("订单号") or "").strip().startswith(ORDER_PREFIX)]
    print(f"  过滤 '{ORDER_PREFIX}' 开头: {total_before} → {len(all_items)} 条明细",
          flush=True)

    session.close()

    print(f"\n  待处理明细: {len(all_items)} 条", flush=True)

    # 6. 过滤 + 去重
    to_write = []
    skipped_dedup = 0
    skipped_ruku = 0
    for item in all_items:
        item_id = str(item.get("id") or "").strip()
        if not item_id:
            continue

        # 去重：飞书表已有
        if item_id in existing_item_ids:
            skipped_dedup += 1
            continue

        # 跳过已入库
        ruku_status = str(item.get("入库操作") or "").strip()
        if args.skip_ruku_done and "已足量" in ruku_status:
            skipped_ruku += 1
            continue

        to_write.append(item)

    print(f"\n  去重跳过: {skipped_dedup}", flush=True)
    print(f"  已入库跳过: {skipped_ruku}", flush=True)
    print(f"  待写入: {len(to_write)} 条", flush=True)

    session.close()

    # 7. 构造飞书记录
    records = []
    for item in to_write:
        fields = item_to_feishu_record(item)
        records.append({"fields": fields})

    # 8. 写入飞书表
    if dry_run:
        print(f"\n[DRY-RUN] 跳过写入，预览前 3 条:", flush=True)
        for i, rec in enumerate(records[:3], 1):
            print(f"\n  --- 第 {i} 条 ---", flush=True)
            for k, v in rec["fields"].items():
                print(f"    {k}: {v}", flush=True)
    else:
        if records:
            print(f"\n[步骤] 写入飞书表 ({len(records)} 条)...", flush=True)
            result = bitable_batch_create(tok, app_token, table_id, records)
            print(f"  ✅ 成功 {result['ok']} 条", flush=True)
            if result["fail"]:
                print(f"  ❌ 失败 {result['fail']} 条", flush=True)
                for e in result["errors"][:3]:
                    print(f"    {e}", flush=True)
        else:
            print("\n  无新记录需要写入", flush=True)

    # 9. 汇总 + 通知
    elapsed = time.time() - t0
    new_count = len(records) if not dry_run else 0
    msg = (
        f"【懂火→飞书 采购订单导入】{'DRY-RUN ' if dry_run else ''}{now_shanghai_str()}\n"
        f"明细总数: {len(all_items_raw)}  C2026-过滤后: {len(all_items)}\n"
        f"去重跳过: {skipped_dedup}  已入库跳过: {skipped_ruku}\n"
        f"新写入: {new_count if not dry_run else len(records)}(预览)\n"
        f"耗时: {elapsed:.1f}s"
    )
    print(f"\n{'='*60}", flush=True)
    print(msg, flush=True)
    print(f"{'='*60}", flush=True)
    feishu_notify(msg)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
