#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一次性诊断：验证懂火 getlist 接口的分页字段语义（只读，不写任何数据）。

为什么需要它：
  backup_all.py 里「服务端每页上限约 300」「超过会丢页」「offset = page * limit」
  这些说法只存在于注释里，没有任何实测依据。要给备份加「完整性总闸」
  （拉到的行数必须 == 接口声明的总数，不等就报错而不是静默上传），
  必须先确认 rtotal / pgtotal 究竟是什么。

验证项：
  1. 每个业务的顶层字段名、rtotal、pgtotal
  2. pgtotal 是否 == ceil(rtotal / limit)，即它是不是「总页数」
  3. limit 取 200 / 300 / 500 时服务端实际返回多少条（验证「上限约 300」）
  4. 挑数据量最小的业务，小分页拉全量，验证「累计条数 == rtotal」
     —— 这一条成立，backup_all.py 才能拿 rtotal 当总闸

用法：python 懂火每周全量备份/_probe_paging.py
跑完可删。
"""

import os
import sys
import json
import math
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "懂火每周全量备份"))

HDR = {"X-Requested-With": "XMLHttpRequest"}
MAX_FULL_PAGES = 30          # 第 4 步最多翻多少页（防失控）
FULL_PAGE_SIZE = 100


def load_dotenv():
    p = os.path.join(ROOT, ".env")
    if not os.path.exists(p):
        print("[WARN] 未找到 .env")
        return
    for line in open(p, encoding="utf-8-sig"):
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


load_dotenv()

from donghuo_login import login_donghuo, BASE_URL          # noqa: E402
import backup_all                                          # noqa: E402


def post(session, path, limit, page):
    return session.post(BASE_URL + path,
                        data={"page": page, "limit": limit},
                        headers=HDR, timeout=60)


def parse(r):
    try:
        d = json.loads(r.text)
        return d if isinstance(d, dict) else None
    except Exception:
        return None


def main():
    user = os.environ.get("DH_USERNAME", "")
    pwd = os.environ.get("DH_PASSWORD", "")
    if not user or not pwd:
        print("[FAIL] .env 里缺 DH_USERNAME / DH_PASSWORD")
        return 2

    print("=" * 72)
    print("懂火 getlist 分页字段语义诊断（只读）")
    print("=" * 72)

    session = login_donghuo(username=user, password=pwd)
    if session is None:
        print("[FAIL] 登录失败")
        return 1
    print("[OK] 登录成功\n")

    tasks = [t for t in backup_all.TASKS if t["api_type"] == "json_paged"]

    # ---------- 步骤 1：每个业务的顶层字段与总数 ----------
    print("-" * 72)
    print("步骤 1：limit=1 探测每个业务的顶层字段 / rtotal / pgtotal")
    print("-" * 72)
    probe_rows = []
    for t in tasks:
        try:
            r = post(session, t["api_path"], 1, 1)
            d = parse(r)
            if d is None:
                print(f"[WARN] {t['biz']}: 非 JSON, HTTP {r.status_code}, {r.text[:120]!r}")
                time.sleep(1)
                continue
            root = d.get(t["records_field"]) or []
            top_keys = [k for k in d.keys() if k != t["records_field"]]
            rtotal = d.get("rtotal")
            pgtotal = d.get(t["total_field"])
            probe_rows.append({
                "biz": t["biz"], "path": t["api_path"],
                "rtotal": rtotal, "pgtotal": pgtotal,
                "top_keys": top_keys, "root_len": len(root),
            })
            print(f"  {t['biz']:<12} rtotal={rtotal!s:<10} pgtotal={pgtotal!s:<8} "
                  f"root={len(root)} 顶层键={top_keys}")
        except Exception as e:
            print(f"[WARN] {t['biz']}: 异常 {e}")
        time.sleep(1)

    if not probe_rows:
        print("[FAIL] 步骤 1 无有效结果，终止")
        return 1

    # ---------- 步骤 2：pgtotal 是否随 limit 变化 ----------
    print("\n" + "-" * 72)
    print("步骤 2：换个 limit 再取 pgtotal，验证「pgtotal == ceil(rtotal/limit)」")
    print("-" * 72)
    t0 = tasks[0]
    for lim in (1, 50, 200):
        try:
            r = post(session, t0["api_path"], lim, 1)
            d = parse(r) or {}
            rt = d.get("rtotal")
            pg = d.get(t0["total_field"])
            expect = None
            try:
                expect = math.ceil(int(rt) / lim)
            except Exception:
                pass
            flag = "一致" if pg == expect else f"不一致(应为 {expect})"
            print(f"  {t0['biz']} limit={lim:<4} rtotal={rt} pgtotal={pg} -> {flag}")
        except Exception as e:
            print(f"  [WARN] limit={lim}: {e}")
        time.sleep(1)

    # ---------- 步骤 3：服务端每页实际上限 ----------
    print("\n" + "-" * 72)
    print("步骤 3：limit=200/300/500/1000 时服务端实际返回多少条")
    print("-" * 72)
    for lim in (200, 300, 500, 1000):
        try:
            r = post(session, t0["api_path"], lim, 1)
            d = parse(r) or {}
            root = d.get(t0["records_field"]) or []
            print(f"  limit={lim:<5} -> 实际返回 {len(root)} 条"
                  f"{'  [截断!]' if len(root) < min(lim, 10**9) and len(root) < lim else ''}")
        except Exception as e:
            print(f"  [WARN] limit={lim}: {e}")
        time.sleep(1)

    # ---------- 步骤 4：最小业务全量对账 ----------
    print("\n" + "-" * 72)
    print("步骤 4：挑数据量最小的业务，小分页翻页，验证「累计条数 == rtotal」")
    print("-" * 72)
    def _int(v):
        try:
            return int(v)
        except Exception:
            return 10 ** 18
    smallest = min(probe_rows, key=lambda x: _int(x["rtotal"]))
    print(f"  选定：{smallest['biz']}  (rtotal={smallest['rtotal']}, pgtotal={smallest['pgtotal']})")

    tsel = next(t for t in tasks if t["biz"] == smallest["biz"])
    seen = set()
    total_rows = 0
    pages_done = 0
    page = 1
    while page <= MAX_FULL_PAGES:
        try:
            r = post(session, tsel["api_path"], FULL_PAGE_SIZE, page)
            d = parse(r)
            if d is None:
                print(f"  [WARN] 第 {page} 页非 JSON，停止")
                break
            root = d.get(tsel["records_field"]) or []
            if not root:
                print(f"  第 {page} 页 root 为空，停止")
                break
            dup = 0
            for rec in root:
                h = hash(json.dumps(rec, sort_keys=True, ensure_ascii=False))
                if h in seen:
                    dup += 1
                seen.add(h)
            total_rows += len(root)
            pages_done += 1
            print(f"  第 {page} 页: +{len(root)} 条, 累计 {total_rows}, 重复 {dup}")
            if page >= _int(d.get(tsel["total_field"])):
                break
            page += 1
            time.sleep(0.8)
        except Exception as e:
            print(f"  [WARN] 第 {page} 页异常: {e}")
            break

    print("\n" + "=" * 72)
    print("对账结论")
    print("=" * 72)
    rt = _int(smallest["rtotal"])
    print(f"  接口声明 rtotal        = {smallest['rtotal']}")
    print(f"  实际拉到（去重前）      = {total_rows}  ({pages_done} 页)")
    print(f"  重复条数              = {total_rows - len(seen)}")
    print(f"  去重后条数            = {len(seen)}")
    if pages_done >= MAX_FULL_PAGES and len(seen) < rt:
        print("  [NOTE] 达到翻页上限被截断，本轮未拉完，无法判定总数是否吻合")
    elif len(seen) == rt:
        print("  [OK] 条数与 rtotal 完全吻合 → rtotal 就是总记录数，可作总闸")
    else:
        print(f"  [WARN] 条数与 rtotal 相差 {rt - len(seen)} → rtotal 语义需要重新判断，"
              f"不能直接当总闸")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())
