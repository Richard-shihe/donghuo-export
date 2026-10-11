#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一次性诊断：汇总字段（extra_top_keys）的值是否随 limit 变化。

背景：backup_all.py 现在把汇总字段名记在 meta 里，然后**重新 POST 一次第一页
（limit=1）**去取它们的值（第 703-712 行），这次请求没有任何重试保护。
若汇总值与 limit 无关，就可以在拉第一页时顺手取值，省掉这次裸请求。

验证：对每个业务，各取 limit=1 和 limit=50 的第一页，比较顶层汇总字段值是否一致。
只读。跑完可删。
"""

import os
import sys
import json
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "懂火每周全量备份"))

HDR = {"X-Requested-With": "XMLHttpRequest"}


def load_dotenv():
    p = os.path.join(ROOT, ".env")
    if not os.path.exists(p):
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


def fetch(session, path, limit):
    r = session.post(BASE_URL + path, data={"page": 1, "limit": limit},
                     headers=HDR, timeout=60)
    try:
        d = json.loads(r.text)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def main():
    session = login_donghuo(username=os.environ.get("DH_USERNAME", ""),
                             password=os.environ.get("DH_PASSWORD", ""))
    if session is None:
        print("[FAIL] 登录失败")
        return 1
    print("[OK] 登录成功\n")

    tasks = [t for t in backup_all.TASKS if t["api_type"] == "json_paged"]
    all_same = True
    for t in tasks:
        std = {t["records_field"], t["total_field"], "page", "rtotal"}
        try:
            d1 = fetch(session, t["api_path"], 1)
            time.sleep(1)
            d50 = fetch(session, t["api_path"], 50)
            time.sleep(1)
        except Exception as e:
            print(f"[WARN] {t['biz']}: {e}")
            continue
        keys = [k for k in d1.keys() if k not in std]
        if not keys:
            print(f"  {t['biz']:<12} 无汇总字段")
            continue
        diffs = []
        for k in keys:
            v1, v50 = d1.get(k), d50.get(k)
            same = (str(v1) == str(v50))
            if not same:
                diffs.append(f"{k}: limit1={v1!r} limit50={v50!r}")
        if diffs:
            all_same = False
            print(f"  {t['biz']:<12} [DIFF] " + " | ".join(diffs))
        else:
            sample = ", ".join(f"{k}={d1.get(k)!r}" for k in keys[:3])
            print(f"  {t['biz']:<12} [SAME] {len(keys)} 个汇总字段：{sample}")

    print("\n" + "=" * 72)
    if all_same:
        print("[OK] 所有业务的汇总字段值都与 limit 无关")
        print("     → 可以直接在拉第一页时顺手取值，去掉那次无重试的裸请求")
    else:
        print("[WARN] 存在随 limit 变化的汇总字段 → 不能简单改，需保留单独取值逻辑")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())
