#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试：
1) 库存 kucungl 接口：POST vs GET，以及是否需要先访问库存管理页面
2) 销售订单 getlist：sleep 几秒再打会不会就不 SQL 断开了
"""
import sys, os, time, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from donghuo_login import login_donghuo, BASE_URL

def dump(label, text, limit=500):
    print(f"\n===== {label} =====")
    print(f"len={len(text)}")
    print(text[:limit])
    if len(text) > limit:
        print(f"... (truncated, total {len(text)} chars)")

def main():
    session = login_donghuo(os.environ["DH_USERNAME"], os.environ["DH_PASSWORD"])
    if session is None:
        print("登录失败"); return

    # ===== 1) 先 GET 一下业务页面，建立服务端上下文 =====
    pages_to_visit = [
        # 库存管理菜单里的页面（猜测）——和 kucunld 对应 v_kucun_ld 类似，kucungl 应该对应某个 iframe 或主页面
        "/view/admin/xiaoshou/v_ifram_kc",    # 库存查询总框架（含公司库存等）
        "/view/admin/xiaoshou/v_kucungl",     # 可能的库存管理页面
    ]
    for p in pages_to_visit:
        try:
            r = session.get(BASE_URL + p, timeout=30)
            print(f"  访问 {p} → HTTP {r.status_code}, len={len(r.text)}")
        except Exception as e:
            print(f"  访问 {p} 失败: {e}")

    # ===== 2) GET 方式的 kucungl（原失败方式） =====
    print("\n--- 2) GET /view/admin/excelbiao/kucungl ---")
    try:
        r = session.get(BASE_URL + "/view/admin/excelbiao/kucungl", timeout=60, allow_redirects=True)
        dump("GET kucungl", r.text, limit=1500)
    except Exception as e:
        print(f"  失败: {e}")

    # ===== 3) POST 方式的 kucungl（类比 kucunld） =====
    print("\n--- 3) POST /view/admin/excelbiao/kucungl (空 data) ---")
    try:
        r = session.post(BASE_URL + "/view/admin/excelbiao/kucungl", data={}, timeout=60, allow_redirects=True)
        dump("POST kucungl", r.text, limit=1500)
    except Exception as e:
        print(f"  失败: {e}")

    # ===== 4) 验证销售订单接口：sleep 6s 后再打 =====
    print("\n--- 4) 销售订单 getlist (先 sleep 6s) ---")
    time.sleep(6)
    try:
        r = session.post(
            BASE_URL + "/model/admin/xiaoshou/m_dindan/getlist",
            data={"page": 1, "limit": 5},
            headers={"X-Requested-With": "XMLHttpRequest"},
            timeout=60,
        )
        dump("xsdd getlist page1 limit5", r.text, limit=800)
    except Exception as e:
        print(f"  失败: {e}")

    # ===== 5) 验证采购订单 =====
    print("\n--- 5) 采购订单 getlist (再 sleep 4s) ---")
    time.sleep(4)
    try:
        r = session.post(
            BASE_URL + "/model/admin/caigou/m_dindan/getlist",
            data={"page": 1, "limit": 5},
            headers={"X-Requested-With": "XMLHttpRequest"},
            timeout=60,
        )
        dump("cgdd getlist page1 limit5", r.text, limit=800)
    except Exception as e:
        print(f"  失败: {e}")

if __name__ == "__main__":
    main()
