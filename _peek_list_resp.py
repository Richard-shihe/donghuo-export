"""看列表 API 返回的到底是什么 HTML/JSON"""
import os, sys, re, requests
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "feishu_uploader"))
import export_lindiao as ld

USER = os.environ["DH_USERNAME"]
PASS = os.environ["DH_PASSWORD"]
LIST_API = f"{ld.BASE_URL}/model/admin/xiaoshou/m_kucun/ld_kucun"

s = ld.create_session()
if not ld.login(s, USER, PASS): raise SystemExit(1)
s.get(ld.LINDIAO_FRAME_PAGE, timeout=15)
s.get(ld.LINDIAO_PAGE, timeout=15)

payload = {"page": 1, "limit": 10, "sxzhuantai": "", "huoquan": "", "canku": "", "pinmin": ""}
r = s.post(LIST_API, data=payload, timeout=30)

print(f"HTTP {r.status_code}")
print(f"Content-Type: {r.headers.get('Content-Type','')}")
print(f"size: {len(r.content)} bytes")
print("\n=== 响应全文 ===")
print(r.text)
print("\n=== 在内容里搜 JSON 结构（找 { 开头 [ 开头）===")
for i, line in enumerate(r.text.splitlines()):
    line = line.strip()
    if not line: continue
    if line.startswith("{") or line.startswith("[") or "\"count\"" in line or "\"list\"" in line:
        print(f"  line {i}: {line[:200]}")
