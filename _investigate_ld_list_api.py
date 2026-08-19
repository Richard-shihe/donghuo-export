"""
侦查：临调库存页面展示的真实数据源。
思路：
  1. 登录懂火
  2. GET 临调库存子页 v_kucun_ld (17KB HTML)
  3. 解析页面里的 JS 代码，找它加载数据用的 API 端点（通常是 /model/.../xjilulist 这类分页接口）
  4. 直接调用这个 JSON API，看能否拿到列表数据（替代 Excel 导出接口）
"""
import os, sys, re, json, requests
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "feishu_uploader"))
import export_lindiao as ld

USER = os.environ["DH_USERNAME"]
PASS = os.environ["DH_PASSWORD"]

s = ld.create_session()
if not ld.login(s, USER, PASS):
    print("登录失败")
    raise SystemExit(1)

# 1. 访问子页
print("=" * 60)
print("[1] GET 临调库存子页 v_kucun_ld")
s.get(ld.LINDIAO_FRAME_PAGE, timeout=15)
r_page = s.get(ld.LINDIAO_PAGE, timeout=15)
html = r_page.text
print(f"  size={len(html)} bytes")

# 2. 找 JS 里的 URL 模式
print("\n" + "=" * 60)
print("[2] 搜索页面中的 API URL 模式 (/model/ 或 xjilulist / list / getlist)")

patterns = [
    r'/model/[^"\'\s]+',            # /model/admin/.../...
    r'url\s*:\s*["\']([^"\']+)["\']',  # url: "xxx"
    r'action\s*=\s*["\']([^"\']+)["\']', # form action
    r'controller/[^"\'\s]+',        # controller/...
]
for pat in patterns:
    matches = re.findall(pat, html, flags=re.I)
    if matches:
        print(f"\n  模式 {pat!r}:")
        uniq = list(dict.fromkeys(matches))[:20]
        for m in uniq:
            print(f"    {m}")

# 3. 找 JS 变量（可能有表名、控制器名）
print("\n" + "=" * 60)
print("[3] 搜索可能的表格/控制器关键字")
keywords = ["kucun", "ld", "v_kucun", "m_kucun", "model_name", "tablename", "table_name", "list_url", "api_url"]
for kw in keywords:
    if kw.lower() in html.lower():
        # 找到前后 60 字符
        idx = html.lower().find(kw.lower())
        ctx = html[max(0,idx-40):idx+len(kw)+40].replace('\n','\\n')
        print(f"  ✅ {kw!r}: ...{ctx}...")

# 4. 找所有 <form> 标签
print("\n" + "=" * 60)
print("[4] 页面中的 <form> 标签")
forms = re.findall(r'<form[^>]*>.*?</form>', html, flags=re.S | re.I)
for i, fm in enumerate(forms):
    action = re.search(r'action=["\']([^"\']+)["\']', fm, flags=re.I)
    method = re.search(r'method=["\']([^"\']+)["\']', fm, flags=re.I)
    inputs = re.findall(r'<input[^>]+>', fm, flags=re.I)
    print(f"  Form {i}: action={action.group(1) if action else '(无)'}, method={method.group(1) if method else 'GET'}, inputs={len(inputs)}")
    for inp in inputs[:10]:
        nm = re.search(r'name=["\']([^"\']+)["\']', inp, flags=re.I)
        vl = re.search(r'value=["\']([^"\']*)["\']', inp, flags=re.I)
        tp = re.search(r'type=["\']([^"\']+)["\']', inp, flags=re.I)
        if nm:
            print(f"    <input> name={nm.group(1)!r}, type={tp.group(1) if tp else 'text'}, value={(vl.group(1) if vl else '')!r}")
