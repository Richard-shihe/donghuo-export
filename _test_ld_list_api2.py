"""
重试列表 API：带上前端需要的筛选参数（sxzhuantai/huoquan/canku/pinmin）
懂火 PHP 代码直接引用 $sxzhuantai 变量，不传就 Fatal error
"""
import os, sys, json, requests, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "feishu_uploader"))
import export_lindiao as ld

USER = os.environ["DH_USERNAME"]
PASS = os.environ["DH_PASSWORD"]
BASE_URL = ld.BASE_URL
LIST_API = f"{BASE_URL}/model/admin/xiaoshou/m_kucun/ld_kucun"

s = ld.create_session()
if not ld.login(s, USER, PASS):
    print("登录失败")
    raise SystemExit(1)

s.get(ld.LINDIAO_FRAME_PAGE, timeout=15)
s.get(ld.LINDIAO_PAGE, timeout=15)

# 带筛选字段（空值 = 全查）
payload = {
    "page": 1,
    "limit": 10,
    "sxzhuantai": "",   # 状态
    "huoquan": "",      # 货权
    "canku": "",        # 仓库
    "pinmin": "",       # 品名
}
print("=" * 60)
print(f"[1] 调用列表 API，带 4 个筛选空字段")
r = s.post(LIST_API, data=payload, timeout=30)
print(f"    HTTP {r.status_code}, Content-Type={r.headers.get('Content-Type','')}")
print(f"    size={len(r.content)} bytes")
try:
    data = r.json()
    print(f"    code={data.get('code')}, msg={data.get('msg')}")
    inner = data.get("data") or {}
    count = inner.get("count") or inner.get("total") or inner.get("recordsTotal")
    rows = inner.get("list") or inner.get("rows") or inner.get("data")
    print(f"    count/total={count}")
    print(f"    rows len={len(rows) if rows else 0}")
    if rows:
        print(f"\n    第 1 条字段: {list(rows[0].keys())}")
        for k, v in list(rows[0].items())[:20]:
            print(f"      {k}: {str(v)[:50]!r}")
except Exception as e:
    print(f"    解析失败: {e}")
    print(f"    响应前 500 字符:\n{r.text[:500]}")

# 拉全量
print("\n" + "=" * 60)
print(f"[2] 翻页拉全量")
all_rows = []
page = 1
page_size = 200
while True:
    p = dict(payload, page=page, limit=page_size)
    r = s.post(LIST_API, data=p, timeout=30)
    try:
        data = r.json()
        inner = data.get("data") or {}
        rows = inner.get("list") or inner.get("rows") or inner.get("data") or []
        all_rows.extend(rows)
        print(f"  page {page}: {len(rows)} 条，累计 {len(all_rows)} 条")
        if len(rows) < page_size:
            break
        page += 1
        time.sleep(0.2)
    except Exception as e:
        print(f"  错误: {e}, 响应前300: {r.text[:300]}")
        break

print(f"\n全量总数: {len(all_rows)}")

# 核对：拿捆包号和品名字段，找对应列
if all_rows:
    r0 = all_rows[0]
    print("\n所有字段名:")
    for i, (k, v) in enumerate(r0.items()):
        print(f"  [{i:2d}] {k} = {str(v)[:40]!r}")
