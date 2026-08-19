"""
验证列表 API /model/admin/xiaoshou/m_kucun/ld_kucun 是否能返回真实数据。
懂火前端用这个接口做分页展示（不是导出 Excel 接口）。
如果 GitHub Actions 上这个接口不返回 0 条，就可以用它替代 /view/admin/excelbiao/kucunld
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

# 预热：访问框架页和子页
s.get(ld.LINDIAO_FRAME_PAGE, timeout=15)
s.get(ld.LINDIAO_PAGE, timeout=15)

# 先看第 1 页
print("=" * 60)
print(f"[1] 调用列表 API: {LIST_API}")
print("    page=1, limit=10 (先看一眼格式)")
r = s.post(LIST_API, data={"page": 1, "limit": 10}, timeout=30)
print(f"    HTTP {r.status_code}, Content-Type={r.headers.get('Content-Type','')}")
print(f"    size={len(r.content)} bytes")
try:
    data = r.json()
    print(f"    JSON keys: {list(data.keys())}")
    print(f"    code={data.get('code')}, msg={data.get('msg')}")
    inner = data.get("data") or {}
    print(f"    data keys: {list(inner.keys()) if isinstance(inner, dict) else type(inner)}")
    if isinstance(inner, dict):
        # 懂火常见格式: data.count + data.list
        count = inner.get("count") or inner.get("total") or inner.get("recordsTotal")
        rows = inner.get("list") or inner.get("rows") or inner.get("data")
        print(f"    count/total={count}")
        print(f"    list/rows type={type(rows).__name__}, len={len(rows) if rows else 0}")
        if rows and isinstance(rows, list) and len(rows) > 0:
            print(f"\n    第 1 条记录字段: {list(rows[0].keys())}")
            print(f"    第 1 条内容（简略）:")
            for k, v in list(rows[0].items())[:15]:
                vs = str(v)[:40]
                print(f"      {k}: {vs!r}")
except Exception as e:
    print(f"    解析失败: {e}")
    print(f"    响应前 500 字符:\n{r.text[:500]}")

# 第 2 步：拉全量（page/limit 翻页），看总数跟 Excel 导出接口（121 条）对不对得上
print("\n" + "=" * 60)
print(f"[2] 翻页拉全量，验证数量")
all_rows = []
page = 1
page_size = 100
while True:
    r = s.post(LIST_API, data={"page": page, "limit": page_size}, timeout=30)
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
        print(f"  page {page} 错误: {e}")
        break

print(f"\n全量总条数: {len(all_rows)}")
if all_rows:
    print(f"第 1 条捆包号: {all_rows[0].get('捆包号') or all_rows[0].get('k_kunbaohao') or '（找不到字段）'}")
    print(f"第 1 条品名:   {all_rows[0].get('品名') or all_rows[0].get('k_pinmin') or '（找不到字段）'}")
    # 打印所有字段名，找和 CSV 23 列对应的中文字段
    print(f"\n所有字段名: {list(all_rows[0].keys())}")
