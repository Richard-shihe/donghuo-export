"""检查今天写入的 115 条记录里品名字段是否有值"""
import os, requests, datetime as dt

APP_TOKEN = "IRCzb5XVganwbUsz9NKcaC0SnRA"
TABLE_ID = "tblaHMNprLueWYDP"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"

r = requests.post(
    f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
    json={"app_id": os.environ["FEISHU_APP_ID"], "app_secret": os.environ["FEISHU_APP_SECRET"]},
    timeout=15,
)
tok = r.json()["tenant_access_token"]
headers = {"Authorization": f"Bearer {tok}"}

# 取最新 5 条记录（按创建时间倒序，看最新写入的）
url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records"
r = requests.get(url, headers=headers, params={"page_size": 5, "sort": ['{"field_id":"时间","desc":true}']}, timeout=30)
data = r.json().get("data") or {}
items = data.get("items") or []
print(f"最新 {len(items)} 条记录（按时间倒序）:")
print()
for i, it in enumerate(items):
    f = it.get("fields") or {}
    print(f"--- 记录 {i+1} (record_id={it.get('record_id','')[:20]}...) ---")
    # 打印所有字段，特别关注品名
    for k, v in f.items():
        # 把列表型 text 字段展平
        if isinstance(v, list) and v and isinstance(v[0], dict) and "text" in v[0]:
            v_display = "".join(seg.get("text","") for seg in v)
        elif isinstance(v, dict):
            v_display = str(v)[:60]
        else:
            v_display = str(v)[:60]
        marker = " ⚠️ 空" if v_display == "" or v_display == "[]" else ""
        if k == "品名":
            print(f"  >>> 品名 = {v_display!r}{marker}  <<<")
        elif k in ("捆包号", "时间", "所属公司", "规格", "材质", "产地", "仓库"):
            print(f"  {k} = {v_display!r}{marker}")
    print()
