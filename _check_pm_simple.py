"""快速查最新 3 条记录的品名字段"""
import os, requests

APP_TOKEN = "IRCzb5XVganwbUsz9NKcaC0SnRA"
TABLE_ID = "tblaHMNprLueWYDP"

r = requests.post(
    "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
    json={"app_id": os.environ["FEISHU_APP_ID"], "app_secret": os.environ["FEISHU_APP_SECRET"]},
    timeout=15,
)
tok = r.json()["tenant_access_token"]
r = requests.get(
    f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records",
    headers={"Authorization": f"Bearer {tok}"},
    params={"page_size": 3},
    timeout=30,
)
data = r.json().get("data") or {}
items = data.get("items") or []
print(f"total={data.get('total')}, 返回 {len(items)} 条")
for i, it in enumerate(items):
    f = it.get("fields") or {}
    pm = f.get("品名")
    if isinstance(pm, list) and pm and isinstance(pm[0], dict):
        pm = "".join(seg.get("text", "") for seg in pm)
    kbh = f.get("捆包号")
    if isinstance(kbh, list) and kbh and isinstance(kbh[0], dict):
        kbh = "".join(seg.get("text", "") for seg in kbh)
    print(f"[{i}] record_id={it.get('record_id', '')[:20]} 捆包号={kbh!r} 品名={pm!r}")
