"""验证 tblaHMNprLueWYDP 的 record 元数据 created_time 是否可获取"""
import os, requests, json

APP_ID = os.environ.get("FEISHU_APP_ID") or "cli_aaf0ce1e9ef89d27"
APP_SECRET = os.environ.get("FEISHU_APP_SECRET") or ""  # 真实 Secret 从环境变量注入，不进仓库
APP_TOKEN = os.environ.get("BITABLE_APP_TOKEN") or "IRCzb5XVganwbUsz9NKcaC0SnRA"
TABLE_ID = os.environ.get("BITABLE_TABLE_ID") or "tblaHMNprLueWYDP"

r = requests.post('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
                  json={'app_id': APP_ID, 'app_secret': APP_SECRET}, timeout=15)
token = r.json()['tenant_access_token']
headers = {'Authorization': f'Bearer {token}'}

# 拿 1 条 record，看完整结构
r = requests.get(
    f'https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records?page_size=1',
    headers=headers, timeout=15)
data = r.json()
print(f"code={data.get('code')}, msg={data.get('msg')}")
if data.get('code') == 0:
    items = data['data'].get('items', [])
    if items:
        item = items[0]
        print(f"\nrecord 完整 key: {list(item.keys())}")
        print(f"  record_id: {item.get('record_id')}")
        print(f"  created_time: {item.get('created_time')}")
        print(f"  last_modified_time: {item.get('last_modified_time')}")
        print(f"  created_by: {item.get('created_by')}")
        print(f"  last_modified_by: {item.get('last_modified_by')}")
        # 看 fields 里的"时间"字段值
        fields = item.get('fields', {})
        print(f"\n  fields 里的'时间'字段: {fields.get('时间')}")
        print(f"  fields 里的'修改时间'字段: {fields.get('修改时间')}")

# 试一下加参数 include_total 看能不能拿到 created_time
print("\n--- 尝试加 query param ---")
for param in ['true', 'false']:
    r = requests.get(
        f'https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records?page_size=1&include_total={param}',
        headers=headers, timeout=15)
    data = r.json()
    if data.get('code') == 0:
        items = data['data'].get('items', [])
        if items:
            print(f"  include_total={param}: created_time={items[0].get('created_time')}")

# 用 search API 试试
print("\n--- 尝试 search API ---")
r = requests.post(
    f'https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records/search?page_size=1',
    headers=headers, json={}, timeout=15)
data = r.json()
print(f"  code={data.get('code')}, msg={data.get('msg')}")
if data.get('code') == 0:
    items = data['data'].get('items', [])
    if items:
        print(f"  search API record keys: {list(items[0].keys())}")
        print(f"  created_time: {items[0].get('created_time')}")
