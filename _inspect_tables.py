"""查 tblUzkPskttsBa0W 和 tblaHMNprLueWYDP 的字段结构 + 前几条记录"""
import os, requests, json

APP_ID = os.environ.get("FEISHU_APP_ID") or "cli_aaf0ce1e9ef89d27"
APP_SECRET = os.environ.get("FEISHU_APP_SECRET") or ""  # 真实 Secret 从环境变量注入，不进仓库
APP_TOKEN = os.environ.get("BITABLE_APP_TOKEN") or "IRCzb5XVganwbUsz9NKcaC0SnRA"

r = requests.post('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
                  json={'app_id': APP_ID, 'app_secret': APP_SECRET}, timeout=15)
token = r.json()['tenant_access_token']
print(f'[token] len={len(token)}')
headers = {'Authorization': f'Bearer {token}'}

for table_id, table_name in [
    ('tblUzkPskttsBa0W', 'AI 反馈表'),
    ('tblaHMNprLueWYDP', '临调库存主表'),
]:
    print(f'\n{"="*60}\n=== {table_name} ({table_id}) 字段 ===')
    r = requests.get(
        f'https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{table_id}/fields',
        headers=headers, timeout=15)
    data = r.json()
    if data.get('code') != 0:
        print(f'ERROR: {data}')
        continue
    for f in data['data']['items']:
        print(f"  {f['field_id']}  name={f['field_name']!r}  type={f['type']}  ui_type={f.get('ui_type','')!r}")

    print(f'\n=== {table_name} 前 3 条记录 ===')
    r = requests.get(
        f'https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{table_id}/records?page_size=3',
        headers=headers, timeout=15)
    data = r.json()
    if data.get('code') != 0:
        print(f'ERROR: {data}')
        continue
    items = data['data'].get('items', [])
    total = data['data'].get('total', 0)
    print(f'共 {total} 条')
    for i, item in enumerate(items):
        print(f'\n  [{i}] record_id={item["record_id"]}  created_time={item.get("created_time")}')
        for k, v in item['fields'].items():
            vs = json.dumps(v, ensure_ascii=False)[:120]
            print(f'      {k}: {vs}')
