"""查 fldr8lk2Lf (时间) 字段的 property 配置"""
import os, requests, json

APP_ID = os.environ.get("FEISHU_APP_ID") or "cli_aaf0ce1e9ef89d27"
APP_SECRET = os.environ.get("FEISHU_APP_SECRET") or ""  # 真实 Secret 请从环境变量注入，不进仓库
APP_TOKEN = os.environ.get("BITABLE_APP_TOKEN") or "IRCzb5XVganwbUsz9NKcaC0SnRA"
TABLE_ID = os.environ.get("BITABLE_TABLE_ID") or "tblaHMNprLueWYDP"
FIELD_ID = os.environ.get("FIELD_ID") or "fldr8lk2Lf"

r = requests.post('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
                  json={'app_id': APP_ID, 'app_secret': APP_SECRET}, timeout=15)
token = r.json()['tenant_access_token']
headers = {'Authorization': f'Bearer {token}'}

r = requests.get(
    f'https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/fields/{FIELD_ID}',
    headers=headers, timeout=15)
data = r.json()
print(json.dumps(data, ensure_ascii=False, indent=2))

# 看最近写入的记录"时间"字段有没有值（按"时间"降序取前 3 条）
print("\n--- 按'时间'字段降序取前 3 条（看最近导入的记录）---")
r = requests.get(
    f'https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records?page_size=3&sort=%5B%7B%22field_name%22%3A%22%E6%97%B6%E9%97%B4%22%2C%22desc%22%3Atrue%7D%5D',
    headers=headers, timeout=15)
data = r.json()
if data.get('code') == 0:
    items = data['data'].get('items', [])
    for i, item in enumerate(items):
        fields = item.get('fields', {})
        print(f"  [{i}] record_id={item['record_id']}  时间={fields.get('时间')}  捆包号={fields.get('捆包号')}")
else:
    print(f"  ERROR: {data}")
