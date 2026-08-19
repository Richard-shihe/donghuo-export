"""
诊断：检查飞书应用是否能正常访问目标多维表格。
- 获取 tenant_access_token
- 列表记录（GET /records）— 验证读权限
- 列字段（GET /fields）— 验证字段可见性
- 尝试创建 1 条测试记录 — 验证写权限（成功后立即删除）
"""
import os, requests, datetime as dt

APP_TOKEN = "IRCzb5XVganwbUsz9NKcaC0SnRA"
TABLE_ID = "tblaHMNprLueWYDP"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"

# 1. tenant_access_token
r = requests.post(
    f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
    json={"app_id": os.environ["FEISHU_APP_ID"],
          "app_secret": os.environ["FEISHU_APP_SECRET"]},
    timeout=15,
)
print(f"[1] auth: HTTP {r.status_code}, code={r.json().get('code')}")
if r.json().get("code") != 0:
    print(f"    ERROR: {r.json()}")
    raise SystemExit(1)
tok = r.json()["tenant_access_token"]
print(f"    token len={len(tok)}")
headers = {"Authorization": f"Bearer {tok}"}

# 2. 列表记录（读权限）
print(f"\n[2] list records (前 1 条)")
r = requests.get(
    f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records",
    headers=headers, params={"page_size": 1}, timeout=30,
)
print(f"    HTTP {r.status_code}, code={r.json().get('code')}, msg={r.json().get('msg')}")
data = r.json().get("data") or {}
print(f"    total={data.get('total')}, has_more={data.get('has_more')}")

# 3. 列字段
print(f"\n[3] list fields")
r = requests.get(
    f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/fields",
    headers=headers, timeout=30,
)
print(f"    HTTP {r.status_code}, code={r.json().get('code')}, msg={r.json().get('msg')}")
fields = (r.json().get("data") or {}).get("items") or []
print(f"    字段数: {len(fields)}")
# 检查关键字段
key_fields = ["捆包号", "时间"]
for kf in key_fields:
    found = [f for f in fields if f.get("field_name") == kf]
    if found:
        f = found[0]
        print(f"    [{kf}] field_id={f.get('field_id')}, type={f.get('type')}, ui_type={f.get('ui_type')}")
    else:
        print(f"    [{kf}] ❌ 字段不存在或无权访问")

# 4. 尝试创建 1 条测试记录
print(f"\n[4] 尝试创建 1 条测试记录（验证写权限）")
test_record = {
    "fields": {
        "捆包号": "TEST_PERM_CHECK_DELETE_ME",
        "时间": int(dt.datetime.now().timestamp() * 1000),
    }
}
r = requests.post(
    f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records",
    headers=headers, json=test_record, timeout=30,
)
print(f"    HTTP {r.status_code}, code={r.json().get('code')}, msg={r.json().get('msg')}")
resp = r.json()
if resp.get("code") == 0:
    rid = (resp.get("data") or {}).get("record", {}).get("record_id")
    print(f"    ✅ 写权限正常, record_id={rid}")
    # 立即删除
    r = requests.delete(
        f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records/{rid}",
        headers=headers, timeout=30,
    )
    print(f"    [清理] 删除测试记录 HTTP {r.status_code}, code={r.json().get('code')}")
else:
    print(f"    ❌ 写入失败！raw={resp}")
    print(f"\n    可能原因：")
    print(f"    - 飞书应用未被加为多维表格的'可编辑'协作者")
    print(f"    - 字段级权限未授权（'捆包号'/'时间'字段）")
    print(f"    解决：到 https://s2v31ke6sl.feishu.cn/base/{APP_TOKEN} 把应用 cli_aaf0ce1e9ef89d27 加为可编辑协作者")
