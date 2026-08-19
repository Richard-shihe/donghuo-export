"""列出多维表格所有字段，确认'品名'是否还在"""
import os, requests

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

r = requests.get(
    f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/fields",
    headers=headers, timeout=30,
)
data = r.json()
fields = (data.get("data") or {}).get("items") or []
print(f"应用可见字段数: {len(fields)}")
print(f"\n=== 所有字段列表 ===")
for i, f in enumerate(fields):
    name = f.get("field_name") or ""
    name_display = repr(name) if not name else name
    print(f"  [{i:2d}] {name_display:20s}  type={f.get('type')}  ui_type={f.get('ui_type')!r}  field_id={f.get('field_id')}")

# 专门检查品名相关字段
print(f"\n=== 查找品名相关 ===")
found_pinming = False
for f in fields:
    name = f.get("field_name") or ""
    if "品名" in name or "品名" in name or "pinmin" in name.lower():
        print(f"  ✅ 找到: name={name!r}, field_id={f.get('field_id')}, type={f.get('type')}")
        found_pinming = True
if not found_pinming:
    print(f"  ❌ 没有找到包含'品名'的字段")

# 也检查空字符串字段名（字段权限未授权的典型症状）
print(f"\n=== 检查空字段名（可能是字段级权限未授权）===")
empty_name_fields = [f for f in fields if not (f.get("field_name") or "").strip()]
if empty_name_fields:
    print(f"  ⚠️  发现 {len(empty_name_fields)} 个 field_name 为空的字段:")
    for f in empty_name_fields:
        print(f"     field_id={f.get('field_id')}, type={f.get('type')}, ui_type={f.get('ui_type')!r}")
    print(f"     → 这是字段级权限未授权的典型症状（字段存在，但应用没权限看到字段名）")
else:
    print(f"  ✅ 没有空字段名")
