"""快速统计今天 8/17 的记录数和总数"""
import os, time, requests, datetime as dt
from collections import Counter, defaultdict

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

today_start_ms = int(dt.datetime(2026, 8, 17).timestamp() * 1000)
today_end_ms = int(dt.datetime(2026, 8, 18).timestamp() * 1000)

url_l = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records"
r = requests.get(url_l, headers=headers, params={"page_size": 1}, timeout=30)
total = (r.json().get("data") or {}).get("total")
print(f"全表总数: {total}")

today_records = []
page_token = ""
total_scanned = 0
while True:
    params = {"page_size": 500}
    if page_token: params["page_token"] = page_token
    r = requests.get(url_l, headers=headers, params=params, timeout=30)
    d = r.json()
    if d.get("code") != 0:
        print(f"[err] {d}"); break
    data = d.get("data") or {}
    items = data.get("items") or []
    total_scanned += len(items)
    for it in items:
        f = it.get("fields") or {}
        t = f.get("时间")
        t_ms = None
        if isinstance(t, (int, float)): t_ms = int(t)
        elif isinstance(t, dict):
            t_ms = t.get("timestamp") or t.get("value")
            if t_ms: t_ms = int(t_ms)
        if not t_ms or not (today_start_ms <= t_ms < today_end_ms):
            continue
        b = f.get("捆包号")
        if isinstance(b, list):
            b = "".join(seg.get("text","") for seg in b if isinstance(seg, dict))
        elif b is None: b = ""
        today_records.append((it.get("record_id",""), t_ms, str(b).strip()))
    if not data.get("has_more"): break
    np = data.get("page_token") or ""
    if not np or np == page_token: break
    page_token = np
    time.sleep(0.15)

print(f"\n今天(8/17)记录数: {len(today_records)}")

if today_records:
    time_counter = Counter(t_ms for _, t_ms, _ in today_records)
    print(f"\n=== 时间戳分布 ===")
    for t_ms, c in time_counter.most_common():
        print(f"  {dt.datetime.fromtimestamp(t_ms/1000).strftime('%Y-%m-%d %H:%M:%S')}: {c} 条")

    bundle_map = defaultdict(list)
    for rid, t_ms, b in today_records:
        bundle_map[b].append((rid, t_ms))
    unique_bundles = len(bundle_map)
    dup_bundles = {b: lst for b, lst in bundle_map.items() if len(lst) > 1}
    print(f"\n=== 捆包号分布 ===")
    print(f"  unique 捆包号: {unique_bundles}")
    print(f"  重复捆包号数: {len(dup_bundles)}")
    if dup_bundles:
        cnt_dist = Counter(len(v) for v in dup_bundles.values())
        print(f"  重复度分布: {dict(cnt_dist)}")
    else:
        print(f"  ✅ 无重复")
