"""核对 L6ER000579 在 xlsx 中的原始数据"""
import pandas as pd
import glob, os

files = sorted(glob.glob("准发下载_*.xlsx"), key=os.path.getmtime, reverse=True)
if not files:
    print("没找到 xlsx")
    raise SystemExit(1)
f = files[0]
print(f"读取: {f}\n")

df = pd.read_excel(f)
mask = df["钢厂订单号"].astype(str).str.strip() == "L6ER000579"
rows = df[mask]

print(f"=== L6ER000579 共 {len(rows)} 行 ===\n")
if rows.empty:
    print("没找到这个订单号")
    raise SystemExit(0)

for i, (_, r) in enumerate(rows.iterrows(), 1):
    print(f"--- 行 {i} ---")
    for col in df.columns:
        print(f"  {col}: {r[col]}")
    print()

if len(rows) > 1:
    print(f"=== 汇总 ===")
    print(f"  准发量合计: {rows['准发量'].sum()}")
    print(f"  出厂量合计: {rows['出厂量'].sum()}")
