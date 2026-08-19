"""
本地验证：
1. 跑新的 ListAPI 导出路径（export_lindiao_listapi），看条数是否 129 左右
2. 跑老的 Excel 导出路径作为对比
3. 比较两者 CSV 的列头是否一致，捆包号集合是否大致相同
"""
import os, sys, csv, io
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "feishu_uploader"))
import export_lindiao as ld

USER = os.environ["DH_USERNAME"]
PASS = os.environ["DH_PASSWORD"]

s = ld.create_session()
if not ld.login(s, USER, PASS):
    print("登录失败")
    raise SystemExit(1)

filters = ld.build_filters_from_env()
print(filters)

print("=" * 60)
print("[A] 走 ListAPI")
rows_a, csv_a, rc_a, cc_a = ld.export_lindiao_listapi(s, filters=filters)
print(f"A: rows={len(rows_a)}, row_count(含表头)={rc_a}, cols={cc_a}, csv_size={len(csv_a)}")
kb_a = set(r.get("捆包号", "") for r in rows_a if r.get("捆包号"))
print(f"A unique 捆包号: {len(kb_a)}")
if rows_a:
    print(f"A sample: {rows_a[0]}")
# 验证 CSV 列头
lines = csv_a.decode("utf-8-sig").splitlines()
print(f"A CSV header: {lines[0]}")

print("\n" + "=" * 60)
print("[B] 走 Excel 导出（老路径回退）")
html_b, info_b = ld.export_lindiao(s, filters=filters)
csv_b, rc_b, cc_b, rows_b = ld.html_table_to_csv_bytes(html_b)
print(f"B: rows={len(rows_b)}, row_count(含表头)={rc_b}, cols={cc_b}, csv_size={len(csv_b)}")
kb_b = set(r.get("捆包号", "") for r in rows_b if r.get("捆包号"))
print(f"B unique 捆包号: {len(kb_b)}")
if rows_b:
    print(f"B sample: {rows_b[0]}")
lines_b = csv_b.decode("utf-8-sig").splitlines()
print(f"B CSV header: {lines_b[0]}")

print("\n" + "=" * 60)
print("比对:")
print(f"  条数差: A={len(rows_a)} vs B={len(rows_b)} (差 {len(rows_a) - len(rows_b)})")
print(f"  列头一致: {lines[0] == lines_b[0]}")
a_only = sorted(kb_a - kb_b)
b_only = sorted(kb_b - kb_a)
print(f"  A 有 B 无的捆包号: {len(a_only)} 个 {'(正常: list API会多一些)' if a_only else '完全一致'}")
if a_only[:5]:
    print(f"    例: {a_only[:5]}")
print(f"  B 有 A 无的捆包号: {len(b_only)} 个")
if b_only[:5]:
    print(f"    例: {b_only[:5]}")
