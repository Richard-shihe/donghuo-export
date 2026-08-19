"""本地测试：不同筛选组合对导出结果的影响"""
import os, sys, csv, io
# 把 feishu_uploader 目录加入 path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "feishu_uploader"))
import export_lindiao as ld

USER = os.environ["DH_USERNAME"]
PASS = os.environ["DH_PASSWORD"]

# Case 1: 无任何环境变量（全默认空）
print("=" * 50)
print("Case 1: 清空 FILTER_* 环境变量，模拟无筛选")
for ev in ["FILTER_SXZHUANTAI", "FILTER_HUOQUAN", "FILTER_CANKU", "FILTER_PINMIN"]:
    if ev in os.environ: del os.environ[ev]

f = ld.build_filters_from_env()
print(f"  build_filters_from_env() => {f!r}")
s = ld.create_session()
if not ld.login(s, USER, PASS):
    print("  登录失败，跳过")
else:
    html, info = ld.export_lindiao(s, filters=f)
    csv_bytes, nrows, ncols, rows = ld.html_table_to_csv_bytes(html)
    print(f"  结果: size={info['size_kb']} KB, rows(含表头)={nrows}, cols={ncols}")
    print(f"  实际数据条数: {len(rows)}")

# Case 2: 环境变量都设成空字符串（模拟 GitHub vars 未设置 fallback ''）
print("\n" + "=" * 50)
print("Case 2: 4 个 FILTER_* 都设为 '' (空字符串)")
os.environ["FILTER_SXZHUANTAI"] = ""
os.environ["FILTER_HUOQUAN"] = ""
os.environ["FILTER_CANKU"] = ""
os.environ["FILTER_PINMIN"] = ""
f = ld.build_filters_from_env()
print(f"  build_filters_from_env() => {f!r}")
s2 = ld.create_session()
if not ld.login(s2, USER, PASS):
    print("  登录失败，跳过")
else:
    html, info = ld.export_lindiao(s2, filters=f)
    csv_bytes, nrows, ncols, rows = ld.html_table_to_csv_bytes(html)
    print(f"  结果: size={info['size_kb']} KB, rows(含表头)={nrows}, cols={ncols}")
    print(f"  实际数据条数: {len(rows)}")

# Case 3: 故意用一个不存在的仓库值（比如 FILTER_SXZHUANTAI='已锁'）
print("\n" + "=" * 50)
print("Case 3: FILTER_SXZHUANTAI='已锁' (故意加上状态筛选，看会不会变 0)")
os.environ["FILTER_SXZHUANTAI"] = "已锁"
f = ld.build_filters_from_env()
print(f"  build_filters_from_env() => {f!r}")
s3 = ld.create_session()
if not ld.login(s3, USER, PASS):
    print("  登录失败，跳过")
else:
    html, info = ld.export_lindiao(s3, filters=f)
    csv_bytes, nrows, ncols, rows = ld.html_table_to_csv_bytes(html)
    print(f"  结果: size={info['size_kb']} KB, rows(含表头)={nrows}, cols={ncols}")
    print(f"  实际数据条数: {len(rows)}")
