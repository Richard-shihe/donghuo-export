import subprocess, re, os
os.chdir(r"c:\Users\H_RQ\Documents\trae_projects\A")

r = subprocess.run(["git","show","HEAD:feishu_uploader/export_lindiao.py"], capture_output=True, text=True)
txt = r.stdout
print("[1] export_lindiao.py (HEAD):")
print("  - export_lindiao_listapi 函数存在:", "def export_lindiao_listapi" in txt)
print("  - MIN_ROWS=20 阈值存在:", "MIN_ROWS = 20" in txt)
print("  - ListAPI失败回退Excel存在:", "回退 Excel 导出接口" in txt)
print("  - [DEBUG] 登录cookie打印存在:", "[DEBUG] 登录响应后 session cookies" in txt)
print("  - [DEBUG] batch_create(保留)存在:", "[DEBUG] batch_create called with" in txt)
print("  - both 交付模式存在:", 'delivery_mode == "both"' in txt)
print("  - DELIVERY_MODE 默认 both 吗？(看 workflow)")

r2 = subprocess.run(["git","show","HEAD:.github/workflows/export_lindiao.yml"], capture_output=True, text=True)
txt2 = r2.stdout
print("\n[2] export_lindiao.yml (HEAD):")
# 找 DELIVERY_MODE 行
for line in txt2.splitlines():
    if "DELIVERY_MODE" in line or "BITABLE_APP_TOKEN" in line or "BITABLE_TABLE_ID" in line or (line.startswith("name:") and "交付" in line):
        print(f"  {line.strip()}")

# 检查 origin/main 和本地 main 的差距
print("\n[3] origin/main vs local main commit")
r3 = subprocess.run(["git", "log", "origin/main..HEAD", "--oneline"], capture_output=True, text=True)
print("  本地领先 origin 的 commits:", r3.stdout.strip() or "(无，本地 == origin)")
r4 = subprocess.run(["git", "log", "HEAD..origin/main", "--oneline"], capture_output=True, text=True)
print("  origin 领先本地的 commits:", r4.stdout.strip() or "(无)")
