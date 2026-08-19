import subprocess, os
os.chdir(r"c:\Users\H_RQ\Documents\trae_projects\A")
# 检查 git 是否认为 yml 文件有改动
r = subprocess.run(["git", "diff", "HEAD", "--", ".github/workflows/export_lindiao.yml"],
                   capture_output=True, text=True)
print("diff stdout len:", len(r.stdout))
print("diff stderr:", r.stderr[:300])
print("--- diff content ---")
print(r.stdout[:2000])

# 检查 git status -s
r2 = subprocess.run(["git", "status", "-s"], capture_output=True, text=True)
print("\nstatus -s:")
print(r2.stdout)

# 检查文件是否被 gitignore 排除
r3 = subprocess.run(["git", "check-ignore", "-v", ".github/workflows/export_lindiao.yml"],
                    capture_output=True, text=True)
print("check-ignore:", r3.stdout or "(not ignored)")
