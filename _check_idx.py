"""使用 git plumbing 命令查清 git index 对这个文件的状态"""
import subprocess, os
os.chdir(r"c:\Users\H_RQ\Documents\trae_projects\A")

# git ls-files --stage 显示 index 里这个文件的 hash
r = subprocess.run(["git", "ls-files", "-s", ".github/workflows/export_lindiao.yml"],
                   capture_output=True, text=True)
print("index entry:")
print(r.stdout)

# git rev-parse HEAD:file 显示 HEAD 里的 blob hash
r2 = subprocess.run(["git", "rev-parse", "HEAD:.github/workflows/export_lindiao.yml"],
                    capture_output=True, text=True)
print("HEAD blob hash:", r2.stdout.strip())

# git hash-object 当前工作区文件
r3 = subprocess.run(["git", "hash-object", ".github/workflows/export_lindiao.yml"],
                    capture_output=True, text=True)
print("working tree blob hash:", r3.stdout.strip())

# 比较
print("\n如果 index hash == HEAD hash 但工作区 hash 不同 → git 认为 index == HEAD，工作区有改动但 add 没生效")
print("如果 index hash != HEAD hash → index 已 staged 改动，但 git diff --cached 因 line endings 显示空")
