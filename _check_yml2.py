import subprocess, os, hashlib
os.chdir(r"c:\Users\H_RQ\Documents\trae_projects\A")

# HEAD 版本
r = subprocess.run(["git", "show", "HEAD:.github/workflows/export_lindiao.yml"],
                   capture_output=True)
head_bytes = r.stdout
print("HEAD yml md5:", hashlib.md5(head_bytes).hexdigest(), "size:", len(head_bytes))
# 工作区版本
with open(".github/workflows/export_lindiao.yml", "rb") as fp:
    local_bytes = fp.read()
print("LOCAL yml md5:", hashlib.md5(local_bytes).hexdigest(), "size:", len(local_bytes))
print("equal:", head_bytes == local_bytes)

# 把 HEAD 版本写入临时文件做 --no-index diff
with open("_head_yml.yml", "wb") as fp:
    fp.write(head_bytes)
r2 = subprocess.run(
    ["git", "diff", "--no-index", "--stat", "_head_yml.yml", ".github/workflows/export_lindiao.yml"],
    capture_output=True, text=True
)
print("\n--no-index diff exit code:", r2.returncode)
print("--no-index diff stdout:")
print(r2.stdout[:1500])
print("--no-index diff stderr:")
print(r2.stderr[:300])
os.unlink("_head_yml.yml")
