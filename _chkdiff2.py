import subprocess, os, hashlib
os.chdir(r"c:\Users\H_RQ\Documents\trae_projects\A")

# HEAD 版本
r = subprocess.run(["git", "show", "HEAD:feishu_uploader/export_lindiao.py"],
                   capture_output=True)
head_bytes = r.stdout
print("HEAD bytes:", len(head_bytes), "md5:", hashlib.md5(head_bytes).hexdigest())

# 工作区版本
with open("feishu_uploader/export_lindiao.py", "rb") as fp:
    local_bytes = fp.read()
print("LOCAL bytes:", len(local_bytes), "md5:", hashlib.md5(local_bytes).hexdigest())
print("equal:", head_bytes == local_bytes)

# 如果不等，git diff 应该有输出
if head_bytes != local_bytes:
    # 直接写入临时文件比较
    with open("_HEAD_export.py", "wb") as fp: fp.write(head_bytes)
    r2 = subprocess.run(["git", "diff", "--no-index", "_HEAD_export.py", "feishu_uploader/export_lindiao.py"],
                        capture_output=True, text=True)
    print("\n--no-index diff stdout len:", len(r2.stdout))
    print(r2.stdout[:800])
    os.unlink("_HEAD_export.py")
