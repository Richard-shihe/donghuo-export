"""再读一次，并打印完整路径和文件内容前后字节"""
import os, hashlib
path = r"c:\Users\H_RQ\Documents\trae_projects\A\.github\workflows\export_lindiao.yml"
print(f"path: {path}")
print(f"exists: {os.path.exists(path)}")
print(f"isfile: {os.path.isfile(path)}")
print(f"size: {os.path.getsize(path)} bytes")

with open(path, "rb") as fp:
    data = fp.read()
print(f"md5: {hashlib.md5(data).hexdigest()}")
print(f"first 50 bytes: {data[:50]!r}")

# 找含 DH1 的行
text = data.decode("utf-8")
hit = False
for i, line in enumerate(text.splitlines(), 1):
    if "DH1" in line or "HONG" in line:
        print(f"  line {i}: {line!r}")
        hit = True
if not hit:
    print("  (no DH1 or HONG found in file)")

# 用 os.stat 查文件最后修改时间
import datetime
mtime = os.path.getmtime(path)
print(f"last modified: {datetime.datetime.fromtimestamp(mtime)}")
