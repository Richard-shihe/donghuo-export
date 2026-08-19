"""直接读工作区 yml 文件，看里面到底是 DH 还是 DH1"""
import os, hashlib
path = r"c:\Users\H_RQ\Documents\trae_projects\A\.github\workflows\export_lindiao.yml"
with open(path, "rb") as fp:
    data = fp.read()
print(f"file size: {len(data)} bytes")
print(f"md5: {hashlib.md5(data).hexdigest()}")

# 找 DH_USERNAME 行
text = data.decode("utf-8")
for i, line in enumerate(text.splitlines(), 1):
    if "DH_USERNAME" in line or "DH_PASSWORD" in line or "DH1_" in line or "懂火系统账号" in line:
        print(f"  line {i}: {line!r}")
