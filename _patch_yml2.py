"""用更宽松的匹配替换 yml 中的两行"""
import os, re
path = r"c:\Users\H_RQ\Documents\trae_projects\A\.github\workflows\export_lindiao.yml"

with open(path, "rb") as fp:
    data = fp.read().decode("utf-8")

# 用正则匹配 DH_USERNAME 和 DH_PASSWORD 两行
pattern = r"(          DH_USERNAME: )\$\{\{ secrets\.DH_USERNAME \}\}(\n          DH_PASSWORD: )\$\{\{ secrets\.DH_PASSWORD \}\}"
match = re.search(pattern, data)
if not match:
    print("正则没匹配到")
    # 直接打印 line 49-50 的字节
    lines = data.splitlines()
    for i in (48, 49, 50):
        if i < len(lines):
            print(f"  line {i+1} bytes: {lines[i].encode('utf-8')!r}")
else:
    print(f"匹配到，位置: {match.start()}-{match.end()}")
    new_text = (match.group(1) + "${{ secrets.DH1_USERNAME }}" +
                match.group(2) + "${{ secrets.DH1_PASSWORD }}")
    new_data = data[:match.start()] + new_text + data[match.end():]
    with open(path, "wb") as fp:
        fp.write(new_data.encode("utf-8"))
    print("替换成功")
    # 校验
    with open(path, "rb") as fp:
        verify = fp.read().decode("utf-8")
    if "DH1_USERNAME" in verify:
        print("校验通过")
    else:
        print("校验失败")
