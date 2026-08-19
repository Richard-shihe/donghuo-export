"""直接用 Python 改 yml 文件，绕过 Edit 工具的 bug"""
import os
path = r"c:\Users\H_RQ\Documents\trae_projects\A\.github\workflows\export_lindiao.yml"

with open(path, "rb") as fp:
    data = fp.read().decode("utf-8")

old_block = """          # ===== 懂火系统账号（在仓库 Secrets 中配置） =====
          DH_USERNAME: ${{ secrets.DH_USERNAME }}
          DH_PASSWORD: ${{ secrets.DH_PASSWORD }}"""

new_block = """          # ===== 懂火系统账号（在仓库 Secrets 中配置） =====
          # 注：DH1_xxx 是旧账号 HONG@shihe.donghuo，已在 GitHub Actions 上建立境外 IP 信任（8/13 验证可拿数据）
          #     DH_xxx 是新账号 R@shihe.donghuo，本地可用但 GitHub Actions 上被风控返回 0 条
          DH_USERNAME: ${{ secrets.DH1_USERNAME }}
          DH_PASSWORD: ${{ secrets.DH1_PASSWORD }}"""

if old_block in data:
    new_data = data.replace(old_block, new_block, 1)
    with open(path, "wb") as fp:
        fp.write(new_data.encode("utf-8"))
    print("替换成功")
    # 校验
    with open(path, "rb") as fp:
        verify = fp.read().decode("utf-8")
    if "DH1_USERNAME" in verify and "HONG@shihe.donghuo" in verify:
        print("校验通过：文件已包含 DH1_USERNAME 和 HONG 注释")
    else:
        print("校验失败")
else:
    print("未找到要替换的块")
    # 看一下当前文件 DH_USERNAME 那部分的内容
    for i, line in enumerate(data.splitlines(), 1):
        if "DH_USERNAME" in line or "DH_PASSWORD" in line:
            print(f"  line {i}: {line!r}")
