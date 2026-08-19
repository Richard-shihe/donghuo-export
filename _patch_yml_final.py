"""给 DH1_xxx 替换处加注释，然后 commit + push"""
import os, re
path = r"c:\Users\H_RQ\Documents\trae_projects\A\.github\workflows\export_lindiao.yml"

with open(path, "rb") as fp:
    text = fp.read().decode("utf-8")

# 在 "# ===== 懂火系统账号" 这行下面加 2 行注释
old_header = "# ===== 懂火系统账号（在仓库 Secrets 中配置） =====\r\n          DH_USERNAME: ${{ secrets.DH1_USERNAME }}"
new_header = """# ===== 懂火系统账号（在仓库 Secrets 中配置） =====
          # 注：DH1_xxx 是旧账号 HONG@shihe.donghuo，已在 GitHub Actions 上建立境外 IP 信任（8/13 验证可拿数据）
          #     DH_xxx 是新账号 R@shihe.donghuo，本地可用但 GitHub Actions 上被风控返回 0 条
          DH_USERNAME: ${{ secrets.DH1_USERNAME }}""".replace("\n", "\r\n")

if old_header in text:
    text = text.replace(old_header, new_header, 1)
    with open(path, "wb") as fp:
        fp.write(text.encode("utf-8"))
    print("注释已加")
else:
    print("找不到原 header（可能已有注释）")

# 用 git 命令确认有改动
import subprocess
r = subprocess.run(["git", "diff", "--stat", path], capture_output=True, text=True)
print("\ngit diff --stat:")
print(r.stdout)
