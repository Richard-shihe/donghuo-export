"""检查行尾并改用 \r\n 兼容"""
import os, re
path = r"c:\Users\H_RQ\Documents\trae_projects\A\.github\workflows\export_lindiao.yml"

with open(path, "rb") as fp:
    data = fp.read()

# 检查行尾
crlf_count = data.count(b"\r\n")
lf_count = data.count(b"\n")
print(f"CRLF count: {crlf_count}, LF total count: {lf_count}")
print(f"file is CRLF: {crlf_count == lf_count}")

# 找 line 49 的字节范围
text = data.decode("utf-8")
idx_49 = text.find("DH_USERNAME: ${{ secrets.DH_USERNAME }}")
print(f"\nDH_USERNAME 位置: {idx_49}")
print(f"前后 10 字节: {text[idx_49-10:idx_59] if False else text[idx_49-10:idx_49+50]!r}")

# 试着用统一的方式（先转 LF 再做正则）
text_lf = text.replace("\r\n", "\n")
pattern = r"          DH_USERNAME: \$\{\{ secrets\.DH_USERNAME \}\}\n          DH_PASSWORD: \$\{\{ secrets\.DH_PASSWORD \}\}"
match = re.search(pattern, text_lf)
print(f"\nLF 模式匹配: {match is not None}")
if match:
    new_text = "          DH_USERNAME: ${{ secrets.DH1_USERNAME }}\n          DH_PASSWORD: ${{ secrets.DH1_PASSWORD }}"
    new_text_lf = text_lf.replace(match.group(0), new_text, 1)
    # 写回 CRLF
    new_text_crlf = new_text_lf.replace("\n", "\r\n")
    with open(path, "wb") as fp:
        fp.write(new_text_crlf.encode("utf-8"))
    print("写入完成")
    # 校验
    with open(path, "rb") as fp:
        verify = fp.read().decode("utf-8")
    if "DH1_USERNAME" in verify:
        print("✅ 校验通过：文件已包含 DH1_USERNAME")
    else:
        print("❌ 校验失败")
