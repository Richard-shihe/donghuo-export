"""解码 tblUzkPskttsBa0W 表里 SourceID 字段，看里面是什么"""
import base64

sources = [
    "NzU3MjIxODIwMzg0MjUyNzIzNjpyZWN2MnJyd3ZTdnpXazo0ZWQ0ZWVkZWUwNDI4NjIwNDUyMzE5MTJhOGZiNWFkNDox",
    "NzU3MjIxODIwMzg0MjUyNzIzNjpyZWN2MnJpbUFtM01vdzo2MTllYmQyOWI0MDg4OGJiOWUyMjFmMGFlNDZlNDNjNjox",
    "NzU3MjIxODIwMzg0MjUyNzIzNjpyZWN2MVF2RzNGZVRiMDpmMDRjZjgyOGFlMzQ0Mjk3MDllNGU0YzkzMTJmMjQ2Nzox",
]
for i, s in enumerate(sources):
    decoded = base64.b64decode(s).decode("utf-8", errors="replace")
    print(f"[{i}] {decoded}")

# 再看一下 record_id 的对应关系
print("\n表里 record_id 对应:")
print("  [0] recv2rwZdad6FJ")
print("  [1] recv2rwZdamoFv")
print("  [2] recv2rwZda3fx7")
