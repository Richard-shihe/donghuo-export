"""Dump the full first row HTML from the contractStatus=80 response to find any hidden status field."""
import os, glob, re
from bs4 import BeautifulSoup

base = os.path.dirname(os.path.abspath(__file__))
files = sorted(glob.glob(os.path.join(base, "_probe_query_结案_*.html")))
p = files[-1]
print(f"file: {p}")
with open(p, "r", encoding="utf-8") as f:
    body = f.read()
soup = BeautifulSoup(body, "html.parser")
trs = soup.find_all("tr")
print(f"parsed {len(trs)} <tr>")
tr = trs[0]
print("\n=== first <tr> full HTML (first 3000 chars) ===")
print(str(tr)[:3000])

# 找所有 hidden input 的 id 和 value
print("\n=== all hidden inputs in first row ===")
for inp in tr.find_all("input"):
    iid = inp.get("id") or ""
    iname = inp.get("name") or ""
    ival = inp.get("value") or ""
    iattrs = {k: v for k, v in inp.attrs.items() if k not in ("id","name","value","type")}
    print(f"  id={iid!r}  name={iname!r}  value={ival!r}  extra={iattrs}")

# 全局搜 80 / 结案 / 合同
print("\n=== body 全局搜 80 / 结案 / 合同完成 ===")
for kw in ["80", "结案", "合同完成", "close", "closed"]:
    idx = body.find(kw)
    print(f"  {kw!r}: idx={idx}")
