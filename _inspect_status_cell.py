"""Check raw HTML for the status cell - is it showing real contractStatus or something else?"""
import os, glob, re
from bs4 import BeautifulSoup

files = sorted(glob.glob(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       "_probe_query_resp_*.html")))
# 取最近 2 个（全量 + contractStatus=80）
for p in files[-2:]:
    print(f"\n=== {os.path.basename(p)} ===")
    with open(p, "r", encoding="utf-8") as f:
        body = f.read()
    # 找 total
    m = re.search(r'<input[^>]+id=["\']total["\'][^>]*value=["\'](\d+)["\']', body, re.I)
    print(f"total: {m.group(1) if m else '?'}")

    soup = BeautifulSoup(body, "html.parser")
    trs = soup.find_all("tr")
    if not trs:
        continue
    tr = trs[0]
    tds = tr.find_all("td", recursive=False)
    if len(tds) < 5:
        tds = tr.find_all("td")
    print(f"first tr: {len(tds)} td")
    # 打印 td[3] (订单状态) 的完整原始 HTML
    if len(tds) > 3:
        print(f"\ntd[3] raw HTML (订单状态):")
        print(str(tds[3])[:600])
    # 找 td 里所有疑似状态字段：data-*, title, hidden input
    print("\n--- 全部 td 的 raw HTML（第 1 行）---")
    for i, td in enumerate(tds[:18]):
        s = str(td).strip()
        if len(s) > 200:
            s = s[:200] + "..."
        print(f"  td[{i}]: {s}")
