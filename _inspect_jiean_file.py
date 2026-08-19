"""Inspect the contractStatus=80 (结案) saved response."""
import os, glob, re
from bs4 import BeautifulSoup
from collections import Counter

base = os.path.dirname(os.path.abspath(__file__))
# 找两个最新的：全量 + 结案
files_quan = sorted(glob.glob(os.path.join(base, "_probe_query_全量_*.html")))
files_jiean = sorted(glob.glob(os.path.join(base, "_probe_query_结案_*.html")))

for label, files in [("全量", files_quan), ("结案=80", files_jiean)]:
    if not files:
        print(f"\n{label}: 没找到文件")
        continue
    p = files[-1]
    print(f"\n=== {label}: {os.path.basename(p)} ===")
    with open(p, "r", encoding="utf-8") as f:
        body = f.read()
    m = re.search(r'<input[^>]+id=["\']total["\'][^>]*value=["\'](\d+)["\']', body, re.I)
    print(f"total: {m.group(1) if m else '?'}")
    soup = BeautifulSoup(body, "html.parser")
    trs = soup.find_all("tr")
    print(f"parsed {len(trs)} <tr>")
    code_counter = Counter()
    text_counter = Counter()
    for tr in trs:
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 5:
            tds = tr.find_all("td")
        if len(tds) < 5:
            continue
        td3 = tds[3]
        code = td3.get("labelcode") or ""
        text = td3.get_text(strip=True)
        code_counter[code] += 1
        if text:
            text_counter[text] += 1
    print(f"td[3] labelcode 分布: {dict(code_counter)}")
    print(f"td[3] text 分布: {dict(text_counter)}")
    # 前 3 行 raw
    print("前 3 行 td[3] raw:")
    for tr in trs[:3]:
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 5:
            tds = tr.find_all("td")
        if len(tds) < 5:
            continue
        print(f"  {str(tds[3])[:150]}")
