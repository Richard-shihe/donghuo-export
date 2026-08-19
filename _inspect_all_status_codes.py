"""Look at td[3] across all rows to understand status mapping."""
import os, glob, re
from bs4 import BeautifulSoup
from collections import Counter

files = sorted(glob.glob(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       "_probe_query_resp_*.html")))
for p in files[-3:]:
    print(f"\n=== {os.path.basename(p)} ===")
    with open(p, "r", encoding="utf-8") as f:
        body = f.read()
    m = re.search(r'<input[^>]+id=["\']total["\'][^>]*value=["\'](\d+)["\']', body, re.I)
    print(f"total: {m.group(1) if m else '?'}")
    soup = BeautifulSoup(body, "html.parser")
    trs = soup.find_all("tr")
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
        # 打印前 5 行的 code/text/raw
    print(f"labelcode 分布: {dict(code_counter)}")
    print(f"text 分布: {dict(text_counter)}")
    # 打印前 5 行 td[3] 的 raw
    print("前 5 行 td[3] raw HTML:")
    for tr in trs[:5]:
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 5:
            tds = tr.find_all("td")
        if len(tds) < 5:
            continue
        print(f"  code={tds[3].get('labelcode')}  text={tds[3].get_text(strip=True)!r}  raw={str(tds[3])[:120]}")
