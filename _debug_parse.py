"""Check why tr regex didn't match - read the saved response file."""
import re, os, glob

files = sorted(glob.glob(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       "_probe_query_resp_*.html")))
p = files[-1]
print(f"file: {p}")
with open(p, "r", encoding="utf-8") as f:
    body = f.read()
print(f"len: {len(body)}")
print(f"first 200 chars: {body[:200]!r}")

# 试几种 tr 解析方式
m1 = re.findall(r'<tr[^>]*>.*?</tr>', body, re.S)
print(f"\nmethod1 <tr>...</tr>: {len(m1)} matches")
m2 = re.findall(r'<tr', body)
print(f"method2 <tr count: {len(m2)}")
m3 = re.findall(r'<tbody[^>]*>(.*?)</tbody>', body, re.S)
print(f"method3 <tbody>...</tbody>: {len(m3)} matches, first len={len(m3[0]) if m3 else 0}")

# 用 BeautifulSoup 试
try:
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(body, "html.parser")
    trs = soup.find_all("tr")
    print(f"\nBeautifulSoup tr count: {len(trs)}")
    if trs:
        first = trs[0]
        tds = first.find_all("td")
        print(f"first tr has {len(tds)} td")
        for i, td in enumerate(tds):
            t = td.get_text(strip=True)
            print(f"  td[{i}]: {t[:80]!r}")
except ImportError:
    print("\n(no bs4)")
