#!/usr/bin/env python3
"""Inspect _probe_production.html for selectForm fields + table headers + pagination."""
import re
import os

p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_probe_production.html")
with open(p, "r", encoding="utf-8") as f:
    html = f.read()

print("=== selectForm assignments (first 40) ===")
m = re.findall(r'selectForm[\.\[][^;\n]{1,140}', html)
for x in m[:40]:
    print(" ", x.strip())

print("\n=== pagination params ===")
m2 = re.findall(r'(pageNum|pageSize|pageNo|limit|offset|page_num|page_size)["\']?\s*[:=]\s*[^,;\n\}]{1,40}', html)
for x in m2[:25]:
    print(" ", x.strip())

print("\n=== <th ...>...</th> headers (first 40) ===")
ths = re.findall(r'<th[^>]*>(.*?)</th>', html, re.S)
for i, t in enumerate(ths[:40], 1):
    t = re.sub(r'<[^>]+>', '', t).strip()
    if t:
        print(f"  [{i}] {t}")

print("\n=== input/select fields with name= (first 40) ===")
fields = re.findall(r'<(?:input|select)[^>]+name=["\']([^"\']+)["\']', html)
seen = []
for n in fields:
    if n not in seen:
        seen.append(n)
for i, n in enumerate(seen[:40], 1):
    print(f"  [{i}] {n}")

print("\n=== 隐藏字段 id=xxx value= (含 settleUserNum 等) ===")
hids = re.findall(r'<input[^>]+type=["\']hidden["\'][^>]*>', html, re.I)
for h in hids[:30]:
    idm = re.search(r'id=["\']([^"\']+)["\']', h)
    valm = re.search(r'value=["\']([^"\']*)["\']', h)
    print(f"  id={idm.group(1) if idm else '?'}  value={valm.group(1) if valm else ''}")
