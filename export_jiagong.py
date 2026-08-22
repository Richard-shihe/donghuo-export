#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火 → 飞书多维表「加工单」表  增量导出脚本（事情 A）
独立于 import_jiagong.py（事情 B：飞书→懂火 写成品明细）

逻辑：
  1. 懂火加工单列表页分页拉全量加工单基本信息
  2. 打开每个加工单的详情页，解析：
       - 表头字段（费用名称/服务商/税率/金额/备注...）
       - 原料明细行（zid=明细行ID, 捆包号/品名/规格/产地/材质/件数/重量）
  3. 飞书加工单表按「明细行ID」去重增量：
       - 新明细行 → INSERT 新增一条
       - 已存在（明细行ID匹配）→ UPDATE 该条记录的字段值（不删除飞书现有记录）
       - 飞书里有但懂火里没有的明细行 → 保留不删（符合"增量更新不要删"）

命令行参数：
  --dry-run        预览模式，只打印计划，不写飞书
  --since J2026-0601   只处理加工单号 >= 这个值的（按字典序，可选）
  --limit N        只处理 N 个加工单（调试用）

环境变量：
  DH_USERNAME / DH_PASSWORD     懂火登录账号（必填）
  FEISHU_APP_ID / FEISHU_APP_SECRET
  BITABLE_APP_TOKEN             默认 CpYZbPbi3a0qo0smUhDcrKFgnGc
  BITABLE_JIAGONG_TABLE_ID      加工单表 table_id（可默认自动匹配表名"加工单"）
"""
import os
import sys
import re
import json
import time
import datetime
import argparse
import traceback
from collections import defaultdict

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from donghuo_login import login_donghuo

BASE_URL = "https://erpa.donghuo.vip"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
JIAGONG_LIST_API = f"{BASE_URL}/model/admin/caigou/m_jiagon/getlist"
DETAIL_PAGE = f"{BASE_URL}/view/admin/jiagon/v_jigondan"
AJAX_HEADERS = {"X-Requested-With": "XMLHttpRequest"}
NO_PROXY = {"http": None, "https": None}


# ============ 工具 ============
def env(name: str, default: str = "") -> str:
    return str(os.environ.get(name, default) or "").strip()


def _flat_v(v) -> str:
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, (int, float, bool)):
        return str(v).strip()
    if isinstance(v, list):
        parts = []
        for seg in v:
            if isinstance(seg, dict):
                if "text" in seg:
                    parts.append(str(seg["text"]).strip())
                elif "name" in seg:
                    parts.append(str(seg["name"]).strip())
            elif isinstance(seg, str):
                parts.append(seg.strip())
        return "".join(p for p in parts if p).strip()
    return str(v).strip()


# ============ 飞书 API ============
def feishu_token() -> str:
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={
        "app_id": env("FEISHU_APP_ID", ""),
        "app_secret": env("FEISHU_APP_SECRET", ""),
    }, timeout=15, proxies=NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"飞书 token 失败: {data}")
    return str(data["tenant_access_token"])


def bitable_auto_find_table(tok: str, app_token: str, name: str) -> str:
    r = requests.get(f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables",
                     headers={"Authorization": f"Bearer {tok}"},
                     timeout=15, proxies=NO_PROXY)
    items = (r.json().get("data") or {}).get("items") or []
    for t in items:
        if t.get("name") == name:
            return t["table_id"]
    for t in items:
        if name in (t.get("name") or ""):
            return t["table_id"]
    return ""


def bitable_read_all(tok: str, app_token: str, table_id: str) -> list:
    rows = []
    page_token = ""
    for _ in range(200):
        url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(url, params=params,
                         headers={"Authorization": f"Bearer {tok}"},
                         timeout=30, proxies=NO_PROXY)
        d = r.json()
        items = (d.get("data") or {}).get("items") or []
        for it in items:
            rec = {"_record_id": it.get("record_id", "")}
            for k, v in (it.get("fields") or {}).items():
                rec[k] = _flat_v(v)
            rows.append(rec)
        d_data = d.get("data") or {}
        if not d_data.get("has_more") or not d_data.get("page_token"):
            break
        page_token = d_data["page_token"]
    return rows


def bitable_insert(tok: str, app_token: str, table_id: str,
                   fields_list: list[dict]) -> tuple[int, int]:
    """批量插入记录。fields_list 是 [{字段: 值}, ...]，返回 (成功数, 失败数)"""
    if not fields_list:
        return 0, 0
    ok_c, fail_c = 0, 0
    # 飞书批量接口最多 500 条 / 次
    batch_size = 500
    for i in range(0, len(fields_list), batch_size):
        batch = fields_list[i:i + batch_size]
        url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_create"
        payload = {"records": [{"fields": f} for f in batch]}
        r = requests.post(url, json=payload,
                          headers={"Authorization": f"Bearer {tok}"},
                          timeout=30, proxies=NO_PROXY)
        resp = r.json()
        if resp.get("code") == 0:
            ok_c += len((resp.get("data") or {}).get("records") or [])
        else:
            # 失败的话逐条重试，记录单条失败
            for f in batch:
                url2 = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
                r2 = requests.post(url2, json={"fields": f},
                                   headers={"Authorization": f"Bearer {tok}"},
                                   timeout=15, proxies=NO_PROXY)
                if r2.json().get("code") == 0:
                    ok_c += 1
                else:
                    fail_c += 1
                    print(f"  ⚠️  插入失败: {json.dumps(f, ensure_ascii=False)[:200]}")
    return ok_c, fail_c


def bitable_update(tok: str, app_token: str, table_id: str, record_id: str,
                   fields: dict) -> bool:
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/{record_id}"
    r = requests.put(url, json={"fields": fields},
                     headers={"Authorization": f"Bearer {tok}"},
                     timeout=15, proxies=NO_PROXY)
    return r.json().get("code") == 0


# ============ 懂火解析 ============
def parse_html_header(html: str) -> dict:
    def get_sel(sel_id):
        m = re.search(rf'<select[^>]*?id=["\']{sel_id}["\'][^>]*?>(.*?)</select>', html, re.S)
        if not m:
            return ""
        sel_html = m.group(1)
        m2 = re.search(r'<option[^>]*?value=["\']([^"\']*)["\'][^>]*?selected[^>]*?>([^<]*)</option>', sel_html, re.S)
        if not m2:
            m2 = re.search(r'<option[^>]*?selected[^>]*?value=["\']([^"\']*)["\'][^>]*?>([^<]*)</option>', sel_html, re.S)
        return m2.group(1).strip() if m2 else ""

    def get_inp(fid):
        m = re.search(rf'<input[^>]*?id=["\']{fid}["\'][^>]*?value=["\']([^"\']*)["\']', html)
        return m.group(1).strip() if m else ""

    jg_beizhu = ""
    m = re.search(r'id=["\']jg_beizhu["\'][^>]*?>(.*?)</div>', html, re.S)
    if m:
        jg_beizhu = re.sub(r'<[^>]+>', '', m.group(1)).replace('&nbsp;', ' ').strip()
        import html as _h
        jg_beizhu = _h.unescape(jg_beizhu)

    return {
        "服务商名称": get_sel("fyname"),   # 懂火 fyname=费用名称，飞书字段叫"服务商名称"（来自字段对应关系表）
        "结算单位": get_sel("jsdanwei"),   # 懂火 jsdanwei=服务商名称，飞书对应"结算单位"
        "税率": get_sel("shuilv"),
        "金额": get_inp("jiner"),
        "备注": get_inp("beizhu"),
        "制单人": get_sel("funame"),
        "所属公司": get_sel("fcompany"),
        "加工备注": jg_beizhu,
    }


def parse_material_rows(html: str) -> list[dict]:
    """解析原料明细行（和飞书加工单表现有 29 条记录的结构一致）。
    返回 list[{
        "明细行ID": zid,       # 懂火的 input name="zid" value
        "捆包号": 列1,
        "品名": 列2,
        "规格": 列3,
        "产地": 列4,
        "材质": 列5,
        "件(张)数": 列6,
        "重量(吨)": 列7  # 字符串
    }]
    """
    import html as _h
    rows = []
    for tr_m in re.finditer(r'<tr[^>]*>(.*?)</tr>', html, re.S):
        tr = tr_m.group(1)
        if 'name="zid"' not in tr and 'name=\'zid\'' not in tr:
            continue
        # 提取 zid
        zid_m = re.search(r'<input[^>]*?name=["\']zid["\'][^>]*?value=["\']([^"\']*)["\']', tr)
        zid = zid_m.group(1).strip() if zid_m else ""
        if not zid:
            continue
        # 提取 td 的非空纯文本列，结构固定为 10 列：
        #   [0] zid hidden（跳过）, [1]捆包号, [2]品名, [3]规格, [4]产地,
        #   [5]材质, [6]件数, [7]重量, [8]备注(可能空), [9]操作按钮
        # 去掉[0]后，前7个有效文本列就是 [捆包号,品名,规格,产地,材质,件数,重量]
        tds = re.findall(r'<td[^>]*>(.*?)</td>', tr, re.S)
        cells = []
        for td in tds:
            if re.search(r'<input[^>]*?name=["\']zid["\']', td):
                continue
            clean = _h.unescape(re.sub(r'<[^>]+>', '', td)).replace('&nbsp;', ' ').strip()
            cells.append(clean)
        # 过滤后 cells 索引对应：0=捆包号,1=品名,2=规格,3=产地,4=材质,5=件数,6=重量,7=备注(可能空),8=操作按钮
        # 只取前 7 个位置，空的留空串（不要合并掉空列，否则错位）
        cells = (cells + [""] * 7)[:7]
        rows.append({
            "明细行ID": zid,
            "捆包号": cells[0],
            "品名": cells[1],
            "规格": cells[2],
            "产地": cells[3],
            "材质": cells[4],
            "件(张)数": cells[5],
            "重量(吨)": cells[6],
        })
    return rows


def get_all_jiagong_list(session) -> list[dict]:
    """分页取懂火加工单列表（基本字段）。返回 list[{加工单号, 制单人, 所属公司, ...}]"""
    result = []
    page = 1
    while True:
        r = session.post(JIAGONG_LIST_API,
                         data={"page": page, "limit": 100},
                         headers=AJAX_HEADERS, timeout=20)
        try:
            d = r.json()
        except Exception:
            print(f"  列表页 {page} 非JSON: {r.text[:150]}")
            break
        rows = d.get("root") or []
        if not rows:
            break
        for row in rows:
            dh = str(row.get("加工单号") or "").strip()
            if dh:
                # 把列表上有的字段直接带到内存中，详情页可能拿不到"加工日期/状态"
                info = {k: str(v).strip() for k, v in row.items() if v is not None}
                info["加工单号"] = dh
                result.append(info)
        total_pages = int(d.get("pgtotal") or 1)
        if page >= total_pages:
            break
        page += 1
        time.sleep(0.1)
    print(f"[懂火] 列表页拉取 {len(result)} 个加工单")
    return result


# ============ 主流程 ============
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--since", default="", help="只处理加工单号>=这个值（字典序），如 J2026-0601")
    ap.add_argument("--limit", type=int, default=0, help="最多处理 N 个加工单（调试用）")
    args = ap.parse_args()
    dry_run = args.dry_run
    since = args.since.strip()
    limit = args.limit

    t0 = time.time()
    app_token = env("BITABLE_APP_TOKEN", "CpYZbPbi3a0qo0smUhDcrKFgnGc")

    print(f"====== 懂火 → 飞书加工单表 增量导出 {'(DRY-RUN 预览模式)' if dry_run else ''} ======")
    if since:
        print(f"  since 过滤: 加工单号 >= {since}")
    if limit > 0:
        print(f"  limit: 只处理 {limit} 个加工单")

    # 1. 飞书：拿到加工单表，读现有记录
    tok = feishu_token()
    table_id = env("BITABLE_JIAGONG_TABLE_ID") or bitable_auto_find_table(tok, app_token, "加工单")
    if not table_id:
        print("[错误] 找不到飞书「加工单」表，请设置 BITABLE_JIAGONG_TABLE_ID")
        return 2
    print(f"  飞书加工单表: {table_id}")
    existing_rows = bitable_read_all(tok, app_token, table_id)
    # 建立"明细行ID → record_id"索引（去重用）
    existing_by_id: dict[str, dict] = {}
    for r in existing_rows:
        mid = (r.get("明细行ID") or "").strip()
        if mid and mid not in existing_by_id:
            existing_by_id[mid] = r
    print(f"  飞书现有 {len(existing_rows)} 条记录，含明细行ID={len(existing_by_id)} 条")

    # 2. 懂火：列表页
    print("\n[步骤] 登录懂火并拉加工单列表 ...")
    s = login_donghuo(username=env("DH_USERNAME", ""),
                      password=env("DH_PASSWORD", ""))
    if s is None:
        print("[错误] 懂火登录失败")
        return 1
    list_rows = get_all_jiagong_list(s)
    if since:
        list_rows = [x for x in list_rows if x["加工单号"] >= since]
    if limit > 0:
        list_rows = sorted(list_rows, key=lambda x: x["加工单号"])[:limit]
    print(f"  过滤后处理 {len(list_rows)} 个加工单: {[x['加工单号'] for x in list_rows[:10]]}"
          f"{'...' if len(list_rows) > 10 else ''}")

    # 3. 逐个加工单打开详情页
    stats = {
        "insert_rows": 0,     # 需要 INSERT 的明细行数
        "update_rows": 0,     # 需要 UPDATE 的明细行数
        "unchanged_rows": 0,  # 已完全一致，跳过
    }
    to_insert: list[dict] = []   # 待 INSERT 的 fields dict 列表
    to_update: list[tuple[str, dict]] = []  # (record_id, fields_to_update)
    processed_danhaos: set[str] = set()

    for idx, basic in enumerate(sorted(list_rows, key=lambda x: x["加工单号"]), 1):
        danhao = basic["加工单号"]
        processed_danhaos.add(danhao)
        print(f"\n--- [{idx}/{len(list_rows)}] {danhao} ---")
        try:
            r = s.get(DETAIL_PAGE, params={"jg_danhao": danhao}, timeout=20)
            if r.status_code != 200:
                print(f"  ⚠️  HTTP {r.status_code}，跳过")
                continue
            html = r.text
            hdr = parse_html_header(html)
            mats = parse_material_rows(html)
            if not mats:
                # 加工单只有表头、没有原料明细（正常，比如刚建好还没加原料）
                print(f"  列表页基本字段已拿到，但没有原料明细行（可能是空单）")
            else:
                print(f"  原料明细: {len(mats)} 行")

            # 合并"列表页基本字段"和"详情页表头字段"，列表页优先（制单人/所属公司/加工日期/状态更准）
            merged_hdr: dict = {}
            # 先塞详情页解析到的
            for k, v in hdr.items():
                if v:
                    merged_hdr[k] = v
            # 再用列表页的覆盖（制单人在列表页可能有不同值）
            for dh_field, fs_field in [
                ("制单人", "制单人"),
                ("所属公司", "所属公司"),
                ("加工日期", ""),  # 飞书加工单表没"加工日期"字段，跳过
                ("新增时间", ""),  # 同上
                ("状态", ""),      # 同上
            ]:
                if not fs_field:
                    continue
                v = str(basic.get(dh_field) or "").strip()
                if v:
                    merged_hdr[fs_field] = v

            # 飞书加工单表中，表头字段在同一行记录（每个明细行都会复制一遍表头）
            # 所以每一行原料明细都附带表头
            for mat in mats:
                mid = mat["明细行ID"]
                # 构造要写入飞书的完整字段
                rec: dict = {"加工单号": danhao}
                for k, v in merged_hdr.items():
                    if v:
                        rec[k] = v
                for k in ("明细行ID", "捆包号", "品名", "规格", "产地", "材质", "件(张)数"):
                    if mat.get(k):
                        rec[k] = mat[k]
                # 重量(吨) 要数字
                wt = (mat.get("重量(吨)") or "").strip()
                if wt:
                    try:
                        rec["重量(吨)"] = float(wt)
                    except ValueError:
                        rec["重量(吨)"] = wt

                # 对比飞书已有记录
                if mid in existing_by_id:
                    old = existing_by_id[mid]
                    # 计算差异
                    diffs = {}
                    for k, new_v in rec.items():
                        old_v = old.get(k)
                        if k == "重量(吨)":
                            # 都转 float 比较
                            try:
                                nf = float(new_v) if new_v not in (None, "") else None
                                of = float(old_v) if old_v not in (None, "") else None
                                if nf != of:
                                    diffs[k] = new_v
                            except (ValueError, TypeError):
                                if str(new_v).strip() != str(old_v).strip():
                                    diffs[k] = new_v
                        else:
                            if str(new_v).strip() != str(old_v).strip():
                                diffs[k] = new_v
                    if diffs:
                        stats["update_rows"] += 1
                        print(f"  ✏️  更新行 明细行ID={mid} 变更: {list(diffs.keys())}")
                        to_update.append((old["_record_id"], diffs))
                    else:
                        stats["unchanged_rows"] += 1
                else:
                    stats["insert_rows"] += 1
                    kb = rec.get("捆包号", "")
                    wt = rec.get("重量(吨)", "")
                    print(f"  ➕ 新增行 捆包号={kb} 明细行ID={mid} 重量={wt}")
                    to_insert.append(rec)
        except Exception as e:
            print(f"  ❌ 异常: {e}")
            traceback.print_exc()
        time.sleep(0.3)

    s.close()

    # 4. 实际写入飞书
    print(f"\n[步骤] 写入飞书加工单表 (INSERT {len(to_insert)} + UPDATE {len(to_update)}) ...")
    ins_ok, ins_fail, up_ok, up_fail = 0, 0, 0, 0
    if dry_run:
        print(f"  [DRY-RUN] 跳过实际写入")
    else:
        if to_insert:
            ins_ok, ins_fail = bitable_insert(tok, app_token, table_id, to_insert)
            print(f"  INSERT: 成功 {ins_ok}, 失败 {ins_fail}")
        if to_update:
            for rid, fields in to_update:
                if bitable_update(tok, app_token, table_id, rid, fields):
                    up_ok += 1
                else:
                    up_fail += 1
                    print(f"  ⚠️  UPDATE {rid} 失败")
            print(f"  UPDATE: 成功 {up_ok}, 失败 {up_fail}")

    # 5. 汇总
    elapsed = time.time() - t0
    print("\n" + "=" * 60)
    print(f"📊 导出汇总（耗时 {elapsed:.1f}s）{' DRY-RUN' if dry_run else ''}")
    print(f"  懂火加工单: 处理 {len(processed_danhaos)} 个")
    print(f"  原料明细行: INSERT {stats['insert_rows']}, UPDATE {stats['update_rows']}, "
          f"未变 {stats['unchanged_rows']}")
    if not dry_run:
        print(f"  飞书实际: INSERT ok={ins_ok} fail={ins_fail}, UPDATE ok={up_ok} fail={up_fail}")
    print(f"  飞书原有记录: {len(existing_rows)} 条（懂火没有的明细行会保留不删）")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n[中断]")
        raise SystemExit(130)
    except Exception as e:
        print(f"\n[致命异常] {e}")
        traceback.print_exc()
        raise SystemExit(99)
