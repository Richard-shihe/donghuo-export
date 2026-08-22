#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
飞书多维表 → 懂火加工单成品明细 批量写入脚本
(两张表：加工单表头 + 加工成品录入，按加工单号关联)

运行方式:
  # 预览模式（DRY-RUN），只打印计划，不实际修改
  python import_jiagong.py --dry-run

  # 实际执行
  python import_jiagong.py

依赖环境变量:
  FEISHU_APP_ID / FEISHU_APP_SECRET
  BITABLE_APP_TOKEN             (多维表 app_token)
  BITABLE_HEADER_TABLE_ID       (加工单表头 表ID, 可默认自动查找)
  BITABLE_FINISHED_TABLE_ID     (加工成品录入 表ID, 可默认自动查找)
  DH_USERNAME / DH_PASSWORD     (懂火登录)
  DRY_RUN=1                     预览模式
  IMPORT_LIMIT=5                最多处理 N 个加工单 (调试用)
"""
import os
import sys
import re
import json
import time
import datetime
from collections import defaultdict

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from donghuo_login import login_donghuo

BASE_URL = "https://erpa.donghuo.vip"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
DETAIL_PAGE = f"{BASE_URL}/view/admin/jiagon/v_jigondan"
UPDATEJG_API = f"{BASE_URL}/controller/admin/jiagon/c_jiagon/updatejg"
AJAX = {"X-Requested-With": "XMLHttpRequest"}
NO_PROXY = {"http": None, "https": None}


# ================== 工具 ==================
def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def js_escape(s: str) -> str:
    """模拟 JavaScript 的 escape()"""
    result = []
    for c in s:
        if c.isalnum() or c in "-_.!~*'()/":
            result.append(c)
        else:
            for byte in c.encode('utf-8'):
                result.append(f'%{byte:02X}')
    return ''.join(result)


def _bitable_flatten_value(v) -> str:
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


# ================== 飞书 ==================
def feishu_token():
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={
        "app_id": env("FEISHU_APP_ID", ""),
        "app_secret": env("FEISHU_APP_SECRET", "")
    }, timeout=15, proxies=NO_PROXY)
    return r.json()["tenant_access_token"]


def bitable_list_tables(tok: str, app_token: str) -> list:
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables"
    r = requests.get(url, headers={"Authorization": f"Bearer {tok}"},
                     timeout=15, proxies=NO_PROXY)
    return (r.json().get("data") or {}).get("items") or []


def bitable_auto_find_table(tok: str, app_token: str,
                            keyword: str, default_table_id: str = "") -> str:
    """根据关键字自动找表名匹配的 table_id"""
    if default_table_id:
        return default_table_id
    tables = bitable_list_tables(tok, app_token)
    # 先精确匹配
    for t in tables:
        if t.get("name") == keyword:
            return t["table_id"]
    # 再模糊匹配
    for t in tables:
        if keyword in (t.get("name") or ""):
            return t["table_id"]
    return ""


def bitable_read(tok: str, app_token: str, table_id: str,
                 filter_formula: str = "") -> list:
    """读取多维表记录（支持公式筛选），返回 list[dict]（列名即 key）"""
    rows = []
    page_token = ""
    for _ in range(200):
        url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/search"
        body = {"page_size": 500}
        if page_token:
            body["page_token"] = page_token
        if filter_formula:
            body["filter"] = {"conjunction": "and",
                              "conditions": [{"field_name": k, "operator": op, "value": [val]}
                                             for k, op, val in _parse_filter(filter_formula)]}
        r = requests.post(url, json=body,
                          headers={"Authorization": f"Bearer {tok}"},
                          timeout=30, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") != 0:
            # search 接口失败，fallback 到普通 list
            return bitable_read_all(tok, app_token, table_id)
        items = (d.get("data") or {}).get("items") or []
        for it in items:
            rec = {"_record_id": it.get("record_id", "")}
            for k, v in (it.get("fields") or {}).items():
                rec[k] = _bitable_flatten_value(v)
            rows.append(rec)
        d_data = d.get("data") or {}
        if not d_data.get("has_more") or not d_data.get("page_token"):
            break
        page_token = d_data["page_token"]
    return rows


def _parse_filter(formula: str):
    """简单 parse '字段名=值' 形式"""
    conditions = []
    for part in formula.split("&&"):
        if "=" in part:
            k, v = part.split("=", 1)
            conditions.append((k.strip(), "is", v.strip()))
    return conditions


def bitable_read_all(tok, app_token, table_id) -> list:
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
                rec[k] = _bitable_flatten_value(v)
            rows.append(rec)
        d_data = d.get("data") or {}
        if not d_data.get("has_more") or not d_data.get("page_token"):
            break
        page_token = d_data["page_token"]
    return rows


def bitable_update(tok: str, app_token: str, table_id: str,
                   record_id: str, fields: dict):
    """更新某一条记录的字段"""
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/records/{record_id}")
    r = requests.put(url, json={"fields": fields},
                     headers={"Authorization": f"Bearer {tok}"},
                     timeout=15, proxies=NO_PROXY)
    return r.json()


# ================== 懂火详情解析 ==================
def parse_materials(html: str) -> dict:
    """返回 { 原料捆包号: {"zid":..., "token":...} }"""
    result = {}
    for m in re.finditer(
        r"addRow\('(\d+)','([^']+)','[^']*','[^']*','[^']*','[^']*','([^']+)'\)",
        html
    ):
        kb = m.group(2)
        token = m.group(3)
        zid_m = re.search(
            rf'<input[^>]*?name=["\']zid["\'][^>]*?value=["\'](\d+)["\'][^>]*?>\s*</td>\s*<td[^>]*>\s*{re.escape(kb)}',
            html
        )
        result[kb] = {
            "zid": zid_m.group(1) if zid_m else "",
            "token": token
        }
    return result


def parse_existing_finished(html: str) -> set:
    """返回成品表中已有的子捆包号集合（用于去重/增量追加）"""
    result = set()
    for tr_m in re.finditer(r"<tr[^>]*>(.*?)</tr>", html, re.S):
        tr = tr_m.group(1)
        if 'id="zid[]"' not in tr:
            continue
        kb_m = re.search(r'id="kunbaohao\[\]"[^>]*value="([^"]*)"', tr)
        if kb_m and kb_m.group(1).strip():
            result.add(kb_m.group(1).strip())
    return result


def parse_header(html: str) -> dict:
    def get_sel(sel_id):
        m = re.search(rf'<select[^>]*?id=["\']{sel_id}["\'][^>]*?>(.*?)</select>', html, re.S)
        if not m:
            return ""
        sel_html = m.group(1)
        sel_m = re.search(r'<option[^>]*?value=["\']([^"\']*)["\'][^>]*?selected[^>]*?>([^<]*)</option>', sel_html, re.S)
        if not sel_m:
            sel_m = re.search(r'<option[^>]*?selected[^>]*?value=["\']([^"\']*)["\'][^>]*?>([^<]*)</option>', sel_html, re.S)
        return sel_m.group(1).strip() if sel_m else ""

    def get_inp(fid):
        m = re.search(rf'<input[^>]*?id=["\']{fid}["\'][^>]*?value=["\']([^"\']*)["\']', html)
        return m.group(1).strip() if m else ""

    jg_beizhu = ""
    m = re.search(r'id=["\']jg_beizhu["\'][^>]*?>(.*?)</div>', html, re.S)
    if m:
        jg_beizhu = m.group(1).strip()

    return {"fyname": get_sel("fyname"), "jsdanwei": get_sel("jsdanwei"),
            "shuilv": get_sel("shuilv"), "jiner": get_inp("jiner"),
            "beizhu": get_inp("beizhu"), "funame": get_sel("funame"),
            "fcompany": get_sel("fcompany"), "jg_beizhu": jg_beizhu}


def fetch_full_options(session) -> dict:
    """通过 AJAX 接口获取 5 个表头下拉框的全量选项（注意：有些下拉是动态加载的，静态HTML里不全）。
    返回 { select_id: {value1, value2, ...} }
    """
    result = {}
    headers = AJAX

    # 1. 费用名称: /model/admin/m_load/getlist  leixin=费用名称
    try:
        r = session.post(f"{BASE_URL}/model/admin/m_load/getlist",
                         data={"leixin": "费用名称"}, headers=headers, timeout=15)
        d = r.json()
        if isinstance(d, list):
            values = {str(item.get('value') or item.get('key') or '').strip() for item in d}
            values.discard("")
            result["fyname"] = values
    except Exception:
        result["fyname"] = set()

    # 2. 服务商名称: /model/admin/m_load/list_fuwu  (返回 264 条，有 name/key 等多种字段)
    try:
        r = session.post(f"{BASE_URL}/model/admin/m_load/list_fuwu",
                         headers=headers, timeout=15)
        d = r.json()
        rows = d if isinstance(d, list) else d.get("rows") or []
        values = set()
        for row in rows:
            if isinstance(row, dict):
                for key in ("name", "title", "text", "value", "key"):
                    v = str(row.get(key) or "").strip()
                    if v:
                        values.add(v)
                for v in row.values():
                    if isinstance(v, str) and v.strip():
                        values.add(v.strip())
            elif isinstance(row, str):
                values.add(row.strip())
        values.discard("")
        result["jsdanwei"] = values
    except Exception:
        result["jsdanwei"] = set()

    # 3. 税率: 详情页 select 就是静态全量（7 个），这里留空，用 parse_select_options 合并补充
    result["shuilv"] = set()

    # 4. 所属公司: /model/admin/m_load/list_vip
    try:
        r = session.post(f"{BASE_URL}/model/admin/m_load/list_vip",
                         headers=headers, timeout=15)
        d = r.json()
        rows = d if isinstance(d, list) else d.get("rows") or d.get("data") or []
        values = set()
        for row in rows:
            if isinstance(row, dict):
                v = str(row.get("value") or row.get("key") or row.get("name") or "").strip()
                if v:
                    values.add(v)
            elif isinstance(row, str):
                values.add(row.strip())
        values.discard("")
        result["fcompany"] = values
    except Exception:
        result["fcompany"] = set()

    # 5. 添加人: /model/admin/m_load/list_user (所属公司变了可能要改)
    try:
        r = session.post(f"{BASE_URL}/model/admin/m_load/list_user",
                         headers=headers, timeout=15)
        d = r.json()
        rows = d if isinstance(d, list) else d.get("rows") or d.get("data") or []
        values = set()
        for row in rows:
            if isinstance(row, dict):
                v = str(row.get("value") or row.get("key") or row.get("name") or "").strip()
                if v:
                    values.add(v)
            elif isinstance(row, str):
                values.add(row.strip())
        values.discard("")
        result["funame"] = values
    except Exception:
        result["funame"] = set()

    return result


def parse_select_options(html: str) -> dict:
    """解析详情页 HTML 中 select 的静态 option（用于税率等非 AJAX 加载的字段）。
    返回 { select_id: {value1, value2, ...} }
    """
    result = {}
    for sel_id in ("fyname", "jsdanwei", "shuilv", "fcompany", "funame"):
        opts = set()
        m = re.search(rf'<select[^>]*?id=["\']{sel_id}["\'][^>]*?>(.*?)</select>', html, re.S)
        if m:
            for opt_m in re.finditer(r'<option[^>]*?value=["\']([^"\']*)["\'][^>]*?>([^<]*)</option>', m.group(1), re.S):
                val = opt_m.group(1).strip()
                if val:
                    opts.add(val)
        result[sel_id] = opts
    return result


def validate_header_options(merged_header: dict, ajax_options: dict, static_options: dict) -> list:
    """校验表头下拉字段的值是否在懂火选项里（AJAX 全量 ∪ 静态 = 完整）。
    返回错误消息列表（空列表表示全部通过）。
    """
    errors = []
    checks = [
        ("费用名称", "fyname", False),
        ("服务商名称", "jsdanwei", False),
        ("税率", "shuilv", False),
        ("所属公司", "fcompany", True),
        ("添加人", "funame", True),
    ]
    for label, sel_id, allow_empty in checks:
        val = (merged_header.get(sel_id) or "").strip()
        if not val:
            if allow_empty:
                continue
            errors.append(f"{label}为空（飞书未填写且懂火无默认值）")
            continue
        full = (ajax_options.get(sel_id) or set()) | (static_options.get(sel_id) or set())
        if not full:
            # 取不到选项，放行（避免误拦）
            continue
        if val not in full:
            sorted_opts = sorted(full)
            preview = "、".join(sorted_opts[:10]) + (f"... 等{len(sorted_opts)}项" if len(sorted_opts) > 10 else "")
            errors.append(f"{label}「{val}」不在懂火下拉选项中（可选: {preview}）")
    return errors


# ================== 主流程 ==================
def process_one_danhao(session, tok, app_token,
                       finished_table_id, header_table_id,
                       danhao: str, header_info: dict, header_record_id: str,
                       finished_rows: list,
                       ajax_options: dict,
                       dry_run: bool) -> tuple:
    """处理一个加工单号。
    参数 ajax_options: 登录后一次性 fetch_full_options 得到的全量下拉选项（避免每个加工单重复请求）
    返回:
      (finished_updates, header_update)
      finished_updates: [(record_id, status, result_text), ...]  成品表回写
      header_update: (record_id, status, result_text) 或 None     表头表回写
    """
    print(f"\n  加工单: {danhao}")

    # 1. 打开详情页
    r = session.get(DETAIL_PAGE, params={"jg_danhao": danhao}, timeout=15)
    if r.status_code != 200:
        msg = f"详情页 HTTP {r.status_code}"
        print(f"    ❌ {msg}")
        finished_updates = [(row["_record_id"], "失败", msg) for row in finished_rows]
        return finished_updates, (header_record_id, "失败", msg) if header_record_id else None

    html = r.text
    materials = parse_materials(html)
    existing = parse_existing_finished(html)
    dh_header = parse_header(html)
    static_options = parse_select_options(html)

    if not materials:
        msg = f"懂火中无原料捆包（无法添加成品，必须先有原料）"
        print(f"    ❌ {msg}")
        finished_updates = [(row["_record_id"], "失败", msg) for row in finished_rows]
        return finished_updates, (header_record_id, "失败", msg) if header_record_id else None

    print(f"    懂火原料: {list(materials.keys())}")
    print(f"    懂火现有成品: {len(existing)} 条 {sorted(existing)[:5]}...")

    # 2. 合并表头：飞书表头优先于懂火已有表头，再 fallback 默认
    merged_header = dict(dh_header)
    header_map = {
        "费用名称": "fyname",
        "服务商名称": "jsdanwei",
        "税率": "shuilv",
        "金额": "jiner",
        "所属公司": "fcompany",
        "添加人": "funame",
    }
    for col, key in header_map.items():
        v = header_info.get(col)
        if v:
            merged_header[key] = str(v).strip()
    if not merged_header.get("fcompany"):
        merged_header["fcompany"] = "上海砚启实业有限公司"
    if not merged_header.get("funame"):
        merged_header["funame"] = "（砚启）黄敬"
    if not merged_header.get("shuilv"):
        merged_header["shuilv"] = "0"
    if not merged_header.get("jiner"):
        merged_header["jiner"] = "0"

    # 2.5 校验表头下拉字段值是否在懂火选项里（AJAX 全量 ∪ 静态选项 = 完整）
    validation_errors = validate_header_options(merged_header, ajax_options, static_options)
    if validation_errors:
        msg = "表头校验失败: " + "; ".join(validation_errors)
        print(f"    ❌ {msg}")
        # 打印各字段的选项数量，方便排查
        for sel_id, label in [("fyname", "费用名称"), ("jsdanwei", "服务商名称"),
                               ("shuilv", "税率"), ("fcompany", "所属公司")]:
            n_ajax = len(ajax_options.get(sel_id) or set())
            n_static = len(static_options.get(sel_id) or set())
            print(f"      {label} 选项数: AJAX={n_ajax} 静态={n_static}")
        finished_updates = [(row["_record_id"], "失败", msg) for row in finished_rows]
        return finished_updates, (header_record_id, "失败", msg) if header_record_id else None
    print(f"    ✅ 表头校验通过")

    # 3. 构造成品行（跳过已存在的 + 父捆包匹配不上的）
    zid_tokens, kunbaohao_list, pinmin_list, guige_list = [], [], [], []
    chandi_list, caizhi_list, x_jianshu_list, x_shulian_list = [], [], [], []
    kuweihao_list, c_danjia_list = [], []
    ok_record_ids = []
    skipped_record_ids = []  # 已存在 或 父号不匹配

    for row in finished_rows:
        rid = row["_record_id"]
        kb = str(row.get("捆包号") or "").strip()
        if not kb:
            continue
        # 父捆包号
        parent = kb.rsplit("-", 1)[0] if "-" in kb else kb
        if parent not in materials:
            skipped_record_ids.append((rid, "失败",
                f"父捆包 {parent} 不在懂火原料中（需先在懂火添加原料）"))
            continue
        if kb in existing:
            skipped_record_ids.append((rid, "已写入",
                f"懂火中已存在此成品捆包，跳过"))
            continue

        token = materials[parent]["token"]
        zid_tokens.append(token)
        kunbaohao_list.append(kb)
        pinmin_list.append(str(row.get("品名") or "成品").strip())
        guige_list.append(str(row.get("规格") or "").strip())
        chandi_list.append(str(row.get("产地") or "").strip())
        caizhi_list.append(str(row.get("材质") or "").strip())
        x_jianshu_list.append(str(int(float(row.get("件(张)数") or 1))).strip())
        x_shulian_list.append(f"{float(row.get('重量(吨)') or 0):.3f}".rstrip('0').rstrip('.') if '.'
                              in f"{float(row.get('重量(吨)') or 0):.3f}" else f"{float(row.get('重量(吨)') or 0):.3f}")
        # 上面格式化简单化，直接保留3位
        x_shulian_list[-1] = f"{float(row.get('重量(吨)') or 0):.3f}"
        kuweihao_list.append(str(row.get("库位号") or "").strip())
        danjia = str(row.get("入库单价") or "").strip()
        c_danjia_list.append(danjia if danjia else "自动")
        ok_record_ids.append(rid)

    # 4. 如果没有新增行，直接返回跳过结果
    if not zid_tokens:
        print(f"    ⏭  没有需要新增的成品行")
        # 表头校验已通过，但无成品可写。若 skipped_record_ids 里有失败则表头标失败，否则标已写入
        has_fail = any(s == "失败" for _, s, _ in skipped_record_ids)
        header_status = "失败" if has_fail else "已写入"
        header_msg = "无新增成品" + ("（部分行失败见明细表）" if has_fail else "（所有明细已存在或无成品）")
        return skipped_record_ids, (header_record_id, header_status, header_msg) if header_record_id else None

    # 5. 构造 updatejg 请求
    update_data = {
        "jg_danhao": danhao,
        "jg_beizhu": js_escape(merged_header["jg_beizhu"]),
        "fyname": merged_header["fyname"],
        "jsdanwei": merged_header["jsdanwei"],
        "shuilv": merged_header["shuilv"],
        "jiner": merged_header["jiner"],
        "beizhu": merged_header["beizhu"],
        "funame": merged_header["funame"],
        "fcompany": merged_header["fcompany"],
        "zid[]": zid_tokens,
        "kunbaohao[]": kunbaohao_list,
        "pinmin[]": pinmin_list,
        "guige[]": guige_list,
        "chandi[]": chandi_list,
        "caizhi[]": caizhi_list,
        "x_jianshu[]": x_jianshu_list,
        "x_shulian[]": x_shulian_list,
        "kuweihao[]": kuweihao_list,
        "c_danjia[]": c_danjia_list,
    }

    print(f"    新表头: fyname='{update_data['fyname']}' jsdanwei='{update_data['jsdanwei']}' "
          f"shuilv={update_data['shuilv']} jiner={update_data['jiner']} fcompany='{update_data['fcompany']}' funame='{update_data['funame']}'")
    print(f"    新增 {len(zid_tokens)} 条成品:")
    for i, kb in enumerate(kunbaohao_list):
        print(f"      [{i+1}] {kb:<20} {pinmin_list[i]:<6} {guige_list[i]:<22} "
              f"件数={x_jianshu_list[i]} 重量={x_shulian_list[i]}")

    if dry_run:
        result_text = f"[DRY-RUN] 待写入 {len(zid_tokens)} 条"
        print(f"    [DRY-RUN] 将调用 updatejg 保存")
        result_updates = [(rid, "待处理", result_text) for rid in ok_record_ids]
        header_result = (header_record_id, "待处理",
                         f"[DRY-RUN] 校验通过，待写入 {len(zid_tokens)} 条成品") if header_record_id else None
        return result_updates + skipped_record_ids, header_result

    # 6. 实际提交
    headers = {"Referer": f"{DETAIL_PAGE}?jg_danhao={danhao}", **AJAX}
    r2 = session.post(UPDATEJG_API, data=update_data, timeout=30, headers=headers)
    text = r2.text.strip()
    ok = False
    try:
        resp = json.loads(text)
        if str(resp.get("code")) == "200":
            ok = True
    except Exception:
        pass

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if ok:
        print(f"    ✅ updatejg 保存成功")
        # 校验页面刷新
        time.sleep(0.5)
        r3 = session.get(DETAIL_PAGE, params={"jg_danhao": danhao}, timeout=15)
        after = parse_existing_finished(r3.text)
        # 检查按钮状态
        over_disabled = 'myButton-disabled' in r3.text.split("overjiagon")[0][-300:]
        print(f"    刷新后成品: {len(after)} 条, overjiagon 按钮: {'禁用' if over_disabled else '可用 ✓ (待人工完成加工)'}")

        result_updates = [(rid, "已写入", f"写入成功 {now}; overjiagon按钮:{'可用' if not over_disabled else '禁用(已完成)'}")
                          for rid in ok_record_ids]
        header_result = (header_record_id, "已写入",
                         f"写入成功 {now}; 成品{len(ok_record_ids)}条; overjiagon:{'可用' if not over_disabled else '禁用'}") if header_record_id else None
        return result_updates + skipped_record_ids, header_result
    else:
        msg = f"写入失败 {now}; {text[:200]}"
        print(f"    ❌ updatejg 失败: {text[:300]}")
        result_updates = [(rid, "失败", msg) for rid in ok_record_ids]
        header_result = (header_record_id, "失败", msg) if header_record_id else None
        return result_updates + skipped_record_ids, header_result


def main():
    app_token = env("BITABLE_APP_TOKEN", "CpYZbPbi3a0qo0smUhDcrKFgnGc")
    dry_run = env("DRY_RUN", "0") != "0" or "--dry-run" in sys.argv
    limit = int(env("IMPORT_LIMIT", "0") or "0")
    print(f"====== 飞书成品明细 → 懂火{' [DRY-RUN]' if dry_run else ' [实际执行]'} ======")

    t0 = time.time()
    tok = feishu_token()

    # 1. 自动找两张表
    header_table_id = bitable_auto_find_table(tok, app_token, "加工单表头",
                                               env("BITABLE_HEADER_TABLE_ID"))
    finished_table_id = bitable_auto_find_table(tok, app_token, "加工成品录入",
                                                 env("BITABLE_FINISHED_TABLE_ID"))
    if not header_table_id:
        print("❌ 找不到「加工单表头」表，请先在飞书中创建（表名必须含 加工单表头）")
        return 2
    if not finished_table_id:
        print("❌ 找不到「加工成品录入」表，请先在飞书中创建（表名必须含 加工成品录入）")
        return 2
    print(f"  表头表: {header_table_id}")
    print(f"  成品录入表: {finished_table_id}")

    # 2. 读飞书数据
    header_rows = bitable_read_all(tok, app_token, header_table_id)
    finished_rows = bitable_read_all(tok, app_token, finished_table_id)
    print(f"  表头记录: {len(header_rows)} 行")
    print(f"  成品明细记录: {len(finished_rows)} 行")

    # 只处理「处理状态!=已写入」的行（或者没填处理状态的）
    pending_finished = [r for r in finished_rows
                        if str(r.get("处理状态") or "").strip() != "已写入"]
    print(f"  待处理成品: {len(pending_finished)} 行（已过滤掉 处理状态=已写入 的行）")

    # 3. 按加工单号分组（同时保留 header_record_id 用于回写表头表状态）
    header_by_danhao = {}
    for r in header_rows:
        dh = str(r.get("加工单号") or "").strip()
        if dh:
            header_by_danhao[dh] = r

    finished_by_danhao = defaultdict(list)
    for r in pending_finished:
        dh = str(r.get("加工单号") or "").strip()
        if dh:
            finished_by_danhao[dh].append(r)

    print(f"  共 {len(finished_by_danhao)} 个加工单号需要处理")
    for dh in sorted(finished_by_danhao.keys()):
        has_header = "✓" if dh in header_by_danhao else "✗(无表头信息，将用懂火已有或默认)"
        print(f"    {dh}: {len(finished_by_danhao[dh])} 条成品 {has_header}")

    if limit > 0:
        keys = sorted(finished_by_danhao.keys())[:limit]
        finished_by_danhao = {k: finished_by_danhao[k] for k in keys}
        print(f"  ⚠️  IMPORT_LIMIT={limit}，只处理前 {limit} 个加工单: {keys}")

    # 4. 登录懂火
    print("\n[步骤] 登录懂火 ...")
    session = login_donghuo(username=env("DH_USERNAME", ""),
                            password=env("DH_PASSWORD", ""))
    if session is None:
        print("登录失败")
        return 1

    # 4.1 预取全量下拉选项（AJAX 动态加载的，只拉一次）
    print("[步骤] 获取全量下拉选项（AJAX接口）...")
    ajax_options = fetch_full_options(session)
    for sel_id, label in [("fyname", "费用名称"), ("jsdanwei", "服务商名称"),
                           ("shuilv", "税率"), ("fcompany", "所属公司"), ("funame", "添加人")]:
        n = len(ajax_options.get(sel_id) or set())
        print(f"  {label}: {n} 个")

    # 5. 逐个加工单处理
    stats = {"ok": 0, "skip": 0, "fail": 0}
    finished_updates_all = []  # 成品表回写列表 [(rid, status, result)]
    header_updates_all = []    # 表头表回写列表 [(rid, status, result)]

    sorted_dhs = sorted(finished_by_danhao.keys())
    for idx, danhao in enumerate(sorted_dhs, 1):
        print(f"\n--- [{idx}/{len(sorted_dhs)}] ---")
        header_row = header_by_danhao.get(danhao, {})
        header_info = {k: v for k, v in header_row.items() if k != "_record_id"}
        header_record_id = header_row.get("_record_id", "")
        try:
            finished_updates, header_update = process_one_danhao(
                session, tok, app_token,
                finished_table_id, header_table_id,
                danhao, header_info, header_record_id,
                finished_by_danhao[danhao],
                ajax_options,
                dry_run
            )
            for (rid, status, result) in finished_updates:
                if status == "已写入":
                    stats["ok"] += 1
                elif status == "待处理":
                    stats["skip"] += 1
                else:
                    stats["fail"] += 1
                finished_updates_all.append((rid, status, result))
            if header_update:
                header_updates_all.append(header_update)
        except Exception as e:
            msg = f"异常: {e}"
            print(f"    ❌ {msg}")
            for r in finished_by_danhao[danhao]:
                finished_updates_all.append((r["_record_id"], "失败", msg))
                stats["fail"] += 1
            if header_record_id:
                header_updates_all.append((header_record_id, "失败", msg))
        time.sleep(0.5)

    session.close()

    # 6. 回写飞书状态
    print(f"\n[步骤] 回写成品表处理状态 ({len(finished_updates_all)} 条) ...")
    if dry_run:
        print(f"  [DRY-RUN] 跳过回写")
    else:
        now_ts_ms = int(datetime.datetime.now().timestamp() * 1000)
        ok_cnt, fail_cnt = 0, 0
        for rid, status, result in finished_updates_all:
            if not rid:
                continue
            fields = {
                "处理状态": status,
                "写入结果": result[:500],
            }
            if status in ("已写入",):
                fields["写入时间"] = now_ts_ms
            resp = bitable_update(tok, app_token, finished_table_id, rid, fields)
            if resp.get("code") == 0:
                ok_cnt += 1
            else:
                fail_cnt += 1
                print(f"  ⚠️  成品表回写 {rid} 失败: {resp.get('msg','')}")
        print(f"  成品表回写成功 {ok_cnt} 条，失败 {fail_cnt} 条")

    # 6.1 回写表头表状态
    print(f"\n[步骤] 回写表头表处理状态 ({len(header_updates_all)} 条) ...")
    if dry_run:
        print(f"  [DRY-RUN] 跳过回写")
    elif header_updates_all:
        now_ts_ms = int(datetime.datetime.now().timestamp() * 1000) if not dry_run else 0
        ok_cnt2, fail_cnt2 = 0, 0
        for rid, status, result in header_updates_all:
            if not rid:
                continue
            fields = {
                "处理状态": status,
                "写入结果": result[:500],
            }
            if status == "已写入":
                fields["写入时间"] = now_ts_ms
            resp = bitable_update(tok, app_token, header_table_id, rid, fields)
            if resp.get("code") == 0:
                ok_cnt2 += 1
                print(f"  ✅ 表头 {rid}: {status}")
            else:
                fail_cnt2 += 1
                print(f"  ⚠️  表头表回写 {rid} 失败: {resp.get('msg','')}")
        print(f"  表头表回写成功 {ok_cnt2} 条，失败 {fail_cnt2} 条")

    # 7. 汇总
    elapsed = time.time() - t0
    print(f"\n{'='*60}")
    print(f"📊 汇总（耗时 {elapsed:.1f}s）{' DRY-RUN' if dry_run else ''}")
    print(f"  成品行: 成功 {stats['ok']} 条, 跳过/DRY-RUN待处理 {stats['skip']} 条, 失败 {stats['fail']} 条")
    if stats["fail"]:
        print(f"\n⚠️  失败明细:")
        for rid, status, result in finished_updates_all:
            if status == "失败":
                print(f"  - 成品行 record_id={rid}: {result[:200]}")
    if header_updates_all:
        print(f"\n📋 表头状态:")
        for rid, status, result in header_updates_all:
            print(f"  - 表头 record_id={rid}: {status} - {result[:200]}")
    print(f"{'='*60}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
