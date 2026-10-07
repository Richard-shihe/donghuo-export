#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
飞书多维表格「费用录入」→ 懂火 ERP「费用管理」批量写入

三个方向、三种提交方式（均来自 ERP 页面源码实测）：

  费用类型  飞书要填                程序解析                          提交接口
  采购      关联单号(采购单号 C…)    校验订单存在                       c_feiyon/save
  销售      关联单号(销售单号 X…)    查出库行(latest) → 按捆包号/牌号    c_feiyon/save
            可选 捆包号 / 牌号        收窄 → xids                         + leixin='销售'
  库存      捆包号                 查库存行 → 按捆包号匹配(必须唯一)     c_feiyon/save_kcfy

金额规则（用户 2026-10-04 确认）：
  优先校验 单价 × 计价重量 是否等于 金额；不一致时以 金额 ÷ 计价重量 覆盖单价。
  计价重量为空或 0 时无法反算 → 判失败回写原因（不猜）。

只新增、不改动：程序只处理「写入状态」为空/待写入的行（--retry-failed 可带上"失败"行）；
已写入的行永不重跑。写入前先标「写入中」，防崩溃后重复建单。

用法（从仓库根目录运行）：
  python 费用录入/import_feiyong_to_donghuo.py --dry-run       # 只解析+打印计划
  python 费用录入/import_feiyong_to_donghuo.py                 # 实际写入
  python 费用录入/import_feiyong_to_donghuo.py --limit 5 -v
  python 费用录入/import_feiyong_to_donghuo.py --retry-failed  # 连"失败"行一起重试

凭据（仓库根目录 .env）：
  FEISHU_APP_ID / FEISHU_APP_SECRET    飞书自建应用（需能编辑目标多维表格）
                                       —— 未配置时回退 FEISHU_NOTIFY_APP_ID / _SECRET
  DH_USERNAME / DH_PASSWORD            懂火账号
  FEIYONG_NOTIFY_UNION_IDS             通知收件人 union_id（逗号/空格分隔；留空=用脚本内置的默认两人）
  FEISHU_WEBHOOK_URL / FEISHU_WEBHOOK_SECRET   飞书群机器人（可选，额外再发一份）
"""
import os
import re
import sys
import json
import time
import base64
import hashlib
import hmac
import argparse
import datetime
import traceback
import pathlib

# ===== stdout/stderr 编码双保险（Windows subprocess 里 print emoji 会崩）=====
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from dotenv import load_dotenv

REPO_ROOT = pathlib.Path(__file__).resolve().parent.parent
load_dotenv(REPO_ROOT / ".env", override=False)
sys.path.insert(0, str(REPO_ROOT))

from donghuo_login import login_donghuo, BASE_URL  # noqa: E402

# ===== 配置 =====
BITABLE_APP_TOKEN = "WJ2IbCPUgax1Phsogk0cC0eNnCd"
TABLE_NAME = "费用录入"
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
AJAX = {"X-Requested-With": "XMLHttpRequest"}
TZ = datetime.timezone(datetime.timedelta(hours=8))

EP_CAIGOU = "/model/admin/caigou/m_dindan/getlist"      # 采购订单（筛选参数 sc_danhao）
EP_FAHUO = "/model/admin/xiaoshou/m_dindan/fahuolist"   # 销售出库行（筛选参数 sx_danhao）
EP_KUCUN = "/model/admin/xiaoshou/m_kucun/gl_kucun"     # 库存行（筛选参数 sxzhuantai）
EP_FUWU = "/model/admin/m_load/list_fuwu"               # 服务商主数据
EP_SAVE = "/controller/admin/caiwu/c_feiyon/save"       # 采购 / 销售
EP_SAVE_KCFY = "/controller/admin/caiwu/c_feiyon/save_kcfy"   # 库存
EP_FEE_LIST = "/model/admin/caiwu/m_feiyon/getlist"     # 回查记录号用

FEE_FIELDS = ["结算日期", "费用名称", "服务商", "税率", "计价方式", "计价重量", "单价", "金额", "备注说明"]
STATUS_FIELD = "写入状态"
STATUS_PENDING, STATUS_DOING = "待写入", "写入中"
STATUS_DONE, STATUS_FAIL = "已写入", "失败"

PAGE_SIZE = 300          # 懂火 getlist 单页上限
KUCUN_MAX_PAGES = 30     # 库存分页保险丝
SUBMIT_SLEEP = 1.0       # 每条之间的间隔（稳定性优先）
SUBMIT_RETRY = 3         # 提交重试次数
READ_RETRY = 3           # 读取（解析行 id）重试次数
DEFAULT_LIMIT = 200      # 单次运行处理的飞书行数上限

VERBOSE = False


def log(m):
    print(m, flush=True)


def vlog(m):
    if VERBOSE:
        print(m, flush=True)


# ==================== 飞书基础 ====================

def fs_session() -> requests.Session:
    """⚠️ POST（batch_update）不挂自动重试：写飞书状态必须明确成败"""
    s = requests.Session()
    retry = Retry(total=3, connect=2, read=2, status=2, backoff_factor=0.5,
                  status_forcelist=[429, 500, 502, 503, 504],
                  allowed_methods=frozenset(["GET"]))
    a = HTTPAdapter(max_retries=retry, pool_connections=4, pool_maxsize=8)
    s.mount("http://", a)
    s.mount("https://", a)
    return s


FS = fs_session()
_TRANSIENT = {1254607, 1254291}   # Data not ready / Write conflict


def fs_get(url, headers, timeout=60, tries=4):
    delay = 2.0
    for i in range(1, tries + 1):
        d = FS.get(url, headers=headers, timeout=timeout).json()
        if d.get("code") == 0 or d.get("code") not in _TRANSIENT:
            return d
        if i == tries:
            return d
        log(f"[飞书] [WARN] 读取偶发错误 {d.get('code')}，{delay:.0f}s 后重试")
        time.sleep(delay)
        delay *= 2
    return {}


def feishu_token() -> str:
    """应用身份 token。FEISHU_APP_ID 优先，回退 FEISHU_NOTIFY_APP_ID（与 数据汇总 一致）"""
    aid = os.environ.get("FEISHU_APP_ID", "").strip()
    sec = os.environ.get("FEISHU_APP_SECRET", "").strip()
    if not aid:
        aid = os.environ.get("FEISHU_NOTIFY_APP_ID", "").strip()
    if not sec:
        sec = os.environ.get("FEISHU_NOTIFY_APP_SECRET", "").strip()
    if not aid or not sec:
        raise SystemExit("[FAIL] 缺少飞书应用凭据（FEISHU_APP_ID/SECRET 或 "
                         "FEISHU_NOTIFY_APP_ID/SECRET，仓库根 .env）")
    d = FS.post(f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
                json={"app_id": aid, "app_secret": sec}, timeout=20).json()
    if d.get("code") != 0:
        raise SystemExit(f"[FAIL] 取 tenant_access_token 失败: {d.get('code')} {d.get('msg')}")
    return d["tenant_access_token"]


def table_id_by_name(token: str, name: str) -> str:
    d = fs_get(f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables",
               {"Authorization": f"Bearer {token}"}, timeout=30)
    if d.get("code") != 0:
        raise SystemExit(f"[FAIL] 读表列表失败: {d}")
    for t in d["data"].get("items", []):
        if t["name"] == name:
            return t["table_id"]
    raise SystemExit(f"[FAIL] base 里没有「{name}」表")


def read_rows(token: str, table_id: str) -> list:
    """读全表记录，返回 [{record_id, fields}]"""
    h = {"Authorization": f"Bearer {token}"}
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{table_id}/records"
    out, seen, pt = [], set(), None
    while True:
        qs = f"page_size=500" + (f"&page_token={pt}" if pt else "")
        d = fs_get(f"{url}?{qs}", h)
        if d.get("code") != 0:
            raise SystemExit(f"[FAIL] 读记录失败: {d}")
        dd = d.get("data") or {}
        items = dd.get("items") or []
        if not items:
            break
        dup = 0
        for it in items:
            rid = it.get("record_id")
            if rid and rid not in seen:
                out.append({"record_id": rid, "fields": it.get("fields") or {}})
                seen.add(rid)
            else:
                dup += 1
        if dup == len(items):
            break
        if not dd.get("has_more") or not dd.get("page_token"):
            break
        pt = dd.get("page_token")
    return out


def update_rows(token: str, table_id: str, updates: list) -> None:
    """批量回写 [{record_id, fields}]"""
    if not updates:
        return
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{table_id}"
           f"/records/batch_update?ignore_consistency_check=true")
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    for i in range(0, len(updates), 500):
        batch = updates[i:i + 500]
        d = FS.post(url, headers=h, json={"records": batch}, timeout=60).json()
        if d.get("code") != 0:
            raise RuntimeError(f"batch_update 失败: {d.get('code')} {d.get('msg')}")
    vlog(f"[飞书] 已回写 {len(updates)} 行")


def flat(v) -> str:
    """飞书单元格值 → 纯文本（文本字段可能是富文本片段数组）"""
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, bool):
        return "是" if v else "否"
    if isinstance(v, (int, float)):
        f = float(v)
        return str(int(f)) if f == int(f) else f"{f:g}"
    if isinstance(v, list):
        parts = []
        for seg in v:
            if isinstance(seg, dict):
                parts.append(str(seg.get("text") or seg.get("name") or "").strip())
            elif isinstance(seg, str):
                parts.append(seg.strip())
        return "".join(p for p in parts if p).strip()
    if isinstance(v, dict):
        return str(v.get("text") or v.get("name") or "").strip()
    return str(v).strip()


def flat_num(v):
    s = flat(v).replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def ms_to_date(v) -> str:
    n = flat_num(v)
    if not n:
        return ""
    return datetime.datetime.fromtimestamp(n / 1000, TZ).strftime("%Y-%m-%d")


# ==================== 懂火基础 ====================

def dh_post(session, path, data, retry=READ_RETRY, timeout=60):
    """懂火 POST，遇 SQL 断连 / 非 JSON 自动重试，返回解析后的 dict（失败返回 None）"""
    url = BASE_URL + path
    last = ""
    for i in range(1, retry + 1):
        try:
            r = session.post(url, data=data, headers=AJAX, timeout=timeout)
        except Exception as exc:
            last = f"请求异常 {exc}"
        else:
            if r.status_code != 200:
                last = f"HTTP {r.status_code}"
            elif any(k in r.text[:2000] for k in
                     ("Fatal error", "SQL Server", "远程主机强迫关闭了一个现有的连接", "com_exception")):
                last = "懂火后端 SQL 断连错误页"
            else:
                try:
                    return r.json()
                except json.JSONDecodeError:
                    last = f"非 JSON 返回: {r.text[:120]}"
        if i < retry:
            log(f"    [重试 {i}/{retry - 1}] {path} —— {last}")
            time.sleep(2 * i)
    log(f"    [FAIL] {path} 重试耗尽：{last}")
    return None


def load_fuwu(session) -> set:
    """服务商主数据（用于提交前校验，挡住错别字）"""
    d = dh_post(session, EP_FUWU, {})
    if not isinstance(d, list):
        return set()
    return {str(it.get("value") or it.get("key")).strip() for it in d if isinstance(it, dict)}


def find_caigou(session, danhao: str):
    d = dh_post(session, EP_CAIGOU, {"sc_danhao": danhao, "page": 1, "limit": 50})
    if not d:
        return None, "采购订单查询失败"
    rows = [r for r in (d.get("root") or []) if flat(r.get("订单号")) == danhao]
    if not rows:
        return None, f"懂火里查不到采购单 {danhao}"
    if len(rows) > 1:
        return None, f"采购单 {danhao} 匹配到 {len(rows)} 条，异常"
    return rows[0], ""


def find_fahuo_rows(session, danhao: str):
    d = dh_post(session, EP_FAHUO, {"sx_danhao": danhao, "page": 1, "limit": 300})
    if not d:
        return None, "出库明细查询失败"
    return d.get("root") or [], ""


def kucun_all(session):
    rows, page = [], 1
    while page <= KUCUN_MAX_PAGES:
        d = dh_post(session, EP_KUCUN, {"page": page, "limit": PAGE_SIZE, "sxzhuantai": ""})
        if not d:
            return None, "库存查询失败"
        batch = d.get("root") or []
        if not batch:
            break
        rows.extend(batch)
        pages = int(d.get("pgtotal") or 1)
        if page >= pages:
            break
        page += 1
        time.sleep(0.2)
    return rows, ""


# ==================== 任务组装 ====================

def build_task(rec: dict) -> dict:
    f = rec["fields"]
    t = {"record_id": rec["record_id"], "raw": f}
    for k in FEE_FIELDS:
        t[k] = flat(f.get(k))
    t["关联单号"] = flat(f.get("关联单号"))
    t["费用类型"] = flat(f.get("费用类型"))
    t["捆包号"] = flat(f.get("捆包号"))
    t["牌号"] = flat(f.get("牌号"))
    t["结算日期"] = ms_to_date(f.get("结算日期"))
    t["计价重量_n"] = flat_num(f.get("计价重量"))
    t["单价_n"] = flat_num(f.get("单价"))
    t["金额_n"] = flat_num(f.get("金额"))
    return t


def validate(t: dict) -> str:
    """返回错误原因，空串 = 通过"""
    if t["费用类型"] not in ("销售", "采购", "库存"):
        return f"费用类型「{t['费用类型']}」无效（应为 销售/采购/库存）"
    missing = [k for k in ("结算日期", "费用名称", "服务商", "税率", "计价方式")
               if not t.get(k)]
    if missing:
        return "缺少必填：" + "、".join(missing)
    if t["金额_n"] is None:
        return "金额为空或不是数字"
    if t["费用类型"] == "库存":
        if not t["捆包号"]:
            return "库存费用必须填捆包号"
    else:
        if not t["关联单号"]:
            return f"{t['费用类型']}费用必须填关联单号"
        prefix = "X" if t["费用类型"] == "销售" else "C"
        if not t["关联单号"].upper().startswith(prefix):
            log(f"    [WARN] 单号 {t['关联单号']} 不以 {prefix} 开头，仍按 {t['费用类型']} 处理")

    # 金额 ↔ 单价×重量
    w, amt, price = t["计价重量_n"], t["金额_n"], t["单价_n"]
    if w is None or w == 0:
        if t["计价方式"] == "重量":
            return "计价方式为「重量」但计价重量为空或 0，无法校验金额"
        t["单价_final"] = price if price is not None else amt
        t["单价_adjusted"] = False
    else:
        calc = amt / w
        if price is None or abs(price - calc) > 0.005:
            t["单价_final"] = round(calc, 4)
            t["单价_adjusted"] = True
            t["单价_old"] = price
        else:
            t["单价_final"] = price
            t["单价_adjusted"] = False
    t["单价_final"] = round(float(t["单价_final"]), 4)
    return ""


def fmt_num(v):
    if v is None:
        return ""
    f = float(v)
    return str(int(f)) if f == int(f) else f"{f:g}"


def submit_payload(t: dict, extra: dict) -> dict:
    return {
        "jstime": t["结算日期"],
        "fyname": t["费用名称"],
        "jsdanwei": t["服务商"],
        "shuilv": t["税率"],
        "jijia": t["计价方式"],
        "zhonlian": fmt_num(t["计价重量_n"]) or "",
        "danjia": fmt_num(t["单价_final"]),
        "jiner": fmt_num(t["金额_n"]),
        "beizhu": t["备注说明"],
        **extra,
    }


def lookup_fee_id(session, t: dict, claimed: set):
    """提交成功后回查费用列表取「懂火记录号」(ERP 内部 id) 和「懂火费用编号」(F 开头单号)。
    保存接口只回 {"code":"200"}，不返回 id，所以只能按
    科目名称 + 费用金额 + 日期 + 服务商名称 回查，并排除已被别的飞书行占用的 id。
    只有唯一命中才回填——宁可留空，也不写一个错的记录号。
    返回 (记录号, 费用编号, 说明)
    """
    params = {"page": 1, "limit": 300}
    if t["费用类型"] in ("采购", "销售"):
        params["danhao"] = t["关联单号"]
        params["leixin"] = t["费用类型"]
    d = dh_post(session, EP_FEE_LIST, params, retry=2)
    if not isinstance(d, dict):
        return "", "", "回查记录号失败"
    cand = []
    for r in (d.get("root") or []):
        rid = flat(r.get("id"))
        if not rid or rid in claimed:
            continue
        if flat(r.get("科目名称")) != t["费用名称"]:
            continue
        if flat(r.get("日期")) != t["结算日期"]:
            continue
        if flat(r.get("服务商名称")) != t["服务商"]:
            continue
        amt = flat_num(r.get("费用金额"))
        if amt is None or abs(amt - t["金额_n"]) > 0.01:
            continue
        cand.append((rid, flat(r.get("编号"))))
    if len(cand) == 1:
        return cand[0][0], cand[0][1], ""
    if not cand:
        return "", "", "提交成功但回查不到记录号（明细一致的费用可能被过滤）"
    return "", "", f"提交成功但匹配到 {len(cand)} 条同类费用，记录号不唯一，未回填"


def resolve_and_submit(session, t: dict, fuwu: set, dry_run: bool, claimed: set):
    """返回 (ok, 懂火记录号, 懂火费用编号, 说明, 明细行列表)"""
    detail = []
    if fuwu and t["服务商"] not in fuwu:
        return False, "", "", f"服务商「{t['服务商']}」不在懂火服务商库里", detail

    if t["费用类型"] == "采购":
        row, err = find_caigou(session, t["关联单号"])
        if err:
            return False, "", "", err, detail
        detail.append(f"采购单 {t['关联单号']}（id={flat(row.get('id'))}，"
                      f"{flat(row.get('所属公司'))}）")
        payload = submit_payload(t, {"danhao": t["关联单号"], "leixin": "采购"})
        path = EP_SAVE

    elif t["费用类型"] == "销售":
        rows, err = find_fahuo_rows(session, t["关联单号"])
        if err:
            return False, "", "", err, detail
        if not rows:
            return False, "", "", f"销售单 {t['关联单号']} 下没有出库行，无法挂费用", detail
        picked = rows
        if t["捆包号"]:
            picked = [r for r in rows if flat(r.get("捆包号")) == t["捆包号"]]
            if not picked:
                return (False, "", "",
                        f"销售单 {t['关联单号']} 下没有捆包号 {t['捆包号']} 的出库行", detail)
        elif t["牌号"]:
            picked = [r for r in rows if flat(r.get("材质")) == t["牌号"]]
            if not picked:
                return (False, "", "",
                        f"销售单 {t['关联单号']} 下没有牌号 {t['牌号']} 的出库行", detail)
        for r in picked:
            detail.append(f"出库行 id={flat(r.get('id'))} 材质={flat(r.get('材质'))} "
                          f"规格={flat(r.get('规格'))} 捆包号={flat(r.get('捆包号'))} "
                          f"重量={flat(r.get('重量(吨)'))}")
        payload = submit_payload(t, {"danhao": t["关联单号"],
                                     "xids": ",".join(flat(r.get("id")) for r in picked),
                                     "leixin": "销售"})
        path = EP_SAVE

    else:  # 库存
        rows, err = kucun_all(session)
        if err:
            return False, "", "", err, detail
        picked = [r for r in rows if flat(r.get("捆包号")) == t["捆包号"]]
        if not picked:
            return False, "", "", f"库存里找不到捆包号 {t['捆包号']}", detail
        if len(picked) > 1:
            return (False, "", "",
                    f"捆包号 {t['捆包号']} 在库存里匹配到 {len(picked)} 条，不唯一，拒绝写入",
                    detail)
        r = picked[0]
        detail.append(f"库存行 id={flat(r.get('id'))} {flat(r.get('品名'))} "
                      f"{flat(r.get('规格'))} {flat(r.get('材质'))} "
                      f"仓库={flat(r.get('仓库'))} 重量={flat(r.get('重量(吨)'))}")
        payload = submit_payload(t, {"zid": flat(r.get("id"))})
        path = EP_SAVE_KCFY

    if dry_run:
        return True, "(dry-run)", "(dry-run)", "dry-run 未提交", detail

    resp = None
    for i in range(1, SUBMIT_RETRY + 1):
        resp = dh_post(session, path, payload, retry=1)
        if isinstance(resp, dict) and str(resp.get("code")) == "200":
            break
        if i < SUBMIT_RETRY:
            log(f"    [重试 {i}/{SUBMIT_RETRY - 1}] 提交失败：{str(resp)[:120]}")
            time.sleep(2 * i)
    if not (isinstance(resp, dict) and str(resp.get("code")) == "200"):
        return False, "", "", f"懂火返回: {str(resp)[:200]}", detail
    vlog(f"    提交响应: {json.dumps(resp, ensure_ascii=False)[:200]}")
    rid, bianhao, note = lookup_fee_id(session, t, claimed)
    if rid:
        claimed.add(rid)
        log(f"    [回填] 懂火记录号={rid}  懂火费用编号={bianhao or '—'}")
    return True, rid, bianhao, (f"OK（{note}）" if note else "OK"), detail


# ==================== 通知卡片 ====================

# 通知收件人（union_id）：内置默认两人，可用 .env / 仓库 Secrets 的
# FEIYONG_NOTIFY_UNION_IDS 覆盖（逗号或空格分隔）。
# ⚠️ 2026-10-07 用户明确要求内置。本仓库是公开仓库，这两个 union_id 会随源码公开；
#    若要收回，把下面两个值删掉、只留环境变量读取即可。
DEFAULT_NOTIFY_UNION_IDS = [
    "on_b09bcbf3e74f5d423900aa9b2f00eb63",   # 洪
    "on_5b8dd7865ab1c9bc1ba8fea8668b068f",   # 陈红
]


def notify_recipients() -> list:
    raw = os.environ.get("FEIYONG_NOTIFY_UNION_IDS", "").strip()
    if raw:
        ids = [x for x in re.split(r"[,\s]+", raw) if x.strip()]
        if ids:
            return ids
    return list(DEFAULT_NOTIFY_UNION_IDS)


def mask_uid(uid: str) -> str:
    """日志与卡片里给 union_id 打码，避免真人身份外泄"""
    return f"{uid[:7]}***" if len(uid) > 10 else "***"


def build_card(title: str, lines: list, template: str = "blue") -> dict:
    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": template,
                   "title": {"tag": "plain_text", "content": title}},
        "elements": [{"tag": "div",
                      "text": {"tag": "lark_md", "content": "\n".join(lines)}}],
    }


def send_dm_card(union_id: str, card: dict, token: str):
    """应用身份私聊发卡片（仓库规则：通知一律走机器人/应用身份）"""
    url = f"{FEISHU_OPEN_BASE}/im/v1/messages?receive_id_type=union_id"
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"receive_id": union_id, "msg_type": "interactive",
            "content": json.dumps(card, ensure_ascii=False)}
    try:
        d = requests.post(url, headers=h, json=body, timeout=20).json()
    except Exception as exc:
        log(f"[通知] [WARN] 发给 {mask_uid(union_id)} 异常: {exc}")
        return
    if d.get("code") != 0:
        log(f"[通知] [WARN] 发给 {mask_uid(union_id)} 失败: {d.get('code')} {d.get('msg')}")
    else:
        vlog(f"[通知] 已私发 {mask_uid(union_id)}")


def send_card(title: str, lines: list, template: str = "blue", token: str = ""):
    """私聊卡片（主通道）+ 群机器人 webhook（可选，额外一份）"""
    card = build_card(title, lines, template)
    try:
        recips = notify_recipients()
        if not recips:
            vlog("[通知] 未配置 FEIYONG_NOTIFY_UNION_IDS，跳过私聊通知")
        tok = token or feishu_token()
        for uid in recips:
            send_dm_card(uid, card, tok)
    except Exception as exc:
        log(f"[通知] [WARN] 私聊卡片发送失败: {exc}")

    webhook = os.environ.get("FEISHU_WEBHOOK_URL", "").strip()
    if not webhook:
        return
    secret = os.environ.get("FEISHU_WEBHOOK_SECRET", "").strip()
    body = {"msg_type": "interactive", "card": card}
    if secret:
        ts = str(int(time.time()))
        sts = f"{ts}\n{secret}"
        body["timestamp"] = ts
        body["sign"] = base64.b64encode(
            hmac.new(sts.encode("utf-8"), digestmod=hashlib.sha256).digest()).decode()
    try:
        d = requests.post(webhook, json=body, timeout=20).json()
        if d.get("code") not in (0, None):
            log(f"[通知] [WARN] webhook 返回: {str(d)[:200]}")
    except Exception as exc:
        log(f"[通知] [WARN] webhook 发送失败: {exc}")


# ==================== 主流程 ====================

def parse_args():
    p = argparse.ArgumentParser(description="飞书费用录入 → 懂火费用管理")
    p.add_argument("--dry-run", action="store_true", help="只解析并打印计划，不提交、不回写")
    p.add_argument("--limit", type=int, default=DEFAULT_LIMIT, help=f"最多处理 N 行（默认 {DEFAULT_LIMIT}）")
    p.add_argument("--retry-failed", action="store_true", help="把「失败」的行也一起重试")
    p.add_argument("--no-notify", action="store_true", help="不发飞书通知")
    p.add_argument("--notify-test", action="store_true",
                   help="只给通知收件人发一条连通性测试卡片后退出")
    p.add_argument("-v", "--verbose", action="store_true")
    return p.parse_args()


def main() -> int:
    global VERBOSE
    args = parse_args()
    VERBOSE = args.verbose
    started = time.time()
    log(f"===== 飞书费用录入 → 懂火费用管理 [{datetime.datetime.now(TZ):%Y-%m-%d %H:%M:%S}] =====")
    log(f"模式: {'dry-run（不写任何东西）' if args.dry_run else '实际写入'}  "
        f"limit={args.limit}{'  retry-failed' if args.retry_failed else ''}")

    try:
        token = feishu_token()

        if args.notify_test:
            recips = notify_recipients()
            log(f"[通知] 收件人 {len(recips)} 人: {', '.join(recips)}")
            send_card("费用管理机器人连通性测试",
                      ["这是一条测试卡片，用于验证机器人能否给收件人发消息。",
                       f"收件人：{'、'.join(recips)}",
                       f"时间：{datetime.datetime.now(TZ):%Y-%m-%d %H:%M:%S}"],
                      "blue", token=token)
            log("[通知] 已发送（若收件人未收到，检查应用是否开通 im:message 权限、"
                "且收件人在应用可用范围内）")
            return 0

        table_id = table_id_by_name(token, TABLE_NAME)
        log(f"[飞书] [OK] 源表「{TABLE_NAME}」 → {table_id}")

        rows = read_rows(token, table_id)
        todo_status = {"", STATUS_PENDING} | ({STATUS_FAIL} if args.retry_failed else set())
        pending = [r for r in rows if flat(r["fields"].get(STATUS_FIELD)) in todo_status]
        log(f"[飞书] 共 {len(rows)} 行，其中待处理 {len(pending)} 行")
        if not pending:
            log("[完成] 没有待写入的行")
            return 0
        if len(pending) > args.limit:
            log(f"[飞书] 本次只处理前 {args.limit} 行（--limit 可调）")
            pending = pending[:args.limit]

        session = login_donghuo()
        if session is None:
            log("[FAIL] 懂火登录失败")
            return 1
        fuwu = load_fuwu(session)
        log(f"[懂火] [OK] 服务商主数据 {len(fuwu)} 项")

        claimed = {flat(r["fields"].get("懂火记录号"))
                   for r in rows if flat(r["fields"].get("懂火记录号"))}
        if claimed:
            log(f"[飞书] 已有 {len(claimed)} 个懂火记录号，回查时排除")
        results = []
        for i, rec in enumerate(pending, 1):
            t = build_task(rec)
            label = f"{t['费用类型'] or '?'}/{t['关联单号'] or t['捆包号'] or '?'}"
            log(f"\n[{i}/{len(pending)}] {label}  费用名称={t['费用名称']}  金额={flat(t['raw'].get('金额'))}")

            err = validate(t)
            if err:
                log(f"    [FAIL] {err}")
                results.append((t, False, "", "", err, []))
                continue
            if t["单价_adjusted"]:
                log(f"    [单价校正] {fmt_num(t.get('单价_old'))} → {fmt_num(t['单价_final'])} "
                    f"（金额 {fmt_num(t['金额_n'])} ÷ 重量 {fmt_num(t['计价重量_n'])}）")
            ok, rid, bianhao, msg, detail = resolve_and_submit(
                session, t, fuwu, args.dry_run, claimed)
            for d in detail:
                log(f"      影响: {d}")
            log(f"    {'[OK] ' + msg if ok else '[FAIL] ' + msg}")
            results.append((t, ok, rid, bianhao, msg, detail))
            if not args.dry_run:
                time.sleep(SUBMIT_SLEEP)

        # ---- 回写飞书 ----
        ok_n = sum(1 for _, ok, _, _, _, _ in results if ok)
        fail_n = len(results) - ok_n
        if not args.dry_run:
            now_ms = int(time.time() * 1000)
            updates = []
            for t, ok, rid, bianhao, msg, _ in results:
                fields = {STATUS_FIELD: STATUS_DONE if ok else STATUS_FAIL}
                if ok:
                    fields["写入时间"] = now_ms
                    if rid:
                        fields["懂火记录号"] = rid
                    if bianhao:
                        fields["懂火费用编号"] = bianhao
                    fields["写入结果"] = (msg or "写入成功")[:200]
                else:
                    fields["写入结果"] = msg[:200]
                updates.append({"record_id": t["record_id"], "fields": fields})
            update_rows(token, table_id, updates)
            log(f"\n[回写] [OK] 已更新 {len(updates)} 行状态")

        elapsed = time.time() - started
        log(f"\n===== 完成：成功 {ok_n} 条，失败 {fail_n} 条，用时 {elapsed:.1f}s =====")
        for t, ok, rid, bianhao, msg, _ in results:
            if not ok:
                log(f"  [FAIL] {t['费用类型']}/{t['关联单号'] or t['捆包号']} — {msg}")

        if not args.no_notify:
            head = "费用写入懂火"
            lines = [f"**成功 {ok_n} 条，失败 {fail_n} 条**",
                     f"用时 {elapsed:.0f}s{'（dry-run，未写入）' if args.dry_run else ''}"]
            for t, ok, rid, bianhao, msg, _ in results:
                tag = "✅" if ok else "❌"
                name = f"{t['费用类型']}/{t['关联单号'] or t['捆包号']}"
                extra = f" → {bianhao}" if bianhao and bianhao != "(dry-run)" else ""
                lines.append(f"{tag} {name} {t['费用名称']} {fmt_num(t['金额_n'])}{extra}")
                if not ok:
                    lines.append(f"　　{msg}")
            send_card(head + ("（dry-run 预览）" if args.dry_run else ""),
                      lines, "orange" if args.dry_run else ("red" if fail_n else "green"),
                      token=token)
        return 0 if fail_n == 0 else 2

    except SystemExit:
        raise
    except Exception:
        log("[FAIL] 运行异常：")
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
