#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火「客户管理」→ 飞书多维表格「数据汇总（2026）/客户 tblCE7zIWs804RR5」增量同步工作流。

流程：
  1. 复用封装的 donghuo_login.py（requests 版，ddddocr 自动过验证码）登录
  2. 调懂火系统自带数据接口 /model/admin/crm/m_kehu/getlist 分页拉取客户全量
     （筛选清空所有条件 = 不带任何过滤参数，rtotal=381 与页面「共 381 条」一致）
  3. 转 CSV（UTF-8-SIG）落本地备份
  4. 读取飞书表现有全部记录，以「客户名称」为主键做增量比对：
       a. 懂火有、飞书没有          → 新增（batch_create）
       b. 两边都有、跟踪字段有变化   → 更新（batch_update）；若曾被标「已删除」则一并取消标记
       c. 飞书有、懂火已没有        → 不删除！在「已删除」字段标记（batch_update），并在通知里列名
  5. 全部写入成功后删除 CSV 备份
  6. 发送飞书通知卡片给洪（新增/更新/标记删除条数、被标记客户名单、耗时）

【为什么不走页面「导出」按钮？】
  2026-09-06 实测：客户管理页（/view/admin/crm/v_kehu）工具栏「导出」按钮的处理函数
  khdown() 首行就是 `return false;`（部署方人为禁用），其本应 POST 的系统导出端点
  /view/admin/excelbiao/kehubaojia 服务端也直接返回「没有权限」（15 字节）。
  而表格自身加载数据用的 /model/admin/crm/m_kehu/getlist 正常可用，返回结构化 JSON，
  字段与多维表完全对应（与第四部分「往来」同一处理方式，非扒网页）。

使用：python sync_kehu_to_bitable.py [--dry-run] [--skip-upload] [--no-notify]
凭据：仓库根目录 .env 里的 DH_USERNAME / DH_PASSWORD + FEISHU_APP_ID / FEISHU_APP_SECRET
"""
import sys, os, json, time, datetime, argparse, math, requests
from pathlib import Path

# ===== stdout/stderr 编码双保险（Windows subprocess 里 print emoji 会崩）=====
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

from dotenv import load_dotenv
load_dotenv(Path(__file__).resolve().parent.parent / ".env", override=False)  # .env 统一放仓库根目录

# 让脚本可以直接 import 仓库根目录的 donghuo_login.py
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from donghuo_login import login_donghuo  # noqa: E402

# ===== 固定配置 =====
BITABLE_APP_TOKEN = "VahHb3YDBaBTwTsCjeAcaAhhnHc"   # 数据汇总（2026）
BITABLE_TABLE_ID  = "tblCE7zIWs804RR5"              # 客户管理
DONGHUO_KEHU_URL  = "https://erpa.donghuo.vip/model/admin/crm/m_kehu/getlist"
PAGE_SIZE = 300            # 接口单页上限 300（传 500 实测也只返回 300）
MAX_PAGES = 50             # 分页保险丝
CSV_DIR = Path(__file__).parent / "csv_backup"
CSV_DIR.mkdir(exist_ok=True)

FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
BATCH_SIZE = 500    # 飞书 bitable batch_create/batch_update 上限

# 同步完成后飞书通知（默认发给 洪 on_b09bcbf3e74f5d423900aa9b2f00eb63）
FEISHU_NOTIFY_UNION_ID = "on_b09bcbf3e74f5d423900aa9b2f00eb63"

# 「已删除」标记值（写进 已删除 文本字段）
DELETED_MARK = "已删除"

# ===== 字段映射：懂火接口字段 → 多维表字段 =====
# 主键：客户名称（text）。
# 跟踪字段（每次同步比对，有变化才更新）；全部为 text。
TRACKED_FIELDS = [
    "客户类型", "所属人", "联系人", "联系人职位", "固定电话", "移动电话",
    "邮箱地址", "所属省份", "联系地址", "主营产品", "采购产品", "备注",
]
# 仅新增时写入、更新时不动的字段
CREATE_ONLY_FIELDS = {
    "新增时间": ("创建时间", "datetime"),   # 懂火「新增时间」→ 多维表「创建时间」
}
# 不由本脚本写入的多维表字段：
#   重复(公式) / 对应(人工) / 参与(人员) / 创建(自动创建时间) / 简称(人工) /
#   分配(单选) / 来自-收付登记(引用) / 已删除(由本脚本增量逻辑管理)


# ============ 工具函数 ============

def log(msg: str):
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def env(name: str) -> str:
    return os.environ.get(name, "").strip()


def feishu_token() -> str:
    """获取飞书 tenant_access_token（每次现取，2 小时有效期）"""
    app_id = os.environ.get("FEISHU_APP_ID") or os.environ.get("FEISHU_NOTIFY_APP_ID", "").strip()
    app_secret = os.environ.get("FEISHU_APP_SECRET") or os.environ.get("FEISHU_NOTIFY_APP_SECRET", "").strip()
    if not app_id or not app_secret:
        raise RuntimeError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET（或 FEISHU_NOTIFY_APP_ID / NOTIFY_APP_SECRET）环境变量")
    r = requests.post(
        f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret},
        timeout=15,
    )
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换 tenant_access_token 失败: {data}")
    return data["tenant_access_token"]


def feishu_send_card(union_id: str, card: dict, token: str):
    """通过飞书机器人给指定用户发卡片消息"""
    url = f"{FEISHU_OPEN_BASE}/im/v1/messages?receive_id_type=union_id"
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "receive_id": union_id,
        "msg_type": "interactive",
        "content": json.dumps(card, ensure_ascii=False),
    }
    r = requests.post(url, headers=h, json=body, timeout=15)
    data = r.json()
    if data.get("code") != 0:
        log(f"[飞书通知] ❌ 发送失败: code={data.get('code')} msg={data.get('msg')}")
    else:
        log(f"[飞书通知] ✅ 卡片已发送给 {union_id}")


# ===== 通知卡片模板 =====
SOURCE_DESC = "懂火「客户管理」全量客户（筛选清空，共 381 家）"
TARGET_DESC = "飞书多维表「数据汇总（2026）/客户 tblCE7zIWs804RR5」"

# 当前执行环节（失败通知卡片里定位用）
CURRENT_STEP = "初始化"


def build_success_card(stats: dict, deleted_names: list, elapsed_s: float) -> dict:
    """同步成功通知卡片（绿色）——增量统计 + 被标记删除客户名单"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    fields = [
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**懂火客户总数**\n{stats['total_donghuo']} 家"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**飞书原有记录**\n{stats['total_feishu']} 条"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**🆕 新增客户**\n{stats['created']} 条"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**♻️ 更新资料**\n{stats['updated']} 条"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**⚠️ 标记已删除**\n{stats['marked_deleted']} 条（不删数据）"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**➖ 恢复标记**\n{stats['restored']} 条"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**无变化**\n{stats['unchanged']} 条"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**耗时**\n{elapsed_s:.1f} 秒"}},
    ]
    elements = [
        {"tag": "div", "fields": fields},
    ]
    if deleted_names:
        show = deleted_names[:30]
        lines = "\n".join(f"· {n}" for n in show)
        if len(deleted_names) > 30:
            lines += f"\n……等共 {len(deleted_names)} 家（多维表「已删除」字段已标记）"
        elements.append({"tag": "hr"})
        elements.append({"tag": "div", "text": {
            "tag": "lark_md",
            "content": f"**⚠️ 懂火中已不存在、已在表内标记「已删除」的客户：**\n{lines}"
        }})
    elements.append({"tag": "hr"})
    elements.append({"tag": "note", "elements": [
        {"tag": "plain_text", "content": f"{TARGET_DESC} ｜ 增量同步（主键：客户名称）｜ {now}"}
    ]})
    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "green", "title": {"tag": "plain_text", "content": "✅ 懂火客户管理增量同步完成"}},
        "elements": elements,
    }


def build_failure_card(step: str, error: str) -> dict:
    """同步失败通知卡片（红色）"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "red", "title": {"tag": "plain_text", "content": "❌ 懂火客户管理同步失败"}},
        "elements": [
            {"tag": "div", "fields": [
                {"is_short": True, "text": {"tag": "lark_md", "content": f"**失败环节**\n{step}"}},
                {"is_short": True, "text": {"tag": "lark_md", "content": f"**数据来源**\n{SOURCE_DESC}"}},
            ]},
            {"tag": "div", "text": {"tag": "lark_md", "content": f"**错误信息**\n{str(error)[:600]}"}},
            {"tag": "hr"},
            {"tag": "note", "elements": [{"tag": "plain_text", "content": f"同步时间 {now}"}]},
        ],
    }


# ============ 懂火 → 拉取客户全量 ============

def fetch_kehu_rows(max_pages: int = MAX_PAGES) -> list:
    """
    登录懂火 → 调 /model/admin/crm/m_kehu/getlist 分页拉全量客户。
    筛选清空所有条件 = 不带任何过滤参数（与页面点「重置」后查询等效）。
    """
    global CURRENT_STEP
    CURRENT_STEP = "登录懂火（requests 版 donghuo_login）"
    session = login_donghuo()
    if session is None:
        raise RuntimeError("懂火登录失败（donghuo_login 返回 None）")
    log("[懂火] ✅ 登录成功，session 已就绪")

    CURRENT_STEP = "拉取客户接口"
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://erpa.donghuo.vip/view/admin/crm/v_kehu",
    }
    all_rows = []
    rtotal = None
    for page_no in range(1, max_pages + 1):
        r = session.post(
            DONGHUO_KEHU_URL,
            data={"page": page_no, "limit": PAGE_SIZE},   # 不带任何筛选参数 = 全量
            headers=headers,
            timeout=30,
        )
        try:
            data = r.json()
        except ValueError:
            raise RuntimeError(f"getlist 第 {page_no} 页返回非 JSON（可能登录态失效）: {r.text[:200]}")
        root = data.get("root") or []
        if rtotal is None:
            rtotal = int(data.get("rtotal") or 0)
            log(f"[懂火] 客户总数 rtotal={rtotal}，每页 {PAGE_SIZE} 条，预计 {math.ceil(rtotal / PAGE_SIZE)} 页")
        all_rows.extend(root)
        log(f"[懂火] 第 {page_no} 页: +{len(root)} 条，累计 {len(all_rows)}/{rtotal}")
        if not root:
            break
        if rtotal and len(all_rows) >= rtotal:
            break
        time.sleep(0.3)

    if rtotal and len(all_rows) < rtotal:
        log(f"[懂火] ⚠️ 仅拉到 {len(all_rows)}/{rtotal} 条，未拉满（请检查）")

    # 重名检查（主键冲突预警）
    seen = {}
    for row in all_rows:
        name = norm_text(row.get("客户名称"))
        if name:
            seen.setdefault(name, 0)
            seen[name] += 1
    dups = {n: c for n, c in seen.items() if c > 1}
    if dups:
        log(f"[懂火] ⚠️ 客户名称重名 {len(dups)} 个: {list(dups.items())[:10]}")
    log(f"[懂火] ✅ 客户拉取完成，共 {len(all_rows)} 条")
    return all_rows


# ============ rows → CSV 备份 ============

def rows_to_csv(rows: list) -> Path:
    """原始客户 dict 列表 → UTF-8-SIG CSV 备份（上传成功后由调用方删除）"""
    import pandas as pd
    df = pd.DataFrame(rows)
    df = df.dropna(how="all")
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = CSV_DIR / f"kehu_export_{now}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    log(f"[备份] CSV 已保存: {csv_path}（{len(df)} 行 × {len(df.columns)} 列）")
    return csv_path


# ============ 值转换与比对工具 ============

def norm_text(v) -> str:
    """文本归一化：None/空白/全角空格 → ''，其余 strip。用于比对与主键。"""
    if v is None:
        return ""
    s = str(v).replace("　", " ").strip()
    return s


import re
_JUNK_RE = re.compile(r"^\d{1,2}$")   # 1~2 位纯数字视为占位垃圾值（懂火里联系人/电话常见 '1' '0' '00'）


def is_meaningful(v) -> bool:
    """值是否有意义：非空且不是 '1'/'0'/'00' 这类占位垃圾（用于重名合并与写入过滤）"""
    s = norm_text(v)
    if not s:
        return False
    if _JUNK_RE.match(s):
        return False
    return True


def to_datetime_ms(raw) -> int | None:
    """'2026-09-02 13:18:43' → 毫秒时间戳（飞书 DateTime 字段格式）"""
    s = norm_text(raw)
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return int(dt.timestamp() * 1000)
        except ValueError:
            continue
    return None


# ============ 飞书多维表：读取 + 增量写入 ============

def bitable_list_all_records(token: str) -> list:
    """
    读取多维表全部记录，返回 [{"record_id":..., "fields":{...}}, ...]。
    ⚠️ search 接口 page_token 会永远不推进，必须用 GET list 接口。
    """
    h = {"Authorization": f"Bearer {token}"}
    base_url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records"
    out = []
    seen = set()
    page_token = None
    while True:
        qs = f"page_size={BATCH_SIZE}"
        if page_token:
            qs += f"&page_token={page_token}"
        r = requests.get(f"{base_url}?{qs}", headers=h, timeout=30)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"list records 失败: {data}")
        d = data.get("data") or {}
        items = d.get("items") or []
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
            log(f"[飞书] ⚠️ page_token 未推进，停止；累计 {len(out)} 条")
            break
        if not d.get("has_more"):
            break
        page_token = d.get("page_token")
        if not page_token:
            break
    log(f"[飞书] 读取现有记录: 共 {len(out)} 条")
    return out


def bitable_batch_create(token: str, records: list) -> int:
    """批量新增，返回成功条数"""
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records/batch_create?ignore_consistency_check=true"
    total = len(records)
    created = 0
    for i in range(0, total, BATCH_SIZE):
        batch = records[i:i + BATCH_SIZE]
        r = requests.post(url, headers=h, json={"records": batch}, timeout=60)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_create 失败 批 {i // BATCH_SIZE + 1}: code={data.get('code')} msg={data.get('msg')} sample={str(data)[:400]}")
        created += len((data.get("data") or {}).get("records") or [])
        log(f"[飞书] create: {len(batch)} 条 (批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE})，累计 {created}")
    return created


def bitable_batch_update(token: str, updates: list) -> int:
    """
    批量更新。updates: [{"record_id":..., "fields":{...}}, ...]
    返回成功条数。
    """
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records/batch_update?ignore_consistency_check=true"
    total = len(updates)
    done = 0
    for i in range(0, total, BATCH_SIZE):
        batch = updates[i:i + BATCH_SIZE]
        r = requests.post(url, headers=h, json={"records": batch}, timeout=60)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_update 失败 批 {i // BATCH_SIZE + 1}: code={data.get('code')} msg={data.get('msg')} sample={str(data)[:400]}")
        done += len((data.get("data") or {}).get("records") or [])
        log(f"[飞书] update: {len(batch)} 条 (批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE})，累计 {done}")
    return done


def build_new_record_fields(row: dict) -> dict:
    """懂火客户 dict → 新增记录的 fields（含创建时间）。垃圾值（'1'/'0' 等）不写入。"""
    fields = {"客户名称": norm_text(row.get("客户名称"))}
    for col in TRACKED_FIELDS:
        v = norm_text(row.get(col))
        if is_meaningful(v):
            fields[col] = v
    # 仅新增时写「创建时间」
    ts = to_datetime_ms(row.get("新增时间"))
    if ts:
        fields["创建时间"] = ts
    return fields


def diff_tracked_fields(row: dict, feishu_fields: dict) -> dict:
    """
    比对懂火行与飞书现有记录的跟踪字段，返回需要更新的 fields（无变化返回 {}）。
    只在懂火值有意义且与飞书现值不同时更新（避免用空值/垃圾值覆盖人工补充）。
    """
    changed = {}
    for col in TRACKED_FIELDS:
        new_v = norm_text(row.get(col))
        old_v = norm_text(feishu_fields.get(col))
        if is_meaningful(new_v) and new_v != old_v:
            changed[col] = new_v
    return changed


def merge_duplicate_rows(rows: list) -> dict:
    """
    懂火里存在同名客户多行（老行真实资料 + 新行占位垃圾值的重复建档）。
    按客户名称分组合并：
      - 基础行 = 跟踪字段「有效值」最多的那行（垃圾值 '1'/'0' 不计）
      - 基础行为空的字段，用其他行的第一个有效值补齐
      - 新增时间取最早（原始建档时间）
    返回 {客户名称: 合并后 row dict}
    """
    from collections import defaultdict
    groups = defaultdict(list)
    for row in rows:
        name = norm_text(row.get("客户名称"))
        if name:
            groups[name].append(row)

    merged = {}
    for name, grp in groups.items():
        if len(grp) == 1:
            merged[name] = grp[0]
            continue

        def score(r):
            return sum(1 for c in TRACKED_FIELDS if is_meaningful(r.get(c)))

        base = max(grp, key=score)
        m = dict(base)
        for c in TRACKED_FIELDS:
            if not is_meaningful(m.get(c)):
                for r in grp:
                    if r is base:
                        continue
                    if is_meaningful(r.get(c)):
                        m[c] = r[c]
                        break
        # 新增时间取最早
        ts_list = [t for t in (to_datetime_ms(r.get("新增时间")) for r in grp) if t]
        if ts_list:
            earliest = datetime.datetime.fromtimestamp(min(ts_list) / 1000)
            m["新增时间"] = earliest.strftime("%Y-%m-%d %H:%M:%S")
        log(f"[懂火] 重名合并: {name}（{len(grp)} 行，基础行 id={base.get('id')}，有效值 {score(base)} 个）")
        merged[name] = m
    return merged


# ============ 增量同步核心 ============

def plan_incremental(donghuo_rows: list, feishu_records: list) -> dict:
    """
    以客户名称为主键生成增量计划：
      to_create:  懂火有、飞书没有 → 新增 fields 列表
      to_update:  两边都有但字段有变化 → [(record_id, fields)]
      to_restore: 飞书曾标「已删除」但懂火又有了 → 更新 fields 时清空 已删除
      to_mark_deleted: 飞书有、懂火没有且未标记 → [(record_id, 客户名称)]
      unchanged:  无变化
    """
    # 懂火侧：名称 → 合并后 row（重名多行按数据丰富度合并，避免垃圾占位行覆盖真实资料）
    dh_map = merge_duplicate_rows(donghuo_rows)

    # 飞书侧：名称 → [record,...]（重名全部更新/标记）
    fs_map = {}
    no_name = 0
    for rec in feishu_records:
        name = norm_text((rec["fields"] or {}).get("客户名称"))
        if not name:
            no_name += 1
            continue
        fs_map.setdefault(name, []).append(rec)
    if no_name:
        log(f"[飞书] ⚠️ {no_name} 条记录无客户名称，跳过比对（不做任何处理）")

    to_create = []
    to_update = []
    to_restore = []
    unchanged = 0

    for name, row in dh_map.items():
        matches = fs_map.get(name)
        if not matches:
            to_create.append(build_new_record_fields(row))
            continue
        for rec in matches:
            ffields = rec["fields"] or {}
            changed = diff_tracked_fields(row, ffields)
            was_deleted = bool(norm_text(ffields.get("AI 提示")))
            if was_deleted:
                # 客户在懂火重新出现 → 取消删除标记 + 同步最新资料
                fields = dict(changed)
                fields["AI 提示"] = ""   # 清空标记
                to_restore.append({"record_id": rec["record_id"], "fields": fields,
                                   "name": name})
            elif changed:
                to_update.append({"record_id": rec["record_id"], "fields": changed,
                                  "name": name})
            else:
                unchanged += 1

    to_mark_deleted = []
    for name, recs in fs_map.items():
        if name in dh_map:
            continue
        for rec in recs:
            ffields = rec["fields"] or {}
            if norm_text(ffields.get("AI 提示")):
                continue   # 已标记过，不重复写
            to_mark_deleted.append({"record_id": rec["record_id"], "name": name})

    return {
        "to_create": to_create,
        "to_update": to_update,
        "to_restore": to_restore,
        "to_mark_deleted": to_mark_deleted,
        "unchanged": unchanged,
        "total_donghuo": len(dh_map),
        "total_feishu": len(feishu_records),
    }


# ============ 主流程 ============

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="拉取+转CSV+打印增量计划，不写飞书")
    ap.add_argument("--skip-upload", action="store_true", help="跳过写入飞书（只拉取+备份）")
    ap.add_argument("--no-notify", action="store_true", help="不发送飞书通知")
    args = ap.parse_args()

    t0 = time.time()
    log("==== 懂火客户管理 → 飞书多维表 增量同步启动 ====")

    # ---- Step 1: 拉取客户全量 ----
    rows = fetch_kehu_rows()
    if not rows:
        raise RuntimeError("客户接口返回空，中止")

    # ---- Step 2: CSV 备份 ----
    CURRENT_STEP = "落 CSV 备份"
    csv_path = rows_to_csv(rows)

    if args.dry_run or args.skip_upload:
        # 即使 dry-run 也读飞书表，把增量计划打印出来供核对
        CURRENT_STEP = "读取飞书表现有记录"
        token = feishu_token()
        feishu_records = bitable_list_all_records(token)
        plan = plan_incremental(rows, feishu_records)
        log("==== 增量计划（DRY-RUN / SKIP-UPLOAD，不写入）====")
        log(f"  新增 {len(plan['to_create'])} 条，更新 {len(plan['to_update'])} 条，"
            f"恢复 {len(plan['to_restore'])} 条，标记删除 {len(plan['to_mark_deleted'])} 条，"
            f"无变化 {plan['unchanged']} 条")
        if plan["to_create"]:
            log(f"  新增样例: {[r['客户名称'] for r in plan['to_create'][:10]]}")
        if plan["to_mark_deleted"]:
            log(f"  将标记已删除: {[r['name'] for r in plan['to_mark_deleted'][:30]]}")
        if plan["to_update"]:
            log(f"  更新样例: {[(r['name'], list(r['fields'].keys())) for r in plan['to_update'][:5]]}")
        elapsed = time.time() - t0
        log(f"==== 完成（{'DRY-RUN' if args.dry_run else 'SKIP-UPLOAD'}），耗时 {elapsed:.1f}s，CSV 保留: {csv_path.name} ====")
        return 0

    # ---- Step 3: 读飞书 + 增量比对 ----
    CURRENT_STEP = "获取飞书凭证并读取现有记录"
    token = feishu_token()
    log("[飞书] ✅ tenant_access_token 已获取")
    feishu_records = bitable_list_all_records(token)

    CURRENT_STEP = "增量比对"
    plan = plan_incremental(rows, feishu_records)
    log(f"[比对] 新增 {len(plan['to_create'])}，更新 {len(plan['to_update'])}，"
        f"恢复 {len(plan['to_restore'])}，标记删除 {len(plan['to_mark_deleted'])}，"
        f"无变化 {plan['unchanged']}")

    # ---- Step 4: 写入（新增 → 更新 → 恢复 → 标记删除）----
    created = updated = restored = marked = 0
    CURRENT_STEP = "写入新增客户"
    if plan["to_create"]:
        recs = [{"fields": f} for f in plan["to_create"]]
        created = bitable_batch_create(token, recs)
    else:
        log("[飞书] 无新增客户")

    CURRENT_STEP = "更新变化客户"
    if plan["to_update"]:
        updated = bitable_batch_update(token, [{"record_id": r["record_id"], "fields": r["fields"]}
                                               for r in plan["to_update"]])
    else:
        log("[飞书] 无资料变化客户")

    CURRENT_STEP = "恢复误标客户"
    if plan["to_restore"]:
        restored = bitable_batch_update(token, [{"record_id": r["record_id"], "fields": r["fields"]}
                                                for r in plan["to_restore"]])
        log(f"[飞书] ♻️ {restored} 条客户在懂火重新出现，已取消「已删除」标记")
    else:
        log("[飞书] 无需恢复标记的客户")

    CURRENT_STEP = "标记已删除客户"
    if plan["to_mark_deleted"]:
        # AI 提示字段写「已删除」；只标记，不删除记录
        marked = bitable_batch_update(token, [
            {"record_id": r["record_id"], "fields": {"AI 提示": DELETED_MARK}}
            for r in plan["to_mark_deleted"]
        ])
        log(f"[飞书] ⚠️ {marked} 条客户在懂火中已不存在，已标记「已删除」（记录保留）")
    else:
        log("[飞书] 无需要标记删除的客户")

    stats = {
        "total_donghuo": plan["total_donghuo"],
        "total_feishu": plan["total_feishu"],
        "created": created,
        "updated": updated,
        "restored": restored,
        "marked_deleted": marked,
        "unchanged": plan["unchanged"],
    }
    deleted_names = [r["name"] for r in plan["to_mark_deleted"]]

    # ---- Step 4.5: 后处理 —— 补参与 + 标记请复检 ----
    CURRENT_STEP = "客户后处理：补参与 + 标记请复检"
    if not args.dry_run:
        fix_updates = []
        all_records = bitable_list_all_records(token)
        for rec in all_records:
            f = rec["fields"] or {}
            if norm_text(f.get("AI 提示")) == DELETED_MARK:
                continue  # 跳过已删除客户
            owner = f.get("参与")
            owner_empty = not owner or (isinstance(owner, list) and len(owner) == 0) or (isinstance(owner, dict) and not owner)
            if not owner_empty:
                continue
            lookup = f.get("来自-收付登记")
            if not isinstance(lookup, dict):
                continue
            users = lookup.get("users") or []
            if not users:
                continue
            first_id = users[0].get("id")
            if not first_id or not first_id.startswith("ou_"):
                continue
            if norm_text(f.get("AI 提示")) == "请复检":
                continue
            fix_updates.append({"record_id": rec["record_id"],
                                "fields": {"参与": [{"id": first_id}], "AI 提示": "请复检"}})
        if fix_updates:
            fixed = bitable_batch_update(token, fix_updates)
            log(f"[飞书] ⚙️ 后处理：补参与 + 标记请复检 {fixed} 条")
            stats["fixed_owner"] = fixed
        else:
            stats["fixed_owner"] = 0

    # ---- Step 5: 全部成功 → 删除 CSV 备份 ----
    CURRENT_STEP = "删除 CSV 备份"
    if csv_path.exists():
        csv_path.unlink()
        log(f"[清理] CSV 已删除: {csv_path.name}")

    elapsed = time.time() - t0
    log(f"==== 增量同步完成，耗时 {elapsed:.1f}s ====")

    # ---- Step 6: 飞书通知 ----
    if not args.no_notify:
        try:
            CURRENT_STEP = "发送飞书通知"
            token = feishu_token()
            card = build_success_card(stats, deleted_names, elapsed)
            feishu_send_card(FEISHU_NOTIFY_UNION_ID, card, token)
        except Exception as e:
            log(f"[飞书通知] ⚠️ 发送异常: {e}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        # 失败通知（--no-notify 时也发失败卡片？与既有脚本一致：失败照发，便于及时发现）
        try:
            tok = feishu_token()
            feishu_send_card(FEISHU_NOTIFY_UNION_ID, build_failure_card(CURRENT_STEP, str(e)), tok)
        except Exception:
            pass
        sys.exit(1)
