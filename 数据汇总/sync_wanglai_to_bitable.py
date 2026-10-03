#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火「往来」（收付款流水）→ 飞书多维表格「数据汇总（2026）/tblbS1dPaDVL3GY8」自动更新工作流。

流程：
  1. 复用封装的 donghuo_login.py（requests 版，ddddocr 自动过验证码）登录
  2. 调懂火系统自带的数据接口 /model/admin/caiwu/m_liushui/getlist 分页拉取
     收付款流水全量记录（确认状态=全部：不传 zhuantai 参数即不过滤状态，
     已确认 8876+ 条 + 待确认 19 条全部返回）
  3. 转 CSV（UTF-8-SIG）落本地备份
  4. 清空飞书多维表 tblbS1dPaDVL3GY8 现有全部记录
  5. 按原字段格式批量写入新数据
  6. 上传成功后删除 CSV 备份
  7. 发送飞书通知卡片给洪（更新条数、耗时等摘要）

【为什么走接口而不是页面导出按钮？】
  2026-09-06 用 R@shihe.donghuo 账号把 erpa / erpb 两个实例的全部 22 个菜单、
  所有容器页二级标签、首页组件、顶部菜单逐个排查过，导航里不存在「往来」页
  （疑似该账号未开放此菜单权限）。但收款/付款弹窗（v_x_addshk / v_c_addfk）
  保存数据的接口正是 /controller/admin/caiwu/c_liushui/save，流水的数据源就是
  /model/admin/caiwu/m_liushui/getlist（rtotal=8895 与往来表 8879 条吻合），
  字段与往来多维表完全对应。因此本脚本调用系统同一数据接口取数，
  等效于页面导出按钮的数据（结构化 JSON，比页面导出更完整）。

使用：python sync_wanglai_to_bitable.py [--dry-run] [--skip-clear] [--skip-upload] [--no-notify] ...
凭据：仓库根目录 .env 里的 DH_USERNAME / DH_PASSWORD + FEISHU_APP_ID / FEISHU_APP_SECRET
"""
import sys, os, json, time, datetime, argparse, traceback, math, requests
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
BITABLE_TABLE_ID  = "tblbS1dPaDVL3GY8"               # 往来（收付款流水）
DONGHUO_LIUSHUI_URL = "https://erpa.donghuo.vip/model/admin/caiwu/m_liushui/getlist"
PAGE_SIZE = 300            # 接口单页上限 300（传 500 实测也只返回 300）
MAX_PAGES = 200            # 分页保险丝
CSV_DIR = Path(__file__).parent / "csv_backup"
CSV_DIR.mkdir(exist_ok=True)

FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
BATCH_SIZE = 500    # 飞书 bitable batch_create/batch_delete 上限

# 同步完成后飞书通知（默认发给 洪 on_b09bcbf3e74f5d423900aa9b2f00eb63）
FEISHU_NOTIFY_UNION_ID = "on_b09bcbf3e74f5d423900aa9b2f00eb63"

# ===== 多维表字段类型映射（2026-09-06 API 探测 tblbS1dPaDVL3GY8）=====
# 接口字段名与多维表字段名一致；编号/参与/字段1~16 不由本脚本写入
BITABLE_FIELD_TYPES = {
    "所属公司":     "text",
    "日期":         "datetime",
    "我方帐户":     "text",
    "交易对方":     "text",
    "交易类型":     "text",
    "科目名称":     "text",
    "结算方式":     "text",
    "金额":         "number",
    "状态":         "text",
    "结算对方":     "text",
    "订单号":       "text",
    "销售人":       "text",
    "备注说明":     "text",
    "提交人":       "text",
    "确认人":       "text",
}


# ============ 工具函数 ============

def log(msg: str):
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


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


def env(name: str) -> str:
    return os.environ.get(name, "").strip()


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
SOURCE_DESC = "懂火「往来」收付款流水（确认状态=全部）"
TARGET_DESC = "飞书多维表「数据汇总（2026）/往来 tblbS1dPaDVL3GY8」"

# 当前执行环节（失败通知卡片里定位用）
CURRENT_STEP = "初始化"


def build_success_card(written: int, cleared, elapsed_s: float) -> dict:
    """同步成功通知卡片（绿色）"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    fields = [
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**更新记录**\n{written} 条"}},
        {"is_short": True, "text": {"tag": "lark_md",
            "content": f"**清空旧记录**\n{cleared} 条" if cleared is not None else "**清空旧记录**\n跳过"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**耗时**\n{elapsed_s:.1f} 秒"}},
        {"is_short": True, "text": {"tag": "lark_md", "content": f"**数据来源**\n{SOURCE_DESC}"}},
    ]
    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "green", "title": {"tag": "plain_text", "content": "✅ 懂火往来流水同步成功"}},
        "elements": [
            {"tag": "div", "fields": fields},
            {"tag": "hr"},
            {"tag": "note", "elements": [
                {"tag": "plain_text", "content": f"{TARGET_DESC} ｜ 同步时间 {now}"}
            ]},
        ],
    }


def build_failure_card(step: str, error: str) -> dict:
    """同步失败通知卡片（红色）"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "red", "title": {"tag": "plain_text", "content": "❌ 懂火往来流水同步失败"}},
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


# ============ 懂火 → 拉取往来流水 ============

def fetch_wanglai_rows(max_pages: int = MAX_PAGES) -> list:
    """
    登录懂火 → 调 /model/admin/caiwu/m_liushui/getlist 分页拉全量流水。
    确认状态=全部：不传 zhuantai 参数（传"已确认"/"待确认"才过滤）。
    返回原始 dict 列表（含 id/提交日期/米数 等全部字段，便于 CSV 备份）。
    """
    global CURRENT_STEP
    CURRENT_STEP = "登录懂火（requests 版 donghuo_login）"
    session = login_donghuo()
    if session is None:
        raise RuntimeError("懂火登录失败（donghuo_login 返回 None）")
    log("[懂火] ✅ 登录成功，session 已就绪")

    CURRENT_STEP = "拉取往来流水接口"
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://erpa.donghuo.vip/view/admin/caiwu/v_x_jiesuan",
    }
    all_rows = []
    rtotal = None
    for page_no in range(1, max_pages + 1):
        r = session.post(
            DONGHUO_LIUSHUI_URL,
            data={"page": page_no, "limit": PAGE_SIZE},   # 确认状态=全部 → 不传 zhuantai
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
            log(f"[懂火] 流水总数 rtotal={rtotal}，每页 {PAGE_SIZE} 条，预计 {math.ceil(rtotal / PAGE_SIZE)} 页")
        all_rows.extend(root)
        log(f"[懂火] 第 {page_no} 页: +{len(root)} 条，累计 {len(all_rows)}/{rtotal}")
        if not root:
            break
        if rtotal and len(all_rows) >= rtotal:
            break
        time.sleep(0.3)   # 对服务器友好一点

    if rtotal and len(all_rows) < rtotal:
        log(f"[懂火] ⚠️ 仅拉到 {len(all_rows)}/{rtotal} 条，未拉满（请检查）")
    log(f"[懂火] ✅ 流水拉取完成，共 {len(all_rows)} 条")
    return all_rows


# ============ rows → CSV 备份 ============

def rows_to_csv(rows: list) -> Path:
    """原始流水 dict 列表 → UTF-8-SIG CSV 备份（上传成功后由调用方删除）"""
    import pandas as pd
    df = pd.DataFrame(rows)
    df = df.dropna(how="all")
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = CSV_DIR / f"wanglai_export_{now}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    log(f"[备份] CSV 已保存: {csv_path}（{len(df)} 行 × {len(df.columns)} 列）")
    return csv_path


# ============ 飞书多维表：清空 + 写入 ============

def bitable_list_all_records(token: str) -> list:
    """返回多维表内所有 record_id（⚠️ search 接口的 page_token 会永远不推进，必须用 GET list 接口 + query string）"""
    h = {"Authorization": f"Bearer {token}"}
    base_url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records"
    all_ids = []
    seen_ids = set()
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
            if rid and rid not in seen_ids:
                all_ids.append(rid)
                seen_ids.add(rid)
            else:
                dup += 1
        if dup == len(items):
            log(f"[飞书] ⚠️ page_token 未推进，停止；累计 {len(all_ids)} 条")
            break
        if len(all_ids) % 5000 == 0:
            log(f"[飞书] list: 累计 {len(all_ids)}")
        if not d.get("has_more"):
            break
        page_token = d.get("page_token")
        if not page_token:
            break
    log(f"[飞书] list 完成: 共 {len(all_ids)} 条（去重后）")
    return all_ids


def bitable_batch_delete(token: str, record_ids: list):
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records/batch_delete?ignore_consistency_check=true"
    total = len(record_ids)
    for i in range(0, total, BATCH_SIZE):
        batch = record_ids[i:i + BATCH_SIZE]
        r = requests.post(url, headers=h, json={"records": batch}, timeout=30)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_delete 失败: batch {i // BATCH_SIZE + 1}, err={data.get('msg')}")
        if (i // BATCH_SIZE + 1) % 10 == 0 or i + BATCH_SIZE >= total:
            log(f"[飞书] delete: {len(batch)} 条 (批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE})")


def _convert_value(raw, ftype: str):
    """把接口返回的单元格值转成飞书 API 接受的字段值"""
    if raw is None:
        return None
    if isinstance(raw, float) and math.isnan(raw):
        return None
    if ftype == "text":
        s = str(raw).strip()
        return s if s else None
    if ftype == "number":
        s = str(raw).strip()
        if not s:
            return None
        try:
            f = float(s)
            return int(f) if f == int(f) else f
        except ValueError:
            return None
    if ftype == "datetime":
        s = str(raw).strip()
        if not s:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
            try:
                dt = datetime.datetime.strptime(s, fmt)
                return int(dt.timestamp() * 1000)
            except ValueError:
                continue
        return None
    return None


def rows_to_bitable_records(rows: list) -> list:
    """把接口原始 dict 列表转成飞书 batch_create 需要的 [{"fields": {...}}, ...] 列表"""
    records = []
    for row in rows:
        fields = {}
        for col, ftype in BITABLE_FIELD_TYPES.items():
            val = _convert_value(row.get(col), ftype)
            if val is not None:
                fields[col] = val
        records.append({"fields": fields})
    return records


def bitable_batch_create(token: str, records: list):
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records/batch_create?ignore_consistency_check=true"
    total = len(records)
    created = 0
    for i in range(0, total, BATCH_SIZE):
        batch = records[i:i + BATCH_SIZE]
        r = requests.post(url, headers=h, json={"records": batch}, timeout=60)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_create 失败 batch {i // BATCH_SIZE + 1}: code={data.get('code')} msg={data.get('msg')} sample={str(data)[:500]}")
        created += len((data.get("data") or {}).get("records") or [])
        log(f"[飞书] create: {len(batch)} 条 (批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE})，累计已建 {created}")
    return created


# ============ 主流程 ============

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--skip-clear", action="store_true", help="跳过清空旧表（仅追加，不推荐，仅调试用）")
    ap.add_argument("--skip-upload", action="store_true", help="跳过写入多维表（只拉取+转 CSV）")
    ap.add_argument("--download-only", action="store_true", help="等价 --skip-upload --skip-clear")
    ap.add_argument("--dry-run", action="store_true", help="拉取+转 CSV+打印样例，不碰飞书")
    ap.add_argument("--no-notify", action="store_true", help="不发送飞书通知（默认同步完成会通知洪）")
    args = ap.parse_args()

    global CURRENT_STEP
    t0 = time.time()
    log("==== 懂火往来流水 → 飞书多维表 同步工作流 启动 ====")

    # ---- Step 1: 登录 + 拉取流水 ----
    rows = fetch_wanglai_rows()
    if not rows:
        raise RuntimeError("往来流水接口返回空，中止")

    # ---- Step 2: CSV 备份 ----
    CURRENT_STEP = "落 CSV 备份"
    csv_path = rows_to_csv(rows)

    if args.dry_run:
        log("[DRY-RUN] 不操作飞书")
        log(f"  字段: {list(rows[0].keys())}")
        log(f"  前 3 行: {json.dumps(rows[:3], ensure_ascii=False)[:1000]}")
        elapsed = time.time() - t0
        log(f"==== 完成（DRY-RUN），耗时 {elapsed:.1f}s ====")
        return 0

    if args.download_only:
        elapsed = time.time() - t0
        log(f"==== 完成（仅拉取+备份），耗时 {elapsed:.1f}s ====")
        return 0

    # ---- Step 3: 飞书 ----
    CURRENT_STEP = "获取飞书凭证"
    token = feishu_token()
    log("[飞书] ✅ tenant_access_token 已获取")

    cleared_count = None
    if not args.skip_clear:
        CURRENT_STEP = "清空多维表旧记录"
        log("[飞书] Step A: 清空现有记录 ...")
        ids = bitable_list_all_records(token)
        log(f"[飞书] 现有记录数: {len(ids)}")
        if ids:
            bitable_batch_delete(token, ids)
            log("[飞书] ✅ 已清空")
        else:
            log("[飞书] 无需清空（表已空）")
        cleared_count = len(ids)
    else:
        log("[飞书] 跳过清空（--skip-clear）")

    if not args.skip_upload:
        CURRENT_STEP = "预检写入（前 5 条）"
        log(f"[飞书] Step B: 写入 {len(rows)} 条新记录 ...")
        records = rows_to_bitable_records(rows)
        # 先小批量试写，验证字段类型/名称没问题
        test_batch = records[:min(5, len(records))]
        h = {"Authorization": f"Bearer {token}"}
        url_test = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records/batch_create"
        r = requests.post(url_test, headers=h, json={"records": test_batch}, timeout=60)
        d = r.json()
        if d.get("code") != 0:
            log(f"[飞书] ❌ 预检写入失败: code={d.get('code')} msg={d.get('msg')}")
            log(f"  样本记录: {json.dumps(test_batch[0], ensure_ascii=False)[:800]}")
            raise RuntimeError(f"预检写入失败: code={d.get('code')} msg={d.get('msg')}")
        log(f"[飞书] ✅ 预检写入 {len(test_batch)} 条成功，继续写剩余 {len(records) - len(test_batch)} 条")
        CURRENT_STEP = "批量写入多维表"
        remaining = records[len(test_batch):]
        if remaining:
            bitable_batch_create(token, remaining)
        total_written = len(test_batch) + len(remaining)
        log(f"[飞书] ✅ 全部写入完成，共 {total_written} 条")
        # 上传成功后删除 CSV 备份
        if csv_path.exists():
            csv_path.unlink()
            log(f"[清理] CSV 已删除: {csv_path.name}")
    else:
        log("[飞书] 跳过写入（--skip-upload）")
        total_written = 0

    elapsed = time.time() - t0
    log(f"==== 完成，耗时 {elapsed:.1f}s ====")

    # 飞书通知（上传成功后才通知，--no-notify 可跳过）
    if not args.no_notify and not args.skip_upload and not args.dry_run and not args.download_only:
        try:
            CURRENT_STEP = "发送飞书通知"
            token = feishu_token()
            card = build_success_card(total_written, cleared_count, elapsed)
            feishu_send_card(FEISHU_NOTIFY_UNION_ID, card, token)
        except Exception as e:
            log(f"[飞书通知] ⚠️ 发送异常: {e}")

    return 0


if __name__ == "__main__":
    try:
        rc = main()
    except Exception as e:
        log(f"❌ 未捕获异常: {e}")
        traceback.print_exc()
        # 失败通知卡片（尽力而为，发送失败也不影响退出码）
        try:
            _t = feishu_token()
            feishu_send_card(FEISHU_NOTIFY_UNION_ID, build_failure_card(CURRENT_STEP, str(e)), _t)
        except Exception as ne:
            log(f"[飞书通知] ⚠️ 失败通知发送异常: {ne}")
        rc = 99
    sys.exit(rc)
