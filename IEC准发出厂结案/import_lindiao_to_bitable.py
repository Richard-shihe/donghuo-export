#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IEC 准发捆包数据 → 飞书多维表（中间表）
=====================================

业务流程（两段式）：
  第一段（本脚本）：IEC bundle → 飞书多维表「入库管理」（中间表）
  第二段（后续）：人工审核中间表 → 脚本同步到懂火临调库存（本次不实现）

脚本动作：
  1. 飞书 Bitable 自动建 27 列字段（如已存在跳过）
  2. 调 download_bundle.py 的 download_bundles(type='bundle') 拉 IEC 捆包数据
  3. 按用户确认的 6 项业务决策映射为懂火 26 字段
  4. 按捆包号去重（拉表里已有捆包号集合，跳过已存在）
  5. 批量 batch_create 写入多维表（分批 500 条）
  6. 飞书机器人通知 + 本地 audit CSV 留底

运行方式:
  # 预览模式（DRY-RUN），只打印计划，不实际写入
  python import_lindiao_to_bitable.py --dry-run

  # 实际执行（默认全量去重写入）
  python import_lindiao_to_bitable.py

  # 只写 5 条（调试用）
  python import_lindiao_to_bitable.py --limit 5

  # 关闭捆包号去重（重复跑允许重复写入）
  python import_lindiao_to_bitable.py --no-dedup

  # 指定 IEC 月份范围
  python import_lindiao_to_bitable.py --start 202607 --end 202609

依赖环境变量:
  IBAO_USERNAME / IBAO_PASSWORD        宝钢 IEC 账号
  FEISHU_APP_ID / FEISHU_APP_SECRET    飞书自建应用凭据
  FEISHU_WEBHOOK_URL                   飞书机器人 webhook（可选，缺则跳过通知）
  FEISHU_WEBHOOK_SECRET                飞书机器人加签密钥（可选）
  BITABLE_APP_TOKEN                    多维表 app_token（默认 OSuobf2ZkaWtUAsXEE9c3aBTnwh）
  BITABLE_TABLE_ID                     多维表 table_id（默认 tblrhqzHuTsAprU3）
  DRY_RUN=1                            预览模式
  IMPORT_LIMIT=5                       最多处理 N 条
  BITABLE_DEDUP=true                   按捆包号去重（默认 true，false 关闭）
  TZ=Asia/Shanghai                     时区

前置准备:
  1. 飞书自建应用（App ID cli_aaf0ce1e9ef89d27）已加为 Bitable OSuobf2ZkaWtUAsXEE9c3aBTnwh 可编辑协作者
  2. 应用已开通 bitable:app 权限
  3. 建好字段后，在多维表「字段权限管理」把所有字段对应用授权可读写
"""
from __future__ import annotations

import argparse
import csv
import datetime
import hashlib
import hmac
import json
import os
import sys
import time
from typing import Optional

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from download_bundle import download_bundles, _default_range, _parse_month


# ============================================================
# 常量
# ============================================================
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}

# 目标多维表（用户指定的中间表）
BITABLE_APP_TOKEN_DEFAULT = "OSuobf2ZkaWtUAsXEE9c3aBTnwh"
BITABLE_TABLE_ID_DEFAULT = "tblrhqzHuTsAprU3"

# 飞书 Bitable field type code
#   1=多行文本、2=数字、3=单选、4=多选、5=日期时间、7=复选框、11=人员、
#  13=电话、15=超链接、17=附件、18=单向关联、19=查找、20=公式、21=双向关联、
#  22=地理位置、23=群组、1001=创建时间、1002=最后更新时间、1003=创建人、1004=修改人、1005=自动编号
TYPE_TEXT = 1
TYPE_NUMBER = 2
TYPE_DATETIME = 5

# 27 列字段定义：(字段名, 飞书 type code)
# 前 26 列对齐懂火临调库存，第 27 列「批次时间」用于去重/批次识别
BITABLE_FIELDS_27: list[tuple[str, int]] = [
    ("所属公司", TYPE_TEXT), ("货权", TYPE_TEXT), ("品名", TYPE_TEXT),
    ("规格", TYPE_TEXT), ("材质", TYPE_TEXT), ("产地", TYPE_TEXT),
    ("等级", TYPE_TEXT), ("锌层", TYPE_TEXT), ("涂料", TYPE_TEXT),
    ("结构", TYPE_TEXT), ("颜色", TYPE_TEXT), ("件(张)数", TYPE_NUMBER),
    ("米数", TYPE_TEXT), ("重量", TYPE_NUMBER), ("供应商", TYPE_TEXT),
    ("采购单价", TYPE_NUMBER), ("成本单价", TYPE_NUMBER),
    ("销售单价", TYPE_NUMBER), ("仓库", TYPE_TEXT), ("库位号", TYPE_TEXT),
    ("捆包号", TYPE_TEXT), ("合同号", TYPE_TEXT), ("车船号", TYPE_TEXT),
    ("提单号", TYPE_TEXT), ("备注", TYPE_TEXT), ("入库日期", TYPE_DATETIME),
    ("批次时间", TYPE_DATETIME),
]

# 业务默认值（用户 6 项决策）
DEFAULT_COMPANY = "上海士禾实业有限公司"
DEFAULT_HUOQUAN = "拥有"

# batch_create 单次最多 500 条
BITABLE_BATCH_SIZE = 500

# 时区：Asia/Shanghai (UTC+8)，硬编码避免依赖系统时区
SHANGHAI_OFFSET_HOURS = 8


# ============================================================
# 工具
# ============================================================
def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def env_bool(name: str, default: bool = False) -> bool:
    v = env(name).lower()
    if not v:
        return default
    return v in ("1", "true", "yes", "on")


def _to_bitable_text(val) -> str:
    s = str(val).strip() if val is not None else ""
    return s


def _to_bitable_number(val):
    s = str(val).strip() if val is not None else ""
    if not s:
        return None
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return None


def _to_bitable_datetime(val):
    """日期字符串 → 毫秒时间戳；支持常见格式"""
    s = str(val).strip() if val is not None else ""
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d",
                "%Y%m%d"):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return int(dt.timestamp() * 1000)
        except ValueError:
            continue
    return None


def _convert_field(val, ftype: int):
    """根据飞书字段 type code 转换值"""
    if ftype == TYPE_TEXT:
        return _to_bitable_text(val)
    if ftype == TYPE_NUMBER:
        return _to_bitable_number(val)
    if ftype == TYPE_DATETIME:
        return _to_bitable_datetime(val)
    return None


def _mm_to_m(val) -> str:
    """毫米 → 米（保留 2 位小数），空则返回空串"""
    n = _to_bitable_number(val)
    if n is None:
        return ""
    m = n / 1000.0
    return f"{m:.2f}"


def _feishu_sign(secret: str, timestamp: str) -> str:
    """飞书机器人加签"""
    import base64 as _b64
    string_to_sign = f"{timestamp}\n{secret}"
    h = hmac.new(secret.encode("utf-8"),
                 string_to_sign.encode("utf-8"),
                 hashlib.sha256)
    return _b64.b64encode(h.digest()).decode("utf-8")


def now_shanghai_ts_ms() -> int:
    """当前时间戳（毫秒，按 Asia/Shanghai 算）"""
    utc_now = datetime.datetime.utcnow()
    sh_now = utc_now + datetime.timedelta(hours=SHANGHAI_OFFSET_HOURS)
    return int(sh_now.timestamp() * 1000)


def now_shanghai_str() -> str:
    """当前 Asia/Shanghai 时间字符串 'YYYY-MM-DD HH:MM:SS'"""
    utc_now = datetime.datetime.utcnow()
    sh_now = utc_now + datetime.timedelta(hours=SHANGHAI_OFFSET_HOURS)
    return sh_now.strftime("%Y-%m-%d %H:%M:%S")


# ============================================================
# 飞书：认证 / 多维表字段 / 记录读写
# ============================================================
def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": app_id, "app_secret": app_secret},
                      timeout=15, proxies=NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换取 tenant_access_token 失败: {data}")
    token = str(data.get("tenant_access_token") or "")
    if not token:
        raise RuntimeError(f"响应中无 tenant_access_token: {data}")
    print(f"[飞书] tenant_access_token 获取成功 (len={len(token)})")
    return token


def bitable_list_fields(token: str, app_token: str, table_id: str) -> list[dict]:
    """列出多维表已有字段"""
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/fields")
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                     timeout=15, proxies=NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"list_fields 失败: code={data.get('code')}, "
                           f"msg={data.get('msg')}")
    return (data.get("data") or {}).get("items") or []


def bitable_create_field(token: str, app_token: str, table_id: str,
                         field_name: str, type_code: int) -> dict:
    """创建单个字段"""
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/fields")
    body = {"field_name": field_name, "type": type_code}
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"},
                      json=body, timeout=15, proxies=NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"创建字段 {field_name!r} 失败: "
                           f"code={data.get('code')}, msg={data.get('msg')}")
    return (data.get("data") or {}).get("field") or {}


def bitable_ensure_fields(token: str, app_token: str, table_id: str,
                          fields_def: list[tuple[str, int]]) -> dict:
    """查重 + 创建缺失字段。返回 {字段名: field_id}"""
    existing = bitable_list_fields(token, app_token, table_id)
    existing_names = {f.get("field_name", "") for f in existing}
    name_to_id = {f.get("field_name", ""): f.get("field_id", "")
                  for f in existing}
    print(f"[bitable] 已有字段 {len(existing_names)} 个: "
          f"{sorted(existing_names)[:10]}...")

    created_count = 0
    for name, type_code in fields_def:
        if name in existing_names:
            continue
        print(f"[bitable] 创建字段 {name!r} (type={type_code})...")
        try:
            f = bitable_create_field(token, app_token, table_id,
                                     name, type_code)
            name_to_id[name] = f.get("field_id", "")
            created_count += 1
            time.sleep(0.2)  # 避免过快被限流
        except Exception as e:
            # 字段已存在/权限不足等不阻断主流程，最后 batch_create 会报具体错
            print(f"  ⚠️  创建字段 {name!r} 失败: {e}")
    print(f"[bitable] 字段就绪：新建 {created_count} 个，"
          f"共 {len(name_to_id)} 个")
    return name_to_id


def bitable_batch_create(token: str, app_token: str, table_id: str,
                          records: list[dict]) -> list[dict]:
    """batch_create API，单次最多 500 条，自动分批"""
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/records/batch_create")
    headers = {"Authorization": f"Bearer {token}"}
    created: list[dict] = []
    total = len(records)
    if total == 0:
        return created
    for start in range(0, total, BITABLE_BATCH_SIZE):
        batch = records[start:start + BITABLE_BATCH_SIZE]
        r = requests.post(url, headers=headers,
                          json={"records": batch}, timeout=60,
                          proxies=NO_PROXY)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(
                f"bitable 批量创建失败: code={data.get('code')}, "
                f"msg={data.get('msg')}, "
                f"raw={json.dumps(data, ensure_ascii=False)[:1200]}"
            )
        items = (data.get("data") or {}).get("records") or []
        created.extend(items)
        print(f"  [bitable] 写入 {start + len(batch)}/{total} "
              f"(本次 {len(items)} 条)", flush=True)
        time.sleep(0.3)  # 限速
    return created


def bitable_list_all_bundle_numbers(token: str, app_token: str,
                                      table_id: str) -> set[str]:
    """拉表里所有记录的「捆包号」字段值，返回 set（用于去重）"""
    headers = {"Authorization": f"Bearer {token}"}
    bundle_set: set[str] = set()
    page_token = ""
    page = 1
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/records")
    while True:
        params: dict = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(url, headers=headers, params=params,
                         timeout=30, proxies=NO_PROXY)
        data = r.json()
        if data.get("code") != 0:
            print(f"  [bitable] list 错误: code={data.get('code')}, "
                  f"msg={data.get('msg')}")
            break
        items = (data.get("data") or {}).get("items") or []
        for it in items:
            v = (it.get("fields") or {}).get("捆包号")
            if isinstance(v, list):
                # 飞书 text 字段返回 [{type:text,text:...}] 结构
                v = "".join(seg.get("text", "") if isinstance(seg, dict)
                            else str(seg) for seg in v)
            if v:
                bundle_set.add(str(v).strip())
        page_token = (data.get("data") or {}).get("page_token") or ""
        has_more = (data.get("data") or {}).get("has_more")
        print(f"  [bitable] 去重拉取 page {page}: +{len(items)}, "
              f"累计捆包号 {len(bundle_set)} 个, has_more={has_more}")
        if not has_more or not page_token:
            break
        page += 1
        time.sleep(0.1)
    return bundle_set


def feishu_send_bot_text(webhook_url: str, secret: str, text: str) -> None:
    if not webhook_url:
        print("[飞书通知] 未配置 FEISHU_WEBHOOK_URL，跳过通知")
        return
    body = {"msg_type": "text", "content": {"text": text}}
    if secret:
        timestamp = str(int(time.time()))
        body["timestamp"] = timestamp
        body["sign"] = _feishu_sign(secret, timestamp)
    r = requests.post(webhook_url, json=body, timeout=15, proxies=NO_PROXY)
    resp = r.json()
    if resp.get("code") not in (0, None) and resp.get("StatusCode") != 0:
        raise RuntimeError(f"飞书机器人通知失败: {resp}")
    print("[飞书通知] 发送成功")


# ============================================================
# IEC → 懂火 字段映射
# ============================================================
def map_iec_to_donghuo(iec_row: dict) -> dict:
    """IEC bundle 行 → 懂火 26 字段（按用户 6 项业务决策）"""
    def _g(key):
        v = iec_row.get(key, "")
        if v is None:
            return ""
        return str(v).strip()

    return {
        "所属公司": DEFAULT_COMPANY,
        "货权": DEFAULT_HUOQUAN,
        "品名": _g("产品大类"),
        "规格": _g("规格"),
        "材质": _g("牌号"),
        "产地": _g("生产厂家"),
        "等级": _g("性能级别"),
        "锌层": _g("锌层重量"),
        "涂料": "/".join(filter(None, [_g("涂料种类(上)"),
                                       _g("涂料种类(下)")])),
        "结构": _g("表面结构"),
        "颜色": "/".join(filter(None, [_g("涂料颜色(上)"),
                                       _g("涂料颜色(下)")])),
        "件(张)数": _g("总张数") or _g("件数"),
        "米数": _mm_to_m(_g("实际长度")),
        "重量": _g("净重"),
        "供应商": _g("生产厂家"),
        # 采购单价/成本单价/销售单价 留空（人工补）
        "仓库": _g("仓库"),
        "库位号": _g("库位"),
        "捆包号": _g("捆包号"),
        "合同号": _g("销售合同号"),
        # 车船号/提单号/备注 留空
        "入库日期": _g("交库日"),
    }


def build_records_payload(mapped_rows: list[dict],
                          batch_ts_ms: int) -> list[dict]:
    """把映射后的 dict 列表转成 batch_create 的 records 格式"""
    type_map = dict(BITABLE_FIELDS_27)
    payload: list[dict] = []
    skipped_fields: dict[str, int] = {}
    for row in mapped_rows:
        fields: dict = {}
        for col_name, ftype in BITABLE_FIELDS_27:
            if col_name == "批次时间":
                fields[col_name] = batch_ts_ms
                continue
            v = _convert_field(row.get(col_name, ""), ftype)
            if v is None or (isinstance(v, str) and not v):
                continue
            fields[col_name] = v
        for k in row.keys():
            if k not in type_map:
                skipped_fields[k] = skipped_fields.get(k, 0) + 1
        payload.append({"fields": fields})
    if skipped_fields:
        print(f"[映射] 跳过的未映射字段: {skipped_fields}")
    return payload


# ============================================================
# audit CSV 留底
# ============================================================
def write_audit_csv(rows: list[dict], stats: dict, ts_str: str) -> str:
    """写本地 audit CSV（捆包号 + 状态 + 统计信息）"""
    filename = f"lindiao_import_audit_{ts_str.replace(':', '').replace(' ', '_')}.csv"
    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["捆包号", "品名", "规格", "重量", "仓库", "合同号",
                    "入库日期", "状态", "说明"])
        for r in rows:
            w.writerow([
                r.get("捆包号", ""),
                r.get("品名", ""),
                r.get("规格", ""),
                r.get("重量", ""),
                r.get("仓库", ""),
                r.get("合同号", ""),
                r.get("入库日期", ""),
                r.get("_status", ""),
                r.get("_msg", ""),
            ])
        w.writerow([])
        w.writerow(["统计项", "值"])
        for k, v in stats.items():
            w.writerow([k, v])
    print(f"[audit] 留底 CSV: {filename}")
    return filename


# ============================================================
# 主流程
# ============================================================
def parse_args():
    default_start, default_end = _default_range()
    p = argparse.ArgumentParser(
        description="IEC bundle → 飞书多维表（中间表）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
示例:
  python import_lindiao_to_bitable.py --dry-run
  python import_lindiao_to_bitable.py --limit 5
  python import_lindiao_to_bitable.py --start {default_start} --end {default_end}
  python import_lindiao_to_bitable.py --no-dedup
""")
    p.add_argument("--dry-run", action="store_true",
                   help="预览模式：只打印计划，不实际写多维表")
    p.add_argument("--limit", type=int, default=0,
                   help="最多写入 N 条（0=不限制，调试用）")
    p.add_argument("--no-dedup", action="store_true",
                   help="关闭捆包号去重（允许重复写入）")
    p.add_argument("-s", "--start", type=_parse_month, default=default_start,
                   help=f"IEC 起始月份 YYYYMM（默认 {default_start}）")
    p.add_argument("-e", "--end", type=_parse_month, default=default_end,
                   help=f"IEC 结束月份 YYYYMM（默认 {default_end}）")
    args = p.parse_args()

    # 环境变量覆盖
    if env_bool("DRY_RUN") and not args.dry_run:
        args.dry_run = True
    if not args.limit:
        v = env("IMPORT_LIMIT")
        if v.isdigit():
            args.limit = int(v)
    if not args.no_dedup and env_bool("BITABLE_DEDUP", default=True) is False:
        args.no_dedup = True
    sm = env("IEC_START_MONTH")
    if sm:
        args.start = _parse_month(sm)
    em = env("IEC_END_MONTH")
    if em:
        args.end = _parse_month(em)
    return args


def main() -> int:
    args = parse_args()
    ts_str = now_shanghai_str()
    ts_ms = now_shanghai_ts_ms()

    print("=" * 60)
    print(f"IEC bundle → 飞书多维表（中间表）")
    print(f"时间: {ts_str}")
    print(f"模式: {'DRY-RUN（预览，不实际写入）' if args.dry_run else '实际写入'}")
    print(f"月份范围: {args.start} ~ {args.end}")
    print(f"limit: {args.limit or '不限'}")
    print(f"去重: {'关闭' if args.no_dedup else '开启（按捆包号）'}")
    print("=" * 60)

    # 1. 检查 IEC 凭据
    ibao_user = env("IBAO_USERNAME")
    ibao_pass = env("IBAO_PASSWORD")
    if not ibao_user or not ibao_pass:
        print("::error:: 缺少 IBAO_USERNAME / IBAO_PASSWORD 环境变量")
        return 1

    # 2. 飞书 token
    fs_app_id = env("FEISHU_APP_ID")
    fs_app_secret = env("FEISHU_APP_SECRET")
    if not fs_app_id or not fs_app_secret:
        print("::error:: 缺少 FEISHU_APP_ID / FEISHU_APP_SECRET 环境变量")
        return 1
    try:
        token = feishu_tenant_access_token(fs_app_id, fs_app_secret)
    except Exception as e:
        print(f"::error:: 飞书认证失败: {e}")
        return 1

    app_token = env("BITABLE_APP_TOKEN", BITABLE_APP_TOKEN_DEFAULT)
    table_id = env("BITABLE_TABLE_ID", BITABLE_TABLE_ID_DEFAULT)
    print(f"[bitable] 目标表: app={app_token}, table={table_id}")

    # 3. 建字段（DRY-RUN 也执行建字段，让用户能看到字段结构）
    try:
        bitable_ensure_fields(token, app_token, table_id, BITABLE_FIELDS_27)
    except Exception as e:
        print(f"::error:: 建字段失败: {e}")
        return 1

    # 4. 拉 IEC bundle 数据
    print(f"\n[IEC] 拉取捆包级数据 {args.start}~{args.end}...")
    try:
        df, xlsx_path = download_bundles(
            start=args.start, end=args.end,
            download_type='bundle',
        )
    except Exception as e:
        print(f"::error:: IEC 下载失败: {e}")
        return 1

    print(f"[IEC] 拉到 {len(df)} 行 × {len(df.columns)} 列")
    if len(df) == 0:
        print("::error:: IEC 返回 0 行数据，无可写入内容")
        return 0
    print(f"[IEC] 列名: {list(df.columns)}")

    # 5. 转为 dict 行 + 字段映射
    iec_rows = df.fillna("").to_dict(orient="records")
    mapped_rows = [map_iec_to_donghuo(r) for r in iec_rows]

    # 6. 去重
    skipped_dup = 0
    if not args.no_dedup:
        print(f"\n[bitable] 拉取表里已有捆包号集合（用于去重）...")
        try:
            existing_bundles = bitable_list_all_bundle_numbers(
                token, app_token, table_id)
        except Exception as e:
            print(f"  ⚠️  拉取已有捆包号失败: {e}，跳过去重")
            existing_bundles = set()
        print(f"[bitable] 表里已有 {len(existing_bundles)} 个捆包号")

        filtered_rows = []
        for r in mapped_rows:
            kb = str(r.get("捆包号", "")).strip()
            if not kb:
                r["_status"] = "失败"
                r["_msg"] = "捆包号为空"
                continue
            if kb in existing_bundles:
                skipped_dup += 1
                r["_status"] = "已存在"
                r["_msg"] = "表里已有此捆包号，跳过"
                continue
            r["_status"] = "待写入"
            r["_msg"] = ""
            filtered_rows.append(r)
        mapped_rows_to_write = filtered_rows
    else:
        for r in mapped_rows:
            r["_status"] = "待写入"
            r["_msg"] = ""
        mapped_rows_to_write = mapped_rows

    print(f"[去重] 待写入 {len(mapped_rows_to_write)} 条，"
          f"已跳过重复 {skipped_dup} 条")

    # 7. 应用 limit
    if args.limit and args.limit > 0:
        mapped_rows_to_write = mapped_rows_to_write[:args.limit]
        print(f"[limit] 截断为前 {len(mapped_rows_to_write)} 条")

    # 8. 打印预览（无论 dry-run 都打印 1-2 行示例）
    if mapped_rows_to_write:
        print("\n[预览] 第 1 条映射结果:")
        sample = {k: v for k, v in mapped_rows_to_write[0].items()
                  if not k.startswith("_")}
        for k, v in sample.items():
            print(f"  {k}: {v!r}")

    # 9. DRY-RUN 在此止步
    if args.dry_run:
        print(f"\n⏭  DRY-RUN 模式：不实际写入。计划写入 "
              f"{len(mapped_rows_to_write)} 条到 Bitable {app_token}/{table_id}")
        stats = {
            "iec_total_rows": len(iec_rows),
            "skipped_duplicate": skipped_dup,
            "to_write": len(mapped_rows_to_write),
            "dry_run": "true",
        }
        write_audit_csv(mapped_rows_to_write, stats, ts_str)
        return 0

    # 10. 批量写入
    if not mapped_rows_to_write:
        print("\n⏭  无新数据可写入（全部已存在或为空）")
        stats = {
            "iec_total_rows": len(iec_rows),
            "skipped_duplicate": skipped_dup,
            "to_write": 0,
            "written": 0,
        }
        write_audit_csv(mapped_rows_to_write, stats, ts_str)
        return 0

    print(f"\n[bitable] 构造写入 payload...")
    records_payload = build_records_payload(mapped_rows_to_write, ts_ms)
    print(f"[bitable] 准备写入 {len(records_payload)} 条到表 {table_id}")

    try:
        created = bitable_batch_create(token, app_token, table_id,
                                       records_payload)
    except Exception as e:
        print(f"::error:: 写入失败: {e}")
        for r in mapped_rows_to_write:
            if r.get("_status") == "待写入":
                r["_status"] = "失败"
                r["_msg"] = str(e)[:200]
        stats = {
            "iec_total_rows": len(iec_rows),
            "skipped_duplicate": skipped_dup,
            "to_write": len(mapped_rows_to_write),
            "written": 0,
            "error": str(e)[:200],
        }
        write_audit_csv(mapped_rows_to_write, stats, ts_str)
        return 1

    # 11. 标记成功
    for r in mapped_rows_to_write:
        r["_status"] = "已写入"
        r["_msg"] = ""
    print(f"\n✅ 成功写入 {len(created)} 条记录")

    # 12. 飞书机器人通知
    stats = {
        "iec_total_rows": len(iec_rows),
        "skipped_duplicate": skipped_dup,
        "to_write": len(mapped_rows_to_write),
        "written": len(created),
        "bitable": f"{app_token}/{table_id}",
        "timestamp": ts_str,
    }
    try:
        notify_text = (
            f"✅ IEC 准发捆包 → 飞书中间表 完成\n"
            f"时间: {ts_str}\n"
            f"IEC 拉取: {stats['iec_total_rows']} 行\n"
            f"去重跳过: {stats['skipped_duplicate']} 条\n"
            f"实际写入: {stats['written']} 条\n"
            f"Bitable: {app_token}/{table_id}\n"
        )
        feishu_send_bot_text(env("FEISHU_WEBHOOK_URL"),
                              env("FEISHU_WEBHOOK_SECRET"),
                              notify_text)
    except Exception as e:
        print(f"⚠️  飞书通知发送失败: {e}")

    # 13. audit CSV 留底
    write_audit_csv(mapped_rows_to_write, stats, ts_str)

    print(f"\n🎉 完成。详情见 audit CSV。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
