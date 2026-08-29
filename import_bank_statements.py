#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银行流水 → 飞书多维表（流水（士禾））
====================================

功能：
  读取 Finance/ 文件夹下的建行和中行每日流水 Excel，
  清洗后批量写入飞书多维表 tblBAUOPWissWQmX。

用法：
  # 预览模式（不落库）
  python import_bank_statements.py --dry-run

  # 实际执行
  python import_bank_statements.py

  # 指定文件
  python import_bank_statements.py --jh 建行流水.xls --zh 中行流水.xls

  # 指定 Finance 目录
  python import_bank_statements.py --finance-dir Finance/

环境变量：
  FEISHU_APP_ID / FEISHU_APP_SECRET   飞书自建应用
  BITABLE_APP_TOKEN                    多维表 app_token (默认 IvUNbxaVCaNrl2sZC5qcfFADn3c)
  BANK_BITABLE_TABLE_ID                多维表 table_id (默认 tblBAUOPWissWQmX)
  DRY_RUN=1                            等价 --dry-run
"""

from __future__ import annotations

import argparse
import datetime
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
import requests

# ============================================================
# 常量
# ============================================================
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}

DEFAULT_APP_TOKEN = "IvUNbxaVCaNrl2sZC5qcfFADn3c"
DEFAULT_TABLE_ID = "tblBAUOPWissWQmX"
DEFAULT_FINANCE_DIR = "Finance"

# Bitable 字段名
FIELD_BANK = "银行"
FIELD_METHOD = "方式"
FIELD_TITLE = "抬头"
FIELD_DATE = "日期"
FIELD_COUNTERPARTY = "对象单位"
FIELD_DEBIT = "借"       # 支出金额（正数）
FIELD_CREDIT = "贷"      # 收入金额（正数）
FIELD_AMOUNT = "本金"
FIELD_SUMMARY = "摘要"
FIELD_ORIGINAL_NO = "收款单号（原）"

# 抬头固定值
TITLE_VALUE = "士禾"

# 中行业务类型 → Bitable 方式字段映射（HISXLS 格式）
# Bitable 单选只有: 现汇, 承兑
BOC_METHOD_MAP = {
    "小额普通": "现汇",
    "小额实时": "现汇",
    "网上支付": "现汇",
    "收费": "现汇",
    "现金": "现汇",
    "转账": "现汇",
    "汇兑": "现汇",
    "承兑": "承兑",
    "委托收款": "现汇",
    "托收承付": "现汇",
    "信用证": "现汇",
}

# 建行摘要关键词 → 方式映射（A058 格式，无业务类型列，用摘要关键词判断）
CCB_METHOD_MAP = {
    "承兑": "承兑",
}

# ============================================================
# 工具函数
# ============================================================
def env(name: str, default: str = "") -> str:
    """读环境变量，支持 .env 文件"""
    val = os.environ.get(name, "").strip()
    if val:
        return val
    # 尝试从 .env 文件加载
    env_path = Path(__file__).parent / ".env"
    if env_path.exists():
        with open(env_path, "r") as f:
            for line in f:
                line = line.strip()
                if line.startswith(f"{name}="):
                    return line.split("=", 1)[1].strip()
    return default


def parse_amount(val) -> float | None:
    """解析金额，处理千分位逗号和负数"""
    if val is None:
        return None
    s = str(val).strip().replace(",", "").replace("¥", "").replace(" ", "")
    if not s or s == "nan":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def date_to_ts_ms(date_str: str) -> int | None:
    """日期字符串 → 毫秒时间戳"""
    if not date_str:
        return None
    s = str(date_str).strip()
    for fmt in (
        "%Y%m%d",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y%m%d %H:%M:%S",
        "%Y%m%d%H%M%S",
    ):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return int(dt.timestamp() * 1000)
        except ValueError:
            continue
    # 尝试只取日期部分（前8位数字）
    match = re.match(r"(\d{8})", s)
    if match:
        try:
            dt = datetime.datetime.strptime(match.group(1), "%Y%m%d")
            return int(dt.timestamp() * 1000)
        except ValueError:
            pass
    return None


def gen_shoukuandan_no(dt_ts_ms: int, seq: int, prefix: str = "SH") -> str:
    """生成收款单号，格式: SH + YYYYMMDD + -XXX"""
    dt = datetime.datetime.fromtimestamp(dt_ts_ms / 1000)
    return f"{prefix}{dt.strftime('%Y%m%d')}-{seq:03d}"


def gen_original_no(dt_ts_ms: int, seq: int, prefix: str = "SH") -> str:
    """生成原始编号，格式: SH + YYYYMMDD + XX + 三位序号"""
    dt = datetime.datetime.fromtimestamp(dt_ts_ms / 1000)
    return f"{prefix}{dt.strftime('%Y%m%d')}{seq:03d}"


# ============================================================
# 飞书
# ============================================================
def feishu_get_token() -> str:
    app_id = env("FEISHU_APP_ID", "")
    app_secret = env("FEISHU_APP_SECRET", "")

    # 回退到通知应用
    if not app_id:
        app_id = env("FEISHU_NOTIFY_APP_ID", "")
        app_secret = env("FEISHU_NOTIFY_APP_SECRET", "")

    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")

    r = requests.post(
        f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret},
        timeout=15, proxies=NO_PROXY,
    )
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换 tenant_access_token 失败: {data}")
    return data["tenant_access_token"]


def bitable_list_existing(tok: str, app_token: str, table_id: str) -> set:
    """读取 Bitable 所有记录，返回 (银行, 对象单位, 日期timestamp, 本金, 摘要) 集合用于去重"""
    existing = set()
    page_token = ""
    for page_idx in range(100):
        url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(
            url, params=params,
            headers={"Authorization": f"Bearer {tok}"},
            timeout=30, proxies=NO_PROXY,
        )
        d = r.json()
        if d.get("code") != 0:
            print(f"  [警告] 读取已有记录失败 code={d.get('code')} msg={d.get('msg')}")
            break
        items = (d.get("data") or {}).get("items") or []
        for it in items:
            fields = it.get("fields") or {}
            bank = str(fields.get(FIELD_BANK, "")).strip()
            counterparty = str(fields.get(FIELD_COUNTERPARTY, "")).strip()
            date_raw = fields.get(FIELD_DATE, "")
            amount = fields.get(FIELD_AMOUNT, "")
            summary = str(fields.get(FIELD_SUMMARY, "")).strip()

            # 日期转 timestamp
            date_ts = 0
            if date_raw:
                try:
                    date_ts = int(date_raw)
                except (ValueError, TypeError):
                    pass

            # 金额标准化
            amt_val = 0
            if amount is not None and amount != "":
                try:
                    amt_val = float(amount)
                except (ValueError, TypeError):
                    pass

            if bank or counterparty or date_ts:
                existing.add((bank, counterparty, date_ts, round(amt_val, 2), summary))

        d_data = d.get("data") or {}
        if not d_data.get("has_more") or not d_data.get("page_token"):
            break
        page_token = d_data["page_token"]
        if page_idx % 5 == 0:
            print(f"  [读取已有记录] 第{page_idx}页, 累计 {len(existing)} 条唯一键", flush=True)
        time.sleep(0.2)

    print(f"  [读取已有记录] 共 {len(existing)} 条唯一键", flush=True)
    return existing


def bitable_batch_create(tok: str, app_token: str, table_id: str,
                         records: list, batch_size: int = 500) -> dict:
    """批量创建记录，返回 {ok, fail, errors}"""
    ok = 0
    fail = 0
    errors = []
    total_batches = (len(records) + batch_size - 1) // batch_size
    for batch_idx, i in enumerate(range(0, len(records), batch_size), 1):
        batch = records[i : i + batch_size]
        url = (
            f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
            f"/tables/{table_id}/records/batch_create"
        )
        r = requests.post(
            url, json={"records": batch},
            headers={"Authorization": f"Bearer {tok}"},
            timeout=60, proxies=NO_PROXY,
        )
        d = r.json()
        if d.get("code") == 0:
            ok += len(batch)
        else:
            fail += len(batch)
            errors.append(f"batch {batch_idx}: code={d.get('code')} msg={d.get('msg')}")
            print(f"  [错误] 批次 {batch_idx}: {d.get('code')} {d.get('msg')}", flush=True)
        if batch_idx % 5 == 0 or batch_idx == total_batches or fail:
            print(
                f"  [写入 {batch_idx}/{total_batches}] 成功 {ok} / 失败 {fail}",
                flush=True,
            )
        time.sleep(0.3)
    return {"ok": ok, "fail": fail, "errors": errors}


# ============================================================
# 中行流水解析（HISXLS 格式，账户 4520...）
# ============================================================
def parse_zhonghang(filepath: str) -> list[dict]:
    """解析中行每日流水 Excel，返回清洗后的记录列表"""
    print(f"[中行] 读取文件: {filepath}")

    # 中行流水前7行是表头信息，第8行是列名，从第9行开始是数据
    df = pd.read_excel(filepath, header=7, dtype=str)
    print(f"[中行] 原始行数: {len(df)}")

    # 过滤有效交易行（排除汇总行）
    valid_mask = df.iloc[:, 0].isin(["来账", "往账"])
    df_valid = df[valid_mask].copy()
    print(f"[中行] 有效交易行: {len(df_valid)} 条")

    records = []
    for _, row in df_valid.iterrows():
        tx_type = str(row.iloc[0]).strip()  # 来账/往账
        biz_type = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        payer_name = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else ""
        payee_name = str(row.iloc[9]).strip() if pd.notna(row.iloc[9]) else ""
        tx_date = str(row.iloc[10]).strip() if pd.notna(row.iloc[10]) else ""
        tx_amount_raw = str(row.iloc[13]).strip() if pd.notna(row.iloc[13]) else ""
        summary = str(row.iloc[23]).strip() if pd.notna(row.iloc[23]) else ""
        purpose = str(row.iloc[24]).strip() if pd.notna(row.iloc[24]) else ""

        # 金额处理
        amount = parse_amount(tx_amount_raw)
        if amount is None:
            continue

        # 日期 → 毫秒时间戳
        date_ts = date_to_ts_ms(tx_date)
        if date_ts is None:
            print(f"  [警告] 无法解析日期: {tx_date}")
            continue

        # 来账(收入)/往账(支出)
        if tx_type == "来账":
            debit_val = None
            credit_val = abs(amount)
            counterparty = payer_name  # 付款人 = 谁给我们打钱
        else:  # 往账
            debit_val = abs(amount)
            credit_val = None
            counterparty = payee_name  # 收款人 = 我们给谁打钱

        # 识别内部转账
        if "士禾" in payer_name and "士禾" in payee_name and payer_name == payee_name:
            counterparty = f"{TITLE_VALUE}（中行）"

        # 对方名称为空时的特殊标记
        if not counterparty or counterparty == "nan" or counterparty == "":
            counterparty = "手续费" if abs(amount) <= 10 else "(未知)"

        summary_final = summary or purpose

        rec = {
            FIELD_BANK: "中行",
            FIELD_METHOD: BOC_METHOD_MAP.get(biz_type, "现汇"),
            FIELD_TITLE: TITLE_VALUE,
            FIELD_DATE: date_ts,
            FIELD_COUNTERPARTY: counterparty,
            FIELD_DEBIT: debit_val,
            FIELD_CREDIT: credit_val,
            FIELD_AMOUNT: abs(amount),
            FIELD_SUMMARY: summary_final,
        }
        records.append(rec)

    print(f"[中行] 解析完成: {len(records)} 条记录")
    return records


# ============================================================
# 建行流水解析（A058 格式，账户 3100...）
# ============================================================
def parse_jianhang(filepath: str) -> list[dict]:
    """解析建行每日流水 Excel，返回清洗后的记录列表"""
    print(f"[建行] 读取文件: {filepath}")

    df = pd.read_excel(filepath, dtype=str)
    print(f"[建行] 原始行数: {len(df)}")

    records = []
    for _, row in df.iterrows():
        tx_time = str(row.get("交易时间", "")).strip()
        debit_raw = str(row.get("借方发生额（支取）", "")).strip()
        credit_raw = str(row.get("贷方发生额（收入）", "")).strip()
        counterparty = str(row.get("对方户名", "")).strip()
        summary = str(row.get("摘要", "")).strip()
        remark = str(row.get("备注", "")).strip()

        # 金额处理
        debit_val = parse_amount(debit_raw) or 0
        credit_val = parse_amount(credit_raw) or 0

        # 日期 → 毫秒时间戳
        date_ts = date_to_ts_ms(tx_time)
        if date_ts is None:
            print(f"  [警告] 无法解析日期: {tx_time}")
            continue

        # 判断方向
        if debit_val > 0:
            debit_val_final = debit_val
            credit_val_final = None
            amount = debit_val
        elif credit_val > 0:
            debit_val_final = None
            credit_val_final = credit_val
            amount = credit_val
        else:
            continue

        if not counterparty or counterparty == "nan":
            counterparty = "(未知)"

        summary_final = summary or remark

        # 判断方式
        method = "现汇"
        for keyword, mapped in CCB_METHOD_MAP.items():
            if keyword in summary:
                method = mapped
                break

        rec = {
            FIELD_BANK: "建行",
            FIELD_METHOD: method,
            FIELD_TITLE: TITLE_VALUE,
            FIELD_DATE: date_ts,
            FIELD_COUNTERPARTY: counterparty,
            FIELD_DEBIT: debit_val_final,
            FIELD_CREDIT: credit_val_final,
            FIELD_AMOUNT: amount,
            FIELD_SUMMARY: summary_final,
        }
        records.append(rec)

    print(f"[建行] 解析完成: {len(records)} 条记录")
    return records


# ============================================================
# 去重 & 构建写入记录
# ============================================================
def build_bitable_records(parsed: list[dict], existing_keys: set) -> tuple:
    """将解析后的记录与已有记录去重，返回 (待写入records, 跳过数, 新序号映射)"""
    to_write = []
    skipped = 0
    # 按日期分组用于生成序号
    date_seq_map: dict[int, int] = {}

    for rec in parsed:
        bank = rec.get(FIELD_BANK, "")
        counterparty = rec.get(FIELD_COUNTERPARTY, "")
        date_ts = rec.get(FIELD_DATE, 0)
        amount = rec.get(FIELD_AMOUNT, 0)
        summary = rec.get(FIELD_SUMMARY, "")

        # 构建唯一键（含银行，避免跨行误去重）
        key = (bank, counterparty, date_ts, round(float(amount or 0), 2), summary)

        if key in existing_keys:
            skipped += 1
            continue

        # 生成序号
        date_seq_map[date_ts] = date_seq_map.get(date_ts, 0) + 1
        seq = date_seq_map[date_ts]

        # 构建 Bitable records
        fields = {}
        for field_name, value in rec.items():
            if value is not None and value != "":
                fields[field_name] = value

        # 生成收款单号
        sh_no = gen_shoukuandan_no(date_ts, seq)
        orig_no = gen_original_no(date_ts, seq)
        fields[FIELD_ORIGINAL_NO] = orig_no

        to_write.append({"fields": fields})
        # 加入已有键防止重复
        existing_keys.add(key)

    return to_write, skipped


# ============================================================
# 主流程
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="银行流水 → 飞书多维表（流水（士禾））"
    )
    parser.add_argument(
        "--finance-dir", default=DEFAULT_FINANCE_DIR,
        help=f"Finance 文件夹路径 (默认: {DEFAULT_FINANCE_DIR})",
    )
    parser.add_argument(
        "--jh", default=None,
        help="建行流水文件名 (在 finance-dir 下，不指定则自动匹配)",
    )
    parser.add_argument(
        "--zh", default=None,
        help="中行流水文件名 (在 finance-dir 下，不指定则自动匹配)",
    )
    parser.add_argument(
        "--app-token", default=env("BITABLE_APP_TOKEN", DEFAULT_APP_TOKEN),
        help=f"多维表 app_token (默认: {DEFAULT_APP_TOKEN})",
    )
    parser.add_argument(
        "--table-id", default=env("BANK_BITABLE_TABLE_ID", DEFAULT_TABLE_ID),
        help=f"多维表 table_id (默认: {DEFAULT_TABLE_ID})",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        default=env("DRY_RUN", "") == "1",
        help="预览模式，不落库",
    )
    args = parser.parse_args()

    finance_dir = Path(args.finance_dir)
    if not finance_dir.exists():
        print(f"[错误] Finance 文件夹不存在: {finance_dir}")
        return 1

    # 查找银行流水文件
    jh_file = None
    zh_file = None

    if args.jh:
        jh_file = finance_dir / args.jh
    else:
        # 建行：文件名含 "A058"（账户 3100...）
        for f in finance_dir.glob("*.xls"):
            if "A058" in f.name or "建行" in f.name or "CCB" in f.name.upper():
                jh_file = f
                break

    if args.zh:
        zh_file = finance_dir / args.zh
    else:
        # 中行：文件名含 "HISXLS"（账户 4520...）
        for f in finance_dir.glob("*.xls"):
            if f != jh_file and ("HISXLS" in f.name or "中行" in f.name or "BOC" in f.name.upper()):
                zh_file = f
                break
        # 兜底：取第二个 .xls 文件
        if not zh_file:
            xls_files = sorted(finance_dir.glob("*.xls"))
            if jh_file and len(xls_files) >= 2:
                zh_file = [f for f in xls_files if f != jh_file][0]

    print(f"=== 银行流水 → Bitable 导入 ===")
    print(f"  App Token: {args.app_token}")
    print(f"  Table ID:  {args.table_id}")
    print(f"  建行文件:   {jh_file if jh_file else '(未找到)'}")
    print(f"  中行文件:   {zh_file if zh_file else '(未找到)'}")
    print(f"  Dry Run:   {args.dry_run}")
    print()

    # Step 1: 解析银行流水
    all_parsed = []

    if jh_file and jh_file.exists():
        try:
            jh_records = parse_jianhang(str(jh_file))
            all_parsed.extend(jh_records)
        except Exception as e:
            print(f"[错误] 解析建行文件失败: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("[警告] 未找到建行流水文件")

    if zh_file and zh_file.exists():
        try:
            zh_records = parse_zhonghang(str(zh_file))
            all_parsed.extend(zh_records)
        except Exception as e:
            print(f"[错误] 解析中行文件失败: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("[警告] 未找到中行流水文件")

    if not all_parsed:
        print("[错误] 没有可导入的记录")
        return 1

    print(f"\n[汇总] 共解析 {len(all_parsed)} 条记录")

    if args.dry_run:
        print("\n=== [DRY RUN] 预览前 10 条 ===")
        for i, rec in enumerate(all_parsed[:10]):
            print(f"\n  记录 {i + 1}:")
            for k, v in rec.items():
                print(f"    {k}: {v}")
        print(f"\n  ... 共 {len(all_parsed)} 条，不会实际写入")
        return 0

    # Step 2: 获取 token & 读取已有记录去重
    print("\n=== 连接飞书 ===")
    try:
        tok = feishu_get_token()
    except Exception as e:
        print(f"[错误] 飞书认证失败: {e}")
        return 1
    print(f"  Token OK")

    print("\n=== 读取已有记录（用于去重）===")
    existing_keys = bitable_list_existing(tok, args.app_token, args.table_id)

    # Step 3: 构建待写入记录
    print("\n=== 构建写入记录 ===")
    to_write, skipped = build_bitable_records(all_parsed, existing_keys)
    print(f"  新记录: {len(to_write)} 条")
    print(f"  已存在跳过: {skipped} 条")

    if not to_write:
        print("\n[完成] 所有记录均已存在，无需写入")
        return 0

    # Step 4: 预览前 5 条将写入的记录
    print("\n=== 将写入的前 5 条记录预览 ===")
    for i, rec in enumerate(to_write[:5]):
        fields = rec["fields"]
        print(f"\n  记录 {i + 1}:")
        for k, v in fields.items():
            val_str = str(v)
            if k == FIELD_DATE:
                try:
                    dt = datetime.datetime.fromtimestamp(int(v) / 1000)
                    val_str = dt.strftime("%Y-%m-%d")
                except Exception:
                    pass
            print(f"    {k}: {val_str}")

    # Step 5: 写入 Bitable
    print(f"\n=== 写入 Bitable ({len(to_write)} 条) ===")
    result = bitable_batch_create(tok, args.app_token, args.table_id, to_write)

    if result["errors"]:
        print(f"\n[警告] 写入出现错误:")
        for err in result["errors"]:
            print(f"  {err}")

    print(f"\n[完成] 成功 {result['ok']} 条 / 失败 {result['fail']} 条")
    return 0 if result["fail"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())