#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
阶段3 · 梳理「已识别」文件夹质保书 → 写 Bitable「质保书」附件字段
==================================================================

以「资源号」为唯一号，把已识别文件夹里的质保书逐条挂到多维表对应记录：

  源文件夹  VeXHfjiZqll255dgSOYcknbdnnf（阶段2 识别成功产出）
  目标表    Bitable app_token=Tz0XbQVzkaZuJasBwb8cRjkfnoe
            table_id=tblvugnoJPS8GrpX
  目标字段  质保书（附件类型）

去重（两层，保证幂等可重跑）：
  1. 文件夹内同名文件只处理一次
  2. 记录已存在同名附件 → 跳过（不重复挂）

流程：
  S1  列源文件夹全部文件（分页）
  S2  拉目标表全量记录，构建 {资源号 → [record_id]} / {record_id → 已有附件}
  S3  文件名提取资源号（已知资源号最长前缀匹配 → 正则回退），按资源号分组
  S4  逐文件：下载 → 上传 bitable 附件 → 合并进「质保书」字段（追加，不覆盖）
  S5  把写入成功的文件移动到归档文件夹（上传成功 + 跳过重复；--no-move 关闭）
  S6  机器人汇报（cert_common.notify）

用法：
  python stage3_certs_organize.py --dry-run            # 预演
  python stage3_certs_organize.py                      # 实际执行
  python stage3_certs_organize.py --resource L6ER000618  # 只处理指定资源号
  python stage3_certs_organize.py --no-move          # 不做归档移动（默认会移）

环境变量：
  FEISHU_APP_ID / FEISHU_APP_SECRET
  CERT_OK_FOLDER_TOKEN        源文件夹（默认 VeXHfjiZqll255dgSOYcknbdnnf）
  BITABLE_APP_TOKEN / BITABLE_TABLE_ID / BITABLE_FIELD_CERT
  ARCHIVE_FOLDER_TOKEN        （可选）覆盖归档文件夹（默认 H24ifEj4alUBF6dzeioctPqinsf）
  CERT_NOTIFY_UNION_IDS       汇报人员 union_id（逗号分隔；回退 FEISHU_UNION_IDS）
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from collections import defaultdict
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

load_dotenv()  # 兼容从仓库根目录运行时读取 .env

_CERT_DIR = Path(__file__).resolve().parent
for _p in (str(_CERT_DIR), str(_CERT_DIR.parent)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from cert_common import (  # noqa: E402
    env, log, notify, tenant_access_token, list_folder_files,
    list_bitable_records, download_file, upload_bitable_media,
    update_record_field, move_file, folder_url,
)

# ============================================================
# 常量
# ============================================================
DEFAULT_SOURCE_FOLDER = "VeXHfjiZqll255dgSOYcknbdnnf"
DEFAULT_APP_TOKEN = "Tz0XbQVzkaZuJasBwb8cRjkfnoe"
DEFAULT_TABLE_ID = "tblvugnoJPS8GrpX"
DEFAULT_ARCHIVE_FOLDER = "H24ifEj4alUBF6dzeioctPqinsf"  # 梳理成功后的归档文件夹
FIELD_ZIYUANHAO = "资源号"
FIELD_CERT = env("BITABLE_FIELD_CERT", "质保书")

UPLOAD_SLEEP = 0.25  # 素材上传 QPS 5，保守间隔

WORK_DIR = _CERT_DIR / "_cert_work"
WORK_DIR.mkdir(parents=True, exist_ok=True)   # CI 全新 checkout 无此目录，必须先建

# 已知资源号模式：X6E0016009 / L4E0006973 / L6ER000579
ZY_REGEX = re.compile(r'^([A-Z]\d+[A-Z]{1,3}\d{3,})')


# ============================================================
# Bitable 记录字段解析
# ============================================================
def _extract_ziyuanhao(fields: dict) -> list[str]:
    """「资源号」Lookup 字段值格式多样：['L4E0006973'] / [{text:...}] / str。"""
    val = fields.get(FIELD_ZIYUANHAO)
    if val is None:
        return []
    if isinstance(val, str):
        v = val.strip()
        return [v] if v else []
    if isinstance(val, list):
        out = []
        for x in val:
            if isinstance(x, str):
                if x.strip():
                    out.append(x.strip())
            elif isinstance(x, dict):
                t = str(x.get("text") or x.get("value") or "").strip()
                if t:
                    out.append(t)
            else:
                t = str(x).strip()
                if t:
                    out.append(t)
        return out
    return [str(val).strip()]


def _extract_cert_names(fields: dict) -> set[str]:
    """「质保书」附件字段已存在的文件名集合。"""
    val = fields.get(FIELD_CERT)
    out: set[str] = set()
    if val is None:
        return out
    if isinstance(val, list):
        for x in val:
            if isinstance(x, dict):
                n = str(x.get("name") or "").strip()
                if n:
                    out.add(n)
            elif isinstance(x, str) and x.strip():
                out.add(x.strip())
    elif isinstance(val, dict):
        n = str(val.get("name") or "").strip()
        if n:
            out.add(n)
    elif isinstance(val, str) and val.strip():
        out.add(val.strip())
    return out


def _extract_cert_tokens(fields: dict) -> list[dict]:
    """现有质保书附件 token 结构（更新时传 file_token 即可保留原附件）。"""
    val = fields.get(FIELD_CERT)
    out: list[dict] = []
    if val is None:
        return out
    if isinstance(val, list):
        for x in val:
            if isinstance(x, dict) and x.get("file_token"):
                out.append({"file_token": x["file_token"]})
            elif isinstance(x, str) and x:
                out.append({"file_token": x})
    elif isinstance(val, dict) and val.get("file_token"):
        out.append({"file_token": val["file_token"]})
    elif isinstance(val, str) and val:
        out.append({"file_token": val})
    return out


def build_index(records: list[dict], log_file: Optional[Path] = None):
    """构建 {资源号 → [record_id]} / {record_id → 已有附件名} / {record_id → 附件token列表}。"""
    zy_to_rids: dict[str, list[str]] = defaultdict(list)
    rid_to_existing_names: dict[str, set[str]] = {}
    rid_to_existing_tokens: dict[str, list[dict]] = {}
    missing_zy = 0
    for rec in records:
        rid = rec.get("record_id") or ""
        if not rid:
            continue
        fields = rec.get("fields") or {}
        zys = _extract_ziyuanhao(fields)
        if not zys:
            missing_zy += 1
        for zy in zys:
            zy_to_rids[zy].append(rid)
        rid_to_existing_names[rid] = _extract_cert_names(fields)
        rid_to_existing_tokens[rid] = _extract_cert_tokens(fields)
    log(f"  索引完成：记录 {len(records)} 条，唯一资源号 {len(zy_to_rids)} 个，"
        f"无资源号 {missing_zy} 条", log_file)
    return dict(zy_to_rids), rid_to_existing_names, rid_to_existing_tokens


def extract_resource_number(filename: str, known_zys: set[str]) -> str:
    """文件名 → 资源号。已知资源号最长前缀匹配 → 正则 → 首段分隔符。"""
    name = os.path.splitext(filename)[0].strip()
    if not name:
        return ""
    candidates = [zy for zy in known_zys if name.startswith(zy)]
    if candidates:
        return max(candidates, key=len)
    m = ZY_REGEX.match(name)
    if m:
        return m.group(1)
    for sep in ['_', '-', ' ', '（', '(', '，', ',']:
        idx = name.find(sep)
        if idx > 0:
            return name[:idx]
    return name


# ============================================================
# Main
# ============================================================
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="阶段3 · 梳理已识别质保书 → Bitable「质保书」附件字段",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--source-folder", default=env("CERT_OK_FOLDER_TOKEN", DEFAULT_SOURCE_FOLDER),
                   help=f"源云盘文件夹（已识别，默认 {DEFAULT_SOURCE_FOLDER}）")
    p.add_argument("--app-token", default=env("BITABLE_APP_TOKEN", DEFAULT_APP_TOKEN),
                   help=f"Bitable app_token（默认 {DEFAULT_APP_TOKEN}）")
    p.add_argument("--table-id", default=env("BITABLE_TABLE_ID", DEFAULT_TABLE_ID),
                   help=f"Bitable table_id（默认 {DEFAULT_TABLE_ID}）")
    p.add_argument("--move-to", default=env("ARCHIVE_FOLDER_TOKEN", DEFAULT_ARCHIVE_FOLDER),
                   help=f"梳理成功后把文件移动到归档文件夹（默认 {DEFAULT_ARCHIVE_FOLDER}）")
    p.add_argument("--no-move", action="store_true",
                   help="禁用归档移动（处理完的文件留在已识别文件夹）")
    p.add_argument("--only-matched", action="store_true",
                   help="只处理能匹配到 Bitable 资源号的文件")
    p.add_argument("--resource", default="",
                   help="只处理指定资源号（逗号分隔多个）")
    p.add_argument("--dry-run", action="store_true",
                   help="预演，不下载/上传/写表/移动")
    p.add_argument("--no-notify", action="store_true", help="跳过机器人汇报")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    t0 = time.time()
    log_file = WORK_DIR / f'stage3_{time.strftime("%Y%m%d_%H%M%S")}.log'

    log("========== 【阶段3】质保书梳理 → Bitable 入库 ==========", log_file)
    log(f"  源文件夹: {args.source_folder}", log_file)
    log(f"  目标表:   {args.app_token} / {args.table_id}", log_file)
    log(f"  目标字段: {FIELD_CERT}", log_file)
    move_enabled = bool(args.move_to) and not args.no_move
    if move_enabled:
        log(f"  归档文件夹: {args.move_to}（写入成功后自动移动）", log_file)
    else:
        log("  归档移动: 已禁用（--no-move）", log_file)
    log(f"  模式:     {'dry-run（预演）' if args.dry_run else '实际执行'}", log_file)

    # 1. token
    try:
        token = tenant_access_token()
    except Exception as e:
        log(f"❌ 取 tenant_access_token 失败: {e}", log_file)
        return 1

    # 2. 列源文件夹
    log("\n[S1] 列出源文件夹文件…", log_file)
    try:
        files = list_folder_files(token, args.source_folder)
    except Exception as e:
        log(f"❌ 列文件失败: {e}", log_file)
        return 1
    log(f"  共 {len(files)} 个文件", log_file)
    if not files:
        log("  源文件夹为空，无事可做", log_file)
        if not args.no_notify:
            notify("【质保书·③入库Bitable】源文件夹为空，无新增。\n"
                   f"源文件夹: {folder_url(args.source_folder)}")
        return 0

    # 3. 拉目标表全量记录 + 索引
    log("\n[S2] 拉取 Bitable 全部记录…", log_file)
    try:
        records = list_bitable_records(token, args.app_token, args.table_id)
    except Exception as e:
        log(f"❌ 拉记录失败: {e}", log_file)
        return 1
    zy_to_rids, rid_to_existing_names, rid_to_existing_tokens = \
        build_index(records, log_file)

    # 4. 按资源号分组
    log("\n[S3] 按资源号分组…", log_file)
    known_zys = set(zy_to_rids.keys())
    zy_to_files: dict[str, list[dict]] = defaultdict(list)
    unmatched: list[dict] = []
    seen_filenames: set[str] = set()
    dup_in_folder = 0
    for f in files:
        fname = f.get("name") or ""
        if not fname:
            continue
        if fname in seen_filenames:      # 文件夹内同名去重
            dup_in_folder += 1
            continue
        seen_filenames.add(fname)
        zy = extract_resource_number(fname, known_zys)
        if zy and zy in zy_to_rids:
            zy_to_files[zy].append(f)
        else:
            unmatched.append(f)
            if zy:
                log(f"  ⚠️ {fname} → 资源号 {zy} 在表中无记录（跳过）", log_file)
            else:
                log(f"  ⚠️ {fname} → 无法提取资源号（跳过）", log_file)

    log(f"  分组完成：{len(zy_to_files)} 个资源号有文件，"
        f"{len(unmatched)} 个文件无法匹配，"
        f"{dup_in_folder} 个文件夹内同名重复", log_file)

    if args.resource:
        wanted = {x.strip() for x in args.resource.split(',') if x.strip()}
        zy_to_files = {zy: fs for zy, fs in zy_to_files.items() if zy in wanted}
        log(f"  按 --resource 过滤后剩余 {len(zy_to_files)} 个资源号", log_file)

    if args.only_matched and not zy_to_files:
        log("  无可匹配文件，退出", log_file)
        return 0

    # 5. 逐资源号处理
    log("\n[S4] 下载 → 上传附件 → 更新「质保书」字段", log_file)
    total_files = sum(len(fs) for fs in zy_to_files.values())
    log(f"  待处理文件 {total_files} 个", log_file)

    summary = {
        "zy_ok": 0, "zy_fail": 0,
        "files_uploaded": 0, "files_skipped_dup": 0, "files_failed": 0,
        "records_updated": 0,
        "files_moved": 0, "files_move_failed": 0,
    }
    organized_names: set[str] = set()   # 已梳理文件名（上传成功或同名跳过）

    for zy_idx, (zy, fs_list) in enumerate(sorted(zy_to_files.items()), 1):
        rids = zy_to_rids.get(zy) or []
        zy_files_failed = 0   # 该资源号本次处理的失败次数
        log(f"\n  [{zy_idx}/{len(zy_to_files)}] 资源号 {zy} → "
            f"{len(fs_list)} 个文件，{len(rids)} 条记录", log_file)
        for f in fs_list:
            fname = f.get("name") or ""
            ftok = f.get("token") or ""
            log(f"    {fname}", log_file)
            for rid in rids:
                existing_names = rid_to_existing_names.get(rid, set())
                if fname in existing_names:
                    summary["files_skipped_dup"] += 1
                    organized_names.add(fname)
                    log(f"      ⏭ record {rid}: 已存在同名附件，跳过", log_file)
                    continue
                if args.dry_run:
                    organized_names.add(fname)
                    log(f"      [dry-run] 将下载 + 上传 + 追加到 record {rid}", log_file)
                    continue
                try:
                    raw = download_file(token, ftok)
                except Exception as e:
                    summary["files_failed"] += 1
                    zy_files_failed += 1
                    log(f"      ❌ record {rid}: 下载失败: {e}", log_file)
                    continue
                try:
                    new_ft = upload_bitable_media(token, args.app_token, fname, raw)
                    time.sleep(UPLOAD_SLEEP)
                except Exception as e:
                    summary["files_failed"] += 1
                    zy_files_failed += 1
                    log(f"      ❌ record {rid}: 上传素材失败: {e}", log_file)
                    continue
                if not new_ft:
                    summary["files_failed"] += 1
                    zy_files_failed += 1
                    log(f"      ❌ record {rid}: 上传素材返回空 file_token", log_file)
                    continue
                # 合并新附件（追加，不覆盖）
                existing_tokens = list(rid_to_existing_tokens.get(rid, []))
                existing_tokens.append({"file_token": new_ft})
                try:
                    update_record_field(token, args.app_token, args.table_id,
                                        rid, FIELD_CERT, existing_tokens)
                    summary["files_uploaded"] += 1
                    summary["records_updated"] += 1
                    organized_names.add(fname)
                    rid_to_existing_names.setdefault(rid, set()).add(fname)
                    rid_to_existing_tokens[rid] = existing_tokens
                    log(f"      ✅ record {rid}: 追加成功 file_token={new_ft}", log_file)
                except Exception as e:
                    summary["files_failed"] += 1
                    zy_files_failed += 1
                    log(f"      ❌ record {rid}: 更新记录失败: {e}", log_file)
                    continue

        if args.dry_run:
            summary["zy_ok"] += 1
        elif zy_files_failed == 0:
            summary["zy_ok"] += 1
        else:
            summary["zy_fail"] += 1

    # 6. 移动"写入成功"的文件到归档文件夹（上传成功 + 跳过重复；失败的留在已识别文件夹）
    if move_enabled:
        log(f"\n[S5] 归档移动 → {args.move_to}", log_file)
        log(f"  已梳理文件名 {len(organized_names)} 个（含同名重复）", log_file)
        moved_idx = 0
        for f in files:
            fname = f.get("name") or ""
            ftok = f.get("token") or ""
            if not fname or not ftok or fname not in organized_names:
                continue
            moved_idx += 1
            if args.dry_run:
                log(f"  [{moved_idx}] [dry-run] 将移动 {fname}", log_file)
                continue
            try:
                move_file(token, ftok, args.move_to)
                summary["files_moved"] += 1
                log(f"  [{moved_idx}] ✅ 移动 {fname}", log_file)
                time.sleep(0.1)
            except Exception as e:
                summary["files_move_failed"] += 1
                log(f"  [{moved_idx}] ❌ 移动失败 {fname}: {e}", log_file)

    # 7. 汇报
    elapsed = int(time.time() - t0)
    mins, secs = divmod(elapsed, 60)
    msg = (f"【质保书·③入库Bitable】完成 {'✅' if not summary['zy_fail'] else '⚠️'}\n"
           f"源文件 {len(files)} 个（同名重复 {dup_in_folder}，无法匹配 {len(unmatched)}）\n"
           f"资源号: 匹配 {len(zy_to_files)} 个（成功 {summary['zy_ok']} / 失败 {summary['zy_fail']}）\n"
           f"附件写入: {summary['files_uploaded']}"
           f" | 跳过重复: {summary['files_skipped_dup']}"
           f" | 失败: {summary['files_failed']}\n"
           f"更新记录: {summary['records_updated']} 次\n"
           f"目标表: {args.app_token} / {args.table_id}\n"
           + (f"归档移动: {summary['files_moved']}（失败 {summary['files_move_failed']}）\n" if move_enabled else "")
           + f"耗时: {mins}分{secs}秒")
    print(msg)
    if not args.dry_run and not args.no_notify:
        notify(msg)

    result_path = WORK_DIR / f'stage3_result_{time.strftime("%Y%m%d_%H%M%S")}.json'
    with open(result_path, "w", encoding="utf-8") as f:
        json.dump({
            "args": vars(args),
            "elapsed_sec": elapsed,
            "files_in_folder": len(files),
            "dup_in_folder": dup_in_folder,
            "unmatched": len(unmatched),
            "zy_matched": len(zy_to_files),
            "summary": summary,
        }, f, ensure_ascii=False, indent=2)
    log(f"📄 结果快照: {result_path}", log_file)
    return 0


if __name__ == "__main__":
    sys.exit(main())
