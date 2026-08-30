#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
阶段2 · A17 识别分类：原件文件夹 → 新识别 / 未识别 → 已处理
================================================================

流程：
  S1  列「待处理」文件夹全部文件（同名去重处理，同名重复件最后一起移送）
  S2  逐个下载 → 本地 A17 识别（process_one）
      - 识别成功 → 重命名 PDF（可多份）上传「新识别」文件夹（同名跳过）
      - 识别失败 → 原件上传「未识别」文件夹（同名跳过）
  S3  把「待处理」文件夹里的全部文件（含同名重复）移送「已处理」文件夹
  S4  机器人汇报（cert_common.notify）

文件夹（默认值，可参数/环境变量覆盖）：
  待处理   YIrbf0NzlloFNKdYAjvcL6gXnhe   ← 阶段1 产出
  新识别   VeXHfjiZqll255dgSOYcknbdnnf   ← 阶段3 输入
  未识别   Ic0Wf1PeelamrJd15GkcPo7Jnlb
  已处理 HfYCfYMZnlhVpFdKI82cTHQ5npb

用法：
  python stage2_certs_classify.py               # 实际执行
  python stage2_certs_classify.py --dry-run     # 只列文件+识别预演，不上传不移动
  python stage2_certs_classify.py --limit 5     # 只处理前 5 个文件
  python stage2_certs_classify.py --no-notify   # 跳过机器人汇报

环境变量：
  FEISHU_APP_ID / FEISHU_APP_SECRET  飞书自建应用
  CERT_RAW_FOLDER_TOKEN / CERT_OK_FOLDER_TOKEN /
  CERT_BAD_FOLDER_TOKEN / CERT_ARCHIVE_FOLDER_TOKEN
  CERT_NOTIFY_UNION_IDS              汇报人员 union_id（逗号分隔；回退 FEISHU_UNION_IDS）
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import time
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
    download_file, upload_to_folder, move_file, folder_url,
)
from A17 import (  # noqa: E402
    process_one as a17_process_one,
    extract_pdf_text, extract_contract_no, extract_grade, extract_entries,
)

# ============================================================
# 默认文件夹
# ============================================================
DEFAULT_RAW_FOLDER = "YIrbf0NzlloFNKdYAjvcL6gXnhe"      # 待处理（阶段1产出）
DEFAULT_OK_FOLDER = "VeXHfjiZqll255dgSOYcknbdnnf"       # 新识别（阶段3输入）
DEFAULT_BAD_FOLDER = "Ic0Wf1PeelamrJd15GkcPo7Jnlb"      # 未识别
DEFAULT_ARCHIVE_FOLDER = "HfYCfYMZnlhVpFdKI82cTHQ5npb"  # 已处理

WORK_DIR = _CERT_DIR / "_cert_work"
WORK_DIR.mkdir(parents=True, exist_ok=True)   # CI 全新 checkout 无此目录，必须先建
RAW_DIR = WORK_DIR / "raw"
A17_OUT_DIR = WORK_DIR / "a17_output"
A17_FRESH_DIR = WORK_DIR / "a17_fresh"
for d in (RAW_DIR, A17_OUT_DIR, A17_FRESH_DIR):
    d.mkdir(parents=True, exist_ok=True)

UPLOAD_SLEEP = 0.1
MOVE_SLEEP = 0.1


# ============================================================
# A17 输出定位
# ============================================================
def _rebuild_a17_names(src: Path) -> list[Path]:
    """A17 同名跳过（status=skipped）时，重现它会生成的 new_name，定位已存在文件。"""
    try:
        text = extract_pdf_text(str(src))
        contract_no, _ = extract_contract_no(str(src), src.name)
        grade = extract_grade(text)
        entries = extract_entries(text)
        if not contract_no or not grade or not entries:
            return []
        outs = []
        for entry in entries:
            weight_ton = entry["weight_kg"] / 1000
            new_name = (f"{contract_no} {entry['coil_no']} {grade} "
                        f"{entry['thick']} {entry['width']} {weight_ton:.3f}吨.pdf")
            new_name = re.sub(r'[\\/:*?"<>|]', "", new_name)
            for base in (A17_FRESH_DIR, A17_OUT_DIR):
                cand = base / new_name
                if cand.exists():
                    outs.append(cand)
                    break
        return outs
    except Exception:
        return []


def resolve_a17_outputs(src: Path) -> tuple[str, list[Path]]:
    """跑 A17 识别，返回 (status, 输出PDF路径列表)。status: ok / failed"""
    status = ""
    gen_names: list[str] = []
    try:
        res = a17_process_one(str(src), str(A17_OUT_DIR), str(A17_FRESH_DIR))
        if isinstance(res, tuple) and len(res) >= 4:
            status = str(res[0])
            gt = res[3]
            if isinstance(gt, list):
                gen_names = [n for n in gt if isinstance(n, str) and n.lower().endswith(".pdf")]
        else:
            status = "failed"
    except Exception as e:
        log(f"    A17 异常: {e}")
        status = "failed"

    outs: list[Path] = []
    for nm in gen_names:
        for base in (A17_FRESH_DIR, A17_OUT_DIR):
            cand = base / nm
            if cand.exists():
                outs.append(cand)
                break
    if not outs and status == "skipped":
        # 同名已存在（之前跑过），重新构造名字定位
        outs = _rebuild_a17_names(src)
    if outs:
        return "ok", outs
    return "failed", []


# ============================================================
# Main
# ============================================================
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="阶段2 · A17 识别分类：原件 → 新识别/未识别 → 已处理",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--raw-folder", default=env("CERT_RAW_FOLDER_TOKEN", DEFAULT_RAW_FOLDER),
                   help=f"待处理文件夹（默认 {DEFAULT_RAW_FOLDER}）")
    p.add_argument("--ok-folder", default=env("CERT_OK_FOLDER_TOKEN", DEFAULT_OK_FOLDER),
                   help=f"新识别文件夹（默认 {DEFAULT_OK_FOLDER}）")
    p.add_argument("--bad-folder", default=env("CERT_BAD_FOLDER_TOKEN", DEFAULT_BAD_FOLDER),
                   help=f"未识别文件夹（默认 {DEFAULT_BAD_FOLDER}）")
    p.add_argument("--archive-folder",
                   default=env("CERT_ARCHIVE_FOLDER_TOKEN", DEFAULT_ARCHIVE_FOLDER),
                   help=f"已处理文件夹（默认 {DEFAULT_ARCHIVE_FOLDER}）")
    p.add_argument("--limit", type=int, default=0,
                   help="最多处理多少个文件（0=全部）")
    p.add_argument("--dry-run", action="store_true",
                   help="只列文件+识别预演，不上传不移动不汇报")
    p.add_argument("--no-notify", action="store_true", help="跳过机器人汇报")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    t0 = time.time()
    log_file = WORK_DIR / f'stage2_{time.strftime("%Y%m%d_%H%M%S")}.log'

    log("========== 【阶段2】A17 识别分类 ==========", log_file)
    log(f"  待处理:   {args.raw_folder}", log_file)
    log(f"  新识别:   {args.ok_folder}", log_file)
    log(f"  未识别:   {args.bad_folder}", log_file)
    log(f"  已处理: {args.archive_folder}", log_file)
    log(f"  dry_run:  {args.dry_run}", log_file)

    # 1. token
    try:
        token = tenant_access_token()
    except Exception as e:
        log(f"❌ 取 tenant_access_token 失败: {e}", log_file)
        return 1

    # 2. 列待处理文件夹
    log("\n[S1] 列出待处理文件夹…", log_file)
    try:
        files = list_folder_files(token, args.raw_folder)
    except Exception as e:
        log(f"❌ 列文件失败: {e}", log_file)
        return 1
    log(f"  共 {len(files)} 个文件", log_file)
    if not files:
        log("  待处理文件夹为空，无事可做", log_file)
        return 0

    # 同名去重（同名重复件不重复识别，但最后一起移送）
    unique: dict[str, dict] = {}
    for f in files:
        name = f.get("name") or ""
        ftok = f.get("token") or ""
        if not name or not ftok:
            continue
        if name not in unique:
            unique[name] = f
    dup_count = len(files) - len(unique)
    log(f"  去重后待处理 {len(unique)} 个（文件夹内同名重复 {dup_count} 个）", log_file)

    if args.limit and args.limit > 0:
        unique_list = list(unique.values())[:args.limit]
        log(f"  --limit {args.limit}：只处理前 {len(unique_list)} 个", log_file)
    else:
        unique_list = list(unique.values())

    # 目标文件夹已有文件名（上传幂等）
    def _names(folder: str) -> set[str]:
        try:
            return {f.get("name") or "" for f in list_folder_files(token, folder)}
        except Exception as e:
            log(f"  ⚠️ 列目标文件夹 {folder} 失败: {e}", log_file)
            return set()

    ok_names = _names(args.ok_folder)
    bad_names = _names(args.bad_folder)

    # 3. 逐个识别 + 上传
    log("\n[S2] 下载 → A17 识别 → 分类上传", log_file)
    summary = {
        "ok_files": 0,       # 识别成功（上传到新识别的重命名 PDF 数）
        "ok_src": 0,         # 识别成功的源文件数
        "bad_src": 0,        # 识别失败（上传到未识别的原件数）
        "skip_ok_dup": 0,    # 新识别文件夹同名跳过
        "skip_bad_dup": 0,   # 未识别文件夹同名跳过
        "failed": 0,         # 下载/上传失败
        "moved": 0,
        "move_failed": 0,
    }
    items: list[dict] = []

    for i, f in enumerate(unique_list, 1):
        name = f.get("name") or ""
        ftok = f.get("token") or ""
        log(f"\n  [{i}/{len(unique_list)}] {name}", log_file)
        item = {"name": name, "file_token": ftok, "recognized": None}

        if args.dry_run:
            # 预演：下载到本地跑 A17，但不上传
            try:
                raw = download_file(token, ftok)
                local = RAW_DIR / name
                local.write_bytes(raw)
                status, outs = resolve_a17_outputs(local)
                item["recognized"] = status == "ok"
                item["outputs"] = [p.name for p in outs]
                if status == "ok":
                    summary["ok_src"] += 1
                    summary["ok_files"] += len(outs)
                    for p in outs:
                        log(f"    [dry-run] ✅ 识别成功 → {p.name}")
                else:
                    summary["bad_src"] += 1
                    log(f"    [dry-run] ⚠️ 未识别 → 将上传原件到未识别文件夹")
            except Exception as e:
                summary["failed"] += 1
                log(f"    ❌ [dry-run] 下载/识别失败: {e}", log_file)
            items.append(item)
            continue

        # 下载
        try:
            raw = download_file(token, ftok)
        except Exception as e:
            summary["failed"] += 1
            log(f"    ❌ 下载失败: {e}", log_file)
            items.append(item)
            continue
        local = RAW_DIR / name
        local.write_bytes(raw)

        # A17 识别
        status, outs = resolve_a17_outputs(local)
        item["recognized"] = status == "ok"
        item["outputs"] = [p.name for p in outs]

        if status == "ok":
            summary["ok_src"] += 1
            for p in outs:
                if p.name in ok_names:
                    summary["skip_ok_dup"] += 1
                    log(f"    ⏭ 新识别文件夹已有同名，跳过: {p.name}", log_file)
                    continue
                try:
                    upload_to_folder(token, args.ok_folder, p.name, p.read_bytes())
                    ok_names.add(p.name)
                    summary["ok_files"] += 1
                    log(f"    ✅ 上传新识别: {p.name}", log_file)
                except Exception as e:
                    summary["failed"] += 1
                    log(f"    ❌ 上传新识别失败 {p.name}: {e}", log_file)
                time.sleep(UPLOAD_SLEEP)
        else:
            summary["bad_src"] += 1
            if name in bad_names:
                summary["skip_bad_dup"] += 1
                log(f"    ⏭ 未识别文件夹已有同名，跳过: {name}", log_file)
            else:
                try:
                    upload_to_folder(token, args.bad_folder, name, raw)
                    bad_names.add(name)
                    log(f"    ⚠️ 未识别，原件上传未识别文件夹: {name}", log_file)
                except Exception as e:
                    summary["failed"] += 1
                    log(f"    ❌ 上传未识别失败 {name}: {e}", log_file)
            time.sleep(UPLOAD_SLEEP)
        items.append(item)

    # 4. 原件移送归档（含同名重复、含处理失败的）
    #    只移"本次已处理"的文件；--limit 截断时，未处理文件留在待处理文件夹
    processed_names = {f.get("name") or "" for f in unique_list}
    log(f"\n[S3] 原件移送归档 → {args.archive_folder}", log_file)
    if args.dry_run:
        will_move = sum(1 for f in files if (f.get("name") or "") in processed_names)
        log(f"  [dry-run] 将移送 {will_move} 个文件", log_file)
        summary["moved"] = 0
    else:
        for i, f in enumerate(files, 1):
            name = f.get("name") or ""
            ftok = f.get("token") or ""
            if not name or not ftok or name not in processed_names:
                continue
            try:
                move_file(token, ftok, args.archive_folder)
                summary["moved"] += 1
                log(f"  [{i}/{len(files)}] ✅ 移送 {name}", log_file)
            except Exception as e:
                summary["move_failed"] += 1
                log(f"  [{i}/{len(files)}] ❌ 移送失败 {name}: {e}", log_file)
            time.sleep(MOVE_SLEEP)
        log(f"  移送完成：成功 {summary['moved']}，失败 {summary['move_failed']}", log_file)

    # 5. 汇报
    elapsed = int(time.time() - t0)
    mins, secs = divmod(elapsed, 60)
    msg = (f"【质保书·②A17识别分类】完成 ✅\n"
           f"待处理 {len(unique)} 个（同名重复 {dup_count}）\n"
           f"识别成功: {summary['ok_src']} 份 → 上传 {summary['ok_files']} 个 PDF"
           + (f"（同名跳过 {summary['skip_ok_dup']}）" if summary["skip_ok_dup"] else "") + "\n"
           f"未识别: {summary['bad_src']} 份 → 上传原件"
           + (f"（同名跳过 {summary['skip_bad_dup']}）" if summary["skip_bad_dup"] else "") + "\n"
           f"原件移送归档: {summary['moved']}"
           + (f"（失败 {summary['move_failed']}）" if summary["move_failed"] else "") + "\n"
           f"新识别: {folder_url(args.ok_folder)}\n"
           f"未识别: {folder_url(args.bad_folder)}\n"
           f"归档: {folder_url(args.archive_folder)}\n"
           f"耗时: {mins}分{secs}秒")
    print(msg)
    if not args.dry_run and not args.no_notify:
        notify(msg)

    result_path = WORK_DIR / f'stage2_result_{time.strftime("%Y%m%d_%H%M%S")}.json'
    with open(result_path, "w", encoding="utf-8") as f:
        json.dump({
            "args": vars(args),
            "elapsed_sec": elapsed,
            "unique": len(unique),
            "dup_in_folder": dup_count,
            "summary": summary,
            "items": items,
        }, f, ensure_ascii=False, indent=2)
    log(f"📄 结果快照: {result_path}", log_file)
    return 0


if __name__ == "__main__":
    sys.exit(main())
