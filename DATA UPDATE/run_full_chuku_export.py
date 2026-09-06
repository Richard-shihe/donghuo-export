#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一键 runner：懂火系统 → 全量出库记录 → CSV → 飞书文件夹 + 本地备份
目标文件夹 token: Uuu0feVP0lbzEydQaLQcElJ9nVg
本地备份目录：本 runner 所在目录（即 DATA UPDATE/）

使用：
  1) 已通过 GitHub Actions Secrets 注入 DH_USERNAME / DH_PASSWORD 时：
        直接触发 .github/workflows/export.yml，参数：
          export_days  = 0
          folder_token = Uuu0feVP0lbzEydQaLQcElJ9nVg
  2) 本地命令行（需要手动把 DH_USERNAME/DH_PASSWORD 设成环境变量，或追加到 .env）：
        python "DATA UPDATE/run_full_chuku_export.py"
     本 runner 会自动加载仓库根目录的 .env，把 EXPORT_DAYS 固定为 0、FEISHU_FOLDER_TOKEN
     固定为目标文件夹，并在脚本跑完后把 CSV 也落一份到 DATA UPDATE/。
"""

import os
import sys
import shutil
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
WORK_DIR = Path(__file__).resolve().parent  # DATA UPDATE/
TARGET_FOLDER_TOKEN = "Uuu0feVP0lbzEydQaLQcElJ9nVg"

# 1) 先尝试从 .env 装载，有则用（override=False，已设的系统 env 优先级更高）
_dotenv = REPO_ROOT / ".env"
if _dotenv.exists():
    try:
        from dotenv import load_dotenv  # type: ignore
        load_dotenv(_dotenv, override=False)
        print(f"[runner] 已加载 {_dotenv}")
    except ImportError:
        print("[runner] 未安装 python-dotenv，跳过 .env 加载（依赖系统环境变量）")

# 2) 强制覆盖与本任务绑定的参数
os.environ["EXPORT_DAYS"] = "0"                    # 全量
os.environ["FEISHU_FOLDER_TOKEN"] = TARGET_FOLDER_TOKEN
os.environ["DELIVERY_MODE"] = os.environ.get("DELIVERY_MODE") or "feishu"

# 3) 用仓库根目录作为工作目录，运行 export_chuku.main()
EXPORT_DIR = REPO_ROOT / "懂火出库导出"          # export_chuku.py 所在目录
sys.path.insert(0, str(EXPORT_DIR))
os.chdir(REPO_ROOT)

import export_chuku  # noqa: E402

# 记录执行前 DATA UPDATE 下已有的 chuku_*.csv，便于事后找出本次生成的文件
before = {p.name for p in WORK_DIR.glob("chuku_*.csv")}

exit_code = export_chuku.main()

# 4) 若脚本在「懂火出库导出/」生成了 chuku_*.csv（Delivery 失败或飞书参数不足的本地兜底），
#    则把新出现的 CSV 拷贝到 DATA UPDATE/ 以便后续清洗。
#    脚本的 CSV 输出目录 = 脚本所在目录（__file__ 相对），故扫描 EXPORT_DIR。
for csv_p in sorted(EXPORT_DIR.glob("chuku_*.csv")):
    target = WORK_DIR / csv_p.name
    if target.exists() and target.stat().st_size == csv_p.stat().st_size:
        continue
    try:
        shutil.copy2(csv_p, target)
        print(f"[runner] 本地备份 → {target}")
    except Exception as exc:
        print(f"[runner][warn] 拷贝 {csv_p.name} 失败: {exc}")

# 再次扫描，汇报最终产物
after = sorted(WORK_DIR.glob("chuku_*.csv"))
new = [p for p in after if p.name not in before]
if new:
    print(f"[runner] 本轮新增 CSV 共 {len(new)} 个（均保存在 DATA UPDATE/）：")
    for p in new:
        size_kb = p.stat().st_size / 1024
        print(f"  - {p.name}  ({size_kb:.1f} KB)")
else:
    print("[runner] 注意：本地 DATA UPDATE/ 目录内没发现新的 chuku_*.csv。"
          "若 DELIVERY_MODE=feishu 且上传成功，说明文件只在飞书云盘里，"
          "可在飞书文件夹下载一份到本地后再开始清洗。")

raise SystemExit(exit_code)
