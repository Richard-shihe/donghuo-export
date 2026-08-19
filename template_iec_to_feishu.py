#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IEC → 飞书云盘 交付模板
======================

封装「IEC 下载准发数据 → 上传飞书云盘 → 机器人通知」完整流程。
作为同类 IEC 任务的起点：复制本文件，按需修改 stage_download() 一段即可。

三段式 pipeline（main 顺序调用）:
  stage_download()   登录 IEC → 调 exportExcel → 下载 xlsx → 返回 (df, file_path)
  stage_upload()     换飞书 token → upload_all 到指定文件夹
  stage_notify()     飞书机器人 Webhook 通知（缺凭据自动跳过）

复制后通常只需改动:
  1. FILE_PREFIX          输出文件名前缀（默认"准发下载"）
  2. stage_download()      数据源/筛选条件/下载模式（download vs bundle）
  3. DEFAULT_FOLDER_TOKEN 目标飞书文件夹 token（或用环境变量 FEISHU_FOLDER_TOKEN）

环境变量:
  IBAO_USERNAME / IBAO_PASSWORD        IEC 账号
  FEISHU_APP_ID / FEISHU_APP_SECRET    飞书自建应用凭据
  FEISHU_FOLDER_TOKEN                  目标文件夹 token（fld 开头）
  FEISHU_WEBHOOK_URL / FEISHU_WEBHOOK_SECRET  机器人通知（可选，缺则跳过）
  EXPORT_START / EXPORT_END            交货期月份 YYYYMM（可选，默认近3月~下月）
  DRY_RUN=1                            预览模式（只下载不上传）

运行:
  python template_iec_to_feishu.py                  # 默认下载模式
  python template_iec_to_feishu.py --type bundle   # 捆包下载模式
  python template_iec_to_feishu.py --dry-run       # 预览
  python template_iec_to_feishu.py --start 202607 --end 202609
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import hmac
import json
import os
import sys
import time
import urllib.parse
from typing import Optional, Tuple

import requests

try:
    import pandas as pd
    _HAS_PANDAS = True
except ImportError:
    _HAS_PANDAS = False

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ibaosteel_client import IEC


# ============================================================
# 配置区（复制后按需修改）
# ============================================================
# 输出文件名前缀：准发下载_YYMMDD_HHMMSS.xlsx
FILE_PREFIX = "准发下载"

# 飞书云盘默认文件夹 token（留空则用环境变量 FEISHU_FOLDER_TOKEN）
DEFAULT_FOLDER_TOKEN = ""

# IEC 接口常量（一般无需改动）
_IECS_INDEX = "https://www.ibaosteel.com/iecs/index"
_QUASI_HAIR_PAGE = "https://www.ibaosteel.com/iecs/freight/quasiHair/quasiHair/initLoads"
_EXPORT_API = "https://www.ibaosteel.com/iecs/common/iec/exportExcel"
_DOWNLOAD_API = "https://www.ibaosteel.com/iecs/common/download"
_SETTLE_USER_NUM = "062122"

# 飞书 Open API
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}
SHANGHAI_OFFSET_HOURS = 8


# ============================================================
# 工具
# ============================================================
def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def now_shanghai_str() -> str:
    utc_now = datetime.datetime.utcnow()
    sh = utc_now + datetime.timedelta(hours=SHANGHAI_OFFSET_HOURS)
    return sh.strftime("%Y-%m-%d %H:%M:%S")


def _default_range() -> Tuple[str, str]:
    """默认交货期范围：当前月往前 3 个月到往后 1 个月"""
    today = datetime.date.today()
    y, m = today.year, today.month
    sm, sy = m - 3, y
    if sm <= 0:
        sm += 12
        sy -= 1
    em, ey = m + 1, y
    if em > 12:
        em -= 12
        ey += 1
    return f"{sy:04d}{sm:02d}", f"{ey:04d}{em:02d}"


def _auto_filename() -> str:
    """生成默认文件名：{FILE_PREFIX}_YYMMDD_HHMMSS.xlsx"""
    now = datetime.datetime.now()
    return f"{FILE_PREFIX}_{now:%y%m%d_%H%M%S}.xlsx"


def _parse_month(s: str) -> str:
    s = str(s).strip()
    if len(s) == 6 and s.isdigit():
        y, m = int(s[:4]), int(s[4:6])
        if 2000 <= y <= 2100 and 1 <= m <= 12:
            return s
    raise argparse.ArgumentTypeError(f"月份格式错误: {s!r}（应为 YYYYMM 如 202607）")


# ============================================================
# 阶段1：IEC 下载（可替换的钩子）
# ============================================================
def stage_download(
    start: str,
    end: str,
    out_path: str,
    download_type: str = "download",
    dry_run: bool = False,
) -> Tuple["pd.DataFrame", str]:
    """登录 IEC → 调 exportExcel → 下载 xlsx → 返回 (df, xlsx_path)

    复制后如需改为其他 IEC 数据源（如订单查询、合约查询），
    只改本函数内的 apiBean/methodName/param 即可，其余阶段无需动。
    """
    if not _HAS_PANDAS:
        raise ImportError("需要 pandas 和 openpyxl：pip install pandas openpyxl")

    # 1. 登录 IEC
    username = env("IBAO_USERNAME")
    password = env("IBAO_PASSWORD")
    if not username or not password:
        raise RuntimeError("缺少 IBAO_USERNAME / IBAO_PASSWORD 环境变量")
    iec = IEC(username, password)
    if not iec.login():
        raise RuntimeError("IEC 登录失败")
    print(f"[下载] IEC 登录成功")

    try:
        s = iec.session
        token = iec.token
        referer = f"{_IECS_INDEX}?token={token}"

        # 2. 建立 iecs 会话
        s.get(f"{_IECS_INDEX}?token={token}", timeout=15,
              headers={"Referer": referer}).raise_for_status()
        s.get(_QUASI_HAIR_PAGE, timeout=15,
              headers={"Referer": referer}).raise_for_status()

        # 3. 选择 apiBean/methodName（download=订单级 / bundle=捆包级）
        if download_type == "bundle":
            api_bean = "com.baosight.iecs.freight.quasiHair.api.IQuasiHairDetailService"
            method_name = "downQuasiHairDetail"
            type_label = "捆包明细"
        else:
            api_bean = "com.baosight.iecs.freight.quasiHair.api.IQuasiHairService"
            method_name = "queryQuasiHairDownload"
            type_label = "订单级别"

        # 4. 调 exportExcel 接口
        param = {
            "contractNum": "", "orderNum": "", "factoryProductId": "",
            "shopsign": "", "prodCodeName": "", "prodCode": "",
            "deliveryDateChrStart": start, "deliveryDateChrEnd": end,
            "custType": "", "memberCodeFlag": "", "machineId": "",
            "apiBean": api_bean, "methodName": method_name,
            "settleUserNum": _SETTLE_USER_NUM, "factoryOrderNums": "",
            "offset": 0, "limit": 1000,
        }
        print(f"[下载] 请求 {start}~{end} {type_label}数据（{download_type}模式）...")
        r = s.post(_EXPORT_API, data=json.dumps(param), timeout=120,
                   headers={
                       "Content-Type": "application/json; charset=utf-8",
                       "Accept": "application/json, text/javascript, */*; q=0.01",
                       "X-Requested-With": "XMLHttpRequest",
                       "Referer": _QUASI_HAIR_PAGE,
                   })
        r.raise_for_status()
        result = r.json()
        if result.get("code") != 0:
            raise RuntimeError(f"下载失败: {result.get('msg', '未知错误')}")
        server_file = result["msg"]
        print(f"[下载] 服务器生成文件: {server_file}")

        # 5. 下载实际文件
        dl_url = f"{_DOWNLOAD_API}?fileName={urllib.parse.quote(server_file)}&delete=true"
        r2 = s.get(dl_url, timeout=120, stream=True,
                   headers={"Referer": _QUASI_HAIR_PAGE})
        r2.raise_for_status()
        with open(out_path, "wb") as f:
            for chunk in r2.iter_content(8192):
                if chunk:
                    f.write(chunk)
        size = os.path.getsize(out_path)
        print(f"[下载] xlsx 已保存: {out_path} ({size:,} bytes)")

        # 6. 读取为 DataFrame（预览用）
        df = pd.read_excel(out_path)
        print(f"[下载] 数据: {len(df)} 行 × {len(df.columns)} 列")
        return df, out_path
    finally:
        iec.save("iecc.json")


# ============================================================
# 阶段2：飞书云盘上传（固定，一般无需改动）
# ============================================================
def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": app_id, "app_secret": app_secret},
                      timeout=15, proxies=NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换 tenant_access_token 失败: {data}")
    token = str(data.get("tenant_access_token") or "")
    if not token:
        raise RuntimeError(f"响应中无 tenant_access_token: {data}")
    print(f"[云盘] tenant_access_token OK (len={len(token)})")
    return token


def feishu_upload_file(token: str, folder_token: str,
                      file_path: str, max_size_mb: int = 20) -> dict:
    if not os.path.exists(file_path):
        raise FileNotFoundError(file_path)
    size = os.path.getsize(file_path)
    if size == 0:
        raise ValueError("上传的文件为空")
    if size > max_size_mb * 1024 * 1024:
        raise ValueError(f"文件 {size/1024/1024:.1f}MB 超过上限 {max_size_mb}MB")
    if not folder_token:
        raise ValueError("缺少 folder_token")

    filename = os.path.basename(file_path)
    with open(file_path, "rb") as fh:
        files = {"file": (filename, fh.read(), "application/octet-stream")}
    data = {"file_name": filename, "parent_type": "explorer",
            "parent_node": folder_token, "size": str(size)}
    r = requests.post(f"{FEISHU_OPEN_BASE}/drive/v1/files/upload_all",
                      data=data, files=files,
                      headers={"Authorization": f"Bearer {token}"},
                      timeout=120, proxies=NO_PROXY)
    resp = r.json()
    if resp.get("code") != 0:
        raise RuntimeError(f"上传失败: code={resp.get('code')}, msg={resp.get('msg')}")
    info = resp.get("data") or {}
    print(f"[云盘] ✅ 上传成功 name={info.get('name') or filename} "
          f"file_token={info.get('file_token') or info.get('token') or ''}")
    return info


def stage_upload(file_path: str, folder_token: str) -> Optional[dict]:
    """上传文件到飞书云盘。缺凭据时跳过不报错。"""
    app_id = env("FEISHU_APP_ID")
    app_secret = env("FEISHU_APP_SECRET")
    if not app_id or not app_secret:
        print("[云盘] 未配置 FEISHU_APP_ID / FEISHU_APP_SECRET，跳过上传")
        return None
    if not folder_token:
        print("[云盘] 未配置 FEISHU_FOLDER_TOKEN，跳过上传")
        return None
    try:
        token = feishu_tenant_access_token(app_id, app_secret)
        return feishu_upload_file(token, folder_token, file_path)
    except Exception as e:
        print(f"[云盘] ⚠️ 上传失败: {e}")
        return None


# ============================================================
# 阶段3：飞书机器人通知（固定，一般无需改动）
# ============================================================
def _feishu_sign(secret: str, timestamp: str) -> str:
    import base64 as _b64
    string_to_sign = f"{timestamp}\n{secret}"
    h = hmac.new(secret.encode("utf-8"),
                 string_to_sign.encode("utf-8"), hashlib.sha256)
    return _b64.b64encode(h.digest()).decode("utf-8")


def stage_notify(message: str):
    """飞书机器人 Webhook 通知。缺 webhook URL 时跳过。"""
    webhook_url = env("FEISHU_WEBHOOK_URL")
    if not webhook_url:
        return
    secret = env("FEISHU_WEBHOOK_SECRET")
    body = {"msg_type": "text", "content": {"text": message}}
    if secret:
        ts = str(int(time.time()))
        body["timestamp"] = ts
        body["sign"] = _feishu_sign(secret, ts)
    try:
        requests.post(webhook_url, json=body, timeout=10, proxies=NO_PROXY)
    except Exception:
        pass


# ============================================================
# 主流程
# ============================================================
def parse_args():
    p = argparse.ArgumentParser(
        description="IEC → 飞书云盘 交付模板",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--type", choices=["download", "bundle"], default="download",
                   help="下载模式（download=订单级 / bundle=捆包级）")
    p.add_argument("--start", type=_parse_month, default="",
                   help="交货期起始月 YYYYMM（默认近3月）")
    p.add_argument("--end", type=_parse_month, default="",
                   help="交货期结束月 YYYYMM（默认下月）")
    p.add_argument("-o", "--output", default="",
                   help="输出文件名（默认自动命名）")
    p.add_argument("--dry-run", action="store_true",
                   help="预览模式（只下载不上传）")
    return p.parse_args()


def main():
    args = parse_args()
    dry_run = args.dry_run or env("DRY_RUN") == "1"

    start = args.start or env("EXPORT_START")
    end = args.end or env("EXPORT_END")
    if not start or not end:
        start, end = _default_range()

    out_path = args.output or _auto_filename()
    folder_token = env("FEISHU_FOLDER_TOKEN", DEFAULT_FOLDER_TOKEN)

    print(f"====== IEC → 飞书云盘 交付模板 ======")
    print(f"  模式: {args.type}  范围: {start}~{end}  {'DRY-RUN' if dry_run else ''}")
    print(f"  输出: {out_path}")
    print(f"  云盘文件夹: {folder_token or '(未配置)'}")

    t0 = time.time()
    upload_info = None
    df = None

    # 阶段1：下载
    print("\n[阶段1] IEC 下载...")
    try:
        df, out_path = stage_download(start, end, out_path,
                                      download_type=args.type,
                                      dry_run=dry_run)
    except Exception as e:
        print(f"[阶段1] ❌ 下载失败: {e}")
        return 1

    # 阶段2：上传（dry-run 跳过）
    print("\n[阶段2] 飞书云盘上传...")
    if dry_run:
        print("  [DRY-RUN] 跳过上传")
    else:
        upload_info = stage_upload(out_path, folder_token)

    # 阶段3：通知
    rows = len(df) if df is not None else 0
    elapsed = time.time() - t0
    msg = (
        f"【IEC→云盘】{now_shanghai_str()} {'DRY-RUN ' if dry_run else ''}\n"
        f"类型: {args.type}  范围: {start}~{end}\n"
        f"数据: {rows} 行  文件: {os.path.basename(out_path)}\n"
        f"云盘: {'已上传' if upload_info else ('跳过' if dry_run else '失败/未配置')}\n"
        f"耗时: {elapsed:.1f}s"
    )
    print(f"\n[阶段3] 机器人通知...")
    stage_notify(msg)

    print(f"\n{'='*60}")
    print(msg)
    print(f"{'='*60}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
