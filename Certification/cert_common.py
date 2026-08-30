#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Certification 公共模块：飞书云盘 API + 机器人汇报
==================================================

被三个阶段脚本共用：
  stage1_download.py   IEC 下载 → 原件文件夹
  stage2_classify.py   A17 识别分类 → 新识别/未识别文件夹 → 已处理
  stage3_organize.py   梳理 → 写 Bitable「质保书」附件字段

汇报（机器人）：
  优先级：1. 私信  2. 群 Webhook  3. 仅打日志
  汇报人员环境变量（不配则不发私信）：
    CERT_NOTIFY_UNION_IDS        逗号/分号/空格分隔（质保书流程专用，优先）
    FEISHU_UNION_IDS             （回退，全仓库通用名单）
  ID 前缀自动识别 receive_id_type：on_→union_id / oc_→open_id / ou_→user_id / @→email
    CERT_NOTIFY_WEBHOOK_URL      （可选）覆盖 FEISHU_WEBHOOK_URL
    FEISHU_WEBHOOK_SECRET        （可选）Webhook 签名密钥
    NOTIFY_APP_ID / NOTIFY_APP_SECRET  发私信用应用（回退 FEISHU_APP_ID/SECRET）

环境变量（.env 或 GitHub Secrets）：
  FEISHU_APP_ID / FEISHU_APP_SECRET  飞书自建应用（drive:drive + 文件夹协作者）
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Optional

import requests

# 控制台非 UTF-8（如 GBK）时，emoji 打印会 UnicodeEncodeError，统一重配为 utf-8
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
NO_PROXY = {"http": None, "https": None}

DRIVE_WEB_BASE = "https://s2v31ke6sl.feishu.cn/drive"


def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default) or default


def log(msg: str, log_file: Optional[Path] = None) -> None:
    line = f'[{time.strftime("%H:%M:%S")}] {msg}'
    print(line, flush=True)
    if log_file:
        try:
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass


def folder_url(folder_token: str) -> str:
    return f"{DRIVE_WEB_BASE}/folder/{folder_token}"


# ============================================================
# tenant_access_token
# ============================================================
def tenant_access_token(app_id: str = "", app_secret: str = "") -> str:
    app_id = app_id or env("FEISHU_APP_ID")
    app_secret = app_secret or env("FEISHU_APP_SECRET")
    if not app_id or not app_secret:
        raise RuntimeError("未配置 FEISHU_APP_ID / FEISHU_APP_SECRET")
    r = requests.post(
        f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret},
        timeout=15, proxies=NO_PROXY,
    )
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"tenant_access_token 失败: {d}")
    tok = str(d.get("tenant_access_token") or "")
    if not tok:
        raise RuntimeError(f"无 token: {d}")
    return tok


# ============================================================
# 云盘文件夹操作
# ============================================================
def list_folder_files(token: str, folder_token: str) -> list[dict]:
    """列出指定文件夹下所有文件（分页）。返回 [{name, token, type, size, ...}]。
    注意：飞书 list 接口的查询参数是 folder_token（不是 folder_id）。"""
    files: list[dict] = []
    page_token = ""
    url = f"{FEISHU_OPEN_BASE}/drive/v1/files"
    headers = {"Authorization": f"Bearer {token}"}
    page = 0
    while True:
        page += 1
        params = {"folder_token": folder_token, "page_size": 50}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(url, headers=headers, params=params,
                         timeout=30, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") != 0:
            raise RuntimeError(f"列文件失败 page={page}: {d}")
        data = d.get("data") or {}
        files.extend(data.get("files") or [])
        if not data.get("has_more"):
            break
        page_token = data.get("next_page_token") or data.get("page_token") or ""
        if not page_token:
            break
    return files


def download_file(token: str, file_token: str) -> bytes:
    """下载云盘文件，返回 bytes。"""
    url = f"{FEISHU_OPEN_BASE}/drive/v1/files/{file_token}/download"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers, timeout=120, proxies=NO_PROXY)
    if r.status_code != 200:
        raise RuntimeError(f"下载失败 {file_token}: HTTP {r.status_code} {r.text[:200]}")
    return r.content


def upload_to_folder(token: str, folder_token: str, file_name: str, raw: bytes) -> str:
    """上传文件到云盘文件夹（parent_type=explorer），返回 file_token。"""
    size = len(raw)
    if size == 0:
        raise ValueError("空文件")
    if size > 200 * 1024 * 1024:
        raise ValueError(f"文件超过 200MB ({size/1024/1024:.1f}MB)")
    data = {
        "file_name": file_name,
        "parent_type": "explorer",
        "parent_node": folder_token,
        "size": str(size),
    }
    files = {"file": (file_name, raw, "application/octet-stream")}
    r = requests.post(
        f"{FEISHU_OPEN_BASE}/drive/v1/files/upload_all",
        data=data, files=files,
        headers={"Authorization": f"Bearer {token}"},
        timeout=180, proxies=NO_PROXY,
    )
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"上传失败 code={d.get('code')} msg={d.get('msg')} data={d}")
    ft = (d.get("data") or {}).get("file_token") or (d.get("data") or {}).get("token") or ""
    if not ft:
        raise RuntimeError(f"上传返回空 file_token: {d}")
    return ft


def move_file(token: str, file_token: str, target_folder: str) -> dict:
    """移动云盘文件到目标文件夹。"""
    url = f"{FEISHU_OPEN_BASE}/drive/v1/files/{file_token}/move"
    body = {"type": "file", "folder_token": target_folder}
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"},
                      json=body, timeout=30, proxies=NO_PROXY)
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"移动文件失败 {file_token}: code={d.get('code')} msg={d.get('msg')}")
    return d.get("data") or {}


# ============================================================
# Bitable（阶段3用）
# ============================================================
def list_bitable_records(token: str, app_token: str, table_id: str) -> list[dict]:
    """分页拉取 Bitable 表全部记录。"""
    records: list[dict] = []
    page_token = ""
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/records")
    headers = {"Authorization": f"Bearer {token}"}
    page = 0
    while True:
        page += 1
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(url, headers=headers, params=params,
                         timeout=30, proxies=NO_PROXY)
        d = r.json()
        if d.get("code") != 0:
            raise RuntimeError(f"拉记录失败 page={page}: {d}")
        data = d.get("data") or {}
        records.extend(data.get("items") or [])
        if not data.get("has_more"):
            break
        page_token = data.get("page_token") or ""
        if not page_token:
            break
    return records


def upload_bitable_media(token: str, app_token: str,
                         file_name: str, raw: bytes) -> str:
    """上传文件到 bitable 附件空间（parent_type=bitable_file），返回 file_token。
    限制：单文件 ≤ 20MB。"""
    size = len(raw)
    if size == 0:
        raise ValueError("空文件")
    if size > 20 * 1024 * 1024:
        raise ValueError(f"文件超过 20MB ({size/1024/1024:.1f}MB)，需分片上传（未实现）")
    data = {
        "file_name": file_name,
        "parent_type": "bitable_file",
        "parent_node": app_token,
        "size": str(size),
    }
    files = {"file": (file_name, raw, "application/octet-stream")}
    r = requests.post(
        f"{FEISHU_OPEN_BASE}/drive/v1/medias/upload_all",
        data=data, files=files,
        headers={"Authorization": f"Bearer {token}"},
        timeout=180, proxies=NO_PROXY,
    )
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"上传素材失败: {d}")
    return (d.get("data") or {}).get("file_token") or ""


def update_record_field(token: str, app_token: str, table_id: str,
                        record_id: str, field_name: str,
                        field_value) -> dict:
    """更新单条记录指定字段。"""
    url = (f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{app_token}"
           f"/tables/{table_id}/records/{record_id}")
    body = {"fields": {field_name: field_value}}
    r = requests.put(url, headers={"Authorization": f"Bearer {token}"},
                     json=body, timeout=30, proxies=NO_PROXY)
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"更新记录失败 {record_id}: {d}")
    return d.get("data") or {}


# ============================================================
# 机器人汇报（union_id 私信 → Webhook → 仅日志）
# ============================================================
def _feishu_sign(secret: str, ts: str) -> str:
    import base64
    h = hmac.new(secret.encode("utf-8"),
                 f"{ts}\n{secret}".encode("utf-8"), hashlib.sha256)
    return base64.b64encode(h.digest()).decode("utf-8")


def _notify_union_ids() -> list[str]:
    raw = env("CERT_NOTIFY_UNION_IDS") or env("FEISHU_UNION_IDS")
    if not raw:
        return []
    return [p for p in re.split(r"[\s,，;；]+", raw.strip()) if p]


def _id_type(oid: str) -> str:
    """按 ID 前缀自动判断 receive_id_type（与结案流程 iec_jiean_to_bitable 一致）：
    oc_ → open_id / on_ → union_id / ou_ → user_id / 含@ → email / 其他 → open_id
    """
    if not oid:
        return "open_id"
    if oid.startswith("oc_"):
        return "open_id"
    if oid.startswith("on_"):
        return "union_id"
    if oid.startswith("ou_"):
        return "user_id"
    if "@" in oid:
        return "email"
    return "open_id"


def notify(message: str) -> dict:
    """发送机器人汇报。返回 {channel, ok_count, fail_count}。"""
    stat = {"channel": "log-only", "ok": 0, "fail": 0}

    # 1) union_id 私信
    union_ids = _notify_union_ids()
    if union_ids:
        app_id = env("NOTIFY_APP_ID") or env("FEISHU_APP_ID")
        app_secret = env("NOTIFY_APP_SECRET") or env("FEISHU_APP_SECRET")
        if app_id and app_secret:
            try:
                tok = tenant_access_token(app_id, app_secret)
                for oid in union_ids:
                    try:
                        r = requests.post(
                            f"{FEISHU_OPEN_BASE}/im/v1/messages",
                            headers={"Authorization": f"Bearer {tok}",
                                     "Content-Type": "application/json; charset=utf-8"},
                            params={"receive_id_type": _id_type(oid)},
                            json={"receive_id": oid, "msg_type": "text",
                                  "content": json.dumps({"text": message}, ensure_ascii=False)},
                            timeout=15, proxies=NO_PROXY,
                        )
                        d = r.json()
                        ok = d.get("code") == 0
                        stat["ok" if ok else "fail"] += 1
                        print(f"[通知·私信] {'OK  ' if ok else 'FAIL'} → {oid[:16]}… "
                              f"msg_id={(d.get('data') or {}).get('message_id', '')}", flush=True)
                    except Exception as e:
                        stat["fail"] += 1
                        print(f"[通知·私信] 异常: {e}", flush=True)
                    time.sleep(0.15)
                stat["channel"] = "dm"
                return stat
            except Exception as e:
                print(f"[通知·私信] 拿 token 失败，降级 Webhook: {e}", flush=True)

    # 2) Webhook
    url = env("CERT_NOTIFY_WEBHOOK_URL") or env("FEISHU_WEBHOOK_URL")
    if url:
        body: dict = {"msg_type": "text", "content": {"text": message}}
        secret = env("FEISHU_WEBHOOK_SECRET")
        if secret:
            ts = str(int(time.time()))
            body["timestamp"] = ts
            body["sign"] = _feishu_sign(secret, ts)
        try:
            r = requests.post(url, json=body, timeout=10, proxies=NO_PROXY)
            d = r.json() if r.content else {}
            ok = r.status_code == 200
            stat["ok" if ok else "fail"] += 1
            stat["channel"] = "webhook"
            print(f"[通知·Webhook] status={r.status_code} resp={d}", flush=True)
            return stat
        except Exception as e:
            print(f"[通知·Webhook] 异常: {e}", flush=True)

    # 3) 仅日志
    print("[通知] 未配置 CERT_NOTIFY_UNION_IDS / FEISHU_WEBHOOK_URL，仅打印：\n" + message, flush=True)
    return stat
