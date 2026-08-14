#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cookie-Editor 导出 JSON → Playwright storage_state.json 转换器

用法：
  1. 用本机 Chrome/Edge 登录欧冶 https://login-ng.ouyeel.com/sso/login?service=https://www.ouyeel.com/
  2. 登录成功后，在浏览器装 Cookie-Editor 扩展（Chrome 应用商店搜 "Cookie-Editor"）
  3. 点扩展图标 → 右下角 Export → Export as JSON → 粘贴到任意文本编辑器存为 cookies.json
  4. 运行：python feishu_uploader/cookie_to_state.py cookies.json
  5. 自动生成 .ouyeel_state.json + .ouyeel_state.json.b64
  6. 把 .b64 内容粘贴到 GitHub Secret OUYEEL_STORAGE_STATE

Playwright 的 storage_state 路径（CI 里用）就是 .ouyeel_state.json。
本脚本只做格式转换，不依赖 Playwright。
"""

import argparse
import base64
import json
import os
import sys


STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".ouyeel_state.json")
B64_FILE = STATE_FILE + ".b64"

# 关键 cookie 域名过滤：只保留欧冶相关域（防止用户误把其他站点 cookie 混入）
ALLOWED_DOMAIN_SUFFIXES = (
    "ouyeel.com",
    "ouyeel.com.cn",
)

# 同 Site 字段映射：Cookie-Editor 用的是浏览器扩展标准，Playwright 用的是 StorageState 标准
SAME_SITE_MAP = {
    "no_restriction": "None",
    "lax": "Lax",
    "strict": "Strict",
    "unspecified": "Lax",
    "none": "None",
}


def convert_cookie_editor_to_playwright(cookies_json: list, verbose: bool = True) -> dict:
    """把 Cookie-Editor 导出的 cookie 数组转为 Playwright storage_state 格式

    Cookie-Editor 格式（每个 cookie）:
        {"domain": ".ouyeel.com", "expirationDate": 123.45, "hostOnly": false,
         "httpOnly": false, "name": "SESSION", "path": "/", "sameSite": "no_restriction",
         "secure": true, "session": false, "value": "abc"}

    Playwright storage_state 格式:
        {"cookies": [{"name": "...", "value": "...", "domain": "...", "path": "...",
                       "expires": 123, "httpOnly": false, "secure": true, "sameSite": "None"}],
         "origins": []}
    """
    if not isinstance(cookies_json, list):
        raise ValueError(f"输入应为 cookie 数组，实际类型: {type(cookies_json).__name__}")

    converted: list[dict] = []
    skipped_other_domain: list[str] = []

    for c in cookies_json:
        if not isinstance(c, dict):
            continue
        name = c.get("name", "")
        value = c.get("value", "")
        domain = c.get("domain", "")
        path = c.get("path", "/")

        # 域名过滤：只保留欧冶相关 cookie
        domain_lower = domain.lstrip(".").lower()
        if not any(domain_lower.endswith(s) for s in ALLOWED_DOMAIN_SUFFIXES):
            skipped_other_domain.append(f"{name}@{domain}")
            continue

        # 过期时间：Cookie-Editor 用 expirationDate（秒，可能带小数）
        # Playwright 用 expires（整数秒，0 或 -1 表示 session cookie）
        expires_raw = c.get("expirationDate")
        if expires_raw is None or c.get("session") is True:
            expires = -1  # session cookie
        else:
            try:
                expires = int(float(expires_raw))
            except (ValueError, TypeError):
                expires = -1

        same_site_raw = (c.get("sameSite") or "lax").lower()
        same_site = SAME_SITE_MAP.get(same_site_raw, "Lax")

        converted.append({
            "name": name,
            "value": value,
            "domain": domain,
            "path": path,
            "expires": expires,
            "httpOnly": bool(c.get("httpOnly", False)),
            "secure": bool(c.get("secure", False)),
            "sameSite": same_site,
        })

    if verbose:
        print(f"[转换] 输入 {len(cookies_json)} 条 cookie")
        print(f"[转换] 欧冶相关保留: {len(converted)} 条")
        if skipped_other_domain:
            print(f"[转换] 跳过非欧冶域名: {len(skipped_other_domain)} 条")
            for s in skipped_other_domain[:5]:
                print(f"        - {s}")
            if len(skipped_other_domain) > 5:
                print(f"        ... 共 {len(skipped_other_domain)} 条")

    return {"cookies": converted, "origins": []}


def validate_state(state: dict) -> list[str]:
    """校验 storage_state 是否合理，返回警告列表（空=无警告）"""
    warnings: list[str] = []
    cookies = state.get("cookies", [])
    if not cookies:
        warnings.append("⚠️ cookie 列表为空！登录可能未完成，或导出的文件有误")
        return warnings

    names = {c["name"] for c in cookies}
    # 欧冶常见的会话 cookie 关键字（不一定都有，但应该至少有一个）
    session_keywords = ("SESSION", "session", "JSESSIONID", "token", "TOKEN",
                         "accessToken", "access_token", "ST cookie", "sso_token")
    has_session = any(kw.lower() in n.lower() for n in names for kw in session_keywords)
    if not has_session:
        warnings.append("⚠️ 未找到任何会话类 cookie（SESSION/token/JSESSIONID 等），登录态可能无效")

    # 过期检查
    import time
    expired = [c["name"] for c in cookies
               if c["expires"] > 0 and c["expires"] < time.time()]
    if expired:
        warnings.append(f"⚠️ {len(expired)} 条 cookie 已过期: {expired[:3]}")

    return warnings


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Cookie-Editor 导出 JSON → Playwright storage_state.json"
    )
    parser.add_argument("input", help="Cookie-Editor 导出的 JSON 文件路径")
    parser.add_argument("-o", "--output", default=STATE_FILE,
                        help=f"输出 storage_state 路径（默认: {STATE_FILE}）")
    parser.add_argument("--no-b64", action="store_true",
                        help="不生成 .b64 文件")
    args = parser.parse_args()

    # 1) 读 Cookie-Editor JSON
    if not os.path.exists(args.input):
        print(f"[错误] 输入文件不存在: {args.input}")
        return 1

    try:
        with open(args.input, "r", encoding="utf-8-sig") as fp:
            content = fp.read().strip()
    except UnicodeDecodeError:
        with open(args.input, "r", encoding="gbk") as fp:
            content = fp.read().strip()
    except Exception as exc:
        print(f"[错误] 读取文件失败: {exc}")
        return 1

    # Cookie-Editor 导出时，外层可能是数组，也可能包了一层
    try:
        data = json.loads(content)
    except json.JSONDecodeError as exc:
        print(f"[错误] JSON 解析失败: {exc}")
        print("[提示] 请确认是 Cookie-Editor 的 Export → Export as JSON 格式（应为数组）")
        return 1

    # 兼容：如果用户从 F12 复制的是单个 cookie 对象
    if isinstance(data, dict):
        if "cookies" in data and isinstance(data["cookies"], list):
            # 已经是 Playwright 格式
            print("[提示] 输入已是 storage_state 格式，直接保存")
            state = data
        else:
            data = [data]
            state = convert_cookie_editor_to_playwright(data)
    else:
        state = convert_cookie_editor_to_playwright(data)

    # 2) 校验
    warnings = validate_state(state)
    for w in warnings:
        print(w)

    if not state.get("cookies"):
        print("\n[错误] 转换后 cookie 为空，未生成 storage_state")
        return 1

    # 3) 保存 storage_state.json
    try:
        with open(args.output, "w", encoding="utf-8") as fp:
            json.dump(state, fp, ensure_ascii=False, indent=2)
        print(f"\n[OK] storage_state 已保存到: {args.output}")
        print(f"      cookie 数: {len(state['cookies'])}")
    except Exception as exc:
        print(f"[错误] 保存失败: {exc}")
        return 1

    # 4) 生成 base64 文件（方便粘贴到 GitHub Secret）
    if not args.no_b64:
        try:
            with open(args.output, "rb") as fp:
                raw = fp.read()
            b64 = base64.b64encode(raw).decode("ascii")
            with open(B64_FILE, "w", encoding="ascii") as fp:
                fp.write(b64)
            print(f"[OK] base64 已保存到: {B64_FILE}")
            print(f"      长度 {len(b64)} 字符")
            print(f"      可直接粘贴到 GitHub Secret OUYEEL_STORAGE_STATE")
        except Exception as exc:
            print(f"[警告] base64 生成失败: {exc}")

    # 5) 给出下一步操作提示
    print("\n下一步：")
    print("  1. 本地测试：set DELIVERY_MODE=local && python feishu_uploader/export_ouyeel.py")
    print("     （会自动加载 .ouyeel_state.json 跑一次数据抓取）")
    print("  2. 配 GitHub Secret OUYEEL_STORAGE_STATE ← .ouyeel_state.json.b64 内容")
    print("  3. 配 GitHub Secret OUYEEL_FOLDER_TOKEN ← 欧冶目标飞书文件夹 token")
    print("  4. Actions 页面手动 Run workflow 验证")

    return 0


if __name__ == "__main__":
    sys.exit(main())
