#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火钢城系统登录模块（共用）

所有需要登录 erpa.donghuo.vip 的脚本均可复用本模块，避免重复编写
create_session / recognize_captcha / login 三段代码。

用法：
    from donghuo_login import login_donghuo

    # 方式一：从环境变量 DH_USERNAME / DH_PASSWORD 读取（推荐，与 CI 一致）
    session = login_donghuo()
    if session is None:
        raise SystemExit("登录失败")

    # 方式二：直接传入账号密码（调试用，请勿提交真实账号）
    session = login_donghuo(username="your_username", password="your_password")

    # 登录后即可用 session 调任意接口
    resp = session.post(
        "https://erpa.donghuo.vip/model/admin/xiaoshou/m_xiaoshou/xjilulist",
        data={"page": 1, "limit": 30},
        headers={"X-Requested-With": "XMLHttpRequest"},
    )

直接运行本文件可快速验证登录是否正常：
    python donghuo_login.py
"""

import os
import time
import json

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_URL = "https://erpa.donghuo.vip"
LOGIN_URL = f"{BASE_URL}/controller/admin/c_longin/index"
CAPTCHA_URL = f"{BASE_URL}/common/captcha"

# ddddocr 单例（懒加载，避免 import 模块时就加载 ONNX 大模型）
# 单例化后每次识别复用同一个实例，避免反复加载模型导致超时
_ocr_instance = None
_ocr_init_error = None


def _get_ocr():
    """获取 ddddocr 单例，首次调用时初始化模型（只需加载一次）"""
    global _ocr_instance, _ocr_init_error
    if _ocr_instance is not None:
        return _ocr_instance
    if _ocr_init_error is not None:
        # 上次初始化失败过，直接返回 None 避免再次尝试
        return None
    try:
        import ddddocr  # type: ignore
        print("[ddddocr] 首次加载 ONNX 模型（仅一次，后续复用）...", flush=True)
        _ocr_instance = ddddocr.DdddOcr(show_ad=False)
        print("[ddddocr] 模型加载完成", flush=True)
        return _ocr_instance
    except ImportError:
        _ocr_init_error = "ImportError"
        print("[警告] 未安装 ddddocr，无法识别验证码", flush=True)
        return None
    except Exception as exc:
        _ocr_init_error = str(exc)
        print(f"[错误] ddddocr 初始化失败: {exc}", flush=True)
        return None


def create_session() -> requests.Session:
    """创建带重试机制的 requests Session"""
    s = requests.Session()
    retry = Retry(total=5,
                  backoff_factor=1,
                  status_forcelist=[500, 502, 503, 504, 429],
                  allowed_methods=["POST", "GET"])
    adapter = HTTPAdapter(max_retries=retry)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    s.headers.update({
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Accept-Language": "zh-CN,zh;q=0.9",
    })
    return s


def recognize_captcha(image_bytes: bytes) -> str:
    """用 ddddocr 识别图形验证码（复用单例 OCR 实例），失败返回空串"""
    ocr = _get_ocr()
    if ocr is None:
        return ""
    try:
        return ocr.classification(image_bytes).strip().replace(" ", "")
    except Exception as exc:
        print(f"[错误] 验证码识别异常: {exc}", flush=True)
        return ""


def login(session: requests.Session,
          username: str,
          password: str,
          max_attempts: int = 10) -> bool:
    """
    登录懂火系统，自动重试验证码。

    登录接口: POST /controller/admin/c_longin/index
    成功返回 True，否则 False。
    """
    for attempt in range(1, max_attempts + 1):
        print(f"[登录] 尝试 {attempt}/{max_attempts} ...", flush=True)

        # 1) 拉取验证码图片（同时建立/刷新 Session Cookie）
        img_resp = session.get(CAPTCHA_URL, timeout=15)
        if img_resp.status_code != 200:
            print(f"  获取验证码失败: HTTP {img_resp.status_code}", flush=True)
            time.sleep(2)
            continue

        # 2) 识别验证码
        captcha_code = recognize_captcha(img_resp.content)
        if not captcha_code:
            captcha_code = "1234"  # 占位，必然失败，但可触发刷新
        print(f"  识别结果: {captcha_code}", flush=True)

        # 3) 提交登录
        resp = session.post(LOGIN_URL,
                            data={"u_name": username,
                                  "u_pass": password,
                                  "captcha": captcha_code},
                            timeout=15)
        text = resp.text.strip()
        try:
            result = json.loads(text)
            if isinstance(result, dict) and str(result.get("code")) == "200":
                print("[登录] 成功", flush=True)
                return True
            msg = result.get("msg") or result.get("message") or text[:100]
            print(f"  失败: {msg}", flush=True)
        except json.JSONDecodeError:
            print(f"  非JSON响应(前200): {text[:200]}", flush=True)

        time.sleep(2)

    print(f"[登录] 已达最大尝试次数 {max_attempts}，登录失败", flush=True)
    return False


def login_donghuo(username: str | None = None,
                  password: str | None = None,
                  max_attempts: int = 10) -> requests.Session | None:
    """
    一键登录懂火系统：创建 session 并登录，返回已登录的 session。

    - 账号/密码未传入时，从环境变量 DH_USERNAME / DH_PASSWORD 读取
      （与 GitHub Actions / CI 约定一致）
    - 登录失败返回 None

    示例：
        session = login_donghuo()
        if session is None:
            raise SystemExit("登录失败")
    """
    username = (username or os.environ.get("DH_USERNAME", "")).strip()
    password = (password or os.environ.get("DH_PASSWORD", "")).strip()
    if not username or not password:
        print("[错误] 缺少账号/密码：请传入 username/password，"
              "或设置环境变量 DH_USERNAME / DH_PASSWORD", flush=True)
        return None

    session = create_session()
    if login(session, username, password, max_attempts=max_attempts):
        return session
    return None


if __name__ == "__main__":
    # 直接运行本文件可验证登录是否正常
    # 必须设置环境变量 DH_USERNAME / DH_PASSWORD
    _u = os.environ.get("DH_USERNAME", "").strip()
    _p = os.environ.get("DH_PASSWORD", "").strip()
    if not _u or not _p:
        print("[自检] 未设置 DH_USERNAME / DH_PASSWORD 环境变量，无法自检")
        raise SystemExit(1)
    print(f"[自检] 使用账号: {_u}")
    s = login_donghuo(username=_u, password=_p)
    if s is not None:
        print("[自检] 登录态有效，session 已就绪")
    else:
        print("[自检] 登录失败")
        raise SystemExit(1)
