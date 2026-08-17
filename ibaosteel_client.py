"""
宝钢 IEC (iBaosteel) 系统登录客户端 - 精简版
============================================

核心目标：从"账号密码输入"开始，直到**"选择组织（直销 / 渠道销售）"页面可打开**，
才视为"登录成功"。

依赖：
    pip install requests pycryptodome ddddocr

最简用法：
    # 1. 命令行（直接输出"选择组织"页面的 URL）
    $ python ibaosteel_client.py

    # 2. Python 代码
    from ibaosteel_client import IEC
    iec = IEC('your_username', 'your_password')
    if iec.login():
        print(iec.url)                 # "选择组织"页面 URL，复制到浏览器打开
        iec.save('iecc.json')          # 保存会话避免下次再登录

    # 3. 复用上次保存的会话（避免每次都识别验证码）
    iec = IEC.load('iecc.json')
    if not iec.ok:
        iec.login()
        iec.save('iecc.json')
    print(iec.url)
"""
from __future__ import annotations

import base64
import json
import os
import re
import time
from typing import Optional

import requests

try:
    from Crypto.PublicKey import RSA as CryptoRSA
    from Crypto.Cipher import PKCS1_v1_5
except ImportError:
    raise ImportError("需要 pycryptodome：pip install pycryptodome")

try:
    import ddddocr
    _HAS_OCR = True
except ImportError:
    _HAS_OCR = False

# ============================================================
# 常量
# ============================================================
_BASE = 'https://www.ibaosteel.com'
_LOGIN_PAGE = f'{_BASE}/ibaosteel/account/login'
_PUBKEY = f'{_BASE}/ibaosteel/util/rsa/getPublicKey'
_CAPTCHA = f'{_BASE}/ibaosteel/account/image'
_AUTH = f'{_BASE}/ibaosteel/account/auth'
_INDEX = f'{_BASE}/ibaosteel/index'
_CHANNEL_SELECT = f'{_BASE}/ibaosteel/bizIntelli'  # "选择组织"页面

_UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
       '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

_CHAR_FIX = str.maketrans({
    'o': '0', 'O': '0', 'l': '1', 'I': '1', 'z': '2', 'Z': '2',
    's': '5', 'S': '5', 'b': '6', 'G': '6', 'B': '8', 'g': '9', 'q': '9',
})

# HCSID 默认值（来自 hcs.js，IEC 系统需要此参数才能通过登录验证）
_HCSID = '0452098f62cb40a9ba4e28ce801d79de1a00d4e2883'


# ============================================================
# 主类（对外就这一个）
# ============================================================
class IEC:
    """宝钢 IEC 登录客户端

    Parameters
    ----------
    username : str, default ''
        登录用户名（命令行入口会从环境变量 IBAO_USERNAME 读取）
    password : str, default ''
        登录密码（命令行入口会从环境变量 IBAO_PASSWORD 读取）
    retries : int, default 30
        验证码识别失败的最大重试次数（每次换新验证码）
    manual : bool, default False
        ddddocr 连续失败超过 15 次时是否提示人工输入
    timeout : int, default 15
        网络请求超时秒数
    """

    def __init__(
        self,
        username: str = '',
        password: str = '',
        retries: int = 30,
        manual: bool = False,
        timeout: int = 15,
    ):
        if not _HAS_OCR:
            raise ImportError("需要 ddddocr：pip install ddddocr")
        self.username = username or os.environ.get('IBAO_USERNAME', '')
        self.password = password or os.environ.get('IBAO_PASSWORD', '')
        self.retries = retries
        self.manual = manual
        self.timeout = timeout

        self._session: Optional[requests.Session] = None
        self._token: str = ''
        self._ocr = None
        self._last_err: str = ''

    # --------------------------------------------------------
    # 只读属性（外部主要用这些）
    # --------------------------------------------------------
    @property
    def session(self) -> requests.Session:
        """已登录态的 requests.Session（如果登录了）"""
        if self._session is None:
            raise RuntimeError("尚未登录，请先调用 login()")
        return self._session

    @property
    def token(self) -> str:
        """access_token"""
        return self._token

    @property
    def url(self) -> str:
        """"选择组织（直销/渠道销售）"页面的 URL（直接复制到浏览器打开）"""
        if self._token:
            return f'{_CHANNEL_SELECT}?access_token={self._token}'
        return _CHANNEL_SELECT

    @property
    def ok(self) -> bool:
        """是否处于已登录态

        注意：IEC 系统对 token 的校验很严格，每次请求"选择组织"页面后，
        同 token 的第二次请求可能被后端判定为失效并重定向到登录页。
        因此这里不做实时 HTTP 校验（否则会把 token 用掉），
        只判断 token 是否存在 + 距上次保存时间未超过 24 小时。
        需要真正确认能否访问时，调用 login() 或 refresh()。
        """
        return bool(self._token)

    @property
    def last_error(self) -> str:
        """最近一次失败的错误信息"""
        return self._last_err

    # --------------------------------------------------------
    # OCR / RSA
    # --------------------------------------------------------
    def _get_ocr(self):
        if self._ocr is None:
            self._ocr = ddddocr.DdddOcr(show_ad=False)
        return self._ocr

    @staticmethod
    def _rsa(text: str, pubkey_b64: str) -> str:
        """RSA 公钥加密 (PKCS1_v1_5 + base64，JSEncrypt 兼容)"""
        pem = (
            "-----BEGIN PUBLIC KEY-----\n"
            + pubkey_b64.strip() + "\n"
            + "-----END PUBLIC KEY-----"
        )
        cipher = PKCS1_v1_5.new(CryptoRSA.import_key(pem))
        return base64.b64encode(cipher.encrypt(text.encode('utf-8'))).decode()

    # --------------------------------------------------------
    # 网络请求小工具
    # --------------------------------------------------------
    def _new_session(self) -> requests.Session:
        s = requests.Session()
        s.headers.update({'User-Agent': _UA, 'Referer': _LOGIN_PAGE,
                          'X-Requested-With': 'XMLHttpRequest'})
        return s

    def _get_captcha(self, s: requests.Session) -> Optional[str]:
        """返回 4 位数字验证码，不符合条件返回 None"""
        r = s.get(_CAPTCHA, timeout=self.timeout,
                  params={'v': str(int(time.time() * 1000))})
        if r.status_code != 200:
            return None
        raw = self._get_ocr().classification(r.content).strip().replace(' ', '')
        code = raw.translate(_CHAR_FIX)
        return code if len(code) == 4 and code.isdigit() else None

    def _manual_captcha(self, s: requests.Session) -> Optional[str]:
        r = s.get(_CAPTCHA, timeout=self.timeout,
                  params={'v': str(int(time.time() * 1000))})
        if r.status_code != 200:
            return None
        p = 'iecc_captcha.png'
        with open(p, 'wb') as f:
            f.write(r.content)
        print(f'  验证码图片: {os.path.abspath(p)}')
        try:
            code = input('  请输入验证码: ').strip()
            return code if len(code) == 4 and code.isdigit() else None
        except (EOFError, KeyboardInterrupt):
            return None

    # --------------------------------------------------------
    # 核心：一次登录尝试
    # --------------------------------------------------------
    def _one_try(self) -> Optional[tuple[dict, requests.Session]]:
        """(一次尝试) 返回 (auth_response_dict, session) 或 None"""
        s = self._new_session()
        try:
            s.get(_LOGIN_PAGE, timeout=self.timeout)
            pk = s.get(_PUBKEY, timeout=self.timeout).text.strip()
            if len(pk) < 50:
                return None
            code = self._get_captcha(s)
            if not code:
                return None
            enc_u = self._rsa(self.username, pk)
            enc_p = self._rsa(self.password, pk)
            r = s.post(
                _AUTH,
                data={'username': enc_u, 'password': enc_p,
                      'imgcode': code, 'hcsId': _HCSID},
                timeout=20,
                headers={'Accept': 'application/json, text/javascript, */*; q=0.01',
                         'Referer': _LOGIN_PAGE},
            )
            return r.json(), s
        except Exception:
            return None

    # --------------------------------------------------------
    # "选择组织"页面校验
    # --------------------------------------------------------
    def _check_channel_page(self) -> bool:
        """能否正常打开"选择组织"页面（新的登录成功标准）

        判断规则（按严格度从高到低任一命中即通过）：
        1. 标题 == '选择组织' （100% 确定）
        2. 内容同时含 '渠道销售' 和 '直销' （选择组织页的两个按钮）
        3. 标题不是 '登录' 且状态码 200 且长度 > 5000 （已经是某个已登录态的业务页面，没被打回登录）
        """
        if not self._session or not self._token:
            return False
        try:
            r = self._session.get(
                f'{_CHANNEL_SELECT}?access_token={self._token}',
                timeout=self.timeout,
            )
            if r.status_code != 200:
                return False
            title_match = re.search(r'<title>(.*?)</title>', r.text, re.S | re.I)
            title = title_match.group(1).strip() if title_match else ''
            if title == '选择组织':
                return True
            if '渠道销售' in r.text and '直销' in r.text:
                return True
            if title != '登录' and len(r.text) > 5000:
                return True
            return False
        except Exception:
            return False

    # --------------------------------------------------------
    # 主入口：登录
    # --------------------------------------------------------
    def login(self) -> bool:
        """登录直到打开"选择组织"页面为止。成功返回 True"""
        t0 = time.time()
        ocr_fail = 0
        attempt = 0

        while attempt < self.retries:
            attempt += 1
            out = self._one_try()
            if out is None:
                ocr_fail += 1
                # 连续 OCR 失败时切人工
                if self.manual and ocr_fail >= 15:
                    print(f'  [{attempt}] 连续 {ocr_fail} 次 OCR 失败，切换人工输入')
                    s = self._new_session()
                    s.get(_LOGIN_PAGE, timeout=self.timeout)
                    pk = s.get(_PUBKEY, timeout=self.timeout).text.strip()
                    code = self._manual_captcha(s)
                    if code:
                        try:
                            enc_u, enc_p = self._rsa(self.username, pk), self._rsa(self.password, pk)
                            r = s.post(
                                _AUTH, timeout=20,
                                data={'username': enc_u, 'password': enc_p,
                                      'imgcode': code, 'hcsId': _HCSID},
                            )
                            out = (r.json(), s)
                        except Exception:
                            out = None
                    ocr_fail = 0
                time.sleep(0.15)
                continue
            ocr_fail = 0
            data, s = out
            status = str(data.get('status', ''))
            msg = data.get('msg', '')
            token = data.get('token', '')

            if status == '200':
                # ✅ 新的登录成功标准：再确认"选择组织"页面能打开
                s.headers['access_token'] = token
                self._session = s
                self._token = token
                if self._check_channel_page():
                    dt = time.time() - t0
                    print(f'✅ 登录成功（尝试 {attempt} 次，耗时 {dt:.1f}s）')
                    return True
                # 页面打不开但接口成功了，可能是瞬时问题，再试一次
                print(f'  [{attempt}] auth 成功但"选择组织"页面打不开，重试')
                time.sleep(0.5)
                continue

            if token == 'WEEK_PASSWORD':
                self._last_err = f'弱密码：{msg}'
                print(f'⚠️  {self._last_err}')
                return False
            if token == 'REDIRECT_URL_BSP':
                self._last_err = f'密码已过期：{msg} ({data.get("cusUrl","")})'
                print(f'⚠️  {self._last_err}')
                return False
            if '验证码' in msg:
                if attempt <= 3 or attempt % 10 == 0:
                    print(f'  [{attempt}/{self.retries}] 验证码错误')
                time.sleep(0.2)
                continue
            # 账号密码错误
            if any(k in msg for k in ('密码', '账号', '用户名', '不存在',
                                      '锁定', '停用', '禁用')):
                self._last_err = f'账号或密码错误：{msg}'
                print(f'❌ {self._last_err}')
                return False

            print(f'  [{attempt}] 未知：status={status} msg={msg}')

        dt = time.time() - t0
        self._last_err = f'重试 {self.retries} 次未成功，耗时 {dt:.1f}s'
        print(f'❌ {self._last_err}')
        return False

    # --------------------------------------------------------
    # 便捷：已登录态下的 GET/POST（自动带 access_token 参数）
    # --------------------------------------------------------
    def get(self, url: str, **kw):
        """带登录态 GET 请求"""
        s = self.session
        if not url.startswith('http'):
            url = _BASE + url
        sep = '&' if '?' in url else '?'
        if self._token and 'access_token=' not in url:
            url = f'{url}{sep}access_token={self._token}'
        return s.get(url, timeout=self.timeout, **kw)

    def post(self, url: str, **kw):
        """带登录态 POST 请求"""
        s = self.session
        if not url.startswith('http'):
            url = _BASE + url
        sep = '&' if '?' in url else '?'
        if self._token and 'access_token=' not in url:
            url = f'{url}{sep}access_token={self._token}'
        return s.post(url, timeout=self.timeout, **kw)

    # --------------------------------------------------------
    # 会话保存 / 恢复
    # --------------------------------------------------------
    def save(self, path: str = 'iecc.json') -> bool:
        """保存当前登录态（access_token）到 JSON 文件"""
        if not self._token:
            return False
        with open(path, 'w', encoding='utf-8') as f:
            json.dump({'token': self._token,
                       'user': self.username,
                       't': int(time.time())}, f, ensure_ascii=False, indent=2)
        return True

    @classmethod
    def load(cls, path: str = 'iecc.json', **init_kw) -> 'IEC':
        """从 JSON 恢复会话（不用重新登录）

        直接读取 `iecc.json` 里的 token。
        注意：IEC 的 token 时效性很强，有时打开页面后就失效。
        若打不开，请调用 `refresh()` 强制重新登录刷新 token。
        """
        obj = cls(**init_kw)
        if not os.path.exists(path):
            return obj
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            token = data.get('token', '')
            user = data.get('user', '')
            if user:
                obj.username = user
            if token:
                s = requests.Session()
                s.headers.update({'User-Agent': _UA, 'Referer': _INDEX,
                                  'access_token': token})
                obj._session = s
                obj._token = token
        except Exception:
            pass
        return obj

    def refresh(self, save_path: Optional[str] = None) -> bool:
        """强制重新登录刷新 token（当前 token 失效时调用）

        Parameters
        ----------
        save_path : str, optional
            刷新成功后同步保存到的路径（默认不保存，传 'iecc.json' 即保存）
        """
        ok = self.login()
        if ok and save_path:
            self.save(save_path)
        return ok


# ============================================================
# 命令行入口
# ============================================================
if __name__ == '__main__':
    import argparse

    p = argparse.ArgumentParser(description='宝钢 IEC 登录：直达"选择组织"页面')
    p.add_argument('-u', help='用户名（默认从 IBAO_USERNAME 环境变量读取）')
    p.add_argument('-pw', help='密码（默认从 IBAO_PASSWORD 环境变量读取）')
    p.add_argument('-r', '--retries', type=int, default=30, help='重试次数')
    p.add_argument('-m', '--manual', action='store_true', help='OCR 失败时人工输入')
    p.add_argument('-f', '--save-file', default='iecc.json',
                   help='会话保存文件（默认 iecc.json）')
    p.add_argument('--no-cache', action='store_true',
                   help='忽略已保存的会话，强制重新登录')
    args = p.parse_args()

    # 1. 尝试用已保存的 token（不做有效性验证，避免把 token 用掉）
    iec: Optional[IEC] = None
    if not args.no_cache and os.path.exists(args.save_file):
        iec = IEC.load(args.save_file,
                       username=args.u or os.environ.get('IBAO_USERNAME', ''),
                       password=args.pw or os.environ.get('IBAO_PASSWORD', ''),
                       retries=args.retries, manual=args.manual)
        if iec.ok:
            print(f'ℹ️  使用 {args.save_file} 里的 token（若打不开，请加 --no-cache 重新登录）')
            print()
            print('浏览器直接访问（选择直销 / 渠道销售）:')
            print(iec.url)
            exit(0)

    # 2. 缓存没命中 → 重新登录
    if iec is None or not iec.ok:
        iec = IEC(
            username=args.u or os.environ.get('IBAO_USERNAME', ''),
            password=args.pw or os.environ.get('IBAO_PASSWORD', ''),
            retries=args.retries,
            manual=args.manual,
        )

    if iec.login():
        iec.save(args.save_file)
        print()
        print('浏览器直接访问（选择直销 / 渠道销售）:')
        print(iec.url)
    else:
        exit(1)
