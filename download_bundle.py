"""
宝钢 IEC 系统"准发下载"工具
============================

动作名称：准发下载
功能：从 IEC 系统 → 货物管理 → 准发 → "下载"或"捆包下载"，
按时间段筛选并下载准发数据，输出为 xlsx，并自动上传到飞书云盘。

输出文件命名规则：准发下载_YYMMDD_HHMMSS.xlsx
    例：2026-08-16 14:30:25 运行 → 准发下载_260816_143025.xlsx
    每次下载生成独立文件，不覆盖历史记录。
    每次下载自动上传一份到飞书云盘文件夹（见下方环境变量）。

两种下载模式：
    - download（默认）："下载"按钮 → 订单级别数据（apiBean=IQuasiHairService, methodName=queryQuasiHairDownload）
    - bundle          ："捆包下载"按钮 → 捆包级别明细（apiBean=IQuasiHairDetailService, methodName=downQuasiHairDetail）

前置条件：
    1. 同目录下已有 ibaosteel_client.py
    2. 安装依赖：pip install requests pycryptodome ddddocr pandas openpyxl
    3. 上传飞书云盘需要环境变量（不传则跳过上传）：
         FEISHU_APP_ID        自建应用 App ID（cli_ 开头）
         FEISHU_APP_SECRET    自建应用 App Secret
         FEISHU_FOLDER_TOKEN  目标文件夹 token（分享 URL /folder/ 后那段，
                              默认 QIvpfoJnqlg8IIdT1PXctnB3nne）
       并确保自建应用开通 drive:drive 权限，且被添加为文件夹"可编辑"协作者。

用法 1 - 命令行：
    # 配置飞书凭据（PowerShell，配置后自动上传）
    $env:FEISHU_APP_ID = "cli_xxxxx"
    $env:FEISHU_APP_SECRET = "xxxxx"

    # 直接运行（默认"下载"模式 + 当前月往前 3 个月到往后 1 个月）
    # 输出文件自动命名为 准发下载_260816_1430.xlsx，并上传到飞书
    python download_bundle.py

    # 用"捆包下载"模式（捆包级别明细）
    python download_bundle.py --type bundle

    # 手动指定月份范围
    python download_bundle.py --start 202607 --end 202609

    # 指定输出文件名（覆盖自动命名）
    python download_bundle.py -o my_bundles.xlsx

    # 只下载某个合同号的数据
    python download_bundle.py --contract QE26701121

用法 2 - Python 代码：
    from download_bundle import download_bundles

    df, xlsx_path = download_bundles(
        start='202607', end='202609',
        download_type='download',  # 'download' 或 'bundle'
    )
    print(f'下载 {len(df)} 行数据，保存在 {xlsx_path}')
"""
from __future__ import annotations

import argparse
import datetime
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

from ibaosteel_client import IEC


# ============================================================
# 常量
# ============================================================
_IECS_INDEX = 'https://www.ibaosteel.com/iecs/index'
_QUASI_HAIR_PAGE = 'https://www.ibaosteel.com/iecs/freight/quasiHair/quasiHair/initLoads'
_EXPORT_API = 'https://www.ibaosteel.com/iecs/common/iec/exportExcel'
_DOWNLOAD_API = 'https://www.ibaosteel.com/iecs/common/download'

# 结算用户编号（当前账号固定值，从准发页面隐藏字段 #settleUserNum 获取）
_SETTLE_USER_NUM = '062122'

# 飞书云盘（默认 folder_token = 分享 URL /folder/ 后面那段）
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
DEFAULT_FOLDER_TOKEN = "QIvpfoJnqlg8IIdT1PXctnB3nne"  # https://s2v31ke6sl.feishu.cn/drive/folder/QIvpfoJnqlg8IIdT1PXctnB3nne
_NO_PROXY = {"http": None, "https": None}


# ============================================================
# 飞书云盘上传辅助（与 export_lindiao / export_ouyeel 保持一致，不抽共用模块）
# ============================================================
def feishu_tenant_access_token(app_id: str, app_secret: str) -> str:
    if not app_id or not app_secret:
        raise ValueError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")
    url = f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": app_id, "app_secret": app_secret},
                      timeout=15, proxies=_NO_PROXY)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换 tenant_access_token 失败: {data}")
    token = str(data.get("tenant_access_token") or "")
    if not token:
        raise RuntimeError(f"响应中无 tenant_access_token: {data}")
    print(f"[飞书云盘] tenant_access_token OK (len={len(token)})")
    return token


def feishu_upload_file(token: str, folder_token: str,
                       file_path: str, *, max_size_mb: int = 20) -> dict:
    if not os.path.exists(file_path):
        raise FileNotFoundError(file_path)
    size = os.path.getsize(file_path)
    if size == 0:
        raise ValueError("上传的文件为空")
    if size > max_size_mb * 1024 * 1024:
        raise ValueError(
            f"文件 {size/1024/1024:.1f}MB 超过 upload_all 上限 {max_size_mb}MB")
    if not folder_token:
        raise ValueError("缺少 folder_token")

    filename = os.path.basename(file_path)
    with open(file_path, "rb") as fh:
        file_bytes = fh.read()

    files = {"file": (filename, file_bytes, "application/octet-stream")}
    data = {
        "file_name": filename,
        "parent_type": "explorer",
        "parent_node": folder_token,
        "size": str(size),
    }
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{FEISHU_OPEN_BASE}/drive/v1/files/upload_all"
    r = requests.post(url, data=data, files=files, headers=headers,
                      timeout=120, proxies=_NO_PROXY)
    resp = r.json()
    if resp.get("code") != 0:
        raise RuntimeError(f"上传飞书云盘失败: code={resp.get('code')}, "
                           f"msg={resp.get('msg')}, raw={resp}")
    info = resp.get("data") or {}
    ft = info.get("file_token") or info.get("token") or ""
    name = info.get("name") or filename
    print(f"[飞书云盘] ✅ 上传成功 name={name}  file_token={ft}")
    return info


def auto_upload_to_feishu(xlsx_path: str, *, folder_token_override: str = "") -> Optional[dict]:
    """读取环境变量，把 xlsx 传到飞书云盘。缺少凭据时跳过不报错。"""
    app_id = (os.environ.get("FEISHU_APP_ID") or "").strip()
    app_secret = (os.environ.get("FEISHU_APP_SECRET") or "").strip()
    folder = (folder_token_override or os.environ.get("FEISHU_FOLDER_TOKEN")
              or DEFAULT_FOLDER_TOKEN).strip()
    if not app_id or not app_secret:
        print("[飞书云盘] 未配置 FEISHU_APP_ID / FEISHU_APP_SECRET，跳过上传")
        return None
    try:
        token = feishu_tenant_access_token(app_id, app_secret)
        return feishu_upload_file(token, folder, xlsx_path)
    except Exception as e:
        # 上传失败不影响下载结果，打印错误继续
        print(f"[飞书云盘] ⚠️  上传失败: {e}")
        return None


# ============================================================
# 主函数
# ============================================================
def _auto_filename() -> str:
    """生成默认输出文件名：准发下载_YYMMDD_HHMMSS.xlsx"""
    now = datetime.datetime.now()
    return f"准发下载_{now:%y%m%d_%H%M%S}.xlsx"


def download_bundles(
    start: str,
    end: str,
    out_path: str = '',
    *,
    download_type: str = 'download',
    contract_num: str = '',
    order_num: str = '',
    factory_product_id: str = '',
    shopsign: str = '',
    prod_code_name: str = '',
    machine_id: str = '',
    factory_order_nums: str = '',
    iec: Optional[IEC] = None,
    verbose: bool = True,
    auto_upload: bool = False,
    folder_token: str = '',
) -> Tuple['pd.DataFrame', str]:
    """下载准发数据，可选自动上传到飞书云盘。

    Parameters
    ----------
    start, end : str
        交货期起止月份，格式 'YYYYMM'（如 '202607'）
    out_path : str
        输出 xlsx 路径（留空自动命名）
    download_type : str
        下载模式：'download'（默认，"下载"按钮，订单级别）或
        'bundle'（"捆包下载"按钮，捆包级别明细）
    contract_num, order_num, factory_product_id, shopsign,
    prod_code_name, machine_id, factory_order_nums : str
        可选的筛选条件，留空则下载全部
    iec : IEC, optional
        已登录的 IEC 实例，未传则自动登录
    verbose : bool
        打印详细日志
    auto_upload : bool
        是否下载完后自动上传到飞书云盘（需 FEISHU_APP_ID / FEISHU_APP_SECRET）
    folder_token : str
        上传到的飞书文件夹 token，留空用默认值

    Returns
    -------
    (pd.DataFrame, xlsx_path)
        下载的数据，以及本地 xlsx 文件绝对/相对路径
    """
    if not _HAS_PANDAS:
        raise ImportError("需要 pandas 和 openpyxl：pip install pandas openpyxl")

    # 0. 默认输出文件名（若未指定）
    if not out_path:
        out_path = _auto_filename()

    # 1. 登录（如未传入 iec 实例）
    own_iec = False
    if iec is None:
        iec = IEC()
        if not iec.login():
            raise RuntimeError('IEC 登录失败')
        own_iec = True

    xlsx_path = out_path
    try:
        s = iec.session
        token = iec.token
        referer = f'{_IECS_INDEX}?token={token}'

        # 2. 建立 iecs session（先访问主页 + 准发页）
        if verbose:
            print('→ 建立 iecs 会话...')
        r0 = s.get(f'{_IECS_INDEX}?token={token}', timeout=15,
                   headers={'Referer': referer})
        r0.raise_for_status()
        r1 = s.get(_QUASI_HAIR_PAGE, timeout=15,
                   headers={'Referer': referer})
        r1.raise_for_status()

        # 3. 根据 download_type 设置 apiBean 和 methodName
        if download_type == 'bundle':
            api_bean = 'com.baosight.iecs.freight.quasiHair.api.IQuasiHairDetailService'
            method_name = 'downQuasiHairDetail'
            type_label = '捆包明细'
        else:
            api_bean = 'com.baosight.iecs.freight.quasiHair.api.IQuasiHairService'
            method_name = 'queryQuasiHairDownload'
            type_label = '订单级别'

        # 4. 调用 exportExcel 接口
        param = {
            'contractNum': contract_num,
            'orderNum': order_num,
            'factoryProductId': factory_product_id,
            'shopsign': shopsign,
            'prodCodeName': prod_code_name,
            'prodCode': '',
            'deliveryDateChrStart': start,
            'deliveryDateChrEnd': end,
            'custType': '',
            'memberCodeFlag': '',
            'machineId': machine_id,
            'apiBean': api_bean,
            'methodName': method_name,
            'settleUserNum': _SETTLE_USER_NUM,
            'factoryOrderNums': factory_order_nums,
            'offset': 0,
            'limit': 1000,
        }
        if verbose:
            print(f'→ 请求下载 {start}~{end} 期间的{type_label}数据（{download_type}模式）...')
        r2 = s.post(
            _EXPORT_API,
            data=json.dumps(param),
            headers={
                'Content-Type': 'application/json; charset=utf-8',
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'X-Requested-With': 'XMLHttpRequest',
                'Referer': _QUASI_HAIR_PAGE,
            },
            timeout=120,
        )
        r2.raise_for_status()

        try:
            result = r2.json()
        except Exception:
            raise RuntimeError(f'exportExcel 响应非 JSON: {r2.text[:200]}')

        if result.get('code') != 0:
            raise RuntimeError(f'下载失败：{result.get("msg", "未知错误")}')

        filename = result['msg']
        if verbose:
            print(f'  服务器生成文件: {filename}')

        # 5. 下载实际文件
        dl_url = f'{_DOWNLOAD_API}?fileName={urllib.parse.quote(filename)}&delete=true'
        if verbose:
            print('→ 下载文件...')
        r3 = s.get(dl_url, timeout=120, stream=True,
                   headers={'Referer': _QUASI_HAIR_PAGE})
        r3.raise_for_status()

        # 6. 保存为 xlsx
        xlsx_path = out_path
        if xlsx_path.lower().endswith('.csv'):
            xlsx_path = xlsx_path.rsplit('.', 1)[0] + '.xlsx'
        with open(xlsx_path, 'wb') as f:
            for chunk in r3.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        size = os.path.getsize(xlsx_path)
        if verbose:
            print(f'  xlsx 已保存: {xlsx_path} ({size:,} bytes)')

        # 7. 读取 xlsx 生成 DataFrame（用于返回 + 预览）
        df = pd.read_excel(xlsx_path)
        if verbose:
            print(f'  数据: {len(df)} 行 × {len(df.columns)} 列')

        # 8. 可选：上传到飞书云盘（缺凭据不报错）
        if auto_upload:
            if verbose:
                print('→ 上传飞书云盘...')
            auto_upload_to_feishu(xlsx_path, folder_token_override=folder_token)

        return df, xlsx_path

    finally:
        if own_iec:
            iec.save('iecc.json')


def _default_range() -> Tuple[str, str]:
    """默认日期范围：当前月份往前推 3 个月到往后推 1 个月。

    例：当前 8 月 → start=5 月，end=9 月（'202605' 到 '202609'）。
    跨年自动处理（如当前 2 月 → start=去年 11 月，end=3 月）。
    """
    today = datetime.date.today()
    y, m = today.year, today.month

    start_m = m - 3
    start_y = y
    if start_m <= 0:
        start_m += 12
        start_y -= 1

    end_m = m + 1
    end_y = y
    if end_m > 12:
        end_m -= 12
        end_y += 1

    return f'{start_y:04d}{start_m:02d}', f'{end_y:04d}{end_m:02d}'


def _parse_month(s: str) -> str:
    """校验月份格式，返回 6 位 YYYYMM"""
    s = str(s).strip()
    if len(s) == 6 and s.isdigit():
        y, m = int(s[:4]), int(s[4:6])
        if 2000 <= y <= 2100 and 1 <= m <= 12:
            return s
    raise argparse.ArgumentTypeError(
        f'月份格式错误: {s!r}（应为 YYYYMM，如 202607）')


# ============================================================
# 命令行入口
# ============================================================
if __name__ == '__main__':
    default_start, default_end = _default_range()
    p = argparse.ArgumentParser(
        description='宝钢 IEC 准发数据下载：按时间段下载"下载"或"捆包下载"数据',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f'''
示例:
  # 直接运行（默认"下载"模式 + 当前月往前 3 个月到往后 1 个月，当前即 {default_start}~{default_end}）
  python download_bundle.py

  # 用"捆包下载"模式（捆包级别明细）
  python download_bundle.py --type bundle

  # 指定月份范围
  python download_bundle.py --start 202607 --end 202609

  # 指定输出文件名（覆盖自动命名）
  python download_bundle.py -o my_bundles.xlsx

  # 只下载某个合同号的数据
  python download_bundle.py -s 202607 -e 202609 --contract QE26701121
''')
    p.add_argument('-t', '--type', choices=['download', 'bundle'], default='download',
                   help='下载模式：download="下载"(订单级别，默认) bundle="捆包下载"(捆包级别明细)')
    p.add_argument('-s', '--start', type=_parse_month, default=default_start,
                   help=f'起始月份 YYYYMM（默认 {default_start}，即当前月往前 3 个月）')
    p.add_argument('-e', '--end', type=_parse_month, default=default_end,
                   help=f'结束月份 YYYYMM（默认 {default_end}，即当前月往后 1 个月）')
    p.add_argument('-o', '--out', default='',
                   help='输出 xlsx 路径（默认：准发下载_YYMMDD_HHMM.xlsx，留空自动按时间戳命名）')
    p.add_argument('--contract', default='',
                   help='按销售合同号筛选（可选）')
    p.add_argument('--order', default='',
                   help='按钢厂订单号筛选（可选）')
    p.add_argument('--shopsign', default='',
                   help='按牌号筛选（可选）')
    p.add_argument('--factory-order', default='',
                   help='按工厂订单号筛选（多个用顿号分隔，可选）')
    p.add_argument('--no-upload', action='store_true',
                   help='只下载不上传飞书云盘（默认：配置了 FEISHU_* 环境变量时自动上传）')
    p.add_argument('--folder-token', default='',
                   help=f'飞书文件夹 token（覆盖默认的 {DEFAULT_FOLDER_TOKEN[:10]}…）')
    args = p.parse_args()

    try:
        df, xlsx_path = download_bundles(
            start=args.start,
            end=args.end,
            out_path=args.out,
            download_type=args.type,
            contract_num=args.contract,
            order_num=args.order,
            shopsign=args.shopsign,
            factory_order_nums=args.factory_order,
            auto_upload=not args.no_upload,
            folder_token=args.folder_token,
        )
        print(f'\n✅ 完成：{len(df)} 行数据，文件: {xlsx_path}')
        # 打印前 3 行预览
        if len(df) > 0:
            print('\n前 3 行预览:')
            print(df.head(3).to_string(index=False))
    except Exception as e:
        print(f'\n❌ 失败：{e}', file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)
