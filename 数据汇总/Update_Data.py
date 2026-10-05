#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
懂火 11 合 1 数据汇总同步工作流（纯 JSON 接口 · 零浏览器 · 一张汇总通知卡）

一次运行、一次登录，把 11 个模块全量取回并写入飞书多维表「数据汇总（2026）」：

  ① 出库记录  xiaoshou/m_xiaoshou/xjilulist  → 出库数据            全量替换
  ② 销售明细  xiaoshou/m_dindan/mxlist       → 订单明细            全量替换（动态字段探测）
  ③ 应收汇总  caiwu/m_yinshou/getlist        → 应收汇总            全量替换
  ④ 往来流水  caiwu/m_liushui/getlist        → 往来                全量替换
  ⑤ 客户管理  crm/m_kehu/getlist             → 客户管理            增量 + 已删除标记
  ⑥ 销售订单  xiaoshou/m_dindan/getlist      → 订单数据            全量替换
  ⑦ 采购订单  caigou/m_dindan/getlist        → 采购订单（自动建表）  全量替换
  ⑧ 采购明细  caigou/m_dindan/mxlist         → 采购明细（自动建表）  全量替换
  ⑨ 库存管理  xiaoshou/m_kucun/gl_kucun      → 库存（自动建表）      全量替换
  ⑩ 费用管理  caiwu/m_feiyon/getlist         → 费用管理（自动建表）  全量替换
  ⑪ 服务商    m_load/list_fuwu               → 服务商（自动建表）    全量替换（下拉型接口）

汇总卡片里每行都带该模块自己的「用时（取数 ／ 写入）」。

【为什么不用浏览器了】
取的是后台表格背后的 getlist 接口本身（不是扒 HTML），所以整条链路不需要
Playwright / ddddocr / 点导出按钮 / 下载目录 / 解析伪 xls。2026-10-03 全量实测：
8 模块接口取数合计 ≈217s，比原「导出按钮」混合方案（≈272s）快约 30%，
且更保真——导出会把连续空格压成一个、把 null 渲染成 0，接口返回的是原值。

【字段名一律同名直连，不做任何别名映射】
约定（2026-10-03 与洪确认）：名字不一致时以接口返回的键名为准，去改飞书表里的列名，
而不是在脚本里做 alias 强行映射。飞书侧已按此对齐：
  采购重量(吨)→采购重量、发货状态→销售状态(⑥)、订单重量→销售重量(②)、
  注释→备注(②)、订单 号→订单号(②)
因此本脚本：接口里是什么键、什么值，就照抄进同名字段。
  · 表里有列、接口没这个键 → 该列留空（例：② 的「合同号」，接口不返回）
  · 接口有键、表里没这个列 → 记日志跳过（2026-10-05 已把 ①③④⑤⑥ 缺的 27 列补齐，
    见 数据汇总/add_missing_columns.py；现在基本不该再出现这一类）
  · 接口有键、但有意不同步 → 写进 EXCLUDED_API_KEYS（目前只有 ① 的「捆包号N」）
  · 公式 / 查找 / 人员 / 自动类型 → 飞书不许 API 写，自动跳过

【分页安全闸】
懂火 getlist 单页上限就是 300 条（实测；写 500 会静默丢页）。翻满页数上限仍凑不齐
rtotal 时该模块直接判失败——全量替换是先删后建，拿残缺数据去写比写失败糟糕得多。

【失败隔离】
11 个模块各自独立 try，互不阻断；⑤ 客户是增量模式，只标记删除不真删。

【并发（2026-10-03 提速改造）】
飞书**同一张表不支持并发写**（并发写报 1254291 Write conflict），但**跨表可以并发**。
所以并发粒度是「表」：11 张表各占一个线程、表内严格串行。配套三条：
  · 进池前断言本次要写的 table_id 互不相同（防未来有人让两个模块指向同一张表）；
  · results 只在主线程写，worker 只返回结构化结果、绝不把异常抛给 future；
  · 飞书 POST 一律不挂自动重试（batch_create 若「成功但响应超时」，重试会重复写入）。
另外取数顺序改为 FETCH_ORDER（写入量降序），让最慢的 ① 最早开工。
⚠️ 流水线模式下「取数一个、写一个」，所以取数中途整体崩掉时，先取完的模块可能已经写完。
   这不是新语义——各模块本来就是独立 try、失败的模块不写成功的照写，只是时序更早。

【2026-10-06 效率审查：实测基线 + 已排除的路】
现在的瓶颈是「懂火取数」，且它是**全局串行、压不动**的：11 模块 167 页，
服务端自己就要 ~2s/页。飞书写入是表级并发的，基本被取数窗口盖住 —— 优化写入
的平均速度对总时长几乎无收益，真正的关键路径是「最后取完的那张表的写入」。
以下几条都**真跑过量过**，别再试：
  · PAGE_SIZE 300→500  ❌ 服务端硬上限就是 300：传 400/500/600/1000 一律只回 300 条；
                          limit=500 翻页会静默丢页（7800/12854 且 366 组订单号重复）。
  · 懂火取数并发        ❌ 两个**独立登录**的 session 并发发请求，墙钟 3.92s > 顺序 2.57s
                          —— 服务端全局串行，不是 PHP session 锁，多账号也绕不过。
  · 飞书 search 替代 list ❌ page_token 不推进（加 sort 也一样），无法翻页。
  · 只取 record_id 加速删前扫描（list 的 field_names 参数确实生效，每行只回 1 字段）
                        ⚠️ 但读 ① 全表只从 76.4s → 67.0s：慢在**服务端分页开销**
                           （26 页 × 2.6s），不是数据量。
  · 启动阶段            ❌ 登录 1.0s + token 0.1s + 解析 11 个 table_id 2.0s = 3.1s。
⚠️ 懂火服务端波动极大（同一接口同样参数，实测 ① 140.8s / 268s，⑧ 29.1s / 76s，
   两轮差 1.9~2.6 倍）。**别拿单次运行的数据调参**，也别按单次结果判断优化有没有用。

使用：
  python Update_Data.py [--dry-run] [--no-notify] [--only a,b,...] [--workers N] [--no-pipeline]
  python Update_Data.py --check-fields [--markdown]     # 只读：字段对照表 / 安全闸
  python Update_Data.py --ensure-tables [--dry-run]     # 幂等建 ⑦⑧⑨⑩⑪ 五张新表
回滚：--workers 1 --no-pipeline 等价于改造前的串行行为。

凭据：仓库根目录 .env 里的 DH_USERNAME / DH_PASSWORD + FEISHU_APP_ID / FEISHU_APP_SECRET
"""
import sys, os, json, time, re, math, datetime, argparse, traceback, threading, requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from collections import defaultdict
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ===== stdout/stderr 编码双保险（Windows subprocess 里 print emoji 会崩）=====
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

from dotenv import load_dotenv
load_dotenv(Path(__file__).resolve().parent.parent / ".env", override=False)  # .env 统一放仓库根目录

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from donghuo_login import login_donghuo  # noqa: E402  懂火登录（全程仅此一次，返回 requests.Session）


# ===== 固定配置 =====
DONGHUO_BASE = "https://erpa.donghuo.vip"
BITABLE_APP_TOKEN = "VahHb3YDBaBTwTsCjeAcaAhhnHc"   # 数据汇总（2026）
FEISHU_OPEN_BASE = "https://open.feishu.cn/open-apis"
FEISHU_NOTIFY_UNION_ID = "on_b09bcbf3e74f5d423900aa9b2f00eb63"   # 洪

EXPORT_START_DATE = "2026-01-01"   # ① 出库筛选起始日期
PAGE_SIZE  = 300    # 懂火 getlist 单页上限就是 300（实测），写 500 会静默丢页
# 翻页之间歇多久。2026-10-06 从 0.3 调到 0.1 —— 实测 11 模块共 167 页，
# 每页服务端就要 ~1.8s，请求本身严格串行（上一页响应回来才发下一页），
# 这个 sleep 防的是「并发轰炸」，而这条链上根本没有并发。0.3s × 156 次
# = 46.8s 纯等待，占取数总时长 13.7%。调低是省关键路径上最便宜的一笔。
PAGE_SLEEP = 0.1
MAX_PAGES  = 50     # 分页保险丝：翻满仍凑不齐 rtotal → 该模块判失败
BATCH_SIZE = 500    # 飞书 bitable batch_* 单次上限
CSV_DIR = Path(__file__).parent / "csv_backup"
CSV_DIR.mkdir(exist_ok=True)
MD_PATH = Path(__file__).parent / "十一合一字段对照表.md"


# ===== 字段映射（飞书字段名 == 懂火接口键名，全部同名直连）=====
# 铁律（2026-10-03 与洪确认）：名字不一致时以接口键名为准，改飞书列名；
# 脚本侧不做 alias、不做业务推导，ERP 上是什么就填什么。
FIELD_TYPE_CODE_MAP = {   # 飞书字段类型码 → 写入策略；其余（User/Lookup/Formula/Url/Attachment/自动）不可写
    1: "text", 2: "number", 3: "text", 4: "multiselect", 5: "datetime", 7: "checkbox", 13: "text",
}
_STRATEGY_TO_CREATE_CODE = {"text": 1, "number": 2, "datetime": 5}   # --ensure-tables 建表用

# DateTime 字段建表时要带的 property（与 ①~⑨ 现有日期列逐字一致）
# 不带的话飞书给默认格式（含时分秒），跟其它表长得不一样
_DATE_PROP = {"auto_fill": False, "date_formatter": "yyyy/MM/dd"}


def _number_formatter(name: str) -> str:
    """Number 字段建表时要用的显示精度。

    注意：建表时不显式给 property.formatter，飞书默认给 "0.0"（只显示 1 位小数），
    会在**屏幕上**把懂火的小数四舍五入掉：2.190 显示成 2.2、0.370 显示成 0.4、
    4.136 显示成 4.1。存储值是对的，坏的是显示 —— 用户逐行看每一件货都不对、合计也不对。
    2026-10-04 用户报「库存每一件货的重量都不对、总重也不对」就是这个原因：
    ⑦⑧⑨ 是 --ensure-tables 建的，29 个 Number 字段全部中招（已用
    数据汇总/fix_number_formatter.py 改回，建表源头在这里堵上）。

    口径与 ①~⑥ 人工建的表一致：重量 3 位、金额/单价/税率 2 位、件数/库龄 0 位。
    """
    if "重量" in name:
        return "0.000"
    if name.endswith("数") or name == "库龄":   # 件(张)数 / 可售件数 / 库龄
        return "0"
    return "0.00"

# ① 出库记录（接口 50 键，丢掉 id 和 4 个主动排除键后 45 列，与接口键序一致）
#   2026-10-05 补齐了此前表里缺的 17 列（状态/销售日期/入库日期/锌层/涂料/结构/颜色/
#   最低价/锁定发票/销售其它/销售费用/合计成本/采购单号/米数/目的港/捆包号N/新增时间）。
#   ⚠️ 「米数」是文本不是数字：实测 4955 行非空，值全是交期/仓库/往来方简称这类备注
#      （形如「挂-××」「XXX-×.××」「×.××离港」），不是量 —— ERP 那个列名是历史包袱，
#      装的根本不是数字（⑨库存同样是文本）。真实值不入库，避免把往来方写进公开仓库。
CHUKU_FIELD_TYPES = {
    "订单号": "text", "所属公司": "text", "客户名称": "text",
    "出库日期": "datetime", "销售日期": "datetime", "入库日期": "datetime",
    "品名": "text", "规格": "text", "产地": "text", "等级": "text", "材质": "text",
    "锌层": "text", "涂料": "text", "结构": "text", "颜色": "text", "备注": "text",
    "件(张)数": "number", "重量(吨)": "number",
    "销售人": "text", "销售税率": "number", "销售单价": "number",
    "挂牌价": "number", "销售金额": "number", "锁定发票": "number", "未开发票": "number",
    "销售费用": "number",
    "采购重量": "number", "采购单价": "number", "采购税率": "number", "采购金额": "number",
    "合计成本": "number", "费用金额": "number", "利润": "number", "市场盈利": "number",
    "供应商": "text", "采购单号": "text", "合同号": "text", "提单号": "text",
    "米数": "text", "目的港": "text", "仓库": "text", "车船号": "text", "库位号": "text",
    "捆包号": "text", "新增时间": "datetime",
}

# 主动不写的接口键 —— 既不入表，也不当成「漏列」报警。
#
#   ①「捆包号N」：96 位十六进制哈希（实测 12846 条 96 位、7 条 64 位、1 条 128 位）。
#     同一行隔一次请求取回来，前 32 位相同、后面全变 —— 是 ERP 给页面用的即时签名，
#     不是业务数据。真业务值是它左边那列「捆包号」（字母数字混排、约 13~15 位的批号串）。
#
#   ①「状态」「最低价」「销售其它」+ ③「期初应收」「期初应开发票」：2026-10-05 补列后
#     实测**没有任何信息量**，用户拍板不要（其余 16 列照补）。逐列实测值：
#       · ①状态     12854 行**全部**是「已出库」，不同值只有 1 个
#       · ①最低价   12854 行**全部**是 0.00，不同值只有 1 个
#       · ①销售其它 12850 行是 0.00，只有 4 行有真值
#       · ③期初应收 / ③期初应开发票  各 577 行**全部**是 0.00，不同值只有 1 个
#
#   ⚠️ 全量替换是先删后建，所以飞书表里这几列会变成全空（列本身保留，要删得去飞书手动删）。
EXCLUDED_API_KEYS = {
    "chuku": ["捆包号N", "状态", "最低价", "销售其它"],
    "yingshou": ["期初应收", "期初应开发票"],
}

# ③ 应收汇总（客户级汇总，非订单级；接口 10 键，丢掉 id 和 2 个排除键后 7 列）
# 表里另有「参与」= Lookup(客户表→参与人)：飞书 API 不许新建/修改 Lookup，脚本不碰
# 2026-10-05 补了「期初应收」「期初应开发票」两列，实测全是 0.00 → 同日用户拍板摘掉
YINGSHOU_FIELD_TYPES = {
    "客户名称": "text", "所属公司": "text", "销售人": "text",
    "应收款": "number", "实收款": "number", "可结算": "number", "未收款": "number",
}

# ④ 往来流水（接口 18 键，丢掉 id 后 17 列）
# 2026-10-05 补了「提交日期」「米数」两列（米数 9146 行全空，按 ①⑨ 口径定文本）
WANGLAI_FIELD_TYPES = {
    "所属公司": "text", "日期": "datetime", "提交日期": "datetime",
    "我方帐户": "text", "交易对方": "text", "交易类型": "text", "科目名称": "text",
    "结算方式": "text", "金额": "number", "状态": "text", "结算对方": "text",
    "订单号": "text", "销售人": "text", "米数": "text",
    "备注说明": "text", "提交人": "text", "确认人": "text",
}

# ⑥ 销售订单（接口 21 键，丢掉 id 后 20 列；表里另有提成项目自建的「利润」「市场利润」，
#    接口无来源 → 每次同步会清空，见 README 已知风险）
# 2026-10-05 补了「销售日期」「发货状态」「应结金额」三列
XSDD_FIELD_TYPES = {
    "订单号": "text", "所属公司": "text", "日期": "datetime", "销售日期": "datetime",
    "销售状态": "text", "发货状态": "text", "销售人": "text", "客户名称": "text",
    "订单重量": "number", "订单金额": "number", "实发重量": "number", "实发金额": "number",
    "销售费用": "number", "其它款项": "number", "合同定金": "number", "应结金额": "number",
    "已结金额": "number", "未结金额": "number", "合同未结": "number",
    "新增时间": "datetime",
}

# ⑦ 采购订单（新表，--ensure-tables 自动建）
CGDD_FIELD_TYPES = {
    "订单号": "text", "所属公司": "text", "日期": "datetime", "采购人": "text",
    "供应商": "text", "发货状态": "text", "采购合同号": "text",
    "订单重量": "number", "订单金额": "number", "入库重量": "number", "入库金额": "number",
    "采购费用": "number", "其它款项": "number", "合同定金": "number", "应结金额": "number",
    "已结金额": "number", "未结金额": "number", "合同未结": "number",
    "新增时间": "datetime",
}

# ⑧ 采购明细（新表，--ensure-tables 自动建）
CGMX_FIELD_TYPES = {
    "订单号": "text", "所属公司": "text", "日期": "datetime", "采购人": "text",
    "供应商": "text", "采购合同号": "text", "提单号": "text",
    "品名": "text", "规格": "text", "材质": "text", "产地": "text", "等级": "text",
    "结构": "text", "涂料": "text", "锌层": "text", "颜色": "text",
    "备注": "text", "入库操作": "text",
    "订单重量": "number", "入库重量": "number",
    "采购单价": "number", "采购税率": "number", "采购金额": "number", "入库金额": "number",
}

# ⑨ 库存管理（新表，--ensure-tables 自动建；接口 41 键，丢掉 id 后 40 列）
# ⚠️ 接口必须显式带 sxzhuantai=""（哪怕空串），不传会 PHP Fatal error: Undefined variable
# 主字段用「捆包号」——ERP 库存列表里的库存标识，实测 1101/1101 行非空
# （合同号有 6 行空，不能当主字段；主字段必须是 text）
KUCUN_FIELD_TYPES = {
    "捆包号": "text", "入库日期": "datetime", "所属公司": "text", "库存类型": "text",
    "销售状态": "text", "采购人": "text", "供应商": "text", "货权": "text",
    "品名": "text", "规格": "text", "材质": "text", "产地": "text", "等级": "text",
    "锌层": "text", "涂料": "text", "结构": "text", "颜色": "text",
    "件(张)数": "number", "重量(吨)": "number", "可售重量": "number", "可售件数": "number",
    "销售单价": "number", "最低售价": "number", "锁定人": "text",
    "采购税率": "number", "采购单价": "number", "采购金额": "number",
    "费用金额": "number", "成本单价": "number",
    "合同号": "text", "捆包号N": "text", "提单号": "text", "车船号": "text",
    "米数": "text", "目的港": "text", "仓库": "text", "库位号": "text",
    "库龄": "number", "备注": "text", "新增时间": "datetime",
}

# ⑩ 费用管理（新表，--ensure-tables 自动建；接口 24 键）
# ⚠️ start_time 与 end_time 必须**成对**传：只传 start_time 等于没传（仍返回全量 11162 行），
#    两个都传才生效 → 3068 行（2026-01-01 起）。实测 11 页 / 11 秒。
# 「米数」在接口里有、页面上根本不显示，实测 3068 行全空；按 ①⑨ 口径定文本。
# 「发票号」必须是文本：值是 19~20 位数字串，甚至有一行是两个号用「，」连写，数字类型会被浮点截精度。
FEIYONG_FIELD_TYPES = {
    "id": "text", "订单号": "text", "编号": "text", "米数": "text",
    "所属公司": "text", "服务商名称": "text", "销售人": "text", "科目名称": "text",
    "计价方式": "text", "税率": "number", "费用单价": "number", "费用金额": "number",
    "计价重量": "number", "计算金额": "number", "日期": "datetime", "备注说明": "text",
    "费用类型": "text", "审核状态": "text", "付款状态": "text", "付款日期": "datetime",
    "发票状态": "text", "发票日期": "datetime", "发票号": "text", "结算人": "text",
}

# ⑪ 服务商（新表，--ensure-tables 自动建）
# 来源不是 getlist，而是「设置 → 服务商」页面的下拉数据源 m_load/list_fuwu：
# 一次性返回裸数组 [{"key": 名称, "value": 名称}]，无分页、无 rtotal（走 list_api 分支）。
# 实测 271 条，key==value 全同、无空值、无重复、无前后空格。
# ⚠️ 完整页面 system/v_fuwu 上还有 地址/电话/备注 三列，但其接口 system/m_fuwu/getlist
#    对当前脚本所用的懂火账号（见 .env 的 DH_USERNAME）返回「没有权限」；
#    用户 2026-10-05 决定就用下拉这份。
FUWUSHANG_FIELD_TYPES = {"服务商名称": "text"}

# ⑤ 客户：主键=客户名称；跟踪字段；仅新增写创建时间；已删除标记
# 「所属公司」2026-10-05 加入跟踪字段（走 kehu_diff 的文本比对）
# 时间类字段单列：值要转毫秒时间戳，不能当文本比，也参与增量比对（首轮会把空列补齐）
KEHU_TRACKED_FIELDS = [
    "客户类型", "所属人", "联系人", "联系人职位", "固定电话", "移动电话",
    "邮箱地址", "所属省份", "联系地址", "主营产品", "采购产品", "备注", "所属公司",
]
KEHU_TIME_FIELDS = ["新增时间", "最后更新"]
DELETED_MARK = "已删除"
_JUNK_RE = re.compile(r"^\d{1,2}$")   # '1'/'0'/'00' 等占位垃圾值


# ===== 十一合一注册表（加模块只改这一处）=====
# api     : /model/admin/ 之后的路径
# page    : 该表在懂火后台的页面地址，用作 Referer（懂火要求带）
# params  : 必须显式传的筛选参数——少一个就返回子集或非 JSON（实测）
# table   : 飞书表 id；写 ensure 的表示按表名自动定位/建表
# dynamic : 写入时实时探测飞书字段类型（② 订单明细用；其余用上面的硬编码映射）
PART_ORDER = ["chuku", "dingdan", "yingshou", "wanglai", "kehu", "xsdd", "cgdd", "cgmx", "kucun",
              "feiyong", "fuwushang"]


def _today() -> str:
    return datetime.date.today().strftime("%Y-%m-%d")


PARTS = {
    "chuku": {
        "name": "① 出库记录", "table": "tblolnj06JZkYNiU",
        "api": "xiaoshou/m_xiaoshou/xjilulist", "page": "xiaoshou/v_xjlall",
        "params": lambda: {"start_time": EXPORT_START_DATE, "end_time": _today()},
        "fields": CHUKU_FIELD_TYPES,
    },
    "dingdan": {
        "name": "② 销售明细", "table": "tblcEZoQatk7lCAO",
        "api": "xiaoshou/m_dindan/mxlist", "page": "xiaoshou/v_xmxhz",
        "params": lambda: {}, "dynamic": True,
    },
    "yingshou": {
        "name": "③ 应收汇总", "table": "tblpjne9dIuif5HD",
        # ⚠️ 模块名是 m_yinshou（不是 yingshou），写错会返回「定义的模块不存在」
        "api": "caiwu/m_yinshou/getlist", "page": "caiwu/v_x_yinshou",
        "params": lambda: {},
        "fields": YINGSHOU_FIELD_TYPES,
    },
    "wanglai": {
        "name": "④ 往来流水", "table": "tblbS1dPaDVL3GY8",
        "api": "caiwu/m_liushui/getlist", "page": "caiwu/v_x_jiesuan",
        # 2026-10-06 加时间过滤，与 ① 出库同口径（用户拍板）。
        # 不加时是全量历史 9146 行（含 2000/2022/2023/2024/2025 共 6547 行），
        # 加 start_time=2026-01-01 后 2599 行 —— 实测逐条核对「日期」字段
        # 100% 落在 2026 年，无泄漏。取数 74s→39s，且 ④ 正是全流程关键路径上
        # 最长的那条尾巴（取完 267s + 写入 131s = 398s，比取数总时长还长）。
        # ⚠️ 全量替换是先删后建：改这里等于把 2026 年以前的往来流水从飞书表里清掉。
        "params": lambda: {"start_time": EXPORT_START_DATE, "end_time": _today()},
        "fields": WANGLAI_FIELD_TYPES,
    },
    "kehu": {
        "name": "⑤ 客户管理", "table": "tblCE7zIWs804RR5",
        "api": "crm/m_kehu/getlist", "page": "crm/v_kehu",
        "params": lambda: {}, "incremental": True,
    },
    "xsdd": {
        "name": "⑥ 销售订单", "table": "tblJZIyNXoe8PZer",
        "api": "xiaoshou/m_dindan/getlist", "page": "xiaoshou/v_dindan",
        # fhzhuantai / shkzhuantai 必须显式传空串，否则只拿子集
        "params": lambda: {"fhzhuantai": "", "shkzhuantai": ""},
        "fields": XSDD_FIELD_TYPES,
    },
    "cgdd": {
        "name": "⑦ 采购订单", "table": None, "ensure": "采购订单",
        "api": "caigou/m_dindan/getlist", "page": "caigou/v_dindan",
        "params": lambda: {}, "fields": CGDD_FIELD_TYPES,
    },
    "cgmx": {
        "name": "⑧ 采购明细", "table": None, "ensure": "采购明细",
        "api": "caigou/m_dindan/mxlist", "page": "caigou/v_dindan",
        "params": lambda: {}, "fields": CGMX_FIELD_TYPES,
    },
    "kucun": {
        "name": "⑨ 库存管理", "table": None, "ensure": "库存",
        # ⚠️ sxzhuantai 必须显式传（哪怕空串），不传会 PHP Fatal error: Undefined variable
        "api": "xiaoshou/m_kucun/gl_kucun", "page": "xiaoshou/v_kucun_gl",
        "params": lambda: {"sxzhuantai": ""}, "fields": KUCUN_FIELD_TYPES,
    },
    "feiyong": {
        "name": "⑩ 费用管理", "table": None, "ensure": "费用管理",
        # ⚠️ start_time 必须与 end_time 成对传，只传一个会被忽略（仍返回全量 11162 行）
        "api": "caiwu/m_feiyon/getlist", "page": "caiwu/v_feiyon",
        "params": lambda: {"start_time": EXPORT_START_DATE, "end_time": _today()},
        "fields": FEIYONG_FIELD_TYPES,
    },
    "fuwushang": {
        "name": "⑪ 服务商", "table": None, "ensure": "服务商",
        # 下拉型接口：返回裸数组 [{"key","value"}]，不是 {root, rtotal} 分页形状 → list_api 分支
        "api": "m_load/list_fuwu", "page": "system/v_fuwu",
        "params": lambda: {}, "list_api": True, "list_key": "服务商名称",
        "fields": FUWUSHANG_FIELD_TYPES,
    },
}

PART_NAMES = {k: v["name"] for k, v in PARTS.items()}
TABLES = {k: v["table"] for k, v in PARTS.items() if v.get("table")}
_ENSURED = {}   # 运行时按表名解析出来的 table_id 缓存：{part: table_id}

# 取数顺序 = 写入量降序。流水线模式下每张表「取数一结束就开始写」，
# 让行数最大的 ① 最早开工，整条链最短（卡片与 results 顺序仍用 PART_ORDER）。
FETCH_ORDER = ["chuku", "dingdan", "wanglai", "cgmx", "feiyong", "cgdd", "xsdd", "kucun",
               "yingshou", "kehu", "fuwushang"]
assert set(FETCH_ORDER) == set(PART_ORDER), "FETCH_ORDER 与 PART_ORDER 不一致（加模块时漏了？）"

DEFAULT_WORKERS = 6   # 并发写表的线程数；① 出库（≈162s）是硬下界，4 个已到底，6 是到达不齐的余量

# ============ 工具函数 ============

_LOG_LOCK = threading.Lock()


def log(msg: str):
    # 单线程时行为与不加锁完全一致；并发写表时保证整行日志不被别的线程劈开
    with _LOG_LOCK:
        print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def env(name: str) -> str:
    return os.environ.get(name, "").strip()


def norm_text(v) -> str:
    """文本归一化：None/空白/全角空格 → ''（⑤ 客户用）"""
    if v is None:
        return ""
    return str(v).replace("\u3000", " ").strip()


def is_meaningful(v) -> bool:
    """值是否有意义：非空且非 1~2 位纯数字占位垃圾（⑤ 客户用）"""
    s = norm_text(v)
    return bool(s) and not _JUNK_RE.match(s)


def to_datetime_ms(raw):
    """'2026-09-02 13:18:43' → 毫秒时间戳（⑤ 客户创建时间用）"""
    s = norm_text(raw)
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
        try:
            return int(datetime.datetime.strptime(s, fmt).timestamp() * 1000)
        except ValueError:
            continue
    return None


def feishu_token() -> str:
    """获取飞书 tenant_access_token（每次现取，2 小时有效期）"""
    app_id = os.environ.get("FEISHU_APP_ID") or os.environ.get("FEISHU_NOTIFY_APP_ID", "").strip()
    app_secret = os.environ.get("FEISHU_APP_SECRET") or os.environ.get("FEISHU_NOTIFY_APP_SECRET", "").strip()
    if not app_id or not app_secret:
        raise RuntimeError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET 环境变量")
    r = requests.post(
        f"{FEISHU_OPEN_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret}, timeout=15,
    )
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"换 tenant_access_token 失败: {data}")
    return data["tenant_access_token"]


def feishu_send_card(union_id: str, card: dict, token: str):
    """通过飞书机器人给指定用户发卡片消息"""
    url = f"{FEISHU_OPEN_BASE}/im/v1/messages?receive_id_type=union_id"
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"receive_id": union_id, "msg_type": "interactive",
            "content": json.dumps(card, ensure_ascii=False)}
    r = requests.post(url, headers=h, json=body, timeout=15)
    data = r.json()
    if data.get("code") != 0:
        log(f"[飞书通知] ❌ 发送失败: code={data.get('code')} msg={data.get('msg')}")
    else:
        log(f"[飞书通知] ✅ 汇总卡片已发送给 {union_id}")


# ============ 懂火 getlist：八模块统一分页拉取 ============

def _getlist_url(part: str) -> str:
    return f"{DONGHUO_BASE}/model/admin/{PARTS[part]['api']}"


def _getlist_headers(part: str) -> dict:
    return {"X-Requested-With": "XMLHttpRequest",
            "Referer": f"{DONGHUO_BASE}/view/admin/{PARTS[part]['page']}"}


def _fetch_list_api(session, part: str) -> list:
    """下拉型接口（⑪ 服务商 /m_load/list_fuwu）：一次性返回裸数组，无分页。

    形状是 [{"key": 名称, "value": 名称}]，不是 getlist 的 {root, rtotal}，
    所以不能走 fetch_all 的翻页逻辑。取 value 落成 [{飞书列名: 值}]，
    之后跟其它模块共用同一条「CSV → 预检 → 清空 → 批量写」链路。
    """
    spec, name = PARTS[part], PARTS[part]["name"]
    col = spec["list_key"]
    r = session.post(_getlist_url(part), data={}, headers=_getlist_headers(part), timeout=60)
    try:
        data = r.json()
    except ValueError:
        raise RuntimeError(f"{name}：返回非 JSON（登录态可能失效）: {r.text[:150]}")
    if not isinstance(data, list):
        raise RuntimeError(f"{name}：期望裸数组，实际 {type(data).__name__}: {str(data)[:150]}")
    rows = []
    for it in data:
        v = norm_text((it or {}).get("value") if isinstance(it, dict) else it)
        if v:
            rows.append({col: v})
    log(f"[{name}] ✅ 拉取完成，共 {len(rows)} 条（原始 {len(data)} 条，空值已丢 {len(data) - len(rows)} 条）")
    return rows


def fetch_all(session, part: str) -> list:
    """按模块从懂火 getlist 接口顺序翻页拉全量（后台页面表格用的就是这些接口）。

    ⚠️ limit 上限是 300（2026-10-03 实测），写 500 会丢页。
    ⚠️ 翻满 max_pages 仍没到 rtotal 直接 raise —— 全量替换是先删后建，
       拿残缺数据去写，比直接失败糟糕得多。
    """
    spec, name = PARTS[part], PARTS[part]["name"]
    if spec.get("list_api"):
        return _fetch_list_api(session, part)
    max_pages = spec.get("max_pages", MAX_PAGES)
    params = spec["params"]()
    all_rows, rtotal = [], None
    for page_no in range(1, max_pages + 1):
        r = session.post(_getlist_url(part), data={"page": page_no, "limit": PAGE_SIZE, **params},
                         headers=_getlist_headers(part), timeout=60)
        try:
            data = r.json()
        except ValueError:
            raise RuntimeError(f"{name}：第 {page_no} 页返回非 JSON（登录态可能失效）: {r.text[:150]}")
        root = data.get("root") or []
        if rtotal is None:
            rtotal = int(data.get("rtotal") or 0)
            log(f"[{name}] 总数 rtotal={rtotal}，每页 {PAGE_SIZE}，预计 {math.ceil(rtotal / PAGE_SIZE)} 页")
        all_rows.extend(root)
        if page_no % 5 == 0 or not root or (rtotal and len(all_rows) >= rtotal):
            log(f"[{name}] 第 {page_no} 页: 累计 {len(all_rows)}/{rtotal}")
        if not root or (rtotal and len(all_rows) >= rtotal):
            break
        time.sleep(PAGE_SLEEP)
    if rtotal is not None and len(all_rows) < rtotal:
        raise RuntimeError(f"{name}：只拉到 {len(all_rows)}/{rtotal} 条（翻满上限 {max_pages} 页）"
                           f"—— 全量替换先删后建，拒绝写入残缺数据")
    log(f"[{name}] ✅ 拉取完成，共 {len(all_rows)} 条")
    if part == "kehu":
        seen = defaultdict(int)
        for row in all_rows:
            n = norm_text(row.get("客户名称"))
            if n:
                seen[n] += 1
        dups = {n: c for n, c in seen.items() if c > 1}
        if dups:
            log(f"[{name}] ⚠️ 名称重名 {len(dups)} 组（增量合并时按数据丰富度处理）")
        log(f"[{name}] 唯一名称 {len(seen)} 个")
    return all_rows


def probe_json_keys(session, part: str) -> list:
    """只拉第 1 页，取接口返回的键名（供 --check-fields 用，不落盘）"""
    spec = PARTS[part]
    if spec.get("list_api"):
        # 下拉型接口没有行键，只有 key/value 两个字段；入表的是 list_key 那一列
        return [spec["list_key"]]
    r = session.post(_getlist_url(part), data={"page": 1, "limit": PAGE_SIZE, **spec["params"]()},
                     headers=_getlist_headers(part), timeout=60)
    root = (r.json().get("root") or [])
    return list(root[0].keys()) if root else []



# ============ CSV 备份（每部分一份，写入成功后自动删除）============

def rows_to_csv(rows: list, csv_prefix: str) -> Path:
    """④⑤ API 原始 dict 列表 → CSV 备份"""
    import pandas as pd
    df = pd.DataFrame(rows).dropna(how="all")
    now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = CSV_DIR / f"{csv_prefix}_export_{now}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    log(f"[备份] CSV 已保存: {csv_path.name}（{len(df)} 行 × {len(df.columns)} 列）")
    return csv_path
# ============ 飞书多维表通用（table_id 参数化）============

# HTTP 连接复用：每个线程各持一个 Session（requests.Session 本身不是线程安全的，
# 用 threading.local 让并发写表的线程各自复用自己的 TCP 连接，仍能省掉反复 TLS 握手）。
#
# ⚠️ POST 绝不能挂 status/read 重试：batch_create 若「实际成功但响应超时」，
#    自动重试会造成重复写入（飞书侧没有幂等键）。只有 GET（读记录 / 读字段 / 列表）
#    是幂等的，才允许自动重试。connect 类重试对 POST 也生效，但连接失败 = 请求从未发出，
#    属于安全重试，故保留。
_TLS = threading.local()


def _fs_session() -> requests.Session:
    s = getattr(_TLS, "session", None)
    if s is None:
        s = requests.Session()
        retry = Retry(total=3, connect=2, read=2, status=2, backoff_factor=0.5,
                      status_forcelist=[429, 500, 502, 503, 504],
                      allowed_methods=frozenset(["GET"]))
        adapter = HTTPAdapter(max_retries=retry, pool_connections=4, pool_maxsize=8)
        s.mount("http://", adapter)
        s.mount("https://", adapter)
        _TLS.session = s
    return s


# 飞书「偶发」业务错误码：HTTP 200，错误码藏在 body 的 code 里 —— 上面 urllib3 那层
# Retry 只看 HTTP 状态码，拦不住它们，必须在这里自己退避重试。
#   1254607 Data not ready, please try again later（表/索引正在重建，官方原话就是「稍后重试」）
#   1254291 Write conflict（并发写期间的读也可能撞上）
# ⚠️ 只有 GET 能用它 —— 读记录 / 读字段 / 列表 都是幂等的。
#    POST 绝不能这样重试：batch_create 若「实际成功但响应超时」，重试会重复写入。
_TRANSIENT_READ_CODES = {1254607, 1254291}


def _fs_get(url: str, headers: dict, timeout: int = 60, tries: int = 4) -> dict:
    """飞书 GET：幂等，遇偶发业务错误自动退避重试，返回解析后的 JSON。"""
    delay = 2.0
    for attempt in range(1, tries + 1):
        d = _fs_session().get(url, headers=headers, timeout=timeout).json()
        code = d.get("code")
        if code == 0 or code not in _TRANSIENT_READ_CODES:
            return d
        if attempt == tries:
            log(f"[飞书] ❌ 读取偶发错误重试 {tries} 次仍失败: {code} {d.get('msg')}")
            return d
        log(f"[飞书] ⏳ 读取偶发错误 {code} {d.get('msg')} —— {delay:.0f}s 后重试（第 {attempt}/{tries - 1} 次）")
        time.sleep(delay)
        delay *= 2
    return {}


def _records_url(table_id: str) -> str:
    return f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{table_id}/records"


def bitable_list_records(token: str, table_id: str) -> list:
    """读取全部记录 [{"record_id", "fields"}, ...]（⚠️ search 接口 page_token 不推进，必须 GET list）"""
    h = {"Authorization": f"Bearer {token}"}
    out, seen, page_token = [], set(), None
    while True:
        qs = f"page_size={BATCH_SIZE}" + (f"&page_token={page_token}" if page_token else "")
        data = _fs_get(f"{_records_url(table_id)}?{qs}", h)
        if data.get("code") != 0:
            raise RuntimeError(f"list records 失败({table_id}): {data}")
        d = data.get("data") or {}
        items = d.get("items") or []
        if not items:
            break
        dup = 0
        for it in items:
            rid = it.get("record_id")
            if rid and rid not in seen:
                out.append({"record_id": rid, "fields": it.get("fields") or {}})
                seen.add(rid)
            else:
                dup += 1
        if dup == len(items):
            log(f"[飞书] ⚠️ page_token 未推进，停止；累计 {len(out)} 条")
            break
        if not d.get("has_more") or not d.get("page_token"):
            break
        page_token = d.get("page_token")
    return out


def bitable_batch_delete(token: str, table_id: str, record_ids: list):
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{_records_url(table_id)}/batch_delete?ignore_consistency_check=true"
    total = len(record_ids)
    for i in range(0, total, BATCH_SIZE):
        batch = record_ids[i:i + BATCH_SIZE]
        r = _fs_session().post(url, headers=h, json={"records": batch}, timeout=90)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_delete 失败: {data.get('msg')}")
        if (i // BATCH_SIZE + 1) % 10 == 0 or i + BATCH_SIZE >= total:
            log(f"[飞书] delete: 批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE}")


def bitable_batch_create(token: str, table_id: str, records: list) -> int:
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{_records_url(table_id)}/batch_create?ignore_consistency_check=true"
    total, created = len(records), 0
    for i in range(0, total, BATCH_SIZE):
        batch = records[i:i + BATCH_SIZE]
        r = _fs_session().post(url, headers=h, json={"records": batch}, timeout=90)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_create 失败 批 {i // BATCH_SIZE + 1}: {data.get('code')} {data.get('msg')} 样本={str(data)[:300]}")
        created += len((data.get("data") or {}).get("records") or [])
        log(f"[飞书] create: {len(batch)} 条 (批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE})，累计 {created}")
    return created


def bitable_batch_update(token: str, table_id: str, updates: list) -> int:
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{_records_url(table_id)}/batch_update?ignore_consistency_check=true"
    total, done = len(updates), 0
    for i in range(0, total, BATCH_SIZE):
        batch = updates[i:i + BATCH_SIZE]
        r = _fs_session().post(url, headers=h, json={"records": batch}, timeout=90)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"batch_update 失败 批 {i // BATCH_SIZE + 1}: {data.get('code')} {data.get('msg')}")
        done += len((data.get("data") or {}).get("records") or [])
        if (i // BATCH_SIZE + 1) % 10 == 0 or i + BATCH_SIZE >= total:
            log(f"[飞书] update: 批 {i // BATCH_SIZE + 1}/{(total + BATCH_SIZE - 1) // BATCH_SIZE}，累计 {done}")
    return done


def bitable_get_field_types(token: str, table_id: str) -> dict:
    """动态探测字段类型（② 订单表用），返回 {字段名: 写入策略}"""
    h = {"Authorization": f"Bearer {token}"}
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{table_id}/fields"
    out, page_token = {}, None
    while True:
        qs = "page_size=100" + (f"&page_token={page_token}" if page_token else "")
        d = _fs_get(f"{url}?{qs}", h)
        if d.get("code") != 0:
            raise RuntimeError(f"获取字段列表失败: {d}")
        dd = d.get("data") or {}
        for f in dd.get("items") or []:
            out[f["field_name"]] = FIELD_TYPE_CODE_MAP.get(f["type"], "unsupported")
        if not dd.get("has_more") or not dd.get("page_token"):
            break
        page_token = dd.get("page_token")
    return out


def convert_value(raw, ftype: str):
    """DataFrame 单元格值 → 飞书字段值（② 订单增强版：千分位/多选/复选/时间戳兜底，兼容其余部分）"""
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return None
    if ftype == "text":
        s = str(raw).strip()
        return s if s else None
    if ftype == "number":
        s = str(raw).strip()
        if not s:
            return None
        try:
            f = float(s.replace(",", ""))
            return int(f) if f == int(f) else f
        except ValueError:
            return None
    if ftype == "datetime":
        if isinstance(raw, (int, float)) and not isinstance(raw, bool):
            return int(raw) if raw > 10**12 else int(raw * 1000)
        s = str(raw).strip()
        if not s:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                    "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
            try:
                return int(datetime.datetime.strptime(s, fmt).timestamp() * 1000)
            except ValueError:
                continue
        return None
    if ftype == "multiselect":
        s = str(raw).strip()
        if not s:
            return None
        parts = [p.strip() for p in re.split(r"[,，、;；]", s)]
        return [p for p in parts if p] or None
    if ftype == "checkbox":
        s = str(raw).strip().lower()
        if s in ("是", "true", "1", "y", "yes", "√", "已"):
            return True
        if s in ("否", "false", "0", "n", "no", ""):
            return False
        return None
    return None


def rows_to_records(rows: list, field_types: dict, tag: str = "", exclude=()) -> list:
    """懂火接口 dict 列表 → 飞书 batch_create 记录列表。

    外层迭代「飞书字段名」，直接取同名接口键——无别名、无推导。
    空值不写（None 会被飞书忽略，等于保留表里原值）。
    """
    writable = [c for c, t in field_types.items() if t != "unsupported"]
    unsupported = [c for c, t in field_types.items() if t == "unsupported"]
    if unsupported:
        log(f"[{tag}] 不可写字段（公式/查找/人员/自动）跳过: {', '.join(unsupported)}")

    json_keys = set()
    for row in rows[:50]:
        json_keys.update(row.keys())
    missing = [c for c in writable if c not in json_keys]
    if missing:
        log(f"[{tag}] ⚠️ 表里有列但接口无此键，该列将留空: {', '.join(missing)}")
    unused = sorted(json_keys - set(writable) - {"id"} - set(exclude))
    if unused:
        log(f"[{tag}] 接口有键但表里没列，不入表: {', '.join(unused)}")
    hit = sorted(json_keys & set(exclude))
    if hit:
        log(f"[{tag}] 🚫 主动排除，不入表: {', '.join(hit)}")

    records = []
    for row in rows:
        fields = {}
        for col in writable:
            val = convert_value(row.get(col), field_types[col])
            if val is not None:
                fields[col] = val
        records.append({"fields": fields})
    return records

# ============ ⑤ 客户：增量计划（与单脚本一致）============

def merge_duplicate_rows(rows: list) -> dict:
    """懂火同名客户多行 → 合并：基础行=有效值最多行；空字段用其他行有效值补；新增时间取最早"""
    groups = defaultdict(list)
    for row in rows:
        name = norm_text(row.get("客户名称"))
        if name:
            groups[name].append(row)
    merged = {}
    for name, grp in groups.items():
        if len(grp) == 1:
            merged[name] = grp[0]
            continue

        def score(r):
            return sum(1 for c in KEHU_TRACKED_FIELDS if is_meaningful(r.get(c)))

        base = max(grp, key=score)
        m = dict(base)
        for c in KEHU_TRACKED_FIELDS:
            if not is_meaningful(m.get(c)):
                for r in grp:
                    if r is not base and is_meaningful(r.get(c)):
                        m[c] = r[c]
                        break
        ts_list = [t for t in (to_datetime_ms(r.get("新增时间")) for r in grp) if t]
        if ts_list:
            m["新增时间"] = datetime.datetime.fromtimestamp(min(ts_list) / 1000).strftime("%Y-%m-%d %H:%M:%S")
        log(f"[⑤ 客户] 重名合并: {name}（{len(grp)} 行 → 基础行 id={base.get('id')}）")
        merged[name] = m
    return merged


def kehu_build_new_fields(row: dict) -> dict:
    """懂火客户行 → 新增记录 fields（垃圾值不写入）"""
    fields = {"客户名称": norm_text(row.get("客户名称"))}
    for col in KEHU_TRACKED_FIELDS:
        v = norm_text(row.get(col))
        if is_meaningful(v):
            fields[col] = v
    ts = to_datetime_ms(row.get("新增时间"))
    if ts:
        fields["创建时间"] = ts
    for col in KEHU_TIME_FIELDS:
        ms = to_datetime_ms(row.get(col))
        if ms:
            fields[col] = ms
    return fields


def kehu_diff(row: dict, feishu_fields: dict) -> dict:
    """比对跟踪字段，返回需要更新的 fields（只写有意义且不同的值）"""
    changed = {}
    for col in KEHU_TRACKED_FIELDS:
        new_v = norm_text(row.get(col))
        old_v = norm_text(feishu_fields.get(col))
        if is_meaningful(new_v) and new_v != old_v:
            changed[col] = new_v
    # 时间类：飞书存的是毫秒时间戳，得跟毫秒比（别拿字符串去比，会永远判不等）
    for col in KEHU_TIME_FIELDS:
        ms = to_datetime_ms(row.get(col))
        if ms and ms != feishu_fields.get(col):
            changed[col] = ms
    return changed


def kehu_plan(donghuo_rows: list, feishu_records: list) -> dict:
    """以客户名称为主键生成增量计划（新增/更新/恢复/标记删除/无变化）"""
    dh_map = merge_duplicate_rows(donghuo_rows)
    fs_map = defaultdict(list)
    no_name = 0
    for rec in feishu_records:
        name = norm_text((rec["fields"] or {}).get("客户名称"))
        if not name:
            no_name += 1
            continue
        fs_map[name].append(rec)
    if no_name:
        log(f"[⑤ 客户] ⚠️ 飞书 {no_name} 条记录无客户名称，跳过比对")

    to_create, to_update, to_restore = [], [], []
    unchanged = 0
    for name, row in dh_map.items():
        matches = fs_map.get(name)
        if not matches:
            to_create.append(kehu_build_new_fields(row))
            continue
        for rec in matches:
            changed = kehu_diff(row, rec["fields"] or {})
            was_deleted = bool(norm_text((rec["fields"] or {}).get("AI 提示")))
            if was_deleted:
                fields = dict(changed)
                fields["AI 提示"] = ""
                to_restore.append({"record_id": rec["record_id"], "fields": fields})
            elif changed:
                to_update.append({"record_id": rec["record_id"], "fields": changed})
            else:
                unchanged += 1

    to_mark_deleted = []
    for name, recs in fs_map.items():
        if name in dh_map:
            continue
        for rec in recs:
            if norm_text((rec["fields"] or {}).get("AI 提示")):
                continue
            to_mark_deleted.append({"record_id": rec["record_id"], "name": name})

    return {"to_create": to_create, "to_update": to_update, "to_restore": to_restore,
            "to_mark_deleted": to_mark_deleted, "unchanged": unchanged,
            "total_donghuo": len(dh_map), "total_feishu": len(feishu_records)}

# ============ 各部分飞书写入（返回 stats dict；失败 raise）============

def run_full_replace(token: str, part: str, payload: dict) -> dict:
    """全量替换：预检 → 清空 → 写入（①②③④⑥⑦⑧⑨⑩⑪ 共用；② 走动态字段探测）

    顺序刻意是「先预检、后清空」：字段映射若有问题，5 条预检就会失败并抛出，
    此时表里原数据一条没动；而不是删光了才发现写不进去（那才是真正的灾难）。
    """
    spec, name = PARTS[part], PARTS[part]["name"]
    table_id = table_id_of(part)

    if spec.get("dynamic"):
        field_types = bitable_get_field_types(token, table_id)
        writable = [k for k, v in field_types.items() if v != "unsupported"]
        log(f"[{name}] 动态字段探测: {len(field_types)} 个字段，可写 {len(writable)} 个")
    else:
        field_types = spec["fields"]

    records = rows_to_records(payload["rows"], field_types, name,
                              exclude=EXCLUDED_API_KEYS.get(part, []))
    if not records:
        raise RuntimeError(f"{name}：待写入记录 0 条，拒绝清空现有数据")

    # 预检：先拿 5 条试写，验证字段名/类型/选项都通
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    test_batch = records[:min(5, len(records))]
    d = _fs_session().post(f"{_records_url(table_id)}/batch_create", headers=h,
                           json={"records": test_batch}, timeout=90).json()
    if d.get("code") != 0:
        raise RuntimeError(f"{name}：预检写入失败（表内原数据未动）: "
                           f"{d.get('code')} {d.get('msg')} "
                           f"样本={json.dumps(test_batch[0], ensure_ascii=False)[:400]}")
    log(f"[{name}] 预检通过（{len(test_batch)} 条）")

    # 预检通过才开始清空 + 全量写入
    log(f"[{name}] 飞书写入开始 ...")
    existing = bitable_list_records(token, table_id)
    ids = [r["record_id"] for r in existing]
    if ids:
        bitable_batch_delete(token, table_id, ids)
    log(f"[{name}] 已清空旧记录 {len(ids)} 条")

    written = bitable_batch_create(token, table_id, records)
    log(f"[{name}] ✅ 写入完成 {written} 条")
    return {"cleared": len(ids), "written": written}

def run_kehu_incremental(token: str, payload: dict) -> dict:
    """⑤ 客户：增量写入 + 已删除标记（与单脚本一致）"""
    table_id = TABLES["kehu"]
    feishu_records = bitable_list_records(token, table_id)
    plan = kehu_plan(payload["rows"], feishu_records)
    log(f"[⑤ 客户] 计划: 新增 {len(plan['to_create'])} · 更新 {len(plan['to_update'])} · "
        f"标记删除 {len(plan['to_mark_deleted'])} · 恢复 {len(plan['to_restore'])} · 无变化 {plan['unchanged']}")

    created = updated = restored = marked = 0
    if plan["to_create"]:
        created = bitable_batch_create(token, table_id, [{"fields": f} for f in plan["to_create"]])
    if plan["to_update"]:
        updated = bitable_batch_update(token, table_id, plan["to_update"])
    if plan["to_restore"]:
        restored = bitable_batch_update(token, table_id, plan["to_restore"])
        log(f"[⑤ 客户] ♻️ {restored} 条重新出现，已取消「已删除」标记")
    if plan["to_mark_deleted"]:
        marked = bitable_batch_update(token, table_id, [
            {"record_id": r["record_id"], "fields": {"AI 提示": DELETED_MARK}}
            for r in plan["to_mark_deleted"]])
        log(f"[⑤ 客户] ⚠️ {marked} 条在懂火已不存在，已标记「已删除」（记录保留）")
    return {"created": created, "updated": updated, "restored": restored,
            "marked_deleted": marked, "unchanged": plan["unchanged"],
            "deleted_names": [r["name"] for r in plan["to_mark_deleted"]],
            "total_donghuo": plan["total_donghuo"], "total_feishu": plan["total_feishu"]}


def run_kehu_fix_owner(token: str) -> int:
    """
    ⑤ 客户后处理：补参与 + 标记请复检。
    条件：参与为空 AND 来自-收付登记（Lookup）有值 AND AI 提示 != "已删除"。
    动作：取 Lookup 里第一个 user 的 open_id → 写入参与 User 字段；AI 提示写 "请复检"。
    返回处理条数。
    """
    table_id = TABLES["kehu"]
    CURRENT_STEP = "客户后处理：补参与 + 标记请复检"
    log("[⑤ 客户] 后处理：扫描参与空白但有来自-收付登记的记录 ...")
    records = bitable_list_records(token, table_id)
    updates = []
    for rec in records:
        f = rec["fields"] or {}
        # 跳过 AI 提示已标记 "已删除" 的客户
        ai_hint = norm_text(f.get("AI 提示"))
        if ai_hint == DELETED_MARK:
            continue
        # 参与为空 = 没有 User 对象
        owner = f.get("参与")
        owner_empty = not owner or (isinstance(owner, list) and len(owner) == 0) or (isinstance(owner, dict) and not owner)
        if not owner_empty:
            continue
        # 来自-收付登记是 Lookup dict，结构 {"users": [{"id": "ou_xxx", "name": "..."}]}
        lookup = f.get("来自-收付登记")
        if not isinstance(lookup, dict):
            continue
        users = lookup.get("users") or []
        if not users or not isinstance(users, list):
            continue
        first_id = users[0].get("id")
        if not first_id or not first_id.startswith("ou_"):
            continue
        # 还要排除 AI 提示已经是 "请复检" 的（避免重复写相同值）
        if ai_hint == "请复检":
            continue
        updates.append({
            "record_id": rec["record_id"],
            "fields": {
                "参与": [{"id": first_id}],
                "AI 提示": "请复检",
            },
        })
    if not updates:
        log("[⑤ 客户] ✅ 无需补参与的记录（0 条）")
        return 0
    log(f"[⑤ 客户] ⚙️ 补参与 + 标记请复检: {len(updates)} 条")
    done = bitable_batch_update(token, table_id, updates)
    log(f"[⑤ 客户] ✅ 已处理 {done} 条")
    return done

# ============ 汇总通知卡片 ============

def _time_md(res: dict) -> str:
    """各模块用时后缀。取数=fetch_all 拉全量；写入=飞书清空+落库。
    CSV 备份夹在中间，是全体共用的批次步骤，不计进任何单个模块。"""
    f = res.get("fetch_s")
    if f is None:
        return ""
    w = res.get("write_s")
    if w is None:
        return f" · ⏱ {f:.1f}s（取数 {f:.1f}s，未写入）"
    return f" · ⏱ {f + w:.1f}s（取数 {f:.1f}s ／ 写入 {w:.1f}s）"


def _part_line_md(key: str, res: dict) -> str:
    """单部分一行的 markdown 文本"""
    name = PART_NAMES[key]
    if not res.get("ok"):
        err = (res.get("error") or "未知错误").replace("\n", " ")
        return f"**{name}** ❌ 失败{_time_md(res)}\n{err[:200]}"
    s = res.get("stats") or {}
    if key == "kehu":
        extra = f" · 补参与+请复检 {s.get('fixed_owner', 0)}" if s.get("fixed_owner") else ""
        return (f"**{name}** ✅ 新增 {s.get('created', 0)} · 更新 {s.get('updated', 0)} · "
                f"标记删除 {s.get('marked_deleted', 0)} · 恢复 {s.get('restored', 0)} · "
                f"无变化 {s.get('unchanged', 0)}{extra}{_time_md(res)}")
    return (f"**{name}** ✅ 清空 {s.get('cleared', 0)} 条 · 写入 {s.get('written', 0)} 条"
            f"{_time_md(res)}")


def build_summary_card(results: dict, elapsed_s: float) -> dict:
    """十一合一汇总卡片：全成功绿 / 部分失败橙 / 全失败红"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total = len(results)
    ok_cnt = sum(1 for r in results.values() if r.get("ok"))
    if ok_cnt == total:
        template, emoji, title = "green", "✅", f"懂火数据汇总同步完成（{ok_cnt}/{total} 全部成功）"
    elif ok_cnt == 0:
        template, emoji, title = "red", "❌", f"懂火数据汇总同步全部失败（0/{total}）"
    else:
        template, emoji, title = "orange", "⚠️", f"懂火数据汇总同步部分完成（{ok_cnt}/{total} 成功）"

    elements = [
        {"tag": "div", "fields": [
            {"is_short": True, "text": {"tag": "lark_md", "content": f"**总耗时**\n{elapsed_s:.1f} 秒"}},
            {"is_short": True, "text": {"tag": "lark_md", "content": f"**完成时间**\n{now}"}},
        ]},
        {"tag": "hr"},
    ]
    for key in PART_ORDER:
        if key in results:
            elements.append({"tag": "div", "text": {"tag": "lark_md", "content": _part_line_md(key, results[key])}})

    # 客户部分被标记删除的名单
    kehu_res = results.get("kehu") or {}
    if kehu_res.get("ok"):
        names = (kehu_res.get("stats") or {}).get("deleted_names") or []
        if names:
            show = names[:20]
            lines = "\n".join(f"· {n}" for n in show)
            if len(names) > 20:
                lines += f"\n……等共 {len(names)} 家"
            elements.append({"tag": "hr"})
            elements.append({"tag": "div", "text": {"tag": "lark_md",
                "content": f"**⚠️ 已在表内标记「已删除」的客户（懂火中已不存在，记录保留）：**\n{lines}"}})

    elements.append({"tag": "hr"})
    elements.append({"tag": "note", "elements": [{
        "tag": "plain_text",
        "content": f"飞书多维表「数据汇总（2026）」 ｜ 一次登录 · {total} 个模块合并同步 ｜ {now}"
    }]})
    return {"config": {"wide_screen_mode": True},
            "header": {"template": template,
                       "title": {"tag": "plain_text", "content": f"{emoji} {title}"}},
            "elements": elements}



# ============ 建表 / 字段对照表 ============

def bitable_list_tables(token: str) -> list:
    """列出 base 下所有表 [{table_id, name}, ...]"""
    h = {"Authorization": f"Bearer {token}"}
    url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables"
    out, page_token = [], None
    while True:
        qs = "page_size=100" + (f"&page_token={page_token}" if page_token else "")
        d = _fs_get(f"{url}?{qs}", h)
        if d.get("code") != 0:
            raise RuntimeError(f"获取表列表失败: {d}")
        dd = d.get("data") or {}
        out.extend(dd.get("items") or [])
        if not dd.get("has_more") or not dd.get("page_token"):
            break
        page_token = dd.get("page_token")
    return out


def table_id_of(part: str) -> str:
    """取该部分的飞书 table_id；写 ensure 的按表名现场解析（结果缓存）"""
    spec = PARTS[part]
    if spec.get("table"):
        return spec["table"]
    if part not in _ENSURED:
        want = spec["ensure"]
        found = {t["name"]: t["table_id"] for t in bitable_list_tables(feishu_token())}
        if want not in found:
            raise RuntimeError(f"飞书里没有「{want}」表 —— 先跑 --ensure-tables 建表")
        _ENSURED[part] = found[want]
        log(f"[{spec['name']}] 按表名定位到「{want}」→ {found[want]}")
    return _ENSURED[part]


def ensure_tables(token: str, dry_run: bool = False) -> dict:
    """按表名幂等引导 ⑦⑧⑨⑩⑪ 五张新表：已有就用，没有就一次建表带全部字段。
    字段名 1:1 照抄懂火接口，不做任何翻译；首字段即主字段，必须是文本。"""
    existing = {t["name"]: t["table_id"] for t in bitable_list_tables(token)}
    resolved = {}
    for part in PART_ORDER:
        spec = PARTS[part]
        want = spec.get("ensure")
        if not want:
            continue
        if want in existing:
            resolved[part] = existing[want]
            log(f"[建表] {spec['name']}: 已存在「{want}」 → {existing[want]}（跳过）")
            continue
        fields = []
        for f, t in spec["fields"].items():
            fd = {"field_name": f, "type": _STRATEGY_TO_CREATE_CODE[t]}
            if t == "number":
                # 必须显式给 formatter，否则飞书默认 "0.0" 把重量/金额的显示砍到 1 位小数
                fd["property"] = {"formatter": _number_formatter(f)}
            elif t == "datetime":
                # 同上：不带 date_formatter 会跟 ①~⑨ 的日期列长得不一样
                fd["property"] = dict(_DATE_PROP)
            fields.append(fd)
        log(f"[建表] {spec['name']}: 新建「{want}」，{len(fields)} 个字段 ...")
        if dry_run:
            for i, f in enumerate(fields, 1):
                log(f"      {i:>2}. {f['field_name']:<12} type={f['type']}")
            resolved[part] = "<dry-run>"
            continue
        h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        url = f"{FEISHU_OPEN_BASE}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables"
        d = requests.post(url, headers=h, timeout=60,
                          json={"table": {"name": want, "default_view_name": "表格",
                                          "fields": fields}}).json()
        if d.get("code") != 0:
            raise RuntimeError(f"建表「{want}」失败: {d.get('code')} {d.get('msg')}")
        tid = (d.get("data") or {}).get("table_id")
        resolved[part] = tid
        log(f"[建表] ✅ 已建「{want}」 → {tid}")
        got = bitable_get_field_types(token, tid)
        log(f"[建表] 回读校验：{len(got)} 个字段 → {', '.join(got)}")
    return resolved


def check_fields(session, token: str, to_markdown: bool = False) -> int:
    """只读：逐模块核对「飞书列 ↔ 接口键 ↔ 脚本映射」，输出告警 + 对照表。
    返回发现的问题数（0 = 干净）。"""
    problems = 0
    md = ["# 十一合一字段对照表", "",
          f"生成时间：{datetime.datetime.now():%Y-%m-%d %H:%M:%S}", "",
          "> 由 `python 数据汇总/Update_Data.py --check-fields --markdown` 生成。",
          "> 约定：**飞书列名 == 懂火接口键名**，脚本不做任何别名或强行映射；",
          "> 名字不一致时以接口键名为准，去改飞书表里的列名。", ""]
    for part in PART_ORDER:
        spec, name = PARTS[part], PARTS[part]["name"]
        try:
            keys = probe_json_keys(session, part)
        except Exception as e:
            log(f"[对照] {name}: ❌ 接口不可用（{type(e).__name__}: {e}）")
            md += [f"## {name}", "", f"❌ 接口不可用：`{e}`", ""]
            problems += 1
            continue
        try:
            table_id = table_id_of(part)
            feishu = bitable_get_field_types(token, table_id)
        except Exception as e:
            log(f"[对照] {name}: ⚠️ 读飞书字段失败（{e}）")
            md += [f"## {name}", "",
                   f"⚠️ 读飞书字段失败：`{e}`", "",
                   f"接口返回 {len(keys)} 键：{', '.join(f'`{k}`' for k in keys)}", ""]
            problems += 1
            continue

        # ② 动态探测、⑤ 增量：映射就取飞书表现有字段；其余用硬编码映射
        effective = feishu if (spec.get("dynamic") or spec.get("incremental")) else spec["fields"]
        writable = {c: t for c, t in effective.items() if t != "unsupported"}
        unsupported = [c for c, t in effective.items() if t == "unsupported"]
        no_source = [c for c in writable if c not in keys]
        excluded = set(EXCLUDED_API_KEYS.get(part, []))
        no_column = [k for k in keys if k not in effective and k != "id" and k not in excluded]
        not_in_feishu = [c for c in effective if c not in feishu]

        log(f"[对照] {name}: 接口 {len(keys)} 键 / 可写映射 {len(writable)} 列 / 表里 {len(feishu)} 字段"
            + ("（动态探测）" if spec.get("dynamic") else ""))
        if unsupported:
            log(f"[对照]   不可写（自动跳过）: {', '.join(unsupported)}")
        if no_source:
            log(f"[对照]   ⬜ 表里有列、接口无键 → 留空: {', '.join(no_source)}")
        if no_column:
            log(f"[对照]   ➖ 接口有键、表里无列 → 不入表: {', '.join(no_column)}")
        if excluded:
            log(f"[对照]   🚫 主动排除（有意不同步）: {', '.join(sorted(excluded))}")
        if not_in_feishu:
            log(f"[对照]   ⚠️ [WARN] 脚本映射了但飞书表里没这列: {', '.join(not_in_feishu)}")
            problems += 1

        md += [f"## {name}", "",
               f"- 接口返回 **{len(keys)}** 键 ｜ 飞书表 **{len(feishu)}** 字段 ｜ 可写 **{len(writable)}** 列",
               ""]
        if unsupported:
            md.append(f"- 不可写（自动跳过）：{', '.join(f'`{c}`' for c in unsupported)}")
        if no_source:
            md.append(f"- ⬜ **表里有列、接口无键 → 该列留空**：{', '.join(f'`{c}`' for c in no_source)}")
        if no_column:
            md.append(f"- ➖ 接口有键、表里无列 → 不入表：{', '.join(f'`{k}`' for k in no_column)}")
        if excluded:
            md.append(f"- 🚫 **主动排除（有意不同步）**："
                      f"{', '.join(f'`{k}`' for k in sorted(excluded))}")
        if not_in_feishu:
            md.append(f"- ⚠️ **脚本映射了但飞书表里没这列（写不进去）**："
                      f"{', '.join(f'`{c}`' for c in not_in_feishu)}")
        md += ["", "| 飞书列 | 接口键 | 写入策略 | 状态 |", "|---|---|---|---|"]
        for c, t in effective.items():
            if t == "unsupported":
                st = "不可写，跳过"
            elif c not in keys:
                st = "⬜ 接口无此键，留空"
            elif c not in feishu:
                st = "⚠️ 飞书表无此列"
            else:
                st = "✅ 同名直连"
            md.append(f"| {c} | {'—' if c not in keys else c} | {t} | {st} |")
        for k in no_column:
            md.append(f"| — | {k} | — | ➖ 表里无列 |")
        md.append("")

    if to_markdown:
        MD_PATH.write_text("\n".join(md) + "\n", encoding="utf-8")
        log(f"[对照] 📄 对照表已写出: {MD_PATH}")
    log(f"[对照] 结论: {'✅ 无问题' if problems == 0 else f'⚠️ {problems} 处需要关注'}")
    return problems



# ============ 主流程 ============

# ============ 并发写入（表级并发：不同飞书表可同时写，同一张表必须串行）============

def _is_tail_part(part: str) -> bool:
    """该模块是否要留到最后串行收尾（不进线程池）。

    ⑤ 客户是唯一的增量模块，且它第一步就读自己的表。而 ②③④ 都有一个
    `参与`(Lookup) 指向⑤ —— 它们被批量重写时，飞书要重算这些 Lookup，
    ⑤ 期间会短暂返回 `1254607 Data not ready`（2026-10-03 CI 实测踩到）。
    串行版里②③④在⑤之前写完，这个窗口不存在，是并发引入的新问题。
    所以⑤不进池，等其它表全部写完后由主线程单独跑（它很小，基本不占时长）。
    """
    return bool(PARTS[part].get("incremental"))


def _write_part(token: str, part: str, payload_part: dict) -> dict:
    """单个模块的飞书写入 —— 线程池任务单元，也是单模块内联执行的同一个入口。

    CSV 落盘也在这里做（2026-10-06 从取数主线程挪进来）：实测 11 个模块落盘
    合计 ~12s，而取数是严格串行的，落在主线程里就是直接挡住下一个模块。
    挪进来之后「先落 CSV、后写飞书」的顺序一字不变，落盘失败照样不写飞书。

    ⚠️ 契约：自己吞掉所有异常、以结构化结果返回，绝不把异常抛给 future。
       否则 as_completed 循环里某个 future 一抛错，主线程就会跳过其余模块的结果，
       卡片显示错误、退出码也会不对。
    """
    name = PARTS[part]["name"]
    t = time.time()
    try:
        csv_path = rows_to_csv(payload_part["rows"], part)
    except Exception as e:
        return {"part": part, "ok": False, "stats": None,
                "error": f"CSV 落盘失败（未写飞书）: {e}", "write_s": time.time() - t,
                "tb": traceback.format_exc(limit=3), "csv": None}
    try:
        if PARTS[part].get("incremental"):
            # ⑤ 的增量写入与「补参与」后处理是不可分割的一个单元：
            # 拆成两个任务会并发写同一张表（1254291 Write conflict）
            stats = run_kehu_incremental(token, payload_part)
            try:
                stats["fixed_owner"] = run_kehu_fix_owner(token)
            except Exception as e2:
                log(f"[{name}] ⚠️ 补参与后处理异常（不阻断主流程）: {e2}")
                stats["fixed_owner"] = 0
        else:
            stats = run_full_replace(token, part, payload_part)
        return {"part": part, "ok": True, "stats": stats, "error": None,
                "write_s": time.time() - t, "tb": None, "csv": csv_path}
    except Exception as e:
        return {"part": part, "ok": False, "stats": None,
                "error": f"飞书写入失败: {e}", "write_s": time.time() - t,
                "tb": traceback.format_exc(limit=3), "csv": csv_path}


def _merge(results: dict, part: str, csv_path, r: dict):
    """把单个模块的写入结果并入 results —— 只在主线程调用，保证 results 无竞态。"""
    name = PARTS[part]["name"]
    # write_s 无条件赋值：卡片里每个提交过的模块都要能打出用时
    results[part].update(ok=bool(r.get("ok")), stats=r.get("stats"), write_s=r.get("write_s"))
    if r.get("ok"):
        # 仅写入成功才删 CSV；失败保留作证据（沿用原有 CSV 纪律）
        if csv_path is not None and Path(csv_path).exists():
            Path(csv_path).unlink()
            log(f"[{name}] [清理] CSV 已删除: {Path(csv_path).name}")
    else:
        results[part]["error"] = r.get("error") or "未知写入错误"
        log(f"[{name}] ❌ {results[part]['error']}")
        if r.get("tb"):
            log(f"    {r['tb']}")
    fs, ws = results[part].get("fetch_s"), results[part].get("write_s")
    log(f"[{name}] ⏱ 取数 " + (f"{fs:.1f}s" if fs is not None else "—")
        + " + 写入 " + (f"{ws:.1f}s" if ws is not None else "—"))


def main():
    ap = argparse.ArgumentParser(description="懂火 11 合 1 数据汇总同步（纯 JSON 接口，零浏览器）")
    ap.add_argument("--dry-run", action="store_true", help="取数 + 落 CSV + 打印计划，不写飞书")
    ap.add_argument("--no-notify", action="store_true", help="不发飞书通知")
    ap.add_argument("--only", default="", help=f"只跑部分：逗号分隔 {','.join(PART_ORDER)}")
    ap.add_argument("--check-fields", action="store_true",
                    help="只读：核对字段对照表并输出告警（需先登录懂火）")
    ap.add_argument("--markdown", action="store_true", help="配合 --check-fields，写出对照表 .md")
    ap.add_argument("--ensure-tables", action="store_true", help="幂等建 ⑦⑧⑨⑩⑪ 五张新表（按表名查，缺则建）")
    ap.add_argument("--workers", type=int, default=DEFAULT_WORKERS,
                    help=f"并发写表的线程数（默认 {DEFAULT_WORKERS}，上限 11）。"
                         f"1 = 串行写，配合 --no-pipeline 等价于改造前的行为")
    ap.add_argument("--no-pipeline", action="store_true",
                    help="关掉取写流水线：仍然并发写表，但先把 11 个模块全部取完才开始写")
    args = ap.parse_args()

    t0 = time.time()
    only = {s.strip() for s in args.only.split(",") if s.strip()} or set(PART_ORDER)
    invalid = only - set(PART_ORDER)
    if invalid:
        raise SystemExit(f"--only 含未知部分: {invalid}（可选: {PART_ORDER}）")

    # ---- 只读模式：建表 / 字段对照表 ----
    if args.check_fields or args.ensure_tables:
        log(f"==== 懂火十一合一 · {'字段对照' if args.check_fields else '建表'} 模式 ====")
        token = feishu_token()
        if args.ensure_tables:
            ensure_tables(token, dry_run=args.dry_run)
        if args.check_fields:
            session = login_donghuo()
            if session is None:
                raise SystemExit("懂火登录失败（字段对照需要读接口键名）")
            rc = check_fields(session, token, to_markdown=args.markdown)
            log(f"==== 对照完成，耗时 {time.time() - t0:.1f}s ====")
            return rc
        log(f"==== 建表完成，耗时 {time.time() - t0:.1f}s ====")
        return 0

    log(f"==== 懂火 {len(PART_ORDER)} 合 1 同步（纯 JSON）启动（部分: {sorted(only)}）====")
    username, password = env("DH_USERNAME"), env("DH_PASSWORD")
    if not username or not password:
        raise SystemExit("缺少 DH_USERNAME / DH_PASSWORD（.env）")

    results = {k: {"ok": False, "error": None, "stats": None, "fetch_s": None, "write_s": None}
               for k in PART_ORDER if k in only}

    # ---- Phase 1: 懂火登录一次（之后各模块复用同一个 session）----
    log("[懂火] 登录中（全程仅此一次）...")
    session = login_donghuo()
    if session is None:
        raise SystemExit("懂火登录失败，终止")

    # ---- 写入侧准备（dry-run 不需要）----
    token, tids, pool = None, {}, None
    if not args.dry_run:
        # 流水线模式下取数中途就要提交写入任务，token 必须提前拿
        log("[飞书] 获取 tenant_access_token ...")
        token = feishu_token()
        # 单线程预解析 ⑦⑧⑨⑩⑪ 的 table_id（这几张要按表名现场查），
        # 避免多个写表线程同时首次解析同一个模块
        log("[飞书] 解析目标表 id ...")
        for part in PART_ORDER:
            if part not in only:
                continue
            try:
                tids[part] = table_id_of(part)
            except Exception as e:
                results[part]["error"] = f"定位飞书表失败: {e}"
                log(f"[{PARTS[part]['name']}] ❌ {e}")
        # 机械保险：飞书同一张表不支持并发写（1254291 Write conflict）。
        # 一旦两个模块指向同一张表就直接拦下，而不是等线上偶发写冲突
        if len(set(tids.values())) != len(tids):
            raise SystemExit(f"多个模块指向同一张飞书表，会触发写冲突: {tids}")
        if args.workers > 1 and len(tids) > 1:
            n = min(args.workers, len(tids))
            pool = ThreadPoolExecutor(max_workers=n)
            tail = [PARTS[p]["name"] for p in PART_ORDER if _is_tail_part(p) and p in tids]
            log(f"[飞书] 并发写表：{n} 个线程"
                f"{'（流水线：取完一个就写一个）' if not args.no_pipeline else '（--no-pipeline：全部取完再写）'}"
                + (f"，{'/'.join(tail)} 留到最后串行收尾" if tail else ""))
        else:
            log("[飞书] 串行写入（未启用并发）")

    # ---- Phase 2+3+4: 逐模块 取数 → 落 CSV → 提交写入（每模块独立 try，互不阻断）----
    # 取数顺序用 FETCH_ORDER（写入量降序）：让最慢的 ① 最早开始写，整条链最短
    payload, futs, pending = {}, {}, []
    for part in FETCH_ORDER:
        if part not in only:
            continue
        t_part = time.time()
        try:
            rows = fetch_all(session, part)
            results[part]["fetch_s"] = time.time() - t_part   # 只计接口翻页，CSV 落盘另算
            if not rows:
                raise RuntimeError("接口返回空")
            # dry-run 没有写入线程，CSV 只能在这里落；正式跑时由 _write_part 落
            # （「先落 CSV、后写飞书」在那边同样成立，且不再挡取数）
            payload[part] = {"rows": rows}
            if args.dry_run:
                payload[part]["csv"] = rows_to_csv(rows, part)
        except Exception as e:
            if results[part]["fetch_s"] is None:
                results[part]["fetch_s"] = time.time() - t_part
            results[part]["error"] = results[part]["error"] or f"取数/CSV 失败: {e}"
            log(f"[{PARTS[part]['name']}] ❌ {e}")
            continue

        if args.dry_run or part not in tids:
            continue
        if pool is not None and not args.no_pipeline and not _is_tail_part(part):
            futs[pool.submit(_write_part, token, part, payload[part])] = part
        else:
            pending.append(part)   # 未进池（--no-pipeline / 串行模式 / ⑤ 收尾模块），稍后统一写

    # ---- dry-run：打印统计 + ⑤ 客户增量计划，到此为止 ----
    if args.dry_run:
        log("==== DRY-RUN 结果（不写飞书，CSV 全部保留）====")
        for part in PART_ORDER:
            if part not in only:
                continue
            name = PARTS[part]["name"]
            if part in payload:
                keys = list(payload[part]["rows"][0].keys())
                log(f"  {name}: {len(payload[part]['rows'])} 行 × {len(keys)} 键")
            else:
                log(f"  {name}: ❌ {results[part]['error']}")
        if "kehu" in payload:
            try:
                token = feishu_token()
                plan = kehu_plan(payload["kehu"]["rows"], bitable_list_records(token, TABLES["kehu"]))
                log(f"  ⑤ 客户增量计划: 新增 {len(plan['to_create'])} · 更新 {len(plan['to_update'])} · "
                    f"标记删除 {len(plan['to_mark_deleted'])} · 恢复 {len(plan['to_restore'])} · "
                    f"无变化 {plan['unchanged']}")
                if plan["to_mark_deleted"]:
                    log(f"    将标记: {[r['name'] for r in plan['to_mark_deleted'][:10]]}")
            except Exception as e:
                log(f"  ⑤ 客户计划读取失败: {e}")
        log(f"==== DRY-RUN 完成，耗时 {time.time() - t0:.1f}s ====")
        return 0

    # ---- Phase 4: 飞书写入 ----
    # --no-pipeline：等取数全部结束后，再把待写模块一次性并发提交
    # ⑤ 等收尾模块（_is_tail_part）无论哪种模式都不进池，留在 pending 里最后由主线程串行跑
    if pool is not None and pending:
        deferred = [p for p in pending if _is_tail_part(p)]
        for part in pending:
            if _is_tail_part(part):
                continue
            futs[pool.submit(_write_part, token, part, payload[part])] = part
        pending = deferred

    # 主线程消费已完成的任务：results 只在这里被写，天然无竞态
    for fut in as_completed(futs):
        part = futs[fut]
        try:
            r = fut.result()
        except BaseException as e:   # 理论不可达（_write_part 自己兜底），防 write_s 丢失
            r = {"part": part, "ok": False, "stats": None,
                 "error": f"写入线程异常: {e}", "write_s": None,
                 "tb": traceback.format_exc(limit=3)}
        _merge(results, part, r.get("csv"), r)
    if pool is not None:
        pool.shutdown(wait=True)

    # 串行模式 / --only 单模块：主线程内联跑同一个 _write_part，行为与池内完全一致
    for part in pending:
        r = _write_part(token, part, payload[part])
        _merge(results, part, r.get("csv"), r)

    elapsed = time.time() - t0
    ok_cnt = sum(1 for r in results.values() if r.get("ok"))
    log(f"==== 全部结束: {ok_cnt}/{len(results)} 部分成功，总耗时 {elapsed:.1f}s ====")

    # ---- Phase 5: 汇总通知（一张卡）----
    if not args.no_notify:
        try:
            feishu_send_card(FEISHU_NOTIFY_UNION_ID,
                             build_summary_card(results, elapsed), feishu_token())
        except Exception as e:
            log(f"[飞书通知] ⚠️ 发送异常: {e}")

    return 0 if ok_cnt == len(results) else (2 if ok_cnt == 0 else 1)

if __name__ == "__main__":
    sys.exit(main())
