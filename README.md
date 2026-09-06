# donghuo-export

懂火 ERP、宝钢 IEC 等业务系统的数据导出 / 同步 / 自动化工具集，按项目分文件夹管理，CI 统一走 GitHub Actions（`.github/workflows/`）。

---

## 一、项目结构

```
A/
├── donghuo_login.py          # 公共库：懂火登录（requests + ddddocr 验证码）
├── ibaosteel_client.py       # 公共库：宝钢 IEC 登录客户端
├── iecc.json                 # IEC 会话缓存（运行时生成）
├── requirements.txt          # 全仓库 Python 依赖
├── .env                      # 本地凭据（gitignore，脚本向上查找）
│
├── 懂火出库同步/             # 项目：数据汇总（5 合 1 同步）
├── 懂火出库导出/             # 项目：出库记录导出 → 飞书云盘/邮件
├── 懂火加工单同步/           # 项目：加工单双向同步（懂火 ↔ 飞书多维表）
├── 懂火每周全量备份/         # 项目：10 类业务数据每周备份 → 飞书云盘
├── 懂火采购入库/             # 项目：采购订单 + IEC 码单 → 飞书中间表 → 懂火入库
├── IEC准发出厂结案/          # 项目：IEC 准发/出厂/结案 → 飞书
├── 开票申请单/               # 项目：懂火开票申请单 → 飞书多维表
├── 临调欧冶/                 # 项目：临调库存导出 + 欧冶模板 Excel 私聊
├── Certification/            # 项目：质保书三段流水线（下载/分类/整理）
├── 欧冶/                     # 项目：欧冶数据导出（本地三入口，无 CI）
├── 标准知识库/               # 项目：国外标准 → 飞书知识库（已完结）
├── 宝钢酸洗/                 # 项目：宝钢在售清单分析 + 每日简报（本地）
├── DATA UPDATE/              # 配套：出库全量导出本地 runner
└── .github/workflows/        # 20 个 GitHub Actions workflow
```

---

## 二、项目清单与对应 Workflow

| 项目文件夹 | 业务 | Workflow |
|---|---|---|
| 懂火出库同步/ | 懂火 5 模块（出库/订单/应收/往来/客户）→ 飞书多维表，一次性登录批量写入 | `Update_Data.yml`、`sync_chuku_to_bitable.yml` |
| 懂火出库导出/ | 懂火出库记录 → CSV → 飞书云盘/邮件 | `export.yml` |
| 懂火加工单同步/ | 加工单导出到飞书 + 加工成品写回懂火 | `export_jiagong.yml`、`import_jiagong.yml` |
| 懂火每周全量备份/ | 10 类业务数据 → XLSX → 飞书云盘，周报通知 | `backup_all.yml` |
| 懂火采购入库/ | 采购订单↔飞书中间表双向 + IEC 码单下载/比对/增量入库 | `export_caigou_to_bitable.yml`、`import_to_donghuo.yml`、`madan_1_download.yml`、`madan_2_compare_ruku.yml`、`madan_full_pipeline.yml` |
| IEC准发出厂结案/ | IEC 准发下载/进度表更新 + 捆包→中间表 + 结案标记 | `export_IEC_zhunfa.yml`、`export_IEC_jiean.yml`、`import_lindiao_to_bitable.yml` |
| 开票申请单/ | 懂火待确认开票单 → 飞书「滚动」多维表 | `export_kaipiao.yml` |
| 临调欧冶/ | 临调库存→多维表 + 同表视图生成欧冶 Excel 私聊发送 | `export_lindiao.yml`、`export_ouyeel_exel.yml` |
| Certification/ | 质保书：下载→A17 分类→整理归档 三段流水线 | `cert_stage1_download.yml`、`cert_stage2_classify.yml`、`cert_stage3_organize.yml` |
| 欧冶/ | 欧冶数据导出（一键导出/飞书监听/循环导出，本地 bat） | 无（本地运行） |
| 标准知识库/ | 国外标准 709 条 → 飞书知识库（已完结归档） | 无 |
| 宝钢酸洗/ | 宝钢在售清单下载/比对/时序分析/每日简报 | 无（本地运行） |

---

## 三、公共约定

- **公共库在根目录**：`donghuo_login.py` / `ibaosteel_client.py`。子文件夹脚本用 `sys.path.insert(0, 仓库根)` 引用，勿移动。
- **workflow 从仓库根调用脚本**：`python 项目文件夹/脚本.py`，无 `working-directory` 设置，脚本的 cwd 一律是仓库根。
- **产物文件不入库**：CSV / 带时间戳 xlsx / 日志等由 .gitignore 统一忽略，上传成功后自动清理。
- **本地调试脚本约定**：`_` 前缀（`/_*.py` 等）仅限仓库根目录，gitignore，用完即删。
- **凭据**：CI 走 GitHub Secrets / Variables；本地走根目录 `.env`（脚本用 `parent.parent` 向上查找）。
- **新项目一律开新子文件夹**，不再往根目录散落文件。

---

## 四、本地运行

```bash
pip install -r requirements.txt

# 配置凭据（参考 .env.example，从仓库根目录读）
# DH_USERNAME / DH_PASSWORD / FEISHU_APP_ID / FEISHU_APP_SECRET / ...

# 例：手动跑数据汇总
python 懂火出库同步/Update_Data.py

# 例：本地全量出库导出（含 CSV 备份到 DATA UPDATE/）
python "DATA UPDATE/run_full_chuku_export.py"
```

各项目的详细参数（dry-run、日期范围、folder token 覆盖等）见各脚本头部 docstring。
