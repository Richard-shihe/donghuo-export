# AGENTS.md — 本仓库 AI 协作规则

> 本文件供 AI 编程助手（Trae / Codex / Claude Code / Cursor 等）在每次会话自动加载，
> 同时也是人类协作者的"入职文档"。规则按重要性排序，红线部分违反会造成真实事故。

## 1. 项目概述

`donghuo-export`：围绕懂火 ERP / IEC / 欧冶 / 宝钢 / 鞍钢等钢铁电商业务的数据导出、同步、备份自动化脚本集合。
技术栈：Python 3.10 + Playwright（登录爬取）+ lark-oapi（飞书）+ GitHub Actions（定时 CI）+ 本地 Windows .bat 守护进程。
**单仓库多项目**：每个业务项目独立子文件夹，公共登录库与配置放根目录。无测试框架、无 lint，验证靠实际运行。

## 2. 仓库结构与红线

### 2.1 根目录公共文件（严禁移动 / 重命名 / 删除）

- `donghuo_login.py` / `ibaosteel_client.py` — 公共登录库，被子文件夹脚本通过 `sys.path` 指向仓库根反向 import
- `iecc.json` — IEC 凭据缓存；`.env` — 所有密钥；`requirements.txt` / `README.md` / `.gitignore`

**改公共库前必须先 grep 全仓库引用点**，评估对全部项目的影响。

### 2.2 项目文件夹（每文件夹 = 一条完整业务链）

| 文件夹 | 业务 |
|---|---|
| `数据汇总/` | 5合1数据汇总（出库/订单/应收/往来/客户 → 飞书多维表格） |
| `懂火出库导出/` | export_chuku.py，配套 `DATA UPDATE/` 本地 runner |
| `懂火加工单同步/` | export + import 双向 |
| `懂火每周全量备份/` | 全量备份 |
| `懂火采购入库/` | 采购订单 + IEC 码单，5 脚本 5 workflow |
| `IEC准发出厂结案/` | 准发 / 出厂 / 结案，6 脚本 3 workflow |
| `开票申请单/` | 开票导出 |
| `临调欧冶/` | export_lindiao + export_ouyeel_exel 同表数据链 |
| `Certification/` | 质保书三段（下载/分类/整理） |
| `欧冶/` | 本地项目无 CI（一键导出 / 飞书监听 / 循环导出） |
| `宝钢酸洗/` | 本地项目（每日简报 / 时序下载） |
| `标准知识库/` | 本地，飞书 wiki 上传工具（未跟踪） |
| `DATA UPDATE/` | 本地 runner |

### 2.3 结构红线

1. **新项目一律开新子文件夹**，根目录不再新增散落文件。每次任务开始先声明所属项目文件夹边界。
2. **以下文件夹特意未跟踪，绝对不要 `git add`**：`欧冶/`、`标准知识库/`、`宝钢酸洗/`
   ——`标准知识库/wiki_upload_all.py` 含硬编码 APP_SECRET，提交会被 GitHub secret scanning 拦截（已实际发生过）。
3. `.env`、`iecc.json`、`*.csv`、`*.xls` 等数据/凭据文件永远不提交。
4. 历史提交中含飞书 App Secret，push 可能被保护拦截——push 前先 `git status` 复核暂存内容。

## 3. 常用命令

```bash
# 所有脚本一律从仓库根运行（无 working-directory，脚本 cwd 恒为仓库根）
python 懂火出库导出/export_chuku.py
python 数据汇总/sync_chuku_to_bitable.py

# 依赖
pip install -r requirements.txt

# GitHub Actions 手动触发用 workflow_dispatch
```

- Workflow 调用规范：`python 项目文件夹/脚本.py`，不设 working-directory。
- 无测试命令。验证方式 = 实际运行目标脚本（数据类脚本先小样本或 dry-run）。

## 4. 硬性约束（Hard Constraints）

1. **改 workflow / 关键脚本后必须 grep 回读验证落盘**——曾发生 Edit 工具回显成功但未写盘的漏改。
2. **CSV 文件上传成功后必须删除**。仅调试参数、预检写入失败时可保留。
3. **密钥只从 `.env` 读取，任何脚本不得硬编码 secret**。飞书监听授权名单读 `.env` 的 `FEISHU_LISTENER_UNION_IDS`。
4. **飞书通知一律走机器人 / 应用身份，禁止用个人账号给任何人发消息。**
5. 懂火出库同步脚本完成后需自动发飞书文本通知（既有行为，勿删）。
6. 不批量格式化、不顺手重构无关代码（外科手术式修改）。

## 5. 编码规范（本仓库特有，均为真实事故换来的）

1. **所有被 subprocess / bat 启动的 Python 脚本，开头统一**：
   ```python
   import sys
   sys.stdout.reconfigure(encoding="utf-8", errors="replace")
   sys.stderr.reconfigure(encoding="utf-8", errors="replace")
   ```
   并在父进程 `os.environ.setdefault("PYTHONIOENCODING", "utf-8")` 双保险。
   原因：Windows cp936 下子进程 print emoji（✅❌⚠️）必抛 UnicodeEncodeError。
2. **print 输出用 `[OK]` / `[FAIL]` / `[WARN]` 文本标记，不用 emoji**。
3. **`.bat` 文件必须存 GBK(CP936) 编码**，删掉 `chcp 65001` 行——cmd 在 chcp 生效前已按 GBK 解析，UTF-8 中文行会变成乱码命令。
4. **变量命名前后必须一致**——曾因公式分支赋值 `cell_out`、写入时引用 `out_cell`，导致公式全部丢失。
5. **格式判断保守**：自定义格式含"月/日/年"字样不一定是日期（"月息1.0%"被误判为日期格式导致整列 #VALUE!）。

6. **飞书多维表建 Number 字段必须显式带 `property.formatter`**。不传的话飞书默认 `"0.0"`，
   把源头的小数在**屏幕上**四舍五入到 1 位——懂火的 `2.190` 显示成 2.2、`0.370` 显示成 0.4、
   `4.136` 显示成 4.1，逐行看每一件货都不对、合计也不对（**存储值本身是好的，坏的只有显示**）。
   口径：**重量 `0.000`、金额/单价/税率 `0.00`、件数/库龄 `0`**。
   建表/加字段后必须回读 `bitable/v1/apps/{app}/tables/{tid}/fields` 的 `property.formatter` 复核。
   事故：`数据汇总/Update_Data.py` 的 `--ensure-tables` 建 ⑦采购订单/⑧采购明细/⑨库存 时漏了 property，
   29 个字段全部中招，2026-10-04 才被用户发现（修法见 `数据汇总/fix_number_formatter.py`；
   源头已在该脚本 `_number_formatter()` 堵上，回读校验 58 个 Number 字段全部吻合）。

## 6. 已知坑（Known Gotchas）

- **Windows 进程计数**：bat 里裸 `python` 经微软商店桩转发，任务管理器会显示两个 python.exe（桩 + 真解释器），**不是**双实例；判重只数命令行含完整解释器路径（pythoncore/Anaconda）的进程，杀进程要杀真解释器。
- **Playwright 过瑞数**：headless 必白屏。必须真实 Chrome（`channel="chrome"`）+ headed + `--disable-blink-features=AutomationControlled` + 去掉 `--enable-automation` + storage_state；登录会话约 2 小时过期，验证码必须人工。
- **lark-oapi 事件字段路径**是 `data.event.message` / `data.event.sender`（不是 `data.message`，错路会 AttributeError）。
- **重复双击 bat 会多实例并存**（循环重复上传 / 监听重复回复）；重启守护前必须按命令行匹配杀掉 `feishu_listener.py` / `loop_export.py` 全部真进程。
- 欧冶「下载资源单」按钮必须登录后点击才生效；会话过期后导航栏变回"登录/注册"。

- **核对「懂火 ERP ↔ 飞书表」数据一致性时，API 读回的存储值一致 ≠ 用户看到的对**。飞书 Number 字段的
  `property.formatter` 会改屏幕显示，只比存储值必然得到"全部一致"的假结论（已实际发生过：连报三次"三方全对"，
  而用户屏幕上每一件货的重量都被四舍五入）。核对至少要包含两层：① 逐字段读 `property.formatter`、
  和源头实际的小数位数对齐；② 归一化别写 `v or ""`——数字 `0` 是 falsy，会被吃成空串，凭空造出满屏假差异。
  另：ERP 的导出接口不一定有全部列（库存导出的 26 列里没有「可售重量/可售件数」），
  拿"导出缺列"当成"导出值为 0"会得出空对空的结论。

## 7. 任务完成标准（Definition of Done）

- [ ] 脚本改动已实际运行验证（或 dry-run 小样本）
- [ ] workflow 改动已 grep 回读确认落盘，调用路径符合 `python 项目文件夹/脚本.py` 规范
- [ ] 未在根目录新增散落文件；未触碰三个未跟踪文件夹
- [ ] 运行产生的 CSV / 临时文件已清理
- [ ] 提交前 `git status` 确认无 .env / iecc.json / 数据文件混入

## 8. 工作方式

- **先思考再动手**：假设说出来，不确定就问，不猜。
- **简单优先**：能解决问题的最少代码；单次使用的逻辑不抽象。
- **目标驱动**：把任务转成可验证目标（跑通脚本 / 回读确认），循环直到验证通过。
- 修改前先读目标文件和相关 workflow，理解既有模式后再动手。
