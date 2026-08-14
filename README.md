# 懂火钢城系统 - 出库记录自动导出

自动登录 `https://erpa.donghuo.vip` → 识别验证码 → 调接口拉取出库记录 → 生成 CSV → **上传到飞书云盘指定文件夹** → （可选）飞书机器人通知你。

每天北京时间 **10:00** 和 **22:00** 各执行一次（可在 `.github/workflows/export.yml` 的 `cron` 中修改）。

---

## 一、文件结构

```
A/
├── export_chuku.py               # 主脚本：登录 / 抓数据 / CSV / 上传飞书云盘 / 通知
├── requirements.txt              # Python 依赖（requests、ddddocr）
└── .github/workflows/export.yml  # GitHub Actions 定时工作流
```

---

## 二、交付模式

通过环境变量 `DELIVERY_MODE` 切换，默认 `feishu`：

| DELIVERY_MODE | 说明 |
|---------------|------|
| `feishu`（默认）| 上传到 **飞书云盘指定文件夹**，（可选）用飞书机器人在群里发通知 |
| `mail`        | 邮件附件模式（之前的方案，保留备用） |

---

## 三、飞书前置准备（一次配置，长期使用）

### 3.1 自建应用：拿到 App ID / App Secret（用于 Open API）

1. 打开 [飞书开放平台](https://open.feishu.cn/app) → `创建企业自建应用` → 填写名称（如"出库记录机器人"）
2. 进入应用 → 左侧 `凭证与基础信息` → 复制 **App ID** 和 **App Secret**
3. 左侧 `权限管理` → 打开以下权限（必须）：
   - 云文档 / 云空间：`查看、评论、编辑和删除云空间所有文件`（`drive:drive`，或按需更细粒度）
   - 云文档 / 云空间：`上传文件到云空间`（`drive:drive.file:upload`）
   - 机器人通知可选权限（应用发消息用，不需要可不勾）
4. 左侧 `版本管理与发布` → `创建版本` → `申请发布`（需企业管理员审核通过，权限才生效）

> ⚠️ 没发布/权限没勾 → 会报 `99991668 无权限` 等错误

### 3.2 获取目标文件夹的 Folder Token

1. **用浏览器（网页版）打开**你想上传到的飞书云盘文件夹（必须是已发布应用的企业内部空间，不能是外部协作空间）
2. 看浏览器地址栏 URL，一般形如：
   ```
   https://xxx.feishu.cn/drive/folder/fldcnABCDEFGHIJKLMNOP
   ```
   **`fldcnABCDEFGHIJKLMNOP` 就是 folder token**
   - 或者在飞书 App 里：右上角 `...` → `分享` → `复制链接` → 粘贴到记事本 → 链接里 `/folder/` 后面的部分就是 token
3. 把文件夹**添加为应用的协作成员**（重要！）：
   - 在云盘里选中该文件夹 → 右上角 `...`（或右键）→ `管理协作者` → `添加协作者`
   - 搜索你刚才创建的应用名（"出库记录机器人"）→ 设为 `可编辑` 或更高权限 → 保存

   否则应用看不到这个文件夹，上传会报 `99991663 父节点不存在` 或 `99991668 无权限`。

### 3.3（可选）创建飞书群自定义机器人（用于通知你上传成功）

1. 打开飞书群 → `设置` → `群机器人` → `添加机器人` → 选 `自定义机器人`
2. 输入名字（如"出库记录通知"）→ 复制 **Webhook 地址**
3. 安全设置建议勾选「签名校验」→ 复制 **签名密钥（Secret）**
4. 保存

> 不用通知可以跳过 3.3。

---

## 四、本地测试（推荐先跑通）

### 4.1 安装依赖

```bash
pip install -r requirements.txt
```

### 4.2 设置环境变量（Windows PowerShell）

```powershell
# ===== 懂火系统 =====
$env:DH_USERNAME     = "你的懂火账号"
$env:DH_PASSWORD     = "你的懂火密码"

# ===== 导出范围 =====
$env:EXPORT_DAYS     = "30"             # 30=最近30天，0=全部历史

# ===== 交付：飞书云盘（默认） =====
$env:DELIVERY_MODE   = "feishu"
$env:FEISHU_APP_ID       = "cli_xxxxxxxxxxxx"          # 来自 3.1
$env:FEISHU_APP_SECRET   = "xxxxxxxxxxxxxxxxxxxxxxxx"   # 来自 3.1
$env:FEISHU_FOLDER_TOKEN = "fldcnxxxxxxxxxxxxxxxx"      # 来自 3.2
# 可选：通知机器人
$env:FEISHU_WEBHOOK_URL    = "https://open.feishu.cn/open-apis/bot/v2/hook/xxxx-xxxx"
$env:FEISHU_WEBHOOK_SECRET = "xxxxxxxx"                 # 没开启签名校验就留空
```

### 4.3 运行

```bash
python export_chuku.py
```

如果飞书参数没配全，脚本会退化为**本地保存 CSV**，方便你查看数据格式。

---

## 五、部署到 GitHub Actions（云端自动执行）

### Step 1：把代码推到一个 GitHub 仓库

比如仓库名叫 `yourname/donghuo-export`。

### Step 2：在仓库配置 Secrets / Variables

打开仓库页面 → `Settings` → `Secrets and variables` → `Actions`

#### **Secrets** 选项卡 → `New repository secret`，逐个添加：

| Secret 名               | 必填？ | 说明                                             | 示例                              |
| ----------------------- | :---: | ------------------------------------------------ | --------------------------------- |
| `DH_USERNAME`           |  ✅   | 懂火系统登录名                                   | `你的懂火账号`                    |
| `DH_PASSWORD`           |  ✅   | 懂火系统登录密码                                 | `你的懂火密码`                    |
| `FEISHU_APP_ID`         |  ✅   | 飞书自建应用 App ID（cli_开头）                  | `cli_a1b2c3d4e5f6`                |
| `FEISHU_APP_SECRET`     |  ✅   | 飞书自建应用 App Secret                          | `a1b2c3d4e5f6g7h8i9j0`            |
| `FEISHU_FOLDER_TOKEN`   |  ✅   | 出库记录目标文件夹 token（fldcn...，见 3.2）     | `fldcnABCDEFGHIJKLMNOP`           |
| `LD_FOLDER_TOKEN`       |  ✅   | **临调库存**目标文件夹 token（与上面可不同）     | `fldcnXXXXXXXXXXXXXXXX`           |
| `FEISHU_WEBHOOK_URL`    | 可选 | 通知机器人 Webhook（空=不通知）                  | `https://open.feishu.cn/.../hook` |
| `FEISHU_WEBHOOK_SECRET` | 可选 | 机器人签名密钥（没开启签名校验就**不建**这个）   | `abcDefG123`                      |
| `SMTP_HOST` / 等邮件相关 | 可选 | 备用邮件模式（`DELIVERY_MODE=mail` 时才用）     | 详见 README 旧版本               |

#### **Variables** 选项卡 → `New repository variable`（非敏感，方便改）：

| Variable 名           | 说明                                           | 默认值 |
| --------------------- | ---------------------------------------------- | :----: |
| `DELIVERY_MODE`       | 出库记录交付方式：`feishu` 或 `mail`           | `feishu` |
| `EXPORT_DAYS`         | 出库记录导出最近 N 天；0=全部历史              |  `30`  |
| `LD_DELIVERY_MODE`    | **临调库存**交付方式：`feishu` 或 `mail`       | `feishu` |
| `LD_FILTER_SXZHUANTAI`| 临调筛选：状态（已锁/未锁，空=全部）           |  空    |
| `LD_FILTER_HUOQUAN`   | 临调筛选：货权（拥有/待赎，空=全部）           |  空    |
| `LD_FILTER_CANKU`     | 临调筛选：仓库（精确匹配，空=全部）            |  空    |
| `LD_FILTER_PINMIN`    | 临调筛选：品名（空=全部）                      |  空    |

### Step 3：手动触发一次验证

仓库 → `Actions` → 左侧选择 **"自动导出库记录并上传飞书云盘"** → 右侧 `Run workflow` 下拉 → `Run workflow`。

运行结束后：
- 去目标飞书文件夹看 CSV 是否上传成功
- 如果配了机器人 Webhook，群里会收到类似消息：
  ```
  ✅ 出库记录导出 - 2026-08-11 10:00 (共304条)
  导出时间: 2026-08-11 10:00:12
  数据来源: https://erpa.donghuo.vip
  数据量: 304 条
  导出范围: 最近 30 天
  文件名: chuku_20260811_1000.csv
  文件位置: 飞书云盘指定文件夹
  file_token: boxcnXXXXXX
  ```

---

## 六、常见错误排查

| 错误代码 / 关键字                               | 原因 / 解决方法                                                                                         |
| ---------------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| `换 tenant_access_token 失败: code=10013`      | App ID 或 App Secret 写错                                                                             |
| 上传失败 `99991668` / `permission denied`      | 应用没开 `drive:drive` 等权限；或应用版本未发布 / 权限审核没通过；或文件夹没把应用加成协作者（3.2 第3步）|
| 上传失败 `99991663` / `父节点不存在`           | folder_token 写错；或文件夹不属于当前企业；或应用看不到此文件夹（去协作者里确认）                    |
| `文件大小 XMB 超过 upload_all 上限 20MB`       | 最近 30 天的数据量过大；把 `EXPORT_DAYS` 调小；或改用分片上传接口（需要改代码）                       |
| 机器人通知 `code=19021 签名匹配失败`           | FEISHU_WEBHOOK_SECRET 和机器人后台显示的不一致；或者时间戳偏差过大（本机时钟不准）                   |
| 通知 `code=9499` / `msg 无效签名`              | 拼接签名时 `string_to_sign` 有多余换行，本脚本已按标准实现，大概率是 Secret 复制错 / 多了空格       |

---

## 七、修改执行频率

打开 `.github/workflows/export.yml`，修改：

```yaml
on:
  schedule:
    - cron: "0 2 * * *"   # UTC 02:00 = 北京时间 10:00
    - cron: "0 14 * * *"  # UTC 14:00 = 北京时间 22:00
```

cron 是 **UTC 时间**，北京时间 = UTC + 8 小时。

---

## 八、工作原理（简要）

1. **登录**：`POST /controller/admin/c_longin/index`，参数 `u_name` / `u_pass` / `captcha`
2. **验证码识别**：`GET /common/captcha` → 用 `ddddocr` 识别（识别失败自动重试，最多 10 次）
3. **抓数据**：`POST /model/admin/xiaoshou/m_xiaoshou/xjilulist`，参数 `page` / `limit`
   - 响应 JSON：`root`=数据数组，`pgtotal`=总页数，`rtotal`=总条数
   - 数据按"出库日期"降序，达到 `EXPORT_DAYS` 之前的数据即提前终止，避免全量拉取
   - CSV 包含 49 个字段：订单号、客户名称、出库日期、品名、规格、材质、产地、件数、重量、销售单价、销售金额、采购单价、采购金额、利润、供应商、仓库等
4. **生成 CSV**：UTF-8-SIG 编码，Excel 打开中文不乱码
5. **上传到飞书云盘**（默认主模式）：
   - `POST /auth/v3/tenant_access_token/internal` 换应用级 token
   - `POST /drive/v1/files/upload_all`（multipart/form-data，`parent_type=explorer` + `parent_node=folder_token`）
6. **通知**（可选）：通过飞书群自定义机器人 Webhook 发文本消息，附带签名校验

---

## 九、欧冶产能预售导出（独立模块）

> 与懂火脚本完全独立，针对 [欧冶平台](https://www.ouyeel.com) 的产能预售明细数据。

### 9.1 与懂火脚本的区别

| 维度 | 懂火脚本 | 欧冶脚本 |
|------|---------|---------|
| 数据源 | `erpa.donghuo.vip` | `www.ouyeel.com` |
| 反爬 | 无 | **瑞数信息（RS-Anti-Bot）**，必须用真浏览器 |
| 技术方案 | `requests` + ddddocr 识别验证码 | **Playwright** 浏览器自动化 |
| 登录 | 账号密码 + 图形验证码 | **storage_state cookie 复用**（避开 SSO 登录流程） |
| 飞书文件夹 | `LD_FOLDER_TOKEN` | `OUYEEL_FOLDER_TOKEN`（不同文件夹） |

### 9.2 文件结构

```
A/
├── feishu_uploader/
│   ├── export_lindiao.py          # 懂火临调库存
│   └── export_ouyeel.py           # 欧冶产能预售（本节）
└── .github/workflows/
    ├── export_lindiao.yml         # 懂火临调 workflow
    └── export_ouyeel.yml          # 欧冶产能 workflow（本节）
```

### 9.3 飞书新文件夹准备

欧冶数据上传到**独立于懂火**的飞书文件夹：

1. 在飞书云盘新建一个文件夹（如"欧冶产能预售"）
2. 按 [3.2 节](#32-获取目标文件夹的-folder-token) 同样方法获取 folder token
3. 把同一个飞书自建应用加为该文件夹协作者（可编辑）
4. 把 folder token 存为 GitHub Secret `OUYEEL_FOLDER_TOKEN`

> 飞书 App ID / App Secret / 机器人 Webhook 与懂火脚本**共用同一套**，无需重新创建应用。

### 9.4 生成本地登录态 storage_state（关键步骤）

欧冶用 SSO 登录 + 瑞数信息反爬，CI 环境无法自动登录（Playwright 起的浏览器会被瑞数识别为自动化并返回空白页）。采用 **真实浏览器手动登录 → 导出 cookie → 转成 storage_state** 方案，完全避开 Playwright 被检测的问题。

**首次配置 / cookie 过期后刷新**：

#### 步骤 1：用本机 Chrome/Edge 登录欧冶

打开**你自己的** Chrome 或 Edge（不是 Playwright），访问：

```
https://login-ng.ouyeel.com/sso/login?service=https://www.ouyeel.com/
```

手动完成登录（账号密码 / 短信 / 验证码都行）。登录成功后跳到 `www.ouyeel.com` 首页即代表登录成功。

#### 步骤 2：装 Cookie-Editor 扩展（5 秒）

在 Chrome 应用商店装 [Cookie-Editor](https://chromewebstore.google.com/detail/cookie-editor/hlkenndednhfkekhgcdicfkddnncbdoh)（开源、免费、无广告）。Edge 也能装同名扩展。

#### 步骤 3：导出 cookie 为 JSON

1. 登录成功后，停留在 `www.ouyeel.com` 域下任意页面
2. 点浏览器右上角 Cookie-Editor 扩展图标
3. 右下角点 **Export** → 选 **Export as JSON**
4. 自动复制到剪贴板，粘贴到任意文本编辑器，保存为 `cookies.json`（路径随意，比如项目根目录）

#### 步骤 4：转成 storage_state

```bash
# 安装依赖
pip install -r requirements.txt

# 转换（会自动过滤只保留 ouyeel.com 相关 cookie）
python feishu_uploader/cookie_to_state.py cookies.json
```

执行后自动生成两个文件：
- `feishu_uploader/.ouyeel_state.json` — Playwright 用的 storage_state
- `feishu_uploader/.ouyeel_state.json.b64` — base64 编码版（用于 GitHub Secret）

转换脚本会校验：
- ✅ 是否找到会话类 cookie（SESSION/token 等）
- ✅ 是否有 cookie 已过期
- ✅ 是否过滤掉了其他域名的 cookie

> ⚠️ 这两个文件在 `.gitignore` 里，**不会**提交到仓库。

### 9.5 本地测试

```powershell
# 方式一：仅本地保存 CSV（不传飞书，先验证数据抓取是否正常）
$env:DELIVERY_MODE = "local"
python feishu_uploader/export_ouyeel.py

# 方式二：上传飞书云盘
$env:DELIVERY_MODE       = "feishu"
$env:FEISHU_APP_ID       = "cli_xxxxxxxxxxxx"
$env:FEISHU_APP_SECRET   = "xxxxxxxxxxxxxxxxxxxxxxxx"
$env:FEISHU_FOLDER_TOKEN = "fldcnxxxxxxxxxxxxxxxx"   # 欧冶文件夹 token
python feishu_uploader/export_ouyeel.py
```

### 9.6 部署到 GitHub Actions

#### 新增 Secrets

打开仓库 `Settings` → `Secrets and variables` → `Actions` → `Secrets`：

| Secret 名 | 必填 | 说明 |
|-----------|:---:|------|
| `OUYEEL_STORAGE_STATE` | ✅ | 上一步 `.ouyeel_state.json.b64` 的完整内容（base64 字符串） |
| `OUYEEL_FOLDER_TOKEN` | ✅ | 欧冶数据目标飞书文件夹 token |
| `FEISHU_APP_ID` / `FEISHU_APP_SECRET` | ✅ | 复用现有 |
| `FEISHU_WEBHOOK_URL` / `FEISHU_WEBHOOK_SECRET` | 可选 | 复用现有，用于告警 |

#### 新增 Variables（可选）

| Variable 名 | 默认 | 说明 |
|-------------|:----:|------|
| `OUYEEL_PAGE_SIZE` | `50` | 每页条数 |
| `OUYEEL_MAX_PAGES` | `10` | 最大页数防失控（288 条 / 50 = 6 页，10 足够） |
| `OUYEEL_DELIVERY_MODE` | `feishu` | 交付方式 |

#### 手动触发验证

仓库 → `Actions` → 选 **"自动导出欧冶产能预售并上传飞书云盘"** → `Run workflow`

### 9.7 cookie 过期处理

欧冶 cookie 有效期约 **1-7 天**，过期后：

1. 脚本会检测到重定向到 `login-ng.ouyeel.com`，自动通过飞书机器人发告警：
   ```
   ❌ 欧冶产能预售导出失败
   原因: cookie 过期，被重定向到 SSO 登录页
   请按 README 9.4 节重新导出 cookie 并更新 GitHub Secret OUYEEL_STORAGE_STATE
   ```
2. 本地重新走 [9.4 节](#94-生成本地登录态-storage_state关键步骤) 的 4 步流程（用本机 Chrome 登录 → Cookie-Editor 导出 → 跑转换脚本）
3. 把新生成的 `.ouyeel_state.json.b64` 内容更新到 GitHub Secret `OUYEEL_STORAGE_STATE`

### 9.8 欧冶脚本常见错误

| 现象 | 原因 / 解决 |
|------|------------|
| `cookie 过期，被重定向到 SSO 登录页` | storage_state 过期，按 [9.7 节](#97-cookie-过期处理) 刷新 |
| `[XHR] 未拦到数据 JSON` | 页面结构可能变化；首跑会打印所有 XHR URL，把日志反馈给开发者收窄过滤 |
| `未抓到任何数据` | 检查 storage_state 是否登录了正确账号；或欧冶改版需调整 DOM 抽取逻辑 |
| Playwright 安装失败 | 检查 `actions/cache` 是否命中；或 `playwright install --with-deps chromium` 系统依赖缺失 |
| 上传失败 `99991668` | 飞书应用没加为 `OUYEEL_FOLDER_TOKEN` 对应文件夹的协作者（参照 3.2 第 3 步） |

