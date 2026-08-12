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
