@echo off
REM ==========================================================================
REM 临调库存导出 → 飞书云盘 启动脚本
REM 由 Windows 计划任务定时调用
REM ==========================================================================
setlocal enabledelayedexpansion

cd /d "%~dp0.."

REM ===== 必要配置 =====
REM  凭证请放在系统环境变量或 run_upload.bat 同目录下的 run_upload.local.bat（后者勿进仓库）。
REM  本文件只做示例，真实 DH_PASSWORD/FEISHU_APP_SECRET 不要硬编码进 git.
set DELIVERY_MODE=feishu
if "%DH_USERNAME%"=="" set DH_USERNAME=HONG@shihe.donghuo
if "%DH_PASSWORD%"=="" set DH_PASSWORD=
if "%FEISHU_APP_ID%"=="" set FEISHU_APP_ID=cli_aaf0ce1e9ef89d27
if "%FEISHU_APP_SECRET%"=="" set FEISHU_APP_SECRET=
if "%FEISHU_FOLDER_TOKEN%"=="" set FEISHU_FOLDER_TOKEN=O8i6fsf5dlQtX3ds3D2cjGQtnVh
if exist "%~dp0run_upload.local.bat" call "%~dp0run_upload.local.bat"

REM ===== 可选筛选（留空=全部）=====
REM 状态：已锁 / 未锁
set FILTER_SXZHUANTAI=
REM 货权：拥有 / 待赎
set FILTER_HUOQUAN=
REM 仓库（精确匹配，如 仲鼎库）
set FILTER_CANKU=
REM 品名（如 酸洗）
set FILTER_PINMIN=

REM ===== 可选：飞书机器人通知 =====
set FEISHU_WEBHOOK_URL=
set FEISHU_WEBHOOK_SECRET=

REM ===== 运行日志 =====
set LOG_FILE=%~dp0run_upload_%date:~0,4%%date:~5,2%%date:~8,2%.log
echo [%date% %time%] 开始执行... > "%LOG_FILE%"
echo [%date% %time%] WORKDIR=%cd% >> "%LOG_FILE%"

python feishu_uploader\export_lindiao.py >> "%LOG_FILE%" 2>&1
set EXIT_CODE=%ERRORLEVEL%

echo [%date% %time%] 执行结束，退出码=%EXIT_CODE% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

exit /b %EXIT_CODE%
