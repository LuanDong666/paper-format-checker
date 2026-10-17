@echo off
chcp 65001 >nul
setlocal
title 论文格式检查系统（临时分享）
cd /d "%~dp0"

set PYTHON=.venv\Scripts\python.exe
if not exist "%PYTHON%" (
    echo [错误] 未找到 .venv。请先运行“启动系统.bat”完成环境检查。
    pause
    exit /b 1
)
where cpolar >nul 2>&1
if errorlevel 1 (
    echo [错误] 未安装 cpolar，无法创建临时访问链接。
    echo 请从 https://www.cpolar.com/ 安装并完成账号授权。
    pause
    exit /b 1
)

echo 此模式会把本机服务临时暴露到公网。
echo 请勿上传含个人信息、未公开论文或 API Key 的文档。
choice /M "确认继续"
if errorlevel 2 exit /b 0

set PYTHONDONTWRITEBYTECODE=1
start "论文格式检查服务" /B "%PYTHON%" -B app.py
timeout /t 2 /nobreak >nul
echo 服务已启动。cpolar 输出中的 URL 可供临时访问；按 Ctrl+C 结束隧道。
cpolar http 5000
