@echo off
chcp 65001 >nul
setlocal
title 论文格式检查系统（局域网共享）
cd /d "%~dp0"

set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
    echo [错误] 未找到项目运行环境 .venv。
    echo 请先运行一次“启动系统.bat”完成环境检查。
    goto :end
)

"%PYTHON%" -c "import flask, docx" >nul 2>&1
if errorlevel 1 (
    echo 正在安装项目依赖...
    "%PYTHON%" -m pip install -r "%~dp0requirements.txt"
    if errorlevel 1 (
        echo [错误] 依赖安装失败，请检查网络后重试。
        goto :end
    )
)

echo.
echo ==========================================================
echo  论文格式检查系统 - 局域网共享模式
echo ==========================================================
echo  仅限可信的同一 Wi-Fi/校园网使用。请勿用于公共网络。
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0显示局域网访问地址.ps1"
echo.
echo  如其他电脑无法打开，请双击“配置局域网防火墙.bat”并允许管理员权限，
echo  然后重新运行本文件。
echo  保持本窗口开启；按 Ctrl+C 可停止共享。
echo.

set PYTHONDONTWRITEBYTECODE=1
set PAPER_CHECKER_HOST=0.0.0.0
set PAPER_CHECKER_PORT=5000
start "" /b powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0open_browser.ps1"
"%PYTHON%" -B "%~dp0app.py"

:end
echo.
pause
