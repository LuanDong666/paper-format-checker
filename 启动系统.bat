@echo off
setlocal
title Paper Format Checker
cd /d "%~dp0"

set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
    echo ERROR: Project Python environment was not found.
    echo Run the setup commands in README.md first.
    goto :end
)

"%PYTHON%" -c "import flask, docx" >nul 2>&1
if errorlevel 1 (
    echo Installing required Python packages...
    "%PYTHON%" -m pip install -r "%~dp0requirements.txt"
    if errorlevel 1 goto :failed
)

echo Starting Paper Format Checker web application...
set PYTHONDONTWRITEBYTECODE=1
start "" /b powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0open_browser.ps1"
set PAPER_CHECKER_HOST=127.0.0.1
"%PYTHON%" -B "%~dp0app.py"
goto :end

:failed
echo ERROR: Dependency installation failed. Check your network connection.

:end
echo.
pause
