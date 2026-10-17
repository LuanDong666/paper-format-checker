@echo off
setlocal
title Paper Format Checker - Web Mode
cd /d "%~dp0"
set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
    echo ERROR: Project Python environment was not found.
    goto :end
)
echo Starting web mode at http://127.0.0.1:5000
set PYTHONDONTWRITEBYTECODE=1
set PAPER_CHECKER_HOST=127.0.0.1
start "" /b powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0open_browser.ps1"
"%PYTHON%" -B "%~dp0app.py"
:end
echo.
pause
