@echo off
setlocal
title Paper Format Checker
cd /d "%~dp0"
set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
  echo Python environment not found.
  pause
  exit /b 1
)
set "PAPER_CHECKER_HOST=127.0.0.1"
set "PAPER_CHECKER_PORT=5000"
set "PYTHONDONTWRITEBYTECODE=1"
echo Starting Paper Format Checker and opening the browser...
"%PYTHON%" -B "%~dp0run_paper_checker.py"
echo.
echo The server has stopped. Press any key to close this window.
pause

