@echo off
setlocal
title Paper Format Checker - Desktop Backup Mode
cd /d "%~dp0"
set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
    echo ERROR: Project Python environment was not found.
    goto :end
)
"%PYTHON%" -B "%~dp0desktop_app.py"
:end
echo.
pause
