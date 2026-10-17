@echo off
chcp 65001 >nul
setlocal
title 配置论文格式检查系统局域网访问

echo 将申请管理员权限，只添加一条“专用网络”入站规则：
echo 允许 TCP 5000 端口用于论文格式检查系统的局域网访问。
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0配置局域网防火墙.ps1"
echo.
echo 若上方显示“Ok”，防火墙规则已设置完成。
pause
