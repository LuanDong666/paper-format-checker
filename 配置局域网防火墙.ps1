$ErrorActionPreference = 'Stop'

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Start-Process -FilePath 'powershell.exe' -Verb RunAs -Wait -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`""
    )
    exit $LASTEXITCODE
}

& netsh advfirewall firewall delete rule name='Paper Format Checker LAN (TCP 5000)' | Out-Null
& netsh advfirewall firewall add rule name='Paper Format Checker LAN (TCP 5000)' dir=in action=allow protocol=TCP localport=5000 profile=private
if ($LASTEXITCODE -ne 0) { throw 'Windows 防火墙规则写入失败。' }

Write-Host ''
Write-Host '已添加专用网络规则：Paper Format Checker LAN (TCP 5000)。' -ForegroundColor Green
Write-Host '现在可重新运行“局域网共享启动.bat”。' -ForegroundColor Cyan
Read-Host '按 Enter 关闭'
