$ErrorActionPreference = 'Stop'

try {
    $config = Get-NetIPConfiguration |
        Where-Object { $_.IPv4DefaultGateway -and $_.IPv4Address.IPAddress -notlike '169.254*' } |
        Select-Object -First 1
    $ip = $config.IPv4Address |
        Where-Object { $_.IPAddress -notlike '169.254*' } |
        Select-Object -First 1 -ExpandProperty IPAddress
} catch {
    $ip = $null
}

if ($ip) {
    Write-Host "  请在其他同一网络的电脑浏览器打开：" -ForegroundColor Cyan
    Write-Host "  http://$($ip):5000" -ForegroundColor Green
} else {
    Write-Host "  未能自动读取局域网 IPv4 地址。" -ForegroundColor Yellow
    Write-Host "  请运行 ipconfig，找到当前 Wi-Fi/以太网的 IPv4 地址后访问：" -ForegroundColor Yellow
    Write-Host "  http://该IPv4地址:5000" -ForegroundColor Green
}
