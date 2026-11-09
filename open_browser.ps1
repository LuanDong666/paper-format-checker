$ErrorActionPreference = 'Stop'

# Probe the local health endpoint without the system proxy. Windows PowerShell 5.1
# does not support Invoke-WebRequest -NoProxy, so use WebRequest directly.
$port = 5000
if ($env:PAPER_CHECKER_PORT) {
    $parsedPort = 0
    if (-not [int]::TryParse($env:PAPER_CHECKER_PORT, [ref]$parsedPort) -or
        $parsedPort -lt 1 -or $parsedPort -gt 65535) {
        throw 'PAPER_CHECKER_PORT must be a number between 1 and 65535.'
    }
    $port = $parsedPort
}
$url = "http://127.0.0.1:$port/"
$healthUrl = "${url}api/health"
$ready = $false
for ($attempt = 0; $attempt -lt 30; $attempt++) {
    try {
        $request = [System.Net.WebRequest]::Create($healthUrl)
        $request.Proxy = $null
        $request.Timeout = 2000
        $response = $request.GetResponse()
        try {
            if ([int]$response.StatusCode -ge 200 -and [int]$response.StatusCode -lt 300) {
                $ready = $true
                break
            }
        } finally {
            $response.Close()
        }
    } catch {
        # Flask may still be importing dependencies; retry briefly.
    }
    Start-Sleep -Seconds 1
}

if (-not $ready) {
    Write-Warning "Paper Format Checker did not become ready at $url. Read the Python error in the startup window."
    exit 1
}

$edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
if (-not (Test-Path -LiteralPath $edge)) {
    $edge = 'C:\Program Files\Microsoft\Edge\Application\msedge.exe'
}

if (Test-Path -LiteralPath $edge) {
    # This dedicated, local-only profile avoids extensions and proxy state in a running browser.
    $profile = Join-Path $PSScriptRoot '.local-browser-profile'
    $arguments = @(
        '--new-window',
        "--user-data-dir=$profile",
        '--proxy-server=direct://',
        '--proxy-bypass-list=localhost;127.0.0.1',
        $url
    )
    Start-Process -FilePath $edge -ArgumentList $arguments
} else {
    Start-Process $url
}
