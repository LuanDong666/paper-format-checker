$ErrorActionPreference = 'Stop'

# Probe the local health endpoint using Windows PowerShell-compatible APIs.
$ready = $false
for ($attempt = 0; $attempt -lt 30; $attempt++) {
    try {
        $response = Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:5000/api/health' -TimeoutSec 2
        if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300) {
            $ready = $true
            break
        }
    } catch {
        # Flask may still be importing dependencies; retry briefly.
    }
    Start-Sleep -Seconds 1
}

if (-not $ready) {
    Write-Warning 'Paper Format Checker did not become ready at http://127.0.0.1:5000. Read the Python error in the startup window.'
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
        'http://127.0.0.1:5000/'
    )
    Start-Process -FilePath $edge -ArgumentList $arguments
} else {
    Start-Process 'http://127.0.0.1:5000/'
}
