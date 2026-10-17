$ErrorActionPreference = 'Stop'
Start-Sleep -Seconds 2

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
