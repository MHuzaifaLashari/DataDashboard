$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$python = Join-Path $PSScriptRoot '.venv313\Scripts\python.exe'
$url = 'http://localhost:8501'
function Test-AppReady {
    try { return (Invoke-WebRequest "$url/_stcore/health" -UseBasicParsing -TimeoutSec 2).StatusCode -eq 200 } catch { return $false }
}
try {
    if (-not (Test-AppReady)) {
        if (-not (Test-Path -LiteralPath $python)) { throw "Python environment missing: $python" }
        $env:TEMP = Join-Path $PSScriptRoot '.tmp'
        $env:TMP = $env:TEMP
        New-Item -ItemType Directory -Path $env:TEMP -Force | Out-Null
        $process = Start-Process -FilePath $python -ArgumentList @('-m', 'streamlit', 'run', 'app.py', '--server.address', 'localhost', '--server.port', '8501', '--server.headless', 'true', '--browser.gatherUsageStats', 'false') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -RedirectStandardOutput (Join-Path $PSScriptRoot 'streamlit.log') -RedirectStandardError (Join-Path $PSScriptRoot 'streamlit-error.log') -PassThru
        $ready = $false
        for ($i = 0; $i -lt 60; $i++) {
            if (Test-AppReady) { $ready = $true; break }
            if ($process.HasExited) { throw 'App stopped. Check streamlit-error.log for details.' }
            Start-Sleep -Milliseconds 500
        }
        if (-not $ready) { throw 'App startup timed out. Check streamlit-error.log for details.' }
    }
    Start-Process $url
} catch {
    Write-Host $_ -ForegroundColor Red
    exit 1
}
