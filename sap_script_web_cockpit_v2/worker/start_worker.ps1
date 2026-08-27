$ErrorActionPreference = "Stop"

$WorkerDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CockpitDir = Split-Path -Parent $WorkerDir
$ProjectDir = Split-Path -Parent $CockpitDir
$EnvFile = Join-Path $ProjectDir ".env"
$PythonCandidates = @(
    (Join-Path $WorkerDir ".venv\Scripts\python.exe"),
    (Join-Path $CockpitDir ".venv\Scripts\python.exe"),
    (Join-Path $ProjectDir ".venv\Scripts\python.exe")
)
$PythonExe = $PythonCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

if (-not $PythonExe) {
    throw "Python do ambiente virtual não encontrado. Verifique .venv em '$WorkerDir', '$CockpitDir' ou '$ProjectDir'."
}

Set-Location -LiteralPath $WorkerDir

chcp 65001 > $null
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
$env:SAP_SCRIPT_PROJECT_DIR = $ProjectDir
$env:SAP_COCKPIT_MODULE = "sap_script_web_cockpit_v2.sap_cockpit_web_ready"
$env:POLL_SECONDS = "1"

# Ler API_BASE_URL e WORKER_TOKEN do .env se não estiverem definidos
if (Test-Path $EnvFile) {
    Get-Content $EnvFile | ForEach-Object {
        if ($_ -match '^\s*([A-Z_][A-Z0-9_]*)\s*=\s*(.*)$') {
            $key = $matches[1]; $val = $matches[2].Trim('"').Trim("'")
            if ($key -eq "WORKER_TOKEN" -and -not $env:WORKER_TOKEN) { $env:WORKER_TOKEN = $val }
        }
    }
}

if (-not $env:API_BASE_URL) { $env:API_BASE_URL = "http://localhost:8010" }
if (-not $env:WORKER_TOKEN) { $env:WORKER_TOKEN = "change-me" }

& $PythonExe "worker.py"
