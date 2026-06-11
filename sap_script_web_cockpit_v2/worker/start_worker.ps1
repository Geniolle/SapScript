$ErrorActionPreference = "Stop"

$ProjectDir = "C:\workspace\sap-script"
$WorkerDir = "$ProjectDir\sap_script_web_cockpit_v2\worker"

Set-Location $WorkerDir

& ".\.venv\Scripts\Activate.ps1"

chcp 65001 > $null
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
$env:SAP_SCRIPT_PROJECT_DIR = $ProjectDir
$env:SAP_COCKPIT_MODULE = "sap_script_web_cockpit_v2.sap_cockpit_web_ready"
$env:POLL_SECONDS = "1"

# Ler API_BASE_URL e WORKER_TOKEN do .env se não estiverem definidos
$EnvFile = "$ProjectDir\.env"
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

python worker.py
