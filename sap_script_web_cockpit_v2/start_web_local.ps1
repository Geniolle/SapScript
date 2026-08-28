$ErrorActionPreference = "Stop"

$ProjectDir = $PSScriptRoot
$Python = Join-Path $ProjectDir ".venv\Scripts\python.exe"
$EnvFile = Join-Path (Split-Path $ProjectDir -Parent) ".env"
$FiBridgePython = Join-Path (Split-Path $ProjectDir -Parent) ".venv-rfc\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $Python -PathType Leaf)) {
    throw "Ambiente Python não encontrado em $Python."
}

if (-not (Test-Path -LiteralPath $EnvFile -PathType Leaf)) {
    throw "Ficheiro de configuração não encontrado em $EnvFile."
}

Set-Location -LiteralPath $ProjectDir
$env:DATA_DIR = Join-Path $ProjectDir "data"
$env:UPLOADS_DIR = Join-Path $ProjectDir "uploads"
$env:JIRA_DOWNLOAD_DIR_CONTAINER = Join-Path $ProjectDir "jira_downloads"
$env:WORKFLOW_PYTHON_EXEC = $Python
if (Test-Path -LiteralPath $FiBridgePython -PathType Leaf) {
    $env:SAP_FI_BRIDGE_PYTHON = $FiBridgePython
}
Write-Host "Cockpit local (sem Docker): http://localhost:8000"
& $Python -m uvicorn web_api.main:app --host 127.0.0.1 --port 8000
