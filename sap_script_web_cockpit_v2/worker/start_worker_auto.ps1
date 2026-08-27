$ErrorActionPreference = "Continue"

$WorkerDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CockpitDir = Split-Path -Parent $WorkerDir
$ProjectDir = Split-Path -Parent $CockpitDir
$LogPath = Join-Path $WorkerDir "worker_auto.log"
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

# Ler WORKER_TOKEN do .env uma vez antes do loop
$TokenFromEnv = "change-me"
if (Test-Path $EnvFile) {
    Get-Content $EnvFile | ForEach-Object {
        if ($_ -match '^\s*WORKER_TOKEN\s*=\s*(.+)$') {
            $TokenFromEnv = $matches[1].Trim('"').Trim("'")
        }
    }
}

while ($true) {
    try {
        $ExistingWorkers = Get-CimInstance Win32_Process |
            Where-Object {
                $_.CommandLine -match "worker.py" -and
                $_.CommandLine -match [regex]::Escape($WorkerDir) -and
                $_.ProcessId -ne $PID
            }

        if ($ExistingWorkers) {
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Encontrado worker ativo. A encerrar versoes anteriores para iniciar nova sessao..." | Out-File $LogPath -Append -Encoding UTF8
            foreach ($W in $ExistingWorkers) {
                Stop-Process -Id $W.ProcessId -Force -ErrorAction SilentlyContinue
            }
            Start-Sleep -Seconds 1
        }

        chcp 65001 > $null
        $OutputEncoding = [System.Text.UTF8Encoding]::new($false)

        $env:PYTHONUTF8 = "1"
        $env:PYTHONIOENCODING = "utf-8"
        $env:API_BASE_URL = "http://localhost:8010"
        $env:WORKER_TOKEN = $TokenFromEnv
        $env:SAP_SCRIPT_PROJECT_DIR = $ProjectDir
        $env:SAP_COCKPIT_MODULE = "sap_script_web_cockpit_v2.sap_cockpit_web_ready"
        $env:POLL_SECONDS = "1"

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Iniciando worker (token lido do .env)..." | Out-File $LogPath -Append -Encoding UTF8

        & $PythonExe -u "worker.py" *>> $LogPath
    }
    catch {
        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERRO: $_" | Out-File $LogPath -Append -Encoding UTF8
    }

    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Worker terminou. A reiniciar em 5 segundos..." | Out-File $LogPath -Append -Encoding UTF8

    Start-Sleep -Seconds 5
}
