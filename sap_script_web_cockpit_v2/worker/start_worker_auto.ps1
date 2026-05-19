$ErrorActionPreference = "Continue"

$WorkerDir = "C:\workspace\sap-script\sap_script_web_cockpit_v2\worker"
$LogPath = Join-Path $WorkerDir "worker_auto.log"

cd $WorkerDir

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

        & ".\.venv\Scripts\Activate.ps1"

        chcp 65001 > $null
        $OutputEncoding = [System.Text.UTF8Encoding]::new($false)

        $env:PYTHONUTF8 = "1"
        $env:PYTHONIOENCODING = "utf-8"
        $env:API_BASE_URL = "http://localhost:8010"
        $env:WORKER_TOKEN = "change-me"
        $env:SAP_SCRIPT_PROJECT_DIR = "C:\workspace\sap-script"
        $env:SAP_COCKPIT_MODULE = "sap_script_web_cockpit_v2.sap_cockpit_web_ready"
        $env:POLL_SECONDS = "1"

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Iniciando worker..." | Out-File $LogPath -Append -Encoding UTF8

        python -u worker.py *>> $LogPath
    }
    catch {
        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERRO: $_" | Out-File $LogPath -Append -Encoding UTF8
    }

    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Worker terminou. A reiniciar em 5 segundos..." | Out-File $LogPath -Append -Encoding UTF8

    Start-Sleep -Seconds 5
}
