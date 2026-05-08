$ErrorActionPreference = "Continue"

$WorkerDir = "C:\workspace\sap-script\sap_script_web_cockpit_v2\worker"
$LogPath = Join-Path $WorkerDir "worker_auto.log"

cd $WorkerDir

while ($true) {
    try {
        $ExistingWorker = Get-CimInstance Win32_Process |
            Where-Object {
                $_.CommandLine -match "worker.py" -and
                $_.CommandLine -match [regex]::Escape($WorkerDir)
            }

        if ($ExistingWorker) {
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Worker já está ativo. A encerrar launcher duplicado." | Out-File $LogPath -Append -Encoding UTF8
            exit 0
        }

        & ".\.venv\Scripts\Activate.ps1"

        $env:API_BASE_URL = "http://localhost:8010"
        $env:WORKER_TOKEN = "change-me"
        $env:SAP_SCRIPT_PROJECT_DIR = "C:\workspace\sap-script"
        $env:SAP_COCKPIT_MODULE = "sap_script_web_cockpit_v2.sap_cockpit_web_ready"
        $env:POLL_SECONDS = "1"

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Iniciando worker..." | Out-File $LogPath -Append -Encoding UTF8

        python worker.py *>> $LogPath
    }
    catch {
        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERRO: $_" | Out-File $LogPath -Append -Encoding UTF8
    }

    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Worker terminou. A reiniciar em 5 segundos..." | Out-File $LogPath -Append -Encoding UTF8

    Start-Sleep -Seconds 5
}
