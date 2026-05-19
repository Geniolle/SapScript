$ErrorActionPreference = "Continue"

$ProjectDir = "C:\workspace\sap-script\sap_script_web_cockpit_v2"
$WorkerDir = Join-Path $ProjectDir "worker"
$LogPath = Join-Path $ProjectDir "start_all.log"

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Iniciando SAP Script Web Cockpit..." | Out-File $LogPath -Append -Encoding UTF8

cd $ProjectDir

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Subindo Docker Compose..." | Out-File $LogPath -Append -Encoding UTF8
docker compose up -d --build *>> $LogPath

Start-Sleep -Seconds 5

$ExistingWorker = Get-CimInstance Win32_Process |
    Where-Object {
        $_.CommandLine -match "worker.py" -and
        $_.CommandLine -match [regex]::Escape($WorkerDir)
    }

if ($ExistingWorker) {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Worker ja esta ativo. Nada a fazer." | Out-File $LogPath -Append -Encoding UTF8
    exit 0
}

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Iniciando worker automatico..." | Out-File $LogPath -Append -Encoding UTF8

Start-Process powershell.exe -ArgumentList @(
    "-NoExit",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$WorkerDir\start_worker_auto.ps1`""
) -WindowStyle Minimized

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Processo concluido." | Out-File $LogPath -Append -Encoding UTF8
