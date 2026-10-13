$ErrorActionPreference = "Continue"

$ProjectDir = $PSScriptRoot
$WorkerDir = Join-Path $ProjectDir "worker"
$LogPath = Join-Path $ProjectDir "start_all.log"

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Iniciando SAP Script Web Cockpit..." | Out-File $LogPath -Append -Encoding UTF8

cd $ProjectDir

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] O Cockpit Web esta hospedado na Oracle Cloud (140.238.123.30)." | Out-File $LogPath -Append -Encoding UTF8
"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Docker local nao e necessario." | Out-File $LogPath -Append -Encoding UTF8

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
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "$WorkerDir\start_worker_auto.ps1"
) -WindowStyle Hidden

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Processo concluido." | Out-File $LogPath -Append -Encoding UTF8
