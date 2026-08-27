$ErrorActionPreference = "Stop"

$ProtocolName = "sap-worker"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$WorkerScript = Join-Path $ProjectDir "worker\start_worker_auto.ps1"
$PowerShellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

if (-not (Test-Path -LiteralPath $PowerShellExe)) {
    $PowerShellExe = "powershell.exe"
}

if (-not (Test-Path -LiteralPath $WorkerScript)) {
    Write-Host "ERRO: O script do worker não foi encontrado em: $WorkerScript" -ForegroundColor Red
    exit 1
}

$Command = "`"$PowerShellExe`" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$WorkerScript`""

$RegistryPath = "HKCU:\Software\Classes\$ProtocolName"
$CommandPath = "$RegistryPath\shell\open\command"

Write-Host "A registar o protocolo '${ProtocolName}://' para o utilizador atual..."

if (-not (Test-Path $RegistryPath)) {
    New-Item -Path $RegistryPath -Force | Out-Null
}
Set-ItemProperty -Path $RegistryPath -Name "(Default)" -Value "URL:SAP Worker Protocol"
Set-ItemProperty -Path $RegistryPath -Name "URL Protocol" -Value ""

if (-not (Test-Path "$RegistryPath\shell")) {
    New-Item -Path "$RegistryPath\shell" -Force | Out-Null
}
if (-not (Test-Path "$RegistryPath\shell\open")) {
    New-Item -Path "$RegistryPath\shell\open" -Force | Out-Null
}
if (-not (Test-Path $CommandPath)) {
    New-Item -Path $CommandPath -Force | Out-Null
}

Set-ItemProperty -Path $CommandPath -Name "(Default)" -Value $Command

Write-Host "Registo concluído com sucesso!" -ForegroundColor Green
Write-Host "Agora podes clicar no botão 'Ligar Worker' no browser." -ForegroundColor Yellow
