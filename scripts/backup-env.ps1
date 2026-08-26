[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Fail([string]$Message, [int]$Code = 1) {
    Write-Error $Message
    exit $Code
}

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$EnvPath = Join-Path $ProjectRoot ".env"
$EncryptedPath = Join-Path $ProjectRoot ".env.enc"
$ConfigPath = Join-Path $ProjectRoot ".sops.yaml"
$DefaultKeyPath = Join-Path $env:APPDATA "sops\age\keys.txt"
$KeyPath = if ($env:SOPS_AGE_KEY_FILE) { $env:SOPS_AGE_KEY_FILE } else { $DefaultKeyPath }
$TemporaryPath = Join-Path $ProjectRoot ".env.enc.tmp"

if (-not (Test-Path -LiteralPath $EnvPath -PathType Leaf)) {
    Fail "Ficheiro .env não encontrado na raiz do projeto." 2
}
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    Fail "Configuração .sops.yaml não encontrada." 3
}
if (-not (Test-Path -LiteralPath $KeyPath -PathType Leaf)) {
    Fail "Chave age não encontrada fora do repositório. Configure SOPS_AGE_KEY_FILE ou restaure %APPDATA%\sops\age\keys.txt." 4
}

$Sops = Get-Command sops -ErrorAction SilentlyContinue
if (-not $Sops) {
    Fail "SOPS não está disponível no PATH. Instale Mozilla.SOPS e abra um novo PowerShell." 5
}

try {
    $env:SOPS_AGE_KEY_FILE = $KeyPath
    & $Sops.Source --encrypt --input-type dotenv --output-type dotenv --output $TemporaryPath $EnvPath
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $TemporaryPath -PathType Leaf)) {
        Fail "SOPS não conseguiu criar o backup encriptado." 6
    }
    Move-Item -LiteralPath $TemporaryPath -Destination $EncryptedPath -Force
    Write-Host "Backup encriptado criado: .env.enc"
}
finally {
    if (Test-Path -LiteralPath $TemporaryPath) {
        Remove-Item -LiteralPath $TemporaryPath -Force
    }
}
