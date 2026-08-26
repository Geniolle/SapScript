[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

function Fail([string]$Message, [int]$Code = 1) {
    Write-Error $Message
    exit $Code
}

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$EnvPath = Join-Path $ProjectRoot ".env"
$EncryptedPath = Join-Path $ProjectRoot ".env.enc"
$DefaultKeyPath = Join-Path $env:APPDATA "sops\age\keys.txt"
$KeyPath = if ($env:SOPS_AGE_KEY_FILE) { $env:SOPS_AGE_KEY_FILE } else { $DefaultKeyPath }
$TemporaryPath = Join-Path $ProjectRoot ".env.restore.tmp"

if (-not (Test-Path -LiteralPath $EncryptedPath -PathType Leaf)) {
    Fail "Backup .env.enc não encontrado na raiz do projeto." 2
}
if ((Test-Path -LiteralPath $EnvPath) -and -not $Force) {
    Fail "O ficheiro .env já existe. Use -Force apenas se pretende substituí-lo pelo backup encriptado." 3
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
    & $Sops.Source --decrypt --input-type dotenv --output-type dotenv --output $TemporaryPath $EncryptedPath
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $TemporaryPath -PathType Leaf)) {
        Fail "SOPS não conseguiu desencriptar o backup." 6
    }
    Move-Item -LiteralPath $TemporaryPath -Destination $EnvPath -Force
    Write-Host "Ficheiro local .env restaurado com sucesso. Ele permanece ignorado pelo Git."
}
finally {
    if (Test-Path -LiteralPath $TemporaryPath) {
        Remove-Item -LiteralPath $TemporaryPath -Force
    }
}
