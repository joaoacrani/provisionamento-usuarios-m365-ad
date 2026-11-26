# bootstrap.ps1
# Carregado (dot-source) por main.ps1 e web.ps1: importa os módulos, lê o config.json
# e inicializa o log. Define $Config no escopo de quem chamou.

$modulePaths = @(
    "modules\Logger.psm1",
    "modules\Validator.psm1",
    "modules\M365User.psm1",
    "modules\ADUser.psm1",
    "modules\Report.psm1",
    "modules\Provisioning.psm1",
    "modules\Offboarding.psm1",
    "modules\Signature.psm1",
    "modules\UI.psm1"
)

foreach ($mod in $modulePaths) {
    $fullPath = Join-Path $PSScriptRoot $mod
    if (-not (Test-Path $fullPath)) {
        Write-Host "ERRO FATAL: Módulo não encontrado: $fullPath" -ForegroundColor Red
        exit 1
    }
    Import-Module $fullPath -Force -DisableNameChecking
}

$configPath = Join-Path $PSScriptRoot "config.json"
if (-not (Test-Path $configPath)) {
    Write-Host "ERRO FATAL: config.json não encontrado em $PSScriptRoot" -ForegroundColor Red
    Write-Host "Copie config.example.json para config.json e preencha com os dados do seu ambiente." -ForegroundColor Yellow
    exit 1
}
$Config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

# Tempo máximo (s) aguardando a caixa no Exchange antes de adicionar às listas
$DistributionListWaitSeconds = if ($Config.PSObject.Properties.Name -contains 'DistributionListWaitSeconds') {
    [int]$Config.DistributionListWaitSeconds
} else { 300 }

Initialize-Logger -LogDir (Join-Path $PSScriptRoot "logs")
