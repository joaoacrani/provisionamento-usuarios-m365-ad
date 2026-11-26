# criar-atalho.ps1
# Cria o atalho "Criar Usuários" na Área de Trabalho e no Menu Iniciar.
# Rode uma vez (clique com o botão direito → Executar com PowerShell). Se mover a pasta, rode de novo.
#
#   pwsh -File criar-atalho.ps1          → atalho normal
#   pwsh -File criar-atalho.ps1 -Demo    → também cria "Criar Usuários (Demo)"

param([switch]$Demo)

$ErrorActionPreference = "Stop"

$pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
if (-not $pwsh) { throw "PowerShell 7 (pwsh.exe) não encontrado. Instale-o antes de criar o atalho." }

$webScript = Join-Path $PSScriptRoot "web.ps1"
if (-not (Test-Path $webScript)) { throw "web.ps1 não encontrado em $PSScriptRoot" }

$destinos = @(
    [Environment]::GetFolderPath('Desktop'),
    [Environment]::GetFolderPath('Programs')     # Menu Iniciar
)

$shell = New-Object -ComObject WScript.Shell

function New-Shortcut([string]$Folder, [string]$Name, [string]$ExtraArgs, [string]$Description) {
    $lnk = $shell.CreateShortcut((Join-Path $Folder "$Name.lnk"))
    $lnk.TargetPath       = $pwsh
    $lnk.Arguments        = "-NoProfile -ExecutionPolicy Bypass -File `"$webScript`" $ExtraArgs".Trim()
    $lnk.WorkingDirectory = $PSScriptRoot
    $lnk.IconLocation     = "$env:SystemRoot\System32\imageres.dll,74"
    $lnk.Description      = $Description
    $lnk.WindowStyle      = 1
    $lnk.Save()
    Write-Host "  [OK] $(Join-Path $Folder "$Name.lnk")" -ForegroundColor Green
}

foreach ($d in $destinos) {
    New-Shortcut $d "Criar Usuários" "" "Provisionamento de usuários M365 + AD (interface web)"
    if ($Demo) { New-Shortcut $d "Criar Usuários (Demo)" "-Demo" "Interface de provisionamento em modo demonstração — não cria nada" }
}

Write-Host "`n  Pronto. Use o atalho 'Criar Usuários' na Área de Trabalho ou no Menu Iniciar." -ForegroundColor Cyan
