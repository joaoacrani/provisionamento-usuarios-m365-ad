#Requires -Version 7.0
<#
.SYNOPSIS
    Setup inicial — Instala módulos, verifica pré-requisitos e guia a configuração.
    Execute este script UMA VEZ antes de usar o sistema.

.NOTES
    Execute como Administrador para instalar módulos globalmente,
    ou sem privilégios para instalar no escopo do usuário atual.
#>

Write-Host @"

  ╔════════════════════════════════════════════════════╗
  ║    SETUP — Sistema de Provisionamento v1.0.0      ║
  ║    contoso.com                               ║
  ╚════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan

# ── 1. Verifica PowerShell 7+ ──────────────────────────────────
Write-Host "  [1/5] Verificando versão do PowerShell..." -ForegroundColor Yellow
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "  ERRO: PowerShell 7+ obrigatório. Versão atual: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    Write-Host "  Baixe em: https://aka.ms/powershell" -ForegroundColor Cyan
    exit 1
}
Write-Host "  OK — PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Green

# ── 2. Define escopo de instalação ────────────────────────────
$isAdmin   = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
              [Security.Principal.WindowsBuiltInRole]::Administrator)
$scope     = if ($isAdmin) { "AllUsers" } else { "CurrentUser" }
Write-Host "`n  Instalando módulos no escopo: $scope" -ForegroundColor Gray

# ── 3. Instala Microsoft.Graph ─────────────────────────────────
Write-Host "`n  [2/5] Verificando Microsoft.Graph..." -ForegroundColor Yellow
$graphMod  = Get-Module -ListAvailable -Name "Microsoft.Graph" | Sort-Object Version -Descending | Select-Object -First 1

if (-not $graphMod) {
    Write-Host "  Instalando Microsoft.Graph (pode demorar alguns minutos)..." -ForegroundColor Cyan
    Install-Module Microsoft.Graph -Scope $scope -Force -AllowClobber
    Write-Host "  Microsoft.Graph instalado com sucesso." -ForegroundColor Green
} elseif ($graphMod.Version -lt [Version]"2.0.0") {
    Write-Host "  Atualizando Microsoft.Graph de $($graphMod.Version) para versão mais recente..." -ForegroundColor Yellow
    Update-Module Microsoft.Graph -Scope $scope -Force
    Write-Host "  Microsoft.Graph atualizado." -ForegroundColor Green
} else {
    Write-Host "  OK — Microsoft.Graph v$($graphMod.Version) já instalado." -ForegroundColor Green
}

# ── 4. Verifica módulo ActiveDirectory (RSAT) ─────────────────
Write-Host "`n  [3/5] Verificando módulo ActiveDirectory..." -ForegroundColor Yellow
$adMod = Get-Module -ListAvailable -Name "ActiveDirectory" | Select-Object -First 1

if (-not $adMod) {
    Write-Host "  AVISO: Módulo ActiveDirectory não encontrado." -ForegroundColor Yellow
    Write-Host "  Este módulo faz parte do RSAT (Remote Server Administration Tools)." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Para instalar no Windows 10/11:" -ForegroundColor Cyan
    Write-Host "  Configurações > Aplicativos > Recursos Opcionais > 'RSAT: Active Directory'" -ForegroundColor White
    Write-Host ""
    Write-Host "  Ou via PowerShell (como Admin):" -ForegroundColor Cyan
    Write-Host "  Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'" -ForegroundColor White
    Write-Host ""
    Write-Host "  Se estiver executando em um Domain Controller, o módulo já está disponível." -ForegroundColor Gray
} else {
    Write-Host "  OK — ActiveDirectory v$($adMod.Version) disponível." -ForegroundColor Green
}

# ── 5. Registra App no Entra ID (instruções) ──────────────────
Write-Host "`n  [4/5] Configuração do App Registration no Entra ID:" -ForegroundColor Yellow
Write-Host @"

  Para autenticar no Microsoft Graph, você precisa de um App Registration.
  Siga os passos abaixo no portal Azure (portal.azure.com):

  1. Vá em: Microsoft Entra ID > App registrations > New registration
  2. Nome: UserProvisioning-Script (ou qualquer nome)
  3. Supported account types: Accounts in this organizational directory only
  4. Redirect URI: http://localhost (tipo: Public client/native)
  5. Clique em Register

  Após criar:
  6. Copie o "Application (client) ID"
  7. Copie o "Directory (tenant) ID"
  8. Vá em API permissions > Add a permission > Microsoft Graph > Delegated:
       - User.ReadWrite.All
       - Directory.ReadWrite.All
       - Organization.Read.All
  9. Clique em "Grant admin consent"

  Depois, edite o arquivo config.json com:
  - "TenantId": "seu-tenant-id"
  - "ClientId": "seu-client-id"
  - "ADServer": "nome-do-seu-domain-controller"

"@ -ForegroundColor Gray

# ── 6. Valida config.json ──────────────────────────────────────
Write-Host "  [5/5] Verificando config.json..." -ForegroundColor Yellow
$configPath = Join-Path $PSScriptRoot "config.json"
if (Test-Path $configPath) {
    $config = Get-Content $configPath -Raw | ConvertFrom-Json
    $issues = @()
    if ($config.TenantId -eq "SEU-TENANT-ID-AQUI")  { $issues += "TenantId não configurado" }
    if ($config.ClientId -eq "SEU-CLIENT-ID-AQUI")  { $issues += "ClientId não configurado" }
    if ($config.ADServer -eq "SEU-DC-AQUI")          { $issues += "ADServer não configurado" }

    if ($issues.Count -gt 0) {
        Write-Host "  PENDENTE — Edite config.json:" -ForegroundColor Yellow
        $issues | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
    } else {
        Write-Host "  OK — config.json configurado." -ForegroundColor Green
    }
} else {
    Write-Host "  AVISO: config.json não encontrado em $PSScriptRoot" -ForegroundColor Red
}

Write-Host @"

  ══════════════════════════════════════════════════════
  Setup concluído!

  Próximos passos:
  1. Configure o config.json com TenantId, ClientId e ADServer
  2. Execute: pwsh -File main.ps1

  Logs serão salvos em: .\logs\
  ══════════════════════════════════════════════════════

"@ -ForegroundColor Cyan
