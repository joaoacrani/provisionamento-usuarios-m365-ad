#Requires -Version 7.0
<#
.SYNOPSIS
    Sistema Corporativo de Provisionamento de Usuários (versão terminal)
    Empresa: Contoso | M365: contoso.com | AD: corp.contoso.local

.DESCRIPTION
    Automatiza criação de usuários no Microsoft 365 (Entra ID) e Active Directory local,
    operando de forma INDEPENDENTE, sem qualquer sincronização híbrida.
    Suporte a DC primário (DC01/PDC) com fallback automático para DC02.

    Todas as escolhas são coletadas primeiro (com opção de voltar entre etapas) e
    revisadas antes de qualquer criação. Para a interface no navegador use web.ps1.

.NOTES
    Versão     : 2.0.0
    PowerShell : 7.x
    Módulos    : Microsoft.Graph, ActiveDirectory, ExchangeOnlineManagement
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "bootstrap.ps1")

function Invoke-SignatureStep {
    <#
    .SYNOPSIS
        Pergunta se deve gerar a assinatura de e-mail; salva em Downloads e, opcionalmente, na rede.
    #>
    param([hashtable]$Plan)

    $sig = Get-SignatureConfig -Config $Config
    if (-not $sig.Templates.Count) { return }

    $ans = (Read-Host "  Gerar a assinatura de e-mail agora? (S/N)").Trim().ToUpper()
    if ($ans -ne 'S') { return }

    $suggested = Get-SignatureTemplateForOU -Templates $sig.Templates -OUPath $Plan.OUPath
    Write-Host ""
    foreach ($t in $sig.Templates) {
        $mark = if ($t.Key -eq $suggested.Key) { "●" } else { " " }
        Write-Host "  $mark [$($t.Key)] $($t.Label)" -ForegroundColor White
    }
    $key = (Read-Host "  Modelo [ENTER = $($suggested.Key)]").Trim().ToUpper()
    $tpl = if ($key) { $sig.Templates | Where-Object { $_.Key -eq $key } | Select-Object -First 1 } else { $suggested }
    if (-not $tpl) { Write-Host "  [ERRO] Modelo inválido. Assinatura não gerada." -ForegroundColor Red; return }

    $name = (Read-Host "  Nome na assinatura [ENTER = $($Plan.FullName)]").Trim()
    if (-not $name) { $name = $Plan.FullName }

    try   { $downloads = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path }
    catch { $downloads = Join-Path $env:USERPROFILE "Downloads" }
    $out = Join-Path $downloads "$($Plan.Login.ToLower()).png"

    try {
        Write-Host "  Gerando no CorelDRAW (~15 s)..." -ForegroundColor Gray
        New-EmailSignature -TemplatePath $tpl.Path -FullName $name -Email $Plan.Email -OutFile $out -ProgId $sig.ProgId | Out-Null
        Write-Host "  [OK] Assinatura salva em: $out" -ForegroundColor Green
        Start-Process $out

        $net = (Read-Host "  Salvar cópia em '$($tpl.SaveFolder)'? (S/N)").Trim().ToUpper()
        if ($net -eq 'S') {
            $r = Save-SignatureToNetwork -SourceFile $out -Folder $tpl.SaveFolder -Login $Plan.Login
            Write-Host "  $($r.Message)" -ForegroundColor $(if ($r.Ok) { "Green" } else { "Yellow" })
        }
    } catch {
        Write-Log "Falha ao gerar assinatura: $_" -Level ERROR
        Write-Host "  [ERRO] $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Start-UserProvisioning {
    Show-Banner
    Test-RequiredModules

    # Autenticação feita UMA ÚNICA VEZ por execução
    Write-Log "Iniciando autenticação Microsoft Graph..." -Level INFO
    Connect-ToMicrosoftGraph -TenantId $Config.TenantId -ClientId $Config.ClientId

    Write-Log "Verificando acesso ao Active Directory..." -Level INFO
    $activeADServer = Test-ADConnectivity -PrimaryServer  $Config.ADPrimaryServer `
                                          -FallbackServer $Config.ADFallbackServer
    Write-Log "DC ativo nesta sessão: $activeADServer" -Level INFO

    # Catálogos carregados uma vez (grupos) — licenças são recarregadas a cada usuário
    Write-Host "  Carregando grupos do M365 e do AD..." -ForegroundColor Gray
    $ous = @(Get-ADUserOUList -Config $Config)
    $m365Groups = @()
    try { $m365Groups = @(Get-M365SecurityGroups) } catch {
        Write-Log "Falha ao listar grupos M365: $_" -Level ERROR
    }
    $adGroups = @()
    try { $adGroups = @(Get-ADOperationalGroups -ADServer $activeADServer) } catch {
        Write-Log "Falha ao listar grupos do AD: $_" -Level ERROR
    }

    $continuar = $true
    while ($continuar) {
        try {
            $licenses = @(Get-M365AvailableLicenses)

            $plan = Invoke-ProvisioningWizard -Domain $Config.Domain -ADServer $activeADServer `
                                              -OUs $ous -Licenses $licenses `
                                              -M365Groups $m365Groups -ADGroups $adGroups
            if ($plan) {
                $result = Invoke-ProvisioningPlan -Plan $plan -ADServer $activeADServer `
                                                  -DistributionListWaitSeconds $DistributionListWaitSeconds `
                                                  -ShowHeaders
                Show-FinalReport -Plan $plan -Result $result -ActiveDC $activeADServer
                $plan.Password = $null
                if ($result.M365.Success -or $result.AD.Success) { Invoke-SignatureStep -Plan $plan }
            }
        } catch {
            Write-Log "ERRO ao provisionar usuário: $_" -Level ERROR
            Write-Host "`n  [ERRO] $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  A sessão Graph/AD permanece ativa. Você pode tentar novamente." -ForegroundColor Yellow
        }

        $again = Read-Host "`n  Deseja criar outro usuário? (S/N)"
        $continuar = ($again.Trim().ToUpper() -eq "S")
    }

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Write-Log "Sessão encerrada com sucesso." -Level INFO
    Show-Goodbye
}

try {
    Start-UserProvisioning
} catch {
    Write-Log "ERRO CRÍTICO NO SISTEMA: $_" -Level ERROR
    Write-Host "`n  [ERRO CRÍTICO] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Verifique o arquivo de log para detalhes." -ForegroundColor Yellow
    exit 1
}
