# modules\Report.psm1
# Módulo de relatório final

function Get-ProvisioningStatus {
    <#
    .SYNOPSIS
        Calcula o status geral a partir do resultado de Invoke-ProvisioningPlan.
        Retorna @{ Level = 'ok'|'partial'|'fail'; Text }.
    #>
    param([Parameter(Mandatory)][hashtable]$Result)

    $failures = @($Result.DistributionLists.Failed).Count + @($Result.M365Groups.Failed).Count + @($Result.ADGroups.Failed).Count

    if ($Result.M365.Success -and $Result.AD.Success -and $Result.M365.LicenseApplied -and $failures -eq 0) {
        return @{ Level = 'ok';      Text = "PROVISIONAMENTO CONCLUÍDO COM SUCESSO" }
    } elseif ($Result.M365.Success -or $Result.AD.Success) {
        return @{ Level = 'partial'; Text = "PROVISIONAMENTO PARCIAL — VERIFICAR ERROS" }
    }
    return @{ Level = 'fail'; Text = "PROVISIONAMENTO FALHOU — VERIFICAR LOGS" }
}

function Show-FinalReport {
    param(
        [Parameter(Mandatory)][hashtable]$Plan,
        [Parameter(Mandatory)][hashtable]$Result,
        [string]$ActiveDC = ""
    )

    $m365Status    = if ($Result.M365.Success)        { "✔ OK" } else { "✘ FALHOU" }
    $mailboxStatus = if ($Result.M365.LicenseApplied) { "✔ Provisionando (até 24h)" } else { "✘ Licença não aplicada" }
    $adStatus      = if ($Result.AD.Success)          { "✔ OK" } else { "✘ FALHOU" }
    $status        = Get-ProvisioningStatus -Result $Result
    $statusColor   = @{ ok = "Green"; partial = "Yellow"; fail = "Red" }[$status.Level]

    $line = "═" * 60
    Write-Host "`n  $line" -ForegroundColor Cyan
    Write-Host "  RELATÓRIO DE PROVISIONAMENTO" -ForegroundColor White
    Write-Host "  $line" -ForegroundColor Cyan

    Write-Host ""
    Write-Host "  DADOS DO USUÁRIO" -ForegroundColor Yellow
    Write-Host "  Nome     : $($Plan.FullName)"         -ForegroundColor White
    Write-Host "  Login    : $($Plan.Login)"            -ForegroundColor White
    Write-Host "  E-mail   : $($Plan.Email)"            -ForegroundColor White
    Write-Host "  OU       : $($Plan.OUPath)"           -ForegroundColor White
    Write-Host "  Licença  : $($Plan.License.Friendly)" -ForegroundColor White
    if ($ActiveDC) {
        $dcColor = if ($ActiveDC -like "*DC01*") { "Green" } else { "Yellow" }
        $dcLabel = if ($ActiveDC -like "*DC01*") { "$ActiveDC (primário)" } else { "$ActiveDC (fallback ⚠)" }
        Write-Host "  DC usado : $dcLabel" -ForegroundColor $dcColor
    }

    Write-Host ""
    Write-Host "  STATUS DOS SERVIÇOS" -ForegroundColor Yellow
    Write-Host "  Microsoft 365    : $m365Status"    -ForegroundColor $(if ($Result.M365.Success) { "Green" } else { "Red" })
    Write-Host "  Mailbox/Teams    : $mailboxStatus" -ForegroundColor $(if ($Result.M365.LicenseApplied) { "Cyan" } else { "Red" })
    Write-Host "  Active Directory : $adStatus"      -ForegroundColor $(if ($Result.AD.Success) { "Green" } else { "Red" })

    $sections = @(
        @{ Title = "LISTAS DE DISTRIBUIÇÃO"; Data = $Result.DistributionLists },
        @{ Title = "GRUPOS MICROSOFT 365";   Data = $Result.M365Groups },
        @{ Title = "GRUPOS DO AD";           Data = $Result.ADGroups }
    )
    foreach ($s in $sections) {
        $added  = @($s.Data.Added)
        $failed = @($s.Data.Failed)
        Write-Host ""
        Write-Host "  $($s.Title)" -ForegroundColor Yellow
        Write-Host "  Adicionados : $(if ($added.Count) { $added -join ', ' } else { 'Nenhum' })" -ForegroundColor Green
        if ($failed.Count) {
            Write-Host "  Falhas      : $(($failed | ForEach-Object { $_.Group }) -join ', ')" -ForegroundColor Red
        }
    }

    if (-not $Result.M365.Success -and $Result.M365.ErrorMessage) {
        Write-Host "`n  DETALHES DE ERRO M365:" -ForegroundColor Red
        Write-Host "  $($Result.M365.ErrorMessage)" -ForegroundColor Red
    }
    if (-not $Result.AD.Success -and $Result.AD.ErrorMessage) {
        Write-Host "`n  DETALHES DE ERRO AD:" -ForegroundColor Red
        Write-Host "  $($Result.AD.ErrorMessage)" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "  $line" -ForegroundColor Cyan
    Write-Host "  $($status.Text)" -ForegroundColor $statusColor
    Write-Host "  $line`n" -ForegroundColor Cyan

    # Bloco de credenciais para enviar ao usuário (somente na tela — nunca no log)
    if (($Result.M365.Success -or $Result.AD.Success) -and $Plan.Password) {
        Write-Host "  DADOS PARA ENVIAR AO USUÁRIO (selecione e copie)" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Usuário: $($Plan.Login.ToUpper())"  -ForegroundColor White
        Write-Host "Senha: $($Plan.Password)"           -ForegroundColor White
        Write-Host "E-mail: $($Plan.Email.ToLower())"   -ForegroundColor White
        Write-Host ""
    }

    Write-ProvisioningReportLog -Plan $Plan -Result $Result
}

function Write-ProvisioningReportLog {
    param(
        [Parameter(Mandatory)][hashtable]$Plan,
        [Parameter(Mandatory)][hashtable]$Result
    )
    $status = Get-ProvisioningStatus -Result $Result
    Write-Log "── RELATÓRIO FINAL ──────────────────────────────────" -Level INFO -NoConsole
    Write-Log "Nome     : $($Plan.FullName)"                              -Level INFO -NoConsole
    Write-Log "Login    : $($Plan.Login)"                                 -Level INFO -NoConsole
    Write-Log "Email    : $($Plan.Email)"                                 -Level INFO -NoConsole
    Write-Log "OU       : $($Plan.OUPath)"                                -Level INFO -NoConsole
    Write-Log "Licença  : $($Plan.License.Friendly)"                      -Level INFO -NoConsole
    Write-Log "M365     : $(if ($Result.M365.Success) {'OK'} else {'FALHOU'})" -Level INFO -NoConsole
    Write-Log "AD       : $(if ($Result.AD.Success) {'OK'} else {'FALHOU'})"   -Level INFO -NoConsole
    Write-Log "Listas   : $(@($Result.DistributionLists.Added) -join ', ')" -Level INFO -NoConsole
    Write-Log "Grp M365 : $(@($Result.M365Groups.Added) -join ', ')"      -Level INFO -NoConsole
    Write-Log "Grp AD   : $(@($Result.ADGroups.Added) -join ', ')"        -Level INFO -NoConsole
    Write-Log "Status   : $($status.Text)"                                -Level INFO -NoConsole
    Write-Log "─────────────────────────────────────────────────────"     -Level INFO -NoConsole
}

Export-ModuleMember -Function Show-FinalReport, Get-ProvisioningStatus, Write-ProvisioningReportLog
