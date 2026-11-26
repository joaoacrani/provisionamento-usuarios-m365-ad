# modules\Provisioning.psm1
# Orquestração: recebe um "plano" com todas as escolhas já feitas pelo operador
# (no terminal ou na interface web) e executa a criação de ponta a ponta.
#
# Formato do plano (hashtable):
#   FullName, Login, Email, Password, OUPath, OULabel
#   License           = @{ SkuId; Friendly; ... }
#   DistributionLists = string[]                 (endereços)
#   M365Groups        = @(@{ Id; Name; Type })
#   ADGroups          = string[]                 (nomes)

function Get-ADUserOUList {
    <#
    .SYNOPSIS
        OUs oferecidas ao operador (config.json → ADUserOUs, ou lista padrão).
    #>
    param($Config)

    $ouList = $null
    if ($Config -and $Config.PSObject.Properties.Name -contains 'ADUserOUs') { $ouList = $Config.ADUserOUs }

    if (-not $ouList) {
        $ouList = @(
            [PSCustomObject]@{ Label = "MTZ — Matriz";       DN = "OU=MTZ,OU=Contoso,DC=corp,DC=contoso,DC=local" },
            [PSCustomObject]@{ Label = "FL1 — Filial 1";  DN = "OU=FL1,OU=Contoso,DC=corp,DC=contoso,DC=local" },
            [PSCustomObject]@{ Label = "FL2 — Filial 2";         DN = "OU=FL2,OU=Contoso,DC=corp,DC=contoso,DC=local" },
            [PSCustomObject]@{ Label = "FL3 — Filial 3";         DN = "OU=FL3,OU=Contoso,DC=corp,DC=contoso,DC=local" },
            [PSCustomObject]@{ Label = "TI  — Tecnologia";       DN = "OU=TI,OU=Contoso,DC=corp,DC=contoso,DC=local" },
            [PSCustomObject]@{ Label = "Contoso (raiz)";          DN = "OU=Contoso,DC=corp,DC=contoso,DC=local" },
            [PSCustomObject]@{ Label = "ContosoExterno";          DN = "OU=ContosoExterno,DC=corp,DC=contoso,DC=local" }
        )
    }

    $i = 1
    $list = foreach ($ou in $ouList) { @{ Index = $i++; Label = [string]$ou.Label; DN = [string]$ou.DN } }
    return $list
}

function Test-LoginAvailability {
    <#
    .SYNOPSIS
        Valida formato do login e verifica duplicidade no AD e no M365.
        Retorna @{ Ok; Login; Email; Message }.
    #>
    param(
        [AllowEmptyString()][string]$Login,
        [Parameter(Mandatory)][string]$Domain,
        [Parameter(Mandatory)][string]$ADServer
    )

    $login = "$Login".Trim().ToLower()
    $email = "$login@$Domain"

    if ([string]::IsNullOrWhiteSpace($login)) {
        return @{ Ok = $false; Login = $login; Email = $null; Message = "Login não pode ser vazio." }
    }
    if ($login -notmatch '^[a-zA-Z0-9_\-\.]{3,20}$') {
        return @{ Ok = $false; Login = $login; Email = $null; Message = "Login inválido. Use apenas letras, números, _ - . (3 a 20 caracteres)." }
    }
    if (-not (Test-LoginAvailableAD -Login $login -ADServer $ADServer)) {
        return @{ Ok = $false; Login = $login; Email = $email; Message = "Login '$login' já existe no Active Directory." }
    }
    if (-not (Test-EmailAvailableM365 -Email $email)) {
        return @{ Ok = $false; Login = $login; Email = $email; Message = "E-mail '$email' já existe no Microsoft 365." }
    }
    return @{ Ok = $true; Login = $login; Email = $email; Message = "Disponível no AD e no M365." }
}

function Invoke-ProvisioningPlan {
    <#
    .SYNOPSIS
        Executa o plano: M365 → licença → listas → grupos M365 → AD → grupos AD.
        Retorna hashtable com o resultado de cada parte.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Plan,
        [Parameter(Mandatory)][string]$ADServer,
        [int]$DistributionListWaitSeconds = 300,
        [switch]$ShowHeaders
    )

    $empty = { @{ Added = @(); Failed = @() } }

    Write-Log "Iniciando provisionamento — Nome: $($Plan.FullName) | Login: $($Plan.Login) | OU: $($Plan.OUPath)" -Level INFO

    # ── Microsoft 365 ────────────────────────────────────────
    if ($ShowHeaders) { Show-StepHeader -Step "A" -Title "Criação no Microsoft 365 / Entra ID" }
    $m365Result = @(New-M365UserAccount -UserData $Plan -License $Plan.License)[-1]

    $dlResult  = & $empty
    $m365Group = & $empty
    if ($m365Result.UserId) {
        if (@($Plan.DistributionLists).Count -gt 0) {
            if ($ShowHeaders) { Show-StepHeader -Step "B" -Title "Listas de Distribuição" }
            $dlResult = Add-UserToDistributionLists -UserPrincipalName $Plan.Email `
                                                    -Addresses @($Plan.DistributionLists) `
                                                    -WaitSeconds $DistributionListWaitSeconds
        }
        if (@($Plan.M365Groups).Count -gt 0) {
            if ($ShowHeaders) { Show-StepHeader -Step "C" -Title "Grupos Microsoft 365 / Segurança" }
            $m365Group = Add-UserToM365SecurityGroups -UserId $m365Result.UserId -Groups @($Plan.M365Groups)
        }
    }

    # ── Active Directory ─────────────────────────────────────
    if ($ShowHeaders) { Show-StepHeader -Step "D" -Title "Criação no Active Directory Local" }
    $adResult = New-ADUserAccount -UserData $Plan -ADServer $ADServer

    $adGroups = & $empty
    if ($adResult.Success -and @($Plan.ADGroups).Count -gt 0) {
        if ($ShowHeaders) { Show-StepHeader -Step "E" -Title "Grupos do Active Directory" }
        $adGroups = Add-UserToADGroupList -Login $Plan.Login -ADServer $ADServer -GroupNames @($Plan.ADGroups)
    }

    return @{
        M365              = $m365Result
        DistributionLists = $dlResult
        M365Groups        = $m365Group
        AD                = $adResult
        ADGroups          = $adGroups
    }
}

Export-ModuleMember -Function Get-ADUserOUList, Test-LoginAvailability, Invoke-ProvisioningPlan
