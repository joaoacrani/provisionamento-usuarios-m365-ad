# modules\M365User.psm1
# Módulo Microsoft 365 — Autenticação Graph, criação de usuário, licenciamento,
# listas de distribuição e grupos.
#
# As funções deste módulo NÃO fazem perguntas ao operador: só listam ou executam.
# As perguntas ficam em UI.psm1 (terminal) e web\index.html (navegador).

$script:FriendlyLicenseNames = @{
    "O365_BUSINESS_ESSENTIALS"    = "Microsoft 365 Business Basic"
    "SMB_BUSINESS_ESSENTIALS"     = "Microsoft 365 Business Basic"
    "O365_BUSINESS_PREMIUM"       = "Microsoft 365 Business Standard"
    "SPB"                         = "Microsoft 365 Business Premium"
    "ENTERPRISEPACK"              = "Office 365 E3"
    "ENTERPRISEPREMIUM"           = "Office 365 E5"
    "EXCHANGESTANDARD"            = "Exchange Online Plan 1"
    "EXCHANGEENTERPRISE"          = "Exchange Online Plan 2"
    "TEAMS_ESSENTIALS"            = "Microsoft Teams Essentials"
    "Microsoft_Teams_Audio_Conferencing_select_dial_out" = "Teams Audio Conferencing"
    "FLOW_FREE"                   = "Power Automate Free"
    "POWER_BI_STANDARD"           = "Power BI Free"
}

# ─────────────────────────────────────────────────────────────
# AUTENTICAÇÃO
# ─────────────────────────────────────────────────────────────

function Connect-ToMicrosoftGraph {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId
    )

    $scopes = @(
        "User.ReadWrite.All",
        "Directory.ReadWrite.All",
        "Organization.Read.All",
        "UserAuthenticationMethod.ReadWrite.All",   # desligamento: remover MFA
        "User.RevokeSessions.All"                   # desligamento: revogar sessões
    )

    # Evita reautenticar se já existe uma sessão válida para o mesmo tenant
    $existingContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($existingContext -and $existingContext.TenantId -eq $TenantId) {
        $missingScopes = @($scopes | Where-Object { $_ -notin $existingContext.Scopes })
        if ($missingScopes.Count -eq 0) {
            Write-Log "Microsoft Graph já conectado — reutilizando sessão existente | Conta: $($existingContext.Account)" -Level SUCCESS
            Write-Host "  Sessão Microsoft Graph já ativa — reutilizando (sem novo login)" -ForegroundColor Green
            return
        }
    }

    try {
        Write-Host "  Conectando ao Microsoft Graph..." -ForegroundColor Cyan
        Write-Host "  (Uma janela de autenticacao pode ser aberta no navegador)`n" -ForegroundColor Gray

        Connect-MgGraph -TenantId $TenantId `
                        -ClientId $ClientId `
                        -Scopes $scopes `
                        -UseDeviceAuthentication `
                        -NoWelcome `
                        -ErrorAction Stop
        Connect-ExchangeOnline -ShowBanner:$false
        $context = Get-MgContext
        Write-Log "Microsoft Graph conectado — Tenant: $($context.TenantId) | Conta: $($context.Account)" -Level SUCCESS

    } catch {
        throw "Falha na autenticacao Microsoft Graph: $_"
    }
}

# ─────────────────────────────────────────────────────────────
# LICENÇAS
# ─────────────────────────────────────────────────────────────

function Get-M365AvailableLicenses {
    <#
    .SYNOPSIS
        Lista as licenças do tenant que ainda têm unidades livres.
        Retorna array de hashtables: Index, SkuId, SkuPart, Friendly, Available, Total.
    #>
    try {
        $skus = @(Get-MgSubscribedSku -ErrorAction Stop |
                  Where-Object { $_.ConsumedUnits -lt $_.PrepaidUnits.Enabled })
    } catch {
        throw "Erro ao listar licencas M365: $_"
    }

    if ($skus.Count -eq 0) {
        throw "Nenhuma licenca com unidades disponiveis encontrada no tenant."
    }

    $i = 1
    $list = foreach ($sku in $skus) {
        $skuPart  = $sku.SkuPartNumber
        $friendly = if ($script:FriendlyLicenseNames[$skuPart]) { $script:FriendlyLicenseNames[$skuPart] } else { $skuPart }
        @{
            Index     = $i++
            SkuId     = [string]$sku.SkuId
            SkuPart   = $skuPart
            Friendly  = $friendly
            Available = [int]($sku.PrepaidUnits.Enabled - $sku.ConsumedUnits)
            Total     = [int]$sku.PrepaidUnits.Enabled
        }
    }
    return $list
}

# ─────────────────────────────────────────────────────────────
# CRIAÇÃO DO USUÁRIO
# ─────────────────────────────────────────────────────────────

function New-M365UserAccount {
    param(
        [Parameter(Mandatory)][hashtable]$UserData,
        [Parameter(Mandatory)][hashtable]$License
    )

    $result = @{
        Success        = $false
        UserId         = $null
        LicenseApplied = $false
        ErrorMessage   = $null
    }

    Write-Host "  Criando usuario no Microsoft Entra ID..." -ForegroundColor Cyan

    $passwordProfile = New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphPasswordProfile
    $passwordProfile.Password                      = $UserData.Password
    $passwordProfile.ForceChangePasswordNextSignIn = $false

    try {
        $newUser = New-MgUser `
            -DisplayName       $UserData.FullName `
            -UserPrincipalName $UserData.Email.ToLower() `
            -MailNickname      $UserData.Login.ToLower() `
            -PasswordProfile   $passwordProfile `
            -AccountEnabled:$true `
            -UsageLocation     "BR" `
            -ErrorAction Stop

        $result.UserId = $newUser.Id
        Write-Log "Usuario M365 criado — UPN: $($UserData.Email) | ID: $($newUser.Id)" -Level SUCCESS
        Write-Host "  Usuario criado no Entra ID" -ForegroundColor Green

    } catch {
        $result.ErrorMessage = "Falha ao criar usuario M365: $_"
        Write-Log $result.ErrorMessage -Level ERROR
        Write-Host "  [ERRO] $($result.ErrorMessage)" -ForegroundColor Red
        return $result
    }

    Write-Host "  Aguardando propagacao do usuario no Entra ID (15s)..." -ForegroundColor Gray
    Start-Sleep -Seconds 15

    Write-Host "  Aplicando licenca '$($License.Friendly)'..." -ForegroundColor Cyan

    $addLicense = New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphAssignedLicense
    $addLicense.SkuId = $License.SkuId

    try {
        Set-MgUserLicense `
            -UserId         $result.UserId `
            -AddLicenses    @($addLicense) `
            -RemoveLicenses @() `
            -ErrorAction Stop | Out-Null

        $result.LicenseApplied = $true
        $result.Success        = $true
        Write-Log "Licenca aplicada — $($License.Friendly) | SkuId: $($License.SkuId)" -Level SUCCESS
        Write-Host "  Licenca aplicada com sucesso" -ForegroundColor Green

    } catch {
        $result.ErrorMessage = "Usuario criado, mas falha ao aplicar licenca: $_"
        $result.Success      = $true
        Write-Log $result.ErrorMessage -Level WARN
        Write-Host "  [AVISO] $($result.ErrorMessage)" -ForegroundColor Yellow
    }

    return $result
}

# ─────────────────────────────────────────────────────────────
# LISTAS DE DISTRIBUIÇÃO
# ─────────────────────────────────────────────────────────────

function ConvertTo-DistributionListAddress {
    <#
    .SYNOPSIS
        Interpreta o texto digitado/colado pelo operador e devolve os endereços das listas.

    .DESCRIPTION
        Aceita qualquer combinação de separadores: vírgula, ponto e vírgula, espaço,
        TAB ou quebra de linha. Exemplos aceitos:
            todos@contoso.com; matriz@contoso.com
            todos, sp, equipe.vendas                      (alias sem domínio)
            Todos <todos@contoso.com>; SP <sp@...> (colado do Outlook)
            todos@contoso.com.                     (ponto final é ignorado)
        Endereços repetidos são removidos. Retorna @{ Valid = string[]; Invalid = string[] }.
    #>
    param(
        [AllowEmptyString()][string]$InputText,
        [Parameter(Mandatory)][string]$Domain
    )

    $valid   = [System.Collections.Generic.List[string]]::new()
    $invalid = [System.Collections.Generic.List[string]]::new()

    if ([string]::IsNullOrWhiteSpace($InputText)) {
        return @{ Valid = @(); Invalid = @() }
    }

    # Formato do Outlook: "Nome <email>" → usa só o que está entre < >
    if ($InputText -match '<[^>]+@[^>]+>') {
        $tokens = [regex]::Matches($InputText, '<([^>]+)>') | ForEach-Object { $_.Groups[1].Value }
    } else {
        $tokens = $InputText -split '[,;\s]+'
    }

    foreach ($raw in $tokens) {
        # Remove pontuação nas pontas (ex.: ponto final, aspas, parênteses)
        $addr = $raw.Trim().Trim('.', ',', ';', ':', '"', "'", '(', ')', '[', ']', '<', '>').ToLower()
        if ([string]::IsNullOrWhiteSpace($addr)) { continue }

        # Alias sem domínio → completa com o domínio da empresa
        if ($addr -notmatch '@') { $addr = "$addr@$Domain" }

        if ($addr -match '^[a-z0-9._%+\-]+@[a-z0-9\-]+(\.[a-z0-9\-]+)*\.[a-z]{2,}$' -and $addr -notmatch '\.\.') {
            if (-not $valid.Contains($addr)) { $valid.Add($addr) }
        } else {
            $invalid.Add($raw.Trim())
        }
    }

    return @{ Valid = @($valid); Invalid = @($invalid) }
}

function Test-DistributionListExists {
    <#
    .SYNOPSIS
        Verifica no Exchange Online se cada endereço é uma lista de distribuição existente.
        Retorna array de @{ Address; Exists; Name }.
    #>
    param([string[]]$Addresses = @())

    $out = foreach ($addr in $Addresses) {
        try {
            $dg = Get-DistributionGroup -Identity $addr -ErrorAction Stop
            @{ Address = $addr; Exists = $true; Name = [string]$dg.DisplayName }
        } catch {
            @{ Address = $addr; Exists = $false; Name = $null }
        }
    }
    return $out
}

function Wait-ExchangeRecipient {
    <#
    .SYNOPSIS
        Aguarda o usuário recém-criado aparecer como destinatário no Exchange Online.
        Sem isso, Add-DistributionGroupMember falha com "Não foi possível localizar o objeto".
    #>
    param(
        [Parameter(Mandatory)][string]$Identity,
        [int]$TimeoutSeconds  = 300,
        [int]$IntervalSeconds = 15
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $first    = $true
    while ((Get-Date) -lt $deadline) {
        $recipient = Get-Recipient -Identity $Identity -ErrorAction SilentlyContinue
        if ($recipient) {
            Write-Log "Caixa/destinatario disponivel no Exchange: $Identity" -Level SUCCESS
            return $true
        }
        if ($first) {
            Write-Log "Aguardando a caixa de '$Identity' ser provisionada no Exchange (ate $TimeoutSeconds s)..." -Level INFO
            $first = $false
        }
        Start-Sleep -Seconds $IntervalSeconds
    }
    return $false
}

function Add-UserToDistributionLists {
    <#
    .SYNOPSIS
        Adiciona o usuário às listas de distribuição (endereços já validados).
        Retorna @{ Added = string[]; Failed = @(@{Group;Error}) }.
    #>
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [string[]]$Addresses = @(),
        [int]$WaitSeconds = 300
    )

    $result = @{ Added = @(); Failed = @() }
    if (-not $Addresses -or $Addresses.Count -eq 0) { return $result }

    Write-Host "  Aguardando a caixa do usuario no Exchange Online..." -ForegroundColor Gray
    $ready = Wait-ExchangeRecipient -Identity $UserPrincipalName -TimeoutSeconds $WaitSeconds

    if (-not $ready) {
        $msg = "Caixa ainda nao provisionada apos $WaitSeconds s. Adicione depois com: Add-DistributionGroupMember -Identity <lista> -Member $UserPrincipalName"
        Write-Log $msg -Level WARN
        foreach ($addr in $Addresses) { $result.Failed += @{ Group = $addr; Error = $msg } }
        return $result
    }

    foreach ($address in $Addresses) {
        try {
            Add-DistributionGroupMember `
                -Identity $address `
                -Member   $UserPrincipalName `
                -BypassSecurityGroupManagerCheck `
                -ErrorAction Stop

            $result.Added += $address
            Write-Log "Adicionado a lista de distribuicao: $address" -Level SUCCESS
            Write-Host "  [OK] Adicionado a: $address" -ForegroundColor Green
        } catch {
            $result.Failed += @{ Group = $address; Error = $_.ToString() }
            Write-Log "Falha ao adicionar a lista de distribuicao '$address': $_" -Level ERROR
            Write-Host "  [ERRO] Falha em '$address': $_" -ForegroundColor Red
        }
    }
    return $result
}

# ─────────────────────────────────────────────────────────────
# GRUPOS M365 / SEGURANÇA (Entra ID)
# ─────────────────────────────────────────────────────────────

function Get-M365SecurityGroups {
    <#
    .SYNOPSIS
        Lista grupos Microsoft 365 (Unified) e de segurança do Entra ID.
        Retorna array de @{ Index; Id; Name; Type }.
    #>
    $groups = @(Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified') or securityEnabled eq true" -ErrorAction Stop |
                Where-Object { $_.DisplayName -and ($_.GroupTypes -contains 'Unified' -or $_.SecurityEnabled -eq $true) } |
                Sort-Object DisplayName)

    $i = 1
    $list = foreach ($grp in $groups) {
        @{
            Index = $i++
            Id    = [string]$grp.Id
            Name  = [string]$grp.DisplayName
            Type  = if ($grp.GroupTypes -contains 'Unified') { "M365" } else { "SEG" }
        }
    }
    return $list
}

function Add-UserToM365SecurityGroups {
    <#
    .SYNOPSIS
        Adiciona o usuário aos grupos M365/segurança informados (@{Id;Name;Type}).
    #>
    param(
        [Parameter(Mandatory)][string]$UserId,
        [object[]]$Groups = @()
    )

    $result = @{ Added = @(); Failed = @() }

    foreach ($grp in $Groups) {
        try {
            New-MgGroupMember -GroupId $grp.Id -DirectoryObjectId $UserId -ErrorAction Stop | Out-Null
            $result.Added += $grp.Name
            Write-Log "Adicionado ao grupo M365 [$($grp.Type)]: $($grp.Name)" -Level SUCCESS
            Write-Host "  [OK] Adicionado a: $($grp.Name)" -ForegroundColor Green
        } catch {
            $result.Failed += @{ Group = $grp.Name; Error = $_.ToString() }
            Write-Log "Falha ao adicionar ao grupo '$($grp.Name)': $_" -Level ERROR
            Write-Host "  [ERRO] Falha em '$($grp.Name)'" -ForegroundColor Red
        }
    }
    return $result
}

Export-ModuleMember -Function `
    Connect-ToMicrosoftGraph, `
    Get-M365AvailableLicenses, `
    New-M365UserAccount, `
    ConvertTo-DistributionListAddress, `
    Test-DistributionListExists, `
    Wait-ExchangeRecipient, `
    Add-UserToDistributionLists, `
    Get-M365SecurityGroups, `
    Add-UserToM365SecurityGroups
