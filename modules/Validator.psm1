# modules\Validator.psm1
# Módulo de validações: login, e-mail, senha, conectividade

function Test-RequiredModules {
    <#
    .SYNOPSIS
        Verifica se os módulos obrigatórios estão instalados.
    #>
    $required = @(
    @{ Name = "Microsoft.Graph"; MinVersion = "2.0.0" },
    @{ Name = "ActiveDirectory"; MinVersion = "1.0.0" },
    @{ Name = "ExchangeOnlineManagement"; MinVersion = "3.0.0" }
)

    $allOk = $true

    foreach ($mod in $required) {
        $installed = Get-Module -ListAvailable -Name $mod.Name |
                     Sort-Object Version -Descending |
                     Select-Object -First 1

        if (-not $installed) {
            Write-Host "  [ERRO] Módulo '$($mod.Name)' não encontrado." -ForegroundColor Red
            Write-Host "         Instale com: Install-Module $($mod.Name) -Scope CurrentUser" -ForegroundColor Yellow
            $allOk = $false
        } elseif ($installed.Version -lt [Version]$mod.MinVersion) {
            Write-Host "  [AVISO] Módulo '$($mod.Name)' versão $($installed.Version) pode estar desatualizado." -ForegroundColor Yellow
        } else {
            Write-Log "Módulo $($mod.Name) v$($installed.Version) — OK" -Level SUCCESS
        }
    }

    if (-not $allOk) {
        throw "Módulos obrigatórios ausentes. Instale-os antes de continuar."
    }
}

function Test-ADConnectivity {
    <#
    .SYNOPSIS
        Testa DC primário (DC01/PDC). Se falhar, testa fallback (DC02).
        Retorna o hostname do DC disponível para ser usado em toda a sessão.
    #>
    param(
        [Parameter(Mandatory)][string]$PrimaryServer,
        [string]$FallbackServer = ""
    )

    # Tenta DC primário (DC01 — PDC Emulator)
    try {
        $domain = Get-ADDomain -Server $PrimaryServer -ErrorAction Stop
        Write-Log "Active Directory OK — DC primário: $PrimaryServer | Domínio: $($domain.DNSRoot)" -Level SUCCESS
        return $PrimaryServer
    } catch {
        Write-Log "DC primário '$PrimaryServer' indisponível: $_" -Level WARN
        Write-Host "  [AVISO] DC primário ($PrimaryServer) não respondeu. Tentando fallback..." -ForegroundColor Yellow
    }

    # Tenta DC fallback (DC02)
    if (-not [string]::IsNullOrWhiteSpace($FallbackServer)) {
        try {
            $domain = Get-ADDomain -Server $FallbackServer -ErrorAction Stop
            Write-Log "Usando DC fallback — $FallbackServer | Domínio: $($domain.DNSRoot)" -Level WARN
            Write-Host "  [OK] Conectado ao DC fallback: $FallbackServer" -ForegroundColor Yellow
            return $FallbackServer
        } catch {
            Write-Log "DC fallback '$FallbackServer' também indisponível: $_" -Level ERROR
        }
    }

    throw "Nenhum Domain Controller disponível. Verifique conectividade com DC01 e DC02."
}

function Test-LoginAvailableAD {
    <#
    .SYNOPSIS
        Retorna $true se o login NÃO existe no AD (disponível).
    #>
    param(
        [Parameter(Mandatory)][string]$Login,
        [Parameter(Mandatory)][string]$ADServer
    )

    $existing = Get-ADUser -Filter { SamAccountName -eq $Login } `
                           -Server $ADServer `
                           -ErrorAction SilentlyContinue

    return ($null -eq $existing)
}

function Test-EmailAvailableM365 {
    <#
    .SYNOPSIS
        Retorna $true se o UPN/e-mail NÃO existe no Entra ID (disponível).
    #>
    param(
        [Parameter(Mandatory)][string]$Email
    )

    try {
        $existing = Get-MgUser -Filter "userPrincipalName eq '$Email'" `
                               -ErrorAction Stop
        return ($null -eq $existing -or $existing.Count -eq 0)
    } catch {
        Write-Log "Aviso ao verificar e-mail no M365: $_" -Level WARN
        return $true  # Conservativo: deixa tentar; o Graph retornará erro se duplicado
    }
}

function Test-PasswordComplexity {
    <#
    .SYNOPSIS
        Valida complexidade mínima da senha.
        Requisitos: mín 8 chars, maiúscula, minúscula, número, especial.
    #>
    param([string]$Password)

    if ($Password.Length -lt 8)                         { return $false }
    if ($Password -notmatch '[A-Z]')                    { return $false }
    if ($Password -notmatch '[a-z]')                    { return $false }
    if ($Password -notmatch '[0-9]')                    { return $false }
    if ($Password -notmatch '[^A-Za-z0-9]')             { return $false }
    return $true
}

function Get-ValidatedPassword {
    <#
    .SYNOPSIS
        Solicita senha ao operador com confirmação e validação de complexidade.
        Retorna a senha como string simples (não SecureString) para compatibilidade com Graph API.
    #>
    while ($true) {
        $pwd1 = Read-Host "  Senha inicial (mín 8 chars, maiúscula, número, especial)" -AsSecureString
        $pwd2 = Read-Host "  Confirme a senha" -AsSecureString

        # Converte para comparação
        $plain1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwd1))
        $plain2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwd2))

        if ($plain1 -ne $plain2) {
            Write-Host "  [ERRO] As senhas não coincidem. Tente novamente." -ForegroundColor Red
            continue
        }

        if (-not (Test-PasswordComplexity $plain1)) {
            Write-Host "  [ERRO] Senha não atende aos requisitos de complexidade." -ForegroundColor Red
            Write-Host "         Precisa ter: 8+ chars, 1 maiúscula, 1 minúscula, 1 número, 1 especial." -ForegroundColor Yellow
            continue
        }

        return $plain1
    }
}

Export-ModuleMember -Function `
    Test-RequiredModules, `
    Test-ADConnectivity, `
    Test-LoginAvailableAD, `
    Test-EmailAvailableM365, `
    Test-PasswordComplexity, `
    Get-ValidatedPassword
