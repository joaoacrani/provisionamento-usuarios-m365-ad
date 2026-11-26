# modules\Offboarding.psm1
# Desligamento de usuário (M365 + AD), de forma independente em cada ambiente.
#
#   Microsoft 365 : altera a senha, remove os métodos de MFA (celular, Authenticator, etc.)
#                   e revoga as sessões (logoff).
#   Active Directory: altera a senha, desabilita a conta e move para a OU Ex-Usuários.
#
# Senha do desligamento: uma senha aleatória e forte para CADA desligamento (nunca uma senha padrão
# previsível). Ela fica só em memória: aparece uma vez na tela e não vai para o log.

$script:OffboardingPassword = $null

function New-OffboardingPassword {
    <# Gera (com gerador criptográfico) e guarda a senha do desligamento em andamento. 16 caracteres,
       com maiúscula, minúscula, número e símbolo, sem caracteres que se confundem (0/O, 1/l/I). #>
    $grupos = @('ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnpqrstuvwxyz', '23456789', '!@#$%&*?-_=+')
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $sorteio = {
            param([int]$max)
            $b = New-Object byte[] 4
            do { $rng.GetBytes($b); $n = [BitConverter]::ToUInt32($b, 0) } while ($n -ge ([uint32]::MaxValue - ([uint32]::MaxValue % $max)))
            return [int]($n % $max)
        }
        $chars = New-Object System.Collections.Generic.List[char]
        foreach ($g in $grupos) { $chars.Add($g[(& $sorteio $g.Length)]) }
        $todos = $grupos -join ''
        while ($chars.Count -lt 16) { $chars.Add($todos[(& $sorteio $todos.Length)]) }
        for ($i = $chars.Count - 1; $i -gt 0; $i--) {
            $j = & $sorteio ($i + 1)
            $t = $chars[$i]; $chars[$i] = $chars[$j]; $chars[$j] = $t
        }
        $script:OffboardingPassword = -join $chars
    } finally { $rng.Dispose() }
    return $script:OffboardingPassword
}

function Get-OffboardingPassword {
    <# Senha do desligamento em andamento (gera uma se ainda não houver). #>
    if (-not $script:OffboardingPassword) { [void](New-OffboardingPassword) }
    return $script:OffboardingPassword
}

function Get-OffboardingClosingText {
    <# Texto padrão de encerramento do desligamento. #>
    return "Senhas de acesso alteradas para `"$(Get-OffboardingPassword)`"."
}

function Resolve-ExUsersOU {
    <#
    .SYNOPSIS
        Descobre o DN da OU Ex-Usuários: config.json → ADExUsersOU, ou busca por nome no AD.
        Retorna @{ Ok; DN; Message }.
    #>
    param($Config, [Parameter(Mandatory)][string]$ADServer)

    if ($Config -and $Config.PSObject.Properties.Name -contains 'ADExUsersOU' -and $Config.ADExUsersOU) {
        return @{ Ok = $true; DN = [string]$Config.ADExUsersOU; Message = "" }
    }

    try {
        $found = @(Get-ADOrganizationalUnit -Server $ADServer -ErrorAction Stop `
                    -Filter "Name -eq 'Ex-Usuários'")
    } catch {
        return @{ Ok = $false; DN = $null; Message = "Falha ao procurar a OU Ex-Usuários: $_" }
    }
    if ($found.Count -eq 1) { return @{ Ok = $true; DN = [string]$found[0].DistinguishedName; Message = "" } }
    if ($found.Count -eq 0) { return @{ Ok = $false; DN = $null; Message = "OU Ex-Usuários não encontrada no AD. Defina 'ADExUsersOU' no config.json." } }
    return @{ Ok = $false; DN = $null; Message = "Mais de uma OU Ex-Usuários encontrada. Defina 'ADExUsersOU' no config.json." }
}

function Find-OffboardingUser {
    <#
    .SYNOPSIS
        Localiza o usuário no M365 e no AD (somente leitura) para conferência antes do desligamento.
    #>
    param(
        [Parameter(Mandatory)][string]$Login,
        [Parameter(Mandatory)][string]$Domain,
        [Parameter(Mandatory)][string]$ADServer
    )

    $login = "$Login".Trim().ToLower()
    $upn   = "$login@$Domain"
    $out   = @{
        Ok = $false; Login = $login; Email = $upn; Errors = @()
        M365 = @{ Found = $false; Id = $null; DisplayName = $null; Enabled = $null }
        AD   = @{ Found = $false; DisplayName = $null; Enabled = $null; DN = $null }
    }

    if ($login -notmatch '^[a-zA-Z0-9_\-\.]{3,20}$') {
        $out.Errors = @("Login inválido. Use apenas letras, números, _ - . (3 a 20 caracteres).")
        return $out
    }

    try {
        $m = Get-MgUser -UserId $upn -Property Id, DisplayName, AccountEnabled -ErrorAction Stop
        $out.M365 = @{ Found = $true; Id = [string]$m.Id; DisplayName = [string]$m.DisplayName; Enabled = [bool]$m.AccountEnabled }
    } catch { }

    try {
        $a = Get-ADUser -Identity $login -Server $ADServer -Properties DisplayName, Enabled -ErrorAction Stop
        $out.AD = @{ Found = $true; DisplayName = [string]$a.DisplayName; Enabled = [bool]$a.Enabled; DN = [string]$a.DistinguishedName }
    } catch { }

    if (-not $out.M365.Found -and -not $out.AD.Found) {
        $out.Errors = @("Usuário '$login' não encontrado no Microsoft 365 nem no Active Directory.")
        return $out
    }
    $out.Ok = $true
    return $out
}

# ─────────────────────────────────────────────────────────────
# MICROSOFT 365
# ─────────────────────────────────────────────────────────────

function Set-M365OffboardPassword {
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][string]$Password)

    $pp = @{ Password = $Password; ForceChangePasswordNextSignIn = $false }
    try {
        Update-MgUser -UserId $UserId -PasswordProfile $pp -ErrorAction Stop
    } catch {
        # Sem permissão de User Administrator/Password Administrator o PATCH dá 403.
        # A API de métodos de autenticação usa o papel Authentication Administrator,
        # o mesmo que já remove o MFA.
        Write-Log "M365: PATCH de senha negado, tentando via métodos de autenticação..." -Level WARN
        Reset-MgUserAuthenticationMethodPassword -UserId $UserId `
            -AuthenticationMethodId '28c10230-6103-485e-b985-444c60001490' `
            -NewPassword $Password -ErrorAction Stop | Out-Null
    }
}

function Remove-M365MfaMethods {
    <#
    .SYNOPSIS
        Remove todos os métodos de autenticação removíveis (celular, Authenticator, OATH,
        FIDO2, Windows Hello, e-mail, TAP). A senha em si não é removível e é ignorada.
        Retorna @{ Removed = string[]; Failed = @(@{Method;Error}) }.
    #>
    param([Parameter(Mandatory)][string]$UserId)

    $result  = @{ Removed = @(); Failed = @() }
    $methods = @(Get-MgUserAuthenticationMethod -UserId $UserId -ErrorAction Stop)

    foreach ($m in $methods) {
        $type = "$($m.AdditionalProperties['@odata.type'])" -replace '^#microsoft\.graph\.', ''
        $id   = [string]$m.Id
        if ($type -eq 'passwordAuthenticationMethod') { continue }   # senha não é removível
        # O Entra devolve 409 (conflict) / 429 / 5xx quando há operações simultâneas no tenant:
        # tenta de novo com espera crescente. "Not found" = já removido (ex.: sumiu junto com outro).
        $maxTries = 5
        for ($try = 1; $try -le $maxTries; $try++) {
            try {
                switch ($type) {
                    'phoneAuthenticationMethod'                { Remove-MgUserAuthenticationPhoneMethod                -UserId $UserId -PhoneAuthenticationMethodId $id -ErrorAction Stop }
                    'microsoftAuthenticatorAuthenticationMethod' { Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $UserId -MicrosoftAuthenticatorAuthenticationMethodId $id -ErrorAction Stop }
                    'softwareOathAuthenticationMethod'         { Remove-MgUserAuthenticationSoftwareOathMethod         -UserId $UserId -SoftwareOathAuthenticationMethodId $id -ErrorAction Stop }
                    'fido2AuthenticationMethod'                { Remove-MgUserAuthenticationFido2Method                -UserId $UserId -Fido2AuthenticationMethodId $id -ErrorAction Stop }
                    'windowsHelloForBusinessAuthenticationMethod' { Remove-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $UserId -WindowsHelloForBusinessAuthenticationMethodId $id -ErrorAction Stop }
                    'emailAuthenticationMethod'                { Remove-MgUserAuthenticationEmailMethod                -UserId $UserId -EmailAuthenticationMethodId $id -ErrorAction Stop }
                    'temporaryAccessPassAuthenticationMethod'  { Remove-MgUserAuthenticationTemporaryAccessPassMethod  -UserId $UserId -TemporaryAccessPassAuthenticationMethodId $id -ErrorAction Stop }
                    default { throw "Tipo de método não suportado: $type" }
                }
                $result.Removed += $type
                Write-Log "MFA removido ($type)" -Level SUCCESS
                break
            } catch {
                $msg = $_.ToString()
                if ($msg -match 'NotFound|ResourceNotFound|\b404\b') {
                    $result.Removed += $type
                    Write-Log "MFA ($type) já removido" -Level INFO
                    break
                }
                $transient = $msg -match '\b409\b|Conflict|concurrent|\b429\b|TooManyRequests|\b50[234]\b|ServiceUnavailable|Timeout'
                if ($transient -and $try -lt $maxTries) {
                    $wait = 3 * $try
                    Write-Log "MFA ($type): erro transitório, nova tentativa $($try+1)/$maxTries em ${wait}s" -Level WARN
                    Start-Sleep -Seconds $wait
                    continue
                }
                $result.Failed += @{ Method = $type; Error = $msg }
                Write-Log "Falha ao remover MFA ($type): $msg" -Level ERROR
                break
            }
        }
        Start-Sleep -Milliseconds 800   # evita requisições concorrentes entre remoções
    }
    return $result
}

function Invoke-M365Offboarding {
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$Password
    )

    $r = @{
        Password = @{ Ok = $false; Error = $null }
        Mfa      = @{ Ok = $false; Removed = @(); Failed = @(); Error = $null }
        Sessions = @{ Ok = $false; Error = $null }
    }

    try {
        Set-M365OffboardPassword -UserId $UserId -Password $Password
        $r.Password.Ok = $true
        Write-Log "M365: senha alterada" -Level SUCCESS
    } catch {
        $r.Password.Error = $_.ToString()
        Write-Log "M365: falha ao alterar a senha: $_" -Level ERROR
    }

    try {
        $mfa = Remove-M365MfaMethods -UserId $UserId
        $r.Mfa.Removed = @($mfa.Removed); $r.Mfa.Failed = @($mfa.Failed)
        $r.Mfa.Ok = (@($mfa.Failed).Count -eq 0)
        Write-Log "M365: MFA — $(@($mfa.Removed).Count) método(s) removido(s), $(@($mfa.Failed).Count) falha(s)" -Level $(if ($r.Mfa.Ok) { 'SUCCESS' } else { 'WARN' })
    } catch {
        $r.Mfa.Error = $_.ToString()
        Write-Log "M365: falha ao listar/remover MFA: $_" -Level ERROR
    }

    # Por último, para derrubar também os tokens emitidos antes da troca de senha
    try {
        Revoke-MgUserSignInSession -UserId $UserId -ErrorAction Stop | Out-Null
        $r.Sessions.Ok = $true
        Write-Log "M365: sessões revogadas (logoff solicitado)" -Level SUCCESS
    } catch {
        $r.Sessions.Error = $_.ToString()
        Write-Log "M365: falha ao revogar sessões: $_" -Level ERROR
    }
    return $r
}

# ─────────────────────────────────────────────────────────────
# ACTIVE DIRECTORY
# ─────────────────────────────────────────────────────────────

function Invoke-ADOffboarding {
    param(
        [Parameter(Mandatory)][string]$Login,
        [Parameter(Mandatory)][string]$Password,
        [Parameter(Mandatory)][string]$ADServer,
        [Parameter(Mandatory)][string]$ExUsersOU
    )

    $r = @{
        Password = @{ Ok = $false; Error = $null }
        Disable  = @{ Ok = $false; Error = $null }
        Move     = @{ Ok = $false; Error = $null; Target = $ExUsersOU }
    }

    try { $user = Get-ADUser -Identity $Login -Server $ADServer -ErrorAction Stop }
    catch {
        $msg = "Usuário não encontrado no AD: $_"
        foreach ($k in 'Password', 'Disable', 'Move') { $r[$k].Error = $msg }
        Write-Log "AD: $msg" -Level ERROR
        return $r
    }

    $secure = ConvertTo-SecureString $Password -AsPlainText -Force
    try {
        Set-ADAccountPassword -Identity $user -Server $ADServer -Reset -NewPassword $secure -ErrorAction Stop
        $r.Password.Ok = $true
        Write-Log "AD: senha alterada ($Login)" -Level SUCCESS
    } catch {
        $r.Password.Error = $_.ToString()
        Write-Log "AD: falha ao alterar a senha ($Login): $_" -Level ERROR
    } finally { $secure = $null }

    try {
        Disable-ADAccount -Identity $user -Server $ADServer -ErrorAction Stop
        $r.Disable.Ok = $true
        Write-Log "AD: conta desabilitada ($Login)" -Level SUCCESS
    } catch {
        $r.Disable.Error = $_.ToString()
        Write-Log "AD: falha ao desabilitar ($Login): $_" -Level ERROR
    }

    try {
        Move-ADObject -Identity $user.DistinguishedName -TargetPath $ExUsersOU -Server $ADServer -ErrorAction Stop
        $r.Move.Ok = $true
        Write-Log "AD: movido para $ExUsersOU ($Login)" -Level SUCCESS
    } catch {
        $r.Move.Error = $_.ToString()
        Write-Log "AD: falha ao mover para Ex-Usuários ($Login): $_" -Level ERROR
    }
    return $r
}

function Invoke-OffboardingPlan {
    <#
    .SYNOPSIS
        Executa o desligamento. $Target vem de Find-OffboardingUser.
        Retorna @{ M365 = ...; AD = ...; Closing } (cada parte pode ser $null se o usuário não existe lá).
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Target,
        [Parameter(Mandatory)][string]$ADServer,
        $Config
    )

    $password = New-OffboardingPassword
    Write-Log "Iniciando desligamento — Login: $($Target.Login)" -Level INFO

    $m365 = $null
    if ($Target.M365.Found) { $m365 = Invoke-M365Offboarding -UserId $Target.M365.Id -Password $password }
    else { Write-Log "Usuário não existe no M365 — etapa ignorada" -Level WARN }

    $ad = $null
    if ($Target.AD.Found) {
        $ou = Resolve-ExUsersOU -Config $Config -ADServer $ADServer
        if ($ou.Ok) {
            $ad = Invoke-ADOffboarding -Login $Target.Login -Password $password -ADServer $ADServer -ExUsersOU $ou.DN
        } else {
            Write-Log "AD: $($ou.Message) — movimentação não feita; demais etapas executadas sem mover." -Level ERROR
            # Executa senha e desabilitar mesmo sem OU destino
            $ad = Invoke-ADOffboarding -Login $Target.Login -Password $password -ADServer $ADServer -ExUsersOU "OU-NAO-ENCONTRADA"
            $ad.Move.Error = $ou.Message
        }
    } else { Write-Log "Usuário não existe no AD — etapa ignorada" -Level WARN }

    return @{ M365 = $m365; AD = $ad; Closing = (Get-OffboardingClosingText) }
}

function Get-OffboardingStatus {
    param([Parameter(Mandatory)][hashtable]$Result)

    $oks = [System.Collections.Generic.List[bool]]::new()
    if ($Result.M365) { $oks.Add($Result.M365.Password.Ok); $oks.Add($Result.M365.Mfa.Ok); $oks.Add($Result.M365.Sessions.Ok) }
    if ($Result.AD)   { $oks.Add($Result.AD.Password.Ok); $oks.Add($Result.AD.Disable.Ok); $oks.Add($Result.AD.Move.Ok) }

    if ($oks.Count -gt 0 -and -not $oks.Contains($false)) { return @{ Level = 'ok';      Text = "DESLIGAMENTO CONCLUÍDO COM SUCESSO" } }
    if ($oks.Contains($true))                             { return @{ Level = 'partial'; Text = "DESLIGAMENTO PARCIAL — VERIFICAR ERROS" } }
    return @{ Level = 'fail'; Text = "DESLIGAMENTO FALHOU — VERIFICAR LOGS" }
}

Export-ModuleMember -Function New-OffboardingPassword, Get-OffboardingPassword, Get-OffboardingClosingText, Resolve-ExUsersOU, `
    Find-OffboardingUser, Invoke-OffboardingPlan, Get-OffboardingStatus
