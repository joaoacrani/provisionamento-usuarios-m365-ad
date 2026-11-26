# modules\ADUser.psm1
# Módulo Active Directory — Criação de usuário e adição a grupos

# ─────────────────────────────────────────────────────────────
# Grupos do sistema Windows/AD que NUNCA devem aparecer no menu
# ─────────────────────────────────────────────────────────────
$script:GruposSistema = @(
    "Administrators", "Administradores",
    "Users", "Usuários",
    "Guests", "Convidados",
    "Domain Admins", "Admins. do domínio",
    "Domain Users", "Usuários do domínio",
    "Domain Guests", "Convidados domínio",
    "Domain Computers", "Computadores do domínio",
    "Domain Controllers", "Controladores de domínio",
    "Schema Admins", "Administradores de esquema",
    "Enterprise Admins", "Administradores de empresa",
    "Group Policy Creator Owners", "Proprietários criadores de diretiva de grupo",
    "Protected Users",
    "Key Admins", "Administradores de Chaves",
    "Enterprise Key Admins", "Administradores de Chaves Empresariais",
    "Read-only Domain Controllers", "Controladores de Domínio Somente Leitura",
    "Enterprise Read-only Domain Controllers", "Controladores de Domínio de Empresa Somente Leitura",
    "Cloneable Domain Controllers", "Controladores de Domínio Clonáveis",
    "Backup Operators", "Operadores de cópia",
    "Print Operators", "Opers. de impressão",
    "Account Operators", "Opers. de contas",
    "Server Operators", "Opers. de servidores",
    "Network Configuration Operators", "Operadores de configuração de rede",
    "Cryptographic Operators", "Operadores criptográficos",
    "Event Log Readers", "Leitores de log de eventos",
    "Performance Log Users", "Usuários de log de desempenho",
    "Performance Monitor Users", "Usuários de monitor de desempenho",
    "Remote Desktop Users", "Usuários da área de trabalho remota",
    "Remote Management Users", "Usuários de Gerenciamento Remoto",
    "Replicator", "Duplicadores",
    "Incoming Forest Trust Builders", "Criadores de confiança de floresta de entrada",
    "Pre-Windows 2000 Compatible Access", "Acesso compatível anterior ao Windows 2000",
    "Certificate Service DCOM Access",
    "Cert Publishers", "Editores de certificados",
    "RAS and IAS Servers", "Servidores RAS e IAS",
    "Allowed RODC Password Replication Group", "Grupo de Replicação de Senha RODC Permitido",
    "Denied RODC Password Replication Group", "Grupo de Replicação de Senha RODC Negado",
    "Windows Authorization Access Group", "Grupo de acesso de autorização Windows",
    "Terminal Server License Servers", "Servidores de licenças do Terminal Server",
    "Terminal Server Computers",
    "RDS Endpoint Servers",
    "RDS Remote Access Servers",
    "RDS Management Servers",
    "Hyper-V Administrators", "Administradores do Hyper-V",
    "Access Control Assistance Operators", "Operadores de Assistência a Controle de Acesso",
    "Storage Replica Administrators",
    "System Managed Accounts Group",
    "DnsAdmins", "DnsUpdateProxy",
    "DHCP Administrators", "Administradores DHCP",
    "DHCP Users", "Usuários DHCP",
    "IIS_IUSRS",
    "Distributed COM Users", "Distributed COM - Usuários",
    "HelpLibraryUpdaters",
    "WinRMRemoteWMIUsers__",
    "Debugger Users",
    "SupportAssistAdmins", "SupportAssistUsers"
)

function New-ADUserAccount {
    param(
        [Parameter(Mandatory)][hashtable]$UserData,
        [Parameter(Mandatory)][string]$ADServer
    )

    $result = @{
        Success      = $false
        ErrorMessage = $null
    }

    Write-Host "  Criando usuario no Active Directory..." -ForegroundColor Cyan
    Write-Host "  OU destino: $($UserData.OUPath)" -ForegroundColor Gray

    $securePassword = ConvertTo-SecureString $UserData.Password -AsPlainText -Force

    $nameParts = $UserData.FullName.Trim() -split '\s+', 2
    $firstName = $nameParts[0]
    $lastName  = if ($nameParts.Count -gt 1) { $nameParts[1] } else { "" }

    $adUPN = "$($UserData.Login.ToLower())@corp.contoso.local"

    try {
        New-ADUser `
            -Server               $ADServer `
            -Name                 $UserData.FullName `
            -GivenName            $firstName `
            -Surname              $lastName `
            -SamAccountName       $UserData.Login.ToLower() `
            -UserPrincipalName    $adUPN `
            -DisplayName          $UserData.FullName `
            -EmailAddress         $UserData.Email `
            -Path                 $UserData.OUPath `
            -AccountPassword      $securePassword `
            -Enabled              $true `
            -PasswordNeverExpires $true `
            -ChangePasswordAtLogon $false `
            -ErrorAction          Stop

        Write-Log "Usuario AD criado — SAM: $($UserData.Login) | OU: $($UserData.OUPath)" -Level SUCCESS
        Write-Host "  Usuario criado no Active Directory" -ForegroundColor Green
        $result.Success = $true

    } catch {
        $result.ErrorMessage = "Falha ao criar usuario no AD: $_"
        Write-Log $result.ErrorMessage -Level ERROR
        Write-Host "  [ERRO] $($result.ErrorMessage)" -ForegroundColor Red
    } finally {
        $securePassword = $null
    }

    return $result
}

function Get-ADOperationalGroups {
    <#
    .SYNOPSIS
        Lista os grupos do AD que podem ser oferecidos ao operador
        (exclui grupos internos do Windows/AD). Retorna array de @{ Index; Name; Description; DN }.
    #>
    param(
        [Parameter(Mandatory)][string]$ADServer
    )

    $todosGrupos = Get-ADGroup -Filter * -Server $ADServer -Properties Description -ErrorAction Stop |
                   Sort-Object Name

    $operacionais = @($todosGrupos | Where-Object {
        $nome = $_.Name
        if ($script:GruposSistema -contains $nome)          { return $false }
        if ($nome -match '^(SQLServer|SupportAssist)')      { return $false }
        if ($nome -match '^(GPO_Upgrade_W11_Teste)$')       { return $false }
        return $true
    })

    $i = 1
    $list = foreach ($grp in $operacionais) {
        @{
            Index       = $i++
            Name        = [string]$grp.Name
            Description = [string]$grp.Description
            DN          = [string]$grp.DistinguishedName
        }
    }
    return $list
}

function Add-UserToADGroupList {
    <#
    .SYNOPSIS
        Adiciona o usuário aos grupos do AD informados (nomes).
        Retorna @{ Added = string[]; Failed = @(@{Group;Error}) }.
    #>
    param(
        [Parameter(Mandatory)][string]$Login,
        [Parameter(Mandatory)][string]$ADServer,
        [string[]]$GroupNames = @()
    )

    $result = @{ Added = @(); Failed = @() }

    if (-not $GroupNames -or $GroupNames.Count -eq 0) {
        Write-Log "Nenhum grupo do AD selecionado para $Login" -Level INFO
        return $result
    }

    foreach ($name in $GroupNames) {
        try {
            Add-ADGroupMember -Identity $name -Members $Login -Server $ADServer -ErrorAction Stop
            $result.Added += $name
            Write-Log "Usuario $Login adicionado ao grupo: $name" -Level SUCCESS
            Write-Host "  [OK] Adicionado a: $name" -ForegroundColor Green
        } catch {
            $result.Failed += @{ Group = $name; Error = $_.ToString() }
            Write-Log "Falha ao adicionar $Login ao grupo ${name}: $_" -Level ERROR
            Write-Host "  [ERRO] Falha em '$name': $_" -ForegroundColor Red
        }
    }
    return $result
}

Export-ModuleMember -Function New-ADUserAccount, Get-ADOperationalGroups, Add-UserToADGroupList
