#Requires -Version 7.0
<#
.SYNOPSIS
    Interface WEB do Sistema de Provisionamento de Usuários.

.DESCRIPTION
    Sobe um servidor HTTP local (somente http://localhost) usando a mesma sessão
    Microsoft Graph / Exchange / AD e os mesmos módulos do main.ps1, e abre o navegador.

    Segurança:
      • Escuta apenas em localhost — não fica acessível pela rede.
      • Cada execução gera um token aleatório; toda chamada /api exige esse token
        (impede que outro site aberto no navegador dispare ações).
      • Senhas nunca são gravadas em log.

.PARAMETER Port
    Porta local (padrão 8765).

.PARAMETER Demo
    Usa dados fictícios, sem conectar ao M365/AD. Serve para testar a interface.

.EXAMPLE
    pwsh -File web.ps1
    pwsh -File web.ps1 -Demo
#>
param(
    [int]$Port = 8765,
    [switch]$Demo,
    [switch]$NoBrowser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Em caso de erro, mantém a janela aberta (importante quando aberto pelo atalho)
trap {
    Write-Host "`n  [ERRO] $($_.Exception.Message)" -ForegroundColor Red
    try { Write-Log "ERRO CRÍTICO NA INTERFACE WEB: $_" -Level ERROR -NoConsole } catch { }
    if (-not $NoBrowser) { Read-Host "`n  Pressione ENTER para fechar" | Out-Null }
    exit 1
}

. (Join-Path $PSScriptRoot "bootstrap.ps1")

$IndexPath = Join-Path $PSScriptRoot "web\index.html"
if (-not (Test-Path $IndexPath)) { throw "Arquivo da interface não encontrado: $IndexPath" }

# ─────────────────────────────────────────────────────────────
# CONEXÕES E CATÁLOGO
# ─────────────────────────────────────────────────────────────

Show-Banner
if ($Demo) {
    Write-Host "  *** MODO DEMONSTRAÇÃO — nada será criado no M365 ou no AD ***`n" -ForegroundColor Magenta
    Write-Log "Interface web iniciada em MODO DEMO" -Level WARN
    $ActiveADServer = "DC01.corp.contoso.local (demo)"
} else {
    Test-RequiredModules
    Write-Log "Iniciando autenticação Microsoft Graph..." -Level INFO
    Connect-ToMicrosoftGraph -TenantId $Config.TenantId -ClientId $Config.ClientId
    Write-Log "Verificando acesso ao Active Directory..." -Level INFO
    $ActiveADServer = Test-ADConnectivity -PrimaryServer $Config.ADPrimaryServer -FallbackServer $Config.ADFallbackServer
    Write-Log "DC ativo nesta sessão: $ActiveADServer" -Level INFO
}

$script:Catalog = @{ M365Groups = @(); ADGroups = @() }

function Update-GroupCatalog {
    if ($Demo) {
        $i = 1
        $script:Catalog.M365Groups = @(
            "Colaboradores Matriz", "Controladoria", "Equipe Comercial", "Financeiro", "Projetos",
            "TI - Administradores", "Todos Contoso", "VPN - Usuários"
        ) | ForEach-Object { @{ Index = $i; Id = [guid]::NewGuid().ToString(); Name = $_; Type = if ($i++ % 3) { "SEG" } else { "M365" } } }
        $i = 1
        $script:Catalog.ADGroups = @(
            @("COLABORADORES", "Acesso às pastas compartilhadas"), @("CONTROLADORIA", ""), @("FINANCEIRO", "Pasta Financeiro"),
            @("IMPRESSORAS_MTZ", "Impressoras da matriz"), @("FL1_USUARIOS", "Usuários da Filial 1"),
            @("TI", "Equipe de tecnologia"), @("VPN_USUARIOS", "Acesso VPN")
        ) | ForEach-Object { @{ Index = $i++; Name = $_[0]; Description = $_[1]; DN = "CN=$($_[0]),OU=Grupos,DC=corp,DC=contoso,DC=local" } }
        return
    }
    try { $script:Catalog.M365Groups = @(Get-M365SecurityGroups) } catch { Write-Log "Falha ao listar grupos M365: $_" -Level ERROR }
    try { $script:Catalog.ADGroups   = @(Get-ADOperationalGroups -ADServer $ActiveADServer) } catch { Write-Log "Falha ao listar grupos do AD: $_" -Level ERROR }
}

function Get-LicenseCatalog {
    if ($Demo) {
        return @(
            @{ Index = 1; SkuId = "00000000-0000-0000-0000-000000000001"; SkuPart = "O365_BUSINESS_ESSENTIALS"; Friendly = "Microsoft 365 Business Basic";    Available = 12; Total = 80 },
            @{ Index = 2; SkuId = "00000000-0000-0000-0000-000000000002"; SkuPart = "O365_BUSINESS_PREMIUM";    Friendly = "Microsoft 365 Business Standard"; Available = 2;  Total = 40 },
            @{ Index = 3; SkuId = "00000000-0000-0000-0000-000000000003"; SkuPart = "SPB";                      Friendly = "Microsoft 365 Business Premium";  Available = 5;  Total = 10 }
        )
    }
    return @(Get-M365AvailableLicenses)
}

Write-Host "  Carregando grupos do M365 e do AD..." -ForegroundColor Gray
Update-GroupCatalog
$OUs = @(Get-ADUserOUList -Config $Config)
$SigConfig = Get-SignatureConfig -Config $Config

# ─────────────────────────────────────────────────────────────
# AÇÕES DA API
# ─────────────────────────────────────────────────────────────

function Invoke-CheckLogin([string]$Login) {
    if ($Demo) {
        $l = "$Login".Trim().ToLower()
        if ($l -notmatch '^[a-zA-Z0-9_\-\.]{3,20}$') { return @{ Ok = $false; Login = $l; Email = $null; Message = "Login inválido. Use apenas letras, números, _ - . (3 a 20 caracteres)." } }
        if ($l -in @('msouza', 'admin', 'teste')) { return @{ Ok = $false; Login = $l; Email = "$l@$($Config.Domain)"; Message = "Login '$l' já existe no Active Directory." } }
        return @{ Ok = $true; Login = $l; Email = "$l@$($Config.Domain)"; Message = "Disponível no AD e no M365." }
    }
    return Test-LoginAvailability -Login $Login -Domain $Config.Domain -ADServer $ActiveADServer
}

function Invoke-CheckLists([string]$Text) {
    $parsed = ConvertTo-DistributionListAddress -InputText $Text -Domain $Config.Domain
    if ($Demo) {
        $existing = @('todos', 'matriz', 'filial1', 'filial2', 'filial3', 'equipe.vendas', 'equipe.projetos', 'financeiro', 'estagiarios.matriz')
        $checks = @($parsed.Valid | ForEach-Object {
            $alias = ($_ -split '@')[0]
            @{ Address = $_; Exists = ($existing -contains $alias); Name = if ($existing -contains $alias) { "Lista $alias" } else { $null } }
        })
    } else {
        $checks = @(Test-DistributionListExists -Addresses $parsed.Valid)
    }
    return @{ Lists = $checks; Invalid = @($parsed.Invalid) }
}

function ConvertTo-PlanFromRequest([hashtable]$Body) {
    # Remonta e REVALIDA o plano no servidor (não confia só na validação do navegador)
    $errors = [System.Collections.Generic.List[string]]::new()

    $fullName = "$($Body['fullName'])".Trim()
    if (-not $fullName) { $errors.Add("Nome completo é obrigatório.") }

    $login = Invoke-CheckLogin "$($Body['login'])"
    if (-not $login.Ok) { $errors.Add($login.Message) }

    $password = "$($Body['password'])"
    if (-not (Test-PasswordComplexity $password)) { $errors.Add("Senha não atende aos requisitos de complexidade.") }

    $ouPath = "$($Body['ouPath'])".Trim()
    if (-not $ouPath) { $errors.Add("Selecione a OU.") }
    $ouMatch = @($OUs | Where-Object { $_.DN -eq $ouPath })
    $ouLabel = if ($ouMatch.Count) { $ouMatch[0].Label } else { "(manual)" }

    $license = @(Get-LicenseCatalog | Where-Object { $_.SkuId -eq "$($Body['skuId'])" })
    if (-not $license.Count)                 { $errors.Add("Selecione uma licença válida.") }
    elseif ($license[0].Available -le 0)     { $errors.Add("Licença sem unidades disponíveis.") }

    $lists = @((ConvertTo-DistributionListAddress -InputText (@($Body['lists']) -join ';') -Domain $Config.Domain).Valid)

    $m365Ids    = @($Body['m365GroupIds'])
    $m365Groups = @($script:Catalog.M365Groups | Where-Object { $m365Ids -contains $_.Id })
    $adNames    = @($Body['adGroups'])
    $adGroups   = @($script:Catalog.ADGroups | Where-Object { $adNames -contains $_.Name } | ForEach-Object { $_.Name })

    if ($errors.Count) { return @{ Ok = $false; Errors = @($errors) } }

    return @{
        Ok   = $true
        Plan = @{
            FullName = $fullName; Login = $login.Login; Email = $login.Email; Password = $password
            OUPath = $ouPath; OULabel = $ouLabel; License = $license[0]
            DistributionLists = $lists; M365Groups = $m365Groups; ADGroups = $adGroups
        }
    }
}

function Invoke-DemoProvisioning([hashtable]$Plan) {
    $ok = { param($m) Write-Log $m -Level SUCCESS; Start-Sleep -Milliseconds 400 }
    Write-Log "Iniciando provisionamento — Nome: $($Plan.FullName) | Login: $($Plan.Login) | OU: $($Plan.OUPath)" -Level INFO
    & $ok "Usuario M365 criado — UPN: $($Plan.Email) | ID: $([guid]::NewGuid())"
    & $ok "Licenca aplicada — $($Plan.License.Friendly)"
    $dl = @{ Added = @(); Failed = @() }
    if (@($Plan.DistributionLists).Count) {
        Write-Log "Aguardando a caixa de '$($Plan.Email)' ser provisionada no Exchange..." -Level INFO; Start-Sleep -Seconds 1
        foreach ($a in $Plan.DistributionLists) { & $ok "Adicionado a lista de distribuicao: $a"; $dl.Added += $a }
    }
    $mg = @{ Added = @(); Failed = @() }
    foreach ($g in $Plan.M365Groups) { & $ok "Adicionado ao grupo M365 [$($g.Type)]: $($g.Name)"; $mg.Added += $g.Name }
    & $ok "Usuario AD criado — SAM: $($Plan.Login) | OU: $($Plan.OUPath)"
    $ag = @{ Added = @(); Failed = @() }
    foreach ($g in $Plan.ADGroups) { & $ok "Usuario $($Plan.Login) adicionado ao grupo: $g"; $ag.Added += $g }
    return @{
        M365 = @{ Success = $true; UserId = "demo"; LicenseApplied = $true; ErrorMessage = $null }
        DistributionLists = $dl; M365Groups = $mg
        AD = @{ Success = $true; ErrorMessage = $null }; ADGroups = $ag
    }
}

function Invoke-Provision([hashtable]$Body) {
    $built = ConvertTo-PlanFromRequest $Body
    if (-not $built.Ok) { return @{ Ok = $false; Errors = $built.Errors } }
    $plan = $built.Plan

    Start-LogCapture
    try {
        $result = if ($Demo) { Invoke-DemoProvisioning $plan }
                  else { Invoke-ProvisioningPlan -Plan $plan -ADServer $ActiveADServer -DistributionListWaitSeconds $DistributionListWaitSeconds }
        Write-ProvisioningReportLog -Plan $plan -Result $result
        $status = Get-ProvisioningStatus -Result $result
    } catch {
        Write-Log "ERRO ao provisionar usuário: $_" -Level ERROR
        $log = Stop-LogCapture
        return @{ Ok = $false; Errors = @("Erro inesperado: $($_.Exception.Message)"); Log = $log }
    } finally {
        $plan.Password = $null
    }
    $log = Stop-LogCapture

    return @{
        Ok     = $true
        Status = $status
        User   = @{ FullName = $plan.FullName; Login = $plan.Login; Email = $plan.Email; OU = $plan.OUPath; License = $plan.License.Friendly }
        Result = @{
            M365 = @{ Success = $result.M365.Success; LicenseApplied = $result.M365.LicenseApplied; Error = $result.M365.ErrorMessage }
            AD   = @{ Success = $result.AD.Success; Error = $result.AD.ErrorMessage }
            DistributionLists = $result.DistributionLists
            M365Groups        = $result.M365Groups
            ADGroups          = $result.ADGroups
        }
        Log = $log
    }
}

function Invoke-OffboardLookup([hashtable]$Body) {
    $login = "$($Body['login'])".Trim().ToLower()
    if ($Demo) {
        if ($login -notmatch '^[a-z0-9_\-\.]{3,20}$') { return @{ Ok = $false; Errors = @("Login inválido.") } }
        return @{ Ok = $true; Login = $login; Email = "$login@$($Config.Domain)"
                  M365 = @{ Found = $true; DisplayName = "Usuário Demo ($login)"; Enabled = $true }
                  AD   = @{ Found = $true; DisplayName = "Usuário Demo ($login)"; Enabled = $true; DN = "CN=Usuário Demo,OU=MTZ,OU=Contoso,DC=corp,DC=contoso,DC=local" }
                  Password = Get-OffboardingPassword }
    }
    $t = Find-OffboardingUser -Login $login -Domain $Config.Domain -ADServer $ActiveADServer
    if (-not $t.Ok) { return @{ Ok = $false; Errors = @($t.Errors) } }
    return @{
        Ok = $true; Login = $t.Login; Email = $t.Email
        M365 = @{ Found = $t.M365.Found; DisplayName = $t.M365.DisplayName; Enabled = $t.M365.Enabled }
        AD   = $t.AD
        Password = Get-OffboardingPassword
    }
}

function Invoke-Offboard([hashtable]$Body) {
    $login = "$($Body['login'])".Trim().ToLower()
    Start-LogCapture
    try {
        if ($Demo) {
            Write-Log "Iniciando desligamento — Login: $login" -Level INFO
            foreach ($m in "M365: senha alterada", "M365: MFA — 2 método(s) removido(s), 0 falha(s)", "M365: sessões revogadas (logoff solicitado)",
                           "AD: senha alterada ($login)", "AD: conta desabilitada ($login)", "AD: movido para OU=Ex-Usuários ($login)") {
                Write-Log $m -Level SUCCESS; Start-Sleep -Milliseconds 300
            }
            $ok = @{ Ok = $true; Error = $null }
            $result = @{
                M365 = @{ Password = $ok; Mfa = @{ Ok = $true; Removed = @('phoneAuthenticationMethod', 'microsoftAuthenticatorAuthenticationMethod'); Failed = @(); Error = $null }; Sessions = $ok }
                AD   = @{ Password = $ok; Disable = $ok; Move = @{ Ok = $true; Error = $null; Target = "OU=Ex-Usuários,OU=Contoso,DC=corp,DC=contoso,DC=local" } }
                Closing = (Get-OffboardingClosingText)
            }
        } else {
            $target = Find-OffboardingUser -Login $login -Domain $Config.Domain -ADServer $ActiveADServer
            if (-not $target.Ok) { $null = Stop-LogCapture; return @{ Ok = $false; Errors = @($target.Errors) } }
            $result = Invoke-OffboardingPlan -Target $target -ADServer $ActiveADServer -Config $Config
        }
        $status = Get-OffboardingStatus -Result $result
        Write-Log "Desligamento de ${login}: $($status.Text)" -Level INFO
    } catch {
        Write-Log "ERRO no desligamento de ${login}: $_" -Level ERROR
        $log = Stop-LogCapture
        return @{ Ok = $false; Errors = @("Erro inesperado: $($_.Exception.Message)"); Log = $log }
    }
    $log = Stop-LogCapture
    return @{ Ok = $true; Status = $status; Login = $login; Result = $result; Log = $log }
}

function Invoke-Signature([hashtable]$Body) {
    $key      = "$($Body['key'])"
    $fullName = "$($Body['fullName'])".Trim()
    $login    = "$($Body['login'])".Trim().ToLower()

    $tpl = @($SigConfig.Templates | Where-Object { $_.Key -eq $key })
    if (-not $tpl.Count)                                   { return @{ Ok = $false; Errors = @("Modelo de assinatura inválido.") } }
    if (-not $fullName)                                    { return @{ Ok = $false; Errors = @("Informe o nome que vai na assinatura.") } }
    if ($login -notmatch '^[a-z0-9_\-\.]{3,20}$')          { return @{ Ok = $false; Errors = @("Login inválido.") } }
    $tpl   = $tpl[0]
    $email = "$login@$($Config.Domain)"

    $dir = Join-Path ([IO.Path]::GetTempPath()) "UserProvisioning-assinaturas"
    $out = Join-Path $dir "$login.png"
    try {
        New-EmailSignature -TemplatePath $tpl.Path -FullName $fullName -Email $email -OutFile $out -ProgId $SigConfig.ProgId | Out-Null
        $png = [Convert]::ToBase64String([IO.File]::ReadAllBytes($out))

        $network = $null
        if ($Body['saveNetwork'] -eq $true) {
            $network = if ($Demo) { @{ Ok = $false; Message = "Modo demo: não salvei na rede." } }
                       else       { Save-SignatureToNetwork -SourceFile $out -Folder $tpl.SaveFolder -Login $login }
        }
        return @{ Ok = $true; FileName = "$login.png"; Png = $png; Network = $network; Template = $tpl.Label }
    } catch {
        Write-Log "Falha ao gerar assinatura de ${login}: $_" -Level ERROR
        return @{ Ok = $false; Errors = @("Falha ao gerar a assinatura: $($_.Exception.Message)") }
    } finally {
        Remove-Item -LiteralPath $out -ErrorAction SilentlyContinue
    }
}

# ─────────────────────────────────────────────────────────────
# SERVIDOR HTTP
# ─────────────────────────────────────────────────────────────

$Token = -join ((1..32) | ForEach-Object { '{0:x2}' -f (Get-Random -Maximum 256) })

function Send-Response($Ctx, [int]$Status, [string]$ContentType, [byte[]]$Bytes) {
    $res = $Ctx.Response
    $res.StatusCode      = $Status
    $res.ContentType     = $ContentType
    $res.Headers["Cache-Control"]          = "no-store"
    $res.Headers["X-Content-Type-Options"] = "nosniff"
    $res.Headers["X-Frame-Options"]        = "DENY"
    $res.ContentLength64 = $Bytes.Length
    $res.OutputStream.Write($Bytes, 0, $Bytes.Length)
    $res.Close()
}

function Send-Json($Ctx, $Object, [int]$Status = 200) {
    $json = ConvertTo-Json -InputObject $Object -Depth 10 -Compress
    Send-Response $Ctx $Status "application/json; charset=utf-8" ([Text.Encoding]::UTF8.GetBytes($json))
}

function Read-JsonBody($Ctx) {
    $reader = [IO.StreamReader]::new($Ctx.Request.InputStream, [Text.Encoding]::UTF8)
    try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
    if ([string]::IsNullOrWhiteSpace($text)) { return @{} }
    return ($text | ConvertFrom-Json -AsHashtable)
}

$script:StopRequested = $false

function Invoke-Request($Ctx) {
    $req    = $Ctx.Request
    $path   = $req.Url.AbsolutePath.TrimEnd('/')
    $method = $req.HttpMethod

    # Proteção contra DNS rebinding: só aceita Host localhost/127.0.0.1
    if ($req.UserHostName -notmatch "^(localhost|127\.0\.0\.1):$Port$") {
        return Send-Json $Ctx @{ Ok = $false; Errors = @("Host não permitido.") } 403
    }

    if ($method -eq 'GET' -and ($path -eq '' -or $path -eq '/index.html')) {
        return Send-Response $Ctx 200 "text/html; charset=utf-8" ([IO.File]::ReadAllBytes($IndexPath))
    }

    if (-not $path.StartsWith('/api/')) {
        return Send-Json $Ctx @{ Ok = $false; Errors = @("Não encontrado.") } 404
    }

    if ($req.Headers["X-Token"] -ne $Token) {
        return Send-Json $Ctx @{ Ok = $false; Errors = @("Sessão inválida. Feche esta aba e use o link aberto pelo web.ps1.") } 401
    }

    switch ("$method $path") {
        'GET /api/catalog' {
            if ($req.QueryString["refresh"] -eq '1') { Update-GroupCatalog }
            $licenses = @(); $licError = $null
            try { $licenses = @(Get-LicenseCatalog) } catch { $licError = $_.Exception.Message }
            return Send-Json $Ctx @{
                Ok = $true; Demo = [bool]$Demo
                Domain = $Config.Domain; Operator = $env:USERNAME; ADServer = $ActiveADServer
                OUs = $OUs; Licenses = $licenses; LicenseError = $licError
                M365Groups = $script:Catalog.M365Groups; ADGroups = $script:Catalog.ADGroups
                SignatureTemplates = @($SigConfig.Templates | ForEach-Object { @{ Key = $_.Key; Label = $_.Label; SaveFolder = $_.SaveFolder; OUMatch = $_.OUMatch } })
            }
        }
        'POST /api/signature' {
            $b = Read-JsonBody $Ctx
            return Send-Json $Ctx (Invoke-Signature $b)
        }
        'POST /api/check-login' {
            $b = Read-JsonBody $Ctx
            return Send-Json $Ctx (Invoke-CheckLogin "$($b['login'])")
        }
        'POST /api/check-lists' {
            $b = Read-JsonBody $Ctx
            return Send-Json $Ctx (Invoke-CheckLists "$($b['text'])")
        }
        'POST /api/provision' {
            $b = Read-JsonBody $Ctx
            return Send-Json $Ctx (Invoke-Provision $b)
        }
        'POST /api/offboard-lookup' {
            $b = Read-JsonBody $Ctx
            return Send-Json $Ctx (Invoke-OffboardLookup $b)
        }
        'POST /api/offboard' {
            $b = Read-JsonBody $Ctx
            return Send-Json $Ctx (Invoke-Offboard $b)
        }
        'POST /api/shutdown' {
            $script:StopRequested = $true
            return Send-Json $Ctx @{ Ok = $true }
        }
    }
    return Send-Json $Ctx @{ Ok = $false; Errors = @("Rota não encontrada: $method $path") } 404
}

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://localhost:$Port/")
try {
    $listener.Start()
} catch {
    throw "Não foi possível abrir a porta $Port (já está em uso?). Tente: pwsh -File web.ps1 -Port 8766. Detalhe: $($_.Exception.Message)"
}

# O token vai no fragmento (#) da URL: não é enviado ao servidor nem fica em log
$url = "http://localhost:$Port/#t=$Token"
Write-Log "Interface web disponível em http://localhost:$Port/" -Level SUCCESS
Write-Host ""
Write-Host "  Abra no navegador (link com a chave desta sessão):" -ForegroundColor White
Write-Host "  $url" -ForegroundColor Cyan
Write-Host "  Para encerrar: botão 'Encerrar' na página ou Ctrl+C aqui.`n" -ForegroundColor Gray
if (-not $NoBrowser) { Start-Process $url }

try {
    while ($listener.IsListening -and -not $script:StopRequested) {
        $task = $listener.GetContextAsync()
        # Espera em fatias curtas para o Ctrl+C continuar funcionando
        while (-not $task.AsyncWaitHandle.WaitOne(300)) { }
        $ctx = $task.GetAwaiter().GetResult()
        try {
            Invoke-Request $ctx
        } catch {
            Write-Log "Erro na requisição $($ctx.Request.HttpMethod) $($ctx.Request.Url.AbsolutePath): $_" -Level ERROR
            try { Send-Json $ctx @{ Ok = $false; Errors = @("Erro interno: $($_.Exception.Message)") } 500 } catch { }
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
    if (-not $Demo) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    Write-Log "Interface web encerrada." -Level INFO
    Show-Goodbye
}
