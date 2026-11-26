# modules\UI.psm1
# Interface de terminal — banners, cabeçalhos e o assistente de coleta de dados.
#
# O assistente coleta TUDO antes de criar qualquer coisa, e em qualquer etapa o
# operador pode digitar:
#     V  → voltar para a etapa anterior
#     X  → cancelar este usuário
# Na tela de revisão ainda dá para digitar o número de uma etapa para editá-la.

$script:Back   = '__VOLTAR__'
$script:Cancel = '__CANCELAR__'

$script:WizardSteps = @(
    "Nome completo",
    "Login",
    "Senha",
    "OU do Active Directory",
    "Licença Microsoft 365",
    "Listas de distribuição",
    "Grupos Microsoft 365",
    "Grupos do Active Directory"
)

# ─────────────────────────────────────────────────────────────
# ELEMENTOS VISUAIS
# ─────────────────────────────────────────────────────────────

function Show-Banner {
    try { Clear-Host } catch { }
    $banner = @"

  ╔══════════════════════════════════════════════════════════════╗
  ║        SISTEMA DE PROVISIONAMENTO DE USUÁRIOS               ║
  ║        Contoso  —  contoso.com                      ║
  ║        PowerShell 7  |  Microsoft Graph  |  Active Directory ║
  ╚══════════════════════════════════════════════════════════════╝

"@
    Write-Host $banner -ForegroundColor Cyan
    Write-Host "  Data/Hora : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
    Write-Host "  Operador  : $env:USERNAME`n" -ForegroundColor Gray
}

function Show-StepHeader {
    param(
        [string]$Step,
        [string]$Title
    )

    $line = "─" * 60
    Write-Host "`n  $line" -ForegroundColor DarkCyan
    Write-Host "  ETAPA $Step — $Title" -ForegroundColor White
    Write-Host "  $line`n" -ForegroundColor DarkCyan
    Write-Log "Iniciando ETAPA $Step — $Title" -Level INFO -NoConsole
}

function Show-Goodbye {
    Write-Host "`n  ╔══════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "  ║  Operação concluída. Até logo!       ║" -ForegroundColor Green
    Write-Host "  ╚══════════════════════════════════════╝`n" -ForegroundColor Green
}

function Show-WizardStep {
    param([int]$Index)

    $total = $script:WizardSteps.Count
    $line  = "─" * 60
    Write-Host "`n  $line" -ForegroundColor DarkCyan
    Write-Host "  ETAPA $($Index + 1)/$total — $($script:WizardSteps[$Index])" -ForegroundColor White
    Write-Host "  [V] voltar   [X] cancelar este usuário" -ForegroundColor DarkGray
    Write-Host "  $line`n" -ForegroundColor DarkCyan
}

# ─────────────────────────────────────────────────────────────
# HELPERS DE ENTRADA
# ─────────────────────────────────────────────────────────────

function Test-NavCommand {
    # Converte V/X digitados em comandos de navegação
    param([string]$Value)
    switch -Regex ("$Value".Trim()) {
        '^(?i)(v|voltar)$'    { return $script:Back }
        '^(?i)(x|cancelar)$'  { return $script:Cancel }
    }
    return $null
}

function Read-WizardValue {
    <#
    .SYNOPSIS
        Lê um valor. ENTER mantém o valor atual (se houver). V volta, X cancela.
    #>
    param(
        [string]$Prompt,
        [string]$Current
    )

    $hint = if ($Current) { " [ENTER mantém: $Current]" } else { "" }
    $raw  = Read-Host "  $Prompt$hint"
    $nav  = Test-NavCommand $raw
    if ($nav) { return $nav }

    $value = "$raw".Trim()
    if (-not $value -and $Current) { return $Current }
    return $value
}

function ConvertFrom-IndexSelection {
    <#
    .SYNOPSIS
        Converte "1,3 5;7-9" em números. Aceita vírgula, ponto e vírgula, espaço e intervalos.
        Retorna @{ Indexes = int[]; Invalid = string[] }.
    #>
    param([string]$Text, [int]$Max)

    $idx     = [System.Collections.Generic.List[int]]::new()
    $invalid = [System.Collections.Generic.List[string]]::new()

    foreach ($tok in ("$Text" -split '[,;\s]+' | Where-Object { $_ })) {
        if ($tok -match '^(\d+)-(\d+)$') {
            $a = [int]$Matches[1]; $b = [int]$Matches[2]
            if ($a -gt $b) { $a, $b = $b, $a }
            foreach ($n in $a..$b) {
                if ($n -ge 1 -and $n -le $Max) { if (-not $idx.Contains($n)) { $idx.Add($n) } } else { $invalid.Add("$n") }
            }
        } elseif ($tok -match '^\d+$' -and [int]$tok -ge 1 -and [int]$tok -le $Max) {
            if (-not $idx.Contains([int]$tok)) { $idx.Add([int]$tok) }
        } else {
            $invalid.Add($tok)
        }
    }
    return @{ Indexes = @($idx); Invalid = @($invalid) }
}

# ─────────────────────────────────────────────────────────────
# ETAPAS DO ASSISTENTE
# Cada etapa altera $State e devolve 'next', $script:Back ou $script:Cancel.
# ─────────────────────────────────────────────────────────────

function Step-FullName {
    param([hashtable]$State)
    while ($true) {
        $v = Read-WizardValue -Prompt "Nome completo" -Current $State.FullName
        if ($v -eq $script:Back -or $v -eq $script:Cancel) { return $v }
        if ([string]::IsNullOrWhiteSpace($v)) {
            Write-Host "  [ERRO] Nome não pode ser vazio." -ForegroundColor Red
            continue
        }
        $State.FullName = $v
        return 'next'
    }
}

function Step-Login {
    param([hashtable]$State, [string]$Domain, [string]$ADServer)
    while ($true) {
        $v = Read-WizardValue -Prompt "Login (ex: msouza)" -Current $State.Login
        if ($v -eq $script:Back -or $v -eq $script:Cancel) { return $v }

        Write-Host "  Verificando disponibilidade no AD e no M365..." -ForegroundColor Gray
        $check = Test-LoginAvailability -Login $v -Domain $Domain -ADServer $ADServer
        if (-not $check.Ok) {
            Write-Host "  [ERRO] $($check.Message)" -ForegroundColor Red
            continue
        }
        $State.Login = $check.Login
        $State.Email = $check.Email
        Write-Host "  OK — e-mail: $($check.Email)" -ForegroundColor Green
        return 'next'
    }
}

function Step-Password {
    param([hashtable]$State)

    $hint = if ($State.Password) { " [ENTER mantém a senha já digitada]" } else { "" }
    while ($true) {
        $s1 = Read-Host "  Senha inicial (8+ chars, maiúscula, minúscula, número, especial)$hint" -AsSecureString
        $p1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s1))

        $nav = Test-NavCommand $p1
        if ($nav) { return $nav }
        if (-not $p1 -and $State.Password) { return 'next' }

        if (-not (Test-PasswordComplexity $p1)) {
            Write-Host "  [ERRO] Senha não atende aos requisitos: 8+ chars, 1 maiúscula, 1 minúscula, 1 número, 1 especial." -ForegroundColor Red
            continue
        }

        $s2 = Read-Host "  Confirme a senha" -AsSecureString
        $p2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s2))
        if ($p1 -cne $p2) {
            Write-Host "  [ERRO] As senhas não coincidem. Tente novamente." -ForegroundColor Red
            continue
        }
        $State.Password = $p1
        return 'next'
    }
}

function Step-OU {
    param([hashtable]$State, [object[]]$OUs)

    foreach ($ou in $OUs) {
        $mark = if ($State.OUPath -eq $ou.DN) { "●" } else { " " }
        Write-Host "  $mark [$($ou.Index)] $($ou.Label)" -ForegroundColor White
        Write-Host "         $($ou.DN)" -ForegroundColor DarkGray
    }
    Write-Host "    [0] Digitar caminho manualmente`n" -ForegroundColor Yellow

    $current = $null
    if ($State.OUPath) {
        $match = @($OUs | Where-Object { $_.DN -eq $State.OUPath })
        if ($match.Count -gt 0) { $current = "$($match[0].Index)" }
    }

    while ($true) {
        $v = Read-WizardValue -Prompt "Selecione a OU" -Current "$current"
        if ($v -eq $script:Back -or $v -eq $script:Cancel) { return $v }

        if ($v -eq '0' -or (-not $v -and $State.OUPath -and -not $current)) {
            $manual = Read-WizardValue -Prompt "DN completo da OU" -Current $State.OUPath
            if ($manual -eq $script:Back) { continue }
            if ($manual -eq $script:Cancel) { return $manual }
            if ([string]::IsNullOrWhiteSpace($manual)) { Write-Host "  [ERRO] Caminho não pode ser vazio." -ForegroundColor Red; continue }
            $State.OUPath  = $manual
            $State.OULabel = "(manual)"
            return 'next'
        }

        $sel = $OUs | Where-Object { "$($_.Index)" -eq $v } | Select-Object -First 1
        if (-not $sel) { Write-Host "  [ERRO] Opção inválida." -ForegroundColor Red; continue }

        $State.OUPath  = $sel.DN
        $State.OULabel = $sel.Label
        Write-Host "  OU selecionada: $($sel.Label)" -ForegroundColor Green
        return 'next'
    }
}

function Step-License {
    param([hashtable]$State, [object[]]$Licenses)

    foreach ($lic in $Licenses) {
        $mark  = if ($State.License -and $State.License.SkuId -eq $lic.SkuId) { "●" } else { " " }
        $color = if ($lic.Available -le 2) { "Yellow" } else { "White" }
        Write-Host "  $mark [$($lic.Index)] $($lic.Friendly)" -ForegroundColor $color
        Write-Host "         Total: $($lic.Total)  |  Disponíveis: $($lic.Available)  |  SKU: $($lic.SkuPart)" -ForegroundColor DarkGray
    }
    Write-Host ""

    $current = if ($State.License) { "$($State.License.Index)" } else { $null }
    while ($true) {
        $v = Read-WizardValue -Prompt "Selecione a licença" -Current $current
        if ($v -eq $script:Back -or $v -eq $script:Cancel) { return $v }

        $sel = $Licenses | Where-Object { "$($_.Index)" -eq $v } | Select-Object -First 1
        if (-not $sel)              { Write-Host "  [ERRO] Opção inválida." -ForegroundColor Red; continue }
        if ($sel.Available -le 0)   { Write-Host "  [ERRO] Licença sem unidades disponíveis." -ForegroundColor Red; continue }

        $State.License = $sel
        Write-Log "Licenca selecionada: $($sel.Friendly) | SkuId: $($sel.SkuId)" -Level INFO -NoConsole
        return 'next'
    }
}

function Step-DistributionLists {
    param([hashtable]$State, [string]$Domain)

    Write-Host "  Digite ou cole os endereços das listas. Separe com vírgula, ponto e vírgula," -ForegroundColor White
    Write-Host "  espaço ou uma por linha. Pode digitar só o alias (ex: todos, sp, equipe.vendas)." -ForegroundColor White
    Write-Host "  Também aceita o formato colado do Outlook: Todos <todos@$Domain>;" -ForegroundColor Gray
    Write-Host "  ENTER = manter  |  - = limpar  |  V = voltar`n" -ForegroundColor Gray

    while ($true) {
        if (@($State.DistributionLists).Count -gt 0) {
            Write-Host "  Atuais: $(@($State.DistributionLists) -join '; ')" -ForegroundColor Cyan
        }
        $raw = Read-Host "  Listas"
        $nav = Test-NavCommand $raw
        if ($nav) { return $nav }

        $text = "$raw".Trim()
        if (-not $text) { return 'next' }                       # mantém o que já tinha (ou nenhuma)
        if ($text -eq '-') { $State.DistributionLists = @(); Write-Host "  Listas removidas." -ForegroundColor Gray; return 'next' }

        $parsed = ConvertTo-DistributionListAddress -InputText $text -Domain $Domain
        foreach ($bad in $parsed.Invalid) { Write-Host "  [IGNORADO] Formato inválido: '$bad'" -ForegroundColor Yellow }
        if ($parsed.Valid.Count -eq 0) { Write-Host "  [ERRO] Nenhum endereço válido. Digite novamente." -ForegroundColor Red; continue }

        Write-Host "  Conferindo no Exchange Online..." -ForegroundColor Gray
        $checks   = @(Test-DistributionListExists -Addresses $parsed.Valid)
        $found    = @($checks | Where-Object { $_.Exists }  | ForEach-Object { $_.Address })
        $notFound = @($checks | Where-Object { -not $_.Exists } | ForEach-Object { $_.Address })

        foreach ($c in $checks) {
            if ($c.Exists) { Write-Host "  [OK]            $($c.Address)  ($($c.Name))" -ForegroundColor Green }
            else           { Write-Host "  [NÃO ENCONTRADA] $($c.Address)" -ForegroundColor Red }
        }

        if ($notFound.Count -gt 0) {
            $ans = (Read-Host "  [S] seguir só com as encontradas   [N] digitar de novo").Trim().ToUpper()
            if ($ans -ne 'S') { continue }
        }

        $State.DistributionLists = $found
        return 'next'
    }
}

function Step-GroupSelection {
    <#
    .SYNOPSIS
        Seleção múltipla numerada (grupos M365 ou AD). Marca os já escolhidos com ●.
    #>
    param(
        [hashtable]$State,
        [string]$StateKey,
        [object[]]$Items,
        [scriptblock]$Label,
        [scriptblock]$Key
    )

    if (-not $Items -or $Items.Count -eq 0) {
        Write-Host "  [AVISO] Nenhum grupo disponível." -ForegroundColor Yellow
        $nav = Test-NavCommand (Read-Host "  ENTER para continuar, V para voltar")
        if ($nav) { return $nav }
        return 'next'
    }

    $selectedKeys = @($State[$StateKey] | ForEach-Object { & $Key $_ })
    foreach ($it in $Items) {
        $mark = if ($selectedKeys -contains (& $Key $it)) { "●" } else { " " }
        Write-Host ("  {0} [{1,3}] {2}" -f $mark, $it.Index, (& $Label $it)) -ForegroundColor White
    }
    Write-Host ""
    Write-Host "  Números separados por vírgula, ponto e vírgula ou espaço; aceita intervalo (ex: 1,3 5-7)" -ForegroundColor Cyan
    Write-Host "  ENTER = manter seleção atual  |  - = nenhum  |  V = voltar`n" -ForegroundColor Gray

    while ($true) {
        $raw = Read-Host "  Grupos"
        $nav = Test-NavCommand $raw
        if ($nav) { return $nav }

        $text = "$raw".Trim()
        if (-not $text)    { return 'next' }
        if ($text -eq '-') { $State[$StateKey] = @(); return 'next' }

        $sel = ConvertFrom-IndexSelection -Text $text -Max $Items.Count
        foreach ($bad in $sel.Invalid) { Write-Host "  [IGNORADO] '$bad' não é uma opção válida." -ForegroundColor Yellow }
        if ($sel.Indexes.Count -eq 0) { Write-Host "  [ERRO] Nenhuma opção válida." -ForegroundColor Red; continue }

        $State[$StateKey] = @($Items | Where-Object { $sel.Indexes -contains $_.Index })
        Write-Host "  Selecionados: $((@($State[$StateKey]) | ForEach-Object { & $Label $_ }) -join ', ')" -ForegroundColor Green
        return 'next'
    }
}

function Show-WizardReview {
    param([hashtable]$State)

    $line = "═" * 60
    Write-Host "`n  $line" -ForegroundColor Cyan
    Write-Host "  REVISÃO — nada foi criado ainda" -ForegroundColor White
    Write-Host "  $line" -ForegroundColor Cyan

    $none = "(nenhum)"
    $rows = @(
        @(1, "Nome",          $State.FullName),
        @(2, "Login/E-mail",  "$($State.Login)  →  $($State.Email)"),
        @(3, "Senha",         $(if ($State.Password) { "•" * 8 } else { "" })),
        @(4, "OU",            "$($State.OULabel)  ($($State.OUPath))"),
        @(5, "Licença",       $(if ($State.License) { $State.License.Friendly } else { "" })),
        @(6, "Listas",        $(if (@($State.DistributionLists).Count) { @($State.DistributionLists) -join '; ' } else { $none })),
        @(7, "Grupos M365",   $(if (@($State.M365Groups).Count) { (@($State.M365Groups) | ForEach-Object { $_.Name }) -join ', ' } else { $none })),
        @(8, "Grupos AD",     $(if (@($State.ADGroups).Count) { @($State.ADGroups) -join ', ' } else { $none }))
    )
    foreach ($r in $rows) {
        Write-Host ("  [{0}] {1,-13}: {2}" -f $r[0], $r[1], $r[2]) -ForegroundColor White
    }
    Write-Host "`n  [C] Confirmar e criar   [1-8] Editar etapa   [V] Voltar   [X] Cancelar" -ForegroundColor Yellow
}

# ─────────────────────────────────────────────────────────────
# ASSISTENTE PRINCIPAL
# ─────────────────────────────────────────────────────────────

function Invoke-ProvisioningWizard {
    <#
    .SYNOPSIS
        Conduz o operador pelas 8 etapas com navegação para trás e revisão final.
        Retorna o plano (hashtable) pronto para Invoke-ProvisioningPlan, ou $null se cancelado.
    #>
    param(
        [Parameter(Mandatory)][string]$Domain,
        [Parameter(Mandatory)][string]$ADServer,
        [Parameter(Mandatory)][object[]]$OUs,
        [Parameter(Mandatory)][object[]]$Licenses,
        [object[]]$M365Groups = @(),
        [object[]]$ADGroups   = @()
    )

    $state = @{
        FullName = $null; Login = $null; Email = $null; Password = $null
        OUPath = $null; OULabel = $null; License = $null
        DistributionLists = @(); M365Groups = @(); ADGroupItems = @(); ADGroups = @()
    }

    $total        = $script:WizardSteps.Count
    $i            = 0
    $backToReview = $false

    while ($true) {

        # ── Tela de revisão ──────────────────────────────────
        if ($i -ge $total) {
            $state.ADGroups = @($state.ADGroupItems | ForEach-Object { $_.Name })
            Show-WizardReview -State $state
            $ans = (Read-Host "  Opção").Trim().ToUpper()
            # (if/elseif em vez de switch: 'continue' dentro de switch não volta ao while)
            if ($ans -eq 'C') {
                Write-Log "Dados confirmados — Nome: $($state.FullName) | Login: $($state.Login) | OU: $($state.OUPath)" -Level INFO
                $state.Remove('ADGroupItems')
                return $state
            } elseif ($ans -eq 'X') {
                Write-Host "  Cancelado. Nada foi criado." -ForegroundColor Yellow
                return $null
            } elseif ($ans -eq 'V') {
                $i = $total - 1; $backToReview = $false
            } elseif ($ans -match "^[1-$total]$") {
                $i = [int]$ans - 1; $backToReview = $true
            } else {
                Write-Host "  [ERRO] Opção inválida." -ForegroundColor Red
            }
            continue
        }

        # ── Etapa atual ──────────────────────────────────────
        Show-WizardStep -Index $i
        $r = switch ($i) {
            0 { Step-FullName -State $state }
            1 { Step-Login    -State $state -Domain $Domain -ADServer $ADServer }
            2 { Step-Password -State $state }
            3 { Step-OU       -State $state -OUs $OUs }
            4 { Step-License  -State $state -Licenses $Licenses }
            5 { Step-DistributionLists -State $state -Domain $Domain }
            6 { Step-GroupSelection -State $state -StateKey 'M365Groups' -Items $M365Groups `
                    -Label { param($g) "[$($g.Type)] $($g.Name)" } -Key { param($g) $g.Id } }
            7 { Step-GroupSelection -State $state -StateKey 'ADGroupItems' -Items $ADGroups `
                    -Label { param($g) if ($g.Description) { "$($g.Name) — $($g.Description)" } else { $g.Name } } `
                    -Key { param($g) $g.Name } }
        }

        if ($r -eq $script:Cancel) {
            $c = (Read-Host "  Cancelar este usuário? Nada foi criado ainda. (S/N)").Trim().ToUpper()
            if ($c -eq 'S') { return $null }
            continue
        }
        if ($r -eq $script:Back) {
            if ($backToReview) { $i = $total; continue }        # veio da revisão: volta pra ela
            if ($i -gt 0) { $i-- } else { Write-Host "  Já está na primeira etapa." -ForegroundColor DarkGray }
            continue
        }

        # 'next'
        if ($backToReview) { $i = $total; $backToReview = $false }
        else               { $i++ }
    }
}

Export-ModuleMember -Function Show-Banner, Show-StepHeader, Show-Goodbye, Invoke-ProvisioningWizard, ConvertFrom-IndexSelection
