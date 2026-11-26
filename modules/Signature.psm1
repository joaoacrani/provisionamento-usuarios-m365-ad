# modules\Signature.psm1
# Geração da assinatura de e-mail (PNG) a partir dos modelos .cdr do CorelDRAW.
#
# Como funciona: abre uma CÓPIA do modelo no CorelDRAW (oculto), troca o texto do
# nome (maior fonte do modelo) e do e-mail (texto com "@"), e exporta PNG 300 dpi
# com fundo transparente, recortado no conteúdo — igual ao processo manual.
#
# Requer CorelDRAW instalado na máquina que roda o script (testado no 2020 / v22).

function Get-SignatureConfig {
    <#
    .SYNOPSIS
        Lê a seção "Signature" do config.json, com valores padrão.
    #>
    param($Config)

    $sig = if ($Config -and $Config.PSObject.Properties.Name -contains 'Signature') { $Config.Signature } else { $null }

    $templatesFolder = if ($sig -and $sig.TemplatesFolder) { $sig.TemplatesFolder } else { '\\fileserver\Compartilhado\Assinaturas\Modelos' }
    $saveRoot        = if ($sig -and $sig.SaveRoot)        { $sig.SaveRoot }        else { '\\fileserver\Compartilhado\Assinaturas\Assinaturas 2026' }
    $progId          = if ($sig -and $sig.CorelProgId)     { $sig.CorelProgId }     else { 'CorelDRAW.Application.22' }

    $templates = if ($sig -and $sig.Templates) { @($sig.Templates) } else {
        @(
            [PSCustomObject]@{ Key = 'MTZ';  Label = 'Matriz';      File = 'Modelo MTZ.cdr';       SaveFolder = 'Matriz';      OUMatch = 'OU=MTZ,' },
            [PSCustomObject]@{ Key = 'FL1';  Label = 'Filial 1'; File = 'Modelo FL1.cdr';       SaveFolder = 'Filial 1'; OUMatch = 'OU=FL1,' },
            [PSCustomObject]@{ Key = 'FL2'; Label = 'Filial 2';       File = 'Modelo FL2.cdr'; SaveFolder = 'Filial 2';       OUMatch = 'OU=FL2,' },
            [PSCustomObject]@{ Key = 'FL3'; Label = 'Filial 3';  File = 'Modelo FL3.cdr';       SaveFolder = 'Filial 3';       OUMatch = 'OU=FL3,' }
        )
    }

    $list = foreach ($t in $templates) {
        @{
            Key        = [string]$t.Key
            Label      = [string]$t.Label
            Path       = Join-Path $templatesFolder ([string]$t.File)
            SaveFolder = Join-Path $saveRoot ([string]$t.SaveFolder)
            OUMatch    = [string]$t.OUMatch
        }
    }

    return @{ ProgId = $progId; SaveRoot = $saveRoot; Templates = @($list) }
}

function Get-SignatureTemplateForOU {
    <#
    .SYNOPSIS
        Sugere o modelo pela OU do usuário (ex.: OU=FL1,... → modelo FL1). Padrão: MTZ.
    #>
    param([object[]]$Templates, [string]$OUPath)
    foreach ($t in $Templates) {
        if ($t.OUMatch -and $OUPath -like "*$($t.OUMatch)*") { return $t }
    }
    return ($Templates | Where-Object { $_.Key -eq 'MTZ' } | Select-Object -First 1) ?? $Templates[0]
}

function Invoke-ComMethod {
    # Chamada COM "crua": o binder do PowerShell não converte bem os parâmetros do Corel
    param($Object, [string]$Method, [object[]]$Arguments = @())
    return [System.__ComObject].InvokeMember($Method, [Reflection.BindingFlags]::InvokeMethod, $null, $Object, $Arguments)
}

function New-EmailSignature {
    <#
    .SYNOPSIS
        Gera o PNG da assinatura. Retorna o caminho do arquivo gerado.
    #>
    param(
        [Parameter(Mandatory)][string]$TemplatePath,
        [Parameter(Mandatory)][string]$FullName,
        [Parameter(Mandatory)][string]$Email,
        [Parameter(Mandatory)][string]$OutFile,
        [string]$ProgId = 'CorelDRAW.Application.22',
        [int]$Dpi = 300
    )

    if (-not (Test-Path -LiteralPath $TemplatePath)) { throw "Modelo de assinatura não encontrado: $TemplatePath" }

    $outDir = Split-Path $OutFile -Parent
    if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force }

    # Trabalha numa cópia local: o modelo original nunca é aberto nem alterado
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("assinatura_" + [guid]::NewGuid().ToString('N') + ".cdr")
    Copy-Item -LiteralPath $TemplatePath -Destination $tmp

    # Se o operador já estiver com o Corel aberto, o COM se conecta a ESSA instância:
    # nesse caso não escondemos a janela nem fechamos o programa dele.
    $before     = @(Get-Process CorelDRW -ErrorAction SilentlyContinue | ForEach-Object Id)
    $wasRunning = $before.Count -gt 0

    Write-Log "Gerando assinatura — $FullName <$Email> | modelo: $(Split-Path $TemplatePath -Leaf)" -Level INFO
    $app = $null; $doc = $null; $mine = @()
    try {
        try { $app = New-Object -ComObject $ProgId }
        catch { throw "Não foi possível iniciar o CorelDRAW ($ProgId). Ele está instalado nesta máquina? Detalhe: $($_.Exception.Message)" }
        if (-not $wasRunning) {
            $app.Visible = $false
            $mine = @(Get-Process CorelDRW -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
        }

        $doc = $app.OpenDocument($tmp)

        $nameShape = $null; $emailShape = $null; $maxSize = 0
        foreach ($sh in $doc.ActivePage.Shapes) {
            if ($sh.Type -ne 6) { continue }                # 6 = texto
            $txt  = [string]$sh.Text.Story.Text
            $size = [double]$sh.Text.Story.Size
            if ($txt -match '@' -and $txt -notmatch '^\s*www') { $emailShape = $sh }
            if ($size -gt $maxSize) { $maxSize = $size; $nameShape = $sh }
        }
        if (-not $nameShape -or -not $emailShape) {
            throw "Não encontrei os textos de nome/e-mail no modelo $(Split-Path $TemplatePath -Leaf)."
        }

        $nameShape.Text.Story.Text  = $FullName.Trim()
        $emailShape.Text.Story.Text = $Email.Trim().ToLower()

        $doc.ActivePage.Shapes.All().CreateSelection()

        # ExportBitmap(arquivo, PNG=802, seleção=2, RGB=4, larg=0, alt=0 (auto), dpiX, dpiY,
        #              antialias=1, dither, transparente, perfil cor, camadas, compressão, paleta, área)
        # Os dois últimos (objetos) precisam ir explicitamente vazios, senão o Corel 2020 recusa.
        $nul  = [System.Runtime.InteropServices.DispatchWrapper]::new($null)
        $callArgs = [object[]]@([string]$OutFile, [int]802, [int]2, [int]4, [int]0, [int]0, [int]$Dpi, [int]$Dpi,
                                [int]1, [bool]$false, [bool]$true, [bool]$false, [bool]$false, [int]0, $nul, $nul)
        $filter = Invoke-ComMethod $doc 'ExportBitmap' $callArgs
        Invoke-ComMethod $filter 'Finish' | Out-Null

    } finally {
        if ($doc) { try { $doc.Close() } catch { } }
        if ($app -and -not $wasRunning) {
            try { $app.Quit() } catch { }
        }
        if ($app) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch { } }
        Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue

        # O Corel às vezes fica residente após Quit — encerra só a instância que NÓS abrimos
        if ($mine.Count) {
            Start-Sleep -Seconds 2
            foreach ($id in $mine) { Get-Process -Id $id -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue }
        }
    }

    if (-not (Test-Path -LiteralPath $OutFile)) { throw "O CorelDRAW não gerou o arquivo da assinatura." }
    Write-Log "Assinatura gerada: $OutFile" -Level SUCCESS
    return $OutFile
}

function Save-SignatureToNetwork {
    <#
    .SYNOPSIS
        Copia o PNG para a pasta da filial na rede, sem sobrescrever arquivo existente.
        Retorna @{ Ok; Path; Message }.
    #>
    param(
        [Parameter(Mandatory)][string]$SourceFile,
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$Login
    )
    $dest = Join-Path $Folder "$($Login.ToLower()).png"
    try {
        if (-not (Test-Path -LiteralPath $Folder)) { return @{ Ok = $false; Path = $dest; Message = "Pasta não encontrada: $Folder" } }
        if (Test-Path -LiteralPath $dest)         { return @{ Ok = $false; Path = $dest; Message = "Já existe uma assinatura com esse nome na rede (não sobrescrevi): $dest" } }
        Copy-Item -LiteralPath $SourceFile -Destination $dest
        Write-Log "Assinatura salva na rede: $dest" -Level SUCCESS
        return @{ Ok = $true; Path = $dest; Message = "Salva em $dest" }
    } catch {
        Write-Log "Falha ao salvar assinatura na rede: $_" -Level WARN
        return @{ Ok = $false; Path = $dest; Message = "Falha ao salvar na rede: $($_.Exception.Message)" }
    }
}

Export-ModuleMember -Function Get-SignatureConfig, Get-SignatureTemplateForOU, New-EmailSignature, Save-SignatureToNetwork
