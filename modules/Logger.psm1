# modules\Logger.psm1
# Módulo de logging estruturado com rotação de arquivos

$script:LogFile = $null

function Initialize-Logger {
    param(
        [Parameter(Mandatory)]
        [string]$LogDir
    )

    if (-not (Test-Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }

    $timestamp    = Get-Date -Format "yyyy-MM-dd"
    $script:LogFile = Join-Path $LogDir "provisioning_$timestamp.log"

    Write-Log "═══════════════════════════════════════════════════" -Level INFO
    Write-Log "  Sistema de Provisionamento Iniciado — v1.0.0" -Level INFO
    Write-Log "  Operador : $env:USERNAME | Host: $env:COMPUTERNAME" -Level INFO
    Write-Log "═══════════════════════════════════════════════════" -Level INFO
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO","WARN","ERROR","SUCCESS","DEBUG")]
        [string]$Level = "INFO",

        [switch]$NoConsole
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine   = "[$timestamp] [$Level] $Message"

    # Grava no arquivo
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logLine -Encoding UTF8
    }

    # Captura em memória (usada pela interface web para mostrar o andamento)
    if ($null -ne $script:Capture) {
        $script:Capture.Add([PSCustomObject]@{
            Time    = Get-Date -Format "HH:mm:ss"
            Level   = $Level
            Message = $Message
        })
    }

    # Exibe no console com cor
    if (-not $NoConsole) {
        $color = switch ($Level) {
            "INFO"    { "Cyan" }
            "WARN"    { "Yellow" }
            "ERROR"   { "Red" }
            "SUCCESS" { "Green" }
            "DEBUG"   { "Gray" }
            default   { "White" }
        }
        Write-Host "  $logLine" -ForegroundColor $color
    }
}

function Write-LogSensitive {
    <#
    .SYNOPSIS
        Registra mensagem com dado sensível mascarado (senha nunca aparece no log).
    #>
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    # Substitui qualquer valor após "senha", "password", "pwd" por ****
    $safe = $Message -replace '(?i)(senha|password|pwd)\s*[=:]\s*\S+', '$1=****'
    Write-Log $safe -Level $Level
}

$script:Capture = $null

function Start-LogCapture {
    $script:Capture = [System.Collections.Generic.List[object]]::new()
}

function Stop-LogCapture {
    $lines = if ($null -ne $script:Capture) { @($script:Capture) } else { @() }
    $script:Capture = $null
    return ,$lines
}

Export-ModuleMember -Function Initialize-Logger, Write-Log, Write-LogSensitive, Start-LogCapture, Stop-LogCapture
