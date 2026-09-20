#Requires -Version 5.0
<#
    Administrador central de rutas — App Territorios (PowerShell).
    BasePath = carpeta del script (servir.ps1).
#>

$script:PathManagerBase = $null

function Initialize-PathManager {
    param([string]$BaseDirectory)
    if ([string]::IsNullOrWhiteSpace($BaseDirectory)) {
        $script:PathManagerBase = $PSScriptRoot
    } else {
        $script:PathManagerBase = [System.IO.Path]::GetFullPath($BaseDirectory)
    }
}

function Get-BasePath {
    if ([string]::IsNullOrWhiteSpace($script:PathManagerBase)) {
        Initialize-PathManager -BaseDirectory $PSScriptRoot
    }
    return $script:PathManagerBase
}

function Get-AppPath {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Parts)
    $ruta = Get-BasePath
    foreach ($p in $Parts) {
        if (-not [string]::IsNullOrWhiteSpace($p)) {
            $ruta = [System.IO.Path]::Combine($ruta, $p)
        }
    }
    return $ruta
}

function Ensure-AppDirectories {
    foreach ($sub in @('data', 'config', 'logs', 'exports', 'assets')) {
        $dir = Get-AppPath $sub
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }
}

function Resolve-AccdbPath {
    param([string]$OverridePath = '')

    if (-not [string]::IsNullOrWhiteSpace($OverridePath)) {
        return [System.IO.Path]::GetFullPath($OverridePath)
    }

    $envPath = [Environment]::GetEnvironmentVariable('TERRITORIOS_ACCDB')
    if (-not [string]::IsNullOrWhiteSpace($envPath) -and (Test-Path -LiteralPath $envPath)) {
        return [System.IO.Path]::GetFullPath($envPath)
    }

    $candidatos = @(
        (Get-AppPath 'Territorios_BaseDeDatos.accdb'),
        (Get-AppPath 'data', 'Territorios_BaseDeDatos.accdb')
    )
    foreach ($c in $candidatos) {
        if (Test-Path -LiteralPath $c) {
            return $c
        }
    }
    return $candidatos[0]
}
