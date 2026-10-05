<#
.SYNOPSIS
    Importa acciones pendientes de Google Sheets (hoja acciones) hacia Access vía servir.ps1.

.DESCRIPTION
    1. Llama a la aplicación web de Apps Script (listarAccionesPendientes).
    2. Las envía a POST /api/importar-cierres-colaborador (servidor local).
    3. Si Access responde OK, marca cada fila como procesada en la hoja.

    La URL por defecto se lee de google/google.js (APPS_SCRIPT_URL).

.EXAMPLE
    .\servir.ps1
    .\Sync-GoogleToAccess.ps1
#>
param(
    [string]$AppsScriptUrl = '',
    [string]$ImportUrl = 'http://localhost:8080/api/importar-cierres-colaborador',
    [int]$Limit = 100
)

$ErrorActionPreference = 'Stop'

function Get-AppsScriptUrlFromConfig {
    $jsPath = Join-Path $PSScriptRoot 'google\google.js'
    if (-not (Test-Path $jsPath)) { return '' }
    $txt = Get-Content -Raw -Encoding UTF8 $jsPath
    if ($txt -match "APPS_SCRIPT_URL:\s*'([^']+)'") { return $Matches[1] }
    if ($txt -match 'APPS_SCRIPT_URL:\s*"([^"]+)"') { return $Matches[1] }
    return ''
}

if ([string]::IsNullOrWhiteSpace($AppsScriptUrl)) {
    $AppsScriptUrl = Get-AppsScriptUrlFromConfig
}

if ([string]::IsNullOrWhiteSpace($AppsScriptUrl) -or $AppsScriptUrl -match '^TU_') {
    throw 'Configure -AppsScriptUrl o APPS_SCRIPT_URL en google/google.js. Ver GITHUB_PAGES.md.'
}

$AppsScriptUrl = $AppsScriptUrl.TrimEnd('/')

function Invoke-GoogleApi {
    param(
        [Parameter(Mandatory = $true)][string]$Action,
        [hashtable]$Body = $null,
        [string]$Method = 'GET'
    )
    if ($Method -eq 'GET') {
        $uri = $AppsScriptUrl + '?action=' + [uri]::EscapeDataString($Action)
        return Invoke-RestMethod -Uri $uri -Method GET
    }
    $payload = @{ action = $Action }
    if ($Body) { foreach ($k in $Body.Keys) { $payload[$k] = $Body[$k] } }
    $json = $payload | ConvertTo-Json -Depth 30 -Compress
    return Invoke-RestMethod -Uri $AppsScriptUrl -Method POST -ContentType 'text/plain; charset=utf-8' -Body $json
}

function Unwrap-GoogleData($resp) {
    if ($null -eq $resp) { return $null }
    if ($resp.ok -eq $false) { throw [string]$resp.error }
    if ($resp.PSObject.Properties.Name -contains 'data') { return $resp.data }
    return $resp
}

Write-Host 'Buscando acciones pendientes en Google Sheets...'
$raw = Invoke-GoogleApi -Action 'listarAccionesPendientes' -Method GET
$accionesDb = @(Unwrap-GoogleData $raw)
if ($accionesDb.Count -gt $Limit) {
    $accionesDb = $accionesDb[0..($Limit - 1)]
}

$bundleRaw = Invoke-GoogleApi -Action 'cargarDatosSemana' -Method GET
$bundle = Unwrap-GoogleData $bundleRaw
$corte = $null
if ($bundle -and $bundle.exportado) {
    try { $corte = [datetime]::Parse([string]$bundle.exportado).ToUniversalTime() } catch { $corte = $null }
}
if ($corte) {
    $accionesDb = @(
        $accionesDb | Where-Object {
            $creado = $null
            try { $creado = [datetime]::Parse([string]$_.creado_en).ToUniversalTime() } catch { $creado = $null }
            if (-not $creado) { return $true }
            return $creado -gt $corte
        }
    )
}

if ($accionesDb.Count -eq 0) {
    Write-Host 'No hay acciones pendientes.'
    return
}

$acciones = @(
    foreach ($row in $accionesDb) {
        [ordered]@{ tipo = [string]$row.tipo; payload = $row.payload }
    }
)

$payloadImport = [ordered]@{
    exportado = (Get-Date).ToUniversalTime().ToString('o')
    version = 1
    origen = 'google-sheets'
    acciones = $acciones
}

Write-Host "Importando $($acciones.Count) accion(es) en Access..."
$importResult = Invoke-RestMethod -Uri $ImportUrl -Method POST -ContentType 'application/json; charset=utf-8' -Body ($payloadImport | ConvertTo-Json -Depth 30)

if ($importResult.ok -eq $false) {
    throw "La importacion local fallo: $($importResult.message)"
}

foreach ($row in $accionesDb) {
    Invoke-GoogleApi -Action 'marcarAccionProcesada' -Method POST -Body @{ id = $row.id } | Out-Null
}

Write-Host "Listo: $($acciones.Count) accion(es) importadas y marcadas como procesadas en Google Sheets."
