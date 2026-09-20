<#
.SYNOPSIS
    Importa acciones pendientes de Supabase (acciones_colaboradores) hacia Access vía servir.ps1.

.DESCRIPTION
    1. Lee filas con procesado = false en Supabase.
    2. Las envía a POST /api/importar-cierres-colaborador (servidor local).
    3. Si Access responde OK, marca cada fila como procesado = true en Supabase.

    Las mismas credenciales que la web están en Netlify/supabase.js (SUPABASE_URL y SUPABASE_ANON_KEY).
    Copiá esos valores a los parámetros -SupabaseUrl y -AnonKey de este script, o dejá los valores por defecto si ya coinciden.

.PARAMETER SupabaseUrl
    Project URL de Supabase (Settings → API → Project URL).

.PARAMETER AnonKey
    Clave anon public (Settings → API). Para producción podés usar service_role solo en esta PC.

.PARAMETER ImportUrl
    URL del endpoint local de importación. Por defecto http://localhost:8080/api/importar-cierres-colaborador

.PARAMETER Limit
    Máximo de acciones a procesar por ejecución (orden: más antiguas primero).

.EXAMPLE
    # Terminal 1 — servidor local (carpeta App Territorios):
    .\servir.ps1

.EXAMPLE
    # Terminal 2 — sincronizar cola Supabase → Access:
    cd "App Territorios"
    .\Sync-SupabaseToAccess.ps1

.EXAMPLE
    # Con URL y clave explícitas (mismas que Netlify/supabase.js):
    .\Sync-SupabaseToAccess.ps1 `
        -SupabaseUrl 'https://TU_PROYECTO.supabase.co' `
        -AnonKey 'TU_ANON_KEY'

.NOTES
    Requisitos:
    - Tablas creadas (sql/create_tables.sql) y políticas RLS que permitan SELECT/UPDATE en acciones_colaboradores.
    - Access abierto o accesible por la cadena de servir.ps1.
    - Al menos una acción guardada desde predicacion-netlify.html (o insertada de prueba en Supabase).
#>
param(
    [string]$SupabaseUrl = 'https://paupgqeavgxameeqdkuv.supabase.co',
    [string]$AnonKey = 'sb_publishable_KtSh6YrHWSOFewrKNwPSTQ_gg_rq3_L',
    [string]$ImportUrl = 'http://localhost:8080/api/importar-cierres-colaborador',
    [int]$Limit = 100
)

$ErrorActionPreference = 'Stop'

function Invoke-SupabaseRest {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Method = 'GET',
        $Body = $null,
        [hashtable]$ExtraHeaders = @{}
    )

    $headers = @{
        apikey = $AnonKey
        Authorization = "Bearer $AnonKey"
        Accept = 'application/json'
        'Content-Type' = 'application/json; charset=utf-8'
    }
    foreach ($k in $ExtraHeaders.Keys) { $headers[$k] = $ExtraHeaders[$k] }

    $uri = ($SupabaseUrl.TrimEnd('/')) + '/rest/v1' + $Path
    $params = @{ Uri = $uri; Method = $Method; Headers = $headers }
    if ($null -ne $Body) { $params.Body = ($Body | ConvertTo-Json -Depth 30) }
    Invoke-RestMethod @params
}

if ([string]::IsNullOrWhiteSpace($SupabaseUrl) -or $SupabaseUrl -match 'TU_PROYECTO') {
    throw 'Configure -SupabaseUrl (mismo valor que SUPABASE_URL en Netlify/supabase.js).'
}
if ([string]::IsNullOrWhiteSpace($AnonKey) -or $AnonKey -eq 'TU_ANON_KEY') {
    throw 'Configure -AnonKey (mismo valor que SUPABASE_ANON_KEY en Netlify/supabase.js).'
}

Write-Host 'Buscando acciones pendientes en Supabase...'
$path = "/acciones_colaboradores?procesado=eq.false&select=id,tipo,payload,creado_en&order=creado_en.asc&limit=$Limit"
$accionesDb = @(Invoke-SupabaseRest -Path $path)

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
    origen = 'supabase'
    acciones = $acciones
}

Write-Host "Importando $($acciones.Count) accion(es) en Access..."
$importResult = Invoke-RestMethod -Uri $ImportUrl -Method POST -ContentType 'application/json; charset=utf-8' -Body ($payloadImport | ConvertTo-Json -Depth 30)

if ($importResult.ok -eq $false) {
    throw "La importacion local fallo: $($importResult.message)"
}

foreach ($row in $accionesDb) {
    $id = [uri]::EscapeDataString([string]$row.id)
    Invoke-SupabaseRest -Path "/acciones_colaboradores?id=eq.$id" -Method PATCH -Body @{
        procesado = $true
        procesado_en = (Get-Date).ToUniversalTime().ToString('o')
    } -ExtraHeaders @{ Prefer = 'return=minimal' } | Out-Null
}

Write-Host "Listo. Acciones importadas y marcadas como procesadas: $($acciones.Count)"
