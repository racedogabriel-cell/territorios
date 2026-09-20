#Requires -Version 5.0
<#
    Territorios - Servidor local en PowerShell
    -------------------------------------------
    - Sirve el frontend estatico (mapa.html, mapa.svg, mapa_meta.json, etc.)
    - Expone una API REST que consulta Territorios_BaseDeDatos.accdb usando
      el driver Microsoft.ACE.OLEDB.12.0.
    - No requiere instalar nada: PowerShell y .NET Framework ya vienen con
      Windows. El unico requisito es tener Office/Access instalado para
      que exista el driver ACE.

    Uso: doble click en servir.cmd (lanza este script), o:
         powershell -ExecutionPolicy Bypass -File servir.ps1
#>

param(
    [int]$Puerto = 8080,
    [string]$Accdb = ""
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PathManager.ps1')
Initialize-PathManager -BaseDirectory $PSScriptRoot
Ensure-AppDirectories
$raiz = Get-BasePath

# ----------------------------------------------------------------
# Localizar la base Access
# ----------------------------------------------------------------
$Accdb = Resolve-AccdbPath -OverridePath $Accdb
if (-not (Test-Path $Accdb)) {
    Write-Host "ADVERTENCIA: no se encontro el .accdb en:" -ForegroundColor Yellow
    Write-Host "  $Accdb" -ForegroundColor Yellow
    Write-Host "El servidor arrancara igual, pero /api/territorios devolvera error." -ForegroundColor Yellow
}

# Cargar el ensamblado de acceso a datos (viene con .NET Framework)
Add-Type -AssemblyName System.Data | Out-Null

$cadenaConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$Accdb;Persist Security Info=False;"

# ----------------------------------------------------------------
# MIME types
# ----------------------------------------------------------------
$mime = @{
    '.html' = 'text/html; charset=utf-8'
    '.htm'  = 'text/html; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.js'   = 'application/javascript; charset=utf-8'
    '.svg'  = 'image/svg+xml; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.csv'  = 'text/csv; charset=utf-8'
    '.png'  = 'image/png'
    '.jpg'  = 'image/jpeg'
    '.jpeg' = 'image/jpeg'
    '.ico'  = 'image/x-icon'
    '.txt'  = 'text/plain; charset=utf-8'
}

# ----------------------------------------------------------------
# Helpers Access / JSON
# ----------------------------------------------------------------
function ConvertTo-Json5 {
    param($obj)
    # ConvertTo-Json de PS 5.1 trunca a 2 niveles; subimos profundidad.
    return (ConvertTo-Json -InputObject $obj -Depth 10 -Compress)
}

function Repair-Text {
    param([object]$v)
    if ($null -eq $v -or $v -is [System.DBNull]) { return '' }
    $s = [string]$v
    if ([string]::IsNullOrWhiteSpace($s)) { return $s }

    # Intenta corregir mojibake tipo "MaipÃº", "JosÃ©", etc.
    $suspect = ($s.Contains([string][char]195) -or $s.Contains([string][char]194))
    if ($suspect) {
        try {
            $bytes = [System.Text.Encoding]::GetEncoding(1252).GetBytes($s)
            $fixed = [System.Text.Encoding]::UTF8.GetString($bytes)
            if (-not [string]::IsNullOrWhiteSpace($fixed)) { return $fixed }
        } catch {}
    }
    return $s
}

function Normalize-NumeroTerritorio {
    param($raw)
    if ($null -eq $raw) { return '' }
    if ($raw -is [string]) { return $raw.Trim() }
    try {
        return ([decimal][double]$raw).ToString([System.Globalization.CultureInfo]::InvariantCulture).Trim()
    } catch {
        return ([string]$raw).Trim()
    }
}

function Param-OleDbNumeroTerritorio {
    param([string]$s)
    $t = if ($null -eq $s) { '' } else { $s.Trim() }
    if ($t -eq '') { return [System.DBNull]::Value }
    $tn = $t
    if ($tn.Contains(',') -and -not $tn.Contains('.')) { $tn = $tn.Replace(',', '.') }
    $d = 0.0
    if ([double]::TryParse($tn, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$d)) { return $d }
    return $t
}

function Ensure-RelacionNumeroTerritorioDouble {
    if (-not (Test-Path $Accdb)) { return }
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $schema = $con.GetSchema('Columns')
        $col = $schema | Where-Object {
            $_.TABLE_NAME -eq 'RelacionTerritoriopuntosdeencuentro' -and
            $_.COLUMN_NAME -eq 'NumeroTerritorio'
        } | Select-Object -First 1
        if ($null -eq $col) { return }

        # DATA_TYPE 3 = Integer. Los territorios como 78.2 se truncaban a 78.
        if ([int]$col.DATA_TYPE -eq 3) {
            $cmd = $con.CreateCommand()
            $cmd.CommandText = 'ALTER TABLE [RelacionTerritoriopuntosdeencuentro] ALTER COLUMN [NumeroTerritorio] DOUBLE'
            [void]$cmd.ExecuteNonQuery()
            Write-Host "  Esquema actualizado: RelacionTerritoriopuntosdeencuentro.NumeroTerritorio ahora acepta decimales." -ForegroundColor Yellow
        }
    }
    finally { $con.Close() }
}

Ensure-RelacionNumeroTerritorioDouble

function Read-Territorios {
    # Jet/ACE no maneja bien subqueries correlacionadas en SELECT, asi que
    # pre-agregamos el Max(Terminado) en un LEFT JOIN.
    # La columna "Campa<enie>a" se construye en tiempo de ejecucion para que
    # el script sea independiente del encoding del archivo .ps1.
    $enie = [char]0x00F1   # 'ñ'
    $sql = @"
SELECT nt.NumeroTerritorio,
       nt.[Campa${enie}a]         AS Campana,
       nt.AbiertoCerrado,
       nt.Rurales,
       nt.Asignado,
       nt.Enviado,
       nt.PuntosDeEncuentros,
       nt.Responsables,
       nt.Grupos,
       nt.ManzanasPendientes,
       nt.FechaManzanasPendientes,
       nt.UltimoEncargado,
       nt.Descripcion,
       m.MaxTerminado             AS UltimaTerminado
FROM NumeroDeTerritorios AS nt
LEFT JOIN (
    SELECT NumeroTerritorio, Max(Terminado) AS MaxTerminado
    FROM Historial
    GROUP BY NumeroTerritorio
) AS m ON nt.NumeroTerritorio = m.NumeroTerritorio
ORDER BY nt.NumeroTerritorio
"@
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $cmd = $con.CreateCommand(); $cmd.CommandText = $sql
        $rd = $cmd.ExecuteReader()
        $lista = New-Object System.Collections.ArrayList
        while ($rd.Read()) {
            $fecha = $null; $dias = $null
            $u = $rd['UltimaTerminado']
            if ($u -isnot [System.DBNull] -and $null -ne $u) {
                $f = [datetime]$u
                $fecha = $f.ToString('yyyy-MM-dd')
                $dias = [int][math]::Floor((([datetime]::Today) - $f.Date).TotalDays)
            }

            # Asignado en Access es DateTime: lo tratamos como booleano
            # (true si tiene fecha, false si es Null) que es lo que el VBA
            # evalua como "datos(4) > 0".
            $asg  = $rd['Asignado']
            $asgB = -not ($asg -is [System.DBNull] -or $null -eq $asg)
            $asgF = $null
            if ($asgB) { $asgF = ([datetime]$asg).ToString('yyyy-MM-dd') }

            $mp = $rd['ManzanasPendientes']
            $mpS = if ($mp -is [System.DBNull] -or $null -eq $mp) { $null } else { [string]$mp }
            $fmp = $rd['FechaManzanasPendientes']
            $fmpS = if ($fmp -is [System.DBNull] -or $null -eq $fmp) { $null } else { ([datetime]$fmp).ToString('yyyy-MM-dd') }
            $ue = $rd['UltimoEncargado']
            $ueS = if ($ue -is [System.DBNull] -or $null -eq $ue) { $null } else { [string]$ue }
            $env = $rd['Enviado']
            $envS = if ($env -is [System.DBNull] -or $null -eq $env) { $null } else { ([datetime]$env).ToString('yyyy-MM-ddTHH:mm:ss') }
            $pto = $rd['PuntosDeEncuentros']
            $ptoS = if ($pto -is [System.DBNull] -or $null -eq $pto) { $null } else { [string]$pto }
            $rsp = $rd['Responsables']
            $rspS = if ($rsp -is [System.DBNull] -or $null -eq $rsp) { $null } else { [string]$rsp }
            $grp = $rd['Grupos']
            $grpS = if ($grp -is [System.DBNull] -or $null -eq $grp) { $null } else { [string]$grp }
            $desc = $rd['Descripcion']
            $descS = if ($desc -is [System.DBNull] -or $null -eq $desc) { $null } else { ([string]$desc).Trim() }
            if ([string]::IsNullOrWhiteSpace($descS)) { $descS = $null }

            [void]$lista.Add([ordered]@{
                numero             = [string]$rd['NumeroTerritorio']
                ultimaVisita       = $fecha
                dias               = $dias
                campana            = [bool]($rd['Campana']        -as [bool])
                abiertoCerrado     = [bool]($rd['AbiertoCerrado'] -as [bool])
                rurales            = [bool]($rd['Rurales']        -as [bool])
                asignado           = $asgB
                fechaAsignado      = $asgF
                enviado            = $envS
                puntoEncuentro     = $ptoS
                responsables       = $rspS
                grupo              = $grpS
                manzanasPendientes = $mpS
                fechaManzanasPendientes = $fmpS
                ultimoEncargado    = $ueS
                descripcion        = $descS
            })
        }
        $rd.Close()
        return ,$lista.ToArray()
    }
    finally { $con.Close() }
}

function Read-SalidasPredicacion {
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        # Evitamos JOIN con NoVisitar para no duplicar salidas cuando hay varias direcciones.
        $dirPorTerr = @{}
        try {
            $colsNv = Resolve-NoVisitarColumns $con
            if (-not [string]::IsNullOrWhiteSpace($colsNv.colDir)) {
                $cmdNv = $con.CreateCommand()
                $cmdNv.CommandText = "SELECT [" + $colsNv.colNum + "] AS Num, [" + $colsNv.colDir + "] AS Dir FROM NoVisitar"
                $rdNv = $cmdNv.ExecuteReader()
                while ($rdNv.Read()) {
                    $num = if ($rdNv['Num'] -is [System.DBNull]) { '' } else { ([string]$rdNv['Num']).Trim() }
                    $dir = if ($rdNv['Dir'] -is [System.DBNull]) { '' } else { ([string]$rdNv['Dir']).Trim() }
                    if ($num -eq '' -or $dir -eq '') { continue }
                    if (-not $dirPorTerr.ContainsKey($num)) { $dirPorTerr[$num] = New-Object System.Collections.ArrayList }
                    if (-not ($dirPorTerr[$num] -contains $dir)) { [void]$dirPorTerr[$num].Add($dir) }
                }
                $rdNv.Close()
            }
        } catch {}

        $sql = "SELECT N.NumeroTerritorio, N.PuntosDeEncuentros, N.Responsables, N.Grupos, N.Enviado, N.Prioridad, N.ManzanasPendientes, N.Descripcion " +
               "FROM NumeroDeTerritorios AS N " +
               "WHERE N.Enviado > 0 " +
               "ORDER BY N.Enviado, N.Grupos, N.Prioridad"

        $cmd = $con.CreateCommand()
        $cmd.CommandText = $sql
        $rd = $cmd.ExecuteReader()
        $lista = New-Object System.Collections.ArrayList
        while ($rd.Read()) {
            $env = if ($rd['Enviado'] -is [System.DBNull]) { $null } else { ([datetime]$rd['Enviado']).ToString('yyyy-MM-ddTHH:mm:ss') }
            $numTerr = if ($rd['NumeroTerritorio'] -is [System.DBNull]) { '' } else { ([string]$rd['NumeroTerritorio']).Trim() }
            $dir = ''
            if ($numTerr -ne '' -and $dirPorTerr.ContainsKey($numTerr) -and $dirPorTerr[$numTerr].Count -gt 0) {
                $dir = [string]::Join(" / ", $dirPorTerr[$numTerr].ToArray())
            }
            $descRaw = $rd['Descripcion']
            $desc = if ($descRaw -is [System.DBNull] -or $null -eq $descRaw) { '' } else { ([string]$descRaw).Trim() }
            [void]$lista.Add([ordered]@{
                NumeroTerritorio = $numTerr
                PuntosDeEncuentros = if ($rd['PuntosDeEncuentros'] -is [System.DBNull]) { '' } else { [string]$rd['PuntosDeEncuentros'] }
                Responsables = if ($rd['Responsables'] -is [System.DBNull]) { '' } else { [string]$rd['Responsables'] }
                Grupos = if ($rd['Grupos'] -is [System.DBNull]) { '' } else { [string]$rd['Grupos'] }
                Enviado = $env
                Prioridad = if ($rd['Prioridad'] -is [System.DBNull]) { $null } else { [int]$rd['Prioridad'] }
                ManzanasPendientes = if ($rd['ManzanasPendientes'] -is [System.DBNull]) { '' } else { [string]$rd['ManzanasPendientes'] }
                Direccion = $dir
                Descripcion = $desc
            })
        }
        $rd.Close()
        return ,$lista.ToArray()
    }
    finally { $con.Close() }
}

function Read-SalidaBundleField {
    <# Lee un campo de una fila 'salida' deserializada desde JSON (casing variable). #>
    param([object]$row, [string[]]$nombres)

    if ($null -eq $row) { return $null }
    if ($row -is [System.Collections.IDictionary]) {
        foreach ($nm in $nombres) {
            foreach ($key in $row.Keys) {
                if ([string]::Equals([string]$key, $nm, [System.StringComparison]::OrdinalIgnoreCase)) {
                    return $row[$key]
                }
            }
        }
        return $null
    }
    foreach ($nm in $nombres) {
        try {
            $prop = $row.PSObject.Properties[$nm]
            if ($null -ne $prop -and $null -ne $prop.Value) { return $prop.Value }
        } catch {}
    }
    foreach ($p in $row.PSObject.Properties) {
        foreach ($nm in $nombres) {
            if ([string]::Equals($p.Name, $nm, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $p.Value
            }
        }
    }
    return $null
}

function Update-SalidaBundleRowAccess {
    param(
        $con,
        [string]$numNorm,
        [string]$pto,
        [string]$gr,
        [string]$resp,
        [datetime]$envDt,
        $priVal,
        [string]$manz
    )

    $sqlSet = "UPDATE NumeroDeTerritorios SET AbiertoCerrado=True, PuntosDeEncuentros=?, Grupos=?, Responsables=?, Enviado=?, Prioridad=?, ManzanasPendientes=? WHERE "

    function Invoke-One([string]$whereClause, $numParam) {
        $cmd = $con.CreateCommand()
        $cmd.CommandText = $sqlSet + $whereClause
        [void]$cmd.Parameters.AddWithValue('@p0', $pto)
        [void]$cmd.Parameters.AddWithValue('@p1', $gr)
        [void]$cmd.Parameters.AddWithValue('@p2', $resp)
        $pe = $cmd.Parameters.Add('@p3', [System.Data.OleDb.OleDbType]::Date)
        $pe.Value = [datetime]$envDt
        $pp = $cmd.Parameters.Add('@p4', [System.Data.OleDb.OleDbType]::Integer)
        if ($priVal -is [System.DBNull]) {
            $pp.Value = [DBNull]::Value
        }
        else {
            $pp.Value = [int]$priVal
        }
        [void]$cmd.Parameters.AddWithValue('@p5', $manz)
        [void]$cmd.Parameters.AddWithValue('@p6', $numParam)
        return [int]$cmd.ExecuteNonQuery()
    }

    # Misma estrategia que Set-Asignacion: comparación numérica directa (Jet/ACE).
    $n = Invoke-One 'NumeroTerritorio=?' (Param-OleDbNumeroTerritorio $numNorm)
    if ($n -gt 0) { return $n }

    # Fallbacks por formato texto del número (coma/punto) con CStr — parámetro TEXTO.
    $variants = New-Object System.Collections.Generic.List[string]
    [void]$variants.Add($numNorm)
    if ($numNorm.Contains('.') -and -not $numNorm.Contains(',')) {
        [void]$variants.Add(($numNorm.Replace('.', ',')))
    }
    if ($numNorm.Contains(',') -and -not $numNorm.Contains('.')) {
        [void]$variants.Add(($numNorm.Replace(',', '.')))
    }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($v in $variants) {
        $vv = ([string]$v).Trim()
        if ([string]::IsNullOrWhiteSpace($vv) -or -not $seen.Add($vv)) { continue }
        $n = Invoke-One 'CStr(NumeroTerritorio)=?' $vv
        if ($n -gt 0) { return $n }
    }
    return 0
}

<#
 Aplica filas de salidas del bundle publicado en Google Sheets (datos_semana) sobre NumeroDeTerritorios.
 Solo actualiza territorios que ya existen en la base. No elimina salidas locales omitidas del bundle.
#>
function Sync-SalidasBundleToAccess {
    param($bundle)

    if ($null -eq $bundle) {
        throw 'Bundle vacio.'
    }

    $salidasRaw = Read-SalidaBundleField $bundle @('salidas', 'Salidas')
    if ($null -eq $salidasRaw) {
        $salidas = @()
    }
    else {
        $salidas = @($salidasRaw)
    }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $aplicados = 0
        $fallos = New-Object System.Collections.ArrayList

        foreach ($s in $salidas) {
            $rawNum = Read-SalidaBundleField $s @(
                'NumeroTerritorio', 'numero_territorio', 'numeroTerritorio', 'numero', 'Numero'
            )
            $num = Normalize-NumeroTerritorio $rawNum

            if ([string]::IsNullOrWhiteSpace($num)) { continue }

            try {
                $pto = [string](Read-SalidaBundleField $s @('PuntosDeEncuentros', 'puntos_de_encuentros', 'puntosDeEncuentros', 'punto_encuentro', 'punto'))
                $resp = [string](Read-SalidaBundleField $s @('Responsables', 'responsables', 'responsable'))
                $gr = [string](Read-SalidaBundleField $s @('Grupos', 'grupos', 'grupo'))
                $manz = [string](Read-SalidaBundleField $s @('ManzanasPendientes', 'manzanas_pendientes', 'manzanasPendientes'))

                $envRaw = Read-SalidaBundleField $s @('Enviado', 'enviado', 'fecha_hora', 'fecha')
                if ($null -eq $envRaw -or [string]::IsNullOrWhiteSpace([string]$envRaw)) {
                    [void]$fallos.Add("Territorio $num : Enviado vacio (se omite)")
                    continue
                }

                $envDt = $null
                try {
                    $envDt = [datetime]::Parse([string]$envRaw, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal)
                }
                catch {
                    try { $envDt = [datetime]::Parse([string]$envRaw) }
                    catch {
                        [void]$fallos.Add("Territorio $num : Enviado invalido")
                        continue
                    }
                }

                $priVal = [System.DBNull]::Value
                $pri = Read-SalidaBundleField $s @('Prioridad', 'prioridad')
                if ($null -ne $pri -and "$pri" -ne '') {
                    $ix = 0
                    if ([int]::TryParse("$pri", [ref]$ix)) {
                        $priVal = $ix
                    }
                }

                $n = Update-SalidaBundleRowAccess -con $con -numNorm $num -pto $pto -gr $gr -resp $resp -envDt $envDt -priVal $priVal -manz $manz
                if ($n -lt 1) {
                    [void]$fallos.Add("Territorio $num : no encontrado en la base (revisá NumeroTerritorio en Access vs Google Sheets)")
                }
                else {
                    $aplicados++
                }
            }
            catch {
                [void]$fallos.Add("Territorio $num : $($_.Exception.Message)")
            }
        }

        return [ordered]@{
            ok      = ($fallos.Count -eq 0)
            aplicados = $aplicados
            fallos  = @($fallos.ToArray())
        }
    }
    finally { $con.Close() }
}

function Read-NoVisitarMap {
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $schema = $con.GetSchema("Columns")
        $colNum = "NumeroTerritorio"
        $colDir = $null
        $colFecha = $null
        foreach ($r in $schema.Rows) {
            if ([string]$r["TABLE_NAME"] -ne "NoVisitar") { continue }
            $c = [string]$r["COLUMN_NAME"]
            if ($c -match '(?i)dire') { $colDir = $c }
            if ($c -match '(?i)numero') { $colNum = $c }
            if ($c -match '(?i)fech|date') { $colFecha = $c }
        }
        if ([string]::IsNullOrWhiteSpace($colDir)) { return @() }

        $cmd = $con.CreateCommand()
        $selFecha = if (-not [string]::IsNullOrWhiteSpace($colFecha)) { ", [" + $colFecha + "] AS Fec" } else { "" }
        $cmd.CommandText = "SELECT [" + $colNum + "] AS Num, [" + $colDir + "] AS Dir" + $selFecha + " FROM NoVisitar"
        $rd = $cmd.ExecuteReader()
        $lista = New-Object System.Collections.ArrayList
        while ($rd.Read()) {
            $num = if ($rd['Num'] -is [System.DBNull]) { '' } else { ([string]$rd['Num']).Trim() }
            $dir = if ($rd['Dir'] -is [System.DBNull]) { '' } else { ([string]$rd['Dir']).Trim() }
            $fec = $null
            try {
                if ($rd['Fec'] -isnot [System.DBNull] -and $null -ne $rd['Fec']) { $fec = ([datetime]$rd['Fec']).ToString('yyyy-MM-ddTHH:mm:ss') }
            } catch {}
            if ($num -ne '' -and $dir -ne '') { [void]$lista.Add([ordered]@{ NumeroTerritorio = $num; Direccion = $dir; FechaRegistro = $fec }) }
        }
        $rd.Close()
        if ($lista.Count -eq 0) { return @() }
        return @($lista.ToArray())
    }
    finally { $con.Close() }
}

function Resolve-NoVisitarColumns {
    param($con)
    $schema = $con.GetSchema("Columns")
    $colNum = "NumeroTerritorio"
    $colDir = $null
    $colFecha = $null
    foreach ($r in $schema.Rows) {
        if ([string]$r["TABLE_NAME"] -ne "NoVisitar") { continue }
        $c = [string]$r["COLUMN_NAME"]
        if ($c -match '(?i)dire') { $colDir = $c }
        if ($c -match '(?i)numero') { $colNum = $c }
        if ($c -match '(?i)fech|date') { $colFecha = $c }
    }
    return [ordered]@{ colNum = $colNum; colDir = $colDir; colFecha = $colFecha }
}

function Add-NoVisitar {
    param([string]$numero, [string]$direccion)
    $num = ([string]$numero).Trim()
    $dir = ([string]$direccion).Trim()
    if ([string]::IsNullOrWhiteSpace($num)) { throw "Numero de territorio requerido." }
    if ([string]::IsNullOrWhiteSpace($dir)) { throw "Direccion requerida." }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $cols = Resolve-NoVisitarColumns $con
        if ([string]::IsNullOrWhiteSpace($cols.colDir)) { throw "La tabla NoVisitar no tiene columna de direccion." }

        $chk = $con.CreateCommand()
        $chk.CommandText = "SELECT COUNT(*) FROM NoVisitar WHERE CStr([" + $cols.colNum + "])=? AND [" + $cols.colDir + "]=?"
        [void]$chk.Parameters.AddWithValue('@n', $num)
        [void]$chk.Parameters.AddWithValue('@d', $dir)
        $count = [int]$chk.ExecuteScalar()
        if ($count -eq 0) {
            $ins = $con.CreateCommand()
            if (-not [string]::IsNullOrWhiteSpace($cols.colFecha)) {
                $ins.CommandText = "INSERT INTO NoVisitar ([" + $cols.colNum + "], [" + $cols.colDir + "], [" + $cols.colFecha + "]) VALUES (?, ?, ?)"
                [void]$ins.Parameters.AddWithValue('@n', $num)
                [void]$ins.Parameters.AddWithValue('@d', $dir)
                $pF = $ins.Parameters.Add('@f', [System.Data.OleDb.OleDbType]::Date)
                $pF.Value = [datetime]::Now
            } else {
                $ins.CommandText = "INSERT INTO NoVisitar ([" + $cols.colNum + "], [" + $cols.colDir + "]) VALUES (?, ?)"
                [void]$ins.Parameters.AddWithValue('@n', $num)
                [void]$ins.Parameters.AddWithValue('@d', $dir)
            }
            [void]$ins.ExecuteNonQuery()
        }
        return [ordered]@{ ok = $true; numero = $num; direccion = $dir }
    }
    finally { $con.Close() }
}

function Delete-NoVisitar {
    param([string]$numero, [string]$direccion)
    $num = ([string]$numero).Trim()
    $dir = ([string]$direccion).Trim()
    if ([string]::IsNullOrWhiteSpace($num)) { throw "Numero de territorio requerido." }
    if ([string]::IsNullOrWhiteSpace($dir)) { throw "Direccion requerida." }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $cols = Resolve-NoVisitarColumns $con
        if ([string]::IsNullOrWhiteSpace($cols.colDir)) { throw "La tabla NoVisitar no tiene columna de direccion." }

        $del = $con.CreateCommand()
        $del.CommandText = "DELETE FROM NoVisitar WHERE CStr([" + $cols.colNum + "])=? AND [" + $cols.colDir + "]=?"
        [void]$del.Parameters.AddWithValue('@n', $num)
        [void]$del.Parameters.AddWithValue('@d', $dir)
        $rows = [int]$del.ExecuteNonQuery()
        return [ordered]@{ ok = $true; numero = $num; eliminados = $rows }
    }
    finally { $con.Close() }
}

function Read-NoVisitarByNumero {
    param([string]$numero)
    $num = ([string]$numero).Trim()
    if ([string]::IsNullOrWhiteSpace($num)) { return @() }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $cols = Resolve-NoVisitarColumns $con
        if ([string]::IsNullOrWhiteSpace($cols.colDir)) { return @() }

        $selectFecha = if (-not [string]::IsNullOrWhiteSpace($cols.colFecha)) { ", [" + $cols.colFecha + "] AS FechaRegistro" } else { "" }
        $cmd = $con.CreateCommand()
        $cmd.CommandText = "SELECT [" + $cols.colDir + "] AS Direccion" + $selectFecha + " FROM NoVisitar WHERE CStr([" + $cols.colNum + "])=? ORDER BY [" + $cols.colDir + "]"
        [void]$cmd.Parameters.AddWithValue('@n', $num)
        $rd = $cmd.ExecuteReader()
        $lista = New-Object System.Collections.ArrayList
        while ($rd.Read()) {
            $dir = if ($rd['Direccion'] -is [System.DBNull]) { '' } else { ([string]$rd['Direccion']).Trim() }
            if ([string]::IsNullOrWhiteSpace($dir)) { continue }
            $fechaTxt = $null
            try {
                $f = $rd['FechaRegistro']
                if ($f -isnot [System.DBNull] -and $null -ne $f) { $fechaTxt = ([datetime]$f).ToString('yyyy-MM-ddTHH:mm:ss') }
            } catch {}
            [void]$lista.Add([ordered]@{
                NumeroTerritorio = $num
                Direccion = $dir
                FechaRegistro = $fechaTxt
            })
        }
        $rd.Close()
        return @($lista.ToArray())
    }
    finally { $con.Close() }
}

function Read-Historial {
    param([string]$numero)
    $sql = "SELECT NumeroTerritorio, Iniciado, Terminado, Encargado FROM Historial WHERE NumeroTerritorio = ? ORDER BY Iniciado DESC"
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $cmd = $con.CreateCommand(); $cmd.CommandText = $sql
        [void]$cmd.Parameters.AddWithValue('@num', $numero)
        $rd = $cmd.ExecuteReader()
        $lista = New-Object System.Collections.ArrayList
        while ($rd.Read()) {
            [void]$lista.Add([ordered]@{
                numero    = [string]$rd['NumeroTerritorio']
                iniciado  = if ($rd['Iniciado']  -is [System.DBNull]) { $null } else { ([datetime]$rd['Iniciado']).ToString('yyyy-MM-dd') }
                terminado = if ($rd['Terminado'] -is [System.DBNull]) { $null } else { ([datetime]$rd['Terminado']).ToString('yyyy-MM-dd') }
                encargado = if ($rd['Encargado'] -is [System.DBNull]) { $null } else { [string]$rd['Encargado'] }
            })
        }
        $rd.Close()
        return ,$lista.ToArray()
    }
    finally { $con.Close() }
}

function Read-FormData {
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $calles = New-Object System.Collections.ArrayList
        $responsables = New-Object System.Collections.ArrayList
        $horarios = New-Object System.Collections.ArrayList
        $gruposSet = New-Object 'System.Collections.Generic.HashSet[string]'

        $cmd = $con.CreateCommand()
        $cmd.CommandText = "SELECT [Id], [NombreCalles], [Horientacion] FROM [Calles] ORDER BY [NombreCalles]"
        $rd = $cmd.ExecuteReader()
        while ($rd.Read()) {
            $ori = if ($rd['Horientacion'] -is [System.DBNull]) { '' } else { ([string]$rd['Horientacion']).Trim().ToUpperInvariant() }
            [void]$calles.Add([ordered]@{
                id = [int]$rd['Id']
                nombre = (Repair-Text $rd['NombreCalles'])
                orientacion = $ori
            })
        }
        $rd.Close()

        $cmd = $con.CreateCommand()
        $cmd.CommandText = "SELECT Responsable FROM Responsables ORDER BY Responsable"
        $rd = $cmd.ExecuteReader()
        while ($rd.Read()) { [void]$responsables.Add((Repair-Text $rd['Responsable'])) }
        $rd.Close()

        $cmd = $con.CreateCommand()
        $cmd.CommandText = "SELECT Horario FROM Horarios ORDER BY Horario"
        $rd = $cmd.ExecuteReader()
        while ($rd.Read()) {
            if ($rd['Horario'] -isnot [System.DBNull]) {
                [void]$horarios.Add(([datetime]$rd['Horario']).ToString('HH:mm'))
            }
        }
        $rd.Close()

        try {
            $cmd = $con.CreateCommand()
            $cmd.CommandText = "SELECT DISTINCT Grupos FROM NumeroDeTerritorios WHERE Grupos IS NOT NULL"
            $rd = $cmd.ExecuteReader()
            while ($rd.Read()) {
                $g = if ($rd['Grupos'] -is [System.DBNull]) { '' } else { (Repair-Text $rd['Grupos']).Trim() }
                if (-not [string]::IsNullOrWhiteSpace($g)) { [void]$gruposSet.Add($g) }
            }
            $rd.Close()
        } catch {}
        try {
            $cmd = $con.CreateCommand()
            $cmd.CommandText = "SELECT Grupo FROM Grupos"
            $rd = $cmd.ExecuteReader()
            while ($rd.Read()) {
                $g = if ($rd[0] -is [System.DBNull]) { '' } else { (Repair-Text $rd[0]).Trim() }
                if (-not [string]::IsNullOrWhiteSpace($g)) { [void]$gruposSet.Add($g) }
            }
            $rd.Close()
        } catch {}
        [void]$gruposSet.Add('Congregacional')

        $pred = Get-PredeterminadosConfig $con

        return [ordered]@{
            calles = ,$calles.ToArray()
            responsables = ,$responsables.ToArray()
            horarios = ,$horarios.ToArray()
            grupos = @($gruposSet | Sort-Object)
            predeterminados = $pred
        }
    }
    finally { $con.Close() }
}

function Ensure-GruposTable {
    param($con)
    try {
        $cmd = $con.CreateCommand()
        $cmd.CommandText = "SELECT TOP 1 Grupo FROM Grupos"
        [void]$cmd.ExecuteScalar()
    } catch {
        $crt = $con.CreateCommand()
        $crt.CommandText = "CREATE TABLE Grupos ([Grupo] TEXT(120))"
        [void]$crt.ExecuteNonQuery()
    }
}

function Ensure-PredeterminadosTable {
    param($con)
    try {
        $cmd = $con.CreateCommand()
        $cmd.CommandText = "SELECT TOP 1 TipoDia, Horario FROM HorariosPredeterminados"
        [void]$cmd.ExecuteScalar()
    } catch {
        $crt = $con.CreateCommand()
        $crt.CommandText = "CREATE TABLE HorariosPredeterminados ([TipoDia] TEXT(30), [Horario] DATETIME, [Orden] INTEGER)"
        [void]$crt.ExecuteNonQuery()
    }
}

function Get-PredeterminadosConfig {
    param($con)
    Ensure-PredeterminadosTable $con
    $cfg = [ordered]@{
        entresemana = @()
        sabado = @()
        domingo = @()
    }
    $cmd = $con.CreateCommand()
    $cmd.CommandText = "SELECT TipoDia, Horario FROM HorariosPredeterminados ORDER BY TipoDia, Orden, Horario"
    $rd = $cmd.ExecuteReader()
    while ($rd.Read()) {
        $tipo = if ($rd['TipoDia'] -is [System.DBNull]) { '' } else { ([string]$rd['TipoDia']).Trim().ToLowerInvariant() }
        if ($tipo -ne 'entresemana' -and $tipo -ne 'sabado' -and $tipo -ne 'domingo') { continue }
        if ($rd['Horario'] -is [System.DBNull]) { continue }
        $cfg[$tipo] += ([datetime]$rd['Horario']).ToString('HH:mm')
    }
    $rd.Close()

    if ($cfg.entresemana.Count -eq 0) { $cfg.entresemana = @('10:00', '17:00') }
    if ($cfg.sabado.Count -eq 0) { $cfg.sabado = @('10:00') }
    if ($cfg.domingo.Count -eq 0) { $cfg.domingo = @('10:30') }
    return $cfg
}

function Set-PredeterminadosConfig {
    param($body)
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        Ensure-PredeterminadosTable $con
        $del = $con.CreateCommand()
        $del.CommandText = "DELETE FROM HorariosPredeterminados"
        [void]$del.ExecuteNonQuery()

        $insertOne = {
            param([string]$tipo, [string[]]$arr)
            $orden = 1
            foreach ($txt in $arr) {
                $m = [regex]::Match(([string]$txt).Trim(), '^(\d{1,2}):(\d{2})$')
                if (-not $m.Success) { continue }
                $h = [int]$m.Groups[1].Value; $mm = [int]$m.Groups[2].Value
                if ($h -lt 0 -or $h -gt 23 -or $mm -lt 0 -or $mm -gt 59) { continue }
                $dt = [datetime]::Today.AddHours($h).AddMinutes($mm)
                $cmd = $con.CreateCommand()
                $cmd.CommandText = "INSERT INTO HorariosPredeterminados (TipoDia, Horario, Orden) VALUES (?, ?, ?)"
                [void]$cmd.Parameters.AddWithValue('@t', $tipo)
                $pH = $cmd.Parameters.Add('@h', [System.Data.OleDb.OleDbType]::Date)
                $pH.Value = $dt
                [void]$cmd.Parameters.AddWithValue('@o', $orden)
                [void]$cmd.ExecuteNonQuery()
                $orden++
            }
        }

        & $insertOne 'entresemana' @($body.entresemana)
        & $insertOne 'sabado' @($body.sabado)
        & $insertOne 'domingo' @($body.domingo)
    }
    finally { $con.Close() }
}

function Add-CatalogoItem {
    param([string]$tipo, $body)
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        switch ($tipo.ToLowerInvariant()) {
            'responsables' {
                $nombre = ([string]$body.nombre).Trim()
                if ([string]::IsNullOrWhiteSpace($nombre)) { throw "Nombre vacío." }
                $cmd = $con.CreateCommand()
                $cmd.CommandText = "INSERT INTO Responsables (Responsable) VALUES (?)"
                [void]$cmd.Parameters.AddWithValue('@n', $nombre)
                [void]$cmd.ExecuteNonQuery()
                break
            }
            'horarios' {
                $txt = ([string]$body.hora).Trim()
                $m = [regex]::Match($txt, '^(\d{1,2}):(\d{2})$')
                if (-not $m.Success) { throw "Hora inválida. Use HH:mm." }
                $h = [int]$m.Groups[1].Value; $mm = [int]$m.Groups[2].Value
                if ($h -lt 0 -or $h -gt 23 -or $mm -lt 0 -or $mm -gt 59) { throw "Hora inválida." }
                $d = [datetime]::Today.AddHours($h).AddMinutes($mm)
                $cmd = $con.CreateCommand()
                $cmd.CommandText = "INSERT INTO Horarios (Horario) VALUES (?)"
                $p = $cmd.Parameters.Add('@h', [System.Data.OleDb.OleDbType]::Date)
                $p.Value = $d
                [void]$cmd.ExecuteNonQuery()
                break
            }
            'calles' {
                $nombre = ([string]$body.nombre).Trim()
                $ori = ([string]$body.orientacion).Trim().ToUpperInvariant()
                if ([string]::IsNullOrWhiteSpace($nombre)) { throw "Nombre de calle vacío." }
                if ($ori -ne 'H' -and $ori -ne 'V' -and $ori -ne '') { throw "Orientación inválida (H/V)." }
                $cmd = $con.CreateCommand()
                $cmd.CommandText = "INSERT INTO Calles (NombreCalles, Horientacion) VALUES (?, ?)"
                [void]$cmd.Parameters.AddWithValue('@n', $nombre)
                [void]$cmd.Parameters.AddWithValue('@o', $ori)
                [void]$cmd.ExecuteNonQuery()
                break
            }
            'grupos' {
                $nombre = ([string]$body.nombre).Trim()
                if ([string]::IsNullOrWhiteSpace($nombre)) { throw "Nombre de grupo vacío." }
                Ensure-GruposTable $con
                $chk = $con.CreateCommand()
                $chk.CommandText = "SELECT 1 FROM Grupos WHERE Grupo = ?"
                [void]$chk.Parameters.AddWithValue('@g', $nombre)
                $ex = $chk.ExecuteScalar()
                if ($null -eq $ex -or $ex -is [System.DBNull]) {
                    $cmd = $con.CreateCommand()
                    $cmd.CommandText = "INSERT INTO Grupos (Grupo) VALUES (?)"
                    [void]$cmd.Parameters.AddWithValue('@g', $nombre)
                    [void]$cmd.ExecuteNonQuery()
                }
                break
            }
            default { throw "Tipo de catálogo no soportado: $tipo" }
        }
    }
    finally { $con.Close() }
}

function Delete-CatalogoItem {
    param([string]$tipo, $body)
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        switch ($tipo.ToLowerInvariant()) {
            'responsables' {
                $nombre = ([string]$body.nombre).Trim()
                $cmd = $con.CreateCommand()
                $cmd.CommandText = "DELETE FROM Responsables WHERE Responsable = ?"
                [void]$cmd.Parameters.AddWithValue('@n', $nombre)
                [void]$cmd.ExecuteNonQuery()
                break
            }
            'horarios' {
                $txt = ([string]$body.hora).Trim()
                $m = [regex]::Match($txt, '^(\d{1,2}):(\d{2})$')
                if (-not $m.Success) { throw "Hora inválida." }
                $h = [int]$m.Groups[1].Value; $mm = [int]$m.Groups[2].Value
                $d = [datetime]::Today.AddHours($h).AddMinutes($mm)
                $cmd = $con.CreateCommand()
                $cmd.CommandText = "DELETE FROM Horarios WHERE Horario = ?"
                $p = $cmd.Parameters.Add('@h', [System.Data.OleDb.OleDbType]::Date)
                $p.Value = $d
                [void]$cmd.ExecuteNonQuery()
                break
            }
            'calles' {
                $id = [int]$body.id
                $cmd = $con.CreateCommand()
                $cmd.CommandText = "DELETE FROM Calles WHERE Id = ?"
                [void]$cmd.Parameters.AddWithValue('@id', $id)
                [void]$cmd.ExecuteNonQuery()
                break
            }
            'grupos' {
                $nombre = ([string]$body.nombre).Trim()
                try {
                    Ensure-GruposTable $con
                    $cmd = $con.CreateCommand()
                    $cmd.CommandText = "DELETE FROM Grupos WHERE Grupo = ?"
                    [void]$cmd.Parameters.AddWithValue('@g', $nombre)
                    [void]$cmd.ExecuteNonQuery()
                } catch {}
                $up = $con.CreateCommand()
                $up.CommandText = "UPDATE NumeroDeTerritorios SET Grupos = NULL WHERE Grupos = ?"
                [void]$up.Parameters.AddWithValue('@g', $nombre)
                [void]$up.ExecuteNonQuery()
                break
            }
            default { throw "Tipo de catálogo no soportado: $tipo" }
        }
    }
    finally { $con.Close() }
}

function Read-Esquinas {
    param([string]$numero)
    $sql = @"
SELECT R.[Id Calle1] AS Id1, R.[Id Calle2] AS Id2,
       C1.[NombreCalles] AS N1, C2.[NombreCalles] AS N2
FROM ([RelacionTerritoriopuntosdeencuentro] AS R
INNER JOIN [Calles] AS C1 ON R.[Id Calle1] = C1.[Id])
LEFT JOIN [Calles] AS C2 ON R.[Id Calle2] = C2.[Id]
WHERE R.[NumeroTerritorio] = ?
ORDER BY C1.[NombreCalles], C2.[NombreCalles]
"@
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $numTrim = Normalize-NumeroTerritorio $numero
        if ([string]::IsNullOrWhiteSpace($numTrim)) { return ,@() }
        $item = $null
        try { $item = Get-TerritorioRow $numTrim $con } catch { return ,@() }
        $pNum = Param-OleDbNumeroTerritorio ([string]$item.numero)
        $cmd = $con.CreateCommand(); $cmd.CommandText = $sql
        [void]$cmd.Parameters.AddWithValue('@n', $pNum)
        $rd = $cmd.ExecuteReader()
        $lista = New-Object System.Collections.ArrayList
        $seen = New-Object 'System.Collections.Generic.HashSet[string]'
        while ($rd.Read()) {
            $id1 = [int]$rd['Id1']
            $id2 = if ($rd['Id2'] -is [System.DBNull]) { 0 } else { [int]$rd['Id2'] }
            if ($id2 -gt 0 -and $id2 -lt $id1) { $tmp = $id1; $id1 = $id2; $id2 = $tmp }
            $k = "$id1|$id2"
            if ($seen.Contains($k)) { continue }
            [void]$seen.Add($k)
            $n1 = (Repair-Text $rd['N1'])
            $n2 = if ($rd['N2'] -is [System.DBNull]) { '' } else { (Repair-Text $rd['N2']) }
            $txt = if ([string]::IsNullOrWhiteSpace($n2)) { $n1 } else { "$n1 y $n2" }
            [void]$lista.Add([ordered]@{ id1 = $id1; id2 = $id2; texto = $txt })
        }
        $rd.Close()
        return ,$lista.ToArray()
    }
    finally { $con.Close() }
}

function Add-Esquina {
    param([string]$numero, [int]$id1, [int]$id2)
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $numTrim = Normalize-NumeroTerritorio $numero
        if ([string]::IsNullOrWhiteSpace($numTrim)) { throw "Numero de territorio requerido." }
        $item = Get-TerritorioRow $numTrim $con
        $pNum = Param-OleDbNumeroTerritorio ([string]$item.numero)
        $chk = $con.CreateCommand()
        if ($id2 -gt 0) {
            $chk.CommandText = "SELECT 1 FROM [RelacionTerritoriopuntosdeencuentro] WHERE [NumeroTerritorio]=? AND (([Id Calle1]=? AND [Id Calle2]=?) OR ([Id Calle1]=? AND [Id Calle2]=?))"
            [void]$chk.Parameters.AddWithValue('@n', $pNum)
            [void]$chk.Parameters.AddWithValue('@a1', $id1)
            [void]$chk.Parameters.AddWithValue('@a2', $id2)
            [void]$chk.Parameters.AddWithValue('@b1', $id2)
            [void]$chk.Parameters.AddWithValue('@b2', $id1)
        } else {
            # Punto simple: en Access se guarda Id Calle2 como NULL, no como 0.
            $chk.CommandText = "SELECT 1 FROM [RelacionTerritoriopuntosdeencuentro] WHERE [NumeroTerritorio]=? AND [Id Calle1]=? AND [Id Calle2] IS NULL"
            [void]$chk.Parameters.AddWithValue('@n', $pNum)
            [void]$chk.Parameters.AddWithValue('@a1', $id1)
        }
        $exists = $chk.ExecuteScalar()
        if ($null -ne $exists -and $exists -isnot [System.DBNull]) { return }

        $cmd = $con.CreateCommand()
        $cmd.CommandText = "INSERT INTO [RelacionTerritoriopuntosdeencuentro] ([NumeroTerritorio], [Id Calle1], [Id Calle2]) VALUES (?, ?, ?)"
        [void]$cmd.Parameters.AddWithValue('@n', $pNum)
        [void]$cmd.Parameters.AddWithValue('@c1', $id1)
        $p = $cmd.Parameters.Add('@c2', [System.Data.OleDb.OleDbType]::Integer)
        if ($id2 -gt 0) { $p.Value = $id2 } else { $p.Value = [DBNull]::Value }
        [void]$cmd.ExecuteNonQuery()
    }
    finally { $con.Close() }
}

function Remove-Esquina {
    param([string]$numero, [int]$id1, [int]$id2)
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $numTrim = Normalize-NumeroTerritorio $numero
        if ([string]::IsNullOrWhiteSpace($numTrim)) { throw "Numero de territorio requerido." }
        $item = Get-TerritorioRow $numTrim $con
        $pNum = Param-OleDbNumeroTerritorio ([string]$item.numero)
        $cmd = $con.CreateCommand()
        if ($id2 -gt 0) {
            # Borrar sin importar el orden de las calles (A,B) o (B,A)
            $cmd.CommandText = "DELETE FROM [RelacionTerritoriopuntosdeencuentro] WHERE [NumeroTerritorio]=? AND (([Id Calle1]=? AND [Id Calle2]=?) OR ([Id Calle1]=? AND [Id Calle2]=?))"
            [void]$cmd.Parameters.AddWithValue('@n', $pNum)
            [void]$cmd.Parameters.AddWithValue('@c1', $id1)
            [void]$cmd.Parameters.AddWithValue('@c2', $id2)
            [void]$cmd.Parameters.AddWithValue('@c1b', $id2)
            [void]$cmd.Parameters.AddWithValue('@c2b', $id1)
        } else {
            # Punto simple: tolerar NULL o 0 en Id Calle2
            $cmd.CommandText = "DELETE FROM [RelacionTerritoriopuntosdeencuentro] WHERE [NumeroTerritorio]=? AND [Id Calle1]=? AND ([Id Calle2] IS NULL OR [Id Calle2]=0)"
            [void]$cmd.Parameters.AddWithValue('@n', $pNum)
            [void]$cmd.Parameters.AddWithValue('@c1', $id1)
        }
        [void]$cmd.ExecuteNonQuery()
    }
    finally { $con.Close() }
}

function Read-Existentes {
    param([string]$punto, [datetime]$fecha, [string]$grupo)
    $puntoRaw = [string]$punto
    $puntoFix = (Repair-Text $puntoRaw)
    $grupoFix = (Repair-Text $grupo)
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $cmd = $con.CreateCommand()
        $cmd.CommandText = "SELECT NumeroTerritorio, Prioridad FROM NumeroDeTerritorios " +
                           "WHERE (PuntosDeEncuentros = ? OR PuntosDeEncuentros = ?) AND Enviado = ? " +
                           "AND IIf(Grupos Is Null OR Trim(Grupos)='', 'Congregacional', Trim(Grupos)) = ? " +
                           "ORDER BY Prioridad"
        [void]$cmd.Parameters.AddWithValue('@p1', $puntoRaw)
        [void]$cmd.Parameters.AddWithValue('@p2', $puntoFix)
        $pf = $cmd.Parameters.Add('@f', [System.Data.OleDb.OleDbType]::Date)
        $pf.Value = $fecha
        [void]$cmd.Parameters.AddWithValue('@g', $grupoFix)
        $rd = $cmd.ExecuteReader()
        $lista = New-Object System.Collections.ArrayList
        while ($rd.Read()) {
            [void]$lista.Add([ordered]@{
                numero = [string]$rd['NumeroTerritorio']
                prioridad = if ($rd['Prioridad'] -is [System.DBNull]) { 0 } else { [int]$rd['Prioridad'] }
            })
        }
        $rd.Close()
        return ,$lista.ToArray()
    }
    finally { $con.Close() }
}

function Set-Asignacion {
    param($body)
    $numero = Normalize-NumeroTerritorio $body.numero
    if ([string]::IsNullOrWhiteSpace($numero)) {
        throw "Numero de territorio requerido."
    }
    $otros = @()
    if ($null -ne $body.otrosNumeros) {
        foreach ($n in $body.otrosNumeros) {
            $ns = ([string]$n).Trim()
            if (-not [string]::IsNullOrWhiteSpace($ns)) { $otros += $ns }
        }
    }
    $numeros = @($numero) + $otros | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    $fecha = [datetime]::Parse([string]$body.diaHora)
    $puntoRaw = [string]$body.punto
    $punto = (Repair-Text $puntoRaw)
    $grupo = (Repair-Text $body.grupo)
    $responsable = (Repair-Text $body.responsable)
    $prioritario = [bool]$body.prioritario
    $prioridadDeseada = 0
    try { $prioridadDeseada = [int]$body.prioridadDeseada } catch { $prioridadDeseada = 0 }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $cmdCount = $con.CreateCommand()
        $cmdCount.CommandText = "SELECT COUNT(*) FROM NumeroDeTerritorios " +
                                "WHERE (PuntosDeEncuentros = ? OR PuntosDeEncuentros = ?) AND Enviado = ? " +
                                "AND IIf(Grupos Is Null OR Trim(Grupos)='', 'Congregacional', Trim(Grupos)) = ?"
        [void]$cmdCount.Parameters.AddWithValue('@p1', $puntoRaw)
        [void]$cmdCount.Parameters.AddWithValue('@p2', $punto)
        $pc = $cmdCount.Parameters.Add('@f', [System.Data.OleDb.OleDbType]::Date)
        $pc.Value = $fecha
        [void]$cmdCount.Parameters.AddWithValue('@g', $grupo)
        $existentes = [int]$cmdCount.ExecuteScalar()

        if ($prioridadDeseada -gt 0 -and $existentes -gt 0) {
            if ($prioridadDeseada -gt ($existentes + 1)) { $prioridadDeseada = $existentes + 1 }
            $cmdUp = $con.CreateCommand()
            $cmdUp.CommandText = "UPDATE NumeroDeTerritorios SET Prioridad = Prioridad + ? " +
                                 "WHERE (PuntosDeEncuentros = ? OR PuntosDeEncuentros = ?) AND Enviado = ? " +
                                 "AND IIf(Grupos Is Null OR Trim(Grupos)='', 'Congregacional', Trim(Grupos)) = ? " +
                                 "AND Prioridad >= ?"
            [void]$cmdUp.Parameters.AddWithValue('@inc', [int]$numeros.Count)
            [void]$cmdUp.Parameters.AddWithValue('@p1', $puntoRaw)
            [void]$cmdUp.Parameters.AddWithValue('@p2', $punto)
            $pu = $cmdUp.Parameters.Add('@f', [System.Data.OleDb.OleDbType]::Date)
            $pu.Value = $fecha
            [void]$cmdUp.Parameters.AddWithValue('@g', $grupo)
            [void]$cmdUp.Parameters.AddWithValue('@pd', $prioridadDeseada)
            [void]$cmdUp.ExecuteNonQuery()
            $prioridadActual = $prioridadDeseada
        }
        elseif ($prioritario -and $existentes -gt 0) {
            $cmdUp = $con.CreateCommand()
            $cmdUp.CommandText = "UPDATE NumeroDeTerritorios SET Prioridad = Prioridad + ? " +
                                 "WHERE (PuntosDeEncuentros = ? OR PuntosDeEncuentros = ?) AND Enviado = ? " +
                                 "AND IIf(Grupos Is Null OR Trim(Grupos)='', 'Congregacional', Trim(Grupos)) = ?"
            [void]$cmdUp.Parameters.AddWithValue('@inc', [int]$numeros.Count)
            [void]$cmdUp.Parameters.AddWithValue('@p1', $puntoRaw)
            [void]$cmdUp.Parameters.AddWithValue('@p2', $punto)
            $pu = $cmdUp.Parameters.Add('@f', [System.Data.OleDb.OleDbType]::Date)
            $pu.Value = $fecha
            [void]$cmdUp.Parameters.AddWithValue('@g', $grupo)
            [void]$cmdUp.ExecuteNonQuery()
            $prioridadActual = 1
        } else {
            $prioridadActual = $existentes + 1
        }

        foreach ($n in $numeros) {
            $cmd = $con.CreateCommand()
            $cmd.CommandText = "UPDATE NumeroDeTerritorios SET AbiertoCerrado=True, Enviado=?, PuntosDeEncuentros=?, Grupos=?, Responsables=?, Asignado=IIf(IsNull(Asignado), ?, Asignado), Prioridad=? WHERE NumeroTerritorio=?"
            $p1 = $cmd.Parameters.Add('@env', [System.Data.OleDb.OleDbType]::Date); $p1.Value = $fecha
            [void]$cmd.Parameters.AddWithValue('@pt', $punto)
            [void]$cmd.Parameters.AddWithValue('@gr', $grupo)
            [void]$cmd.Parameters.AddWithValue('@re', $responsable)
            $p2 = $cmd.Parameters.Add('@as', [System.Data.OleDb.OleDbType]::Date); $p2.Value = $fecha
            [void]$cmd.Parameters.AddWithValue('@pr', [int]$prioridadActual)
            [void]$cmd.Parameters.AddWithValue('@n', (Param-OleDbNumeroTerritorio ([string]$n)))
            [void]$cmd.ExecuteNonQuery()
            $prioridadActual++
        }

        return [ordered]@{
            ok = $true
            asignados = @($numeros)
            fecha = $fecha.ToString('yyyy-MM-ddTHH:mm:ss')
        }
    }
    finally { $con.Close() }
}

function Get-TerritorioRow {
    param([string]$numeroRaw, $con)

    $numeroTxt = Normalize-NumeroTerritorio $numeroRaw
    if ([string]::IsNullOrWhiteSpace($numeroTxt)) {
        throw "Numero de territorio requerido."
    }

    $candidatos = New-Object System.Collections.ArrayList
    [void]$candidatos.Add($numeroTxt)
    if ($numeroTxt.Contains('.') -and -not $numeroTxt.Contains(',')) {
        [void]$candidatos.Add($numeroTxt.Replace('.', ','))
    }
    if ($numeroTxt.Contains(',') -and -not $numeroTxt.Contains('.')) {
        [void]$candidatos.Add($numeroTxt.Replace(',', '.'))
    }

    $vistos = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($cand in $candidatos) {
        $c = ([string]$cand).Trim()
        if ([string]::IsNullOrWhiteSpace($c) -or $vistos.Contains($c)) { continue }
        [void]$vistos.Add($c)

        $cmdSel = $con.CreateCommand()
        $cmdSel.CommandText = "SELECT TOP 1 * FROM NumeroDeTerritorios WHERE CStr(NumeroTerritorio)=?"
        [void]$cmdSel.Parameters.AddWithValue('@num', $c)
        $rd = $cmdSel.ExecuteReader()
        if ($rd.Read()) {
            $item = [ordered]@{
                numero = [string]$rd['NumeroTerritorio']
                enviado = $rd['Enviado']
                asignado = $rd['Asignado']
                responsables = $rd['Responsables']
            }
            $rd.Close()
            return $item
        }
        $rd.Close()
    }

    $pNum = Param-OleDbNumeroTerritorio $numeroTxt
    $cmdSel = $con.CreateCommand()
    $cmdSel.CommandText = "SELECT TOP 1 * FROM NumeroDeTerritorios WHERE NumeroTerritorio=?"
    [void]$cmdSel.Parameters.AddWithValue('@num', $pNum)
    $rd = $cmdSel.ExecuteReader()
    if ($rd.Read()) {
        $item = [ordered]@{
            numero = [string]$rd['NumeroTerritorio']
            enviado = $rd['Enviado']
            asignado = $rd['Asignado']
            responsables = $rd['Responsables']
        }
        $rd.Close()
        return $item
    }
    $rd.Close()

    throw "No se encontro el territorio $numeroTxt."
}

function Add-HistorialSingleFromTerritorio {
    param([string]$numero, $enviadoObj, $asignadoObj, $encargadoObj, $con)

    $cmdIns = $con.CreateCommand()
    $cmdIns.CommandText = "INSERT INTO Historial (NumeroTerritorio, Terminado, Encargado, Iniciado) VALUES (?, ?, ?, ?)"
    [void]$cmdIns.Parameters.AddWithValue('@n', $numero)

    $pTerm = $cmdIns.Parameters.Add('@t', [System.Data.OleDb.OleDbType]::Date)
    if ($enviadoObj -is [System.DBNull] -or $null -eq $enviadoObj) { $pTerm.Value = [DBNull]::Value } else { $pTerm.Value = [datetime]$enviadoObj }

    $encargadoTxt = ''
    if ($encargadoObj -isnot [System.DBNull] -and $null -ne $encargadoObj) {
        $encargadoTxt = [string]$encargadoObj
    }
    [void]$cmdIns.Parameters.AddWithValue('@e', $encargadoTxt)

    $pIni = $cmdIns.Parameters.Add('@i', [System.Data.OleDb.OleDbType]::Date)
    if ($asignadoObj -is [System.DBNull] -or $null -eq $asignadoObj) { $pIni.Value = [DBNull]::Value } else { $pIni.Value = [datetime]$asignadoObj }

    [void]$cmdIns.ExecuteNonQuery()
}

function Apply-TerritorioAjusteSalida {
    param([string]$numeroTxt, $body, $con)

    if ($null -eq $body) { return }

    $respIn = $body.responsableAjuste
    $envIn = $body.enviadoAjuste

    $hasResp = $null -ne $respIn -and -not [string]::IsNullOrWhiteSpace([string]$respIn)
    $hasEnv = $null -ne $envIn -and -not [string]::IsNullOrWhiteSpace([string]$envIn)

    if (-not $hasResp -and -not $hasEnv) { return }

    $item = Get-TerritorioRow $numeroTxt $con
    $key = [string]$item.numero

    $newResp = if ($hasResp) {
        [string]$respIn.Trim()
    }
    else {
        if ($item.responsables -is [System.DBNull]) { '' } else { [string]$item.responsables }
    }

    $newEnvVal = $null
    if ($hasEnv) {
        $s = [string]$envIn.Trim()
        $parsed = $null
        try {
            $parsed = [datetime]::Parse($s, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal)
        }
        catch {
            try { $parsed = [datetime]::Parse($s) } catch { throw "Fecha/hora invalida: $s" }
        }
        $newEnvVal = $parsed
    }
    else {
        if ($item.enviado -is [System.DBNull]) {
            $newEnvVal = [DBNull]::Value
        }
        else {
            $newEnvVal = [datetime]$item.enviado
        }
    }

    $cmdUp = $con.CreateCommand()
    if ($hasResp -and $hasEnv) {
        $cmdUp.CommandText = "UPDATE NumeroDeTerritorios SET Responsables=?, Enviado=? WHERE CStr(NumeroTerritorio)=?"
        [void]$cmdUp.Parameters.AddWithValue('@r', $newResp)
        $pE = $cmdUp.Parameters.Add('@e', [System.Data.OleDb.OleDbType]::Date)
        if ($newEnvVal -is [System.DBNull]) { $pE.Value = [DBNull]::Value } else { $pE.Value = [datetime]$newEnvVal }
        [void]$cmdUp.Parameters.AddWithValue('@num', $key)
    }
    elseif ($hasResp) {
        $cmdUp.CommandText = "UPDATE NumeroDeTerritorios SET Responsables=? WHERE CStr(NumeroTerritorio)=?"
        [void]$cmdUp.Parameters.AddWithValue('@r', $newResp)
        [void]$cmdUp.Parameters.AddWithValue('@num', $key)
    }
    else {
        $cmdUp.CommandText = "UPDATE NumeroDeTerritorios SET Enviado=? WHERE CStr(NumeroTerritorio)=?"
        $pE = $cmdUp.Parameters.Add('@e', [System.Data.OleDb.OleDbType]::Date)
        if ($newEnvVal -is [System.DBNull]) { $pE.Value = [DBNull]::Value } else { $pE.Value = [datetime]$newEnvVal }
        [void]$cmdUp.Parameters.AddWithValue('@num', $key)
    }
    [void]$cmdUp.ExecuteNonQuery()
}

function Save-TerritorioEdicionSalida {
    param($body)

    $numeroTxt = ([string]$body.numero).Trim()
    if ([string]::IsNullOrWhiteSpace($numeroTxt)) {
        throw "Numero de territorio requerido."
    }

    $respIn = $body.responsableAjuste
    $envIn = $body.enviadoAjuste
    $hasResp = $null -ne $respIn -and -not [string]::IsNullOrWhiteSpace([string]$respIn)
    $hasEnv = $null -ne $envIn -and -not [string]::IsNullOrWhiteSpace([string]$envIn)
    if (-not $hasResp -and -not $hasEnv) {
        throw "Indique responsable y/o fecha y hora de salida."
    }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        Apply-TerritorioAjusteSalida $numeroTxt $body $con
        $item = Get-TerritorioRow $numeroTxt $con
        return [ordered]@{
            ok = $true
            numero = [string]$item.numero
        }
    }
    finally { $con.Close() }
}

function Save-TerritorioGrupo {
    param($body)

    $numeroTxt = ([string]$body.numero).Trim()
    if ([string]::IsNullOrWhiteSpace($numeroTxt)) {
        throw "Numero de territorio requerido."
    }

    $grupoRaw = $body.grupo
    $grupoTxt = if ($null -eq $grupoRaw) { '' } else { (Repair-Text ([string]$grupoRaw)).Trim() }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $item = Get-TerritorioRow $numeroTxt $con
        $key = [string]$item.numero

        $cmdUp = $con.CreateCommand()
        $cmdUp.CommandText = "UPDATE NumeroDeTerritorios SET Grupos=? WHERE CStr(NumeroTerritorio)=?"
        $pG = $cmdUp.Parameters.Add('@g', [System.Data.OleDb.OleDbType]::VarWChar, 120)
        if ([string]::IsNullOrWhiteSpace($grupoTxt)) {
            $pG.Value = [DBNull]::Value
        }
        else {
            $pG.Value = $grupoTxt
        }
        [void]$cmdUp.Parameters.AddWithValue('@num', $key)
        $rows = [int]$cmdUp.ExecuteNonQuery()

        return [ordered]@{
            ok = ($rows -gt 0)
            numero = $key
            grupo = if ([string]::IsNullOrWhiteSpace($grupoTxt)) { $null } else { $grupoTxt }
        }
    }
    finally { $con.Close() }
}

function Close-TerritorioSinHecho {
    param($body)

    $numeroTxt = ([string]$body.numero).Trim()
    if ([string]::IsNullOrWhiteSpace($numeroTxt)) {
        throw "Numero de territorio requerido."
    }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        Apply-TerritorioAjusteSalida $numeroTxt $body $con

        $cmdSel = $con.CreateCommand()
        $cmdSel.CommandText = "SELECT TOP 1 NumeroTerritorio, Enviado, Asignado FROM NumeroDeTerritorios WHERE CStr(NumeroTerritorio)=?"
        [void]$cmdSel.Parameters.AddWithValue('@num', $numeroTxt)
        $rd = $cmdSel.ExecuteReader()

        if (-not $rd.Read()) {
            $rd.Close()
            throw "No se encontro el territorio $numeroTxt."
        }

        $dbNumero = [string]$rd['NumeroTerritorio']
        $enviadoObj = $rd['Enviado']
        $asignadoObj = $rd['Asignado']
        $rd.Close()

        $limpiarAsignado = $false
        if (
            $enviadoObj -isnot [System.DBNull] -and
            $asignadoObj -isnot [System.DBNull] -and
            ([datetime]$enviadoObj -eq [datetime]$asignadoObj)
        ) {
            # Regla original del VBA para BOTON 2:
            # solo limpia Asignado si Enviado = Asignado.
            $limpiarAsignado = $true
        }

        $sql = "UPDATE NumeroDeTerritorios SET " +
               "AbiertoCerrado=False, Enviado=Null, PuntosDeEncuentros=Null, Grupos=Null, Responsables=Null, Prioridad=Null" +
               $(if ($limpiarAsignado) { ", Asignado=Null" } else { "" }) +
               " WHERE CStr(NumeroTerritorio)=?"
        $cmdUp = $con.CreateCommand()
        $cmdUp.CommandText = $sql
        [void]$cmdUp.Parameters.AddWithValue('@num', $numeroTxt)
        $rows = [int]$cmdUp.ExecuteNonQuery()

        return [ordered]@{
            ok = ($rows -gt 0)
            numero = $dbNumero
            asignadoLimpiado = $limpiarAsignado
        }
    }
    finally { $con.Close() }
}

function Close-TerritorioConAccion {
    param($body)

    $numeroTxt = ([string]$body.numero).Trim()
    $accion = ([string]$body.accion).Trim().ToLowerInvariant()
    $manzanasTexto = ([string]$body.manzanasPendientes).Trim()

    if ([string]::IsNullOrWhiteSpace($accion)) {
        throw "Accion requerida."
    }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        Apply-TerritorioAjusteSalida $numeroTxt $body $con

        $item = Get-TerritorioRow $numeroTxt $con
        $numero = [string]$item.numero
        $enviadoObj = $item.enviado
        $asignadoObj = $item.asignado
        $responsablesObj = $item.responsables

        switch ($accion) {
            'completado' {
                $limpiarAsignado = $true
                if (
                    $enviadoObj -is [System.DBNull] -or $null -eq $enviadoObj -or
                    $asignadoObj -is [System.DBNull] -or $null -eq $asignadoObj
                ) {
                    $limpiarAsignado = $true
                } elseif ([datetime]$enviadoObj -eq [datetime]$asignadoObj) {
                    $limpiarAsignado = $true
                }

                $sql = "UPDATE NumeroDeTerritorios SET " +
                       "AbiertoCerrado=False, Enviado=Null, PuntosDeEncuentros=Null, Grupos=Null, Responsables=Null, Prioridad=Null, " +
                       "ManzanasPendientes=Null, FechaManzanasPendientes=Null, UltimoEncargado=Null" +
                       $(if ($limpiarAsignado) { ", Asignado=Null" } else { "" }) +
                       " WHERE CStr(NumeroTerritorio)=?"
                $cmdUp = $con.CreateCommand()
                $cmdUp.CommandText = $sql
                [void]$cmdUp.Parameters.AddWithValue('@num', $numeroTxt)
                [void]$cmdUp.ExecuteNonQuery()

                Add-HistorialSingleFromTerritorio $numero $enviadoObj $asignadoObj $responsablesObj $con
                return [ordered]@{ ok = $true; numero = $numero; accion = 'completado' }
            }
            'completado-campana' {
                $enie = [char]0x00F1
                $sql = "UPDATE NumeroDeTerritorios SET " +
                       "AbiertoCerrado=False, [Campa${enie}a]=True, Enviado=Null, PuntosDeEncuentros=Null, Grupos=Null, Responsables=Null, " +
                       "Asignado=Null, Prioridad=Null, ManzanasPendientes=Null, FechaManzanasPendientes=Null, UltimoEncargado=Null " +
                       "WHERE CStr(NumeroTerritorio)=?"
                $cmdUp = $con.CreateCommand()
                $cmdUp.CommandText = $sql
                [void]$cmdUp.Parameters.AddWithValue('@num', $numeroTxt)
                [void]$cmdUp.ExecuteNonQuery()

                Add-HistorialSingleFromTerritorio $numero $enviadoObj $asignadoObj $responsablesObj $con
                return [ordered]@{ ok = $true; numero = $numero; accion = 'completado-campana' }
            }
            'incompleto' {
                if ([string]::IsNullOrWhiteSpace($manzanasTexto)) {
                    throw "Para incompleto debe ingresar Manzanas Pendientes."
                }

                $fechaRefObj = if ($enviadoObj -is [System.DBNull] -or $null -eq $enviadoObj) { [datetime]::Today } else { [datetime]$enviadoObj }
                $encargado = if ($responsablesObj -is [System.DBNull] -or $null -eq $responsablesObj) { [DBNull]::Value } else { [string]$responsablesObj }

                $sql = "UPDATE NumeroDeTerritorios SET " +
                       "AbiertoCerrado=False, Enviado=Null, PuntosDeEncuentros=Null, Grupos=Null, Responsables=Null, Prioridad=Null, " +
                       "ManzanasPendientes=?, FechaManzanasPendientes=?, UltimoEncargado=? " +
                       "WHERE CStr(NumeroTerritorio)=?"
                $cmdUp = $con.CreateCommand()
                $cmdUp.CommandText = $sql
                [void]$cmdUp.Parameters.AddWithValue('@mp', $manzanasTexto)
                $pFecha = $cmdUp.Parameters.Add('@fmp', [System.Data.OleDb.OleDbType]::Date)
                $pFecha.Value = $fechaRefObj
                [void]$cmdUp.Parameters.AddWithValue('@ue', $encargado)
                [void]$cmdUp.Parameters.AddWithValue('@num', $numeroTxt)
                [void]$cmdUp.ExecuteNonQuery()

                return [ordered]@{ ok = $true; numero = $numero; accion = 'incompleto' }
            }
            default {
                throw "Accion no soportada: $accion"
            }
        }
    }
    finally { $con.Close() }
}

function Remove-TerritorioIncompleto {
    param($body)

    $numeroTxt = ([string]$body.numero).Trim()
    if ([string]::IsNullOrWhiteSpace($numeroTxt)) {
        throw "Numero de territorio requerido."
    }
    $guardarHistorial = [bool]$body.guardarHistorial

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $item = Get-TerritorioRow $numeroTxt $con
        $numero = [string]$item.numero

        $cmdDatos = $con.CreateCommand()
        $cmdDatos.CommandText = "SELECT TOP 1 AbiertoCerrado, Asignado, ManzanasPendientes, FechaManzanasPendientes, UltimoEncargado FROM NumeroDeTerritorios WHERE CStr(NumeroTerritorio)=?"
        [void]$cmdDatos.Parameters.AddWithValue('@n', $numeroTxt)
        $rd = $cmdDatos.ExecuteReader()
        if (-not $rd.Read()) {
            $rd.Close()
            throw "No se encontro el territorio $numeroTxt."
        }
        $abiertoCerrado = [bool]($rd['AbiertoCerrado'] -as [bool])
        $asignadoObj = $rd['Asignado']
        $mpObj = $rd['ManzanasPendientes']
        $fechaMpObj = $rd['FechaManzanasPendientes']
        $ultEncObj = $rd['UltimoEncargado']
        $rd.Close()

        $mpTxt = if ($mpObj -is [System.DBNull] -or $null -eq $mpObj) { '' } else { ([string]$mpObj).Trim() }
        if ([string]::IsNullOrWhiteSpace($mpTxt)) {
            throw "El territorio $numero no tiene registro de manzanas pendientes."
        }
        if ($abiertoCerrado) {
            throw "El territorio $numero esta enviado/asignado. Solo se puede borrar incompleto cuando no esta enviado."
        }

        if ($guardarHistorial) {
            $terminadoObj = $fechaMpObj
            if ($terminadoObj -is [System.DBNull] -or $null -eq $terminadoObj -or [string]::IsNullOrWhiteSpace([string]$terminadoObj)) {
                $terminadoObj = [datetime]::Today
            }

            $cmdIns = $con.CreateCommand()
            $cmdIns.CommandText = "INSERT INTO Historial (NumeroTerritorio, Terminado, Iniciado, Encargado) VALUES (?, ?, ?, ?)"
            [void]$cmdIns.Parameters.AddWithValue('@n', $numero)
            $pTer = $cmdIns.Parameters.Add('@t', [System.Data.OleDb.OleDbType]::Date)
            $pTer.Value = [datetime]$terminadoObj
            $pIni = $cmdIns.Parameters.Add('@i', [System.Data.OleDb.OleDbType]::Date)
            if ($asignadoObj -is [System.DBNull] -or $null -eq $asignadoObj -or [string]::IsNullOrWhiteSpace([string]$asignadoObj)) { $pIni.Value = [DBNull]::Value } else { $pIni.Value = [datetime]$asignadoObj }
            if ($ultEncObj -is [System.DBNull] -or $null -eq $ultEncObj) {
                [void]$cmdIns.Parameters.AddWithValue('@e', [DBNull]::Value)
            } else {
                [void]$cmdIns.Parameters.AddWithValue('@e', [string]$ultEncObj)
            }
            [void]$cmdIns.ExecuteNonQuery()
        }

        $cmdUp = $con.CreateCommand()
        $cmdUp.CommandText = "UPDATE NumeroDeTerritorios SET Asignado=Null, ManzanasPendientes=Null, FechaManzanasPendientes=Null, UltimoEncargado=Null WHERE CStr(NumeroTerritorio)=?"
        [void]$cmdUp.Parameters.AddWithValue('@n', $numeroTxt)
        [void]$cmdUp.ExecuteNonQuery()

        return [ordered]@{
            ok = $true
            numero = $numero
            historialGuardado = $guardarHistorial
        }
    }
    finally { $con.Close() }
}

function Set-TerritorioRural {
    param($body)

    $numeroTxt = ([string]$body.numero).Trim()
    if ([string]::IsNullOrWhiteSpace($numeroTxt)) {
        throw "Numero de territorio requerido."
    }
    $esRural = [bool]$body.rural

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $cmd = $con.CreateCommand()
        $cmd.CommandText = "UPDATE NumeroDeTerritorios SET Rurales=? WHERE CStr(NumeroTerritorio)=?"
        [void]$cmd.Parameters.AddWithValue('@r', $esRural)
        [void]$cmd.Parameters.AddWithValue('@n', $numeroTxt)
        $rows = [int]$cmd.ExecuteNonQuery()
        if ($rows -le 0) {
            throw "No se encontro el territorio $numeroTxt."
        }
        return [ordered]@{
            ok = $true
            numero = $numeroTxt
            rural = $esRural
        }
    }
    finally { $con.Close() }
}

function Update-TerritorioManzanasPendientes {
    param($body)

    $numeroTxt = ([string]$body.numero).Trim()
    $manzanasTexto = ([string]$body.manzanasPendientes).Trim()
    if ([string]::IsNullOrWhiteSpace($numeroTxt)) {
        throw "Numero de territorio requerido."
    }
    if ([string]::IsNullOrWhiteSpace($manzanasTexto)) {
        throw "El texto de manzanas pendientes no puede estar vacio."
    }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $cmdDatos = $con.CreateCommand()
        $cmdDatos.CommandText = "SELECT TOP 1 ManzanasPendientes FROM NumeroDeTerritorios WHERE CStr(NumeroTerritorio)=?"
        [void]$cmdDatos.Parameters.AddWithValue('@n', $numeroTxt)
        $rd = $cmdDatos.ExecuteReader()
        if (-not $rd.Read()) {
            $rd.Close()
            throw "No se encontro el territorio $numeroTxt."
        }
        $mpObj = $rd['ManzanasPendientes']
        $rd.Close()
        if ($mpObj -is [System.DBNull] -or $null -eq $mpObj -or [string]::IsNullOrWhiteSpace([string]$mpObj)) {
            throw "El territorio $numeroTxt no tiene manzanas pendientes registradas."
        }

        $cmd = $con.CreateCommand()
        $cmd.CommandText = "UPDATE NumeroDeTerritorios SET ManzanasPendientes=? WHERE CStr(NumeroTerritorio)=?"
        [void]$cmd.Parameters.AddWithValue('@mp', $manzanasTexto)
        [void]$cmd.Parameters.AddWithValue('@n', $numeroTxt)
        [void]$cmd.ExecuteNonQuery()

        return [ordered]@{
            ok = $true
            numero = $numeroTxt
            manzanasPendientes = $manzanasTexto
        }
    }
    finally { $con.Close() }
}

function Execute-AgregadoAccion {
    param($body)

    $numeroTxt = ([string]$body.numero).Trim()
    $accion = ([string]$body.accion).Trim().ToLowerInvariant()
    $responsableTxt = ([string]$body.responsable).Trim()
    $manzanasTexto = ([string]$body.manzanasPendientes).Trim()

    if ([string]::IsNullOrWhiteSpace($numeroTxt)) { throw "Numero de territorio requerido." }
    if ([string]::IsNullOrWhiteSpace($accion)) { throw "Accion requerida." }
    if ([string]::IsNullOrWhiteSpace($responsableTxt)) { throw "Responsable requerido." }

    $fechaRef = $null
    try {
        $fechaRaw = [string]$body.fecha
        if ([string]::IsNullOrWhiteSpace($fechaRaw)) {
            $fechaRef = [datetime]::Today
        } else {
            $fechaRef = [datetime]::Parse($fechaRaw)
        }
    } catch {
        throw "Fecha invalida."
    }

    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $cmdChk = $con.CreateCommand()
        $cmdChk.CommandText = "SELECT TOP 1 NumeroTerritorio FROM NumeroDeTerritorios WHERE CStr(NumeroTerritorio)=?"
        [void]$cmdChk.Parameters.AddWithValue('@n', $numeroTxt)
        $rd = $cmdChk.ExecuteReader()
        if (-not $rd.Read()) {
            $rd.Close()
            throw "No se encontro el territorio $numeroTxt."
        }
        $numero = [string]$rd['NumeroTerritorio']
        $rd.Close()

        switch ($accion) {
            'completado' {
                $cmdIns = $con.CreateCommand()
                $cmdIns.CommandText = "INSERT INTO Historial (NumeroTerritorio, Terminado, Encargado, Iniciado) VALUES (?, ?, ?, ?)"
                [void]$cmdIns.Parameters.AddWithValue('@n', $numero)
                $pT = $cmdIns.Parameters.Add('@t', [System.Data.OleDb.OleDbType]::Date); $pT.Value = $fechaRef
                [void]$cmdIns.Parameters.AddWithValue('@e', $responsableTxt)
                $pI = $cmdIns.Parameters.Add('@i', [System.Data.OleDb.OleDbType]::Date); $pI.Value = $fechaRef
                [void]$cmdIns.ExecuteNonQuery()
                return [ordered]@{ ok = $true; numero = $numero; accion = 'completado' }
            }
            'completado-campana' {
                $enie = [char]0x00F1
                $cmdUp = $con.CreateCommand()
                $cmdUp.CommandText = "UPDATE NumeroDeTerritorios SET AbiertoCerrado=False, [Campa${enie}a]=True, Enviado=Null, PuntosDeEncuentros=Null, Grupos=Null, Responsables=Null, Asignado=Null, ManzanasPendientes=Null, FechaManzanasPendientes=Null, UltimoEncargado=Null WHERE CStr(NumeroTerritorio)=?"
                [void]$cmdUp.Parameters.AddWithValue('@n', $numeroTxt)
                [void]$cmdUp.ExecuteNonQuery()

                $cmdIns = $con.CreateCommand()
                $cmdIns.CommandText = "INSERT INTO Historial (NumeroTerritorio, Terminado, Encargado, Iniciado) VALUES (?, ?, ?, ?)"
                [void]$cmdIns.Parameters.AddWithValue('@n', $numero)
                $pT = $cmdIns.Parameters.Add('@t', [System.Data.OleDb.OleDbType]::Date); $pT.Value = $fechaRef
                [void]$cmdIns.Parameters.AddWithValue('@e', $responsableTxt)
                $pI = $cmdIns.Parameters.Add('@i', [System.Data.OleDb.OleDbType]::Date); $pI.Value = $fechaRef
                [void]$cmdIns.ExecuteNonQuery()
                return [ordered]@{ ok = $true; numero = $numero; accion = 'completado-campana' }
            }
            'incompleto' {
                if ([string]::IsNullOrWhiteSpace($manzanasTexto)) {
                    throw "Para incompleto debe ingresar Manzanas Pendientes."
                }
                $cmdUp = $con.CreateCommand()
                $cmdUp.CommandText = "UPDATE NumeroDeTerritorios SET AbiertoCerrado=False, Enviado=Null, PuntosDeEncuentros=Null, Grupos=Null, Responsables=Null, Prioridad=Null, Asignado=IIf(IsNull(Asignado), ?, Asignado), ManzanasPendientes=?, FechaManzanasPendientes=?, UltimoEncargado=? WHERE CStr(NumeroTerritorio)=?"
                $pA = $cmdUp.Parameters.Add('@a', [System.Data.OleDb.OleDbType]::Date); $pA.Value = $fechaRef
                [void]$cmdUp.Parameters.AddWithValue('@mp', $manzanasTexto)
                $pF = $cmdUp.Parameters.Add('@f', [System.Data.OleDb.OleDbType]::Date); $pF.Value = $fechaRef
                [void]$cmdUp.Parameters.AddWithValue('@u', $responsableTxt)
                [void]$cmdUp.Parameters.AddWithValue('@n', $numeroTxt)
                [void]$cmdUp.ExecuteNonQuery()
                return [ordered]@{ ok = $true; numero = $numero; accion = 'incompleto' }
            }
            default {
                throw "Accion no soportada: $accion"
            }
        }
    }
    finally { $con.Close() }
}

function Import-CierresColaborador {
    param($body)

    $lista = $body.acciones
    if ($null -eq $lista) {
        throw "Se requiere acciones (array)."
    }
    $detalle = New-Object System.Collections.ArrayList
    $n = 0
    foreach ($item in @($lista)) {
        $n++
        $tipo = [string]$item.tipo
        $payload = $item.payload
        if ([string]::IsNullOrWhiteSpace($tipo)) {
            throw "La accion $n no tiene tipo."
        }
        try {
            $res = $null
            switch ($tipo) {
                'cerrar-sin-hecho' { $res = Close-TerritorioSinHecho $payload }
                'cerrar-con-accion' { $res = Close-TerritorioConAccion $payload }
                'guardar-edicion-salida' { $res = Save-TerritorioEdicionSalida $payload }
                'agregado-accion' { $res = Execute-AgregadoAccion $payload }
                default { throw "Tipo no soportado: $tipo" }
            }
            [void]$detalle.Add([ordered]@{ i = $n; ok = $true; tipo = $tipo })
        }
        catch {
            [void]$detalle.Add([ordered]@{ i = $n; ok = $false; tipo = $tipo; error = $_.Exception.Message })
            return [ordered]@{
                ok = $false
                message = $_.Exception.Message
                hasta = $n
                detalle = $detalle.ToArray()
            }
        }
    }
    return [ordered]@{
        ok = $true
        count = $detalle.Count
        detalle = $detalle.ToArray()
    }
}

function Clear-CampanaGlobal {
    $con = New-Object System.Data.OleDb.OleDbConnection $cadenaConexion
    $con.Open()
    try {
        $enie = [char]0x00F1
        $cmd = $con.CreateCommand()
        $cmd.CommandText = "UPDATE NumeroDeTerritorios SET [Campa${enie}a]=False"
        $rows = [int]$cmd.ExecuteNonQuery()
        return [ordered]@{
            ok = $true
            actualizados = $rows
        }
    }
    finally { $con.Close() }
}

function Read-JsonBody {
    param($req)
    $sr = New-Object System.IO.StreamReader($req.InputStream, $req.ContentEncoding)
    try {
        $txt = $sr.ReadToEnd()
        if ([string]::IsNullOrWhiteSpace($txt)) { return $null }
        return ($txt | ConvertFrom-Json)
    }
    finally { $sr.Close() }
}

function Write-Json {
    param($resp, $obj, [int]$status = 200)
    $resp.StatusCode = $status
    $resp.ContentType = 'application/json; charset=utf-8'
    $json = ConvertTo-Json5 $obj
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $resp.ContentLength64 = $bytes.Length
    $resp.OutputStream.Write($bytes, 0, $bytes.Length)
}

function Resolve-SafeStaticPath {
    param([string]$root, [string]$relativePath)

    $rootFull = [System.IO.Path]::GetFullPath($root)
    $candidate = [System.IO.Path]::GetFullPath((Join-Path $rootFull $relativePath))
    if ($candidate.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $candidate
    }
    return $null
}

# ----------------------------------------------------------------
# Estado portable (JSON en la carpeta del programa; viaja al copiar la carpeta)
# Referencias del mapa + calibracion SVG (mismo criterio que localStorage previo)
# ----------------------------------------------------------------
$PortableUiPath = Get-AppPath 'territorios_portable.json'

function Get-PortableUiDefault {
    return [ordered]@{
        version      = 1
        ajusteGlobal = $null
        referencias  = @()
    }
}

function Read-PortableUi {
    if (-not (Test-Path -LiteralPath $PortableUiPath)) {
        return (Get-PortableUiDefault)
    }
    try {
        $enc = [System.Text.UTF8Encoding]::new($false)
        $txt = [System.IO.File]::ReadAllText($PortableUiPath, $enc)
        if ([string]::IsNullOrWhiteSpace($txt)) {
            return (Get-PortableUiDefault)
        }
        $o = $txt | ConvertFrom-Json
        $refs = @()
        if ($null -ne $o -and $null -ne $o.referencias) {
            $refs = @($o.referencias)
        }
        $aj = $null
        if ($null -ne $o -and $null -ne $o.ajusteGlobal) {
            $aj = $o.ajusteGlobal
        }
        return [ordered]@{
            version      = 1
            ajusteGlobal = $aj
            referencias  = $refs
        }
    }
    catch {
        return (Get-PortableUiDefault)
    }
}

function Save-PortableUi {
    param($doc)
    $toWrite = [ordered]@{
        version      = 1
        ajusteGlobal = $doc.ajusteGlobal
        referencias  = @($doc.referencias)
    }
    $json = ConvertTo-Json5 $toWrite
    $enc = [System.Text.UTF8Encoding]::new($false)
    $tmp = $PortableUiPath + '.tmp'
    [System.IO.File]::WriteAllText($tmp, $json, $enc)
    if (Test-Path -LiteralPath $PortableUiPath) {
        Remove-Item -LiteralPath $PortableUiPath -Force
    }
    Move-Item -LiteralPath $tmp -Destination $PortableUiPath -Force
}

# ----------------------------------------------------------------
# Servidor HTTP
# ----------------------------------------------------------------
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Puerto/")
try {
    $listener.Start()
} catch {
    Write-Host "ERROR: no se pudo abrir el puerto $Puerto." -ForegroundColor Red
    Write-Host "Si ya hay un servidor corriendo, cerralo o cambia el puerto con: servir.ps1 -Puerto 8081" -ForegroundColor Yellow
    Read-Host "Presiona Enter para salir"
    exit 1
}

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Territorios - servidor local" -ForegroundColor Cyan
Write-Host "  URL:     http://localhost:$Puerto/" -ForegroundColor Cyan
Write-Host "  Carpeta: $raiz" -ForegroundColor Gray
Write-Host "  Access:  $Accdb" -ForegroundColor Gray
Write-Host "  Existe:  $(Test-Path $Accdb)" -ForegroundColor Gray
Write-Host "  CTRL+C para detener" -ForegroundColor Gray
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

Start-Process "http://localhost:$Puerto/mapa.html"
$script:StopRequested = $false

try {
    while ($listener.IsListening -and -not $script:StopRequested) {
        $ctx  = $listener.GetContext()
        $req  = $ctx.Request
        $resp = $ctx.Response
        # Normalizar mayúsculas (AbsolutePath puede llegar como /API/... y los elseif son sensibles al caso)
        $ruta = ([Uri]::UnescapeDataString($req.Url.AbsolutePath)).TrimEnd('/').ToLowerInvariant()

        Write-Host "$($req.HttpMethod) $ruta" -ForegroundColor DarkGray

        try {
            # ------------ API ------------
            if ($ruta -eq '/api/ping') {
                Write-Json $resp ([ordered]@{
                    ok     = $true
                    accdb  = $Accdb
                    existe = (Test-Path $Accdb)
                    driver = 'Microsoft.ACE.OLEDB.12.0'
                    ps     = "$($PSVersionTable.PSVersion)"
                    bitness = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
                })
            }
            elseif ($ruta -eq '/api/portable-ui' -and $req.HttpMethod -eq 'GET') {
                $datos = Read-PortableUi
                Write-Json $resp $datos
            }
            elseif ($ruta -eq '/api/portable-ui' -and $req.HttpMethod -eq 'PUT') {
                $body = Read-JsonBody $req
                if ($null -eq $body) { $body = @{} }
                $existing = Read-PortableUi
                $norm = [ordered]@{
                    version      = 1
                    ajusteGlobal = $existing.ajusteGlobal
                    referencias  = @($existing.referencias)
                }
                if ($null -ne $body.ajusteGlobal) {
                    $norm.ajusteGlobal = $body.ajusteGlobal
                }
                if ($null -ne $body.referencias) {
                    $norm.referencias = @($body.referencias)
                }
                Save-PortableUi $norm
                Write-Json $resp $norm
            }
            elseif ($ruta -eq '/api/territorios') {
                $datos = Read-Territorios
                Write-Json $resp $datos
            }
            elseif ($ruta -eq '/api/salidas') {
                $datos = Read-SalidasPredicacion
                Write-Json $resp $datos
            }
            elseif ($ruta -eq '/api/no-visitar' -and $req.HttpMethod -eq 'GET') {
                $datos = Read-NoVisitarMap
                if ($null -eq $datos) { $datos = @() }
                Write-Json $resp $datos
            }
            elseif ($ruta -match '^/api/no-visitar/([^/]+)$' -and $req.HttpMethod -eq 'GET') {
                $n = $Matches[1]
                $datos = Read-NoVisitarByNumero $n
                Write-Json $resp $datos
            }
            elseif ($ruta -eq '/api/no-visitar' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Add-NoVisitar ([string]$body.numero) ([string]$body.direccion)
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/no-visitar' -and $req.HttpMethod -eq 'DELETE') {
                $body = Read-JsonBody $req
                $res = Delete-NoVisitar ([string]$body.numero) ([string]$body.direccion)
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/form-data') {
                $datos = Read-FormData
                Write-Json $resp $datos
            }
            elseif ($ruta -eq '/api/predeterminados' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                Set-PredeterminadosConfig $body
                $datos = Read-FormData
                Write-Json $resp $datos
            }
            elseif ($ruta -match '^/api/catalogo/([a-zA-Z]+)$' -and $req.HttpMethod -eq 'POST') {
                $tipo = $Matches[1]
                $body = Read-JsonBody $req
                Add-CatalogoItem $tipo $body
                $datos = Read-FormData
                Write-Json $resp $datos
            }
            elseif ($ruta -match '^/api/catalogo/([a-zA-Z]+)$' -and $req.HttpMethod -eq 'DELETE') {
                $tipo = $Matches[1]
                $body = Read-JsonBody $req
                Delete-CatalogoItem $tipo $body
                $datos = Read-FormData
                Write-Json $resp $datos
            }
            elseif ($ruta -match '^/api/historial/(.+)$') {
                $n = $Matches[1]
                $datos = Read-Historial $n
                Write-Json $resp $datos
            }
            elseif ($ruta -match '^/api/esquinas/(.+)$' -and $req.HttpMethod -eq 'GET') {
                $n = [Uri]::UnescapeDataString($Matches[1]).Trim()
                $datos = Read-Esquinas $n
                Write-Json $resp $datos
            }
            elseif ($ruta -eq '/api/esquinas' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $nTerr = Normalize-NumeroTerritorio $body.numero
                Add-Esquina $nTerr ([int]$body.id1) ([int]$body.id2)
                $datos = Read-Esquinas $nTerr
                Write-Json $resp $datos
            }
            elseif ($ruta -eq '/api/esquinas' -and $req.HttpMethod -eq 'DELETE') {
                $body = Read-JsonBody $req
                $nTerr = Normalize-NumeroTerritorio $body.numero
                Remove-Esquina $nTerr ([int]$body.id1) ([int]$body.id2)
                $datos = Read-Esquinas $nTerr
                Write-Json $resp $datos
            }
            elseif ($ruta -eq '/api/existentes') {
                $punto = [string]$req.QueryString['punto']
                $fechaRaw = [string]$req.QueryString['fecha']
                $grupo = [string]$req.QueryString['grupo']
                $fecha = [datetime]::Parse($fechaRaw)
                $datos = Read-Existentes $punto $fecha $grupo
                Write-Json $resp $datos
            }
            elseif ($ruta -eq '/api/asignar' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Set-Asignacion $body
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/territorios/guardar-edicion-salida' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Save-TerritorioEdicionSalida $body
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/territorios/guardar-grupo' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Save-TerritorioGrupo $body
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/territorios/cerrar-sin-hecho' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Close-TerritorioSinHecho $body
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/territorios/cerrar-con-accion' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Close-TerritorioConAccion $body
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/territorios/borrar-incompleto' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Remove-TerritorioIncompleto $body
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/territorios/actualizar-manzanas-pendientes' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Update-TerritorioManzanasPendientes $body
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/territorios/set-rural' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Set-TerritorioRural $body
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/territorios/agregado-accion' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Execute-AgregadoAccion $body
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/recibir-datos-semana-bundle' -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Sync-SalidasBundleToAccess $body
                $code = if ($res.ok) { 200 } else { 200 }
                Write-Json $resp $res $code
            }
            elseif (($ruta -eq '/api/territorios/importar-cierres-colaborador' -or $ruta -eq '/api/importar-cierres-colaborador') -and $req.HttpMethod -eq 'POST') {
                $body = Read-JsonBody $req
                $res = Import-CierresColaborador $body
                $code = if ($res.ok) { 200 } else { 400 }
                Write-Json $resp $res $code
            }
            elseif ($ruta -eq '/api/campana/limpiar' -and $req.HttpMethod -eq 'POST') {
                $res = Clear-CampanaGlobal
                Write-Json $resp $res
            }
            elseif ($ruta -eq '/api/stop') {
                Write-Json $resp ([ordered]@{
                    ok = $true
                    message = 'Servidor detenido por solicitud del usuario.'
                })
                $script:StopRequested = $true
            }
            # ------------ Estaticos ------------
            else {
                $rel = $ruta.TrimStart('/')
                if ([string]::IsNullOrWhiteSpace($rel)) { $rel = 'mapa.html' }
                $archivo = Resolve-SafeStaticPath $raiz $rel

                if ($null -eq $archivo) {
                    $resp.StatusCode = 403
                    $msg = [System.Text.Encoding]::UTF8.GetBytes("403 - Ruta no permitida")
                    $resp.OutputStream.Write($msg, 0, $msg.Length)
                }
                elseif ((Test-Path $archivo) -and (-not (Get-Item $archivo).PSIsContainer)) {
                    $ext = [System.IO.Path]::GetExtension($archivo).ToLower()
                    $ct  = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { 'application/octet-stream' }
                    $bytes = [System.IO.File]::ReadAllBytes($archivo)
                    $resp.ContentType = $ct
                    $resp.ContentLength64 = $bytes.Length
                    $resp.OutputStream.Write($bytes, 0, $bytes.Length)
                } else {
                    $resp.StatusCode = 404
                    $msg = [System.Text.Encoding]::UTF8.GetBytes("404 - $rel")
                    $resp.OutputStream.Write($msg, 0, $msg.Length)
                }
            }
        }
        catch {
            Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
            try {
                Write-Json $resp ([ordered]@{
                    error = $_.Exception.Message
                    tipo  = $_.Exception.GetType().FullName
                }) 500
            } catch {}
        }
        finally {
            try { $resp.OutputStream.Close() } catch {}
        }
    }
}
finally {
    $listener.Stop()
    $listener.Close()
}
