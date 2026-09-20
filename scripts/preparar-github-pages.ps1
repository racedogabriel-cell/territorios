# Prepares docs/ for GitHub Pages from the Netlify public pages.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path (Join-Path $root 'Netlify\predicacion-netlify.html'))) {
    $root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
}
$docs = Join-Path $root 'docs'
New-Item -ItemType Directory -Force -Path $docs | Out-Null

$utf8 = New-Object System.Text.UTF8Encoding $false

function Write-Utf8([string]$path, [string]$text) {
    [System.IO.File]::WriteAllText($path, $text, $utf8)
}

function Read-Utf8([string]$path) {
    return [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
}

function Convert-NetlifyToGoogle([string]$text) {
    $text = $text.Replace('./supabase.js', './google.js')
    $text = $text.Replace('Netlify/supabase.js', 'google/google.js')
    $text = $text.Replace('Sync-SupabaseToAccess.ps1', 'Sync-GoogleToAccess.ps1')
    $text = $text.Replace('SUPABASE.md', 'GITHUB_PAGES.md')
    $text = $text.Replace('sql/create_tables.sql', 'GITHUB_PAGES.md')
    $text = $text.Replace("msg.includes('SUPABASE_SIN_DATOS') ||", "msg.includes('SUPABASE_SIN_DATOS') || msg.includes('GOOGLE_SIN_DATOS') || msg.includes('GOOGLE_CONFIG') ||")
    $text = [regex]::Replace($text, '(?<![A-Za-z])Supabase(?![A-Za-z])', 'Google Sheets')
    $text = $text.Replace('Enviar a Google Sheets', 'Enviar a Google')
    $text = $text.Replace('Recibir de Google Sheets', 'Recibir de Google')
    $text = $text.Replace('modulo Google Sheets', 'modulo Google')
    $text = $text.Replace('dulo Google Sheets', 'dulo Google')
    $text = [regex]::Replace($text, '\(URL y ANON KEY\)[^.]*\.', '(APPS_SCRIPT_URL o SPREADSHEET_ID + API_KEY).')
    $text = [regex]::Replace($text, 'SQL Editor[^<]*<code>GITHUB_PAGES.md</code>', 'Apps Script (ver <code>GITHUB_PAGES.md</code>)')
    return $text
}

$pred = Convert-NetlifyToGoogle (Read-Utf8 (Join-Path $root 'Netlify\predicacion-netlify.html'))
$mapa = Convert-NetlifyToGoogle (Read-Utf8 (Join-Path $root 'Netlify\mapa-semanal.html'))

$pred = $pred.Replace(
    '<button type="button" class="btn btn-secondary" id="btn-agregados">Agregados</button>',
    '<button type="button" class="btn btn-secondary" id="btn-agregados">Agregados</button>' + "`r`n            <a class=`"btn btn-secondary`" href=`"mapa-semanal.html`" style=`"display:inline-flex;align-items:center;text-decoration:none;`">Mapa semanal</a>"
)
$mapa = $mapa.Replace(
    '<h1>Mapa semanal de territorios</h1>',
    '<h1>Mapa semanal de territorios</h1>' + "`r`n  <a href=`"./`" class=`"btn-top`" style=`"text-decoration:none`">Salidas</a>"
)

Write-Utf8 (Join-Path $docs 'index.html') $pred
Write-Utf8 (Join-Path $docs 'mapa-semanal.html') $mapa
Copy-Item -Force (Join-Path $root 'Netlify\mapa.svg') (Join-Path $docs 'mapa.svg')
Copy-Item -Force (Join-Path $root 'Netlify\mapa_meta.json') (Join-Path $docs 'mapa_meta.json')
Copy-Item -Force (Join-Path $root 'google\google.js') (Join-Path $docs 'google.js')
Write-Utf8 (Join-Path $docs '.nojekyll') "`n"

Write-Host "GitHub Pages listo en $docs"
