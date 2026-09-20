# Prepares docs/ for GitHub Pages from the public Netlify pages (already Google-backed).
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

$pred = Read-Utf8 (Join-Path $root 'Netlify\predicacion-netlify.html')
$mapa = Read-Utf8 (Join-Path $root 'Netlify\mapa-semanal.html')

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
Copy-Item -Force (Join-Path $root 'google\google.js') (Join-Path $root 'Netlify\google.js')
Write-Utf8 (Join-Path $docs '.nojekyll') "`n"

Write-Host "GitHub Pages listo en $docs"
