@echo off
setlocal
cd /d "%~dp0"

set "PS64=%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS64%" set "PS64=powershell"

rem Intento elegante: pedir apagado por API
"%PS64%" -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-RestMethod -Uri 'http://localhost:8080/api/stop' -TimeoutSec 2 | Out-Null } catch {}"

rem Fallback por puerto
"%PS64%" -NoProfile -ExecutionPolicy Bypass -Command "$p = Get-NetTCPConnection -LocalPort 8080 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique; foreach ($id in $p) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }"

echo Territorios cerrado.
