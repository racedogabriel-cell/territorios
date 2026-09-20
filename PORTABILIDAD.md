# Portabilidad — App Territorios

**Actualizado:** 2026-05-17 (post-limpieza)

La carpeta **`App Territorios`** es el programa completo. En `Programa Territorios` solo debe quedar esta carpeta (el resto obsoleto fue eliminado).

## Portabilidad

- Rutas desde **`PathManager.ps1`** y **`$PSScriptRoot`**.
- Sin rutas absolutas en código ejecutable.
- Copiar toda la carpeta y abrir **`Abrir Territorios.cmd`**.

## Archivos esenciales

| Grupo | Archivos |
|-------|----------|
| Arranque | `Abrir Territorios.cmd`, `Cerrar Territorios.cmd`, `abrir_mapa.vbs`, `servir.ps1`, `PathManager.ps1` |
| Mapa | `mapa.html`, `mapa.svg`, `mapa_meta.json`, `territorios_portable.json` |
| Salidas local | `predicacion.html` |
| Base | `Territorios_BaseDeDatos.accdb` |
| Supabase / web | `Netlify/predicacion-netlify.html`, `Netlify/mapa-semanal.html`, `Netlify/supabase.js`, `Netlify/mapa.svg`, `Netlify/mapa_meta.json`, `Sync-SupabaseToAccess.ps1`, `sql/*.sql` |
| GitHub Pages / Google | `docs/` (sitio público), `google/google.js`, `google/Codigo.gs`, `Sync-GoogleToAccess.ps1`, `GITHUB_PAGES.md` |
| Docs | `LEEME - Territorios.txt`, `SUPABASE.md`, `GITHUB_PAGES.md` |

## Dependencias externas

- Driver **Microsoft ACE.OLEDB.12.0**
- **PowerShell 5+**
- **Supabase** (opcional, flujo web Netlify)
- **Google Sheets / GitHub Pages** (opcional, flujo web; ver `GITHUB_PAGES.md`)
- Abrir siempre por HTTP con **`Abrir Territorios.cmd`** (no `file://`)

## Mover a otra PC

1. Copiar toda **`App Territorios`**.
2. Incluir **`Territorios_BaseDeDatos.accdb`**.
3. Ejecutar **`Abrir Territorios.cmd`**.
