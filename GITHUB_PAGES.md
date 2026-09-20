# GitHub Pages + Google Sheets

La página pública de Netlify (`Netlify/predicacion-netlify.html` y `Netlify/mapa-semanal.html`) queda replicada en **GitHub Pages**. Los datos salen de una **Hoja de cálculo de Google** a través de una API (Apps Script o Google Sheets API).

## Qué se publica

| URL en GitHub Pages | Equivale a |
|---------------------|------------|
| `/` (`docs/index.html`) | Salidas / cierre (`predicacion-netlify.html`) |
| `/mapa-semanal.html` | Mapa semanal |

Carpeta de origen: `docs/`. Regenerarla con:

```powershell
.\scripts\preparar-github-pages.ps1
```

## 1. Crear la hoja de Google

1. Abrí [Google Sheets](https://sheets.google.com) y creá una hoja en blanco, por ejemplo **Territorios**.
2. **Extensiones → Apps Script**.
3. Borrá el código por defecto y pegá el contenido de `google/Codigo.gs`.
4. Guardá el proyecto.
5. En el selector de funciones elegí `inicializar` y pulsá **Ejecutar**. Autorizá el acceso a la hoja.
6. **Implementar → Nueva implementación → Aplicación web**:
   - Descripción: `api-territorios`
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquiera**
7. Copiá la URL (`https://script.google.com/macros/s/.../exec`).

## 2. Configurar el cliente

En `google/google.js` (y de nuevo en `docs/google.js` después de regenerar) completá:

```javascript
const GOOGLE_CONFIG = {
    SPREADSHEET_ID: 'TU_SPREADSHEET_ID',      // opcional si usás Apps Script
    API_KEY: 'TU_API_KEY',                    // opcional; solo lectura directa
    APPS_SCRIPT_URL: 'https://script.google.com/macros/s/XXXX/exec'
};
```

Con **solo** `APPS_SCRIPT_URL` alcanza para leer, publicar y guardar acciones.

### Lectura directa con Google Sheets API (opcional)

Si también querés que el navegador lea la hoja sin pasar por Apps Script:

1. [Google Cloud Console](https://console.cloud.google.com/) → nuevo proyecto (o uno existente).
2. Habilitá **Google Sheets API**.
3. Credenciales → **Clave de API**. Restringila a Sheets API y, si el sitio ya está en línea, al referrer `https://TU_USUARIO.github.io/*`.
4. El ID de la hoja está en la URL: `https://docs.google.com/spreadsheets/d/ESTE_ID/edit`.
5. Compartí la hoja: **Cualquiera con el enlace → Lector**.

Sin `APPS_SCRIPT_URL` la página **solo lee**. Guardar acciones y **Enviar a Google** requieren la aplicación web.

## 3. Publicar datos desde la PC

1. `Abrir Territorios.cmd` (servidor local).
2. En `mapa.html`: **Enviar a Google**.
3. En GitHub Pages, **Actualizar**.

## 4. Subir el sitio a GitHub Pages

No hace falta `gh` en esta PC. Desde la carpeta del programa:

```powershell
git init -b main
git add .
git commit -m "Sitio publico en GitHub Pages con datos de Google Sheets"
```

En GitHub.com: **New repository** (por ejemplo `territorios`), sin README. Después:

```powershell
git remote add origin https://github.com/TU_USUARIO/territorios.git
git push -u origin main
```

En el repositorio: **Settings → Pages**:

- Source: **GitHub Actions** (usa `.github/workflows/pages.yml`), o
- Source: **Deploy from a branch** → `main` → folder `/docs`.

La URL queda:

`https://racedogabriel-cell.github.io/territorios/`

## 5. Acciones de colaboradores → Access

```powershell
.\servir.ps1
.\Sync-GoogleToAccess.ps1
```

## Hojas que crea Apps Script

| Hoja | Contenido |
|------|-----------|
| `datos_semana` | Canal `publica`, catálogos y metadatos del mapa |
| `salidas` | Una fila por territorio asignado |
| `mapa_territorios` | Estado del mapa semanal (un territorio por fila) |
| `acciones` | Cola de cierres/ajustes pendientes |
