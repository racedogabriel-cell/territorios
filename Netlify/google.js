(function () {
    'use strict';

    // =============================================================================
    // CONFIGURACIÓN GOOGLE — editar estas líneas
    //
    // Opción A (recomendada, lectura + escritura):
    //   1. Crear una Hoja de cálculo de Google.
    //   2. Extensiones → Apps Script → pegar google/Codigo.gs
    //   3. Ejecutar inicializar() una vez.
    //   4. Implementar → Nueva implementación → Aplicación web
    //      - Ejecutar como: Yo
    //      - Quién tiene acceso: Cualquiera
    //   5. Pegar la URL de la aplicación web en APPS_SCRIPT_URL.
    //
    // Opción B (solo lectura, API de Sheets):
    //   - Habilitar Google Sheets API en Google Cloud.
    //   - Crear una clave de API.
    //   - Compartir la hoja: "Cualquiera con el enlace" → Lector.
    //   - Pegar SPREADSHEET_ID y API_KEY.
    // =============================================================================
    const GOOGLE_CONFIG = {
        SPREADSHEET_ID: 'TU_SPREADSHEET_ID',
        API_KEY: 'TU_API_KEY',
        APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbxIjJCUJINarSxaecBnFZlSQiexoSAgHKxGLj-65-ratqMo5WFScvqQz7ZAQpb1rHHe/exec'
    };
    // =============================================================================

    const SHEETS = {
        DATOS: 'datos_semana',
        SALIDAS: 'salidas',
        MAPA: 'mapa_territorios',
        ACCIONES: 'acciones'
    };
    const CANAL_DATOS_PUBLICOS = 'publica';
    const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets/';

    function cfg() {
        return {
            spreadsheetId: String(GOOGLE_CONFIG.SPREADSHEET_ID || '').trim(),
            apiKey: String(GOOGLE_CONFIG.API_KEY || '').trim(),
            scriptUrl: String(GOOGLE_CONFIG.APPS_SCRIPT_URL || '').trim()
        };
    }

    function esPlaceholder(valor, token) {
        const v = String(valor || '').trim();
        return !v || v === token || v.indexOf('TU_') === 0;
    }

    function puedeAppsScript() {
        return !esPlaceholder(cfg().scriptUrl, 'TU_APPS_SCRIPT_URL');
    }

    function puedeSheetsApi() {
        const c = cfg();
        return !esPlaceholder(c.spreadsheetId, 'TU_SPREADSHEET_ID') &&
            !esPlaceholder(c.apiKey, 'TU_API_KEY');
    }

    function assertAlgunaConfig() {
        if (!puedeAppsScript() && !puedeSheetsApi()) {
            throw new Error(
                'GOOGLE_CONFIG: Configurá APPS_SCRIPT_URL (recomendado) o SPREADSHEET_ID + API_KEY en google/google.js. Ver GITHUB_PAGES.md.'
            );
        }
    }

    function assertEscritura() {
        if (!puedeAppsScript()) {
            throw new Error(
                'GOOGLE_CONFIG: Para guardar acciones o publicar datos hace falta APPS_SCRIPT_URL (aplicación web de Apps Script). Ver GITHUB_PAGES.md.'
            );
        }
    }

    function parseJsonMaybe(val, fallback) {
        if (val == null || val === '') return fallback;
        if (typeof val === 'object') return val;
        const txt = String(val).trim();
        if (!txt) return fallback;
        try { return JSON.parse(txt); }
        catch (_) { return fallback; }
    }

    function truthySheet(val) {
        const t = String(val == null ? '' : val).trim().toLowerCase();
        return t === 'true' || t === '1' || t === 'verdadero' || t === 'si' || t === 'sí';
    }

    function mapearSalida(row) {
        if (Array.isArray(row)) {
            return {
                NumeroTerritorio: row[0] ?? '',
                PuntosDeEncuentros: row[1] ?? '',
                Responsables: row[2] ?? '',
                Grupos: row[3] ?? '',
                Enviado: row[4] ?? '',
                Prioridad: row[5] === '' || row[5] == null ? null : row[5],
                ManzanasPendientes: row[6] ?? '',
                Direccion: row[7] ?? '',
                Descripcion: row[8] ?? ''
            };
        }
        return {
            NumeroTerritorio: row.NumeroTerritorio ?? row.numero_territorio ?? row.numeroTerritorio ?? row.numero ?? '',
            PuntosDeEncuentros: row.PuntosDeEncuentros ?? row.puntos_de_encuentros ?? row.punto_encuentro ?? row.punto ?? '',
            Responsables: row.Responsables ?? row.responsables ?? row.responsable ?? '',
            Grupos: row.Grupos ?? row.grupos ?? row.grupo ?? '',
            Enviado: row.Enviado ?? row.enviado ?? row.fecha_hora ?? row.fecha ?? '',
            Prioridad: row.Prioridad ?? row.prioridad ?? null,
            ManzanasPendientes: row.ManzanasPendientes ?? row.manzanas_pendientes ?? '',
            Direccion: row.Direccion ?? row.direccion ?? '',
            Descripcion: row.Descripcion ?? row.descripcion ?? ''
        };
    }

    function indiceColumnas(header) {
        const map = {};
        (header || []).forEach((name, i) => {
            map[String(name || '').trim().toLowerCase()] = i;
        });
        return map;
    }

    function celda(row, idx, nombres, def) {
        for (const n of nombres) {
            if (idx[n] == null) continue;
            const v = row[idx[n]];
            if (v != null && String(v).trim() !== '') return v;
        }
        return def;
    }

    async function sheetsValuesGet(range) {
        const c = cfg();
        const url = SHEETS_API + encodeURIComponent(c.spreadsheetId) +
            '/values/' + encodeURIComponent(range) +
            '?key=' + encodeURIComponent(c.apiKey);
        const resp = await fetch(url, { cache: 'no-store' });
        const body = await resp.json().catch(() => ({}));
        if (!resp.ok) {
            const msg = body && (body.error && body.error.message) ? body.error.message : ('HTTP ' + resp.status);
            throw new Error('Google Sheets API: ' + msg);
        }
        return Array.isArray(body.values) ? body.values : [];
    }

    function unwrapScriptData(body) {
        if (body == null) return body;
        if (typeof body === 'object' && Object.prototype.hasOwnProperty.call(body, 'ok')) {
            if (!body.ok) throw new Error(body.error || 'Error en Apps Script.');
            return body.data;
        }
        return body;
    }

    async function callAppsScript(action, payload, method) {
        assertEscritura();
        const c = cfg();
        const base = c.scriptUrl.replace(/\/+$/, '');
        const verbo = (method || (payload ? 'POST' : 'GET')).toUpperCase();
        let resp;
        if (verbo === 'GET') {
            const q = new URLSearchParams({ action: action });
            if (payload && typeof payload === 'object') {
                Object.keys(payload).forEach((k) => {
                    if (payload[k] == null) return;
                    q.set(k, typeof payload[k] === 'object' ? JSON.stringify(payload[k]) : String(payload[k]));
                });
            }
            resp = await fetch(base + '?' + q.toString(), { method: 'GET', cache: 'no-store', redirect: 'follow' });
        } else {
            const body = Object.assign({ action: action }, payload || {});
            resp = await fetch(base, {
                method: 'POST',
                cache: 'no-store',
                redirect: 'follow',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify(body)
            });
        }
        const text = await resp.text();
        let parsed = null;
        if (text) {
            try { parsed = JSON.parse(text); }
            catch (_) { parsed = text; }
        }
        if (!resp.ok) {
            const msg = parsed && typeof parsed === 'object'
                ? (parsed.error || parsed.message || JSON.stringify(parsed))
                : (parsed || ('HTTP ' + resp.status));
            throw new Error(String(msg));
        }
        return unwrapScriptData(parsed);
    }

    async function cargarDatosSemanaDesdeSheets() {
        const metaRows = await sheetsValuesGet(SHEETS.DATOS + '!A1:H20');
        if (!metaRows.length) {
            throw new Error('GOOGLE_SIN_DATOS: La hoja datos_semana está vacía. Publicá desde mapa.html (Enviar a Google).');
        }
        const idx = indiceColumnas(metaRows[0]);
        const dataRows = metaRows.slice(1).filter((r) => (r || []).some((c) => String(c || '').trim()));
        if (!dataRows.length) {
            throw new Error('GOOGLE_SIN_DATOS: La hoja datos_semana no tiene filas. Publicá desde mapa.html (Enviar a Google).');
        }
        let chosen = dataRows[0];
        for (const row of dataRows) {
            const canal = String(celda(row, idx, ['canal'], '') || '').trim().toLowerCase();
            if (canal === CANAL_DATOS_PUBLICOS) {
                chosen = row;
                break;
            }
        }

        const blob = celda(chosen, idx, ['datos', 'data', 'bundle', 'json'], null);
        const parsedBlob = parseJsonMaybe(blob, null);
        if (parsedBlob && (Array.isArray(parsedBlob.salidas) || parsedBlob.mapaSemanal)) {
            return parsedBlob;
        }

        const exportado = celda(chosen, idx, ['exportado'], new Date().toISOString());
        const version = Number(celda(chosen, idx, ['version'], 2)) || 2;
        const responsables = parseJsonMaybe(celda(chosen, idx, ['responsables', 'responsables_json'], '[]'), []);
        const territoriosCat = parseJsonMaybe(celda(chosen, idx, ['territorios', 'territorios_json'], '[]'), []);
        const descripciones = parseJsonMaybe(celda(chosen, idx, ['descripciones', 'descripciones_json'], '{}'), {});
        const mapaExtra = parseJsonMaybe(celda(chosen, idx, ['mapa_extra', 'mapa_extra_json'], '{}'), {});

        const salidasRows = await sheetsValuesGet(SHEETS.SALIDAS + '!A1:I5000');
        const salidas = [];
        if (salidasRows.length > 1) {
            const sIdx = indiceColumnas(salidasRows[0]);
            const looksHeader = String(salidasRows[0][0] || '').toLowerCase().indexOf('numero') >= 0 ||
                sIdx.numeroterritorio != null;
            const start = looksHeader ? 1 : 0;
            for (let i = start; i < salidasRows.length; i++) {
                const r = salidasRows[i] || [];
                if (!r.some((c) => String(c || '').trim())) continue;
                if (looksHeader) {
                    salidas.push({
                        NumeroTerritorio: celda(r, sIdx, ['numeroterritorio', 'numero', 'territorio'], ''),
                        PuntosDeEncuentros: celda(r, sIdx, ['puntosdeencuentros', 'punto', 'puntoencuentro'], ''),
                        Responsables: celda(r, sIdx, ['responsables', 'responsable'], ''),
                        Grupos: celda(r, sIdx, ['grupos', 'grupo'], ''),
                        Enviado: celda(r, sIdx, ['enviado', 'fecha', 'fechahora'], ''),
                        Prioridad: celda(r, sIdx, ['prioridad'], null),
                        ManzanasPendientes: celda(r, sIdx, ['manzanaspendientes', 'manzanas'], ''),
                        Direccion: celda(r, sIdx, ['direccion'], ''),
                        Descripcion: celda(r, sIdx, ['descripcion'], '')
                    });
                } else {
                    salidas.push(mapearSalida(r));
                }
            }
        }

        const mapaRows = await sheetsValuesGet(SHEETS.MAPA + '!A1:B5000');
        const mapaTerritorios = {};
        if (mapaRows.length) {
            const start = String(mapaRows[0][0] || '').toLowerCase() === 'numero' ? 1 : 0;
            for (let i = start; i < mapaRows.length; i++) {
                const r = mapaRows[i] || [];
                const num = String(r[0] || '').trim();
                if (!num) continue;
                mapaTerritorios[num] = parseJsonMaybe(r[1], r[1] || {});
            }
        }

        const bundle = {
            exportado,
            version,
            salidas,
            responsables: Array.isArray(responsables) ? responsables : [],
            territorios: Array.isArray(territoriosCat) ? territoriosCat : [],
            descripciones: descripciones && typeof descripciones === 'object' ? descripciones : {}
        };
        if (mapaExtra && typeof mapaExtra === 'object') {
            bundle.mapaSemanal = Object.assign({}, mapaExtra, {
                territorios: mapaTerritorios
            });
            if (!bundle.mapaSemanal.exportado) bundle.mapaSemanal.exportado = exportado;
        } else if (Object.keys(mapaTerritorios).length) {
            bundle.mapaSemanal = {
                version: 1,
                exportado,
                territorios: mapaTerritorios
            };
        }
        return bundle;
    }

    async function listarAccionesDesdeSheets() {
        const rows = await sheetsValuesGet(SHEETS.ACCIONES + '!A1:G5000');
        if (rows.length < 2) return [];
        const idx = indiceColumnas(rows[0]);
        const out = [];
        for (let i = 1; i < rows.length; i++) {
            const r = rows[i] || [];
            if (truthySheet(celda(r, idx, ['procesado'], 'false'))) continue;
            out.push({
                id: celda(r, idx, ['id'], i),
                tipo: celda(r, idx, ['tipo'], ''),
                payload: parseJsonMaybe(celda(r, idx, ['payload', 'payload_json'], '{}'), {}),
                origen: celda(r, idx, ['origen'], 'predicacion-github'),
                creado_en: celda(r, idx, ['creado_en', 'creadoen'], '')
            });
        }
        out.sort((a, b) => String(a.creado_en).localeCompare(String(b.creado_en)));
        return out;
    }

    async function cargarDatosSemana() {
        assertAlgunaConfig();
        if (puedeAppsScript()) {
            try {
                const data = await callAppsScript('cargarDatosSemana', null, 'GET');
                if (data) return data;
            } catch (err) {
                if (!puedeSheetsApi()) throw err;
            }
        }
        return cargarDatosSemanaDesdeSheets();
    }

    function normalizarPayloadAccion(payload) {
        const p = payload && typeof payload === 'object' ? Object.assign({}, payload) : {};
        if (p.numero != null && String(p.numero).trim()) {
            p.numero = String(p.numero).trim();
        }
        return p;
    }

    async function guardarAccion(tipo, payload) {
        const row = await callAppsScript('guardarAccion', {
            tipo: tipo,
            payload: normalizarPayloadAccion(payload)
        }, 'POST');
        return row;
    }

    async function listarAccionesPendientes() {
        assertAlgunaConfig();
        if (puedeAppsScript()) {
            try {
                const rows = await callAppsScript('listarAccionesPendientes', null, 'GET');
                return Array.isArray(rows) ? rows : [];
            } catch (err) {
                if (!puedeSheetsApi()) throw err;
            }
        }
        return listarAccionesDesdeSheets();
    }

    async function actualizarAccion(id, tipo, payload) {
        return callAppsScript('actualizarAccion', {
            id: id,
            tipo: tipo,
            payload: normalizarPayloadAccion(payload)
        }, 'POST');
    }

    async function eliminarAccion(id) {
        return callAppsScript('eliminarAccion', { id: id }, 'POST');
    }

    async function eliminarAccionesPendientesPorNumero(numero) {
        const num = String(numero || '').trim();
        if (!num) throw new Error('Número de territorio inválido.');
        return callAppsScript('eliminarAccionesPendientesPorNumero', { numero: num }, 'POST');
    }

    async function marcarAccionProcesada(id) {
        return callAppsScript('marcarAccionProcesada', { id: id }, 'POST');
    }

    async function upsertDatosSemana(bundle) {
        if (!bundle || typeof bundle !== 'object') {
            throw new Error('upsertDatosSemana: bundle inválido.');
        }
        return callAppsScript('upsertDatosSemana', { bundle: bundle }, 'POST');
    }

    const api = {
        url: cfg().scriptUrl || cfg().spreadsheetId,
        cargarDatosSemana,
        guardarAccion,
        listarAccionesPendientes,
        actualizarAccion,
        eliminarAccion,
        eliminarAccionesPendientesPorNumero,
        marcarAccionProcesada,
        upsertDatosSemana,
        CANAL_DATOS_PUBLICOS,
        config: GOOGLE_CONFIG
    };

    window.TerritoriosGoogle = api;
}());
