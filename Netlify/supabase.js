(function () {
    'use strict';

    // =============================================================================
    // CONFIGURACIÓN SUPABASE — editar solo estas dos líneas
    // Panel: https://supabase.com/dashboard → tu proyecto → Settings → API
    //   - Project URL  → SUPABASE_URL
    //   - anon public  → SUPABASE_ANON_KEY
    // =============================================================================
    const SUPABASE_URL = 'https://paupgqeavgxameeqdkuv.supabase.co';
    const SUPABASE_ANON_KEY = 'sb_publishable_KtSh6YrHWSOFewrKNwPSTQ_gg_rq3_L';
    // =============================================================================

    const REST_URL = SUPABASE_URL.replace(/\/+$/, '') + '/rest/v1';

    function assertConfig() {
        const url = String(SUPABASE_URL || '').trim();
        const key = String(SUPABASE_ANON_KEY || '').trim();
        if (!url || url.indexOf('TU_PROYECTO') >= 0 || !/^https:\/\/.+\.supabase\.co/i.test(url)) {
            throw new Error('SUPABASE_CONFIG: Configure SUPABASE_URL en Netlify/supabase.js (Settings → API → Project URL).');
        }
        if (!key || key === 'TU_ANON_KEY' || key.length < 20) {
            throw new Error('SUPABASE_CONFIG: Configure SUPABASE_ANON_KEY en Netlify/supabase.js (Settings → API → anon public).');
        }
    }

    function headers(extra) {
        return Object.assign({
            apikey: SUPABASE_ANON_KEY,
            Authorization: 'Bearer ' + SUPABASE_ANON_KEY,
            'Content-Type': 'application/json;charset=UTF-8',
            Accept: 'application/json'
        }, extra || {});
    }

    async function request(path, options) {
        assertConfig();
        const resp = await fetch(REST_URL + path, Object.assign({ headers: headers() }, options || {}));
        const text = await resp.text();
        let body = null;
        if (text) {
            try { body = JSON.parse(text); }
            catch (_) { body = text; }
        }
        if (!resp.ok) {
            const msg = body && typeof body === 'object'
                ? (body.message || body.error || body.details || body.hint || JSON.stringify(body))
                : (body || ('HTTP ' + resp.status));
            throw new Error(msg);
        }
        return body;
    }

    function extraerCampoJson(row) {
        if (!row || typeof row !== 'object') return null;
        const keys = ['datos', 'data', 'payload', 'bundle', 'json', 'contenido'];
        for (const key of keys) {
            if (!Object.prototype.hasOwnProperty.call(row, key)) continue;
            const val = row[key];
            if (val && typeof val === 'object') return val;
            if (typeof val === 'string' && val.trim()) return JSON.parse(val);
        }
        return null;
    }

    function mapearSalida(row) {
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

    function normalizarDatosSemana(rows) {
        if (!Array.isArray(rows) || rows.length === 0) {
            throw new Error(
                'SUPABASE_SIN_DATOS: La tabla datos_semana no tiene filas. ' +
                'En Supabase → SQL Editor ejecutá sql/create_tables.sql y publicá datos desde mapa.html (Enviar a Supabase).'
            );
        }
        const firstJson = extraerCampoJson(rows[0]);
        if (firstJson) return firstJson;
        if (rows[0] && Array.isArray(rows[0].salidas)) return rows[0];
        return {
            exportado: new Date().toISOString(),
            version: 2,
            salidas: rows.map(mapearSalida),
            responsables: [],
            territorios: []
        };
    }

    async function cargarDatosSemana() {
        const queries = [
            '/datos_semana?select=*&order=exportado.desc&limit=1',
            '/datos_semana?select=*&order=created_at.desc&limit=1',
            '/datos_semana?select=*&order=creado_en.desc&limit=1',
            '/datos_semana?select=*'
        ];
        let lastError = null;
        for (const query of queries) {
            try { return normalizarDatosSemana(await request(query)); }
            catch (err) { lastError = err; }
        }
        throw lastError || new Error('No se pudo cargar datos_semana desde Supabase.');
    }

    function normalizarPayloadAccion(payload) {
        const p = payload && typeof payload === 'object' ? Object.assign({}, payload) : {};
        if (p.numero != null && String(p.numero).trim()) {
            p.numero = String(p.numero).trim();
        }
        return p;
    }

    function filasEliminadas(body) {
        if (Array.isArray(body)) return body;
        if (body && typeof body === 'object') return [body];
        return [];
    }

    async function guardarAccion(tipo, payload) {
        const body = await request('/acciones_colaboradores', {
            method: 'POST',
            headers: headers({ Prefer: 'return=representation' }),
            body: JSON.stringify({
                tipo,
                payload: normalizarPayloadAccion(payload),
                origen: 'predicacion-netlify',
                procesado: false,
                creado_en: new Date().toISOString()
            })
        });
        if (Array.isArray(body) && body[0] && typeof body[0] === 'object') return body[0];
        return body;
    }

    async function listarAccionesPendientes() {
        return request(
            '/acciones_colaboradores?procesado=eq.false&select=id,tipo,payload,creado_en,origen&order=creado_en.asc'
        );
    }

    async function actualizarAccion(id, tipo, payload) {
        const q = encodeURIComponent(String(id));
        return request('/acciones_colaboradores?id=eq.' + q, {
            method: 'PATCH',
            headers: headers({ Prefer: 'return=minimal' }),
            body: JSON.stringify({ tipo, payload: normalizarPayloadAccion(payload) })
        });
    }

    async function eliminarAccion(id) {
        const q = encodeURIComponent(String(id));
        const body = await request('/acciones_colaboradores?id=eq.' + q + '&procesado=eq.false', {
            method: 'DELETE',
            headers: headers({ Prefer: 'return=representation' })
        });
        const rows = filasEliminadas(body);
        if (!rows.length) {
            throw new Error(
                'No se eliminó ninguna fila (id=' + id + '). ' +
                'Ejecutá en Supabase SQL Editor: sql/alter_acciones_delete_anon.sql'
            );
        }
        return rows;
    }

    /** Borra todas las acciones no procesadas de un territorio (evita duplicados en cola). */
    async function eliminarAccionesPendientesPorNumero(numero) {
        const num = String(numero || '').trim();
        if (!num) throw new Error('Número de territorio inválido.');
        const numQ = encodeURIComponent(num);
        const path =
            '/acciones_colaboradores?procesado=eq.false&payload->>numero=eq.' + numQ;
        const body = await request(path, {
            method: 'DELETE',
            headers: headers({ Prefer: 'return=representation' })
        });
        const rows = filasEliminadas(body);
        if (!rows.length) {
            throw new Error(
                'No se eliminó ninguna acción pendiente para el territorio ' + num + '. ' +
                'Revisá la política RLS de DELETE (sql/alter_acciones_delete_anon.sql).'
            );
        }
        return rows;
    }

    /** Igual que Sync-SupabaseToAccess.ps1 tras import OK en Access. */
    async function marcarAccionProcesada(id) {
        const q = encodeURIComponent(String(id));
        return request('/acciones_colaboradores?id=eq.' + q, {
            method: 'PATCH',
            headers: headers({ Prefer: 'return=minimal' }),
            body: JSON.stringify({
                procesado: true,
                procesado_en: new Date().toISOString()
            })
        });
    }

        /** Paquete semanal: { exportado, version, salidas, responsables, territorios } */
    const CANAL_DATOS_PUBLICOS = 'publica';

    async function upsertDatosSemana(bundle) {
        if (!bundle || typeof bundle !== 'object') {
            throw new Error('upsertDatosSemana: bundle inválido.');
        }
        const exportadoIso = bundle.exportado || new Date().toISOString();
        const row = {
            canal: CANAL_DATOS_PUBLICOS,
            exportado: exportadoIso,
            datos: bundle
        };
        try {
            return await request('/datos_semana?on_conflict=canal', {
                method: 'POST',
                headers: headers({
                    Prefer: 'resolution=merge-duplicates,return=minimal'
                }),
                body: JSON.stringify([row])
            });
        } catch (err) {
            const msg = String(err && err.message ? err.message : err);
            if (/canal/i.test(msg) && (/schema cache/i.test(msg) || /could not find/i.test(msg))) {
                throw new Error(
                    'Tu proyecto Supabase aún no tiene la columna canal en datos_semana (o la API no refrescó el esquema). ' +
                    'Abrí Dashboard → SQL Editor, pegá y ejecutá todo el archivo sql/alter_datos_semana_canal_upsert.sql. ' +
                    'Si el error sigue, esperá 1–2 minutos y reintentá, o en Dashboard → Settings → API revisá que el proyecto esté activo.'
                );
            }
            throw err;
        }
    }

    window.TerritoriosSupabase = {
        url: SUPABASE_URL,
        cargarDatosSemana,
        guardarAccion,
        listarAccionesPendientes,
        actualizarAccion,
        eliminarAccion,
        eliminarAccionesPendientesPorNumero,
        marcarAccionProcesada,
        upsertDatosSemana,
        CANAL_DATOS_PUBLICOS
    };
}());
