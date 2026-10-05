/* Cola de ediciones de la página de salidas.
   Se escribe en el teléfono al instante y se envía a Google en segundo plano. */
(function () {
    'use strict';

    const DB_NAME = 'territorios-salidas';
    const DB_VER = 1;
    const ESPERA_ERROR_MS = 12000;

    let dbp = null;
    let memoria = [];
    let lista = false;
    let corriendo = false;
    let apiRef = null;
    let timer = null;
    const oyentes = [];
    const resultados = [];

    function abrir() {
        if (dbp) return dbp;
        dbp = new Promise(function (resolve, reject) {
            const req = indexedDB.open(DB_NAME, DB_VER);
            req.onupgradeneeded = function () {
                const db = req.result;
                if (!db.objectStoreNames.contains('cola')) db.createObjectStore('cola', { keyPath: 'localId' });
                if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta');
            };
            req.onsuccess = function () { resolve(req.result); };
            req.onerror = function () { reject(req.error); };
        });
        return dbp;
    }

    function txDone(tx) {
        return new Promise(function (resolve, reject) {
            tx.oncomplete = function () { resolve(); };
            tx.onerror = function () { reject(tx.error); };
            tx.onabort = function () { reject(tx.error || new Error('transacción abortada')); };
        });
    }

    async function leerTodos() {
        const db = await abrir();
        const tx = db.transaction('cola', 'readonly');
        const req = tx.objectStore('cola').getAll();
        const rows = await new Promise(function (resolve, reject) {
            req.onsuccess = function () { resolve(req.result || []); };
            req.onerror = function () { reject(req.error); };
        });
        await txDone(tx);
        return rows;
    }

    async function hidratar() {
        if (lista) return memoria.slice();
        try {
            memoria = await leerTodos();
        } catch (_) {
            memoria = [];
        }
        lista = true;
        let cambio = false;
        memoria.forEach(function (r) {
            if (r.estado !== 'enviando') return;
            r.estado = r.modo === 'eliminar' ? 'eliminar' : 'pendiente';
            r.intento = 0;
            cambio = true;
        });
        if (cambio) {
            try {
                const db = await abrir();
                const tx = db.transaction('cola', 'readwrite');
                const store = tx.objectStore('cola');
                memoria.forEach(function (r) { store.put(r); });
                await txDone(tx);
            } catch (_) {}
        }
        return memoria.slice();
    }

    async function guardarReg(reg) {
        const db = await abrir();
        const tx = db.transaction('cola', 'readwrite');
        tx.objectStore('cola').put(reg);
        await txDone(tx);
        const i = memoria.findIndex(function (r) { return r.localId === reg.localId; });
        if (i >= 0) memoria[i] = reg;
        else memoria.push(reg);
    }

    async function quitar(localId) {
        const db = await abrir();
        const tx = db.transaction('cola', 'readwrite');
        tx.objectStore('cola').delete(localId);
        await txDone(tx);
        memoria = memoria.filter(function (r) { return r.localId !== localId; });
    }

    async function obtener(localId) {
        await hidratar();
        try {
            const db = await abrir();
            const tx = db.transaction('cola', 'readonly');
            const req = tx.objectStore('cola').get(localId);
            const row = await new Promise(function (resolve, reject) {
                req.onsuccess = function () { resolve(req.result || null); };
                req.onerror = function () { reject(req.error); };
            });
            await txDone(tx);
            if (!row) {
                memoria = memoria.filter(function (r) { return r.localId !== localId; });
                return null;
            }
            const i = memoria.findIndex(function (r) { return r.localId === localId; });
            if (i >= 0) memoria[i] = row;
            else memoria.push(row);
            return row;
        } catch (_) {
            return memoria.find(function (r) { return r.localId === localId; }) || null;
        }
    }

    function emitir() {
        const copia = memoria.slice();
        oyentes.forEach(function (fn) {
            try { fn(copia); } catch (_) {}
        });
    }

    function avisar(evento) {
        resultados.forEach(function (fn) {
            try { fn(evento); } catch (_) {}
        });
    }

    function nuevoId() {
        if (window.crypto && crypto.randomUUID) return 'local-' + crypto.randomUUID();
        return 'local-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
    }

    function clonar(obj) {
        try { return JSON.parse(JSON.stringify(obj || {})); }
        catch (_) { return {}; }
    }

    function resumen() {
        const enviando = memoria.filter(function (r) { return r.estado === 'enviando'; }).length;
        const esperando = memoria.some(function (r) {
            return r.estado === 'pendiente' || r.estado === 'eliminar';
        });
        const conError = memoria.some(function (r) { return r.estado === 'error'; });
        return {
            pendientes: enviando,
            enCurso: enviando > 0,
            error: conError && enviando === 0 && !esperando
        };
    }

    function programar(ms) {
        if (timer) clearTimeout(timer);
        timer = setTimeout(function () {
            timer = null;
            drenar(apiRef);
        }, ms);
    }

    function elegir() {
        const ahora = Date.now();
        return memoria.find(function (r) {
            if (r.reintentarEn && r.reintentarEn > ahora) return false;
            return r.estado === 'pendiente' || r.estado === 'eliminar' || r.estado === 'error';
        }) || null;
    }

    async function enviarUno(item) {
        const api = apiRef;
        if (!api) return;
        const rev = item.revision;
        const localId = item.localId;
        const tipo = item.tipo;
        const payload = clonar(item.payload);
        const serverId = item.serverId;
        const numero = item.numero;
        const eraEliminar = item.estado === 'eliminar' || item.modo === 'eliminar';
        const esReintento = (item.intento || 0) > 0;
        item.intento = (item.intento || 0) + 1;
        item.actualizado_en = new Date().toISOString();
        if (!esReintento) {
            item.estado = 'enviando';
            item.error = '';
            await guardarReg(item);
            emitir();
        } else {
            await guardarReg(item);
        }
        try {
            let row = null;
            if (eraEliminar) {
                await api.eliminar(serverId, numero);
            } else if (serverId != null) {
                row = await api.actualizar(serverId, tipo, payload);
            } else {
                row = await api.guardar(tipo, payload);
            }
            const now = await obtener(localId);
            if (!now) {
                if (!eraEliminar && serverId == null && row && row.id != null && api.eliminar) {
                    try { await api.eliminar(row.id, numero); } catch (_) {}
                }
                return;
            }
            if (now.revision !== rev) {
                if (serverId == null && row && row.id != null) now.serverId = row.id;
                if (now.estado === 'enviando') now.estado = 'pendiente';
                await guardarReg(now);
                emitir();
                return;
            }
            await quitar(localId);
            if (eraEliminar) avisar({ tipo: 'eliminado', numero: numero });
            else {
                avisar({
                    tipo: 'confirmado',
                    numero: numero,
                    row: {
                        id: (row && row.id != null) ? row.id : serverId,
                        tipo: tipo,
                        payload: payload,
                        creado_en: (row && row.creado_en) || item.creado_en || new Date().toISOString(),
                        origen: (row && row.origen) || 'predicacion-github'
                    }
                });
            }
            emitir();
        } catch (err) {
            const now = await obtener(localId);
            if (!now || now.revision !== rev) return;
            const msg = String(err && err.message ? err.message : err);
            const yaNoEsta = /No se encontró la acción|No se eliminó ninguna|ya procesada/i.test(msg);
            if (eraEliminar && yaNoEsta) {
                await quitar(localId);
                emitir();
                return;
            }
            if (yaNoEsta) now.serverId = null;
            now.estado = eraEliminar ? 'eliminar' : (yaNoEsta ? 'pendiente' : 'error');
            now.error = msg;
            now.reintentarEn = Date.now() + ESPERA_ERROR_MS;
            now.actualizado_en = new Date().toISOString();
            await guardarReg(now);
            emitir();
            avisar({ tipo: 'error', numero: numero, error: now.error });
            programar(ESPERA_ERROR_MS);
        }
    }

    async function ciclo() {
        for (let i = 0; i < 40; i++) {
            const item = elegir();
            if (!item) return;
            await enviarUno(item);
        }
    }

    async function drenar(api) {
        if (api) apiRef = api;
        if (!apiRef) return;
        if (corriendo) return;
        corriendo = true;
        try {
            await hidratar();
            await ciclo();
        } catch (_) {
            programar(ESPERA_ERROR_MS);
        } finally {
            corriendo = false;
            if (elegir()) drenar(apiRef);
        }
    }

    async function encolar(opts) {
        await hidratar();
        const num = String(opts && opts.numero || '').trim();
        if (!num) throw new Error('Número de territorio inválido.');
        const prev = memoria.find(function (r) { return r.numero === num; });
        const ahora = new Date().toISOString();
        let serverId = opts && opts.serverId != null ? opts.serverId : null;
        if (serverId == null && prev && prev.serverId != null) serverId = prev.serverId;
        const reg = {
            localId: prev ? prev.localId : nuevoId(),
            numero: num,
            tipo: opts.tipo,
            payload: clonar(opts.payload),
            serverId: serverId,
            estado: 'pendiente',
            modo: 'guardar',
            error: '',
            revision: prev ? (Number(prev.revision) || 0) + 1 : 1,
            creado_en: prev ? prev.creado_en : ahora,
            actualizado_en: ahora,
            reintentarEn: 0
        };
        await guardarReg(reg);
        emitir();
        drenar(apiRef);
        return reg;
    }

    async function encolarEliminacion(numero, serverId) {
        await hidratar();
        const num = String(numero || '').trim();
        if (!num) throw new Error('Número de territorio inválido.');
        const prev = memoria.find(function (r) { return r.numero === num; });
        const sid = serverId != null ? serverId : (prev && prev.serverId != null ? prev.serverId : null);
        if (sid == null) {
            if (prev) await quitar(prev.localId);
            emitir();
            return null;
        }
        const ahora = new Date().toISOString();
        const reg = {
            localId: prev ? prev.localId : nuevoId(),
            numero: num,
            tipo: prev ? prev.tipo : 'eliminar',
            payload: prev ? clonar(prev.payload) : { numero: num },
            serverId: sid,
            estado: 'eliminar',
            modo: 'eliminar',
            error: '',
            revision: prev ? (Number(prev.revision) || 0) + 1 : 1,
            creado_en: prev ? prev.creado_en : ahora,
            actualizado_en: ahora,
            reintentarEn: 0
        };
        await guardarReg(reg);
        emitir();
        drenar(apiRef);
        return reg;
    }

    async function guardarSnapshot(doc) {
        const db = await abrir();
        const tx = db.transaction('meta', 'readwrite');
        tx.objectStore('meta').put({
            bundle: doc && doc.bundle ? doc.bundle : null,
            pendientes: doc && Array.isArray(doc.pendientes) ? doc.pendientes : [],
            guardado_en: new Date().toISOString()
        }, 'snapshot');
        await txDone(tx);
    }

    async function leerSnapshot() {
        try {
            const db = await abrir();
            const tx = db.transaction('meta', 'readonly');
            const req = tx.objectStore('meta').get('snapshot');
            const row = await new Promise(function (resolve, reject) {
                req.onsuccess = function () { resolve(req.result || null); };
                req.onerror = function () { reject(req.error); };
            });
            await txDone(tx);
            return row;
        } catch (_) {
            return null;
        }
    }

    function firmaAccion(tipo, payload) {
        const p = payload && typeof payload === 'object' ? payload : {};
        const limpio = {};
        Object.keys(p).sort().forEach(function (k) {
            const v = p[k];
            if (v == null) return;
            if (typeof v === 'string' && v.trim() === '') return;
            limpio[k] = typeof v === 'string' ? v.trim() : v;
        });
        return String(tipo || '') + '\n' + JSON.stringify(limpio);
    }

    function numeroDeFilaServidor(row) {
        return String((row && row.payload && row.payload.numero) || '').trim();
    }

    async function reconciliar(filas) {
        await hidratar();
        const lista = Array.isArray(filas) ? filas : [];
        const copia = memoria.slice();
        let cambio = false;
        for (let i = 0; i < copia.length; i++) {
            const reg = copia[i];
            if (!reg || reg.estado === 'enviando') continue;
            const num = String(reg.numero || '').trim();
            const delNum = lista.filter(function (row) { return numeroDeFilaServidor(row) === num; });
            const esBorrado = reg.estado === 'eliminar' || reg.modo === 'eliminar';
            if (esBorrado) {
                if (!delNum.length) {
                    await quitar(reg.localId);
                    cambio = true;
                }
                continue;
            }
            const firmaLocal = firmaAccion(reg.tipo, reg.payload);
            const yaEsta = delNum.some(function (row) {
                return firmaAccion(row.tipo, row.payload) === firmaLocal;
            });
            if (yaEsta) {
                await quitar(reg.localId);
                cambio = true;
                continue;
            }
            const sigueEnGoogle = reg.serverId != null && delNum.some(function (row) {
                return String(row.id) === String(reg.serverId);
            });
            if (reg.serverId != null && !sigueEnGoogle) {
                reg.serverId = null;
                if (reg.estado === 'error') reg.estado = 'pendiente';
                await guardarReg(reg);
                cambio = true;
            }
        }
        if (cambio) emitir();
    }

    function onCambio(fn) { if (typeof fn === 'function') oyentes.push(fn); }
    function onResultado(fn) { if (typeof fn === 'function') resultados.push(fn); }

    function arrancar(api) {
        apiRef = api;
        hidratar().then(function () { emitir(); }).catch(function () {});
        window.addEventListener('online', function () { drenar(apiRef); });
    }

    window.ColaLocal = {
        hidratar: hidratar,
        listar: async function () { return hidratar(); },
        resumen: resumen,
        encolar: encolar,
        encolarEliminacion: encolarEliminacion,
        reconciliar: reconciliar,
        drenar: drenar,
        arrancar: arrancar,
        onCambio: onCambio,
        onResultado: onResultado,
        guardarSnapshot: guardarSnapshot,
        leerSnapshot: leerSnapshot
    };
}());
