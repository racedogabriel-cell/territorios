/* Mapa Leaflet del modal de predicación: formas de territorio + pin GPS anclado. */
(function (global) {
    'use strict';

    const GOOGLE_TILE_SUBS = '0123';
    // Misma línea de vía que mapa.html / mapa-semanal.html
    const LINEA_VIA_SVG = { x1: 28, y1: 392, x2: 1040, y2: 366 };
    const AJUSTE_GRUPOS_INFO = { arriba: { nombre: 'Arriba' }, abajo: { nombre: 'Abajo' } };
    const BOUNDS_CALIB_VERSION = 3;
    const NUMEROS_EXCLUIDOS_BOUNDS_CALIBRACION = new Set(['205', '206', '207']);

    let MAPA = null;
    let CAPA = null;
    let MARKER = null;
    let META = { rurales: {}, gps: {} };
    let TERRITORIOS = [];
    let cargandoGeometria = null;
    let AJUSTE_GLOBAL = { groups: null, compat: null, boundsBase: null, centroBloqueRef: null, boundsCalibVersion: 0 };
    let CENTRO_BLOQUE_BASE = null;
    let BOUNDS_LOCAL_BASE = null;
    let GRUPOS_AJUSTE_BASE = {};

    function urlTileGoogleMapaCalle() {
        return 'https://mt{s}.google.com/vt/lyrs=m&hl=es&x={x}&y={y}&z={z}';
    }

    function leerLatLon(metaEntry) {
        if (!metaEntry) return null;
        const lat = Number(metaEntry.lat ?? metaEntry.latitud ?? metaEntry.latitude);
        const lon = Number(metaEntry.lon ?? metaEntry.lng ?? metaEntry.longitud ?? metaEntry.longitude);
        if (!Number.isFinite(lat) || !Number.isFinite(lon) || (lat === 0 && lon === 0)) return null;
        return { lat, lon };
    }

    function metrosPorGradoLon(lat) {
        return Math.cos(lat * Math.PI / 180) * 111320;
    }
    function geoALocal(lat, lon, ref) {
        return { x: (lon - ref.lon) * metrosPorGradoLon(ref.lat), y: (lat - ref.lat) * 111320 };
    }
    function localAGeo(x, y, ref) {
        return { lat: ref.lat + y / 111320, lon: ref.lon + x / metrosPorGradoLon(ref.lat) };
    }
    function boundsDesdePuntosLocal(localPts) {
        if (!localPts.length) return null;
        return {
            minX: Math.min(...localPts.map((p) => p.x)),
            maxX: Math.max(...localPts.map((p) => p.x)),
            minY: Math.min(...localPts.map((p) => p.y)),
            maxY: Math.max(...localPts.map((p) => p.y))
        };
    }
    function cornersDesdeBounds(b) {
        if (!b) return null;
        return {
            nw: { x: b.minX, y: b.maxY },
            ne: { x: b.maxX, y: b.maxY },
            se: { x: b.maxX, y: b.minY },
            sw: { x: b.minX, y: b.minY }
        };
    }
    function normalizarCorners(c) {
        if (!c || typeof c !== 'object') return null;
        const o = {};
        for (const k of ['nw', 'ne', 'se', 'sw']) {
            const p = c[k];
            if (!p) return null;
            const x = Number(p.x);
            const y = Number(p.y);
            if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
            o[k] = { x, y };
        }
        return o;
    }
    function normalizarGruposAjuste(g) {
        if (!g || typeof g !== 'object') return null;
        const o = {};
        for (const k of Object.keys(AJUSTE_GRUPOS_INFO)) {
            const c = normalizarCorners(g[k]);
            if (c) o[k] = c;
        }
        return Object.keys(o).length ? o : null;
    }
    function normalizarBounds(b) {
        if (!b || typeof b !== 'object') return null;
        const minX = Number(b.minX);
        const maxX = Number(b.maxX);
        const minY = Number(b.minY);
        const maxY = Number(b.maxY);
        if (![minX, maxX, minY, maxY].every(Number.isFinite)) return null;
        return { minX, maxX, minY, maxY };
    }
    function normalizarBoundsBase(o) {
        if (!o || typeof o !== 'object') return null;
        const out = {};
        for (const k of Object.keys(AJUSTE_GRUPOS_INFO)) {
            const b = normalizarBounds(o[k]);
            if (b) out[k] = b;
        }
        return Object.keys(out).length ? out : null;
    }
    function boundsGrupoBase(k) {
        return GRUPOS_AJUSTE_BASE[k]?.bounds || null;
    }
    function cornersGrupoBase(k) {
        return cornersDesdeBounds(boundsGrupoBase(k));
    }
    function transformarLocalPorCorners(x, y, bounds, corners) {
        if (!bounds || !corners) return { x, y };
        const denX = (bounds.maxX - bounds.minX) || 1;
        const denY = (bounds.maxY - bounds.minY) || 1;
        const s = (x - bounds.minX) / denX;
        const t = (y - bounds.minY) / denY;
        return {
            x: (1 - s) * (1 - t) * corners.sw.x + s * (1 - t) * corners.se.x + (1 - s) * t * corners.nw.x + s * t * corners.ne.x,
            y: (1 - s) * (1 - t) * corners.sw.y + s * (1 - t) * corners.se.y + (1 - s) * t * corners.nw.y + s * t * corners.ne.y
        };
    }
    function transformarLocalCompat(x, y) {
        const v = AJUSTE_GLOBAL.compat;
        if (!v) return { x, y };
        if (v.mode === 'single-bilinear') return transformarLocalPorCorners(x, y, BOUNDS_LOCAL_BASE, v.corners);
        const dx = Number.isFinite(Number(v.dx)) ? Number(v.dx) : 0;
        const dy = Number.isFinite(Number(v.dy)) ? Number(v.dy) : 0;
        const rot = Number.isFinite(Number(v.rot)) ? Number(v.rot) : 0;
        const scaleCompat = Number.isFinite(Number(v.scale)) ? Number(v.scale) : 1;
        const scaleX = Number.isFinite(Number(v.scaleX)) ? Number(v.scaleX) : scaleCompat;
        const scaleY = Number.isFinite(Number(v.scaleY)) ? Number(v.scaleY) : scaleCompat;
        const ang = rot * Math.PI / 180;
        const cos = Math.cos(ang);
        const sin = Math.sin(ang);
        const xs = scaleX * x;
        const ys = scaleY * y;
        return { x: cos * xs - sin * ys + dx, y: sin * xs + cos * ys + dy };
    }
    function asegurarGruposAjuste() {
        if (AJUSTE_GLOBAL.groups || !Object.keys(GRUPOS_AJUSTE_BASE).length) return;
        const groups = {};
        for (const key of Object.keys(AJUSTE_GRUPOS_INFO)) {
            const base = cornersGrupoBase(key);
            if (!base) continue;
            if (AJUSTE_GLOBAL.compat) {
                const next = {};
                for (const [corner, p] of Object.entries(base)) next[corner] = transformarLocalCompat(p.x, p.y);
                groups[key] = next;
            } else {
                groups[key] = base;
            }
        }
        AJUSTE_GLOBAL.groups = groups;
        AJUSTE_GLOBAL.compat = null;
    }
    function aplicarValorAjusteGlobal(v) {
        if (!v || typeof v !== 'object') return;
        AJUSTE_GLOBAL.groups = normalizarGruposAjuste(v.groups);
        AJUSTE_GLOBAL.boundsBase = normalizarBoundsBase(v.boundsBase);
        const cb = v.centroBloque;
        if (cb && Number.isFinite(Number(cb.lat)) && Number.isFinite(Number(cb.lon))) {
            AJUSTE_GLOBAL.centroBloqueRef = { lat: Number(cb.lat), lon: Number(cb.lon) };
        } else {
            AJUSTE_GLOBAL.centroBloqueRef = null;
        }
        AJUSTE_GLOBAL.boundsCalibVersion = Number(v.boundsCalibVersion) || 0;
        AJUSTE_GLOBAL.compat = null;
        if (!AJUSTE_GLOBAL.groups && v.corners) {
            AJUSTE_GLOBAL.compat = { mode: 'single-bilinear', corners: normalizarCorners(v.corners) };
        }
        if (!AJUSTE_GLOBAL.groups && !AJUSTE_GLOBAL.compat && ('dx' in v || 'rot' in v || 'scaleX' in v || 'scaleY' in v || 'scale' in v)) {
            AJUSTE_GLOBAL.compat = v;
        }
    }
    function aplicarAjusteGeo(lat, lon, groupKey) {
        if (!CENTRO_BLOQUE_BASE || !groupKey) return { lat, lon };
        asegurarGruposAjuste();
        const bounds = (AJUSTE_GLOBAL.boundsBase && AJUSTE_GLOBAL.boundsBase[groupKey]) || boundsGrupoBase(groupKey);
        const corners = AJUSTE_GLOBAL.groups?.[groupKey];
        if (!bounds || !corners) return { lat, lon };
        const ref = AJUSTE_GLOBAL.centroBloqueRef || CENTRO_BLOQUE_BASE;
        const p = geoALocal(lat, lon, ref);
        const adj = transformarLocalPorCorners(p.x, p.y, bounds, corners);
        return localAGeo(adj.x, adj.y, ref);
    }
    function calcularSnapshotBoundsCalibracion(territorios, excluirSet) {
        const base = (territorios || []).filter((t) => !excluirSet.has(String(t.numero)));
        const todos = base.flatMap((t) => t.latlngsBase);
        if (!todos.length) return { boundsBase: null, centroBloque: null };
        const avg = todos.reduce((acc, p) => ({ lat: acc.lat + p[0], lon: acc.lon + p[1] }), { lat: 0, lon: 0 });
        const centroBloque = { lat: avg.lat / todos.length, lon: avg.lon / todos.length };
        const boundsBase = {};
        for (const key of Object.keys(AJUSTE_GRUPOS_INFO)) {
            const ptsGrupo = base
                .filter((t) => t.groupKey === key)
                .flatMap((t) => t.latlngsBase)
                .map((p) => geoALocal(p[0], p[1], centroBloque));
            if (ptsGrupo.length) {
                const b = boundsDesdePuntosLocal(ptsGrupo);
                if (b) boundsBase[key] = { minX: b.minX, maxX: b.maxX, minY: b.minY, maxY: b.maxY };
            }
        }
        return { centroBloque, boundsBase: Object.keys(boundsBase).length ? boundsBase : null };
    }
    function reconciliarBoundsCalibracion(territorios) {
        if (!AJUSTE_GLOBAL.groups) return;
        if (AJUSTE_GLOBAL.boundsCalibVersion === BOUNDS_CALIB_VERSION && AJUSTE_GLOBAL.boundsBase && AJUSTE_GLOBAL.centroBloqueRef) return;
        const snap = calcularSnapshotBoundsCalibracion(territorios, NUMEROS_EXCLUIDOS_BOUNDS_CALIBRACION);
        if (snap.boundsBase) AJUSTE_GLOBAL.boundsBase = snap.boundsBase;
        if (snap.centroBloque) AJUSTE_GLOBAL.centroBloqueRef = snap.centroBloque;
        AJUSTE_GLOBAL.boundsCalibVersion = BOUNDS_CALIB_VERSION;
    }
    function recalcularGeometriasAjustadas() {
        asegurarGruposAjuste();
        for (const t of TERRITORIOS) {
            t.latlngs = t.latlngsBase.map(([lat, lon]) => {
                const p = aplicarAjusteGeo(lat, lon, t.groupKey);
                return [p.lat, p.lon];
            });
            const lbl = t.labelBase ? aplicarAjusteGeo(t.labelBase.lat, t.labelBase.lon, t.groupKey) : null;
            t.labelPos = lbl;
        }
    }

    function extraerPuntosPolygon(shape) {
        return String(shape.getAttribute('points') || '')
            .trim()
            .split(/\s+/)
            .filter(Boolean)
            .map((par) => par.split(',').map(Number))
            .filter((pt) => pt.length >= 2 && Number.isFinite(pt[0]) && Number.isFinite(pt[1]));
    }
    function centroRect(shape) {
        const x = Number(shape.getAttribute('x'));
        const y = Number(shape.getAttribute('y'));
        const w = Number(shape.getAttribute('width'));
        const h = Number(shape.getAttribute('height'));
        return { cx: x + w / 2, cy: y + h / 2, minDim: Math.min(w, h) };
    }
    function centroPolygonSimple(pts) {
        if (!pts.length) return { cx: 0, cy: 0, minDim: 0 };
        const xs = pts.map((v) => v[0]);
        const ys = pts.map((v) => v[1]);
        return {
            cx: (Math.min(...xs) + Math.max(...xs)) / 2,
            cy: (Math.min(...ys) + Math.max(...ys)) / 2,
            minDim: Math.min(Math.max(...xs) - Math.min(...xs), Math.max(...ys) - Math.min(...ys))
        };
    }
    function extraerGeometriaSvg(shape) {
        const tag = shape.tagName.toLowerCase();
        if (tag === 'rect') {
            const x = Number(shape.getAttribute('x'));
            const y = Number(shape.getAttribute('y'));
            const w = Number(shape.getAttribute('width'));
            const h = Number(shape.getAttribute('height'));
            return {
                tipo: 'rect',
                puntos: [[x, y], [x + w, y], [x + w, y + h], [x, y + h]],
                centro: centroRect(shape)
            };
        }
        if (tag === 'polygon') {
            const puntos = extraerPuntosPolygon(shape);
            return { tipo: 'polygon', puntos, centro: centroPolygonSimple(puntos) };
        }
        return null;
    }
    function resolverSistema3x3(A, b) {
        const m = [
            [A[0][0], A[0][1], A[0][2], b[0]],
            [A[1][0], A[1][1], A[1][2], b[1]],
            [A[2][0], A[2][1], A[2][2], b[2]]
        ];
        for (let col = 0; col < 3; col++) {
            let piv = col;
            for (let row = col + 1; row < 3; row++) {
                if (Math.abs(m[row][col]) > Math.abs(m[piv][col])) piv = row;
            }
            if (Math.abs(m[piv][col]) < 1e-12) return null;
            if (piv !== col) [m[col], m[piv]] = [m[piv], m[col]];
            const div = m[col][col];
            for (let j = col; j < 4; j++) m[col][j] /= div;
            for (let row = 0; row < 3; row++) {
                if (row === col) continue;
                const fac = m[row][col];
                for (let j = col; j < 4; j++) m[row][j] -= fac * m[col][j];
            }
        }
        return [m[0][3], m[1][3], m[2][3]];
    }
    function ajustarTransformacionAfin(controles) {
        if (!controles || controles.length < 3) return null;
        let sxx = 0, sxy = 0, sx = 0, syy = 0, sy = 0, n = 0;
        let blx = 0, bly = 0, bl1 = 0, bax = 0, bay = 0, ba1 = 0;
        for (const p of controles) {
            sxx += p.x * p.x; sxy += p.x * p.y; sx += p.x;
            syy += p.y * p.y; sy += p.y; n += 1;
            blx += p.x * p.lon; bly += p.y * p.lon; bl1 += p.lon;
            bax += p.x * p.lat; bay += p.y * p.lat; ba1 += p.lat;
        }
        const A = [[sxx, sxy, sx], [sxy, syy, sy], [sx, sy, n]];
        const lonCoef = resolverSistema3x3(A, [blx, bly, bl1]);
        const latCoef = resolverSistema3x3(A, [bax, bay, ba1]);
        if (!lonCoef || !latCoef) return null;
        return { lonCoef, latCoef };
    }
    function aplicarTransformacion(tr, x, y) {
        return {
            lon: tr.lonCoef[0] * x + tr.lonCoef[1] * y + tr.lonCoef[2],
            lat: tr.latCoef[0] * x + tr.latCoef[1] * y + tr.latCoef[2]
        };
    }
    function yLineaViaEnSvg(x) {
        const den = (LINEA_VIA_SVG.x2 - LINEA_VIA_SVG.x1) || 1;
        return LINEA_VIA_SVG.y1 + (LINEA_VIA_SVG.y2 - LINEA_VIA_SVG.y1) * ((x - LINEA_VIA_SVG.x1) / den);
    }
    function grupoSegunViaSvg(x, y) {
        return y <= yLineaViaEnSvg(x) ? 'arriba' : 'abajo';
    }

    function construirTerritoriosDesdeSvg(svgTxt) {
        const doc = new DOMParser().parseFromString(svgTxt, 'image/svg+xml');
        const svg = doc.querySelector('svg');
        if (!svg) throw new Error('No se pudo leer mapa.svg');
        const shapes = [];
        svg.querySelectorAll('.territorio').forEach((shape) => {
            const numero = String(shape.getAttribute('data-n') || '').trim();
            if (!numero || numero === '500') return;
            const geom = extraerGeometriaSvg(shape);
            if (!geom || !geom.puntos || geom.puntos.length < 3) return;
            shapes.push({
                numero,
                tipo: geom.tipo,
                puntosSvg: geom.puntos,
                centroSvg: geom.centro,
                isRural: Number(numero) >= 200 && !!(META.rurales && META.rurales[numero])
            });
        });
        const controlesRect = shapes
            .filter((s) => s.tipo === 'rect' && !s.isRural)
            .map((s) => {
                const gps = leerLatLon(META.gps && META.gps[s.numero]);
                return gps ? { x: s.centroSvg.cx, y: s.centroSvg.cy, lat: gps.lat, lon: gps.lon } : null;
            })
            .filter(Boolean);
        const controles = controlesRect.length >= 8
            ? controlesRect
            : shapes
                .map((s) => {
                    const gps = leerLatLon(META.gps && META.gps[s.numero]);
                    return gps && !s.isRural
                        ? { x: s.centroSvg.cx, y: s.centroSvg.cy, lat: gps.lat, lon: gps.lon }
                        : null;
                })
                .filter(Boolean);
        const tr = ajustarTransformacionAfin(controles);
        if (!tr) throw new Error('No se pudo calibrar el mapa.');
        const territorios = shapes.map((shape) => {
            const latlngsBase = shape.puntosSvg
                .map(([x, y]) => aplicarTransformacion(tr, x, y))
                .filter((p) => Number.isFinite(p.lat) && Number.isFinite(p.lon))
                .map((p) => [p.lat, p.lon]);
            const centroGeo = aplicarTransformacion(tr, shape.centroSvg.cx, shape.centroSvg.cy);
            return {
                numero: shape.numero,
                isRural: shape.isRural,
                groupKey: grupoSegunViaSvg(shape.centroSvg.cx, shape.centroSvg.cy),
                latlngsBase,
                labelBase: centroGeo,
                latlngs: latlngsBase.map((p) => [p[0], p[1]]),
                labelPos: centroGeo
            };
        });
        const todos = territorios.flatMap((t) => t.latlngsBase);
        if (todos.length) {
            const avg = todos.reduce((acc, p) => ({ lat: acc.lat + p[0], lon: acc.lon + p[1] }), { lat: 0, lon: 0 });
            CENTRO_BLOQUE_BASE = { lat: avg.lat / todos.length, lon: avg.lon / todos.length };
            const localPts = todos.map((p) => geoALocal(p[0], p[1], CENTRO_BLOQUE_BASE));
            BOUNDS_LOCAL_BASE = boundsDesdePuntosLocal(localPts);
            GRUPOS_AJUSTE_BASE = {};
            for (const key of Object.keys(AJUSTE_GRUPOS_INFO)) {
                const ptsGrupo = territorios
                    .filter((t) => t.groupKey === key)
                    .flatMap((t) => t.latlngsBase)
                    .map((p) => geoALocal(p[0], p[1], CENTRO_BLOQUE_BASE));
                if (ptsGrupo.length) GRUPOS_AJUSTE_BASE[key] = { bounds: boundsDesdePuntosLocal(ptsGrupo) };
            }
        }
        return territorios;
    }

    async function cargarAjustePortable() {
        // 1) API local Access (si está)
        try {
            const r = await fetch('/api/portable-ui', { cache: 'no-store' });
            if (r.ok) {
                const doc = await r.json();
                if (doc && doc.ajusteGlobal) {
                    aplicarValorAjusteGlobal(doc.ajusteGlobal);
                    return;
                }
            }
        } catch (_) {}
        // 2) Archivo portable en la raíz del proyecto / docs
        for (const url of ['../../territorios_portable.json', '../territorios_portable.json', '/territorios_portable.json']) {
            try {
                const r = await fetch(url, { cache: 'no-store' });
                if (!r.ok) continue;
                const doc = await r.json();
                if (doc && doc.ajusteGlobal) {
                    aplicarValorAjusteGlobal(doc.ajusteGlobal);
                    return;
                }
            } catch (_) {}
        }
    }

    function iconoPuntoEncuentro() {
        return L.divIcon({
            className: 'mm-leaflet-pin-wrap',
            html: '<div class="mm-leaflet-pin"></div>',
            iconSize: [18, 18],
            iconAnchor: [9, 9]
        });
    }

    function asegurarMapa(containerEl) {
        if (MAPA) return MAPA;
        MAPA = L.map(containerEl, {
            zoomControl: true,
            attributionControl: false,
            doubleClickZoom: false,
            scrollWheelZoom: true,
            worldCopyJump: false
        });
        // Fondo tipo calles (como My Maps), sin satélite.
        L.tileLayer(urlTileGoogleMapaCalle(), {
            subdomains: GOOGLE_TILE_SUBS,
            maxZoom: 21,
            maxNativeZoom: 20,
            opacity: 1
        }).addTo(MAPA);
        return MAPA;
    }

    async function asegurarGeometria() {
        if (TERRITORIOS.length) return;
        if (cargandoGeometria) return cargandoGeometria;
        cargandoGeometria = (async () => {
            const [metaResp, svgResp] = await Promise.all([
                fetch('../mapa_meta.json', { cache: 'no-store' }),
                fetch('../mapa.svg', { cache: 'no-store' })
            ]);
            if (!metaResp.ok) throw new Error('HTTP ' + metaResp.status + ' en mapa_meta.json');
            if (!svgResp.ok) throw new Error('HTTP ' + svgResp.status + ' en mapa.svg');
            META = await metaResp.json();
            if (!META || typeof META !== 'object') META = { rurales: {}, gps: {} };
            const svgTxt = await svgResp.text();
            TERRITORIOS = construirTerritoriosDesdeSvg(svgTxt);
            await cargarAjustePortable();
            reconciliarBoundsCalibracion(TERRITORIOS);
            recalcularGeometriasAjustadas();
        })().finally(() => {
            cargandoGeometria = null;
        });
        return cargandoGeometria;
    }

    function gpsDeTerritorio(numero) {
        const num = String(numero ?? '').trim();
        const gps = META && META.gps && typeof META.gps === 'object' ? META.gps : {};
        const directo = leerLatLon(gps[num]);
        if (directo) return directo;
        if (num.includes('.')) return leerLatLon(gps[num.split('.')[0]]);
        return null;
    }

    function esperarLayout() {
        return new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    }

    async function mostrar(opts) {
        const {
            containerEl,
            numero,
            puntoTxt,
            gpsPunto
        } = opts || {};
        if (!containerEl) throw new Error('Falta contenedor del mapa');
        await asegurarGeometria();
        await esperarLayout();
        asegurarMapa(containerEl);
        MAPA.invalidateSize(true);

        if (CAPA) {
            MAPA.removeLayer(CAPA);
            CAPA = null;
        }
        MARKER = null;
        CAPA = L.layerGroup().addTo(MAPA);

        const keySel = String(numero || '').trim().replace(',', '.');
        let polySel = null;
        const polys = [];
        for (const t of TERRITORIOS) {
            const sel = String(t.numero) === keySel;
            if (!t.latlngs || t.latlngs.length < 3) continue;
            const poly = L.polygon(t.latlngs, {
                color: sel ? '#1a73e8' : '#5f6368',
                weight: sel ? 3.5 : 1.2,
                opacity: sel ? 1 : 0.55,
                fillColor: sel ? '#1a73e8' : '#9aa0a6',
                fillOpacity: sel ? 0.28 : 0.08,
                interactive: false
            }).addTo(CAPA);
            polys.push(poly);
            if (sel) polySel = poly;
        }

        if (gpsPunto && Number.isFinite(gpsPunto.lat) && Number.isFinite(gpsPunto.lon)) {
            MARKER = L.marker([gpsPunto.lat, gpsPunto.lon], {
                icon: iconoPuntoEncuentro(),
                keyboard: false,
                zIndexOffset: 900
            }).addTo(CAPA);
            const tip = puntoTxt || 'Punto de encuentro';
            MARKER.bindTooltip(tip, { direction: 'top', offset: [0, -10], permanent: true, className: 'mm-leaflet-tip' });
        }

        await esperarLayout();
        MAPA.invalidateSize(true);

        if (polySel && polySel.getBounds && polySel.getBounds().isValid()) {
            MAPA.fitBounds(polySel.getBounds(), { padding: [36, 36], maxZoom: 18 });
        } else if (gpsPunto) {
            MAPA.setView([gpsPunto.lat, gpsPunto.lon], 17);
        } else {
            const gpsTerr = gpsDeTerritorio(keySel);
            if (gpsTerr) MAPA.setView([gpsTerr.lat, gpsTerr.lon], 17);
            else if (polys.length) {
                const b = L.featureGroup(polys).getBounds();
                if (b && b.isValid()) MAPA.fitBounds(b, { padding: [20, 20], maxZoom: 16 });
            }
        }
    }

    function limpiar() {
        if (CAPA && MAPA) {
            try { MAPA.removeLayer(CAPA); } catch (_) {}
        }
        CAPA = null;
        MARKER = null;
    }

    global.PredicacionMapaModal = {
        mostrar,
        limpiar,
        gpsDeTerritorio: function (numero) {
            return gpsDeTerritorio(numero);
        },
        asegurarGeometria
    };
})(window);
