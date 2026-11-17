/* Mapa Leaflet del modal de predicación: formas reales de My Maps + pin GPS anclado. */
(function (global) {
    'use strict';

    const GOOGLE_TILE_SUBS = '0123';
    const GEOJSON_URLS = [
        '../territorios_mymaps.geojson',
        '/docs/territorios_mymaps.geojson',
        '/territorios_mymaps.geojson'
    ];
    const META_URLS = ['../mapa_meta.json', '/docs/mapa_meta.json', '/mapa_meta.json'];

    let MAPA = null;
    let CAPA = null;
    let MARKER = null;
    let META = { rurales: {}, gps: {} };
    let TERRITORIOS = [];
    let cargandoGeometria = null;

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

    function ringToLatLngs(ring) {
        if (!Array.isArray(ring) || ring.length < 3) return null;
        const pts = [];
        for (const p of ring) {
            if (!Array.isArray(p) || p.length < 2) continue;
            const lon = Number(p[0]);
            const lat = Number(p[1]);
            if (!Number.isFinite(lat) || !Number.isFinite(lon)) continue;
            pts.push([lat, lon]);
        }
        return pts.length >= 3 ? pts : null;
    }

    function geometryToLatLngsList(geometry) {
        if (!geometry || !geometry.type) return [];
        const out = [];
        if (geometry.type === 'Polygon') {
            const outer = ringToLatLngs(geometry.coordinates && geometry.coordinates[0]);
            if (outer) out.push(outer);
        } else if (geometry.type === 'MultiPolygon') {
            for (const poly of geometry.coordinates || []) {
                const outer = ringToLatLngs(poly && poly[0]);
                if (outer) out.push(outer);
            }
        }
        return out;
    }

    function construirTerritoriosDesdeGeoJson(doc) {
        const features = Array.isArray(doc && doc.features) ? doc.features : [];
        const byNum = new Map();
        for (const f of features) {
            const props = f && f.properties ? f.properties : {};
            const numero = String(props.numero || props.name || props.Name || '').trim();
            if (!numero) continue;
            const rings = geometryToLatLngsList(f.geometry);
            if (!rings.length) continue;
            if (!byNum.has(numero)) {
                byNum.set(numero, {
                    numero,
                    rings: [],
                    isRural: Number(numero) >= 200
                });
            }
            byNum.get(numero).rings.push(...rings);
        }
        return Array.from(byNum.values());
    }

    async function fetchPrimeroOk(urls) {
        let lastErr = null;
        for (const url of urls) {
            try {
                const r = await fetch(url, { cache: 'no-store' });
                if (!r.ok) {
                    lastErr = new Error('HTTP ' + r.status + ' en ' + url);
                    continue;
                }
                return await r.json();
            } catch (e) {
                lastErr = e;
            }
        }
        throw lastErr || new Error('No se pudo cargar el recurso');
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
            const [geoDoc, metaDoc] = await Promise.all([
                fetchPrimeroOk(GEOJSON_URLS),
                fetchPrimeroOk(META_URLS).catch(() => ({ rurales: {}, gps: {} }))
            ]);
            META = metaDoc && typeof metaDoc === 'object' ? metaDoc : { rurales: {}, gps: {} };
            TERRITORIOS = construirTerritoriosDesdeGeoJson(geoDoc);
            if (!TERRITORIOS.length) throw new Error('No hay formas de territorio en My Maps.');
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
        if (num.includes(',')) return leerLatLon(gps[num.split(',')[0]]);
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
        const keyAlt = String(numero || '').trim().replace('.', ',');
        let polySelGroup = null;
        const polys = [];

        for (const t of TERRITORIOS) {
            const sel = t.numero === keySel || t.numero === keyAlt
                || String(t.numero).replace(',', '.') === keySel;
            for (const latlngs of t.rings) {
                if (!latlngs || latlngs.length < 3) continue;
                const poly = L.polygon(latlngs, {
                    color: sel ? '#1a73e8' : '#5f6368',
                    weight: sel ? 3.5 : 1.2,
                    opacity: sel ? 1 : 0.55,
                    fillColor: sel ? '#1a73e8' : '#9aa0a6',
                    fillOpacity: sel ? 0.28 : 0.08,
                    interactive: false
                }).addTo(CAPA);
                polys.push(poly);
                if (sel) {
                    if (!polySelGroup) polySelGroup = L.featureGroup();
                    polySelGroup.addLayer(poly);
                }
            }
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

        if (polySelGroup) {
            const b = polySelGroup.getBounds();
            if (b && b.isValid()) {
                MAPA.fitBounds(b, { padding: [36, 36], maxZoom: 18 });
            }
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
