/* Service worker de la PWA Predicación (consulta). Scope: ./predicacion/ */
const CACHE = 'predicacion-shell-v5';
const ASSETS = [
    './',
    './index.html',
    './manifest.json',
    './mapa-modal.js',
    './icon-192.png',
    './icon-512.png',
    './icon-maskable-512.png',
    '../google.js',
    '../mapa_meta.json',
    '../territorios_mymaps.geojson'
];

function esNavegacion(request) {
    return request.mode === 'navigate';
}

async function redPrimero(request, respaldo) {
    const cache = await caches.open(CACHE);
    try {
        const fresco = await fetch(request);
        if (fresco && fresco.ok) {
            cache.put(request, fresco.clone()).catch(function () {});
        }
        return fresco;
    } catch (err) {
        const guardado = await cache.match(request) || (respaldo ? await cache.match(respaldo) : null);
        if (guardado) return guardado;
        throw err;
    }
}

self.addEventListener('install', function (event) {
    event.waitUntil(
        caches.open(CACHE).then(function (cache) {
            return Promise.all(ASSETS.map(function (url) {
                return cache.add(url).catch(function () { return null; });
            }));
        }).then(function () {
            return self.skipWaiting();
        })
    );
});

self.addEventListener('activate', function (event) {
    event.waitUntil(
        caches.keys().then(function (keys) {
            return Promise.all(keys.filter(function (k) {
                return k.startsWith('predicacion-shell-') && k !== CACHE;
            }).map(function (k) {
                return caches.delete(k);
            }));
        }).then(function () {
            return self.clients.claim();
        })
    );
});

self.addEventListener('fetch', function (event) {
    const req = event.request;
    if (req.method !== 'GET') return;
    const url = new URL(req.url);
    if (url.origin !== self.location.origin) return;

    if (esNavegacion(req)) {
        event.respondWith(redPrimero(req, './index.html'));
        return;
    }

    if (/\/(predicacion\/(index\.html|manifest\.json|mapa-modal\.js|icon-.*\.png|sw\.js)|google\.js|mapa_meta\.json|territorios_mymaps\.geojson)$/i.test(url.pathname)) {
        event.respondWith(redPrimero(req));
    }
});

