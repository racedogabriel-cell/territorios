/* Service worker solo para la página de salidas.
   mapa-semanal.html no se cachea: sus pedidos salen directo a la red. */
const CACHE = 'salidas-shell-v7';
const ASSETS = [
    './',
    './index.html',
    './google.js',
    './cola-local.js',
    './manifest.json',
    './icon-192.png',
    './icon-512.png',
    './mapa.svg',
    './mapa_meta.json'
];

function esMapa(pathname) {
    return /\/mapa-semanal\.html$/i.test(pathname);
}

function esInicio(pathname) {
    const base = new URL(self.registration.scope).pathname;
    return pathname === base || pathname === base + 'index.html';
}

function esAssetSalidas(pathname) {
    return /\/(index\.html|google\.js|cola-local\.js|manifest\.json|icon-192\.png|icon-512\.png|mapa\.svg|mapa_meta\.json)$/i.test(pathname);
}

async function desdeMapa(clientId) {
    if (!clientId) return false;
    const client = await self.clients.get(clientId);
    if (!client) return false;
    return esMapa(new URL(client.url).pathname);
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
            return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) {
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

    if (req.mode === 'navigate') {
        if (esMapa(url.pathname) || !esInicio(url.pathname)) return;
        event.respondWith(redPrimero(req, './index.html'));
        return;
    }

    if (url.origin !== self.location.origin) {
        if (!/unpkg\.com\/leaflet@/i.test(url.href)) return;
        event.respondWith((async function () {
            if (await desdeMapa(event.clientId)) return fetch(req);
            return redPrimero(req);
        })());
        return;
    }

    if (esMapa(url.pathname) || !esAssetSalidas(url.pathname)) return;

    event.respondWith((async function () {
        if (await desdeMapa(event.clientId)) return fetch(req);
        return redPrimero(req);
    })());
});
