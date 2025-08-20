const CACHE_NAME = 'rota-avail-v2'; // bump when assets change
const ASSETS = [
  './',
  './index.html',
  './app.js',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './ARMSLOGO.PNG'
];

// Install: pre-cache core assets
self.addEventListener('install', (e) => {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
});

// Activate: clean old caches
self.addEventListener('activate', (e) => {
  e.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)));
      await self.clients.claim();
    })()
  );
});

// Fetch: prefer cache for same-origin static; fall back to network, then cache
self.addEventListener('fetch', (e) => {
  const { request } = e;
  const url = new URL(request.url);

  // Only handle same-origin requests
  if (url.origin !== location.origin) return;

  // Strategy:
  // - HTML: Network first, cache fallback, update cache with fresh copy
  // - JS/CSS/Images: Cache first, then network; cache successful responses
  if (request.destination === 'document') {
    e.respondWith(
      (async () => {
        try {
          const network = await fetch(request);
          // Cache fresh HTML for offline
          const cache = await caches.open(CACHE_NAME);
          cache.put(request, network.clone());
          return network;
        } catch {
          const cached = await caches.match(request);
          return cached || caches.match('./index.html');
        }
      })()
    );
    return;
  }

  if (['script', 'style', 'image', 'font'].includes(request.destination)) {
    e.respondWith(
      (async () => {
        const cached = await caches.match(request);
        if (cached) return cached;
        try {
          const network = await fetch(request);
          // Only cache OK, same-origin, non-opaque responses
          if (network.ok && network.type === 'basic') {
            const cache = await caches.open(CACHE_NAME);
            cache.put(request, network.clone());
          }
          return network;
        } catch {
          // Last-resort: return whatever we have (maybe a precached icon)
          return caches.match('./icon-192.png') || Response.error();
        }
      })()
    );
    return;
  }

  // Default: just pass-through (e.g., POSTs to API, etc.)
});
