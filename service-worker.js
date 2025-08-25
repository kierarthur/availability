// sw.js
// PWA service worker for Availability app

// Bump this any time you change assets or caching logic
const CACHE_NAME = 'rota-avail-v3';

// Core assets to pre-cache for offline
const ASSETS = [
  './',
  './index.html',
  './app.js',
  './manifest.json',
  './APPICON-192.png',
  './APPICON-512.png',
  './ARMSLOGO.png',
];

// Utility: safe cache put (ignore opaque or error responses)
async function cachePutSafe(cache, request, response) {
  try {
    if (response && response.ok && response.type === 'basic') {
      await cache.put(request, response.clone());
    }
  } catch {
    // ignore put failures (e.g., quota)
  }
}

// --- Install: pre-cache core assets ---
self.addEventListener('install', (event) => {
  // Immediately activate the new SW on install
  self.skipWaiting();
  event.waitUntil(
    (async () => {
      const cache = await caches.open(CACHE_NAME);
      await cache.addAll(ASSETS);
    })()
  );
});

// --- Activate: clean old caches + enable navigation preload ---
self.addEventListener('activate', (event) => {
  event.waitUntil(
    (async () => {
      // Delete any old caches
      const keys = await caches.keys();
      await Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)));

      // Optional but nice: enable navigation preload when supported
      if ('navigationPreload' in self.registration) {
        try { await self.registration.navigationPreload.enable(); } catch {}
      }

      // Take control of all clients immediately
      await self.clients.claim();
    })()
  );
});

// Optional: allow the page to ask the SW to activate immediately after update
self.addEventListener('message', (event) => {
  if (!event || !event.data) return;
  if (event.data === 'SKIP_WAITING' || (event.data && event.data.type === 'SKIP_WAITING')) {
    self.skipWaiting();
  }
});

// --- Fetch handler ---
// Strategy:
// - For navigations / HTML: Network-first → cache fallback → (as last resort) cached index.html
// - For same-origin static assets (script/style/image/font): Cache-first → network; cache success
// - For everything else (e.g., POSTs, cross-origin, API): pass-through (do not cache)
self.addEventListener('fetch', (event) => {
  const { request } = event;
  const url = new URL(request.url);

  // Only handle same-origin requests; never touch cross-origin (e.g., your API on script.google.com)
  if (url.origin !== location.origin) return;

  // Handle HTML navigations (works across browsers)
  const isNavigation =
    request.mode === 'navigate' ||
    (request.destination === 'document') ||
    (request.headers.get('accept') || '').includes('text/html');

  if (isNavigation) {
    event.respondWith((async () => {
      // Try navigation preload (if enabled) for faster responses
      try {
        const preload = await event.preloadResponse;
        if (preload) {
          // Update cache in background
          caches.open(CACHE_NAME).then(cache => cachePutSafe(cache, request, preload.clone()));
          return preload;
        }
      } catch {}

      // Network first
      try {
        const network = await fetch(request);
        // Update cache with fresh HTML
        const cache = await caches.open(CACHE_NAME);
        cachePutSafe(cache, request, network.clone());
        return network;
      } catch {
        // Fallback to cached page, then to cached index
        const cache = await caches.open(CACHE_NAME);
        const cached = await cache.match(request);
        return cached || cache.match('./index.html');
      }
    })());
    return;
  }

  // For static assets, use cache-first
  if (['script', 'style', 'image', 'font'].includes(request.destination)) {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_NAME);
      const cached = await cache.match(request);
      if (cached) return cached;

      try {
        const network = await fetch(request);
        await cachePutSafe(cache, request, network.clone());
        return network;
      } catch {
        // Fallback to an app icon as a last resort if it's an image request
        if (request.destination === 'image') {
          return (await cache.match('./APPICON-192.png')) || Response.error();
        }
        return Response.error();
      }
    })());
    return;
  }

  // Default: let it go to network (e.g., POSTs, xhr/fetch to API, etc.)
  // We deliberately do not cache API requests to avoid auth token leakage / staleness
});
