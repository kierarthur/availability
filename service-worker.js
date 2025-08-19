const CACHE_NAME = 'rota-avail-v1';
const ASSETS = [
  './',
  './index.html',
  './app.js',
  './manifest.json'
  // add icons when you have them: './icon-192.png','./icon-512.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
});

self.addEventListener('fetch', (e) => {
  const { request } = e;
  // Only cache same-origin static assets; let API requests go to network.
  const url = new URL(request.url);
  if (url.origin === location.origin) {
    e.respondWith(
      caches.match(request).then(cached =>
        cached || fetch(request).then(resp => {
          if (request.method === 'GET' && resp.ok && request.headers.get('accept')?.includes('text/html')) {
            const clone = resp.clone();
            caches.open(CACHE_NAME).then(c => c.put(request, clone));
          }
          return resp;
        }).catch(() => cached)
      )
    );
  }
});
