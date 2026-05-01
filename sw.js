const CACHE = 'lex-v2';
const STATIC = [
  '/', '/index.html', '/app.css',
  '/dist/main.js', '/dist/pwaVersion.js',
  '/manifest.json'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(STATIC)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.hostname === 'api.anthropic.com' ||
      url.hostname.endsWith('microsoftonline.com') ||
      url.hostname.endsWith('graph.microsoft.com') ||
      url.hostname.endsWith('live.com') ||
      url.hostname.endsWith('jsdelivr.net')) return;
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(resp => {
        if (resp.ok && e.request.method === 'GET') {
          caches.open(CACHE).then(c => c.put(e.request, resp.clone()));
        }
        return resp;
      });
    })
  );
});
