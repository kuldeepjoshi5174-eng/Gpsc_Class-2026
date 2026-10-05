// Data (/api) is NEVER cached, so everything stays live. Only the page shell is kept for offline start.
const CACHE = 'gpsc-shell-v3';
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.add('/')).catch(() => {})); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.pathname.startsWith('/api/')) return;
  if (e.request.mode === 'navigate') {
    e.respondWith(fetch(e.request).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put('/', cp)); return r; })
      .catch(() => caches.match('/')));
  }
});
