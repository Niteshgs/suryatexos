// Suryatex OS service worker: fast shell, offline fallback. Never touches /api/*.
const V = 'sx-v32';
const SHELL = ['/', '/logo.svg', '/manifest.json', '/icon-192.png', '/icon-512.png', '/apple-touch-icon.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(V).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== V).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (u.origin !== location.origin || u.pathname.startsWith('/api/')) return;

  // Pages: try the network for up to 3s (so updates arrive), else the cached copy.
  if (r.mode === 'navigate') {
    e.respondWith(
      Promise.race([
        fetch(r).then((res) => {
          if (res.ok) { const c = res.clone(); caches.open(V).then((ch) => ch.put('/', c)); }
          return res;
        }),
        new Promise((_, rej) => setTimeout(() => rej(new Error('slow')), 3000)),
      ]).catch(() => caches.match('/').then((hit) => hit || fetch(r)))
    );
    return;
  }

  // Static assets: cached copy first, refreshed in the background.
  e.respondWith(
    caches.match(r).then((hit) => {
      const net = fetch(r).then((res) => {
        if (res.ok) { const c = res.clone(); caches.open(V).then((ch) => ch.put(r, c)); }
        return res;
      }).catch(() => hit);
      return hit || net;
    })
  );
});
