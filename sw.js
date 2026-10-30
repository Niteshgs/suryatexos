// Suryatex OS service worker: fast shell, offline fallback. Never touches /api/*.
const V = 'sx-v68';
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

// Phone notifications (Attendance complete alert)
self.addEventListener('push', (e) => {
  let d = {}; try { d = e.data ? e.data.json() : {}; } catch (x) {}
  e.waitUntil(self.registration.showNotification(d.title || 'Suryatex OS', {
    body: d.body || '', icon: '/icon-192.png', badge: '/icon-192.png', tag: 'att-done', data: { url: d.url || '/' },
  }));
});
self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((list) => {
    for (const c of list) { if ('focus' in c) return c.focus(); }
    return self.clients.openWindow((e.notification.data && e.notification.data.url) || '/');
  }));
});
