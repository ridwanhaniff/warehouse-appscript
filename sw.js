// Service worker Ina Warehouse: cache "app shell" supaya aplikasi cepat terbuka.
// Request POST ke Google Apps Script TIDAK pernah disentuh/di-cache.
const CACHE = 'ina-warehouse-v1';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];
const CDN = [
  'https://cdn.tailwindcss.com',
  'https://cdn.jsdelivr.net/npm/html5-qrcode@2.3.8/html5-qrcode.min.js'
];

self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    await c.addAll(SHELL);
    await Promise.all(CDN.map((u) =>
      fetch(u, { mode: 'no-cors' }).then((r) => c.put(u, r)).catch(() => {})
    ));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const host = new URL(req.url).hostname;
  if (host.endsWith('script.google.com') || host.endsWith('googleusercontent.com')) return;

  // Halaman: ambil versi terbaru dari jaringan, jatuh ke cache bila offline
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put('index.html', copy));
          return res;
        })
        .catch(() => caches.match('index.html'))
    );
    return;
  }

  // Aset statis: cache dulu, lalu jaringan
  e.respondWith(
    caches.match(req).then((hit) => hit || fetch(req).then((res) => {
      if (res && (res.ok || res.type === 'opaque')) {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy));
      }
      return res;
    }))
  );
});
