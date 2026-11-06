// Naikkan versi ini setiap kali Anda mengubah kode secara besar (contoh: v3, v4)
const CACHE = 'ina-warehouse-v3';
const SHELL = [
  './',
  'index.html',
  'manifest.webmanifest',
  'logo.png',
  'logo-192.png',
  'logo-512.png'
];
const CDN = [
  'https://cdn.tailwindcss.com',
  'https://cdn.jsdelivr.net/npm/html5-qrcode@2.3.8/html5-qrcode.min.js'
];

self.addEventListener('install', (e) => {
  // PENTING: Langsung paksa service worker baru untuk aktif (skip waiting)
  self.skipWaiting();
  e.waitUntil((async () => {
    const c = await caches.open(CACHE);
    await Promise.all(SHELL.map((u) => c.add(u).catch((err) => console.warn('Cache warning:', u, err))));
    await Promise.all(CDN.map((u) =>
      fetch(u, { mode: 'no-cors' }).then((r) => c.put(u, r)).catch(() => { })
    ));
  })());
});

self.addEventListener('activate', (e) => {
  // Langsung ambil alih kontrol client
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

  // Stale-While-Revalidate untuk halaman utama. 
  // Akan selalu menampilkan halaman cepat, lalu mendownload pembaruan di latar belakang
  if (req.mode === 'navigate' || req.headers.get('accept').includes('text/html')) {
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
          return res;
        })
        .catch(() => caches.match(req))
    );
    return;
  }

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