// Ina Basic Store - Service Worker
// Naikkan CACHE_NAME setiap kali file shell berubah.
const CACHE_NAME = 'ina-warehouse-v6';

// File milik sendiri. Satu file hilang TIDAK boleh menggagalkan instalasi.
const SHELL = [
    './',
    './index.html',
    './penjahit.html',
    './manifest.webmanifest',
    './manifest-penjahit.webmanifest',
    './logo.png',
    './logo-192.png',
    './logo-512.png'
];
const CDN = [
    'https://cdn.tailwindcss.com',
    'https://cdn.jsdelivr.net/npm/html5-qrcode@2.3.8/html5-qrcode.min.js'
];

self.addEventListener('install', (event) => {
    // Tidak memanggil skipWaiting() otomatis. Versi baru menunggu sampai
    // halaman mengirim pesan SKIP_WAITING agar input yang sedang berjalan tidak terputus.
    event.waitUntil((async () => {
        const cache = await caches.open(CACHE_NAME);
        await Promise.all(SHELL.map((url) =>
            cache.add(url).catch((err) => console.warn('SW: gagal cache', url, err && err.message))
        ));
        await Promise.all(CDN.map((url) =>
            fetch(url, { mode: 'no-cors' }).then((res) => cache.put(url, res)).catch(() => { })
        ));
    })());
});

self.addEventListener('message', (event) => {
    if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil((async () => {
        const keys = await caches.keys();
        await Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)));
        await self.clients.claim();
    })());
});

self.addEventListener('fetch', (event) => {
    const req = event.request;
    if (req.method !== 'GET') return;

    const url = new URL(req.url);
    // Jangan pernah mencegat atau menyimpan respons Apps Script (data transaksi).
    if (url.hostname.endsWith('script.google.com') || url.hostname.endsWith('googleusercontent.com')) return;

    const accept = req.headers.get('accept') || '';
    const isPage = req.mode === 'navigate' || accept.indexOf('text/html') !== -1;

    // Pustaka CDN: cache dulu, ambil jaringan bila belum ada.
    if (url.origin !== self.location.origin) {
        event.respondWith(
            caches.match(req).then((hit) => hit || fetch(req).then((res) => {
                if (res && (res.ok || res.type === 'opaque')) {
                    const copy = res.clone();
                    caches.open(CACHE_NAME).then((c) => c.put(req, copy));
                }
                return res;
            }))
        );
        return;
    }

    // File milik sendiri: jaringan dulu, cache sebagai cadangan offline.
    event.respondWith(
        fetch(req)
            .then((res) => {
                if (res && res.status === 200 && res.type === 'basic') {
                    const copy = res.clone();
                    caches.open(CACHE_NAME).then((c) => c.put(req, copy));
                }
                return res;
            })
            .catch(() => caches.match(req).then((hit) => {
                if (hit) return hit;
                if (isPage) {
                    const fallback = url.pathname.includes('penjahit') ? './penjahit.html' : './index.html';
                    return caches.match(fallback);
                }
                return undefined;
            }))
    );
});
