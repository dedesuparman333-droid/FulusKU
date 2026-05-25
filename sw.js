// ── FulusKUpro Service Worker ─────────────────────────────────────
// Strategy:
//   - Aset lokal (HTML/CSS/font/icon) → Cache First
//   - Request ke Google Apps Script / googleapis → Network Only
//   - Fallback offline → tampilkan index.html dari cache
// ─────────────────────────────────────────────────────────────────

const CACHE_NAME = 'fulusku-v1';

// Aset yang di-cache saat install
const STATIC_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  // CDN fonts & icons (opsional — dihapus jika tidak mau cache CDN)
  'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap',
  'https://cdn.tailwindcss.com',
  'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css'
];

// Pola URL yang SELALU diarahkan ke network (jangan di-cache)
const NETWORK_ONLY_PATTERNS = [
  'script.google.com',   // Google Apps Script endpoint
  'googleapis.com',       // Google APIs lainnya
  'accounts.google.com'  // Google Auth
];

// ── INSTALL: cache semua static assets ───────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        // addAll akan gagal jika salah satu URL error — gunakan individual add untuk CDN
        return cache.addAll(['./index.html', './manifest.json', './icon-192.png', './icon-512.png', './'])
          .then(() => {
            // CDN assets: coba cache, tapi jangan gagalkan install jika tidak bisa
            return Promise.allSettled(
              STATIC_ASSETS.filter(url => url.startsWith('http'))
                .map(url => cache.add(url).catch(() => null))
            );
          });
      })
  );
  self.skipWaiting();
});

// ── ACTIVATE: hapus cache lama ────────────────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(key => key !== CACHE_NAME)
          .map(key => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

// ── FETCH: routing strategy ───────────────────────────────────────
self.addEventListener('fetch', event => {
  const url = event.request.url;

  // 1. Google Apps Script & APIs → Network Only (selalu butuh koneksi)
  const isNetworkOnly = NETWORK_ONLY_PATTERNS.some(pattern => url.includes(pattern));
  if (isNetworkOnly) {
    event.respondWith(
      fetch(event.request).catch(() => {
        // Jika offline dan request ke GAS gagal, kembalikan JSON error
        return new Response(
          JSON.stringify({
            success: false,
            data: null,
            message: 'Tidak ada koneksi internet. Silakan periksa jaringan Anda dan coba lagi.'
          }),
          {
            status: 503,
            headers: { 'Content-Type': 'application/json' }
          }
        );
      })
    );
    return;
  }

  // 2. Non-GET request → langsung ke network
  if (event.request.method !== 'GET') {
    event.respondWith(fetch(event.request));
    return;
  }

  // 3. Aset lokal & CDN → Cache First, fallback ke network lalu cache
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;

      return fetch(event.request)
        .then(networkResponse => {
          // Simpan ke cache jika response valid
          if (networkResponse && networkResponse.status === 200 && networkResponse.type !== 'opaque') {
            const cloned = networkResponse.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(event.request, cloned));
          }
          return networkResponse;
        })
        .catch(() => {
          // Fallback: jika request adalah navigasi halaman, kembalikan index.html
          if (event.request.mode === 'navigate') {
            return caches.match('./index.html');
          }
          return new Response('Offline', { status: 503 });
        });
    })
  );
});

// ── MESSAGE: handle skipWaiting dari client ───────────────────────
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
