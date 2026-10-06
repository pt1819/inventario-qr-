// Cache dell'app per funzionare anche senza connessione
const CACHE = 'inventario-qr-v2';
const FILES = [
  './', './index.html', './manifest.json', './core.js',
  './lib/jsQR.js', './lib/xlsx.full.min.js',
  './icon-192.png', './icon-512.png', './icon-maskable.png', './apple-touch-icon.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then(hit =>
      hit || fetch(e.request).catch(() => caches.match('./index.html'))
    )
  );
});
