// Cache dell'app per funzionare anche senza connessione.
// Prima prova la rete (così gli aggiornamenti arrivano subito), se manca usa la copia salvata.
const CACHE = 'inventario-qr-v4';
const FILES = [
  './', './index.html', './manifest.json', './core.js',
  './lib/jsQR.js', './lib/xlsx.full.min.js',
  './icon-192.png', './icon-512.png', './icon-maskable.png', './apple-touch-icon.png'
];
// File per leggere le scritte dell'etichetta: grandi, si scaricano dopo, senza bloccare l'installazione
const OCR_FILES = [
  './lib/ocr/tesseract.min.js', './lib/ocr/worker.min.js', './lib/ocr/eng.traineddata.gz',
  './lib/ocr/tesseract-core-simd-lstm.wasm.js', './lib/ocr/tesseract-core-lstm.wasm.js'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
      .then(() => caches.open(CACHE)).then(c => c.addAll(OCR_FILES)).catch(() => {})
  );
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || !e.request.url.startsWith(self.location.origin)) return;
  e.respondWith(
    fetch(e.request)
      .then(res => {
        if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true })
        .then(hit => hit || caches.match('./index.html')))
  );
});
