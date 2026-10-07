/* Service worker : met l'application en cache pour un fonctionnement hors réseau.
 * Il ne voit passer aucune donnée de ronde (tout est en base locale). */
const CACHE = 'rondes-cachan-1.1.0';
const ASSETS = ['./', './index.html', './app.js', './manifest.webmanifest', './icon-192.png', './icon-512.png',
  './lib/fflate.js', './lib/qrcode.js', './lib/jsQR.js', './lib/xlsx.mini.min.js'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  if (req.mode === 'navigate') {
    e.respondWith(caches.match('./index.html').then(r => r || fetch(req)));
    return;
  }
  e.respondWith(caches.match(req, { ignoreSearch: true }).then(r => r || fetch(req)));
});
