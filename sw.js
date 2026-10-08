/* Service worker — Rondes parkings
 * - Programme (page, app.js) : toujours la dernière version publiée dès qu'il y a du réseau ;
 *   la copie locale ne sert qu'hors réseau (ou si le réseau ne répond pas en 4 s).
 * - Bibliothèques et icônes : copie locale (elles changent avec le numéro de version ci-dessous).
 * - À l'installation d'une nouvelle version, les pages ouvertes sont rechargées automatiquement.
 * Il ne voit passer aucune donnée de ronde (tout est en base locale ou chiffré vers le relais).
 */
const CACHE = 'rondes-cachan-1.5.0';
const SHELL = ['./', './index.html', './app.js', './manifest.webmanifest'];
const STATIC = ['./icon-192.png', './icon-512.png', './lib/fflate.js', './lib/qrcode.js', './lib/jsQR.js', './lib/xlsx.mini.min.js', './lib/chart.umd.min.js'];
const NET_TIMEOUT = 4000;

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE)
    .then(c => c.addAll([...SHELL, ...STATIC].map(u => new Request(u, { cache: 'no-cache' }))))
    .then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    const anciens = keys.filter(k => k.startsWith('rondes-cachan') && k !== CACHE);
    await Promise.all(anciens.map(k => caches.delete(k)));
    await self.clients.claim();
    if (anciens.length) {   // mise à jour : recharger les pages ouvertes avec l'ancienne version
      const wins = await self.clients.matchAll({ type: 'window' });
      wins.forEach(w => { try { w.navigate(w.url); } catch (err) { } });
    }
  })());
});

function estProgramme(url, req) {
  if (req.mode === 'navigate') return true;
  const p = url.pathname;
  return p.endsWith('/') || p.endsWith('/index.html') || p.endsWith('/app.js') || p.endsWith('/manifest.webmanifest');
}
async function reseauDabord(req, cleCache) {
  const cache = await caches.open(CACHE);
  const reseau = fetch(req, { cache: 'no-cache' }).then(r => {
    if (r && r.ok) cache.put(cleCache, r.clone());
    return r;
  });
  const delai = new Promise(res => setTimeout(() => res(null), NET_TIMEOUT));
  try {
    const r = await Promise.race([reseau, delai]);
    if (r && r.ok) return r;
  } catch (err) { /* hors réseau */ }
  const copie = await cache.match(cleCache, { ignoreSearch: true });
  if (copie) return copie;
  return reseau;   // pas de copie locale : attendre le réseau
}
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (estProgramme(url, req)) {
    e.respondWith(reseauDabord(req, req.mode === 'navigate' ? './index.html' : req));
    return;
  }
  e.respondWith(caches.match(req, { ignoreSearch: true }).then(r => r || fetch(req)));
});
