/* Rondes parkings — Ville de Cachan, DPMS
 * Application web hors ligne (PWA). Aucune donnée n'est envoyée à un serveur :
 * tout est stocké sur le téléphone et transmis par fichier au superviseur.
 * L'horodatage et la position de chaque saisie sont chiffrés avec la clé publique
 * du superviseur ; seul le téléphone superviseur peut les lire.
 */
'use strict';

const APP_VERSION = '1.0.0';
const FMT_PAQUET = 'rondes-cachan-paquet';
const FMT_ETAT = 'rondes-cachan-etat';
const FMT_SAUVEGARDE = 'rondes-cachan-sauvegarde';
const FMT_CLE = 'rondes-cachan-cle';
const HISTO_JOURS = 92;          // conservation de l'historique lisible sur le téléphone agents
const PBKDF2_ITER = 600000;

/* ====================================================================== */
/* Utilitaires                                                            */
/* ====================================================================== */
const te = new TextEncoder(), td = new TextDecoder();

const b64u = {
  enc(buf) {
    const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  },
  dec(str) {
    const s = str.replace(/-/g, '+').replace(/_/g, '/');
    const bin = atob(s + '==='.slice((s.length + 3) % 4));
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
};

function uuid() { return crypto.randomUUID(); }
function randCode(n) {
  const A = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const r = crypto.getRandomValues(new Uint8Array(n));
  return Array.from(r, x => A[x % A.length]).join('');
}
function pad(n) { return String(n).padStart(2, '0'); }
function localDate(d = new Date()) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function fmtJour(j) { if (!j) return ''; const [y, m, d] = j.split('-'); return `${d}/${m}/${y}`; }
function fmtDT(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function fmtHeure(iso) { if (!iso) return ''; const d = new Date(iso); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; }
function fileStamp(d = new Date()) { return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`; }

function canon(v) {
  if (v === null || typeof v !== 'object') return JSON.stringify(v);
  if (Array.isArray(v)) return '[' + v.map(canon).join(',') + ']';
  return '{' + Object.keys(v).filter(k => v[k] !== undefined).sort()
    .map(k => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}';
}
async function sha256hex(data) {
  let buf;
  if (typeof data === 'string') buf = te.encode(data);
  else if (data instanceof Blob) buf = await data.arrayBuffer();
  else buf = data;
  const h = await crypto.subtle.digest('SHA-256', buf);
  return Array.from(new Uint8Array(h), b => b.toString(16).padStart(2, '0')).join('');
}
function blobToDataURL(blob) {
  return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(blob); });
}
function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 60000);
}
async function shareOrDownload(blob, name, title) {
  const file = new File([blob], name, { type: blob.type || 'application/octet-stream' });
  if (navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title }); return 'partage'; }
    catch (e) { if (e.name === 'AbortError') return 'annule'; }
  }
  download(blob, name);
  return 'telechargement';
}
function pickFile(accept) {
  return new Promise(res => {
    const i = document.createElement('input');
    i.type = 'file'; i.accept = accept;
    i.onchange = () => res(i.files[0] || null);
    i.click();
  });
}
function appBaseURL() { return location.origin + location.pathname.replace(/index\.html$/, ''); }

/* ====================================================================== */
/* Base de données locale (IndexedDB)                                     */
/* ====================================================================== */
const DB_NAME = 'rondes-cachan', DB_VER = 1;
let _db = null;
function openDB() {
  if (_db) return Promise.resolve(_db);
  return new Promise((res, rej) => {
    const r = indexedDB.open(DB_NAME, DB_VER);
    r.onupgradeneeded = () => {
      const d = r.result;
      d.createObjectStore('kv');
      d.createObjectStore('events', { keyPath: 'seq' });     // agent : journal chaîné
      d.createObjectStore('photos', { keyPath: 'id' });
      d.createObjectStore('sigs', { keyPath: 'id' });        // signalements
      d.createObjectStore('histo', { keyPath: 'k' });        // agent : historique lisible
      d.createObjectStore('sup_events', { keyPath: 'k' });   // superviseur : journal importé
      d.createObjectStore('rondes', { keyPath: 'rid' });     // superviseur : rondes reconstituées
    };
    r.onsuccess = () => { _db = r.result; res(_db); };
    r.onerror = () => rej(r.error);
  });
}
function reqP(r) { return new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); }
function txDone(t) { return new Promise((res, rej) => { t.oncomplete = () => res(); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error); }); }
async function dbGet(s, k) { const d = await openDB(); return reqP(d.transaction(s).objectStore(s).get(k)); }
async function dbAll(s) { const d = await openDB(); return reqP(d.transaction(s).objectStore(s).getAll()); }
async function dbPut(s, v, k) { const d = await openDB(); const t = d.transaction(s, 'readwrite'); k === undefined ? t.objectStore(s).put(v) : t.objectStore(s).put(v, k); return txDone(t); }
async function dbPutMany(s, vals) { const d = await openDB(); const t = d.transaction(s, 'readwrite'); const o = t.objectStore(s); vals.forEach(v => o.put(v)); return txDone(t); }
async function dbDel(s, k) { const d = await openDB(); const t = d.transaction(s, 'readwrite'); t.objectStore(s).delete(k); return txDone(t); }
async function dbClear(s) { const d = await openDB(); const t = d.transaction(s, 'readwrite'); t.objectStore(s).clear(); return txDone(t); }
const kvGet = k => dbGet('kv', k);
const kvSet = (k, v) => dbPut('kv', v, k);
const kvDel = k => dbDel('kv', k);

let _lock = Promise.resolve();
function withLock(fn) { const p = _lock.then(fn, fn); _lock = p.catch(() => { }); return p; }

/* ====================================================================== */
/* Cryptographie (WebCrypto)                                              */
/* ECDH P-256 éphémère + HKDF-SHA256 + AES-256-GCM (schéma de type ECIES) */
/* ====================================================================== */
const EC = { name: 'ECDH', namedCurve: 'P-256' };
const HKDF_INFO = te.encode('rondes-cachan v1');

async function hkdfKey(bits, salt, usages) {
  const base = await crypto.subtle.importKey('raw', bits, 'HKDF', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'HKDF', hash: 'SHA-256', salt, info: HKDF_INFO }, base, { name: 'AES-GCM', length: 256 }, false, usages);
}
let _pubKey = null, _pubJwkStr = null;
async function getPubKey() {
  const jwk = state.cfg.pub;
  const s = JSON.stringify(jwk);
  if (_pubKey && _pubJwkStr === s) return _pubKey;
  _pubKey = await crypto.subtle.importKey('jwk', jwk, EC, false, []);
  _pubJwkStr = s;
  return _pubKey;
}
async function seal(obj) {
  const pub = await getPubKey();
  const eph = await crypto.subtle.generateKey(EC, true, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'ECDH', public: pub }, eph.privateKey, 256);
  const epkRaw = new Uint8Array(await crypto.subtle.exportKey('raw', eph.publicKey));
  const key = await hkdfKey(bits, epkRaw, ['encrypt']);
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, te.encode(JSON.stringify(obj)));
  return { epk: b64u.enc(epkRaw), iv: b64u.enc(iv), ct: b64u.enc(ct) };
}
async function unseal(priv, sec) {
  const epkRaw = b64u.dec(sec.epk);
  const epk = await crypto.subtle.importKey('raw', epkRaw, EC, false, []);
  const bits = await crypto.subtle.deriveBits({ name: 'ECDH', public: epk }, priv, 256);
  const key = await hkdfKey(bits, epkRaw, ['decrypt']);
  const pt = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b64u.dec(sec.iv) }, key, b64u.dec(sec.ct));
  return JSON.parse(td.decode(pt));
}
async function passKey(pass, salt, usages) {
  const base = await crypto.subtle.importKey('raw', te.encode(pass), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', hash: 'SHA-256', salt, iterations: PBKDF2_ITER }, base, { name: 'AES-GCM', length: 256 }, false, usages);
}
async function encryptWithPass(pass, obj) {
  const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
  const key = await passKey(pass, salt, ['encrypt']);
  const ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, te.encode(JSON.stringify(obj)));
  return { kdf: 'PBKDF2-SHA256', iter: PBKDF2_ITER, salt: b64u.enc(salt), iv: b64u.enc(iv), ct: b64u.enc(ct) };
}
async function decryptWithPass(pass, box) {
  const key = await passKey(pass, b64u.dec(box.salt), ['decrypt']);
  const pt = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b64u.dec(box.iv) }, key, b64u.dec(box.ct));
  return JSON.parse(td.decode(pt));
}
async function getPrivKey() { return kvGet('privKey'); }

/* ====================================================================== */
/* Configuration par défaut                                               */
/* ====================================================================== */
function defaultConfig(pubJwk) {
  return {
    v: 1,
    cfgId: randCode(8),
    pub: pubJwk,
    sites: [
      { id: 'henouille', code: 'HEN', nom: 'Hénouille', token: randCode(10) },
      { id: 'dumotel', code: 'DUM', nom: 'Dumotel', token: randCode(10) },
      { id: 'arobase', code: 'ARO', nom: 'Arobase', token: randCode(10) }
    ],
    agents: ['Agent 1', 'Agent 2', 'Agent 3', 'Agent 4', 'Vacataire'],   // noms réels saisis dans « Listes et coordonnées », jamais dans le code publié
    dests: ['Ateliers', 'DST', 'Skidata', 'Police municipale', 'DPMS'],
    cats: ['Éclairage', 'Propreté', 'Dégradation / vandalisme', 'Barrière / caisse / borne',
      'Sécurité incendie / issues de secours', 'Véhicule (ventouse, épave, gênant)',
      'Présence suspecte / occupation', 'Fuite / infiltration', 'Autre'],
    checklist: [
      { lbl: 'Éclairage', cat: 'Éclairage' },
      { lbl: 'Barrières, caisses et bornes', cat: 'Barrière / caisse / borne' },
      { lbl: 'Propreté', cat: 'Propreté' },
      { lbl: 'Issues de secours et extincteurs', cat: 'Sécurité incendie / issues de secours' },
      { lbl: 'Accès piétons (portes, escaliers, ascenseur)', cat: 'Dégradation / vandalisme' },
      { lbl: 'Véhicules (ventouses, épaves)', cat: 'Véhicule (ventouse, épave, gênant)' }
    ],
    urgenceTel: '',
    urgenceMail: ''
  };
}
const VEHICULE_RE = /v[ée]hicule/i;
function siteById(id) { return state.cfg.sites.find(s => s.id === id); }
function siteNom(id) { const s = siteById(id); return s ? s.nom : id; }
function agentCfgPayload(cfg) {
  // Ce qui part vers le téléphone des agents (jamais la clé privée)
  return { v: cfg.v, cfgId: cfg.cfgId, pub: cfg.pub, sites: cfg.sites, agents: cfg.agents, dests: cfg.dests, cats: cfg.cats, checklist: cfg.checklist, urgenceTel: cfg.urgenceTel, urgenceMail: cfg.urgenceMail };
}
function packCfg(cfg) { return b64u.enc(fflate.deflateSync(te.encode(JSON.stringify(agentCfgPayload(cfg))), { level: 9 })); }
function unpackCfg(s) { return JSON.parse(td.decode(fflate.inflateSync(b64u.dec(s)))); }
function validCfg(c) {
  return c && c.pub && c.pub.kty === 'EC' && Array.isArray(c.sites) && Array.isArray(c.agents) && Array.isArray(c.dests) && Array.isArray(c.cats) && Array.isArray(c.checklist);
}

/* ====================================================================== */
/* Géolocalisation : uniquement pendant une ronde, jamais en continu hors */
/* ronde. Seule la position au moment de chaque saisie est enregistrée.   */
/* ====================================================================== */
const Geo = {
  watchId: null, last: null, err: null,
  start() {
    if (this.watchId !== null || !('geolocation' in navigator)) return;
    this.err = null;
    this.watchId = navigator.geolocation.watchPosition(
      p => { this.last = { lat: +p.coords.latitude.toFixed(6), lon: +p.coords.longitude.toFixed(6), acc: Math.round(p.coords.accuracy), ts: p.timestamp }; this.err = null; },
      e => { this.err = ({ 1: 'refusée', 2: 'indisponible', 3: 'délai dépassé' })[e.code] || 'erreur'; },
      { enableHighAccuracy: true, maximumAge: 30000, timeout: 30000 });
  },
  stop() { if (this.watchId !== null) navigator.geolocation.clearWatch(this.watchId); this.watchId = null; this.last = null; },
  async waitFirst(ms) { const t0 = Date.now(); while (!this.last && !this.err && Date.now() - t0 < ms) await new Promise(r => setTimeout(r, 200)); },
  snap() {
    if (!('geolocation' in navigator)) return { err: 'non prise en charge' };
    if (this.last) return { ...this.last, age: Math.round((Date.now() - this.last.ts) / 1000) };
    return { err: this.err || 'pas de position' };
  }
};

/* ====================================================================== */
/* État de l'application et rendu                                         */
/* ====================================================================== */
const state = { mode: null, cfg: null, view: 'home', params: {}, cleanup: null };

function h(tag, attrs, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'style') el.style.cssText = v;
    else if (k === 'html') el.innerHTML = v;
    else if (k === 'value') el.value = v;
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat(Infinity)) { if (c == null || c === false) continue; el.append(c instanceof Node ? c : String(c)); }
  return el;
}
function safe(fn) {
  return async (...a) => {
    try { await fn(...a); }
    catch (e) { console.error(e); hideBusy(); toast(e && e.message ? e.message : String(e), true); }
  };
}
let _toastT = null;
function toast(msg, bad) {
  document.querySelectorAll('.toast').forEach(t => t.remove());
  const t = h('div', { class: 'toast' + (bad ? ' bad' : ''), role: 'status' }, msg);
  document.body.append(t);
  clearTimeout(_toastT); _toastT = setTimeout(() => t.remove(), bad ? 6000 : 3500);
}
function showBusy(msg) { hideBusy(); document.body.append(h('div', { class: 'overlay', id: 'busy' }, msg || 'Patientez…')); }
function hideBusy() { const b = document.getElementById('busy'); if (b) b.remove(); }

function header(title, opts) {
  opts = opts || {};
  const tag = state.mode === 'agent' ? 'Téléphone agents' : state.mode === 'superviseur' ? 'Superviseur DPMS' : null;
  return h('header', { class: 'top' },
    opts.back ? h('button', { class: 'back', 'aria-label': 'Retour', onclick: () => opts.back === true ? go('home') : opts.back() }, '‹') : null,
    h('h1', null, title),
    tag ? h('span', { class: 'tag' }, tag) : null);
}
function page(title, opts, ...content) { return h('div', null, header(title, opts), h('main', null, ...content)); }

function go(view, params = {}, opts = {}) {
  if (state.cleanup) { try { state.cleanup(); } catch (e) { } state.cleanup = null; }
  state.view = view; state.params = params;
  if (!opts.noPush && view !== 'home') history.pushState({ v: view }, '');
  render();
  window.scrollTo(0, 0);
}
async function render() {
  const root = document.getElementById('app');
  const fn = VIEWS[state.view] || VIEWS.home;
  try {
    const node = await fn(state.params);
    root.replaceChildren(node);
  } catch (e) {
    console.error(e);
    root.replaceChildren(page('Erreur', { back: true }, h('div', { class: 'banner bad' }, String(e && e.message || e))));
  }
}
window.addEventListener('popstate', () => {
  if (state.view === 'signalement' && !confirm('Abandonner ce signalement ?')) { history.pushState({ v: state.view }, ''); return; }
  if (state.view !== 'home') go('home', {}, { noPush: true });
});

/* ====================================================================== */
/* Photos                                                                 */
/* ====================================================================== */
function toJpeg(bmp, max, q) {
  const s = Math.min(1, max / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas');
  c.width = Math.round(bmp.width * s); c.height = Math.round(bmp.height * s);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  return new Promise(r => c.toBlob(r, 'image/jpeg', q));
}
async function processPhoto(file) {
  // Ré-encodage : réduit la taille et supprime les métadonnées EXIF (dont le GPS du fichier)
  const bmp = await createImageBitmap(file);
  const full = await toJpeg(bmp, 1600, 0.72);
  const thumb = await toJpeg(bmp, 320, 0.6);
  if (bmp.close) bmp.close();
  const id = uuid();
  const sha = await sha256hex(full);
  const ageMin = file.lastModified ? Math.round((Date.now() - file.lastModified) / 60000) : null;
  return { id, blob: full, sha, thumb: await blobToDataURL(thumb), ageMin };
}
async function thumbFromBlob(blob) {
  const bmp = await createImageBitmap(blob);
  const t = await toJpeg(bmp, 320, 0.6);
  if (bmp.close) bmp.close();
  return blobToDataURL(t);
}
function photoInput(onFile) {
  return safe(async () => {
    const i = document.createElement('input');
    i.type = 'file'; i.accept = 'image/*'; i.setAttribute('capture', 'environment');
    i.onchange = safe(async () => { if (i.files[0]) await onFile(i.files[0]); });
    i.click();
  });
}

/* ====================================================================== */
/* AGENT : journal chaîné                                                 */
/* ====================================================================== */
async function appendEvent(type, base, data, secretExtra = {}, photos = []) {
  return withLock(async () => {
    const head = (await kvGet('head')) || { seq: 0, hash: '0'.repeat(64) };
    const now = new Date();
    const secret = { ts: now.toISOString(), tzo: now.getTimezoneOffset(), geo: Geo.snap(), ...secretExtra };
    const ev = {
      v: 1, dev: await kvGet('dev'), cfgId: state.cfg.cfgId, seq: head.seq + 1, t: type, jour: localDate(now),
      rid: base.rid || null, site: base.site || null, agent: base.agent || null,
      data, photos: photos.map(p => ({ id: p.id, sha: p.sha })),
      sec: await seal(secret), prev: head.hash
    };
    ev.hash = await sha256hex(canon(ev));
    const d = await openDB();
    const t = d.transaction(['events', 'kv'], 'readwrite');
    t.objectStore('events').put(ev);
    t.objectStore('kv').put({ seq: ev.seq, hash: ev.hash }, 'head');
    await txDone(t);
    return ev;
  });
}
async function addHisto(ev, row) { await dbPut('histo', { k: ev.seq, jour: ev.jour, site: ev.site, agent: ev.agent, ...row }); }
async function savePhotos(list) { await dbPutMany('photos', list.map(p => ({ id: p.id, blob: p.blob, sha: p.sha }))); }

const VERDICTS = { resolu: 'Résolu', encours: 'Toujours en cours', aggrave: 'Aggravé' };
const STATUTS = { ouvert: 'Ouvert', a_cloturer: 'Résolu selon l’agent — à clôturer', clos: 'Clos' };
function statutBadge(s) {
  const cls = s === 'ouvert' ? 'b-warn' : s === 'a_cloturer' ? 'b-info' : 'b-ok';
  return h('span', { class: 'badge ' + cls }, s === 'a_cloturer' ? 'À clôturer' : STATUTS[s] || s);
}
function applyVerdict(sig, v) {
  if (v.verdict === 'resolu' && sig.statut === 'ouvert') sig.statut = 'a_cloturer';
  if (v.verdict === 'aggrave') { sig.aggrave = true; if (sig.statut === 'a_cloturer') sig.statut = 'ouvert'; }
  if (v.verdict === 'encours' && sig.statut === 'a_cloturer') sig.statut = 'ouvert';
}

/* ====================================================================== */
/* VUES                                                                   */
/* ====================================================================== */
const VIEWS = {};

/* ---------- Accueil (aiguillage) ---------- */
VIEWS.home = async () => {
  if (!state.mode) return viewSetup();
  return state.mode === 'agent' ? agentHome() : supHome();
};

/* ---------- Première configuration ---------- */
async function viewSetup() {
  const pending = state.pendingCfg;
  if (pending) {
    return page('Configuration', null,
      h('div', { class: 'card' },
        h('h3', null, 'Configurer ce téléphone comme téléphone des agents ?'),
        h('p', { class: 'muted' }, `Parkings : ${pending.sites.map(s => s.nom).join(', ')}.`),
        h('p', { class: 'muted' }, 'Les saisies seront chiffrées pour le seul téléphone superviseur qui a produit ce code.'),
        h('button', { class: 'ok', onclick: safe(async () => { await setupAgent(pending); }) }, 'Oui, configurer'),
        h('button', { class: 'sec', onclick: () => { state.pendingCfg = null; render(); } }, 'Annuler')));
  }
  return page('Rondes parkings', null,
    h('div', { class: 'card' },
      h('h3', null, 'Téléphone des agents'),
      h('p', { class: 'muted' }, 'Sur le téléphone superviseur, ouvrez « QR codes » puis scannez le QR de configuration avec ce téléphone.'),
      h('button', { class: 'big', onclick: () => go('scan') }, 'Scanner le QR de configuration')),
    h('div', { class: 'card' },
      h('h3', null, 'Téléphone superviseur (DPMS)'),
      h('p', { class: 'muted' }, 'Crée la clé qui seule permet de lire l’horodatage et la position des saisies.'),
      h('button', { class: 'sec', onclick: () => go('supInit') }, 'Créer le superviseur'),
      h('button', { class: 'sec', onclick: () => go('supRestore') }, 'Restaurer une sauvegarde superviseur')),
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
}

async function setupAgent(cfg) {
  if (!validCfg(cfg)) throw new Error('Configuration invalide.');
  if (state.mode === 'superviseur') throw new Error('Ce téléphone est le superviseur.');
  if (state.mode === 'agent') {
    if (JSON.stringify(state.cfg.pub) !== JSON.stringify(cfg.pub)) throw new Error('Ce code provient d’un autre superviseur. Réinitialisez d’abord le téléphone.');
    await kvSet('cfg', cfg); state.cfg = cfg; state.pendingCfg = null;
    toast('Configuration mise à jour.'); go('home'); return;
  }
  await kvSet('mode', 'agent');
  await kvSet('cfg', cfg);
  if (!(await kvGet('dev'))) await kvSet('dev', 'T-' + randCode(6));
  state.mode = 'agent'; state.cfg = cfg; state.pendingCfg = null;
  await requestPersist();
  toast('Téléphone des agents configuré.');
  go('home');
}
async function requestPersist() {
  try { if (navigator.storage && navigator.storage.persist) return await navigator.storage.persist(); } catch (e) { }
  return false;
}

/* ---------- Scanner QR ---------- */
VIEWS.scan = async () => {
  const video = h('video', { playsinline: true, muted: true, autoplay: true });
  const msg = h('p', { class: 'muted' }, 'Visez le QR code.');
  const wrap = h('div', { class: 'scanner' }, video, h('div', { class: 'frame' }));
  let stream = null, stop = false, timer = null;
  state.cleanup = () => { stop = true; clearTimeout(timer); if (stream) stream.getTracks().forEach(t => t.stop()); };
  const node = page('Scanner', { back: true }, wrap, msg,
    state.mode === 'agent' ? h('button', { class: 'link', onclick: () => go('siteManuel') }, 'QR code illisible ? Démarrer sans QR') : null);
  (async () => {
    try {
      stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' }, audio: false });
      if (stop) { stream.getTracks().forEach(t => t.stop()); return; }
      video.srcObject = stream; await video.play();
      const detector = ('BarcodeDetector' in window) ? new BarcodeDetector({ formats: ['qr_code'] }) : null;
      const canvas = document.createElement('canvas'); const ctx = canvas.getContext('2d', { willReadFrequently: true });
      const tick = async () => {
        if (stop) return;
        let text = null;
        try {
          if (video.readyState >= 2) {
            if (detector) { const r = await detector.detect(video); if (r.length) text = r[0].rawValue; }
            else {
              const w = video.videoWidth, hh = video.videoHeight, s = Math.min(1, 800 / Math.max(w, hh));
              canvas.width = w * s; canvas.height = hh * s; ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
              const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
              const r = jsQR(img.data, img.width, img.height, { inversionAttempts: 'dontInvert' }); if (r) text = r.data;
            }
          }
        } catch (e) { }
        if (text) { state.cleanup(); state.cleanup = null; handleLink(text, true); return; }
        timer = setTimeout(tick, 220);
      };
      tick();
    } catch (e) {
      msg.textContent = 'Caméra indisponible : ' + (e.message || e.name) + '. Autorisez la caméra pour cette application.';
      msg.className = 'banner bad';
    }
  })();
  return node;
};

/* Traite un lien (QR scanné ou URL d'ouverture) */
function parseLink(text) {
  let hash = '';
  try { hash = new URL(text, appBaseURL()).hash; } catch (e) { hash = ''; }
  if (!hash && text.startsWith('#')) hash = text;
  const m = /^#(cfg|site)=(.+)$/.exec(hash);
  return m ? { kind: m[1], val: m[2] } : null;
}
async function handleLink(text, fromScan) {
  const l = parseLink(text);
  if (!l) { toast('QR code non reconnu.', true); if (fromScan) go('home'); return; }
  if (l.kind === 'cfg') {
    let cfg; try { cfg = unpackCfg(l.val); } catch (e) { toast('QR de configuration illisible.', true); go('home'); return; }
    if (state.mode === 'superviseur') { toast('Ce téléphone est le superviseur : configuration ignorée.', true); go('home'); return; }
    if (state.mode === 'agent') {
      if (confirm('Mettre à jour la configuration de ce téléphone ?')) await safe(setupAgent)(cfg); else go('home');
      return;
    }
    state.pendingCfg = cfg; go('home'); return;
  }
  if (l.kind === 'site') {
    if (state.mode !== 'agent') { toast('QR de parking : à scanner avec le téléphone des agents.', true); go('home'); return; }
    const [id, token] = l.val.split('.');
    const site = siteById(id);
    if (!site || site.token !== token) { toast('QR code de parking inconnu ou périmé.', true); go('home'); return; }
    return startRondeFlow(site.id, true);
  }
}

/* ====================================================================== */
/* AGENT                                                                  */
/* ====================================================================== */
async function agentHome() {
  const ronde = await kvGet('ronde');
  const sigs = await dbAll('sigs');
  const head = (await kvGet('head')) || { seq: 0 };
  const ack = (await kvGet('ack')) || 0;
  const lastExport = await kvGet('lastExport');
  const lastEtat = await kvGet('lastEtatImport');
  const nonTransmis = head.seq - ack;
  const cfg = state.cfg;

  const perSite = cfg.sites.map(s => {
    const o = sigs.filter(x => x.site === s.id && x.statut === 'ouvert').length;
    const a = sigs.filter(x => x.site === s.id && x.statut === 'a_cloturer').length;
    return h('div', { class: 'stat' }, h('span', null, s.nom), h('span', null,
      h('b', null, o), h('span', { class: 'muted small' }, o > 1 ? ' ouverts' : ' ouvert'),
      a ? h('span', { class: 'badge b-info' }, a + ' à clôturer') : null));
  });

  return page('Rondes parkings', null,
    ronde ? h('div', { class: 'card', style: 'border:2px solid var(--accent)' },
      h('h3', null, `Ronde en cours : ${siteNom(ronde.site)}`),
      h('p', { class: 'muted' }, `${ronde.agent} — commencée à ${fmtHeure(ronde.debut)}`),
      h('button', { class: 'accent big', onclick: () => resumeRonde(ronde) }, 'Reprendre la ronde'))
      : h('div', null,
        h('button', { class: 'big', onclick: () => go('scan') }, 'Scanner le QR du parking'),
        h('button', { class: 'link', onclick: () => go('siteManuel') }, 'QR code illisible ? Démarrer sans QR')),
    h('div', { class: 'card' }, h('h2', null, 'Signalements en cours'), ...perSite,
      h('button', { class: 'sec', onclick: () => go('agentSigs') }, 'Voir les signalements')),
    h('div', { class: 'card' }, h('h2', null, 'Transmission au DPMS'),
      nonTransmis > 0
        ? h('div', { class: 'banner warn' }, `${nonTransmis} saisie(s) non encore reçue(s) par le DPMS.`)
        : h('div', { class: 'banner ok' }, 'Tout a été reçu par le DPMS.'),
      lastExport ? h('p', { class: 'muted small' }, `Dernier envoi : ${fmtDT(lastExport.at)} (saisies ${lastExport.from} à ${lastExport.to}).`) : null,
      lastEtat ? h('p', { class: 'muted small' }, `Dernière mise à jour reçue du DPMS : ${fmtDT(lastEtat)}.`) : null,
      h('button', { disabled: nonTransmis <= 0 || !!ronde, onclick: safe(agentExport) }, 'Transmettre au DPMS'),
      ronde ? h('p', { class: 'muted small' }, 'Terminez la ronde avant de transmettre.') : null,
      h('button', { class: 'sec', onclick: safe(agentImportEtat) }, 'Recevoir la mise à jour du DPMS')),
    h('div', { class: 'card' }, h('h2', null, 'Extrait Excel'),
      h('div', { class: 'row' },
        h('button', { class: 'sec', onclick: safe(() => agentExcel(7)) }, '7 jours'),
        h('button', { class: 'sec', onclick: safe(() => agentExcel(31)) }, '31 jours'),
        h('button', { class: 'sec', onclick: safe(() => agentExcel(HISTO_JOURS)) }, 'Tout'))),
    (cfg.urgenceTel || cfg.urgenceMail) ? h('div', { class: 'card' }, h('h2', null, 'Urgence'),
      h('p', { class: 'muted small' }, 'Pour une urgence, prévenez directement. En cas de danger immédiat : 17 ou 112.'),
      cfg.urgenceTel ? h('a', { class: 'btn danger', href: 'tel:' + cfg.urgenceTel }, 'Appeler le DPMS') : null,
      cfg.urgenceMail ? h('a', { class: 'btn sec', href: 'mailto:' + cfg.urgenceMail + '?subject=' + encodeURIComponent('URGENT — parking') }, 'Écrire au DPMS') : null) : null,
    h('button', { class: 'link', onclick: () => go('reglages') }, 'Réglages'),
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
}

VIEWS.siteManuel = async () => page('Choisir le parking', { back: true },
  h('div', { class: 'banner warn' }, 'Le démarrage sans QR code est signalé au DPMS. Signalez le QR abîmé dans la ronde.'),
  ...state.cfg.sites.map(s => h('button', { class: 'big sec', onclick: () => startRondeFlow(s.id, false) }, s.nom)));

async function startRondeFlow(siteId, qr) {
  const cur = await kvGet('ronde');
  if (cur) {
    if (cur.site === siteId) { resumeRonde(cur); return; }
    if (!confirm(`Une ronde est en cours à ${siteNom(cur.site)}. La terminer et démarrer à ${siteNom(siteId)} ?`)) { go('home'); return; }
    await finishRonde(cur, true);
  }
  go('choixAgent', { site: siteId, qr });
}

VIEWS.choixAgent = async ({ site, qr }) => page(`Ronde : ${siteNom(site)}`, { back: true },
  h('p', null, 'Qui fait la ronde ?'),
  ...state.cfg.agents.map(a => h('button', { class: 'big sec', onclick: safe(() => beginRonde(site, a, qr)) }, a)));

async function beginRonde(site, agent, qr) {
  showBusy('Démarrage de la ronde…');
  Geo.start();
  await Geo.waitFirst(4000);
  const now = new Date();
  const ronde = { rid: uuid(), site, agent, qr, debut: now.toISOString(), step: 'revue', checklist: {}, sigsCrees: [] };
  await appendEvent('ronde_debut', ronde, { qr }, { qr });
  await kvSet('ronde', ronde);
  hideBusy();
  resumeRonde(ronde);
}
function resumeRonde(ronde) {
  Geo.start();
  if (ronde.step === 'revue') go('revue');
  else if (ronde.step === 'checklist') go('checklist');
  else go('ronde');
}

/* ---------- Revue des signalements ouverts ---------- */
const revueDraft = {};
VIEWS.revue = async () => {
  const ronde = await kvGet('ronde');
  if (!ronde) return agentHome();
  const open = (await dbAll('sigs')).filter(s => s.site === ronde.site && s.statut === 'ouvert')
    .sort((a, b) => (a.jour || '').localeCompare(b.jour || ''));
  if (!open.length) {
    ronde.step = 'checklist'; await kvSet('ronde', ronde);
    return VIEWS.checklist();
  }
  const btnValider = h('button', { class: 'ok' }, 'Valider la revue');
  const refresh = () => { btnValider.disabled = !open.every(s => revueDraft[s.id] && revueDraft[s.id].verdict); };
  const cards = open.map(s => {
    const d = revueDraft[s.id] || (revueDraft[s.id] = { verdict: null, comment: '', photos: [] });
    const thumbs = h('div', { class: 'thumbs' });
    const drawThumbs = () => thumbs.replaceChildren(...d.photos.map((p, i) =>
      h('div', { class: 't' }, h('img', { src: p.thumb, alt: '' }), h('button', { onclick: () => { d.photos.splice(i, 1); drawThumbs(); } }, '×'))));
    drawThumbs();
    const vbtns = Object.entries(VERDICTS).map(([k, lbl]) => h('button', {
      class: d.verdict === k ? 'sel-' + k : '',
      onclick: (e) => { d.verdict = k; [...e.target.parentNode.children].forEach(b => b.className = ''); e.target.className = 'sel-' + k; refresh(); }
    }, lbl));
    return h('div', { class: 'card' },
      h('h3', null, s.cat, s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('p', { class: 'muted small' }, `${s.ref} — signalé le ${fmtJour(s.jour)}${s.agent ? ' par ' + s.agent : ''}`),
      s.desc ? h('p', null, s.desc) : null,
      s.plaque ? h('p', null, h('b', null, 'Plaque : '), s.plaque) : null,
      s.thumb ? h('div', { class: 'thumbs' }, h('img', { src: s.thumb, alt: 'Photo du signalement' })) : null,
      s.notesAgents ? h('div', { class: 'banner info' }, 'DPMS : ' + s.notesAgents) : null,
      h('p', { style: 'margin:10px 0 6px;font-weight:600' }, 'Aujourd’hui :'),
      h('div', { class: 'verdicts' }, ...vbtns),
      h('input', { type: 'text', placeholder: 'Commentaire (facultatif)', value: d.comment, style: 'margin-top:8px', oninput: e => d.comment = e.target.value }),
      thumbs,
      h('button', { class: 'sec', onclick: photoInput(async f => { showBusy('Photo…'); d.photos.push(await processPhoto(f)); hideBusy(); drawThumbs(); }) }, 'Ajouter une photo'));
  });
  btnValider.onclick = safe(async () => {
    showBusy('Enregistrement…');
    for (const s of open) {
      const d = revueDraft[s.id];
      await savePhotos(d.photos);
      const ev = await appendEvent('revue', ronde, { sig: s.id, ref: s.ref, verdict: d.verdict, comment: d.comment.trim() },
        { photoAges: d.photos.map(p => ({ id: p.id, ageMin: p.ageMin })) }, d.photos);
      const sig = await dbGet('sigs', s.id);
      sig.suivi = sig.suivi || [];
      sig.suivi.push({ jour: ev.jour, agent: ronde.agent, verdict: d.verdict, comment: d.comment.trim() });
      applyVerdict(sig, d);
      await dbPut('sigs', sig);
      await addHisto(ev, { type: 'Constat de suivi', ref: s.ref, cat: s.cat, desc: s.desc, plaque: s.plaque, verdict: VERDICTS[d.verdict], comment: d.comment.trim() });
      delete revueDraft[s.id];
    }
    ronde.step = 'checklist'; await kvSet('ronde', ronde);
    hideBusy(); go('checklist');
  });
  refresh();
  return page(`Revue : ${siteNom(ronde.site)}`, null,
    h('div', { class: 'banner info' }, `${open.length} signalement(s) en cours sur ce parking. Indiquez leur état aujourd’hui.`),
    ...cards, h('div', { style: 'height:70px' }),
    h('div', { class: 'sticky-bottom' }, h('div', null, btnValider)));
};

/* ---------- Check-list ---------- */
VIEWS.checklist = async () => {
  const ronde = await kvGet('ronde');
  if (!ronde) return agentHome();
  const items = state.cfg.checklist;
  const btnValider = h('button', { class: 'ok' }, 'Valider la check-list');
  const refresh = () => { btnValider.disabled = !items.every(it => ronde.checklist[it.lbl]); };
  const rows = items.map(it => {
    const v = ronde.checklist[it.lbl];
    const bR = h('button', { class: v === 'RAS' ? 'sel-ras' : '' }, 'RAS');
    const bA = h('button', { class: v === 'Anomalie' ? 'sel-ano' : '' }, 'Anomalie');
    bR.onclick = safe(async () => { ronde.checklist[it.lbl] = 'RAS'; bR.className = 'sel-ras'; bA.className = ''; await kvSet('ronde', ronde); refresh(); });
    bA.onclick = safe(async () => {
      ronde.checklist[it.lbl] = 'Anomalie'; await kvSet('ronde', ronde);
      go('signalement', { cat: it.cat, from: 'checklist', item: it.lbl });
    });
    return h('div', { class: 'chk' }, h('span', { class: 'lbl' }, it.lbl), bR, bA);
  });
  btnValider.onclick = safe(async () => {
    await appendEvent('checklist', ronde, { items: ronde.checklist });
    ronde.step = 'ronde'; await kvSet('ronde', ronde);
    go('ronde');
  });
  refresh();
  return page(`Check-list : ${siteNom(ronde.site)}`, null,
    h('div', { class: 'card' }, ...rows),
    h('p', { class: 'muted small' }, '« Anomalie » ouvre directement un signalement.'),
    btnValider);
};

/* ---------- Ronde en cours ---------- */
VIEWS.ronde = async () => {
  const ronde = await kvGet('ronde');
  if (!ronde) return agentHome();
  const sigs = (await dbAll('sigs')).filter(s => ronde.sigsCrees.includes(s.id));
  return page(`Ronde : ${siteNom(ronde.site)}`, null,
    h('p', { class: 'muted' }, `${ronde.agent} — commencée à ${fmtHeure(ronde.debut)}`),
    h('button', { class: 'big accent', onclick: () => go('signalement', { from: 'ronde' }) }, '+ Nouveau signalement'),
    h('div', { class: 'card' }, h('h2', null, 'Signalements de cette ronde'),
      sigs.length ? sigs.map(s => h('div', { class: 'stat' }, h('span', null, s.cat, s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null), h('span', { class: 'muted small' }, s.ref)))
        : h('p', { class: 'muted' }, 'Aucun pour l’instant.')),
    h('button', { class: 'ok', onclick: safe(async () => { if (confirm('Terminer la ronde ?')) { await finishRonde(ronde); } }) }, 'Terminer la ronde'));
};
async function finishRonde(ronde, silent) {
  showBusy('Clôture de la ronde…');
  await appendEvent('ronde_fin', ronde, { nbSig: ronde.sigsCrees.length });
  await kvDel('ronde');
  Geo.stop();
  hideBusy();
  if (!silent) { toast('Ronde terminée et enregistrée.'); go('home'); }
}

/* ---------- Nouveau signalement ---------- */
let sigDraft = null;
VIEWS.signalement = async ({ cat, from, item }) => {
  const ronde = await kvGet('ronde');
  if (!ronde) return agentHome();
  const cfg = state.cfg;
  if (!sigDraft || sigDraft.item !== item) sigDraft = { cat: cat || null, desc: '', plaque: '', photos: [], dests: [], urgent: false, item };
  const d = sigDraft;
  const plaqueWrap = h('div', null, h('label', { class: 'f' }, 'Plaque d’immatriculation'),
    h('input', { type: 'text', value: d.plaque, placeholder: 'AB-123-CD', autocapitalize: 'characters', oninput: e => d.plaque = e.target.value }));
  const showPlaque = () => plaqueWrap.style.display = VEHICULE_RE.test(d.cat || '') ? '' : 'none';
  const catChips = h('div', { class: 'chips' }, ...cfg.cats.map(c => h('button', {
    class: 'chip' + (d.cat === c ? ' on' : ''),
    onclick: e => { d.cat = c; [...catChips.children].forEach(b => b.classList.remove('on')); e.target.classList.add('on'); showPlaque(); }
  }, c)));
  const destChips = h('div', { class: 'chips' }, ...cfg.dests.map(x => h('button', {
    class: 'chip' + (d.dests.includes(x) ? ' on' : ''),
    onclick: e => { const i = d.dests.indexOf(x); if (i >= 0) d.dests.splice(i, 1); else d.dests.push(x); e.target.classList.toggle('on'); }
  }, x)));
  const urg = h('button', { class: 'chip' + (d.urgent ? ' on-bad' : ''), onclick: e => { d.urgent = !d.urgent; e.target.className = 'chip' + (d.urgent ? ' on-bad' : ''); } }, 'Urgent');
  const thumbs = h('div', { class: 'thumbs' });
  const drawThumbs = () => thumbs.replaceChildren(...d.photos.map((p, i) =>
    h('div', { class: 't' }, h('img', { src: p.thumb, alt: '' }), h('button', { onclick: () => { d.photos.splice(i, 1); drawThumbs(); } }, '×'))));
  drawThumbs();
  showPlaque();
  const back = () => { if (!d.photos.length && !d.desc || confirm('Abandonner ce signalement ?')) { sigDraft = null; go(from === 'checklist' ? 'checklist' : 'ronde'); } };

  const save = safe(async () => {
    if (!d.cat) throw new Error('Choisissez une catégorie.');
    if (!d.desc.trim() && !d.photos.length) throw new Error('Ajoutez une description ou une photo.');
    if (!d.dests.length) throw new Error('Choisissez au moins un destinataire.');
    showBusy('Enregistrement…');
    const site = siteById(ronde.site);
    const sig = {
      id: uuid(), ref: `${site.code}-${randCode(5)}`, site: ronde.site, cat: d.cat, desc: d.desc.trim(),
      plaque: VEHICULE_RE.test(d.cat) ? d.plaque.trim().toUpperCase() : '', dests: [...d.dests], urgent: d.urgent,
      statut: 'ouvert', jour: localDate(), agent: ronde.agent, thumb: d.photos[0] ? d.photos[0].thumb : null, suivi: []
    };
    await savePhotos(d.photos);
    const ev = await appendEvent('signalement', ronde,
      { id: sig.id, ref: sig.ref, cat: sig.cat, desc: sig.desc, plaque: sig.plaque, dests: sig.dests, urgent: sig.urgent, checklist: item || null },
      { photoAges: d.photos.map(p => ({ id: p.id, ageMin: p.ageMin })) }, d.photos);
    await dbPut('sigs', sig);
    await addHisto(ev, { type: 'Signalement', ref: sig.ref, cat: sig.cat, desc: sig.desc, plaque: sig.plaque, dests: sig.dests.join(', '), urgent: sig.urgent ? 'Oui' : '' });
    ronde.sigsCrees.push(sig.id);
    await kvSet('ronde', ronde);
    sigDraft = null;
    hideBusy();
    toast(`Signalement ${sig.ref} enregistré.`);
    go(from === 'checklist' ? 'checklist' : 'ronde');
  });

  return page('Nouveau signalement', { back },
    item ? h('div', { class: 'banner warn' }, `Anomalie de la check-list : ${item}`) : null,
    h('label', { class: 'f' }, 'Catégorie'), catChips,
    plaqueWrap,
    h('label', { class: 'f' }, 'Description'),
    h('textarea', { placeholder: 'Ce qui ne va pas, où exactement (niveau, place, porte…). Le micro du clavier permet de dicter.', oninput: e => d.desc = e.target.value, value: d.desc }),
    h('label', { class: 'f' }, 'Photos'),
    thumbs,
    h('button', { class: 'sec', onclick: photoInput(async f => { if (d.photos.length >= 4) { toast('4 photos maximum.', true); return; } showBusy('Photo…'); d.photos.push(await processPhoto(f)); hideBusy(); drawThumbs(); }) }, 'Prendre une photo'),
    h('label', { class: 'f' }, 'À signaler à'), destChips,
    h('label', { class: 'f' }, 'Priorité'), h('div', { class: 'chips' }, urg),
    d.urgent && (state.cfg.urgenceTel) ? h('p', { class: 'muted small' }, 'Urgent : prévenez aussi par téléphone depuis l’accueil.') : null,
    h('div', { style: 'height:70px' }),
    h('div', { class: 'sticky-bottom' }, h('div', null, h('button', { class: 'ok', onclick: save }, 'Enregistrer le signalement'))));
};

/* ---------- Liste des signalements (agent) ---------- */
VIEWS.agentSigs = async () => {
  const sigs = (await dbAll('sigs')).filter(s => s.statut !== 'clos').sort((a, b) => (b.jour || '').localeCompare(a.jour || ''));
  return page('Signalements en cours', { back: true },
    sigs.length ? sigs.map(s => h('div', { class: 'card' },
      h('h3', null, s.cat, statutBadge(s.statut), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null),
      h('p', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ref} — ${fmtJour(s.jour)}`),
      s.desc ? h('p', null, s.desc) : null,
      s.plaque ? h('p', null, h('b', null, 'Plaque : '), s.plaque) : null,
      s.thumb ? h('div', { class: 'thumbs' }, h('img', { src: s.thumb, alt: '' })) : null,
      s.notesAgents ? h('div', { class: 'banner info' }, 'DPMS : ' + s.notesAgents) : null))
      : h('p', { class: 'muted' }, 'Aucun signalement en cours.'));
};

/* ---------- Transmission (agent → superviseur) ---------- */
async function agentExport() {
  const ack = (await kvGet('ack')) || 0;
  const events = (await dbAll('events')).filter(e => e.seq > ack).sort((a, b) => a.seq - b.seq);
  if (!events.length) { toast('Rien à transmettre.'); return; }
  showBusy('Préparation du fichier…');
  const files = {};
  const ids = new Set(events.flatMap(e => e.photos.map(p => p.id)));
  for (const id of ids) {
    const p = await dbGet('photos', id);
    if (p) files[`photos/${id}.jpg`] = [new Uint8Array(await p.blob.arrayBuffer()), { level: 0 }];
  }
  const dev = await kvGet('dev');
  const paquet = { format: FMT_PAQUET, v: 1, app: APP_VERSION, dev, cfgId: state.cfg.cfgId, cree: new Date().toISOString(), from: events[0].seq, to: events[events.length - 1].seq, events };
  files['paquet.json'] = te.encode(JSON.stringify(paquet));
  const zip = fflate.zipSync(files);
  const name = `rondes_${dev}_${fileStamp()}_${paquet.from}-${paquet.to}.zip`;
  await kvSet('lastExport', { at: new Date().toISOString(), from: paquet.from, to: paquet.to });
  hideBusy();
  const r = await shareOrDownload(new Blob([zip], { type: 'application/zip' }), name, 'Rondes parkings');
  if (r !== 'annule') toast('Fichier prêt. Il reste à l’envoyer au DPMS (Quick Share, courriel…).');
  render();
}

/* ---------- Mise à jour reçue du superviseur ---------- */
async function agentImportEtat() {
  const f = await pickFile('.json,application/json');
  if (!f) return;
  showBusy('Lecture de la mise à jour…');
  const etat = JSON.parse(await f.text());
  if (etat.format !== FMT_ETAT) throw new Error('Ce fichier n’est pas une mise à jour du DPMS.');
  if (!validCfg(etat.cfg)) throw new Error('Configuration du fichier invalide.');
  if (JSON.stringify(etat.cfg.pub) !== JSON.stringify(state.cfg.pub)) throw new Error('Fichier produit par un autre superviseur : refusé.');
  const dev = await kvGet('dev');
  const ack = Math.max((await kvGet('ack')) || 0, (etat.ack && etat.ack[dev]) || 0);
  await kvSet('cfg', etat.cfg); state.cfg = etat.cfg;
  await withLock(async () => {
    // 1. Purge des saisies reçues par le DPMS
    const events = await dbAll('events');
    const d = await openDB();
    const t = d.transaction(['events', 'photos', 'sigs', 'kv'], 'readwrite');
    const keepPhotos = new Set();
    for (const e of events) {
      if (e.seq <= ack) t.objectStore('events').delete(e.seq);
      else e.photos.forEach(p => keepPhotos.add(p.id));
    }
    // 2. État des signalements : le DPMS fait foi
    t.objectStore('sigs').clear();
    const sigs = new Map(etat.sigs.map(s => [s.id, s]));
    // 3. Réapplication des saisies locales non encore reçues
    for (const e of events.filter(e => e.seq > ack).sort((a, b) => a.seq - b.seq)) {
      if (e.t === 'signalement' && !sigs.has(e.data.id)) {
        const local = await dbGet('sigs', e.data.id);
        sigs.set(e.data.id, local || { id: e.data.id, ref: e.data.ref, site: e.site, cat: e.data.cat, desc: e.data.desc, plaque: e.data.plaque, dests: e.data.dests, urgent: e.data.urgent, statut: 'ouvert', jour: e.jour, agent: e.agent, suivi: [] });
      }
      if (e.t === 'revue' && sigs.has(e.data.sig)) {
        const s = sigs.get(e.data.sig);
        if (s.statut !== 'clos') { s.suivi = s.suivi || []; s.suivi.push({ jour: e.jour, agent: e.agent, verdict: e.data.verdict, comment: e.data.comment }); applyVerdict(s, e.data); }
      }
    }
    sigs.forEach(s => t.objectStore('sigs').put(s));
    t.objectStore('kv').put(ack, 'ack');
    t.objectStore('kv').put(new Date().toISOString(), 'lastEtatImport');
    await txDone(t);
    const photos = await dbAll('photos');
    const d2 = await openDB(); const t2 = d2.transaction('photos', 'readwrite');
    photos.forEach(p => { if (!keepPhotos.has(p.id)) t2.objectStore('photos').delete(p.id); });
    await txDone(t2);
  });
  hideBusy();
  toast('Mise à jour du DPMS appliquée.');
  go('home');
}

/* ---------- Extrait Excel (agent) : sans horodatage précis ni position ---------- */
async function agentExcel(jours) {
  const lim = localDate(new Date(Date.now() - jours * 86400000));
  const rows = (await dbAll('histo')).filter(r => r.jour >= lim).sort((a, b) => a.k - b.k);
  if (!rows.length) { toast('Aucune saisie sur la période.'); return; }
  const aoa = [['Date', 'Parking', 'Agent', 'Type', 'Réf.', 'Catégorie', 'Description', 'Plaque', 'Destinataires', 'Urgent', 'Constat', 'Commentaire']];
  rows.forEach(r => aoa.push([fmtJour(r.jour), siteNom(r.site), r.agent || '', r.type, r.ref || '', r.cat || '', r.desc || '', r.plaque || '', r.dests || '', r.urgent || '', r.verdict || '', r.comment || '']));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [10, 12, 12, 12, 11, 24, 50, 12, 28, 8, 16, 30].map(w => ({ wch: w }));
  ws['!autofilter'] = { ref: ws['!ref'] };
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Rondes');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  await shareOrDownload(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `extrait_rondes_${fileStamp()}.xlsx`, 'Extrait des rondes');
}

/* ---------- Réglages (commun) ---------- */
VIEWS.reglages = async () => {
  let persisted = null;
  try { persisted = navigator.storage && navigator.storage.persisted ? await navigator.storage.persisted() : null; } catch (e) { }
  let est = null; try { est = navigator.storage && navigator.storage.estimate ? await navigator.storage.estimate() : null; } catch (e) { }
  const dev = await kvGet('dev');
  return page('Réglages', { back: true },
    h('div', { class: 'card' }, h('dl', { class: 'kv' },
      h('dt', null, 'Mode'), h('dd', null, state.mode === 'agent' ? 'Téléphone agents' : 'Superviseur'),
      dev ? [h('dt', null, 'Identifiant'), h('dd', null, dev)] : null,
      state.cfg ? [h('dt', null, 'Configuration'), h('dd', null, state.cfg.cfgId)] : null,
      h('dt', null, 'Stockage protégé'), h('dd', null, persisted === true ? 'Oui' : persisted === false ? 'Non (installez l’application sur l’écran d’accueil)' : 'Inconnu'),
      est ? [h('dt', null, 'Espace utilisé'), h('dd', null, (est.usage / 1048576).toFixed(1) + ' Mo')] : null,
      h('dt', null, 'Version'), h('dd', null, APP_VERSION))),
    persisted === false ? h('button', { class: 'sec', onclick: safe(async () => { const ok = await requestPersist(); toast(ok ? 'Stockage protégé.' : 'Refusé par le navigateur.' , !ok); render(); }) }, 'Demander la protection du stockage') : null,
    state.mode === 'agent' ? h('button', { class: 'sec', onclick: () => go('scan') }, 'Scanner une nouvelle configuration') : null,
    h('hr'),
    h('div', { class: 'card' }, h('h2', null, 'Réinitialiser'),
      h('p', { class: 'muted small' }, state.mode === 'agent'
        ? 'Efface toutes les données de ce téléphone, y compris les saisies non transmises.'
        : 'Efface la clé et toutes les données. Sans sauvegarde, rien ne pourra être relu.'),
      h('button', {
        class: 'danger', onclick: safe(async () => {
          const r = prompt('Tapez EFFACER pour confirmer.');
          if (r !== 'EFFACER') return;
          Geo.stop();
          for (const s of ['kv', 'events', 'photos', 'sigs', 'histo', 'sup_events', 'rondes']) await dbClear(s);
          state.mode = null; state.cfg = null;
          toast('Téléphone réinitialisé.'); go('home');
        })
      }, 'Réinitialiser ce téléphone')));
};

/* ====================================================================== */
/* SUPERVISEUR                                                            */
/* ====================================================================== */
VIEWS.supInit = async () => {
  const p1 = h('input', { type: 'password', autocomplete: 'new-password' });
  const p2 = h('input', { type: 'password', autocomplete: 'new-password' });
  return page('Créer le superviseur', { back: true },
    h('div', { class: 'banner info' }, 'La phrase de passe protège la clé de secours. Elle sera demandée pour restaurer le superviseur sur un autre téléphone. Notez-la et remettez-la sous pli fermé à la DGS avec la clé de secours.'),
    h('label', { class: 'f' }, 'Phrase de passe (12 caractères minimum)'), p1,
    h('label', { class: 'f' }, 'Confirmation'), p2,
    h('button', {
      class: 'ok', onclick: safe(async () => {
        if (p1.value.length < 12) throw new Error('12 caractères minimum.');
        if (p1.value !== p2.value) throw new Error('Les deux saisies diffèrent.');
        showBusy('Création de la clé…');
        const kp = await crypto.subtle.generateKey(EC, true, ['deriveBits']);
        const privJwk = await crypto.subtle.exportKey('jwk', kp.privateKey);
        const pubFull = await crypto.subtle.exportKey('jwk', kp.publicKey);
        const pub = { kty: pubFull.kty, crv: pubFull.crv, x: pubFull.x, y: pubFull.y };
        const privEnc = await encryptWithPass(p1.value, privJwk);
        const priv = await crypto.subtle.importKey('jwk', privJwk, EC, false, ['deriveBits']);
        const cfg = defaultConfig(pub);
        await kvSet('privKey', priv); await kvSet('privEnc', privEnc);
        await kvSet('cfg', cfg); await kvSet('mode', 'superviseur'); await kvSet('devs', {});
        state.mode = 'superviseur'; state.cfg = cfg;
        await requestPersist();
        hideBusy();
        toast('Superviseur créé.');
        go('supCle', { first: true });
      })
    }, 'Créer la clé'));
};

VIEWS.supRestore = async () => page('Restaurer le superviseur', { back: true },
  h('p', { class: 'muted' }, 'Choisissez une sauvegarde complète (.zip) ou une clé de secours (.json).'),
  h('button', {
    onclick: safe(async () => {
      const f = await pickFile('.zip,.json');
      if (!f) return;
      const pass = prompt('Phrase de passe de la clé :');
      if (!pass) return;
      showBusy('Restauration…');
      let cle, db = null, photos = {};
      if (f.name.toLowerCase().endsWith('.zip')) {
        const z = fflate.unzipSync(new Uint8Array(await f.arrayBuffer()));
        const dump = JSON.parse(td.decode(z['sauvegarde.json']));
        if (dump.format !== FMT_SAUVEGARDE) throw new Error('Sauvegarde invalide.');
        cle = dump.cle; db = dump;
        for (const [k, v] of Object.entries(z)) if (k.startsWith('photos/')) photos[k.slice(7, -4)] = v;
      } else {
        cle = JSON.parse(await f.text());
        if (cle.format !== FMT_CLE) throw new Error('Clé de secours invalide.');
      }
      let privJwk;
      try { privJwk = await decryptWithPass(pass, cle.privEnc); } catch (e) { throw new Error('Phrase de passe incorrecte.'); }
      const priv = await crypto.subtle.importKey('jwk', privJwk, EC, false, ['deriveBits']);
      await kvSet('privKey', priv); await kvSet('privEnc', cle.privEnc);
      const cfg = (db && db.cfg) || cle.cfg;
      await kvSet('cfg', cfg); await kvSet('mode', 'superviseur'); await kvSet('devs', (db && db.devs) || {});
      if (db) {
        await dbPutMany('sigs', db.sigs); await dbPutMany('sup_events', db.events); await dbPutMany('rondes', db.rondes);
        const shaMap = {};
        db.events.forEach(e => (e.photos || []).forEach(p => shaMap[p.id] = p.sha));
        await dbPutMany('photos', Object.entries(photos).map(([id, bytes]) => ({ id, blob: new Blob([bytes], { type: 'image/jpeg' }), sha: shaMap[id] || null })));
      }
      state.mode = 'superviseur'; state.cfg = cfg;
      await requestPersist();
      hideBusy(); toast('Superviseur restauré.'); go('home');
    })
  }, 'Choisir le fichier'));

async function supHome() {
  const sigs = await dbAll('sigs');
  const devs = (await kvGet('devs')) || {};
  const lastImport = await kvGet('lastImport');
  const lastBackup = await kvGet('lastBackup');
  const n = st => sigs.filter(s => s.statut === st).length;
  const urg = sigs.filter(s => s.statut === 'ouvert' && s.urgent).length;
  const backupOld = !lastBackup || (Date.now() - new Date(lastBackup).getTime()) > 7 * 86400000;
  return page('Rondes parkings', null,
    backupOld ? h('div', { class: 'banner warn' }, lastBackup ? `Dernière sauvegarde le ${fmtDT(lastBackup)}. Faites-en une et versez-la sur un stockage de la Ville.` : 'Aucune sauvegarde : faites-en une et versez-la sur un stockage de la Ville.') : null,
    h('button', { class: 'big', onclick: safe(supImport) }, 'Importer une transmission'),
    lastImport ? h('p', { class: 'muted small' }, `Dernier import : ${fmtDT(lastImport)}.`) : null,
    h('div', { class: 'card' }, h('h2', null, 'Signalements'),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'ouvert' }) }, h('div', { class: 'stat' }, h('span', null, 'Ouverts', urg ? h('span', { class: 'badge b-bad' }, urg + ' urgent(s)') : null), h('b', null, n('ouvert')))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'a_cloturer' }) }, h('div', { class: 'stat' }, h('span', null, 'Résolus selon l’agent, à clôturer'), h('b', null, n('a_cloturer')))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'clos' }) }, h('div', { class: 'stat' }, h('span', null, 'Clos'), h('b', null, n('clos'))))),
    h('div', { class: 'card' }, h('h2', null, 'Retour vers le téléphone des agents'),
      h('p', { class: 'muted small' }, 'Après chaque import : envoie vos clôtures et accuse réception, ce qui libère la mémoire du téléphone.'),
      h('button', { onclick: safe(supExportEtat) }, 'Préparer la mise à jour')),
    h('div', { class: 'card' }, h('h2', null, 'Exports et sauvegarde'),
      h('button', { class: 'sec', onclick: safe(supExcel) }, 'Excel complet (heures et positions)'),
      h('button', { class: 'sec', onclick: safe(supBackup) }, 'Sauvegarde complète (.zip)'),
      h('button', { class: 'sec', onclick: () => go('supCle') }, 'Clé de secours')),
    h('div', { class: 'card' }, h('h2', null, 'Paramétrage'),
      h('button', { class: 'sec', onclick: () => go('supQR') }, 'QR codes (parkings et configuration)'),
      h('button', { class: 'sec', onclick: () => go('supConfig') }, 'Listes et coordonnées'),
      h('button', { class: 'sec', onclick: () => go('supRondes') }, 'Historique des rondes')),
    Object.keys(devs).length ? h('div', { class: 'card' }, h('h2', null, 'Téléphones connus'),
      ...Object.entries(devs).map(([d, v]) => h('div', { class: 'stat' }, h('span', null, d), h('span', { class: 'muted small' }, `saisie n° ${v.seq} — ${fmtDT(v.at)}`)))) : null,
    h('button', { class: 'link', onclick: () => go('reglages') }, 'Réglages'),
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
}

/* ---------- Import d'une transmission ---------- */
async function supImport() {
  const f = await pickFile('.zip,application/zip');
  if (!f) return;
  showBusy('Vérification et déchiffrement…');
  const z = fflate.unzipSync(new Uint8Array(await f.arrayBuffer()));
  if (!z['paquet.json']) throw new Error('Ce fichier n’est pas une transmission de rondes.');
  const paquet = JSON.parse(td.decode(z['paquet.json']));
  if (paquet.format !== FMT_PAQUET) throw new Error('Format de transmission inconnu.');
  const priv = await getPrivKey();
  const devs = (await kvGet('devs')) || {};
  const dv = devs[paquet.dev] || { seq: 0, hash: '0'.repeat(64) };
  const rapport = { importes: 0, dejaRecus: 0, nouveauxSigs: 0, revues: 0, rondes: 0, alertes: [] };
  const evs = [...paquet.events].sort((a, b) => a.seq - b.seq);
  let expSeq = dv.seq + 1, expPrev = dv.hash;
  const sigsMap = new Map((await dbAll('sigs')).map(s => [s.id, s]));
  const rondesMap = new Map((await dbAll('rondes')).map(r => [r.rid, r]));
  const toStore = [], photosToStore = [];

  for (const ev of evs) {
    const k = `${paquet.dev}:${String(ev.seq).padStart(8, '0')}`;
    const integ = [];
    const { hash, ...rest } = ev;
    if (await sha256hex(canon(rest)) !== hash) integ.push('empreinte invalide (saisie modifiée)');
    if (ev.seq <= dv.seq) {
      const prevStored = await dbGet('sup_events', k);
      if (prevStored && prevStored.hash === hash) { rapport.dejaRecus++; continue; }
      integ.push('saisie déjà reçue avec un contenu différent');
    } else {
      if (ev.seq !== expSeq) integ.push(`rupture : saisies ${expSeq} à ${ev.seq - 1} manquantes`);
      else if (ev.prev !== expPrev) integ.push('chaînage rompu avec la saisie précédente');
      expSeq = ev.seq + 1; expPrev = hash;
    }
    let dec = null;
    try { dec = await unseal(priv, ev.sec); } catch (e) { integ.push('données chiffrées illisibles'); }
    for (const p of ev.photos || []) {
      const bytes = z[`photos/${p.id}.jpg`];
      if (!bytes) { integ.push(`photo ${p.id.slice(0, 8)} absente`); continue; }
      if (await sha256hex(bytes) !== p.sha) integ.push(`photo ${p.id.slice(0, 8)} modifiée`);
      photosToStore.push({ id: p.id, blob: new Blob([bytes], { type: 'image/jpeg' }), sha: p.sha });
    }
    if (dec) {
      if (dec.qr === false) integ.push('ronde démarrée sans QR code');
      (dec.photoAges || []).forEach(a => { if (a.ageMin != null && a.ageMin > 15) integ.push(`photo ${a.id.slice(0, 8)} prise ${a.ageMin} min avant la saisie`); });
      if (dec.geo && dec.geo.err) integ.push('position : ' + dec.geo.err);
    }
    integ.forEach(m => rapport.alertes.push(`n° ${ev.seq} (${ev.t}) : ${m}`));
    toStore.push({ k, ...ev, dec, integ });
    rapport.importes++;

    // Reconstitution des signalements et des rondes
    if (ev.t === 'signalement') {
      if (!sigsMap.has(ev.data.id)) {
        sigsMap.set(ev.data.id, {
          id: ev.data.id, ref: ev.data.ref, site: ev.site, cat: ev.data.cat, desc: ev.data.desc, plaque: ev.data.plaque,
          dests: ev.data.dests, urgent: ev.data.urgent, statut: 'ouvert', jour: ev.jour, agent: ev.agent,
          photos: (ev.photos || []).map(p => p.id), ts: dec && dec.ts, geo: dec && dec.geo, dev: paquet.dev, seq: ev.seq,
          suivi: [], journalDPMS: [], notes: '', notesAgents: ''
        });
        rapport.nouveauxSigs++;
      }
    } else if (ev.t === 'revue') {
      const s = sigsMap.get(ev.data.sig);
      if (s) {
        s.suivi = s.suivi || [];
        s.suivi.push({ jour: ev.jour, ts: dec && dec.ts, agent: ev.agent, verdict: ev.data.verdict, comment: ev.data.comment, photos: (ev.photos || []).map(p => p.id), geo: dec && dec.geo });
        if (s.statut !== 'clos') applyVerdict(s, ev.data);
        rapport.revues++;
      } else rapport.alertes.push(`n° ${ev.seq} : revue d’un signalement inconnu (${ev.data.ref})`);
    }
    if (ev.rid) {
      const r = rondesMap.get(ev.rid) || { rid: ev.rid, site: ev.site, agent: ev.agent, dev: paquet.dev, jour: ev.jour, nbSig: 0, nbRevue: 0, alertes: [] };
      if (ev.t === 'ronde_debut') { r.debut = dec && dec.ts; r.qr = ev.data.qr; r.geoDebut = dec && dec.geo; rapport.rondes++; }
      if (ev.t === 'ronde_fin') { r.fin = dec && dec.ts; }
      if (ev.t === 'checklist') r.checklist = ev.data.items;
      if (ev.t === 'signalement') r.nbSig++;
      if (ev.t === 'revue') r.nbRevue++;
      r.alertes.push(...integ);
      rondesMap.set(ev.rid, r);
    }
  }
  // Enregistrement
  await dbPutMany('photos', photosToStore);
  await dbPutMany('sup_events', toStore);
  await dbPutMany('sigs', [...sigsMap.values()]);
  await dbPutMany('rondes', [...rondesMap.values()]);
  const last = toStore.filter(e => e.seq > dv.seq).pop();
  if (last) devs[paquet.dev] = { seq: last.seq, hash: last.hash, at: new Date().toISOString() };
  await kvSet('devs', devs);
  await kvSet('lastImport', new Date().toISOString());
  hideBusy();
  go('supRapport', { rapport, nom: f.name });
}
VIEWS.supRapport = async ({ rapport, nom }) => page('Import', { back: true },
  h('div', { class: 'banner ' + (rapport.alertes.length ? 'warn' : 'ok') },
    rapport.alertes.length ? `${rapport.alertes.length} point(s) d’attention.` : 'Transmission intègre : chaînage, empreintes et photos vérifiés.'),
  h('div', { class: 'card' }, h('h2', null, nom),
    h('div', { class: 'stat' }, h('span', null, 'Saisies importées'), h('b', null, rapport.importes)),
    h('div', { class: 'stat' }, h('span', null, 'Déjà reçues (ignorées)'), h('b', null, rapport.dejaRecus)),
    h('div', { class: 'stat' }, h('span', null, 'Rondes'), h('b', null, rapport.rondes)),
    h('div', { class: 'stat' }, h('span', null, 'Nouveaux signalements'), h('b', null, rapport.nouveauxSigs)),
    h('div', { class: 'stat' }, h('span', null, 'Constats de suivi'), h('b', null, rapport.revues))),
  rapport.alertes.length ? h('div', { class: 'card' }, h('h2', null, 'Points d’attention'), ...rapport.alertes.map(a => h('p', { class: 'small' }, '• ' + a))) : null,
  h('button', { onclick: safe(supExportEtat) }, 'Préparer la mise à jour du téléphone agents'),
  h('button', { class: 'sec', onclick: () => go('supSigs', { f: 'ouvert' }) }, 'Voir les signalements'));

/* ---------- Liste et fiche des signalements ---------- */
VIEWS.supSigs = async ({ f = 'ouvert', site = '' }) => {
  const all = await dbAll('sigs');
  const list = all.filter(s => (f === 'tous' || s.statut === f) && (!site || s.site === site))
    .sort((a, b) => (b.urgent - a.urgent) || (b.ts || b.jour || '').localeCompare(a.ts || a.jour || ''));
  const fchips = [['ouvert', 'Ouverts'], ['a_cloturer', 'À clôturer'], ['clos', 'Clos'], ['tous', 'Tous']].map(([k, l]) =>
    h('button', { class: 'chip' + (f === k ? ' on' : ''), onclick: () => go('supSigs', { f: k, site }, { noPush: true }) }, l));
  const schips = [['', 'Tous parkings'], ...state.cfg.sites.map(s => [s.id, s.nom])].map(([k, l]) =>
    h('button', { class: 'chip' + (site === k ? ' on' : ''), onclick: () => go('supSigs', { f, site: k }, { noPush: true }) }, l));
  return page('Signalements', { back: true },
    h('div', { class: 'chips' }, ...fchips), h('div', { class: 'chips', style: 'margin-bottom:12px' }, ...schips),
    list.length ? list.map(s => h('button', { class: 'list-item', onclick: () => go('supSig', { id: s.id }) },
      h('div', { class: 'l1' }, `${s.ref} · ${s.cat}`, statutBadge(s.statut), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('div', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ts ? fmtDT(s.ts) : fmtJour(s.jour)} — ${s.agent || ''} — ${(s.dests || []).join(', ')}`),
      s.desc ? h('div', { class: 'small' }, s.desc.length > 120 ? s.desc.slice(0, 120) + '…' : s.desc) : null))
      : h('p', { class: 'muted' }, 'Aucun signalement.'));
};
function geoLink(geo) {
  if (!geo || geo.err || geo.lat == null) return h('span', { class: 'muted' }, geo && geo.err ? 'position ' + geo.err : '—');
  return h('a', { href: `https://www.openstreetmap.org/?mlat=${geo.lat}&mlon=${geo.lon}#map=19/${geo.lat}/${geo.lon}`, target: '_blank', rel: 'noopener' },
    `${geo.lat}, ${geo.lon} (± ${geo.acc} m)`);
}
async function photoImgs(ids) {
  const out = [];
  for (const id of ids || []) {
    const p = await dbGet('photos', id);
    if (p) out.push(h('img', { class: 'photo-full', src: URL.createObjectURL(p.blob), alt: 'Photo' }));
  }
  return out;
}
VIEWS.supSig = async ({ id }) => {
  const s = await dbGet('sigs', id);
  if (!s) return page('Signalement', { back: true }, h('p', null, 'Introuvable.'));
  const notes = h('textarea', { value: s.notes || '', placeholder: 'Suites données, transmissions, interventions (interne DPMS)' });
  const notesAg = h('input', { type: 'text', value: s.notesAgents || '', placeholder: 'Ex. : intervention Ateliers prévue mardi' });
  const saveWith = async (statut, action) => {
    s.notes = notes.value; s.notesAgents = notesAg.value;
    if (statut) s.statut = statut;
    s.journalDPMS = s.journalDPMS || [];
    s.journalDPMS.push({ ts: new Date().toISOString(), action, note: notes.value });
    if (statut === 'clos') s.closLe = new Date().toISOString();
    if (statut === 'ouvert') { s.closLe = null; }
    await dbPut('sigs', s);
    toast('Enregistré.');
    render();
  };
  const suivi = [];
  for (const v of s.suivi || []) {
    suivi.push(h('div', { class: 'card' },
      h('div', null, h('b', null, VERDICTS[v.verdict] || v.verdict), ` — ${v.ts ? fmtDT(v.ts) : fmtJour(v.jour)} — ${v.agent || ''}`),
      v.comment ? h('p', null, v.comment) : null,
      h('div', { class: 'small' }, 'Position : ', geoLink(v.geo)),
      ...(await photoImgs(v.photos))));
  }
  return page(`Signalement ${s.ref}`, { back: () => go('supSigs', { f: s.statut }) },
    h('div', { class: 'card' },
      h('h3', null, s.cat, statutBadge(s.statut), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('dl', { class: 'kv' },
        h('dt', null, 'Parking'), h('dd', null, siteNom(s.site)),
        h('dt', null, 'Signalé'), h('dd', null, s.ts ? fmtDT(s.ts) : fmtJour(s.jour), s.agent ? ' par ' + s.agent : ''),
        h('dt', null, 'Position'), h('dd', null, geoLink(s.geo)),
        h('dt', null, 'À signaler à'), h('dd', null, (s.dests || []).join(', ')),
        s.plaque ? [h('dt', null, 'Plaque'), h('dd', null, s.plaque)] : null,
        s.closLe ? [h('dt', null, 'Clos le'), h('dd', null, fmtDT(s.closLe))] : null),
      s.desc ? h('p', null, s.desc) : null,
      ...(await photoImgs(s.photos))),
    suivi.length ? h('h2', { style: 'font-size:16px;color:var(--navy)' }, 'Constats de suivi') : null, ...suivi,
    h('div', { class: 'card' }, h('h2', null, 'Suites DPMS'),
      h('label', { class: 'f' }, 'Notes internes'), notes,
      h('label', { class: 'f' }, 'Message affiché aux agents'), notesAg,
      h('button', { class: 'sec', onclick: safe(() => saveWith(null, 'note')) }, 'Enregistrer les notes'),
      s.statut !== 'clos' ? h('button', { class: 'ok', onclick: safe(() => saveWith('clos', 'clôture')) }, 'Clôturer') : null,
      s.statut !== 'ouvert' ? h('button', { class: 'sec', onclick: safe(() => saveWith('ouvert', 'réouverture')) }, 'Remettre en cours') : null,
      s.plaque && s.statut === 'clos' ? h('button', { class: 'link', onclick: safe(async () => { if (confirm('Effacer la plaque de ce signalement clos ?')) { s.plaque = ''; await dbPut('sigs', s); toast('Plaque effacée.'); render(); } }) }, 'Effacer la plaque (dossier traité)') : null),
    (s.journalDPMS || []).length ? h('div', { class: 'card' }, h('h2', null, 'Historique DPMS'),
      ...s.journalDPMS.map(j => h('p', { class: 'small' }, `${fmtDT(j.ts)} — ${j.action}`))) : null);
};

/* ---------- Historique des rondes ---------- */
VIEWS.supRondes = async () => {
  const rs = (await dbAll('rondes')).sort((a, b) => (b.debut || '').localeCompare(a.debut || '')).slice(0, 200);
  return page('Rondes', { back: true },
    rs.length ? rs.map(r => {
      const dur = r.debut && r.fin ? Math.round((new Date(r.fin) - new Date(r.debut)) / 60000) : null;
      const anos = r.checklist ? Object.entries(r.checklist).filter(([, v]) => v === 'Anomalie').map(([k]) => k) : [];
      return h('div', { class: 'card' },
        h('div', { class: 'l1' }, h('b', null, `${siteNom(r.site)} — ${r.agent}`), r.qr === false ? h('span', { class: 'badge b-warn' }, 'sans QR') : null, !r.fin ? h('span', { class: 'badge b-warn' }, 'fin non reçue') : null),
        h('div', { class: 'muted small' }, `${fmtDT(r.debut)}${r.fin ? ' → ' + fmtHeure(r.fin) : ''}${dur != null ? ` (${dur} min)` : ''} — ${r.nbSig} signalement(s), ${r.nbRevue} constat(s)`),
        anos.length ? h('div', { class: 'small' }, 'Anomalies : ' + anos.join(', ')) : null,
        h('div', { class: 'small' }, 'Position au départ : ', geoLink(r.geoDebut)));
    }) : h('p', { class: 'muted' }, 'Aucune ronde importée.'));
};

/* ---------- Mise à jour pour le téléphone agents ---------- */
async function supExportEtat() {
  showBusy('Préparation…');
  const sigs = (await dbAll('sigs')).filter(s => s.statut !== 'clos');
  const out = [];
  for (const s of sigs) {
    let thumb = s.thumb || null;
    if (!thumb && s.photos && s.photos[0]) {
      const p = await dbGet('photos', s.photos[0]);
      if (p) { thumb = await thumbFromBlob(p.blob); s.thumb = thumb; await dbPut('sigs', s); }
    }
    out.push({ id: s.id, ref: s.ref, site: s.site, cat: s.cat, desc: s.desc, plaque: s.plaque, dests: s.dests, urgent: s.urgent, aggrave: !!s.aggrave, statut: s.statut, jour: s.jour, agent: s.agent, thumb, notesAgents: s.notesAgents || '', suivi: (s.suivi || []).map(v => ({ jour: v.jour, agent: v.agent, verdict: v.verdict, comment: v.comment })) });
  }
  const devs = (await kvGet('devs')) || {};
  const ack = Object.fromEntries(Object.entries(devs).map(([d, v]) => [d, v.seq]));
  const etat = { format: FMT_ETAT, v: 1, cree: new Date().toISOString(), cfg: agentCfgPayload(state.cfg), ack, sigs: out };
  hideBusy();
  await shareOrDownload(new Blob([JSON.stringify(etat)], { type: 'application/json' }), `maj_rondes_${fileStamp()}.json`, 'Mise à jour rondes');
  toast('À ouvrir sur le téléphone agents : « Recevoir la mise à jour du DPMS ».');
}

/* ---------- Excel complet (superviseur) ---------- */
async function supExcel() {
  showBusy('Construction du classeur…');
  const sigs = (await dbAll('sigs')).sort((a, b) => (a.ts || a.jour || '').localeCompare(b.ts || b.jour || ''));
  const rondes = (await dbAll('rondes')).sort((a, b) => (a.debut || '').localeCompare(b.debut || ''));
  const evs = (await dbAll('sup_events')).sort((a, b) => a.k.localeCompare(b.k));
  const g = (geo, k) => geo && !geo.err && geo[k] != null ? geo[k] : '';
  const A1 = [['Réf.', 'Parking', 'Catégorie', 'Description', 'Plaque', 'Destinataires', 'Urgent', 'Aggravé', 'Statut', 'Signalé le', 'Agent', 'Latitude', 'Longitude', 'Précision (m)', 'Photos', 'Dernier constat', 'Date dernier constat', 'Message aux agents', 'Notes DPMS', 'Clos le']];
  sigs.forEach(s => {
    const last = (s.suivi || []).slice(-1)[0];
    A1.push([s.ref, siteNom(s.site), s.cat, s.desc, s.plaque || '', (s.dests || []).join(', '), s.urgent ? 'Oui' : '', s.aggrave ? 'Oui' : '', STATUTS[s.statut], s.ts ? fmtDT(s.ts) : fmtJour(s.jour), s.agent || '',
      g(s.geo, 'lat'), g(s.geo, 'lon'), g(s.geo, 'acc'), (s.photos || []).length, last ? VERDICTS[last.verdict] : '', last ? (last.ts ? fmtDT(last.ts) : fmtJour(last.jour)) : '', s.notesAgents || '', s.notes || '', s.closLe ? fmtDT(s.closLe) : '']);
  });
  const A2 = [['Parking', 'Agent', 'Début', 'Fin', 'Durée (min)', 'Démarrage par QR', 'Anomalies check-list', 'Signalements', 'Constats', 'Lat. départ', 'Lon. départ', 'Précision (m)', 'Points d’attention', 'Téléphone']];
  rondes.forEach(r => {
    const dur = r.debut && r.fin ? Math.round((new Date(r.fin) - new Date(r.debut)) / 60000) : '';
    const anos = r.checklist ? Object.entries(r.checklist).filter(([, v]) => v === 'Anomalie').map(([k]) => k).join(', ') : '';
    A2.push([siteNom(r.site), r.agent, fmtDT(r.debut), fmtDT(r.fin), dur, r.qr === false ? 'Non' : 'Oui', anos, r.nbSig, r.nbRevue, g(r.geoDebut, 'lat'), g(r.geoDebut, 'lon'), g(r.geoDebut, 'acc'), [...new Set(r.alertes || [])].join(' ; '), r.dev]);
  });
  const TYPES = { ronde_debut: 'Début de ronde', ronde_fin: 'Fin de ronde', checklist: 'Check-list', signalement: 'Signalement', revue: 'Constat de suivi' };
  const A3 = [['Téléphone', 'N°', 'Type', 'Jour déclaré', 'Horodatage appareil', 'Parking', 'Agent', 'Détail', 'Latitude', 'Longitude', 'Précision (m)', 'Âge position (s)', 'Intégrité']];
  evs.forEach(e => {
    let det = '';
    if (e.t === 'signalement') det = `${e.data.ref} ${e.data.cat}${e.data.desc ? ' — ' + e.data.desc : ''}`;
    if (e.t === 'revue') det = `${e.data.ref} : ${VERDICTS[e.data.verdict]}${e.data.comment ? ' — ' + e.data.comment : ''}`;
    if (e.t === 'checklist') det = Object.entries(e.data.items || {}).map(([k, v]) => `${k} : ${v}`).join(' ; ');
    if (e.t === 'ronde_debut') det = e.data.qr ? 'QR scanné' : 'sans QR';
    if (e.t === 'ronde_fin') det = `${e.data.nbSig} signalement(s)`;
    const geo = e.dec && e.dec.geo;
    A3.push([e.dev, e.seq, TYPES[e.t] || e.t, fmtJour(e.jour), e.dec ? fmtDT(e.dec.ts) : 'illisible', siteNom(e.site), e.agent || '', det, g(geo, 'lat'), g(geo, 'lon'), g(geo, 'acc'), g(geo, 'age'), (e.integ || []).length ? e.integ.join(' ; ') : 'OK']);
  });
  const wb = XLSX.utils.book_new();
  const add = (aoa, name, widths) => { const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = widths.map(w => ({ wch: w })); ws['!autofilter'] = { ref: ws['!ref'] }; XLSX.utils.book_append_sheet(wb, ws, name); };
  add(A1, 'Signalements', [11, 12, 24, 50, 12, 28, 8, 8, 22, 16, 12, 11, 11, 10, 7, 16, 16, 30, 40, 16]);
  add(A2, 'Rondes', [12, 12, 16, 16, 10, 10, 30, 12, 10, 11, 11, 10, 40, 10]);
  add(A3, 'Journal', [10, 6, 16, 11, 16, 12, 12, 50, 11, 11, 10, 10, 40]);
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  hideBusy();
  await shareOrDownload(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `rondes_parkings_complet_${fileStamp()}.xlsx`, 'Rondes parkings');
}

/* ---------- Sauvegarde et clé de secours ---------- */
async function supBackup() {
  showBusy('Sauvegarde…');
  const dump = {
    format: FMT_SAUVEGARDE, v: 1, app: APP_VERSION, cree: new Date().toISOString(),
    cfg: state.cfg, devs: (await kvGet('devs')) || {},
    cle: { format: FMT_CLE, privEnc: await kvGet('privEnc') },
    sigs: await dbAll('sigs'), events: await dbAll('sup_events'), rondes: await dbAll('rondes')
  };
  const files = { 'sauvegarde.json': te.encode(JSON.stringify(dump)) };
  for (const p of await dbAll('photos')) files[`photos/${p.id}.jpg`] = [new Uint8Array(await p.blob.arrayBuffer()), { level: 0 }];
  const zip = fflate.zipSync(files);
  await kvSet('lastBackup', new Date().toISOString());
  hideBusy();
  await shareOrDownload(new Blob([zip], { type: 'application/zip' }), `sauvegarde_rondes_${fileStamp()}.zip`, 'Sauvegarde rondes');
  toast('Versez ce fichier sur un stockage de la Ville (OneDrive professionnel, lecteur réseau).');
}
VIEWS.supCle = async ({ first } = {}) => page('Clé de secours', { back: true },
  first ? h('div', { class: 'banner ok' }, 'Superviseur créé. Étape suivante : la clé de secours.') : null,
  h('div', { class: 'card' },
    h('p', null, 'Sans cette clé, si ce téléphone est perdu ou réinitialisé, aucune donnée transmise ne pourra plus être lue.'),
    h('p', { class: 'muted' }, 'Le fichier est chiffré par votre phrase de passe. Conservez-le sur un stockage de la Ville et remettez la phrase de passe sous pli fermé à la DGS.'),
    h('button', {
      onclick: safe(async () => {
        const cle = { format: FMT_CLE, v: 1, cree: new Date().toISOString(), privEnc: await kvGet('privEnc'), cfg: state.cfg };
        await shareOrDownload(new Blob([JSON.stringify(cle)], { type: 'application/json' }), 'cle_secours_rondes.json', 'Clé de secours rondes');
      })
    }, 'Télécharger la clé de secours')),
  first ? h('button', { class: 'sec', onclick: () => go('supQR') }, 'Continuer : QR codes') : null);

/* ---------- QR codes ---------- */
function qrSvg(text, ecc = 'M') {
  const q = qrcode(0, ecc); q.addData(text, 'Byte'); q.make();
  return q.createSvgTag({ cellSize: 4, margin: 4, scalable: true });
}
VIEWS.supQR = async () => {
  const cfgURL = appBaseURL() + '#cfg=' + packCfg(state.cfg);
  const show = h('div', { class: 'qr' });
  return page('QR codes', { back: true },
    h('div', { class: 'card' }, h('h2', null, 'Configuration du téléphone agents'),
      h('p', { class: 'muted small' }, 'À scanner avec l’appareil photo du téléphone agents, ou depuis son écran de configuration. À refaire après chaque modification des listes, ou envoyez une mise à jour.'),
      h('button', { onclick: () => { show.innerHTML = qrSvg(cfgURL, 'L'); } }, 'Afficher le QR de configuration'), show),
    h('div', { class: 'card' }, h('h2', null, 'Parkings'),
      h('p', { class: 'muted small' }, 'Un QR par parking, à afficher à l’entrée, plastifié, hors de portée. Scanné à l’arrivée, il démarre la ronde.'),
      ...state.cfg.sites.map(s => h('div', { class: 'stat' }, h('span', null, s.nom), h('span', { class: 'muted small' }, 'code ' + s.token.slice(0, 4) + '…'))),
      h('button', { class: 'accent', onclick: () => printSiteQR() }, 'Imprimer les QR des parkings'),
      h('button', {
        class: 'link', onclick: safe(async () => {
          if (!confirm('Générer de nouveaux codes ? Les QR affichés dans les parkings ne fonctionneront plus.')) return;
          state.cfg.sites.forEach(s => s.token = randCode(10)); state.cfg.cfgId = randCode(8);
          await kvSet('cfg', state.cfg); toast('Nouveaux codes générés : réimprimez et mettez à jour le téléphone agents.'); render();
        })
      }, 'Régénérer les codes (QR perdu ou copié)')));
};
function printSiteQR() {
  const area = document.getElementById('print-area');
  area.replaceChildren(...state.cfg.sites.map(s => {
    const url = appBaseURL() + `#site=${s.id}.${s.token}`;
    return h('div', { class: 'print-page' },
      h('p', null, 'VILLE DE CACHAN — Rondes de surveillance'),
      h('h1', null, 'Parking ' + s.nom),
      h('div', { html: qrSvg(url, 'M') }),
      h('p', null, 'Agents : scannez ce code à votre arrivée pour démarrer la ronde.'));
  }));
  setTimeout(() => window.print(), 100);
}

/* ---------- Listes et coordonnées ---------- */
VIEWS.supConfig = async () => {
  const c = state.cfg;
  const ta = (lines) => h('textarea', { value: lines.join('\n'), style: 'min-height:140px' });
  const tAgents = ta(c.agents), tDests = ta(c.dests), tCats = ta(c.cats);
  const tChk = ta(c.checklist.map(x => `${x.lbl} | ${x.cat}`));
  const sitesInputs = c.sites.map(s => h('input', { type: 'text', value: s.nom }));
  const tel = h('input', { type: 'tel', value: c.urgenceTel || '', placeholder: '01 …' });
  const mail = h('input', { type: 'email', value: c.urgenceMail || '', placeholder: 'prenom.nom@ville-cachan.fr' });
  const lines = t => t.value.split('\n').map(x => x.trim()).filter(Boolean);
  return page('Listes et coordonnées', { back: true },
    h('div', { class: 'banner info' }, 'Après modification : renvoyez une mise à jour au téléphone agents.'),
    h('label', { class: 'f' }, 'Noms des parkings'), ...sitesInputs,
    h('label', { class: 'f' }, 'Agents (un par ligne)'), tAgents,
    h('label', { class: 'f' }, 'Destinataires (un par ligne)'), tDests,
    h('label', { class: 'f' }, 'Catégories (une par ligne)'), tCats,
    h('label', { class: 'f' }, 'Check-list (« libellé | catégorie » par ligne)'), tChk,
    h('label', { class: 'f' }, 'Téléphone d’urgence affiché aux agents'), tel,
    h('label', { class: 'f' }, 'Courriel d’urgence'), mail,
    h('button', {
      class: 'ok', onclick: safe(async () => {
        const cats = lines(tCats);
        const chk = lines(tChk).map(l => { const [lbl, cat] = l.split('|').map(x => x.trim()); return { lbl, cat: cats.includes(cat) ? cat : (cats[cats.length - 1] || 'Autre') }; });
        if (!lines(tAgents).length || !lines(tDests).length || !cats.length) throw new Error('Listes vides.');
        sitesInputs.forEach((i, k) => { if (i.value.trim()) c.sites[k].nom = i.value.trim(); });
        Object.assign(c, { agents: lines(tAgents), dests: lines(tDests), cats, checklist: chk, urgenceTel: tel.value.trim(), urgenceMail: mail.value.trim(), cfgId: randCode(8) });
        await kvSet('cfg', c);
        toast('Enregistré. Pensez à mettre à jour le téléphone agents.');
        go('home');
      })
    }, 'Enregistrer'));
};

/* ====================================================================== */
/* Démarrage                                                              */
/* ====================================================================== */
async function boot() {
  if (!window.isSecureContext || !crypto.subtle) {
    document.getElementById('app').replaceChildren(page('Rondes parkings', null,
      h('div', { class: 'banner bad' }, 'Cette application doit être ouverte depuis son adresse https:// (pas depuis un fichier).')));
    return;
  }
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(e => console.warn('SW', e));
  state.mode = (await kvGet('mode')) || null;
  state.cfg = (await kvGet('cfg')) || null;
  if (state.mode === 'agent' && (await kvGet('ronde'))) Geo.start();
  const hash = location.hash;
  if (hash && /^#(cfg|site)=/.test(hash)) {
    history.replaceState(null, '', location.pathname + location.search);
    await handleLink(hash, false);
    if (state.view === 'home') render();
    return;
  }
  render();
}
window.addEventListener('hashchange', () => {
  const hash = location.hash;
  if (hash && /^#(cfg|site)=/.test(hash)) { history.replaceState(null, '', location.pathname + location.search); handleLink(hash, false); }
});
boot();
