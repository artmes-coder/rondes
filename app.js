/* Rondes parkings — Ville de Cachan, DPMS — version 1.2
 * Application web hors ligne (PWA).
 * - L'outil appartient aux agents : ils signalent, suivent et clôturent.
 * - Le relais (Google Apps Script) conserve un journal chiffré partagé. Chaque appareil
 *   y publie ses saisies et reconstitue l'état commun à partir de ce journal :
 *   plusieurs téléphones agents et plusieurs appareils superviseurs sont possibles,
 *   et un appareil réinstallé retrouve toutes les données.
 * - Clé d'équipe (téléphones agents + superviseurs) : saisies, photos, interventions.
 *   Clé superviseur (superviseurs seuls) : horodatage, position, notes internes.
 *   Le relais ne peut rien lire.
 */
'use strict';

const APP_VERSION = '1.2.0';
const HISTO_JOURS = 92;            // période maximale de l'extrait Excel des agents
const PBKDF2_ITER = 600000;
const LOT_MAX_OCTETS = 4000000;    // taille maximale d'un envoi (photos comprises)
const SYNC_PERIODE_MS = 120000;

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
function fmtQuand(iso) { if (!iso) return ''; return localDate(new Date(iso)) === localDate() ? 'aujourd’hui à ' + fmtHeure(iso) : 'le ' + fmtDT(iso); }
function fileStamp(d = new Date()) { return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`; }
function canon(v) {
  if (v === null || typeof v !== 'object') return JSON.stringify(v);
  if (Array.isArray(v)) return '[' + v.map(canon).join(',') + ']';
  return '{' + Object.keys(v).filter(k => v[k] !== undefined).sort().map(k => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}';
}
async function sha256hex(data) {
  let buf;
  if (typeof data === 'string') buf = te.encode(data);
  else if (data instanceof Blob) buf = await data.arrayBuffer();
  else buf = data;
  const h = await crypto.subtle.digest('SHA-256', buf);
  return Array.from(new Uint8Array(h), b => b.toString(16).padStart(2, '0')).join('');
}
function blobToDataURL(blob) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(blob); }); }
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
function appBaseURL() { return location.origin + location.pathname.replace(/index\.html$/, ''); }

/* ====================================================================== */
/* Base de données locale (IndexedDB)                                     */
/* ====================================================================== */
const DB_NAME = 'rondes-cachan', DB_VER = 2;
const STORES = {
  kv: null,
  events: { keyPath: 'seq' },     // agents : saisies pas encore publiées sur le relais
  photos: { keyPath: 'id' },      // agents : photos pas encore publiées ; superviseurs : photos téléchargées
  entries: { keyPath: 'k' },      // journal relevé sur le relais (déchiffré)
  thumbs: { keyPath: 'id' },      // vignettes des signalements
  sigs: { keyPath: 'id' },        // état reconstitué des signalements
  rondes: { keyPath: 'rid' }      // superviseurs : rondes reconstituées
};
let _db = null;
function openDB() {
  if (_db) return Promise.resolve(_db);
  return new Promise((res, rej) => {
    const r = indexedDB.open(DB_NAME, DB_VER);
    r.onupgradeneeded = () => {
      const d = r.result;
      // Version 1.2 : nouveau format de données ; les données de test des versions antérieures sont effacées.
      [...d.objectStoreNames].forEach(n => d.deleteObjectStore(n));
      for (const [n, o] of Object.entries(STORES)) o ? d.createObjectStore(n, o) : d.createObjectStore(n);
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
async function dbReplaceAll(s, vals) { const d = await openDB(); const t = d.transaction(s, 'readwrite'); const o = t.objectStore(s); o.clear(); vals.forEach(v => o.put(v)); return txDone(t); }
const kvGet = k => dbGet('kv', k);
const kvSet = (k, v) => dbPut('kv', v, k);
const kvDel = k => dbDel('kv', k);
let _lock = Promise.resolve();
function withLock(fn) { const p = _lock.then(fn, fn); _lock = p.catch(() => { }); return p; }

/* ====================================================================== */
/* Cryptographie (WebCrypto)                                              */
/* ====================================================================== */
const EC = { name: 'ECDH', namedCurve: 'P-256' };
const HKDF_INFO = te.encode('rondes-cachan v1');
async function hkdfKey(bits, salt, usages) {
  const base = await crypto.subtle.importKey('raw', bits, 'HKDF', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'HKDF', hash: 'SHA-256', salt, info: HKDF_INFO }, base, { name: 'AES-GCM', length: 256 }, false, usages);
}
const _pubCache = new Map();
async function importPub(jwk) {
  const s = JSON.stringify(jwk);
  if (!_pubCache.has(s)) _pubCache.set(s, await crypto.subtle.importKey('jwk', jwk, EC, false, []));
  return _pubCache.get(s);
}
function getPubKey() { return importPub(state.conn.pub); }
/* Chiffré pour les seuls superviseurs : [1][clé éphémère 65][iv 12][chiffré] */
async function sealBytesFor(pubKey, bytes) {
  const eph = await crypto.subtle.generateKey(EC, true, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'ECDH', public: pubKey }, eph.privateKey, 256);
  const epkRaw = new Uint8Array(await crypto.subtle.exportKey('raw', eph.publicKey));
  const key = await hkdfKey(bits, epkRaw, ['encrypt']);
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, bytes));
  const out = new Uint8Array(78 + ct.length);
  out[0] = 1; out.set(epkRaw, 1); out.set(iv, 66); out.set(ct, 78);
  return out;
}
async function unsealBytes(priv, data) {
  if (data[0] !== 1 || data.length < 95) throw new Error('format chiffré inconnu');
  const epkRaw = data.slice(1, 66);
  const epk = await crypto.subtle.importKey('raw', epkRaw, EC, false, []);
  const bits = await crypto.subtle.deriveBits({ name: 'ECDH', public: epk }, priv, 256);
  const key = await hkdfKey(bits, epkRaw, ['decrypt']);
  return new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: data.slice(66, 78) }, key, data.slice(78)));
}
async function seal(obj) {
  const r = await sealBytesFor(await getPubKey(), te.encode(JSON.stringify(obj)));
  return { epk: b64u.enc(r.subarray(1, 66)), iv: b64u.enc(r.subarray(66, 78)), ct: b64u.enc(r.subarray(78)) };
}
async function unseal(priv, sec) {
  const parts = [new Uint8Array([1]), b64u.dec(sec.epk), b64u.dec(sec.iv), b64u.dec(sec.ct)];
  const all = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0;
  parts.forEach(p => { all.set(p, o); o += p.length; });
  return JSON.parse(td.decode(await unsealBytes(priv, all)));
}
/* Chiffré pour l'équipe (clé symétrique partagée) : [2][iv 12][chiffré] */
let _teamKey = null, _teamRaw = null;
async function teamKey() {
  if (_teamKey && _teamRaw === state.conn.team) return _teamKey;
  _teamKey = await crypto.subtle.importKey('raw', b64u.dec(state.conn.team), 'AES-GCM', false, ['encrypt', 'decrypt']);
  _teamRaw = state.conn.team;
  return _teamKey;
}
async function teamEnc(bytes) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await teamKey(), bytes));
  const out = new Uint8Array(13 + ct.length); out[0] = 2; out.set(iv, 1); out.set(ct, 13);
  return out;
}
async function teamDec(data) {
  if (data[0] !== 2) throw new Error('format chiffré inconnu');
  return new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: data.slice(1, 13) }, await teamKey(), data.slice(13)));
}
const encJSON = async o => b64u.enc(await teamEnc(fflate.deflateSync(te.encode(JSON.stringify(o)))));
const decJSON = async s => JSON.parse(td.decode(fflate.inflateSync(await teamDec(b64u.dec(s)))));
async function newKeyPair() {
  const kp = await crypto.subtle.generateKey(EC, true, ['deriveBits']);
  const privJwk = await crypto.subtle.exportKey('jwk', kp.privateKey);
  const p = await crypto.subtle.exportKey('jwk', kp.publicKey);
  const priv = await crypto.subtle.importKey('jwk', privJwk, EC, false, ['deriveBits']); // non exportable
  return { priv, privJwk, pub: { kty: p.kty, crv: p.crv, x: p.x, y: p.y } };
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

/* ====================================================================== */
/* Relais                                                                 */
/* Requête « simple » (text/plain) : pas de pré-vol CORS, compatible avec */
/* la redirection des applications web Apps Script.                       */
/* ====================================================================== */
async function relayCall(url, body, timeoutMs = 60000) {
  const ctl = new AbortController();
  const t = setTimeout(() => ctl.abort(), timeoutMs);
  let r;
  try {
    r = await fetch(url, { method: 'POST', body: JSON.stringify(body), headers: { 'Content-Type': 'text/plain;charset=utf-8' }, redirect: 'follow', signal: ctl.signal, cache: 'no-store' });
  } catch (e) { throw new Error('relais injoignable'); }
  finally { clearTimeout(t); }
  if (!r.ok) throw new Error('relais : erreur ' + r.status);
  let j; try { j = await r.json(); } catch (e) { throw new Error('relais : réponse illisible (vérifiez l’adresse et l’accès « Tout le monde »)'); }
  if (!j.ok) { const e = new Error(j.err === 'deja-initialise' ? 'Ce relais est déjà activé par un superviseur.' : 'relais : ' + (j.err || 'refus')); e.code = j.err; throw e; }
  return j;
}
function rel(op, extra = {}, timeout) {
  const c = state.conn;
  return relayCall(c.url, { op, tok: state.mode === 'superviseur' ? c.supTok : c.agentTok, ...extra }, timeout);
}

/* ====================================================================== */
/* Configuration                                                          */
/* ====================================================================== */
function defaultConfig() {
  return {
    v: 2, cfgId: randCode(8),
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
    urgenceTel: '', urgenceMail: ''
  };
}
const VEHICULE_RE = /v[ée]hicule/i;
function siteById(id) { return state.cfg && state.cfg.sites.find(s => s.id === id); }
function siteNom(id) { const s = siteById(id); return s ? s.nom : id; }
function validCfg(c) { return c && Array.isArray(c.sites) && Array.isArray(c.agents) && Array.isArray(c.dests) && Array.isArray(c.cats) && Array.isArray(c.checklist); }
/* QR de configuration des téléphones agents : accès au relais et clés publiques/équipe uniquement */
function packJoin(c) { return b64u.enc(fflate.deflateSync(te.encode(JSON.stringify({ v: 2, url: c.url, tok: c.agentTok, team: c.team, pub: c.pub })), { level: 9 })); }
function unpackJoin(s) { return JSON.parse(td.decode(fflate.inflateSync(b64u.dec(s)))); }

/* ====================================================================== */
/* Géolocalisation : uniquement pendant une ronde                         */
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
const state = { mode: null, conn: null, cfg: null, priv: null, view: 'home', params: {}, cleanup: null, sync: { busy: false } };

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
  clearTimeout(_toastT); _toastT = setTimeout(() => t.remove(), bad ? 7000 : 3500);
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
  if (!opts.keepScroll) window.scrollTo(0, 0);
}
let _renderSeq = 0;
async function render() {
  const root = document.getElementById('app');
  const fn = VIEWS[state.view] || VIEWS.home;
  const my = ++_renderSeq;
  try {
    const node = await fn(state.params);
    if (my !== _renderSeq) return;
    root.replaceChildren(node);
  } catch (e) {
    console.error(e);
    root.replaceChildren(page('Erreur', { back: true }, h('div', { class: 'banner bad' }, String(e && e.message || e))));
  }
}
function refreshIfHome() { if (state.view === 'home' && !document.getElementById('busy')) render(); }
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
  const ageMin = file.lastModified ? Math.round((Date.now() - file.lastModified) / 60000) : null;
  return { id: uuid(), blob: full, sha: await sha256hex(full), thumb: await blobToDataURL(thumb), ageMin };
}
function photoInput(onFile) {
  return safe(async () => {
    const i = document.createElement('input');
    i.type = 'file'; i.accept = 'image/*'; i.setAttribute('capture', 'environment');
    i.onchange = safe(async () => { if (i.files[0]) await onFile(i.files[0]); });
    i.click();
  });
}
async function thumbOf(s) {
  if (s.thumb) return s.thumb;
  const t = await dbGet('thumbs', s.id);
  return t ? t.data : null;
}

/* ====================================================================== */
/* Statuts                                                                */
/* ====================================================================== */
const VERDICTS = { resolu: 'Résolu', encours: 'Toujours en cours', aggrave: 'Aggravé' };
const STATUTS = { ouvert: 'Ouvert', clos: 'Clos' };
function statutBadge(s) { return h('span', { class: 'badge ' + (s === 'clos' ? 'b-ok' : 'b-warn') }, STATUTS[s] || s); }
function applyVerdict(sig, verdict, quand, agent) {
  if (verdict === 'resolu') { sig.statut = 'clos'; sig.closLe = quand; sig.closPar = agent || 'agent'; }
  else { sig.statut = 'ouvert'; sig.closLe = null; sig.closPar = null; }
  if (verdict === 'aggrave') sig.aggrave = true;
}

/* ====================================================================== */
/* Reconstitution de l'état à partir du journal                           */
/* Ordre : numéro attribué par le relais, puis saisies locales en attente.*/
/* ====================================================================== */
const ORD_LOCAL = 1e15;
async function fold() {
  const sup = state.mode === 'superviseur';
  const entries = await dbAll('entries');
  const items = [];
  for (const e of entries) items.push({ ord: e.n * 100000 + (e.idx || 0), e });
  if (!sup) for (const ev of await dbAll('events')) items.push({ ord: ORD_LOCAL + ev.seq, e: { kind: 'ev', ev, local: true } });
  if (sup) ((await kvGet('decQueue')) || []).forEach((q, i) => items.push({ ord: 2 * ORD_LOCAL + i, e: { kind: q.kind === 'n' ? 'note' : 'dec', dec: q.d, local: true } }));
  items.sort((a, b) => a.ord - b.ord);

  const sigs = new Map(), rondes = new Map(), devs = {}, chains = {};
  let cfg = null;
  const alertes = [];
  for (const { e } of items) {
    if (e.kind === 'ev') {
      const ev = e.ev, dec = e.dec || null, ts = dec && dec.ts;
      if (sup) {
        (chains[ev.dev] = chains[ev.dev] || []).push(e);
        const dv = devs[ev.dev] = devs[ev.dev] || { seq: 0, at: null };
        if (ev.seq > dv.seq) { dv.seq = ev.seq; dv.at = e.at; dv.agent = ev.agent; }
      }
      if (ev.t === 'signalement') {
        if (!sigs.has(ev.data.id)) sigs.set(ev.data.id, {
          id: ev.data.id, ref: ev.data.ref, site: ev.site, cat: ev.data.cat, desc: ev.data.desc, plaque: ev.data.plaque,
          dests: ev.data.dests, urgent: ev.data.urgent, statut: 'ouvert', jour: ev.jour, agent: ev.agent,
          thumb: e.local && typeof ev.data.thumb === 'string' ? ev.data.thumb : null,
          photos: (ev.photos || []).map(p => ({ id: p.id, sha: p.sha, n: e.n })), ts, geo: dec && dec.geo, dev: ev.dev, seq: ev.seq,
          suivi: [], journalDPMS: [], notes: '', notesAgents: '', lieu: e.local ? 'local' : 'relais'
        });
      } else if (ev.t === 'revue') {
        const s = sigs.get(ev.data.sig);
        if (s) {
          s.suivi.push({ jour: ev.jour, ts, agent: ev.agent, verdict: ev.data.verdict, comment: ev.data.comment, photos: (ev.photos || []).map(p => ({ id: p.id, sha: p.sha, n: e.n })), geo: dec && dec.geo });
          applyVerdict(s, ev.data.verdict, ts || ev.jour, ev.agent);
        } else if (sup) alertes.push({ key: `orph:${ev.dev}:${ev.seq}`, at: e.at, msg: `${siteNom(ev.site)}, ${ev.agent}, ${fmtJour(ev.jour)} — constat sur un signalement inconnu (${ev.data.ref})` });
      }
      if (sup && ev.rid) {
        const r = rondes.get(ev.rid) || { rid: ev.rid, site: ev.site, agent: ev.agent, dev: ev.dev, jour: ev.jour, nbSig: 0, nbRevue: 0, alertes: [] };
        if (ev.t === 'ronde_debut') { r.debut = ts; r.qr = ev.data.qr; r.geoDebut = dec && dec.geo; }
        if (ev.t === 'ronde_fin') r.fin = ts;
        if (ev.t === 'checklist') r.checklist = ev.data.items;
        if (ev.t === 'signalement') r.nbSig++;
        if (ev.t === 'revue') r.nbRevue++;
        rondes.set(ev.rid, r);
      }
    } else if (e.kind === 'dec') {
      const d = e.dec;
      if (d.type === 'cfg') { if (validCfg(d.cfg)) cfg = d.cfg; }
      else if (d.type === 'sig') {
        const s = sigs.get(d.sig);
        if (s) {
          Object.assign(s, d.patch);
          s.journalDPMS.push({ ts: d.ts, action: d.action, par: d.par || 'DPMS' });
        }
      }
    } else if (e.kind === 'note' && sup) {
      const s = sigs.get(e.dec.sig);
      if (s) s.notes = e.dec.text;
    }
  }

  if (sup) {
    // Contrôles d'intégrité par téléphone : continuité de la série et chaînage
    for (const [dev, list] of Object.entries(chains)) {
      list.sort((a, b) => a.ev.seq - b.ev.seq);
      let prev = null;
      for (const e of list) {
        const ev = e.ev, lieu = `${siteNom(ev.site)}, ${ev.agent || '?'}, ${fmtJour(ev.jour)}`;
        const add = (m) => {
          alertes.push({ key: `${dev}:${ev.seq}:${m}`, at: e.at, msg: `${lieu} — saisie n° ${ev.seq} (${dev}) : ${m}` });
          if (ev.rid && rondes.has(ev.rid)) rondes.get(ev.rid).alertes.push(m);
        };
        if (!e.hashOk) add('empreinte invalide (saisie modifiée)');
        if (prev === null) { if (ev.seq !== 1) add(`saisies 1 à ${ev.seq - 1} manquantes`); else if (ev.prev !== '0'.repeat(64)) add('chaînage rompu'); }
        else if (ev.seq !== prev.ev.seq + 1) add(`saisies ${prev.ev.seq + 1} à ${ev.seq - 1} manquantes`);
        else if (ev.prev !== prev.ev.hash) add('chaînage rompu avec la saisie précédente');
        if (!e.dec) add('horodatage et position illisibles');
        else {
          if (e.dec.qr === false) add('ronde démarrée sans QR code');
          (e.dec.photoAges || []).forEach(a => { if (a.ageMin != null && a.ageMin > 15) add(`photo prise ${a.ageMin} min avant la saisie`); });
          if (ev.t === 'ronde_debut' && e.dec.geo && e.dec.geo.err) add('position : ' + e.dec.geo.err);
        }
        prev = e;
      }
    }
    for (const r of rondes.values()) if (r.debut && !r.fin && Date.now() - new Date(r.debut).getTime() > 6 * 3600000) r.alertes.push('fin de ronde non reçue');
    await dbReplaceAll('rondes', [...rondes.values()]);
    await kvSet('alertes', alertes.sort((a, b) => (b.at || '').localeCompare(a.at || '')));
    await kvSet('devs', devs);
  }
  await dbReplaceAll('sigs', [...sigs.values()]);
  if (cfg && JSON.stringify(cfg) !== JSON.stringify(state.cfg)) { state.cfg = cfg; await kvSet('cfg', cfg); }
}

/* ====================================================================== */
/* AGENTS : journal chaîné local                                          */
/* ====================================================================== */
async function appendEvent(type, base, data, secretExtra = {}, photos = []) {
  return withLock(async () => {
    const head = (await kvGet('head')) || { seq: 0, hash: '0'.repeat(64) };
    const now = new Date();
    const secret = { ts: now.toISOString(), tzo: now.getTimezoneOffset(), geo: Geo.snap(), ...secretExtra };
    const ev = {
      v: 2, dev: await kvGet('dev'), seq: head.seq + 1, t: type, jour: localDate(now),
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
async function savePhotos(list) { await dbPutMany('photos', list.map(p => ({ id: p.id, blob: p.blob, sha: p.sha }))); }

/* Entrée de journal : la vignette est rangée à part, l'empreinte reste vérifiable */
async function storeEvEntry(ev, n, idx, at, extra = {}) {
  const k = `e:${ev.dev}:${String(ev.seq).padStart(8, '0')}`;
  const old = await dbGet('entries', k);
  if (old && old.ev.hash === ev.hash && old.n <= n) return false;    // déjà connue (renvoi après coupure)
  if (ev.t === 'signalement' && typeof ev.data.thumb === 'string') await dbPut('thumbs', { id: ev.data.id, data: ev.data.thumb });
  const light = ev.t === 'signalement' && typeof ev.data.thumb === 'string' ? { ...ev, data: { ...ev.data, thumb: true } } : ev;
  await dbPut('entries', { k, kind: 'ev', n, idx, at, ev: light, ...extra });
  return true;
}
function restoreThumbForHash(ev, thumb) {
  return ev.t === 'signalement' && ev.data.thumb === true ? { ...ev, data: { ...ev.data, thumb } } : ev;
}

/* ---------- Synchronisation ---------- */
let _syncP = null, _syncAgain = null;
function syncNow() {
  // Une synchronisation déjà en cours a pu démarrer avant la dernière saisie : on en relance une à sa suite.
  if (_syncP) { if (!_syncAgain) _syncAgain = _syncP.then(() => { _syncAgain = null; return syncNow(); }); return _syncAgain; }
  _syncP = (async () => {
    await null;            // garantit que _syncP est affecté avant toute sortie anticipée
    state.sync.busy = true;
    try {
      if (!state.conn) return { skipped: true };
      if (!navigator.onLine) return { offline: true };
      const r = state.mode === 'agent' ? await agentSync() : await supSync();
      await kvDel('syncErr');
      await kvSet('lastSync', new Date().toISOString());
      return r;
    } catch (e) { await kvSet('syncErr', { at: new Date().toISOString(), msg: e.message }); return { err: e.message }; }
    finally { state.sync.busy = false; _syncP = null; }
  })();
  return _syncP;
}
async function agentSync() {
  const dev = await kvGet('dev');
  let sent = 0;
  // 1. Publication des saisies en attente
  for (let guard = 0; guard < 50; guard++) {
    const pending = (await dbAll('events')).sort((a, b) => a.seq - b.seq);
    if (!pending.length) break;
    const batch = []; let bytes = 0; const files = {};
    for (const ev of pending) {
      batch.push(ev);
      for (const p of ev.photos) {
        const ph = await dbGet('photos', p.id);
        if (ph) { files[`${p.id}.jpg`] = [new Uint8Array(await ph.blob.arrayBuffer()), { level: 0 }]; bytes += ph.blob.size; }
      }
      if (bytes > LOT_MAX_OCTETS) break;
    }
    const data = await encJSON({ v: 2, dev, events: batch });
    const media = Object.keys(files).length ? b64u.enc(await teamEnc(fflate.zipSync(files))) : null;
    const r = await rel('put', { dev, data, media }, 180000);
    for (let i = 0; i < batch.length; i++) await storeEvEntry(batch[i], r.n, i, r.at);
    const d = await openDB(); const t = d.transaction(['events', 'photos'], 'readwrite');
    batch.forEach(ev => { t.objectStore('events').delete(ev.seq); ev.photos.forEach(p => t.objectStore('photos').delete(p.id)); });
    await txDone(t);
    sent += batch.length;
  }
  // 2. Relève : saisies des autres téléphones, interventions et configuration du DPMS
  const got = await pullAll(['p', 'd']);
  if (sent || got) await fold();
  return { sent, got };
}
async function pullAll(kinds) {
  let after = (await kvGet('cursor')) || 0, got = 0;
  for (let guard = 0; guard < 400; guard++) {
    const r = await rel('pull', { after, kinds }, 180000);
    for (const x of r.entries) { if (x.data) await ingest(x); got++; }
    after = r.last;
    await kvSet('cursor', after);
    if (!r.more) break;
  }
  return got;
}
async function ingest(x) {
  if (x.k === 'p') {
    let pk; try { pk = await decJSON(x.data); } catch (e) { await noteIllisible(x); return; }
    for (let i = 0; i < pk.events.length; i++) {
      const ev = pk.events[i];
      const extra = { media: !!x.media };
      if (state.mode === 'superviseur') {
        const { hash, ...rest } = ev;
        extra.hashOk = (await sha256hex(canon(rest))) === hash;
        try { extra.dec = await unseal(state.priv, ev.sec); } catch (e) { extra.dec = null; }
      }
      await storeEvEntry(ev, x.n, i, x.at, extra);
    }
  } else if (x.k === 'd') {
    let d; try { d = await decJSON(x.data); } catch (e) { await noteIllisible(x); return; }
    await dbPut('entries', { k: `d:${x.n}`, kind: 'dec', n: x.n, at: x.at, dec: d });
  } else if (x.k === 'n' && state.mode === 'superviseur') {
    let d; try { d = JSON.parse(td.decode(fflate.inflateSync(await unsealBytes(state.priv, b64u.dec(x.data))))); } catch (e) { await noteIllisible(x); return; }
    await dbPut('entries', { k: `n:${x.n}`, kind: 'note', n: x.n, at: x.at, dec: d });
  }
}
async function noteIllisible(x) {
  const l = (await kvGet('illisibles')) || [];
  l.push({ n: x.n, k: x.k, at: x.at }); await kvSet('illisibles', l.slice(-100));
}

/* ====================================================================== */
/* VUES                                                                   */
/* ====================================================================== */
const VIEWS = {};
VIEWS.home = async () => {
  if (!state.mode) return viewSetup();
  return state.mode === 'agent' ? agentHome() : supHome();
};

/* ---------- Premier lancement ---------- */
async function viewSetup() {
  const pending = state.pendingJoin;
  if (pending) {
    return page('Configuration', null,
      h('div', { class: 'card' },
        h('h3', null, 'Configurer cet appareil comme téléphone des agents ?'),
        h('p', { class: 'muted' }, 'Il se connectera au relais et récupérera la configuration et les signalements en cours.'),
        h('button', { class: 'ok', onclick: safe(async () => { await setupAgent(pending); }) }, 'Oui, configurer'),
        h('button', { class: 'sec', onclick: () => { state.pendingJoin = null; render(); } }, 'Annuler')));
  }
  return page('Rondes parkings', null,
    h('div', { class: 'card' },
      h('h3', null, 'Téléphone des agents'),
      h('p', { class: 'muted' }, 'Sur un appareil superviseur, ouvrez « QR codes » puis scannez le QR de configuration avec ce téléphone. Plusieurs téléphones agents peuvent être configurés.'),
      h('button', { class: 'big', onclick: () => go('scan') }, 'Scanner le QR de configuration')),
    h('div', { class: 'card' },
      h('h3', null, 'Appareil superviseur (DPMS)'),
      h('p', { class: 'muted' }, 'Téléphone ou ordinateur. Plusieurs appareils superviseurs peuvent être utilisés.'),
      h('button', { class: 'sec', onclick: () => go('supJoin') }, 'Ajouter cet appareil comme superviseur'),
      h('button', { class: 'sec', onclick: () => go('supInit') }, 'Premier superviseur : activer un relais neuf')),
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
}
async function setupAgent(j) {
  if (!j || !j.url || !j.tok || !j.team || !j.pub) throw new Error('QR de configuration invalide.');
  if (state.mode === 'superviseur') throw new Error('Cet appareil est superviseur.');
  if (state.mode === 'agent' && JSON.stringify(state.conn.pub) === JSON.stringify(j.pub) && state.conn.url === j.url) {
    state.pendingJoin = null; toast('Ce téléphone est déjà configuré.'); go('home'); return;
  }
  if (state.mode === 'agent' && ((await dbAll('events')).length)) throw new Error('Des saisies ne sont pas encore envoyées : envoyez-les avant de changer de configuration.');
  showBusy('Connexion au relais…');
  const conn = { url: j.url, agentTok: j.tok, team: j.team, pub: j.pub };
  await relayCall(conn.url, { op: 'ping', tok: conn.agentTok });   // vérifie l'accès avant d'enregistrer
  if (state.mode === 'agent') { for (const s of ['entries', 'thumbs', 'sigs']) await dbClear(s); await kvDel('cursor'); }
  if (!(await kvGet('dev'))) await kvSet('dev', 'T-' + randCode(6));
  await kvSet('conn', conn); await kvSet('mode', 'agent');
  state.mode = 'agent'; state.conn = conn; state.pendingJoin = null;
  showBusy('Récupération de la configuration et des signalements…');
  const r = await syncNow();
  hideBusy();
  await requestPersist();
  if (r && r.err) toast('Configuré, mais la récupération a échoué : ' + r.err, true);
  else toast(state.cfg ? 'Téléphone des agents configuré.' : 'Configuré, mais aucune configuration trouvée sur le relais.', !state.cfg);
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
function parseLink(text) {
  let hash = '';
  try { hash = new URL(text, appBaseURL()).hash; } catch (e) { hash = ''; }
  if (!hash && text.startsWith('#')) hash = text;
  const m = /^#(join|site)=(.+)$/.exec(hash);
  return m ? { kind: m[1], val: m[2] } : null;
}
async function handleLink(text, fromScan) {
  const l = parseLink(text);
  if (!l) { toast('QR code non reconnu (ancienne version ?).', true); if (fromScan) go('home'); return; }
  if (l.kind === 'join') {
    let j; try { j = unpackJoin(l.val); } catch (e) { toast('QR de configuration illisible.', true); go('home'); return; }
    if (state.mode === 'superviseur') { toast('Cet appareil est superviseur : configuration ignorée.', true); go('home'); return; }
    if (state.mode === 'agent') { if (confirm('Reconfigurer ce téléphone avec ce QR ?')) await safe(setupAgent)(j); else go('home'); return; }
    state.pendingJoin = j; go('home'); return;
  }
  if (l.kind === 'site') {
    if (state.mode !== 'agent') { toast('QR de parking : à scanner avec un téléphone des agents.', true); go('home'); return; }
    if (!state.cfg) { toast('Configuration non encore reçue : connectez le téléphone au réseau.', true); go('home'); return; }
    const [id, token] = l.val.split('.');
    const site = siteById(id);
    if (!site || site.token !== token) { toast('QR code de parking inconnu ou périmé.', true); go('home'); return; }
    return startRondeFlow(site.id, true);
  }
}

/* ====================================================================== */
/* AGENTS                                                                 */
/* ====================================================================== */
function syncCard(enAttente, lastSync, syncErr) {
  return h('div', { class: 'card' }, h('h2', null, 'Envoi au DPMS'),
    enAttente > 0
      ? h('div', { class: 'banner warn' }, `${enAttente} saisie(s) en attente d’envoi. L’envoi est automatique dès que le réseau est disponible.`)
      : h('div', { class: 'banner ok' }, 'Tout est envoyé.'),
    lastSync ? h('p', { class: 'muted small' }, `Dernière liaison ${fmtQuand(lastSync)}.`) : null,
    syncErr ? h('p', { class: 'muted small' }, `Dernier essai : ${syncErr.msg}.`) : null,
    h('button', { class: 'sec', disabled: state.sync.busy, onclick: safe(async () => { showBusy('Synchronisation…'); const r = await syncNow(); hideBusy(); toast(r && r.err ? 'Échec : ' + r.err : r && r.offline ? 'Pas de réseau.' : 'Synchronisé.', !!(r && (r.err || r.offline))); render(); }) }, enAttente > 0 ? 'Envoyer maintenant' : 'Actualiser'));
}
async function agentHome() {
  const lastSync = await kvGet('lastSync');
  const syncErr = await kvGet('syncErr');
  const enAttente = (await dbAll('events')).length;
  if (!state.cfg) {
    return page('Rondes parkings', null,
      h('div', { class: 'banner warn' }, 'Configuration pas encore reçue du relais. Connectez le téléphone au réseau.'),
      syncCard(enAttente, lastSync, syncErr),
      h('button', { class: 'link', onclick: () => go('reglages') }, 'Réglages'));
  }
  const ronde = await kvGet('ronde');
  const sigs = await dbAll('sigs');
  const cfg = state.cfg;
  const perSite = cfg.sites.map(s => {
    const o = sigs.filter(x => x.site === s.id && x.statut === 'ouvert').length;
    return h('div', { class: 'stat' }, h('span', null, s.nom), h('span', null, h('b', null, o), h('span', { class: 'muted small' }, o > 1 ? ' ouverts' : ' ouvert')));
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
    syncCard(enAttente, lastSync, syncErr),
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
  ...(state.cfg ? state.cfg.sites : []).map(s => h('button', { class: 'big sec', onclick: () => startRondeFlow(s.id, false) }, s.nom)));

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
  const ronde = { rid: uuid(), site, agent, qr, debut: new Date().toISOString(), step: 'revue', checklist: {}, sigsCrees: [] };
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
  if (!open.length) { ronde.step = 'checklist'; await kvSet('ronde', ronde); return VIEWS.checklist(); }
  const btnValider = h('button', { class: 'ok' }, 'Valider la revue');
  const refresh = () => { btnValider.disabled = !open.every(s => revueDraft[s.id] && revueDraft[s.id].verdict); };
  const cards = [];
  for (const s of open) {
    const d = revueDraft[s.id] || (revueDraft[s.id] = { verdict: null, comment: '', photos: [] });
    const thumbs = h('div', { class: 'thumbs' });
    const drawThumbs = () => thumbs.replaceChildren(...d.photos.map((p, i) =>
      h('div', { class: 't' }, h('img', { src: p.thumb, alt: '' }), h('button', { onclick: () => { d.photos.splice(i, 1); drawThumbs(); } }, '×'))));
    drawThumbs();
    const vbtns = Object.entries(VERDICTS).map(([k, lbl]) => h('button', {
      class: d.verdict === k ? 'sel-' + k : '',
      onclick: (e) => { d.verdict = k; [...e.target.parentNode.children].forEach(b => b.className = ''); e.target.className = 'sel-' + k; refresh(); }
    }, lbl));
    const th = await thumbOf(s);
    cards.push(h('div', { class: 'card' },
      h('h3', null, s.cat, s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('p', { class: 'muted small' }, `${s.ref} — signalé le ${fmtJour(s.jour)}${s.agent ? ' par ' + s.agent : ''}`),
      s.desc ? h('p', null, s.desc) : null,
      s.plaque ? h('p', null, h('b', null, 'Plaque : '), s.plaque) : null,
      th ? h('div', { class: 'thumbs' }, h('img', { src: th, alt: 'Photo du signalement' })) : null,
      s.notesAgents ? h('div', { class: 'banner info' }, 'DPMS : ' + s.notesAgents) : null,
      h('p', { style: 'margin:10px 0 6px;font-weight:600' }, 'Aujourd’hui :'),
      h('div', { class: 'verdicts' }, ...vbtns),
      h('input', { type: 'text', placeholder: 'Commentaire (facultatif)', value: d.comment, style: 'margin-top:8px', oninput: e => d.comment = e.target.value }),
      thumbs,
      h('button', { class: 'sec', onclick: photoInput(async f => { showBusy('Photo…'); d.photos.push(await processPhoto(f)); hideBusy(); drawThumbs(); }) }, 'Ajouter une photo')));
  }
  btnValider.onclick = safe(async () => {
    showBusy('Enregistrement…');
    for (const s of open) {
      const d = revueDraft[s.id];
      await savePhotos(d.photos);
      await appendEvent('revue', ronde, { sig: s.id, ref: s.ref, verdict: d.verdict, comment: d.comment.trim() },
        { photoAges: d.photos.map(p => ({ id: p.id, ageMin: p.ageMin })) }, d.photos);
      delete revueDraft[s.id];
    }
    await fold();
    ronde.step = 'checklist'; await kvSet('ronde', ronde);
    hideBusy(); go('checklist');
  });
  refresh();
  return page(`Revue : ${siteNom(ronde.site)}`, null,
    h('div', { class: 'banner info' }, `${open.length} signalement(s) en cours sur ce parking. Indiquez leur état aujourd’hui. « Résolu » clôt le signalement.`),
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
    bA.onclick = safe(async () => { ronde.checklist[it.lbl] = 'Anomalie'; await kvSet('ronde', ronde); go('signalement', { cat: it.cat, from: 'checklist', item: it.lbl }); });
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
    h('button', { class: 'ok', onclick: safe(async () => { if (confirm('Terminer la ronde ?')) await finishRonde(ronde); }) }, 'Terminer la ronde'));
};
async function finishRonde(ronde, silent) {
  showBusy('Clôture de la ronde…');
  await appendEvent('ronde_fin', ronde, { nbSig: ronde.sigsCrees.length });
  await kvDel('ronde');
  Geo.stop();
  hideBusy();
  if (!silent) {
    toast('Ronde terminée. Envoi au DPMS en cours…');
    go('home');
    syncNow().then(r => {
      if (r && r.sent) toast('Ronde envoyée au DPMS.');
      else if (r && (r.offline || r.err)) toast('Pas de réseau : la ronde partira automatiquement plus tard.');
      refreshIfHome();
    });
  }
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
    const id = uuid(), ref = `${site.code}-${randCode(5)}`;
    await savePhotos(d.photos);
    await appendEvent('signalement', ronde,
      { id, ref, cat: d.cat, desc: d.desc.trim(), plaque: VEHICULE_RE.test(d.cat) ? d.plaque.trim().toUpperCase() : '', dests: [...d.dests], urgent: d.urgent, checklist: item || null, thumb: d.photos[0] ? d.photos[0].thumb : null },
      { photoAges: d.photos.map(p => ({ id: p.id, ageMin: p.ageMin })) }, d.photos);
    await fold();
    ronde.sigsCrees.push(id);
    await kvSet('ronde', ronde);
    const urgent = d.urgent;
    sigDraft = null;
    hideBusy();
    toast(`Signalement ${ref} enregistré.`);
    go(from === 'checklist' ? 'checklist' : 'ronde');
    if (urgent) syncNow();     // un signalement urgent part sans attendre la fin de la ronde
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
    h('div', { style: 'height:70px' }),
    h('div', { class: 'sticky-bottom' }, h('div', null, h('button', { class: 'ok', onclick: save }, 'Enregistrer le signalement'))));
};

/* ---------- Liste des signalements (agents) ---------- */
VIEWS.agentSigs = async () => {
  const sigs = (await dbAll('sigs')).filter(s => s.statut === 'ouvert').sort((a, b) => (b.jour || '').localeCompare(a.jour || ''));
  const cards = [];
  for (const s of sigs) {
    const th = await thumbOf(s);
    cards.push(h('div', { class: 'card' },
      h('h3', null, s.cat, s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('p', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ref} — ${fmtJour(s.jour)}${s.agent ? ' — ' + s.agent : ''}`),
      s.desc ? h('p', null, s.desc) : null,
      s.plaque ? h('p', null, h('b', null, 'Plaque : '), s.plaque) : null,
      th ? h('div', { class: 'thumbs' }, h('img', { src: th, alt: '' })) : null,
      s.notesAgents ? h('div', { class: 'banner info' }, 'DPMS : ' + s.notesAgents) : null));
  }
  return page('Signalements en cours', { back: true }, cards.length ? cards : h('p', { class: 'muted' }, 'Aucun signalement en cours.'));
};

/* ---------- Extrait Excel (agents) : toutes les rondes, sans horodatage précis ni position ---------- */
async function agentExcel(jours) {
  const lim = localDate(new Date(Date.now() - jours * 86400000));
  const evs = [...(await dbAll('entries')).filter(e => e.kind === 'ev').map(e => ({ ord: e.n * 100000 + e.idx, ev: e.ev })),
  ...(await dbAll('events')).map(ev => ({ ord: ORD_LOCAL + ev.seq, ev }))].sort((a, b) => a.ord - b.ord);
  const refs = new Map();
  evs.forEach(({ ev }) => { if (ev.t === 'signalement') refs.set(ev.data.id, ev.data); });
  const aoa = [['Date', 'Parking', 'Agent', 'Type', 'Réf.', 'Catégorie', 'Description', 'Plaque', 'Destinataires', 'Urgent', 'Constat', 'Commentaire']];
  for (const { ev } of evs) {
    if (ev.jour < lim) continue;
    if (ev.t === 'signalement') aoa.push([fmtJour(ev.jour), siteNom(ev.site), ev.agent, 'Signalement', ev.data.ref, ev.data.cat, ev.data.desc, ev.data.plaque || '', (ev.data.dests || []).join(', '), ev.data.urgent ? 'Oui' : '', '', '']);
    else if (ev.t === 'revue') { const s = refs.get(ev.data.sig) || {}; aoa.push([fmtJour(ev.jour), siteNom(ev.site), ev.agent, 'Constat de suivi', ev.data.ref, s.cat || '', s.desc || '', s.plaque || '', '', '', VERDICTS[ev.data.verdict], ev.data.comment || '']); }
    else if (ev.t === 'checklist') { const an = Object.entries(ev.data.items || {}).filter(([, v]) => v === 'Anomalie').map(([k]) => k); aoa.push([fmtJour(ev.jour), siteNom(ev.site), ev.agent, 'Check-list', '', '', an.length ? 'Anomalies : ' + an.join(', ') : 'RAS', '', '', '', '', '']); }
  }
  if (aoa.length === 1) { toast('Aucune saisie sur la période.'); return; }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [10, 12, 12, 16, 11, 24, 50, 12, 28, 8, 16, 30].map(w => ({ wch: w }));
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
  const enAttente = state.mode === 'agent' ? (await dbAll('events')).length : 0;
  return page('Réglages', { back: true },
    h('div', { class: 'card' }, h('dl', { class: 'kv' },
      h('dt', null, 'Mode'), h('dd', null, state.mode === 'agent' ? 'Téléphone agents' : 'Superviseur'),
      dev ? [h('dt', null, 'Identifiant'), h('dd', null, dev)] : null,
      state.conn ? [h('dt', null, 'Relais'), h('dd', { class: 'small' }, state.conn.url)] : null,
      h('dt', null, 'Stockage protégé'), h('dd', null, persisted === true ? 'Oui' : persisted === false ? 'Non (installez l’application sur l’écran d’accueil)' : 'Inconnu'),
      est ? [h('dt', null, 'Espace utilisé'), h('dd', null, (est.usage / 1048576).toFixed(1) + ' Mo')] : null,
      h('dt', null, 'Version'), h('dd', null, APP_VERSION))),
    persisted === false ? h('button', { class: 'sec', onclick: safe(async () => { const ok = await requestPersist(); toast(ok ? 'Stockage protégé.' : 'Refusé par le navigateur.', !ok); render(); }) }, 'Demander la protection du stockage') : null,
    state.mode === 'agent' ? h('button', { class: 'sec', onclick: () => go('scan') }, 'Scanner un nouveau QR de configuration') : null,
    h('hr'),
    h('div', { class: 'card' }, h('h2', null, 'Réinitialiser cet appareil'),
      h('p', { class: 'muted small' }, state.mode === 'agent'
        ? (enAttente ? `Attention : ${enAttente} saisie(s) pas encore envoyée(s) seront perdues. ` : '') + 'Les données déjà envoyées restent sur le relais et seront retrouvées après reconfiguration.'
        : 'Efface cet appareil. Les données restent sur le relais : l’appareil pourra être rajouté avec l’adresse du relais et la phrase de passe.'),
      h('button', {
        class: 'danger', onclick: safe(async () => {
          const r = prompt('Tapez EFFACER pour confirmer.');
          if (r !== 'EFFACER') return;
          Geo.stop();
          for (const s of Object.keys(STORES)) await dbClear(s);
          Object.assign(state, { mode: null, conn: null, cfg: null, priv: null });
          toast('Appareil réinitialisé.'); go('home');
        })
      }, 'Réinitialiser cet appareil')));
};

/* ====================================================================== */
/* SUPERVISEUR                                                            */
/* ====================================================================== */
const ETAPES_RELAIS = [
  'Avec le compte Google dédié, ouvrez script.google.com, créez un projet et collez le contenu du fichier relais/relais.gs.',
  'Déployer → Nouveau déploiement → Application web. Exécuter en tant que : moi. Accès : tout le monde. Autorisez l’accès à Drive.',
  'Copiez l’adresse de l’application web (elle se termine par /exec).'
];
VIEWS.supInit = async () => {
  const url = h('input', { type: 'text', placeholder: 'https://script.google.com/macros/s/…/exec' });
  const p1 = h('input', { type: 'password', autocomplete: 'new-password' });
  const p2 = h('input', { type: 'password', autocomplete: 'new-password' });
  const err = h('div');
  return page('Activer un relais neuf', { back: true },
    h('div', { class: 'banner info' }, 'À faire une seule fois, par le premier superviseur. Les autres appareils superviseurs utiliseront « Ajouter cet appareil comme superviseur ».'),
    h('div', { class: 'card' }, h('h2', null, 'Relais'), ...ETAPES_RELAIS.map((t, i) => h('p', { class: 'small' }, `${i + 1}. ${t}`)),
      h('label', { class: 'f' }, 'Adresse du relais'), url),
    h('div', { class: 'card' }, h('h2', null, 'Phrase de passe superviseur'),
      h('p', { class: 'small muted' }, 'Elle permet d’ajouter d’autres appareils superviseurs. Elle est la seule protection des clés : longue, notée, remise sous pli fermé à la DGS.'),
      h('label', { class: 'f' }, 'Phrase de passe (12 caractères minimum)'), p1,
      h('label', { class: 'f' }, 'Confirmation'), p2),
    err,
    h('button', {
      class: 'ok', onclick: safe(async () => {
        err.replaceChildren();
        const u = url.value.trim();
        if (!/^https:\/\/\S+$/.test(u) && !/^http:\/\/localhost[:/]/.test(u)) throw new Error('Adresse du relais invalide.');
        if (p1.value.length < 12) throw new Error('Phrase de passe : 12 caractères minimum.');
        if (p1.value !== p2.value) throw new Error('Les deux phrases de passe diffèrent.');
        showBusy('Création des clés et activation du relais…');
        const kp = await newKeyPair();
        const conn = { url: u, supTok: randCode(32), agentTok: randCode(32), team: b64u.enc(crypto.getRandomValues(new Uint8Array(32))), pub: kp.pub };
        try { await relayCall(u, { op: 'init', supTok: conn.supTok, agentTok: conn.agentTok }); }
        catch (e) {
          hideBusy();
          if (e.code === 'deja-initialise') {
            err.replaceChildren(h('div', { class: 'banner warn' },
              h('b', null, 'Ce relais est déjà activé. '), 'Deux possibilités :',
              h('p', { class: 'small' }, '• Ajouter cet appareil au superviseur existant : « Ajouter cet appareil comme superviseur », avec la même phrase de passe.'),
              h('p', { class: 'small' }, '• Repartir de zéro (efface tout le relais) : dans l’éditeur Apps Script, choisir la fonction « reinitialiserRelais » dans la barre d’outils, cliquer sur « Exécuter », puis recommencer ici.'),
              h('button', { class: 'sec', onclick: () => go('supJoin', { url: u }, { noPush: true }) }, 'Ajouter cet appareil comme superviseur')));
            return;
          }
          throw e;
        }
        const vault = await encryptWithPass(p1.value, { v: 2, privJwk: kp.privJwk, ...conn });
        state.mode = 'superviseur'; state.conn = conn; state.priv = kp.priv;
        await relayCall(u, { op: 'putVault', tok: conn.supTok, data: JSON.stringify(vault) });
        const cfg = defaultConfig();
        await kvSet('privKey', kp.priv); await kvSet('conn', conn); await kvSet('mode', 'superviseur');
        state.cfg = cfg; await kvSet('cfg', cfg);
        await queueDec('d', { type: 'cfg', ts: new Date().toISOString(), cfg });
        await syncNow();
        await requestPersist();
        hideBusy();
        toast('Relais activé.');
        go('supQR', { first: true });
      })
    }, 'Créer et activer'));
};
VIEWS.supJoin = async ({ url: u0 } = {}) => {
  const url = h('input', { type: 'text', value: u0 || '', placeholder: 'https://script.google.com/macros/s/…/exec' });
  const pass = h('input', { type: 'password', autocomplete: 'current-password' });
  return page('Ajouter un superviseur', { back: true },
    h('p', { class: 'muted' }, 'L’adresse du relais figure dans les Réglages de tout appareil déjà configuré.'),
    h('label', { class: 'f' }, 'Adresse du relais'), url,
    h('label', { class: 'f' }, 'Phrase de passe superviseur'), pass,
    h('button', {
      class: 'ok', onclick: safe(async () => {
        const u = url.value.trim();
        if (!u) throw new Error('Adresse du relais manquante.');
        showBusy('Récupération des clés…');
        const r = await relayCall(u, { op: 'getVault' });
        let v;
        try { v = await decryptWithPass(pass.value, JSON.parse(r.data)); } catch (e) { throw new Error('Phrase de passe incorrecte.'); }
        const priv = await crypto.subtle.importKey('jwk', v.privJwk, EC, false, ['deriveBits']);
        const conn = { url: u, supTok: v.supTok, agentTok: v.agentTok, team: v.team, pub: v.pub };
        for (const s of Object.keys(STORES)) await dbClear(s);
        await kvSet('privKey', priv); await kvSet('conn', conn); await kvSet('mode', 'superviseur');
        Object.assign(state, { mode: 'superviseur', conn, priv, cfg: null });
        showBusy('Récupération des données…');
        const s = await syncNow();
        await requestPersist();
        hideBusy();
        if (s && s.err) toast('Ajouté, mais la récupération a échoué : ' + s.err, true); else toast('Appareil superviseur ajouté.');
        go('home');
      })
    }, 'Ajouter cet appareil'));
};

/* ---------- Synchronisation superviseur ---------- */
async function queueDec(kind, d) {
  const q = (await kvGet('decQueue')) || [];
  q.push({ kind, d }); await kvSet('decQueue', q);
}
async function supSync() {
  // 1. Publication des interventions en attente
  let q = (await kvGet('decQueue')) || [];
  while (q.length) {
    const it = q[0];
    const data = it.kind === 'n'
      ? b64u.enc(await sealBytesFor(await getPubKey(), fflate.deflateSync(te.encode(JSON.stringify(it.d)))))
      : await encJSON(it.d);
    const r = await rel('putDec', { kind: it.kind, data });
    await dbPut('entries', { k: `${it.kind}:${r.n}`, kind: it.kind === 'n' ? 'note' : 'dec', n: r.n, at: r.at, dec: it.d });
    q = q.slice(1); await kvSet('decQueue', q);
  }
  // 2. Relève
  const got = await pullAll(['p', 'd', 'n']);
  await fold();
  return { got };
}

async function supHome() {
  const sigs = await dbAll('sigs');
  const devs = (await kvGet('devs')) || {};
  const lastSync = await kvGet('lastSync');
  const syncErr = await kvGet('syncErr');
  const vues = new Set((await kvGet('alertesVues')) || []);
  const alertes = ((await kvGet('alertes')) || []).filter(a => !vues.has(a.key));
  const ouverts = sigs.filter(s => s.statut === 'ouvert');
  const urg = ouverts.filter(s => s.urgent).length;
  const rondes = (await dbAll('rondes')).sort((a, b) => (b.debut || '').localeCompare(a.debut || '')).slice(0, 5);
  const queue = (await kvGet('decQueue')) || [];
  return page('Rondes parkings', null,
    h('div', { class: 'card' },
      h('div', { class: 'stat' },
        h('span', null, state.sync.busy ? 'Relève en cours…' : lastSync ? `Relevé ${fmtQuand(lastSync)}` : 'Jamais relevé'),
        h('button', { class: 'chip', style: 'width:auto', disabled: state.sync.busy, onclick: safe(async () => { const r = await syncNow(); toast(r && r.err ? 'Relève impossible : ' + r.err : r && r.offline ? 'Pas de réseau.' : r && r.got ? `${r.got} élément(s) reçu(s).` : 'Rien de nouveau.', !!(r && (r.err || r.offline))); render(); }) }, 'Actualiser')),
      syncErr ? h('p', { class: 'muted small' }, 'Dernier essai : ' + syncErr.msg) : null,
      queue.length ? h('p', { class: 'muted small' }, `${queue.length} intervention(s) en attente d’envoi.`) : null),
    alertes.length ? h('button', { class: 'list-item', style: 'border-color:var(--warn)', onclick: () => go('supAlertes') },
      h('div', { class: 'l1' }, `${alertes.length} point(s) d’attention`, h('span', { class: 'badge b-warn' }, 'à lire')),
      h('div', { class: 'muted small' }, alertes[0].msg)) : null,
    h('div', { class: 'card' }, h('h2', null, 'Signalements'),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'ouvert' }) }, h('div', { class: 'stat' }, h('span', null, 'En cours', urg ? h('span', { class: 'badge b-bad' }, urg + ' urgent(s)') : null), h('b', null, ouverts.length))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'clos' }) }, h('div', { class: 'stat' }, h('span', null, 'Clos'), h('b', null, sigs.length - ouverts.length)))),
    h('div', { class: 'card' }, h('h2', null, 'Dernières rondes'),
      rondes.length ? rondes.map(rondeLigne) : h('p', { class: 'muted' }, 'Aucune ronde reçue.'),
      rondes.length ? h('button', { class: 'sec', onclick: () => go('supRondes') }, 'Toutes les rondes') : null),
    h('div', { class: 'card' }, h('h2', null, 'Exports'),
      h('button', { class: 'sec', onclick: safe(supExcel) }, 'Excel complet (heures et positions)'),
      h('button', { class: 'sec', onclick: safe(supArchive) }, 'Archive complète avec photos (.zip)')),
    h('div', { class: 'card' }, h('h2', null, 'Paramétrage'),
      h('button', { class: 'sec', onclick: () => go('supQR') }, 'QR codes (parkings et téléphones agents)'),
      h('button', { class: 'sec', onclick: () => go('supConfig') }, 'Listes et coordonnées'),
      h('button', { class: 'sec', onclick: () => go('supRelais') }, 'Relais et appareils')),
    Object.keys(devs).length ? h('div', { class: 'card' }, h('h2', null, 'Téléphones agents'),
      ...Object.entries(devs).map(([d, v]) => h('div', { class: 'stat' }, h('span', { style: 'white-space:nowrap;margin-right:8px' }, d), h('span', { class: 'muted small', style: 'text-align:right' }, `${v.seq} saisies — dernier envoi ${fmtQuand(v.at)}`)))) : null,
    h('button', { class: 'link', onclick: () => go('reglages') }, 'Réglages'),
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
}
function rondeLigne(r) {
  const dur = r.debut && r.fin ? Math.round((new Date(r.fin) - new Date(r.debut)) / 60000) : null;
  const anos = r.checklist ? Object.entries(r.checklist).filter(([, v]) => v === 'Anomalie').map(([k]) => k) : [];
  return h('div', { class: 'stat', style: 'display:block' },
    h('div', null, h('b', null, `${siteNom(r.site)} — ${r.agent}`), r.qr === false ? h('span', { class: 'badge b-warn' }, 'sans QR') : null, !r.fin ? h('span', { class: 'badge b-warn' }, 'fin non reçue') : null),
    h('div', { class: 'muted small' }, `${fmtDT(r.debut)}${r.fin ? ' → ' + fmtHeure(r.fin) : ''}${dur != null ? ` (${dur} min)` : ''} — ${r.nbSig} signalement(s), ${r.nbRevue} constat(s)`),
    anos.length ? h('div', { class: 'small' }, 'Anomalies : ' + anos.join(', ')) : null);
}
VIEWS.supAlertes = async () => {
  const al = (await kvGet('alertes')) || [];
  const vues = new Set((await kvGet('alertesVues')) || []);
  const node = page('Points d’attention', { back: true },
    h('p', { class: 'muted small' }, 'Contrôles automatiques : intégrité des saisies, continuité des séries, démarrages sans QR, photos anciennes, absence de position.'),
    al.length ? al.slice(0, 300).map(a => h('div', { class: 'card', style: vues.has(a.key) ? '' : 'border-color:var(--warn)' }, h('div', { class: 'muted small' }, fmtDT(a.at)), h('div', null, a.msg)))
      : h('p', { class: 'muted' }, 'Aucun.'));
  await kvSet('alertesVues', [...new Set([...vues, ...al.map(a => a.key)])].slice(-5000));
  return node;
};

/* ---------- Relais et appareils ---------- */
VIEWS.supRelais = async () => page('Relais et appareils', { back: true },
  h('div', { class: 'card' }, h('h2', null, 'Adresse du relais'), h('p', { class: 'small', style: 'word-break:break-all' }, state.conn.url)),
  h('div', { class: 'card' }, h('h2', null, 'Ajouter un appareil'),
    h('p', { class: 'small' }, h('b', null, 'Superviseur (téléphone ou ordinateur) : '), 'ouvrir l’application → « Ajouter cet appareil comme superviseur » → adresse ci-dessus + phrase de passe.'),
    h('p', { class: 'small' }, h('b', null, 'Téléphone agents : '), 'ouvrir l’application → scanner le QR de configuration (menu « QR codes »).'),
    h('p', { class: 'small muted' }, 'Un appareil ajouté ou réinstallé retrouve toutes les données conservées sur le relais.')),
  h('div', { class: 'card' }, h('h2', null, 'Repartir de zéro'),
    h('p', { class: 'small' }, 'Dans l’éditeur Apps Script du relais : choisir la fonction « reinitialiserRelais » dans la barre d’outils, cliquer sur « Exécuter ». Tout le contenu du relais est mis à la corbeille du compte Google. Réinitialiser ensuite chaque appareil (Réglages), puis activer un relais neuf.')));

/* ---------- Liste et fiche des signalements ---------- */
VIEWS.supSigs = async ({ f = 'ouvert', site = '' }) => {
  const all = await dbAll('sigs');
  const list = all.filter(s => (f === 'tous' || s.statut === f) && (!site || s.site === site))
    .sort((a, b) => (b.urgent - a.urgent) || (b.ts || b.jour || '').localeCompare(a.ts || a.jour || ''));
  const fchips = [['ouvert', 'En cours'], ['clos', 'Clos'], ['tous', 'Tous']].map(([k, l]) =>
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
/* Photos : téléchargées à la demande depuis le relais, vérifiées, conservées sur l'appareil */
async function fetchMedia(n, wanted) {
  const r = await rel('media', { n }, 180000);
  const z = fflate.unzipSync(await teamDec(b64u.dec(r.data)));
  const out = [];
  for (const [name, bytes] of Object.entries(z)) {
    const id = name.replace(/\.jpg$/, '');
    const sha = await sha256hex(bytes);
    out.push({ id, blob: new Blob([bytes], { type: 'image/jpeg' }), sha, ok: !wanted || !wanted[id] || wanted[id] === sha });
  }
  await dbPutMany('photos', out);
  return out;
}
async function photoBlocks(list) {
  const out = [];
  const missing = new Map();
  for (const p of list || []) { const c = await dbGet('photos', p.id); if (!c && p.n) missing.set(p.n, true); }
  for (const n of missing.keys()) { try { await fetchMedia(n); } catch (e) { out.push(h('p', { class: 'muted small' }, 'Photos indisponibles : ' + e.message)); } }
  for (const p of list || []) {
    const c = await dbGet('photos', p.id);
    if (!c) continue;
    if (p.sha && c.sha !== p.sha) out.push(h('div', { class: 'banner bad' }, 'Photo modifiée après la saisie (empreinte différente).'));
    out.push(h('img', { class: 'photo-full', src: URL.createObjectURL(c.blob), alt: 'Photo' }));
  }
  return out;
}
async function supDecision(s, action, patch) {
  await queueDec('d', { type: 'sig', id: uuid(), ts: new Date().toISOString(), par: 'DPMS', action, sig: s.id, patch });
  await fold();
  syncNow();     // l'état local est déjà à jour : pas de rafraîchissement d'écran (il effacerait une saisie en cours)
}
VIEWS.supSig = async ({ id, edit }) => {
  const s = await dbGet('sigs', id);
  if (!s) return page('Signalement', { back: true }, h('p', null, 'Introuvable.'));
  const back = () => go('supSigs', { f: s.statut });
  const done = (msg) => { toast(msg); go('supSig', { id }, { noPush: true, keepScroll: true }); };
  if (edit) {
    const cat = h('select', null, ...state.cfg.cats.map(c => h('option', { value: c, selected: c === s.cat ? true : null }, c)));
    const desc = h('textarea', { value: s.desc || '' });
    const plaque = h('input', { type: 'text', value: s.plaque || '' });
    const dests = [...(s.dests || [])];
    const destChips = h('div', { class: 'chips' }, ...state.cfg.dests.map(x => h('button', {
      class: 'chip' + (dests.includes(x) ? ' on' : ''),
      onclick: e => { const i = dests.indexOf(x); if (i >= 0) dests.splice(i, 1); else dests.push(x); e.target.classList.toggle('on'); }
    }, x)));
    let urgent = !!s.urgent;
    const urg = h('button', { class: 'chip' + (urgent ? ' on-bad' : ''), onclick: e => { urgent = !urgent; e.target.className = 'chip' + (urgent ? ' on-bad' : ''); } }, 'Urgent');
    const msg = h('input', { type: 'text', value: s.notesAgents || '', placeholder: 'Ex. : intervention Ateliers prévue mardi' });
    return page(`Modifier ${s.ref}`, { back: () => go('supSig', { id }, { noPush: true }) },
      h('label', { class: 'f' }, 'Catégorie'), cat,
      h('label', { class: 'f' }, 'Description'), desc,
      h('label', { class: 'f' }, 'Plaque'), plaque,
      h('label', { class: 'f' }, 'À signaler à'), destChips,
      h('label', { class: 'f' }, 'Priorité'), h('div', { class: 'chips' }, urg),
      h('label', { class: 'f' }, 'Message affiché aux agents'), msg,
      h('button', {
        class: 'ok', onclick: safe(async () => {
          await supDecision(s, 'modification', { cat: cat.value, desc: desc.value.trim(), plaque: plaque.value.trim().toUpperCase(), dests: [...dests], urgent, notesAgents: msg.value.trim() });
          done('Modifié. Transmis aux agents.');
        })
      }, 'Enregistrer et transmettre aux agents'));
  }
  const notes = h('textarea', { value: s.notes || '', placeholder: 'Notes internes DPMS (lisibles par les seuls superviseurs)' });
  const suivi = [];
  for (const v of s.suivi || []) {
    suivi.push(h('div', { class: 'card' },
      h('div', null, h('b', null, VERDICTS[v.verdict] || v.verdict), ` — ${v.ts ? fmtDT(v.ts) : fmtJour(v.jour)} — ${v.agent || ''}`),
      v.comment ? h('p', null, v.comment) : null,
      h('div', { class: 'small' }, 'Position : ', geoLink(v.geo)),
      ...(await photoBlocks(v.photos))));
  }
  return page(`Signalement ${s.ref}`, { back },
    h('div', { class: 'card' },
      h('h3', null, s.cat, statutBadge(s.statut), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('dl', { class: 'kv' },
        h('dt', null, 'Parking'), h('dd', null, siteNom(s.site)),
        h('dt', null, 'Signalé'), h('dd', null, s.ts ? fmtDT(s.ts) : fmtJour(s.jour), s.agent ? ' par ' + s.agent : ''),
        h('dt', null, 'Position'), h('dd', null, geoLink(s.geo)),
        h('dt', null, 'À signaler à'), h('dd', null, (s.dests || []).join(', ')),
        s.plaque ? [h('dt', null, 'Plaque'), h('dd', null, s.plaque)] : null,
        s.statut === 'clos' ? [h('dt', null, 'Clos'), h('dd', null, `${s.closLe && s.closLe.length > 10 ? fmtDT(s.closLe) : fmtJour(s.closLe)}${s.closPar ? ' par ' + s.closPar : ''}`)] : null,
        s.notesAgents ? [h('dt', null, 'Message aux agents'), h('dd', null, s.notesAgents)] : null),
      s.desc ? h('p', null, s.desc) : null,
      ...(await photoBlocks(s.photos))),
    suivi.length ? h('h2', { style: 'font-size:16px;color:var(--navy)' }, 'Constats des agents') : null, ...suivi,
    h('div', { class: 'card' }, h('h2', null, 'Intervenir (facultatif)'),
      h('p', { class: 'muted small' }, 'Toute intervention est transmise automatiquement aux téléphones des agents et aux autres superviseurs.'),
      s.statut !== 'clos'
        ? h('button', { class: 'ok', onclick: safe(async () => { if (!confirm('Clôturer ce signalement ?')) return; await supDecision(s, 'clôture', { statut: 'clos', closLe: new Date().toISOString(), closPar: 'DPMS' }); done('Clos. Transmis aux agents.'); }) }, 'Clôturer')
        : h('button', { class: 'sec', onclick: safe(async () => { await supDecision(s, 'réouverture', { statut: 'ouvert', closLe: null, closPar: null }); done('Rouvert. Transmis aux agents.'); }) }, 'Rouvrir'),
      h('button', { class: 'sec', onclick: () => go('supSig', { id, edit: true }, { noPush: true }) }, 'Modifier ou écrire aux agents'),
      s.plaque && s.statut === 'clos' ? h('button', { class: 'link', onclick: safe(async () => { if (confirm('Effacer la plaque de ce signalement clos ?')) { await supDecision(s, 'effacement de la plaque', { plaque: '' }); done('Plaque effacée.'); } }) }, 'Effacer la plaque (dossier traité)') : null),
    h('div', { class: 'card' }, h('h2', null, 'Notes internes'), notes,
      h('button', { class: 'sec', onclick: safe(async () => { await queueDec('n', { type: 'note', ts: new Date().toISOString(), sig: s.id, text: notes.value }); await fold(); syncNow(); toast('Notes enregistrées.'); }) }, 'Enregistrer les notes')),
    (s.journalDPMS || []).length ? h('div', { class: 'card' }, h('h2', null, 'Interventions DPMS'),
      ...s.journalDPMS.map(j => h('p', { class: 'small' }, `${fmtDT(j.ts)} — ${j.action}`))) : null);
};

/* ---------- Historique des rondes ---------- */
VIEWS.supRondes = async () => {
  const rs = (await dbAll('rondes')).sort((a, b) => (b.debut || '').localeCompare(a.debut || '')).slice(0, 300);
  return page('Rondes', { back: true },
    rs.length ? rs.map(r => h('div', { class: 'card' }, rondeLigne(r), h('div', { class: 'small' }, 'Position au départ : ', geoLink(r.geoDebut)),
      r.alertes && r.alertes.length ? h('div', { class: 'small', style: 'color:var(--warn)' }, [...new Set(r.alertes)].join(' ; ')) : null))
      : h('p', { class: 'muted' }, 'Aucune ronde reçue.'));
};

/* ---------- Excel complet ---------- */
async function supExcel() {
  showBusy('Construction du classeur…');
  const sigs = (await dbAll('sigs')).sort((a, b) => (a.ts || a.jour || '').localeCompare(b.ts || b.jour || ''));
  const rondes = (await dbAll('rondes')).sort((a, b) => (a.debut || '').localeCompare(b.debut || ''));
  const evs = (await dbAll('entries')).filter(e => e.kind === 'ev').sort((a, b) => a.k.localeCompare(b.k));
  const g = (geo, k) => geo && !geo.err && geo[k] != null ? geo[k] : '';
  const A1 = [['Réf.', 'Parking', 'Catégorie', 'Description', 'Plaque', 'Destinataires', 'Urgent', 'Aggravé', 'Statut', 'Signalé le', 'Agent', 'Latitude', 'Longitude', 'Précision (m)', 'Photos', 'Dernier constat', 'Date dernier constat', 'Clos le', 'Clos par', 'Message aux agents', 'Notes DPMS']];
  sigs.forEach(s => {
    const last = (s.suivi || []).slice(-1)[0];
    A1.push([s.ref, siteNom(s.site), s.cat, s.desc, s.plaque || '', (s.dests || []).join(', '), s.urgent ? 'Oui' : '', s.aggrave ? 'Oui' : '', STATUTS[s.statut] || s.statut, s.ts ? fmtDT(s.ts) : fmtJour(s.jour), s.agent || '',
      g(s.geo, 'lat'), g(s.geo, 'lon'), g(s.geo, 'acc'), (s.photos || []).length, last ? VERDICTS[last.verdict] : '', last ? (last.ts ? fmtDT(last.ts) : fmtJour(last.jour)) : '',
      s.closLe ? (s.closLe.length > 10 ? fmtDT(s.closLe) : fmtJour(s.closLe)) : '', s.closPar || '', s.notesAgents || '', s.notes || '']);
  });
  const A2 = [['Parking', 'Agent', 'Début', 'Fin', 'Durée (min)', 'Démarrage par QR', 'Anomalies check-list', 'Signalements', 'Constats', 'Lat. départ', 'Lon. départ', 'Précision (m)', 'Points d’attention', 'Téléphone']];
  rondes.forEach(r => {
    const dur = r.debut && r.fin ? Math.round((new Date(r.fin) - new Date(r.debut)) / 60000) : '';
    const anos = r.checklist ? Object.entries(r.checklist).filter(([, v]) => v === 'Anomalie').map(([k]) => k).join(', ') : '';
    A2.push([siteNom(r.site), r.agent, fmtDT(r.debut), fmtDT(r.fin), dur, r.qr === false ? 'Non' : 'Oui', anos, r.nbSig, r.nbRevue, g(r.geoDebut, 'lat'), g(r.geoDebut, 'lon'), g(r.geoDebut, 'acc'), [...new Set(r.alertes || [])].join(' ; '), r.dev]);
  });
  const TYPES = { ronde_debut: 'Début de ronde', ronde_fin: 'Fin de ronde', checklist: 'Check-list', signalement: 'Signalement', revue: 'Constat de suivi' };
  const A3 = [['Téléphone', 'N°', 'Type', 'Jour déclaré', 'Horodatage appareil', 'Reçu par le relais', 'Parking', 'Agent', 'Détail', 'Latitude', 'Longitude', 'Précision (m)', 'Âge position (s)', 'Empreinte']];
  evs.forEach(e => {
    const ev = e.ev; let det = '';
    if (ev.t === 'signalement') det = `${ev.data.ref} ${ev.data.cat}${ev.data.desc ? ' — ' + ev.data.desc : ''}`;
    if (ev.t === 'revue') det = `${ev.data.ref} : ${VERDICTS[ev.data.verdict]}${ev.data.comment ? ' — ' + ev.data.comment : ''}`;
    if (ev.t === 'checklist') det = Object.entries(ev.data.items || {}).map(([k, v]) => `${k} : ${v}`).join(' ; ');
    if (ev.t === 'ronde_debut') det = ev.data.qr ? 'QR scanné' : 'sans QR';
    if (ev.t === 'ronde_fin') det = `${ev.data.nbSig} signalement(s)`;
    const geo = e.dec && e.dec.geo;
    A3.push([ev.dev, ev.seq, TYPES[ev.t] || ev.t, fmtJour(ev.jour), e.dec ? fmtDT(e.dec.ts) : 'illisible', fmtDT(e.at), siteNom(ev.site), ev.agent || '', det, g(geo, 'lat'), g(geo, 'lon'), g(geo, 'acc'), g(geo, 'age'), e.hashOk ? 'OK' : 'INVALIDE']);
  });
  const wb = XLSX.utils.book_new();
  const add = (aoa, name, widths) => { const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = widths.map(w => ({ wch: w })); ws['!autofilter'] = { ref: ws['!ref'] }; XLSX.utils.book_append_sheet(wb, ws, name); };
  add(A1, 'Signalements', [11, 12, 24, 50, 12, 28, 8, 8, 10, 16, 12, 11, 11, 10, 7, 16, 16, 16, 12, 30, 40]);
  add(A2, 'Rondes', [12, 12, 16, 16, 10, 10, 30, 12, 10, 11, 11, 10, 40, 10]);
  add(A3, 'Journal', [10, 6, 16, 11, 16, 16, 12, 12, 50, 11, 11, 10, 10, 10]);
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  hideBusy();
  await shareOrDownload(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `rondes_parkings_complet_${fileStamp()}.xlsx`, 'Rondes parkings');
}
/* ---------- Archive complète (pour versement sur un stockage de la Ville) ---------- */
async function supArchive() {
  showBusy('Téléchargement des photos…');
  const entries = (await dbAll('entries')).filter(e => e.kind === 'ev' && e.media);
  const ns = [...new Set(entries.map(e => e.n))];
  for (const n of ns) {
    const ids = entries.filter(e => e.n === n).flatMap(e => e.ev.photos.map(p => p.id));
    let have = true; for (const id of ids) if (!(await dbGet('photos', id))) { have = false; break; }
    if (!have) await fetchMedia(n);
  }
  showBusy('Construction de l’archive…');
  const dump = { format: 'rondes-cachan-archive', v: 2, app: APP_VERSION, cree: new Date().toISOString(), cfg: state.cfg, sigs: await dbAll('sigs'), rondes: await dbAll('rondes'), journal: (await dbAll('entries')).map(e => ({ ...e })) };
  const files = { 'archive.json': te.encode(JSON.stringify(dump, null, 1)) };
  for (const p of await dbAll('photos')) files[`photos/${p.id}.jpg`] = [new Uint8Array(await p.blob.arrayBuffer()), { level: 0 }];
  const zip = fflate.zipSync(files);
  hideBusy();
  await shareOrDownload(new Blob([zip], { type: 'application/zip' }), `archive_rondes_${fileStamp()}.zip`, 'Archive rondes');
  toast('Archive en clair : à verser uniquement sur un stockage de la Ville.');
}

/* ---------- QR codes ---------- */
function qrSvg(text, ecc = 'M') {
  const q = qrcode(0, ecc); q.addData(text, 'Byte'); q.make();
  return q.createSvgTag({ cellSize: 4, margin: 4, scalable: true });
}
VIEWS.supQR = async ({ first } = {}) => {
  const joinURL = appBaseURL() + '#join=' + packJoin(state.conn);
  const show = h('div', { class: 'qr' });
  return page('QR codes', { back: true },
    first ? h('div', { class: 'banner ok' }, 'Relais activé. Étapes suivantes : saisir les prénoms des agents (Listes et coordonnées), imprimer les QR des parkings, configurer le téléphone des agents.') : null,
    h('div', { class: 'card' }, h('h2', null, 'Configuration des téléphones agents'),
      h('p', { class: 'muted small' }, 'À scanner une fois avec l’appareil photo de chaque téléphone agents. Les modifications de listes leur parviennent ensuite seules.'),
      h('button', { onclick: () => { show.innerHTML = qrSvg(joinURL, 'L'); } }, 'Afficher le QR de configuration'), show,
      h('p', { class: 'muted small' }, 'Ce QR donne accès aux signalements de l’équipe : ne le diffusez pas.')),
    h('div', { class: 'card' }, h('h2', null, 'Parkings'),
      h('p', { class: 'muted small' }, 'Un QR par parking, à afficher à l’entrée, plastifié, hors de portée. Scanné à l’arrivée, il démarre la ronde.'),
      ...state.cfg.sites.map(s => h('div', { class: 'stat' }, h('span', null, s.nom), h('span', { class: 'muted small' }, 'code ' + s.token.slice(0, 4) + '…'))),
      h('button', { class: 'accent', onclick: () => printSiteQR() }, 'Imprimer les QR des parkings'),
      h('button', {
        class: 'link', onclick: safe(async () => {
          if (!confirm('Générer de nouveaux codes ? Les QR affichés dans les parkings ne fonctionneront plus.')) return;
          const cfg = JSON.parse(JSON.stringify(state.cfg));
          cfg.sites.forEach(s => s.token = randCode(10)); cfg.cfgId = randCode(8);
          await publishCfg(cfg);
          toast('Nouveaux codes générés : réimprimez les QR.'); render();
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
async function publishCfg(cfg) {
  state.cfg = cfg; await kvSet('cfg', cfg);
  await queueDec('d', { type: 'cfg', ts: new Date().toISOString(), cfg });
  syncNow();
}

/* ---------- Listes et coordonnées ---------- */
VIEWS.supConfig = async () => {
  const c = JSON.parse(JSON.stringify(state.cfg));
  const ta = (lines) => h('textarea', { value: lines.join('\n'), style: 'min-height:140px' });
  const tAgents = ta(c.agents), tDests = ta(c.dests), tCats = ta(c.cats);
  const tChk = ta(c.checklist.map(x => `${x.lbl} | ${x.cat}`));
  const sitesInputs = c.sites.map(s => h('input', { type: 'text', value: s.nom }));
  const tel = h('input', { type: 'tel', value: c.urgenceTel || '', placeholder: '01 …' });
  const mail = h('input', { type: 'email', value: c.urgenceMail || '', placeholder: 'prenom.nom@ville-cachan.fr' });
  const lines = t => t.value.split('\n').map(x => x.trim()).filter(Boolean);
  return page('Listes et coordonnées', { back: true },
    h('div', { class: 'banner info' }, 'Les modifications sont transmises automatiquement à tous les appareils.'),
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
        await publishCfg(c);
        toast('Enregistré. Transmis à tous les appareils.');
        go('home');
      })
    }, 'Enregistrer'));
};

/* ====================================================================== */
/* Démarrage et synchronisation périodique                                */
/* ====================================================================== */
async function boot() {
  if (!window.isSecureContext || !crypto.subtle) {
    document.getElementById('app').replaceChildren(page('Rondes parkings', null,
      h('div', { class: 'banner bad' }, 'Cette application doit être ouverte depuis son adresse https:// (pas depuis un fichier).')));
    return;
  }
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(e => console.warn('SW', e));
  state.mode = (await kvGet('mode')) || null;
  state.conn = (await kvGet('conn')) || null;
  state.cfg = (await kvGet('cfg')) || null;
  state.priv = state.mode === 'superviseur' ? await kvGet('privKey') : null;
  if (state.mode && !state.conn) { state.mode = null; }
  if (state.mode === 'agent' && (await kvGet('ronde'))) Geo.start();
  const hash = location.hash;
  if (hash && /^#(join|site)=/.test(hash)) {
    history.replaceState(null, '', location.pathname + location.search);
    await handleLink(hash, false);
    if (state.view === 'home') render();
  } else render();
  if (state.mode) syncNow().then(refreshIfHome);
}
window.addEventListener('hashchange', () => {
  const hash = location.hash;
  if (hash && /^#(join|site)=/.test(hash)) { history.replaceState(null, '', location.pathname + location.search); handleLink(hash, false); }
});
window.addEventListener('online', () => { if (state.mode) syncNow().then(refreshIfHome); });
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && state.mode) syncNow().then(refreshIfHome); });
setInterval(() => { if (document.visibilityState === 'visible' && state.mode) syncNow().then(refreshIfHome); }, SYNC_PERIODE_MS);
boot();
