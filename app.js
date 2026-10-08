/* Rondes parkings — Ville de Cachan, DPMS — version 1.3
 * Application web hors ligne (PWA).
 * - L'outil appartient aux agents : rondes, comptage des véhicules, signalements (catégories et
 *   sous-catégories, plusieurs par catégorie, en ronde ou hors ronde), barrières laissées ouvertes.
 * - Police municipale : suivi des signalements de véhicules (observations, traité / non traité).
 * - Superviseurs : consultation, interventions ponctuelles, exports.
 * - Le relais (Google Apps Script) conserve un journal chiffré commun. Clé d'équipe : agents, PM,
 *   superviseurs. Clé superviseur : horodatage exact, positions, notes internes. Le relais ne lit rien.
 */
'use strict';

const APP_VERSION = '1.3.0';
const HISTO_JOURS = 92;            // période maximale de l'extrait Excel des agents
const PBKDF2_ITER = 600000;
const LOT_MAX_OCTETS = 4000000;    // taille maximale d'un envoi (photos comprises)
const SYNC_PERIODE_MS = 120000;
const RETRO_MIN = 10;              // écart (min) au-delà duquel une heure déclarée est « a posteriori »

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
function localDT(d = new Date()) { return `${localDate(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`; }
function fmtJour(j) { if (!j) return ''; const [y, m, d] = j.slice(0, 10).split('-'); return `${d}/${m}/${y}`; }
function fmtDT(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function fmtHeure(iso) { if (!iso) return ''; const d = new Date(iso); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; }
function fmtQuand(iso) { if (!iso) return ''; return localDate(new Date(iso)) === localDate() ? 'aujourd’hui à ' + fmtHeure(iso) : 'le ' + fmtDT(iso); }
function fmtDuree(ms) {
  if (ms == null || isNaN(ms)) return '';
  const m = Math.max(0, Math.round(ms / 60000));
  if (m < 60) return `${m} min`;
  const hh = Math.floor(m / 60), mm = m % 60;
  if (hh < 48) return `${hh} h ${pad(mm)}`;
  return `${Math.floor(hh / 24)} j ${hh % 24} h`;
}
function depuisJours(jour) {
  const n = Math.round((new Date(localDate() + 'T00:00:00') - new Date(jour.slice(0, 10) + 'T00:00:00')) / 86400000);
  return n <= 0 ? 'aujourd’hui' : n === 1 ? 'hier' : `il y a ${n} jours`;
}
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
function xlsxBlob(sheets) {
  const wb = XLSX.utils.book_new();
  for (const [name, aoa, widths] of sheets) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = widths.map(w => ({ wch: w }));
    ws['!autofilter'] = { ref: ws['!ref'] };
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  return new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/* ====================================================================== */
/* Base de données locale (IndexedDB)                                     */
/* ====================================================================== */
const DB_NAME = 'rondes-cachan', DB_VER = 2;
const STORES = {
  kv: null,
  events: { keyPath: 'seq' },     // saisies pas encore publiées sur le relais
  photos: { keyPath: 'id' },      // photos pas encore publiées ; ou téléchargées pour affichage
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
    r.onupgradeneeded = (ev) => {
      const d = r.result;
      // Passage à la 1.2 : nouveau format ; les données d'essai des versions antérieures sont effacées.
      [...d.objectStoreNames].forEach(n => d.deleteObjectStore(n));
      for (const [n, o] of Object.entries(STORES)) o ? d.createObjectStore(n, o) : d.createObjectStore(n);
      if (ev.oldVersion > 0) r.transaction.objectStore('kv').put(true, 'migre');
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
  if (!j.ok) {
    const msg = j.err === 'deja-initialise' ? 'Ce relais est déjà activé par un superviseur.'
      : j.err === 'aucun superviseur enregistré' ? 'Ce relais a été activé, mais l’activation n’est pas allée à son terme (aucune clé enregistrée). Réinitialisez le relais (fonction reinitialiserRelais), puis activez-le à nouveau.'
        : 'relais : ' + (j.err || 'refus');
    const e = new Error(msg); e.code = j.err; throw e;
  }
  return j;
}
function rel(op, extra = {}, timeout) {
  const c = state.conn;
  return relayCall(c.url, { op, tok: state.mode === 'superviseur' ? c.supTok : c.agentTok, ...extra }, timeout);
}

/* ====================================================================== */
/* Configuration : parkings, agents, destinataires, catégories            */
/* Format des catégories (modifiable par le superviseur) :                */
/*   Catégorie | destinataire suggéré, …                                  */
/*   - sous-catégorie            (« ! » final = urgent par défaut)        */
/* ====================================================================== */
const TAXO_DEFAUT = `Sécurité incendie | Ateliers
- Extincteur absent ou déplacé
- Extincteur utilisé, vide ou scellé rompu
- Extincteur : contrôle périodique dépassé (étiquette)
- Accès à un extincteur encombré
- Bac d’absorbant vide ou absent
- Déclencheur manuel d’alarme cassé, masqué ou enclenché
- Alarme incendie en dérangement ou qui sonne !
- Porte coupe-feu calée ouverte ou ferme-porte hors service
- Commande de désenfumage dégradée ou inaccessible
- Bouche de ventilation ou de désenfumage obstruée
- Colonne sèche : prise dégradée ou capot manquant
- Plans ou consignes de sécurité absents ou illisibles
- Dépôt de matières combustibles ou d’encombrants
- Odeur de fumée, de brûlé ou de carburant !
Issues de secours et évacuation | Ateliers
- Issue de secours encombrée !
- Issue de secours verrouillée ou bloquée !
- Barre antipanique ou porte d’issue défectueuse
- Éclairage de sécurité (blocs, balisage) éteint ou cassé
- Signalétique d’évacuation absente ou masquée
- Escalier encombré, sale ou dégradé
Éclairage et électricité | Ateliers
- Luminaire(s) éteint(s)
- Zone entièrement dans le noir
- Luminaire cassé ou pendant
- Câble électrique apparent ou arraché !
- Armoire ou coffret électrique ouvert
Ventilation et qualité de l’air | Ateliers
- Ventilation à l’arrêt ou bruit anormal
- Détecteur ou alarme de pollution (CO) en défaut
- Odeur persistante de gaz d’échappement
Barrières, bornes et caisses | Skidata
- Barrière en panne (bloquée ouverte)
- Barrière en panne (ne s’ouvre pas)
- Lisse de barrière cassée
- Borne d’entrée hors service
- Borne de sortie hors service
- Caisse automatique hors service
- Caisse forcée ou dégradée !
- Interphone d’appel hors service
- Affichage des places disponibles erroné
Accès piétons et ascenseurs | Ateliers
- Ascenseur en panne
- Alarme ou téléphone de cabine d’ascenseur hors service
- Porte d’accès piéton forcée ou serrure hors service
- Porte ou rideau d’accès véhicules en panne
- Lecteur de badge ou digicode hors service
- Vitre ou porte vitrée cassée
Propreté et hygiène
- Déchets ou détritus
- Dépôt sauvage ou encombrants
- Urine ou déjections
- Seringues ou matériel de consommation (ne pas toucher) !
- Tags ou graffitis
- Flaque d’huile ou d’hydrocarbures
- Nuisibles (rats, pigeons)
- Poubelles pleines
Bâtiment et infiltrations | Ateliers
- Fuite d’eau ou infiltration
- Inondation ou eau stagnante
- Fissure, éclat de béton ou ferraille apparente
- Chute de matériaux du plafond !
- Garde-corps ou main courante dégradé
- Avaloir ou caniveau bouché
Signalisation et marquage | Ateliers
- Panneau de hauteur maximale absent ou endommagé
- Gabarit de hauteur arraché ou tordu
- Marquage au sol effacé
- Signalisation des places PMR effacée
- Panneau directionnel ou sens de circulation absent
- Numérotation des places ou des niveaux illisible
Véhicules | Police municipale
- Véhicule ventouse (stationnement prolongé)
- Épave ou véhicule hors d’usage
- Véhicule mal stationné ou gênant
- Véhicule sur place PMR sans carte
- Véhicule sur place de recharge sans recharger
- Véhicule fracturé ou vitre brisée
- Fuite sous un véhicule
- Personne dans un véhicule
Bornes de recharge électrique | Ateliers
- Borne hors service
- Câble de recharge arraché ou dégradé
- Extincteur à proximité absent
- Chaleur, fumée ou odeur de brûlé !
Sûreté et présences | Police municipale
- Personnes installées (squat, regroupement)
- Personne endormie ou en difficulté !
- Comportement suspect
- Trace d’effraction
- Caméra de vidéoprotection dégradée ou masquée
- Vandalisme en cours (appeler le 17) !`;
const MOTIFS_DEFAUT = ['Panne de la barrière', 'Panne de borne ou de caisse', 'Intervention technique', 'Forte affluence', 'Consigne de la hiérarchie', 'Autre'];
const CAT_AUTRE = 'Autre';
const SUB_AUTRE = 'Autre (préciser)';
const VEHICULE_RE = /v[ée]hicule/i;
const PM_RE = /police/i;

function parseTaxo(text) {
  const cats = [];
  for (const raw of text.split('\n')) {
    const line = raw.trim();
    if (!line) continue;
    if (/^[-•*]/.test(line)) {
      if (!cats.length) continue;
      let lbl = line.replace(/^[-•*]\s*/, '').trim(), urgent = false;
      if (lbl.endsWith('!')) { urgent = true; lbl = lbl.slice(0, -1).trim(); }
      if (lbl) cats[cats.length - 1].subs.push({ lbl, urgent });
    } else {
      const [nom, d] = line.split('|');
      cats.push({ nom: nom.trim(), dests: (d || '').split(',').map(x => x.trim()).filter(Boolean), subs: [] });
    }
  }
  return cats.filter(c => c.nom && c.nom !== CAT_AUTRE);
}
function taxoToText(cats) {
  return cats.map(c => [c.nom + (c.dests.length ? ' | ' + c.dests.join(', ') : ''), ...c.subs.map(s => '- ' + s.lbl + (s.urgent ? ' !' : ''))].join('\n')).join('\n');
}
function defaultConfig() {
  return normalizeCfg({
    v: 3, cfgId: randCode(8),
    sites: [
      { id: 'henouille', code: 'HEN', nom: 'Hénouille', token: randCode(10) },
      { id: 'dumotel', code: 'DUM', nom: 'Dumotel', token: randCode(10) },
      { id: 'arobase', code: 'ARO', nom: 'Arobase', token: randCode(10) }
    ],
    agents: ['Agent 1', 'Agent 2', 'Agent 3', 'Agent 4', 'Vacataire'],   // prénoms saisis dans « Listes et coordonnées », jamais dans le code publié
    dests: ['Ateliers', 'DST', 'Skidata', 'Police municipale', 'DPMS'],
    urgenceTel: '', urgenceMail: ''
  });
}
/* Mise au format courant (déterministe : tous les appareils obtiennent le même résultat) */
function normalizeCfg(c) {
  if (!c) return c;
  const n = JSON.parse(JSON.stringify(c));
  if (!Array.isArray(n.cats) || !n.cats.length || typeof n.cats[0] === 'string') n.cats = parseTaxo(TAXO_DEFAUT);
  delete n.checklist;
  n.sites = (n.sites || []).map(s => ({ ...s, barrieres: Array.isArray(s.barrieres) && s.barrieres.length ? s.barrieres : ['Entrée', 'Sortie'] }));
  if (!Array.isArray(n.motifs) || !n.motifs.length) n.motifs = MOTIFS_DEFAUT.slice();
  n.v = 3;
  return n;
}
function siteById(id) { return state.cfg && state.cfg.sites.find(s => s.id === id); }
function siteNom(id) { const s = siteById(id); return s ? s.nom : id; }
function catByNom(nom) { return (state.cfg.cats || []).find(c => c.nom === nom) || null; }
function isVehicule(cat) { return VEHICULE_RE.test(cat || ''); }
function sigTitre(s) { return s.sub ? `${s.cat} — ${s.sub}` : s.cat; }
function validCfg(c) { return c && Array.isArray(c.sites) && Array.isArray(c.agents) && Array.isArray(c.dests) && Array.isArray(c.cats); }
/* QR de configuration : accès au relais, clés d'équipe et publique ; rôle agents ou police municipale */
function packJoin(c, role) { return b64u.enc(fflate.deflateSync(te.encode(JSON.stringify({ v: 2, url: c.url, tok: c.agentTok, team: c.team, pub: c.pub, role: role || 'agent' })), { level: 9 })); }
function unpackJoin(s) { return JSON.parse(td.decode(fflate.inflateSync(b64u.dec(s)))); }

/* ====================================================================== */
/* Géolocalisation : uniquement pendant une saisie ou une ronde           */
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
async function geoStopSiLibre() { if (!(await kvGet('ronde'))) Geo.stop(); }

/* ====================================================================== */
/* État de l'application et rendu                                         */
/* ====================================================================== */
const state = { mode: null, conn: null, cfg: null, priv: null, view: 'home', params: {}, cleanup: null, sync: { busy: false } };
const estEquipe = () => state.mode === 'agent' || state.mode === 'pm';

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
  const tag = { agent: 'Agents', superviseur: 'Superviseur DPMS', pm: 'Police municipale' }[state.mode];
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
const VUES_SAISIE = ['signalement', 'barriere'];
window.addEventListener('popstate', () => {
  if (VUES_SAISIE.includes(state.view) && !confirm('Abandonner cette saisie ?')) { history.pushState({ v: state.view }, ''); return; }
  if (VUES_SAISIE.includes(state.view)) { sigDraft = null; barrDraft = null; }
  if (VUES_SAISIE.includes(state.view) && state.params && state.params.from === 'ronde') { go('ronde', {}, { noPush: true }); return; }
  if (state.view !== 'home') go('home', {}, { noPush: true });
});
/* Boutons-puces : choix unique ou multiple */
function chips(options, selected, onChange, opts = {}) {
  const box = h('div', { class: 'chips' });
  const draw = () => box.replaceChildren(...options.map(o => {
    const val = typeof o === 'object' ? o.val : o, lbl = typeof o === 'object' ? o.lbl : o;
    const on = opts.multi ? selected.includes(val) : selected === val;
    return h('button', {
      class: 'chip' + (on ? (opts.cls || ' on') : ''),
      onclick: () => {
        if (opts.multi) { const i = selected.indexOf(val); if (i >= 0) selected.splice(i, 1); else selected.push(val); }
        else selected = val;
        draw(); onChange(selected);
      }
    }, lbl);
  }));
  draw();
  box.set = v => { selected = v; draw(); };
  return box;
}

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
const PM_STATUTS = { obs: 'Observation', traite: 'Traité', non_traite: 'Non traité' };
function statutBadge(s) { return h('span', { class: 'badge ' + (s === 'clos' ? 'b-ok' : 'b-warn') }, STATUTS[s] || s); }
function applyVerdict(sig, verdict, quand, agent) {
  if (verdict === 'resolu') { sig.statut = 'clos'; sig.closLe = quand; sig.closPar = agent || 'agent'; }
  else { sig.statut = 'ouvert'; sig.closLe = null; sig.closPar = null; }
  if (verdict === 'aggrave') sig.aggrave = true;
}
function pmEtat(s) {
  if (s.pm && s.pm.statut === 'traite') return 'Traité';
  if (s.pm && s.pm.statut === 'non_traite') return 'Non traité';
  if (s.statut === 'clos') return 'Clos';
  return 'À traiter';
}
function pmBadge(s) {
  const e = pmEtat(s);
  return h('span', { class: 'badge ' + (e === 'Traité' ? 'b-ok' : e === 'Non traité' ? 'b-bad' : e === 'Clos' ? 'b-info' : 'b-warn') }, 'PM : ' + e);
}
function pmBanner(s) {
  if (!s.pm || !s.pm.hist.length) return null;
  const l = s.pm.hist[s.pm.hist.length - 1];
  return h('div', { class: 'banner ' + (s.pm.statut === 'non_traite' ? 'warn' : 'info') }, `Police municipale (${fmtJour(l.jour)}) : ${PM_STATUTS[l.statut]}${l.obs ? ' — ' + l.obs : ''}`);
}

/* ====================================================================== */
/* Reconstitution de l'état à partir du journal                           */
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

  const sigs = new Map(), rondes = new Map(), barr = new Map(), devs = {}, chains = {};
  let cfg = null;
  const alertes = [];
  for (const { e } of items) {
    if (e.kind === 'ev') {
      const ev = e.ev, dec = e.dec || null, ts = dec && dec.ts;
      if (sup) {
        (chains[ev.dev] = chains[ev.dev] || []).push(e);
        const dv = devs[ev.dev] = devs[ev.dev] || { seq: 0, at: null, role: ev.role || 'agent' };
        if (ev.seq > dv.seq) { dv.seq = ev.seq; dv.at = e.at; }
      }
      if (ev.t === 'signalement') {
        if (!sigs.has(ev.data.id)) sigs.set(ev.data.id, {
          id: ev.data.id, ref: ev.data.ref, site: ev.site, cat: ev.data.cat, sub: ev.data.sub || '', desc: ev.data.desc,
          plaque: ev.data.plaque || '', emplacement: ev.data.emplacement || '', vehicule: ev.data.vehicule || '',
          dests: ev.data.dests, urgent: ev.data.urgent, horsRonde: !!ev.data.horsRonde, statut: 'ouvert', jour: ev.jour, agent: ev.agent,
          thumb: e.local && typeof ev.data.thumb === 'string' ? ev.data.thumb : null,
          photos: (ev.photos || []).map(p => ({ id: p.id, sha: p.sha, n: e.n })), ts, geo: dec && dec.geo, dev: ev.dev, seq: ev.seq,
          suivi: [], journalDPMS: [], notes: '', notesAgents: '', pm: null
        });
      } else if (ev.t === 'revue') {
        const s = sigs.get(ev.data.sig);
        if (s) {
          s.suivi.push({ jour: ev.jour, ts, agent: ev.agent, verdict: ev.data.verdict, comment: ev.data.comment, photos: (ev.photos || []).map(p => ({ id: p.id, sha: p.sha, n: e.n })), geo: dec && dec.geo });
          applyVerdict(s, ev.data.verdict, ts || ev.jour, ev.agent);
        } else if (sup) alertes.push({ key: `orph:${ev.dev}:${ev.seq}`, at: e.at, msg: `${siteNom(ev.site)}, ${ev.agent}, ${fmtJour(ev.jour)} — constat sur un signalement inconnu (${ev.data.ref})` });
      } else if (ev.t === 'pm') {
        const s = sigs.get(ev.data.sig);
        if (s) {
          s.pm = s.pm || { statut: null, hist: [] };
          s.pm.hist.push({ jour: ev.jour, ts, statut: ev.data.statut, obs: ev.data.obs || '', agent: ev.agent });
          if (ev.data.statut !== 'obs') s.pm.statut = ev.data.statut;
          if (ev.data.statut === 'traite') { s.statut = 'clos'; s.closLe = ts || ev.jour; s.closPar = 'Police municipale'; }
          else if (ev.data.statut === 'non_traite' && s.closPar === 'Police municipale') { s.statut = 'ouvert'; s.closLe = null; s.closPar = null; }
        }
      } else if (ev.t === 'barriere') {
        const d = ev.data;
        const b = barr.get(d.bid) || { bid: d.bid, site: ev.site, barriere: d.barriere, ouverte: null, fermee: null };
        const rec = { quand: d.quand, agent: ev.agent, motif: d.motif || '', comment: d.comment || '', ts, dev: ev.dev };
        if (d.action === 'ouverte' && !b.ouverte) b.ouverte = rec;
        if (d.action === 'fermee' && !b.fermee) b.fermee = rec;
        if (b.ouverte) barr.set(d.bid, b);
      }
      if (sup && ev.rid) {
        const r = rondes.get(ev.rid) || { rid: ev.rid, site: ev.site, agent: ev.agent, dev: ev.dev, jour: ev.jour, nbSig: 0, nbRevue: 0, compte: null, alertes: [] };
        if (ev.t === 'ronde_debut') { r.debut = ts; r.qr = ev.data.qr; r.geoDebut = dec && dec.geo; }
        if (ev.t === 'ronde_fin') { r.fin = ts; if (r.compte == null && ev.data.compte != null) r.compte = ev.data.compte; }
        if (ev.t === 'checklist') r.checklist = ev.data.items;
        if (ev.t === 'comptage') r.compte = ev.data.total;
        if (ev.t === 'signalement') r.nbSig++;
        if (ev.t === 'revue') r.nbRevue++;
        rondes.set(ev.rid, r);
      }
    } else if (e.kind === 'dec') {
      const d = e.dec;
      if (d.type === 'cfg') { if (validCfg(normalizeCfg(d.cfg))) cfg = normalizeCfg(d.cfg); }
      else if (d.type === 'sig') {
        const s = sigs.get(d.sig);
        if (s) { Object.assign(s, d.patch); s.journalDPMS.push({ ts: d.ts, action: d.action, par: d.par || 'DPMS' }); }
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
        const ev = e.ev, lieu = `${siteNom(ev.site) || '—'}, ${ev.agent || '?'}, ${fmtJour(ev.jour)}`;
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
  await kvSet('barrieres', [...barr.values()]);
  if (cfg && JSON.stringify(cfg) !== JSON.stringify(state.cfg)) { state.cfg = cfg; await kvSet('cfg', cfg); }
}

/* ====================================================================== */
/* Saisies : journal chaîné local (agents et police municipale)           */
/* ====================================================================== */
async function appendEvent(type, base, data, secretExtra = {}, photos = []) {
  return withLock(async () => {
    const head = (await kvGet('head')) || { seq: 0, hash: '0'.repeat(64) };
    const now = new Date();
    const secret = { ts: now.toISOString(), tzo: now.getTimezoneOffset(), geo: Geo.snap(), ...secretExtra };
    const ev = {
      v: 3, dev: await kvGet('dev'), role: state.mode, seq: head.seq + 1, t: type, jour: localDate(now),
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
async function storeEvEntry(ev, n, idx, at, extra = {}) {
  const k = `e:${ev.dev}:${String(ev.seq).padStart(8, '0')}`;
  const old = await dbGet('entries', k);
  if (old && old.ev.hash === ev.hash && old.n <= n) return false;
  if (ev.t === 'signalement' && typeof ev.data.thumb === 'string') await dbPut('thumbs', { id: ev.data.id, data: ev.data.thumb });
  const light = ev.t === 'signalement' && typeof ev.data.thumb === 'string' ? { ...ev, data: { ...ev.data, thumb: true } } : ev;
  await dbPut('entries', { k, kind: 'ev', n, idx, at, ev: light, ...extra });
  return true;
}
async function agentPrefere() {
  const r = await kvGet('ronde');
  if (r) return r.agent;
  const l = await kvGet('dernierAgent');
  return l && Date.now() - l.at < 12 * 3600000 ? l.nom : null;
}

/* ---------- Synchronisation ---------- */
let _syncP = null, _syncAgain = null;
function syncNow() {
  if (_syncP) { if (!_syncAgain) _syncAgain = _syncP.then(() => { _syncAgain = null; return syncNow(); }); return _syncAgain; }
  _syncP = (async () => {
    await null;
    state.sync.busy = true;
    try {
      if (!state.conn) return { skipped: true };
      if (!navigator.onLine) return { offline: true };
      const r = estEquipe() ? await agentSync() : await supSync();
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
    const data = await encJSON({ v: 3, dev, events: batch });
    const media = Object.keys(files).length ? b64u.enc(await teamEnc(fflate.zipSync(files))) : null;
    const r = await rel('put', { dev, data, media }, 180000);
    for (let i = 0; i < batch.length; i++) await storeEvEntry(batch[i], r.n, i, r.at, { media: !!media });
    const d = await openDB(); const t = d.transaction(['events', 'photos'], 'readwrite');
    batch.forEach(ev => { t.objectStore('events').delete(ev.seq); ev.photos.forEach(p => t.objectStore('photos').delete(p.id)); });
    await txDone(t);
    sent += batch.length;
  }
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
  if (state.mode === 'agent') return agentHome();
  if (state.mode === 'pm') return pmHome();
  return supHome();
};

/* ---------- Premier lancement ---------- */
async function viewSetup() {
  const pending = state.pendingJoin;
  const migre = await kvGet('migre');
  if (pending) {
    const pm = pending.role === 'pm';
    return page('Configuration', null,
      h('div', { class: 'card' },
        h('h3', null, pm ? 'Configurer cet appareil pour la police municipale ?' : 'Configurer cet appareil comme téléphone des agents ?'),
        h('p', { class: 'muted' }, pm ? 'Il affichera les signalements de véhicules, avec plaques et photos.' : 'Il se connectera au relais et récupérera la configuration et les signalements en cours.'),
        h('button', { class: 'ok', onclick: safe(async () => { await setupEquipe(pending); }) }, 'Oui, configurer'),
        h('button', { class: 'sec', onclick: () => { state.pendingJoin = null; render(); } }, 'Annuler')));
  }
  return page('Rondes parkings', null,
    migre ? h('div', { class: 'banner info' }, 'Application mise à jour : cet appareil doit être configuré à nouveau. Les données déjà envoyées au relais seront retrouvées.') : null,
    h('div', { class: 'card' },
      h('h3', null, 'Téléphone des agents ou police municipale'),
      h('p', { class: 'muted' }, 'Sur un appareil superviseur, ouvrez « QR codes » puis scannez le QR correspondant avec cet appareil.'),
      h('button', { class: 'big', onclick: () => go('scan') }, 'Scanner un QR de configuration')),
    h('div', { class: 'card' },
      h('h3', null, 'Appareil superviseur (DPMS)'),
      h('p', { class: 'muted' }, 'Téléphone ou ordinateur. Plusieurs appareils superviseurs peuvent être utilisés.'),
      h('button', { class: 'sec', onclick: () => go('supJoin') }, 'Ajouter cet appareil comme superviseur'),
      h('button', { class: 'sec', onclick: () => go('supInit') }, 'Premier superviseur : activer un relais neuf')),
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
}
async function setupEquipe(j) {
  if (!j || !j.url || !j.tok || !j.team || !j.pub) throw new Error('QR de configuration invalide.');
  if (state.mode === 'superviseur') throw new Error('Cet appareil est superviseur.');
  const role = j.role === 'pm' ? 'pm' : 'agent';
  if (estEquipe() && state.mode === role && JSON.stringify(state.conn.pub) === JSON.stringify(j.pub) && state.conn.url === j.url) {
    state.pendingJoin = null; toast('Cet appareil est déjà configuré.'); go('home'); return;
  }
  if (estEquipe() && ((await dbAll('events')).length)) throw new Error('Des saisies ne sont pas encore envoyées : envoyez-les avant de changer de configuration.');
  showBusy('Connexion au relais…');
  const conn = { url: j.url, agentTok: j.tok, team: j.team, pub: j.pub };
  await relayCall(conn.url, { op: 'ping', tok: conn.agentTok });
  if (estEquipe()) { for (const s of ['entries', 'thumbs', 'sigs', 'photos']) await dbClear(s); await kvDel('cursor'); await kvDel('head'); }
  await kvSet('dev', (role === 'pm' ? 'PM-' : 'T-') + randCode(6));
  await kvSet('conn', conn); await kvSet('mode', role); await kvDel('migre');
  state.mode = role; state.conn = conn; state.pendingJoin = null;
  showBusy('Récupération de la configuration et des signalements…');
  const r = await syncNow();
  hideBusy();
  await requestPersist();
  if (r && r.err) toast('Configuré, mais la récupération a échoué : ' + r.err, true);
  else toast(state.cfg ? 'Appareil configuré.' : 'Configuré, mais aucune configuration trouvée sur le relais.', !state.cfg);
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
  const m = /^#(join|site|cfg)=(.+)$/.exec(hash);
  return m ? { kind: m[1], val: m[2] } : null;
}
async function handleLink(text, fromScan) {
  const l = parseLink(text);
  if (!l) { toast('QR code non reconnu.', true); if (fromScan) go('home'); return; }
  if (l.kind === 'cfg') { toast('Ce QR a été produit par une ancienne version de l’application. Rechargez la page sur l’appareil qui l’affiche, puis affichez un nouveau QR.', true); go('home'); return; }
  if (l.kind === 'join') {
    let j; try { j = unpackJoin(l.val); } catch (e) { toast('QR de configuration illisible.', true); go('home'); return; }
    if (state.mode === 'superviseur') { toast('Cet appareil est superviseur : configuration ignorée.', true); go('home'); return; }
    if (estEquipe()) { if (confirm('Reconfigurer cet appareil avec ce QR ?')) await safe(setupEquipe)(j); else go('home'); return; }
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
async function barrieresOuvertes() { return ((await kvGet('barrieres')) || []).filter(b => b.ouverte && !b.fermee).sort((a, b) => a.ouverte.quand.localeCompare(b.ouverte.quand)); }
function barriereLigne(b, withBtn) {
  return h('div', { class: 'stat', style: 'display:block' },
    h('div', null, h('b', null, `${siteNom(b.site)} — barrière ${b.barriere}`), h('span', { class: 'badge b-bad' }, 'ouverte')),
    h('div', { class: 'muted small' }, `depuis ${fmtQuand(b.ouverte.quand)} (${fmtDuree(Date.now() - new Date(b.ouverte.quand))}) — ${b.ouverte.agent || ''}${b.ouverte.motif ? ' — ' + b.ouverte.motif : ''}`),
    withBtn ? h('button', { class: 'ok', onclick: () => go('barriere', { bid: b.bid }) }, 'Barrière refermée') : null);
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
  const ouvertes = await barrieresOuvertes();
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
    h('div', { class: 'row' },
      h('button', { class: 'sec', onclick: () => go('signalement', { from: 'hr' }) }, 'Signalement hors ronde'),
      h('button', { class: 'sec', onclick: () => go('barriere', {}) }, 'Barrière ouverte')),
    ouvertes.length ? h('div', { class: 'card', style: 'border-color:var(--bad)' }, h('h2', null, 'Barrières ouvertes'), ...ouvertes.map(b => barriereLigne(b, true))) : null,
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
  const ronde = { rid: uuid(), site, agent, qr, debut: new Date().toISOString(), step: 'revue', checklist: {}, sigsCrees: [], compte: 0 };
  await appendEvent('ronde_debut', ronde, { qr }, { qr });
  await kvSet('ronde', ronde);
  await kvSet('dernierAgent', { nom: agent, at: Date.now() });
  hideBusy();
  resumeRonde(ronde);
}
function resumeRonde(ronde) {
  Geo.start();
  go(ronde.step === 'revue' ? 'revue' : 'ronde');
}

/* ---------- Revue des signalements ouverts ---------- */
const revueDraft = {};
VIEWS.revue = async () => {
  const ronde = await kvGet('ronde');
  if (!ronde) return agentHome();
  const open = (await dbAll('sigs')).filter(s => s.site === ronde.site && s.statut === 'ouvert')
    .sort((a, b) => (a.jour || '').localeCompare(b.jour || ''));
  if (!open.length) { ronde.step = 'ronde'; await kvSet('ronde', ronde); return VIEWS.ronde(); }
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
      h('h3', null, sigTitre(s), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('p', { class: 'muted small' }, `${s.ref} — signalé le ${fmtJour(s.jour)}${s.agent ? ' par ' + s.agent : ''}`),
      s.desc ? h('p', null, s.desc) : null,
      s.plaque || s.emplacement ? h('p', null, s.plaque ? [h('b', null, 'Plaque : '), s.plaque, ' '] : null, s.emplacement ? [h('b', null, 'Emplacement : '), s.emplacement] : null) : null,
      th ? h('div', { class: 'thumbs' }, h('img', { src: th, alt: 'Photo du signalement' })) : null,
      s.notesAgents ? h('div', { class: 'banner info' }, 'DPMS : ' + s.notesAgents) : null,
      pmBanner(s),
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
    ronde.step = 'ronde'; await kvSet('ronde', ronde);
    hideBusy(); go('ronde');
  });
  refresh();
  return page(`Revue : ${siteNom(ronde.site)}`, null,
    h('div', { class: 'banner info' }, `${open.length} signalement(s) en cours sur ce parking. Indiquez leur état aujourd’hui. « Résolu » clôt le signalement.`),
    ...cards, h('div', { style: 'height:70px' }),
    h('div', { class: 'sticky-bottom' }, h('div', null, btnValider)));
};

/* ---------- Ronde : compteur, check-list par catégorie, signalements ---------- */
VIEWS.ronde = async () => {
  const ronde = await kvGet('ronde');
  if (!ronde) return agentHome();
  if (ronde.compte == null) ronde.compte = 0;
  const cfg = state.cfg;
  const sigs = (await dbAll('sigs')).filter(s => ronde.sigsCrees.includes(s.id));
  const parCat = {};
  sigs.forEach(s => { (parCat[s.cat] = parCat[s.cat] || []).push(s); });

  // Compteur de véhicules
  const nb = h('div', { class: 'compteur' }, String(ronde.compte));
  const majCompte = async (v) => { ronde.compte = Math.max(0, v); nb.textContent = ronde.compte; await kvSet('ronde', ronde); };
  const compteur = h('div', { class: 'card' },
    h('h2', null, 'Véhicules comptés'),
    nb,
    h('button', { class: 'plus1', onclick: safe(async () => { await majCompte(ronde.compte + 1); if (navigator.vibrate) navigator.vibrate(30); }) }, '+1 véhicule'),
    h('div', { class: 'row' },
      h('button', { class: 'sec', onclick: safe(() => majCompte(ronde.compte - 1)) }, '−1'),
      h('button', { class: 'sec', onclick: safe(async () => { const v = prompt('Nombre de véhicules :', ronde.compte); if (v !== null && /^\d+$/.test(v.trim())) await majCompte(parseInt(v, 10)); }) }, 'Corriger')));

  // Check-list : une ligne par catégorie
  const rows = cfg.cats.map(c => {
    const list = parCat[c.nom] || [];
    const ras = ronde.checklist[c.nom] === 'RAS' && !list.length;
    const etat = list.length ? h('span', { class: 'badge b-bad' }, `${list.length} signalement${list.length > 1 ? 's' : ''}`) : ras ? h('span', { class: 'badge b-ok' }, 'RAS') : h('span', { class: 'badge b-warn' }, 'à contrôler');
    return h('div', { class: 'chk' },
      h('span', { class: 'lbl' }, c.nom, ' ', etat),
      list.length ? null : h('button', {
        class: ras ? 'sel-ras' : '', onclick: safe(async () => {
          if (ras) delete ronde.checklist[c.nom]; else ronde.checklist[c.nom] = 'RAS';
          await kvSet('ronde', ronde); go('ronde', {}, { noPush: true, keepScroll: true });
        })
      }, 'RAS'),
      h('button', { class: list.length ? 'sel-ano' : '', onclick: () => go('signalement', { from: 'ronde', cat: c.nom }) }, '+ Signaler'));
  });
  const restant = cfg.cats.filter(c => !(parCat[c.nom] || []).length && ronde.checklist[c.nom] !== 'RAS').length;
  const autres = (parCat[CAT_AUTRE] || []).length;

  return page(`Ronde : ${siteNom(ronde.site)}`, { back: () => go('home') },
    h('p', { class: 'muted' }, `${ronde.agent} — commencée à ${fmtHeure(ronde.debut)}`),
    compteur,
    h('div', { class: 'card' }, h('h2', null, `Contrôles (${cfg.cats.length - restant}/${cfg.cats.length})`),
      h('p', { class: 'muted small' }, '« RAS » si rien à signaler. « + Signaler » autant de fois que nécessaire.'),
      ...rows),
    h('div', { class: 'row' },
      h('button', { class: 'sec', onclick: () => go('signalement', { from: 'ronde', cat: CAT_AUTRE }) }, `+ Autre signalement${autres ? ` (${autres})` : ''}`),
      h('button', { class: 'sec', onclick: () => go('barriere', { site: ronde.site, from: 'ronde' }) }, 'Barrière ouverte')),
    h('button', {
      class: 'ok big', onclick: safe(async () => {
        if (restant) { toast(`Encore ${restant} contrôle(s) à faire : « RAS » ou « + Signaler ».`, true); return; }
        if (!confirm(`Terminer la ronde ?\nVéhicules comptés : ${ronde.compte}${ronde.compte === 0 ? ' (aucun)' : ''}`)) return;
        await finishRonde(ronde);
      })
    }, 'Terminer la ronde'));
};
async function finishRonde(ronde, silent) {
  showBusy('Clôture de la ronde…');
  const items = {};
  const sigs = (await dbAll('sigs')).filter(s => ronde.sigsCrees.includes(s.id));
  for (const c of state.cfg.cats) items[c.nom] = sigs.some(s => s.cat === c.nom) ? 'Anomalie' : (ronde.checklist[c.nom] || 'Non contrôlé');
  await appendEvent('checklist', ronde, { items });
  await appendEvent('comptage', ronde, { total: ronde.compte || 0 });
  await appendEvent('ronde_fin', ronde, { nbSig: ronde.sigsCrees.length, compte: ronde.compte || 0 });
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

/* ---------- Signalement (en ronde ou hors ronde) ---------- */
let sigDraft = null;
VIEWS.signalement = async ({ cat, from }) => {
  const ronde = from === 'ronde' ? await kvGet('ronde') : null;
  if (from === 'ronde' && !ronde) return agentHome();
  const cfg = state.cfg;
  const cle = `${from}|${cat || ''}`;
  if (!sigDraft || sigDraft.cle !== cle) {
    sigDraft = { cle, site: ronde ? ronde.site : (state.dernierSite || (cfg.sites.length === 1 ? cfg.sites[0].id : null)), agent: ronde ? ronde.agent : await agentPrefere(),
      cat: cat || null, sub: null, desc: '', plaque: '', emplacement: '', vehicule: '', photos: [], dests: [], urgent: false, destsTouche: false };
    if (cat) appliquerCat(sigDraft, cat);
  }
  if (from !== 'ronde') Geo.start();
  const d = sigDraft;
  const catsNoms = [...cfg.cats.map(c => c.nom), CAT_AUTRE];

  const subsBox = h('div');
  const vehBox = h('div', null,
    h('label', { class: 'f' }, 'Plaque d’immatriculation'),
    h('input', { type: 'text', value: d.plaque, placeholder: 'AB-123-CD', autocapitalize: 'characters', oninput: e => d.plaque = e.target.value }),
    h('label', { class: 'f' }, 'Emplacement (numéro de place, niveau)'),
    h('input', { type: 'text', value: d.emplacement, placeholder: 'Ex. : place 112, niveau -1', oninput: e => d.emplacement = e.target.value }),
    h('label', { class: 'f' }, 'Véhicule (marque, modèle, couleur)'),
    h('input', { type: 'text', value: d.vehicule, placeholder: 'Ex. : Renault Clio grise', oninput: e => d.vehicule = e.target.value }));
  const destBox = chips(cfg.dests, d.dests, () => { d.destsTouche = true; }, { multi: true });
  const urg = h('button', { class: 'chip' + (d.urgent ? ' on-bad' : ''), onclick: e => { d.urgent = !d.urgent; d.urgentAuto = false; e.target.className = 'chip' + (d.urgent ? ' on-bad' : ''); } }, 'Urgent');
  const descLbl = h('label', { class: 'f' });
  const drawSubs = () => {
    const c = catByNom(d.cat);
    vehBox.style.display = isVehicule(d.cat) ? '' : 'none';
    descLbl.textContent = d.sub === SUB_AUTRE || d.cat === CAT_AUTRE ? 'Description (obligatoire)' : 'Précisions (facultatif)';
    if (!d.cat || d.cat === CAT_AUTRE) { subsBox.replaceChildren(); return; }
    const opts = [...(c ? c.subs.map(s => s.lbl) : []), SUB_AUTRE];
    subsBox.replaceChildren(h('label', { class: 'f' }, 'Quoi ?'), chips(opts, d.sub, v => {
      d.sub = v;
      const s = c && c.subs.find(x => x.lbl === v);
      if (s && s.urgent) { d.urgent = true; d.urgentAuto = true; } else if (d.urgentAuto) { d.urgent = false; d.urgentAuto = false; }
      urg.className = 'chip' + (d.urgent ? ' on-bad' : '');
      descLbl.textContent = v === SUB_AUTRE ? 'Description (obligatoire)' : 'Précisions (facultatif)';
    }));
  };
  drawSubs();
  const catBox = h('div');
  const drawCats = () => {
    if (d.cat && !d.choixCat) {
      catBox.replaceChildren(h('div', { class: 'chips' }, h('button', { class: 'chip on' }, d.cat), h('button', { class: 'link', style: 'width:auto;min-height:44px;margin:0', onclick: () => { d.choixCat = true; drawCats(); } }, 'changer')));
      return;
    }
    catBox.replaceChildren(chips(catsNoms, d.cat, v => { appliquerCat(d, v); d.choixCat = false; drawCats(); drawSubs(); destBox.set(d.dests); urg.className = 'chip' + (d.urgent ? ' on-bad' : ''); }));
  };
  drawCats();
  const thumbs = h('div', { class: 'thumbs' });
  const drawThumbs = () => thumbs.replaceChildren(...d.photos.map((p, i) =>
    h('div', { class: 't' }, h('img', { src: p.thumb, alt: '' }), h('button', { onclick: () => { d.photos.splice(i, 1); drawThumbs(); } }, '×'))));
  drawThumbs();
  const retour = () => { sigDraft = null; if (from === 'ronde') go('ronde', {}, { noPush: true }); else { geoStopSiLibre(); go('home'); } };
  const back = () => { if (!d.photos.length && !d.desc || confirm('Abandonner ce signalement ?')) retour(); };

  const save = (encore) => safe(async () => {
    if (!d.site) throw new Error('Choisissez le parking.');
    if (!d.agent) throw new Error('Indiquez qui fait le signalement.');
    if (!d.cat) throw new Error('Choisissez une catégorie.');
    if (d.cat !== CAT_AUTRE && !d.sub) throw new Error('Choisissez ce qui ne va pas (« Quoi ? »).');
    if ((d.cat === CAT_AUTRE || d.sub === SUB_AUTRE) && !d.desc.trim()) throw new Error('Décrivez le problème.');
    if (!d.dests.length) throw new Error('Choisissez au moins un destinataire.');
    showBusy('Enregistrement…');
    const site = siteById(d.site);
    const id = uuid(), ref = `${site.code}-${randCode(5)}`;
    const veh = isVehicule(d.cat);
    await savePhotos(d.photos);
    await appendEvent('signalement', { rid: ronde ? ronde.rid : null, site: d.site, agent: d.agent },
      {
        id, ref, cat: d.cat, sub: d.sub && d.sub !== SUB_AUTRE ? d.sub : (d.sub === SUB_AUTRE ? 'Autre' : ''), desc: d.desc.trim(),
        plaque: veh ? d.plaque.trim().toUpperCase() : '', emplacement: veh ? d.emplacement.trim() : '', vehicule: veh ? d.vehicule.trim() : '',
        dests: [...d.dests], urgent: d.urgent, horsRonde: !ronde, thumb: d.photos[0] ? d.photos[0].thumb : null
      },
      { photoAges: d.photos.map(p => ({ id: p.id, ageMin: p.ageMin })) }, d.photos);
    if (ronde) { const r = await kvGet('ronde'); r.sigsCrees.push(id); await kvSet('ronde', r); }
    else { await kvSet('dernierAgent', { nom: d.agent, at: Date.now() }); state.dernierSite = d.site; }
    await fold();
    const urgent = d.urgent, catNom = d.cat;
    sigDraft = null;
    hideBusy();
    toast(`Signalement ${ref} enregistré.`);
    if (encore) go('signalement', { from, cat: catNom }, { noPush: true });
    else retour();
    if (urgent || !ronde) syncNow().then(refreshIfHome);   // urgent ou hors ronde : envoi immédiat
  });

  return page(ronde ? 'Nouveau signalement' : 'Signalement hors ronde', { back },
    !ronde ? [h('label', { class: 'f' }, 'Parking'), chips(cfg.sites.map(s => ({ val: s.id, lbl: s.nom })), d.site, v => d.site = v),
      h('label', { class: 'f' }, 'Signalé par'), chips(cfg.agents, d.agent, v => d.agent = v)] : null,
    h('label', { class: 'f' }, 'Catégorie'), catBox,
    subsBox,
    descLbl,
    h('textarea', { placeholder: 'Où exactement (niveau, place, porte…), ce qui ne va pas. Le micro du clavier permet de dicter.', oninput: e => d.desc = e.target.value, value: d.desc }),
    vehBox,
    h('label', { class: 'f' }, 'Photos'),
    thumbs,
    h('button', { class: 'sec', onclick: photoInput(async f => { if (d.photos.length >= 4) { toast('4 photos maximum.', true); return; } showBusy('Photo…'); d.photos.push(await processPhoto(f)); hideBusy(); drawThumbs(); }) }, 'Prendre une photo'),
    h('label', { class: 'f' }, 'À signaler à'), destBox,
    h('label', { class: 'f' }, 'Priorité'), h('div', { class: 'chips' }, urg),
    h('div', { style: 'height:130px' }),
    h('div', { class: 'sticky-bottom' }, h('div', null,
      h('button', { class: 'ok', onclick: save(false) }, 'Enregistrer'),
      h('button', { class: 'sec', onclick: save(true) }, 'Enregistrer et signaler autre chose ici'))));
};
function appliquerCat(d, nom) {
  d.cat = nom; d.sub = null;
  const c = catByNom(nom);
  if (!d.destsTouche) { d.dests.splice(0, d.dests.length, ...((c && c.dests) || []).filter(x => state.cfg.dests.includes(x))); }
}

/* ---------- Barrière ouverte / refermée ---------- */
let barrDraft = null;
VIEWS.barriere = async ({ site, bid, from }) => {
  const cfg = state.cfg;
  const toutes = (await kvGet('barrieres')) || [];
  const b = bid ? toutes.find(x => x.bid === bid) : null;
  if (bid && (!b || b.fermee)) { toast('Cette barrière est déjà signalée refermée.'); return agentHome(); }
  const cle = bid || 'new|' + (site || '');
  if (!barrDraft || barrDraft.cle !== cle) {
    barrDraft = { cle, site: site || (b && b.site) || (cfg.sites.length === 1 ? cfg.sites[0].id : null), barriere: null, motif: null, comment: '', agent: await agentPrefere(), maintenant: true, quand: localDT() };
  }
  Geo.start();
  const d = barrDraft;
  const quandIn = h('input', { type: 'datetime-local', value: d.quand, max: localDT(new Date(Date.now() + 60000)), style: d.maintenant ? 'display:none' : '', oninput: e => d.quand = e.target.value });
  const quandBox = chips([{ val: true, lbl: 'Maintenant' }, { val: false, lbl: 'Plus tôt (oubli)' }], d.maintenant, v => { d.maintenant = v; quandIn.style.display = v ? 'none' : ''; });
  const barrBox = h('div');
  const drawBarr = () => {
    const s = siteById(d.site);
    if (!s) { barrBox.replaceChildren(); return; }
    const ouvertes = toutes.filter(x => x.site === s.id && !x.fermee).map(x => x.barriere);
    barrBox.replaceChildren(h('label', { class: 'f' }, 'Quelle barrière ?'),
      chips(s.barrieres.map(x => ({ val: x, lbl: x + (ouvertes.includes(x) ? ' (déjà signalée ouverte)' : '') })), d.barriere, v => d.barriere = v));
  };
  drawBarr();
  const quitter = () => { barrDraft = null; if (from === 'ronde') go('ronde', {}, { noPush: true }); else { geoStopSiLibre(); go('home'); } };

  const save = safe(async () => {
    const quand = d.maintenant ? new Date() : new Date(d.quand);
    if (isNaN(quand)) throw new Error('Heure invalide.');
    if (quand.getTime() > Date.now() + 120000) throw new Error('L’heure ne peut pas être dans le futur.');
    if (quand.getTime() < Date.now() - 7 * 86400000) throw new Error('Au-delà de 7 jours, prévenez le DPMS.');
    if (!d.agent) throw new Error('Indiquez qui fait la déclaration.');
    let ev;
    if (b) {
      if (quand < new Date(b.ouverte.quand)) throw new Error(`La fermeture ne peut pas précéder l’ouverture (${fmtDT(b.ouverte.quand)}).`);
      ev = { bid: b.bid, action: 'fermee', barriere: b.barriere, quand: quand.toISOString(), comment: d.comment.trim(), retro: !d.maintenant };
      await appendEvent('barriere', { site: b.site, agent: d.agent }, ev);
    } else {
      if (!d.site) throw new Error('Choisissez le parking.');
      if (!d.barriere) throw new Error('Choisissez la barrière.');
      if (toutes.some(x => x.site === d.site && x.barriere === d.barriere && !x.fermee)) throw new Error('Cette barrière est déjà signalée ouverte.');
      if (!d.motif) throw new Error('Indiquez pourquoi elle est ouverte.');
      ev = { bid: uuid(), action: 'ouverte', barriere: d.barriere, quand: quand.toISOString(), motif: d.motif, comment: d.comment.trim(), retro: !d.maintenant };
      await appendEvent('barriere', { site: d.site, agent: d.agent }, ev);
    }
    await kvSet('dernierAgent', { nom: d.agent, at: Date.now() });
    await fold();
    toast(b ? 'Fermeture enregistrée.' : 'Ouverture enregistrée.');
    quitter();
    syncNow().then(refreshIfHome);
  });

  return page(b ? 'Barrière refermée' : 'Barrière ouverte', { back: quitter },
    b ? h('div', { class: 'card' }, barriereLigne(b, false)) : [
      h('label', { class: 'f' }, 'Parking'), chips(cfg.sites.map(s => ({ val: s.id, lbl: s.nom })), d.site, v => { d.site = v; d.barriere = null; drawBarr(); }),
      barrBox,
      h('label', { class: 'f' }, 'Pourquoi ?'), chips(cfg.motifs, d.motif, v => d.motif = v)],
    h('label', { class: 'f' }, b ? 'Refermée quand ?' : 'Ouverte depuis quand ?'), quandBox, quandIn,
    h('label', { class: 'f' }, 'Déclarée par'), chips(cfg.agents, d.agent, v => d.agent = v),
    h('label', { class: 'f' }, 'Commentaire (facultatif)'),
    h('input', { type: 'text', value: d.comment, oninput: e => d.comment = e.target.value }),
    h('button', { class: b ? 'ok big' : 'danger big', onclick: save }, b ? 'Enregistrer la fermeture' : 'Enregistrer l’ouverture'));
};

/* ---------- Liste des signalements (agents) ---------- */
VIEWS.agentSigs = async () => {
  const sigs = (await dbAll('sigs')).filter(s => s.statut === 'ouvert').sort((a, b) => (b.jour || '').localeCompare(a.jour || ''));
  const cards = [];
  for (const s of sigs) {
    const th = await thumbOf(s);
    cards.push(h('div', { class: 'card' },
      h('h3', null, sigTitre(s), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('p', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ref} — ${fmtJour(s.jour)}${s.agent ? ' — ' + s.agent : ''}${s.horsRonde ? ' — hors ronde' : ''}`),
      s.desc ? h('p', null, s.desc) : null,
      s.plaque || s.emplacement ? h('p', null, s.plaque ? [h('b', null, 'Plaque : '), s.plaque, ' '] : null, s.emplacement ? [h('b', null, 'Emplacement : '), s.emplacement] : null) : null,
      th ? h('div', { class: 'thumbs' }, h('img', { src: th, alt: '' })) : null,
      s.notesAgents ? h('div', { class: 'banner info' }, 'DPMS : ' + s.notesAgents) : null,
      pmBanner(s)));
  }
  return page('Signalements en cours', { back: true }, cards.length ? cards : h('p', { class: 'muted' }, 'Aucun signalement en cours.'));
};

/* ---------- Journal lisible (agents, PM) ---------- */
async function journalEquipe() {
  return [...(await dbAll('entries')).filter(e => e.kind === 'ev').map(e => ({ ord: e.n * 100000 + e.idx, ev: e.ev })),
  ...(await dbAll('events')).map(ev => ({ ord: ORD_LOCAL + ev.seq, ev }))].sort((a, b) => a.ord - b.ord).map(x => x.ev);
}
/* Extrait Excel (agents) : sans horodatage précis ni position */
async function agentExcel(jours) {
  const lim = localDate(new Date(Date.now() - jours * 86400000));
  const evs = await journalEquipe();
  const refs = new Map();
  evs.forEach(ev => { if (ev.t === 'signalement') refs.set(ev.data.id, ev.data); });
  const aoa = [['Date', 'Parking', 'Agent', 'Type', 'Réf.', 'Catégorie', 'Sous-catégorie', 'Description', 'Plaque', 'Emplacement', 'Destinataires', 'Urgent', 'Constat / valeur', 'Commentaire']];
  for (const ev of evs) {
    if (ev.jour < lim) continue;
    const base = [fmtJour(ev.jour), siteNom(ev.site), ev.agent || ''];
    if (ev.t === 'signalement') aoa.push([...base, ev.data.horsRonde ? 'Signalement hors ronde' : 'Signalement', ev.data.ref, ev.data.cat, ev.data.sub || '', ev.data.desc, ev.data.plaque || '', ev.data.emplacement || '', (ev.data.dests || []).join(', '), ev.data.urgent ? 'Oui' : '', '', '']);
    else if (ev.t === 'revue') { const s = refs.get(ev.data.sig) || {}; aoa.push([...base, 'Constat de suivi', ev.data.ref, s.cat || '', s.sub || '', s.desc || '', s.plaque || '', s.emplacement || '', '', '', VERDICTS[ev.data.verdict], ev.data.comment || '']); }
    else if (ev.t === 'checklist') { const an = Object.entries(ev.data.items || {}).filter(([, v]) => v !== 'RAS').map(([k, v]) => `${k} : ${v}`); aoa.push([...base, 'Contrôles', '', '', '', an.length ? an.join(' ; ') : 'Tout RAS', '', '', '', '', '', '']); }
    else if (ev.t === 'comptage') aoa.push([...base, 'Comptage des véhicules', '', '', '', '', '', '', '', '', ev.data.total, '']);
    else if (ev.t === 'barriere') aoa.push([...base, ev.data.action === 'ouverte' ? 'Barrière ouverte' : 'Barrière refermée', '', '', ev.data.barriere, ev.data.motif || '', '', '', '', '', fmtDT(ev.data.quand), ev.data.comment || '']);
    else if (ev.t === 'pm') { const s = refs.get(ev.data.sig) || {}; aoa.push([...base, 'Police municipale', ev.data.ref, s.cat || '', s.sub || '', s.desc || '', s.plaque || '', s.emplacement || '', '', '', PM_STATUTS[ev.data.statut], ev.data.obs || '']); }
  }
  if (aoa.length === 1) { toast('Aucune saisie sur la période.'); return; }
  await shareOrDownload(xlsxBlob([['Rondes', aoa, [10, 12, 12, 20, 11, 22, 30, 45, 12, 16, 24, 8, 16, 30]]]), `extrait_rondes_${fileStamp()}.xlsx`, 'Extrait des rondes');
}

/* ====================================================================== */
/* POLICE MUNICIPALE                                                      */
/* ====================================================================== */
function pourPM(s) { return isVehicule(s.cat) || (s.dests || []).some(d => PM_RE.test(d)); }
async function pmHome() {
  if (!state.cfg) {
    return page('Police municipale', null,
      h('div', { class: 'banner warn' }, 'Configuration pas encore reçue du relais. Connectez l’appareil au réseau.'),
      syncCard((await dbAll('events')).length, await kvGet('lastSync'), await kvGet('syncErr')));
  }
  const f = state.params.f || 'a_traiter';
  const portee = state.params.portee || 'vehicules';
  const all = (await dbAll('sigs')).filter(s => portee === 'vehicules' ? isVehicule(s.cat) : pourPM(s));
  const filtre = { a_traiter: s => pmEtat(s) === 'À traiter', non_traite: s => pmEtat(s) === 'Non traité', traite: s => pmEtat(s) === 'Traité', tous: () => true }[f];
  const list = all.filter(filtre).sort((a, b) => (a.jour || '').localeCompare(b.jour || ''));
  const n = k => all.filter({ a_traiter: s => pmEtat(s) === 'À traiter', non_traite: s => pmEtat(s) === 'Non traité', traite: s => pmEtat(s) === 'Traité', tous: () => true }[k]).length;
  const cards = [];
  for (const s of list) {
    const th = await thumbOf(s);
    cards.push(h('button', { class: 'list-item', onclick: () => go('pmSig', { id: s.id }) },
      h('div', { style: 'display:flex;gap:10px;align-items:flex-start' },
        th ? h('img', { src: th, alt: '', style: 'width:72px;height:72px;object-fit:cover;border-radius:8px;flex:none' }) : null,
        h('div', { style: 'flex:1;min-width:0' },
          h('div', { class: 'l1' }, s.plaque ? h('span', { class: 'plaque' }, s.plaque) : 'sans plaque', ' ', pmBadge(s)),
          h('div', { class: 'small' }, sigTitre(s)),
          h('div', { class: 'muted small' }, `${siteNom(s.site)}${s.emplacement ? ' — ' + s.emplacement : ''} — signalé le ${fmtJour(s.jour)} (${depuisJours(s.jour)})`)))));
  }
  return page('Police municipale', null,
    h('div', { class: 'chips' }, ...[['vehicules', 'Véhicules'], ['pm', 'Tout ce qui est adressé à la PM']].map(([k, l]) =>
      h('button', { class: 'chip' + (portee === k ? ' on' : ''), onclick: () => { state.params = { f, portee: k }; render(); } }, l))),
    h('div', { class: 'chips', style: 'margin-bottom:12px' }, ...[['a_traiter', 'À traiter'], ['non_traite', 'Non traités'], ['traite', 'Traités'], ['tous', 'Tous']].map(([k, l]) =>
      h('button', { class: 'chip' + (f === k ? ' on' : ''), onclick: () => { state.params = { f: k, portee }; render(); } }, `${l} (${n(k)})`))),
    cards.length ? cards : h('p', { class: 'muted' }, 'Aucun signalement.'),
    syncCard((await dbAll('events')).length, await kvGet('lastSync'), await kvGet('syncErr')),
    h('button', { class: 'sec', onclick: safe(pmExcel) }, 'Export Excel'),
    h('button', { class: 'link', onclick: () => go('reglages') }, 'Réglages'),
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
}
VIEWS.pmSig = async ({ id }) => {
  const s = await dbGet('sigs', id);
  if (!s) return page('Signalement', { back: true }, h('p', null, 'Introuvable.'));
  const obs = h('textarea', { placeholder: 'Observations : constat sur place, démarches (identification du titulaire, fourrière…), suite donnée' });
  const enregistrer = statut => safe(async () => {
    if (statut !== 'traite' && !obs.value.trim()) throw new Error('Saisissez une observation.');
    if (statut === 'traite' && !confirm('Marquer ce signalement comme traité ? Il sera clos pour les agents.')) return;
    await appendEvent('pm', { site: s.site, agent: 'Police municipale' }, { sig: s.id, ref: s.ref, statut, obs: obs.value.trim() });
    await fold();
    toast('Enregistré.');
    syncNow();
    go('pmSig', { id }, { noPush: true });
  });
  const suivi = (s.suivi || []).map(v => h('p', { class: 'small' }, `${fmtJour(v.jour)} — ${v.agent || ''} : ${VERDICTS[v.verdict]}${v.comment ? ' — ' + v.comment : ''}`));
  return page(s.plaque || s.ref, { back: () => go('home') },
    h('div', { class: 'card' },
      h('h3', null, sigTitre(s), ' ', pmBadge(s)),
      h('dl', { class: 'kv' },
        h('dt', null, 'Plaque'), h('dd', null, s.plaque ? h('span', { class: 'plaque' }, s.plaque) : '—'),
        h('dt', null, 'Emplacement'), h('dd', null, s.emplacement || '—'),
        h('dt', null, 'Véhicule'), h('dd', null, s.vehicule || '—'),
        h('dt', null, 'Parking'), h('dd', null, siteNom(s.site)),
        h('dt', null, 'Signalé'), h('dd', null, `${fmtJour(s.jour)} par ${s.agent || '?'} (${s.ref})`),
        h('dt', null, 'Statut'), h('dd', null, STATUTS[s.statut] + (s.closPar ? ` (par ${s.closPar})` : ''))),
      s.desc ? h('p', null, s.desc) : null,
      ...(await photoBlocks(s.photos))),
    suivi.length ? h('div', { class: 'card' }, h('h2', null, 'Constats des agents'), ...suivi) : null,
    s.pm && s.pm.hist.length ? h('div', { class: 'card' }, h('h2', null, 'Suivi police municipale'),
      ...s.pm.hist.map(x => h('p', { class: 'small' }, h('b', null, `${fmtJour(x.jour)} — ${PM_STATUTS[x.statut]}`), x.obs ? ' : ' + x.obs : ''))) : null,
    h('div', { class: 'card' }, h('h2', null, 'Observation'), obs,
      h('button', { class: 'sec', onclick: enregistrer('obs') }, 'Enregistrer l’observation'),
      h('div', { class: 'row' },
        h('button', { class: 'ok', onclick: enregistrer('traite') }, 'Traité'),
        h('button', { class: 'danger', onclick: enregistrer('non_traite') }, 'Non traité')),
      h('p', { class: 'muted small' }, '« Traité » clôt le signalement. « Non traité » le laisse ouvert, avec votre observation visible des agents.')));
};
async function pmExcel() {
  const sigs = (await dbAll('sigs')).filter(pourPM).sort((a, b) => (a.jour || '').localeCompare(b.jour || ''));
  const aoa = [['Réf.', 'Parking', 'Catégorie', 'Sous-catégorie', 'Plaque', 'Emplacement', 'Véhicule', 'Description', 'Signalé le', 'Par', 'Statut', 'Suivi PM', 'Dernière observation PM']];
  sigs.forEach(s => { const l = s.pm && s.pm.hist.length ? s.pm.hist[s.pm.hist.length - 1] : null; aoa.push([s.ref, siteNom(s.site), s.cat, s.sub || '', s.plaque || '', s.emplacement || '', s.vehicule || '', s.desc || '', fmtJour(s.jour), s.agent || '', STATUTS[s.statut], pmEtat(s), l ? `${fmtJour(l.jour)} : ${l.obs}` : '']); });
  await shareOrDownload(xlsxBlob([['Police municipale', aoa, [11, 12, 16, 30, 12, 16, 20, 40, 11, 12, 8, 12, 40]]]), `signalements_PM_${fileStamp()}.xlsx`, 'Signalements PM');
}

/* ---------- Réglages (commun) ---------- */
VIEWS.reglages = async () => {
  let persisted = null;
  try { persisted = navigator.storage && navigator.storage.persisted ? await navigator.storage.persisted() : null; } catch (e) { }
  let est = null; try { est = navigator.storage && navigator.storage.estimate ? await navigator.storage.estimate() : null; } catch (e) { }
  const dev = await kvGet('dev');
  const enAttente = estEquipe() ? (await dbAll('events')).length : 0;
  return page('Réglages', { back: true },
    h('div', { class: 'card' }, h('dl', { class: 'kv' },
      h('dt', null, 'Mode'), h('dd', null, { agent: 'Téléphone agents', pm: 'Police municipale', superviseur: 'Superviseur' }[state.mode]),
      dev ? [h('dt', null, 'Identifiant'), h('dd', null, dev)] : null,
      state.conn ? [h('dt', null, 'Relais'), h('dd', { class: 'small', style: 'word-break:break-all' }, state.conn.url)] : null,
      h('dt', null, 'Stockage protégé'), h('dd', null, persisted === true ? 'Oui' : persisted === false ? 'Non (installez l’application sur l’écran d’accueil)' : 'Inconnu'),
      est ? [h('dt', null, 'Espace utilisé'), h('dd', null, (est.usage / 1048576).toFixed(1) + ' Mo')] : null,
      h('dt', null, 'Version'), h('dd', null, APP_VERSION))),
    persisted === false ? h('button', { class: 'sec', onclick: safe(async () => { const ok = await requestPersist(); toast(ok ? 'Stockage protégé.' : 'Refusé par le navigateur.', !ok); render(); }) }, 'Demander la protection du stockage') : null,
    estEquipe() ? h('button', { class: 'sec', onclick: () => go('scan') }, 'Scanner un nouveau QR de configuration') : null,
    h('hr'),
    h('div', { class: 'card' }, h('h2', null, 'Réinitialiser cet appareil'),
      h('p', { class: 'muted small' }, estEquipe()
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
        // L'appareil est enregistré avant l'envoi du coffre : si l'envoi échoue, il est retenté à la synchronisation.
        const vault = JSON.stringify(await encryptWithPass(p1.value, { v: 2, privJwk: kp.privJwk, ...conn }));
        state.mode = 'superviseur'; state.conn = conn; state.priv = kp.priv;
        await kvSet('privKey', kp.priv); await kvSet('conn', conn); await kvSet('vaultPending', vault); await kvSet('mode', 'superviseur'); await kvDel('migre');
        const cfg = defaultConfig();
        state.cfg = cfg; await kvSet('cfg', cfg);
        await queueDec('d', { type: 'cfg', ts: new Date().toISOString(), cfg });
        const r = await syncNow();
        await requestPersist();
        hideBusy();
        if (r && r.err) toast('Relais activé, mais l’envoi a échoué : ' + r.err + '. Nouvel essai automatique.', true); else toast('Relais activé.');
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
  const vault = await kvGet('vaultPending');
  if (vault) { await rel('putVault', { data: vault }); await kvDel('vaultPending'); }
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
  const barrOuv = await barrieresOuvertes();
  const vehPM = sigs.filter(s => isVehicule(s.cat) && pmEtat(s) === 'À traiter').length;
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
    barrOuv.length ? h('div', { class: 'card', style: 'border-color:var(--bad)' }, h('h2', null, 'Barrières ouvertes en ce moment'), ...barrOuv.map(b => barriereLigne(b, false))) : null,
    h('div', { class: 'card' }, h('h2', null, 'Signalements'),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'ouvert' }) }, h('div', { class: 'stat' }, h('span', null, 'En cours', urg ? h('span', { class: 'badge b-bad' }, urg + ' urgent(s)') : null), h('b', null, ouverts.length))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'clos' }) }, h('div', { class: 'stat' }, h('span', null, 'Clos'), h('b', null, sigs.length - ouverts.length))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'vehicules' }) }, h('div', { class: 'stat' }, h('span', null, 'Véhicules en attente de la PM'), h('b', null, vehPM)))),
    h('div', { class: 'card' }, h('h2', null, 'Dernières rondes'),
      rondes.length ? rondes.map(rondeLigne) : h('p', { class: 'muted' }, 'Aucune ronde reçue.'),
      h('div', { class: 'row' },
        rondes.length ? h('button', { class: 'sec', onclick: () => go('supRondes') }, 'Toutes les rondes') : null,
        h('button', { class: 'sec', onclick: () => go('supBarrieres') }, 'Barrières'))),
    h('div', { class: 'card' }, h('h2', null, 'Exports'),
      h('button', { class: 'sec', onclick: safe(supExcel) }, 'Excel complet (heures et positions)'),
      h('button', { class: 'sec', onclick: safe(supArchive) }, 'Archive complète avec photos (.zip)')),
    h('div', { class: 'card' }, h('h2', null, 'Paramétrage'),
      h('button', { class: 'sec', onclick: () => go('supQR') }, 'QR codes (parkings, agents, police municipale)'),
      h('button', { class: 'sec', onclick: () => go('supConfig') }, 'Listes, catégories et coordonnées'),
      h('button', { class: 'sec', onclick: () => go('supRelais') }, 'Relais et appareils')),
    Object.keys(devs).length ? h('div', { class: 'card' }, h('h2', null, 'Appareils de l’équipe'),
      ...Object.entries(devs).map(([d, v]) => h('div', { class: 'stat' }, h('span', { style: 'white-space:nowrap;margin-right:8px' }, d + (v.role === 'pm' ? ' (PM)' : '')), h('span', { class: 'muted small', style: 'text-align:right' }, `${v.seq} saisies — dernier envoi ${fmtQuand(v.at)}`)))) : null,
    h('button', { class: 'link', onclick: () => go('reglages') }, 'Réglages'),
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
}
function rondeLigne(r) {
  const dur = r.debut && r.fin ? Math.round((new Date(r.fin) - new Date(r.debut)) / 60000) : null;
  const anos = r.checklist ? Object.entries(r.checklist).filter(([, v]) => v !== 'RAS').map(([k, v]) => v === 'Anomalie' ? k : `${k} (${v.toLowerCase()})`) : [];
  return h('div', { class: 'stat', style: 'display:block' },
    h('div', null, h('b', null, `${siteNom(r.site)} — ${r.agent}`), r.qr === false ? h('span', { class: 'badge b-warn' }, 'sans QR') : null, !r.fin ? h('span', { class: 'badge b-warn' }, 'fin non reçue') : null),
    h('div', { class: 'muted small' }, `${fmtDT(r.debut)}${r.fin ? ' → ' + fmtHeure(r.fin) : ''}${dur != null ? ` (${dur} min)` : ''} — ${r.compte != null ? r.compte + ' véhicule(s), ' : ''}${r.nbSig} signalement(s), ${r.nbRevue} constat(s)`),
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
function retroInfo(rec) {
  if (!rec || !rec.ts) return null;
  const ecart = (new Date(rec.ts) - new Date(rec.quand)) / 60000;
  return ecart > RETRO_MIN ? `déclarée a posteriori (saisie ${fmtQuand(rec.ts)})` : null;
}
VIEWS.supBarrieres = async () => {
  const list = ((await kvGet('barrieres')) || []).sort((a, b) => b.ouverte.quand.localeCompare(a.ouverte.quand));
  const total = {};
  list.forEach(b => { const fin = b.fermee ? new Date(b.fermee.quand) : new Date(); const k = `${siteNom(b.site)} — ${b.barriere}`; total[k] = (total[k] || 0) + (fin - new Date(b.ouverte.quand)); });
  return page('Barrières ouvertes', { back: true },
    Object.keys(total).length ? h('div', { class: 'card' }, h('h2', null, 'Durée cumulée d’ouverture'),
      ...Object.entries(total).map(([k, v]) => h('div', { class: 'stat' }, h('span', null, k), h('b', null, fmtDuree(v))))) : null,
    list.length ? list.map(b => {
      const ro = retroInfo(b.ouverte), rf = retroInfo(b.fermee);
      return h('div', { class: 'card', style: b.fermee ? '' : 'border-color:var(--bad)' },
        h('div', null, h('b', null, `${siteNom(b.site)} — barrière ${b.barriere}`), b.fermee ? h('span', { class: 'badge b-ok' }, 'refermée') : h('span', { class: 'badge b-bad' }, 'ouverte')),
        h('p', { class: 'small' }, `Ouverte ${fmtQuand(b.ouverte.quand)} par ${b.ouverte.agent || '?'} — ${b.ouverte.motif}${b.ouverte.comment ? ' — ' + b.ouverte.comment : ''}`, ro ? h('span', { class: 'badge b-warn' }, ro) : null),
        b.fermee ? h('p', { class: 'small' }, `Refermée ${fmtQuand(b.fermee.quand)} par ${b.fermee.agent || '?'}${b.fermee.comment ? ' — ' + b.fermee.comment : ''}`, rf ? h('span', { class: 'badge b-warn' }, rf) : null) : null,
        h('p', { class: 'small' }, h('b', null, 'Durée : '), fmtDuree((b.fermee ? new Date(b.fermee.quand) : new Date()) - new Date(b.ouverte.quand)), b.fermee ? '' : ' (en cours)'));
    }) : h('p', { class: 'muted' }, 'Aucune barrière signalée ouverte.'));
};

/* ---------- Relais et appareils ---------- */
VIEWS.supRelais = async () => page('Relais et appareils', { back: true },
  h('div', { class: 'card' }, h('h2', null, 'Adresse du relais'), h('p', { class: 'small', style: 'word-break:break-all' }, state.conn.url)),
  h('div', { class: 'card' }, h('h2', null, 'Ajouter un appareil'),
    h('p', { class: 'small' }, h('b', null, 'Superviseur (téléphone ou ordinateur) : '), 'ouvrir l’application → « Ajouter cet appareil comme superviseur » → adresse ci-dessus + phrase de passe.'),
    h('p', { class: 'small' }, h('b', null, 'Téléphone agents ou police municipale : '), 'ouvrir l’application → scanner le QR correspondant (menu « QR codes »).'),
    h('p', { class: 'small muted' }, 'Un appareil ajouté ou réinstallé retrouve toutes les données conservées sur le relais.')),
  h('div', { class: 'card' }, h('h2', null, 'Repartir de zéro'),
    h('p', { class: 'small' }, 'Dans l’éditeur Apps Script du relais : choisir la fonction « reinitialiserRelais » dans la barre d’outils, cliquer sur « Exécuter ». Tout le contenu du relais est mis à la corbeille du compte Google. Réinitialiser ensuite chaque appareil (Réglages), puis activer un relais neuf.')));

/* ---------- Liste et fiche des signalements ---------- */
VIEWS.supSigs = async ({ f = 'ouvert', site = '' }) => {
  const all = await dbAll('sigs');
  const test = { ouvert: s => s.statut === 'ouvert', clos: s => s.statut === 'clos', vehicules: s => isVehicule(s.cat), tous: () => true }[f];
  const list = all.filter(s => test(s) && (!site || s.site === site))
    .sort((a, b) => (b.urgent - a.urgent) || (b.ts || b.jour || '').localeCompare(a.ts || a.jour || ''));
  const fchips = [['ouvert', 'En cours'], ['clos', 'Clos'], ['vehicules', 'Véhicules'], ['tous', 'Tous']].map(([k, l]) =>
    h('button', { class: 'chip' + (f === k ? ' on' : ''), onclick: () => go('supSigs', { f: k, site }, { noPush: true }) }, l));
  const schips = [['', 'Tous parkings'], ...state.cfg.sites.map(s => [s.id, s.nom])].map(([k, l]) =>
    h('button', { class: 'chip' + (site === k ? ' on' : ''), onclick: () => go('supSigs', { f, site: k }, { noPush: true }) }, l));
  return page('Signalements', { back: true },
    h('div', { class: 'chips' }, ...fchips), h('div', { class: 'chips', style: 'margin-bottom:12px' }, ...schips),
    list.length ? list.map(s => h('button', { class: 'list-item', onclick: () => go('supSig', { id: s.id }) },
      h('div', { class: 'l1' }, `${s.ref} · ${sigTitre(s)}`, statutBadge(s.statut), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null, isVehicule(s.cat) ? pmBadge(s) : null),
      h('div', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ts ? fmtDT(s.ts) : fmtJour(s.jour)} — ${s.agent || ''}${s.horsRonde ? ' (hors ronde)' : ''} — ${(s.dests || []).join(', ')}`),
      s.plaque ? h('div', { class: 'small' }, h('span', { class: 'plaque' }, s.plaque), s.emplacement ? ' — ' + s.emplacement : '') : null,
      s.desc ? h('div', { class: 'small' }, s.desc.length > 120 ? s.desc.slice(0, 120) + '…' : s.desc) : null))
      : h('p', { class: 'muted' }, 'Aucun signalement.'));
};
function geoLink(geo) {
  if (!geo || geo.err || geo.lat == null) return h('span', { class: 'muted' }, geo && geo.err ? 'position ' + geo.err : '—');
  return h('a', { href: `https://www.openstreetmap.org/?mlat=${geo.lat}&mlon=${geo.lon}#map=19/${geo.lat}/${geo.lon}`, target: '_blank', rel: 'noopener' },
    `${geo.lat}, ${geo.lon} (± ${geo.acc} m)`);
}
/* Photos : téléchargées à la demande depuis le relais, vérifiées, conservées sur l'appareil */
async function fetchMedia(n) {
  let r;
  try { r = await rel('media', { n }, 180000); }
  catch (e) { if (/interdit/.test(e.message)) throw new Error('le relais doit être mis à jour (version 3) pour afficher les photos sur cet appareil'); throw e; }
  const z = fflate.unzipSync(await teamDec(b64u.dec(r.data)));
  const out = [];
  for (const [name, bytes] of Object.entries(z)) out.push({ id: name.replace(/\.jpg$/, ''), blob: new Blob([bytes], { type: 'image/jpeg' }), sha: await sha256hex(bytes) });
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
  syncNow();
}
VIEWS.supSig = async ({ id, edit }) => {
  const s = await dbGet('sigs', id);
  if (!s) return page('Signalement', { back: true }, h('p', null, 'Introuvable.'));
  const back = () => go('supSigs', { f: s.statut });
  const done = (msg) => { toast(msg); go('supSig', { id }, { noPush: true, keepScroll: true }); };
  if (edit) {
    const catSel = h('select', null, ...[...state.cfg.cats.map(c => c.nom), CAT_AUTRE].map(c => h('option', { value: c, selected: c === s.cat ? true : null }, c)));
    const sub = h('input', { type: 'text', value: s.sub || '' });
    const desc = h('textarea', { value: s.desc || '' });
    const plaque = h('input', { type: 'text', value: s.plaque || '' });
    const empl = h('input', { type: 'text', value: s.emplacement || '' });
    const dests = [...(s.dests || [])];
    let urgent = !!s.urgent;
    const urg = h('button', { class: 'chip' + (urgent ? ' on-bad' : ''), onclick: e => { urgent = !urgent; e.target.className = 'chip' + (urgent ? ' on-bad' : ''); } }, 'Urgent');
    const msg = h('input', { type: 'text', value: s.notesAgents || '', placeholder: 'Ex. : intervention Ateliers prévue mardi' });
    return page(`Modifier ${s.ref}`, { back: () => go('supSig', { id }, { noPush: true }) },
      h('label', { class: 'f' }, 'Catégorie'), catSel,
      h('label', { class: 'f' }, 'Sous-catégorie'), sub,
      h('label', { class: 'f' }, 'Description'), desc,
      h('label', { class: 'f' }, 'Plaque'), plaque,
      h('label', { class: 'f' }, 'Emplacement'), empl,
      h('label', { class: 'f' }, 'À signaler à'), chips(state.cfg.dests, dests, () => { }, { multi: true }),
      h('label', { class: 'f' }, 'Priorité'), h('div', { class: 'chips' }, urg),
      h('label', { class: 'f' }, 'Message affiché aux agents'), msg,
      h('button', {
        class: 'ok', onclick: safe(async () => {
          await supDecision(s, 'modification', { cat: catSel.value, sub: sub.value.trim(), desc: desc.value.trim(), plaque: plaque.value.trim().toUpperCase(), emplacement: empl.value.trim(), dests: [...dests], urgent, notesAgents: msg.value.trim() });
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
      h('h3', null, sigTitre(s), statutBadge(s.statut), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('dl', { class: 'kv' },
        h('dt', null, 'Parking'), h('dd', null, siteNom(s.site)),
        h('dt', null, 'Signalé'), h('dd', null, s.ts ? fmtDT(s.ts) : fmtJour(s.jour), s.agent ? ' par ' + s.agent : '', s.horsRonde ? ' (hors ronde)' : ''),
        h('dt', null, 'Position'), h('dd', null, geoLink(s.geo)),
        h('dt', null, 'À signaler à'), h('dd', null, (s.dests || []).join(', ')),
        s.plaque ? [h('dt', null, 'Plaque'), h('dd', null, h('span', { class: 'plaque' }, s.plaque))] : null,
        s.emplacement ? [h('dt', null, 'Emplacement'), h('dd', null, s.emplacement)] : null,
        s.vehicule ? [h('dt', null, 'Véhicule'), h('dd', null, s.vehicule)] : null,
        s.statut === 'clos' ? [h('dt', null, 'Clos'), h('dd', null, `${s.closLe && s.closLe.length > 10 ? fmtDT(s.closLe) : fmtJour(s.closLe)}${s.closPar ? ' par ' + s.closPar : ''}`)] : null,
        s.notesAgents ? [h('dt', null, 'Message aux agents'), h('dd', null, s.notesAgents)] : null),
      s.desc ? h('p', null, s.desc) : null,
      ...(await photoBlocks(s.photos))),
    s.pm && s.pm.hist.length ? h('div', { class: 'card' }, h('h2', null, 'Police municipale'),
      ...s.pm.hist.map(x => h('p', { class: 'small' }, h('b', null, `${x.ts ? fmtDT(x.ts) : fmtJour(x.jour)} — ${PM_STATUTS[x.statut]}`), x.obs ? ' : ' + x.obs : ''))) : null,
    suivi.length ? h('h2', { style: 'font-size:16px;color:var(--navy)' }, 'Constats des agents') : null, ...suivi,
    h('div', { class: 'card' }, h('h2', null, 'Intervenir (facultatif)'),
      h('p', { class: 'muted small' }, 'Toute intervention est transmise automatiquement aux agents, à la PM et aux autres superviseurs.'),
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
  const barr = ((await kvGet('barrieres')) || []).sort((a, b) => a.ouverte.quand.localeCompare(b.ouverte.quand));
  const evs = (await dbAll('entries')).filter(e => e.kind === 'ev').sort((a, b) => a.k.localeCompare(b.k));
  const g = (geo, k) => geo && !geo.err && geo[k] != null ? geo[k] : '';
  const A1 = [['Réf.', 'Parking', 'Catégorie', 'Sous-catégorie', 'Description', 'Plaque', 'Emplacement', 'Véhicule', 'Destinataires', 'Urgent', 'Aggravé', 'Hors ronde', 'Statut', 'Signalé le', 'Agent', 'Latitude', 'Longitude', 'Précision (m)', 'Photos', 'Dernier constat', 'Date dernier constat', 'Clos le', 'Clos par', 'Suivi PM', 'Observations PM', 'Message aux agents', 'Notes DPMS']];
  sigs.forEach(s => {
    const last = (s.suivi || []).slice(-1)[0];
    A1.push([s.ref, siteNom(s.site), s.cat, s.sub || '', s.desc, s.plaque || '', s.emplacement || '', s.vehicule || '', (s.dests || []).join(', '), s.urgent ? 'Oui' : '', s.aggrave ? 'Oui' : '', s.horsRonde ? 'Oui' : '', STATUTS[s.statut] || s.statut, s.ts ? fmtDT(s.ts) : fmtJour(s.jour), s.agent || '',
      g(s.geo, 'lat'), g(s.geo, 'lon'), g(s.geo, 'acc'), (s.photos || []).length, last ? VERDICTS[last.verdict] : '', last ? (last.ts ? fmtDT(last.ts) : fmtJour(last.jour)) : '',
      s.closLe ? (s.closLe.length > 10 ? fmtDT(s.closLe) : fmtJour(s.closLe)) : '', s.closPar || '',
      isVehicule(s.cat) || s.pm ? pmEtat(s) : '', s.pm ? s.pm.hist.map(x => `${fmtJour(x.jour)} ${PM_STATUTS[x.statut]}${x.obs ? ' : ' + x.obs : ''}`).join(' | ') : '', s.notesAgents || '', s.notes || '']);
  });
  const A2 = [['Parking', 'Agent', 'Début', 'Fin', 'Durée (min)', 'Véhicules comptés', 'Démarrage par QR', 'Contrôles non RAS', 'Signalements', 'Constats', 'Lat. départ', 'Lon. départ', 'Précision (m)', 'Points d’attention', 'Téléphone']];
  rondes.forEach(r => {
    const dur = r.debut && r.fin ? Math.round((new Date(r.fin) - new Date(r.debut)) / 60000) : '';
    const anos = r.checklist ? Object.entries(r.checklist).filter(([, v]) => v !== 'RAS').map(([k, v]) => `${k} (${v})`).join(', ') : '';
    A2.push([siteNom(r.site), r.agent, fmtDT(r.debut), fmtDT(r.fin), dur, r.compte != null ? r.compte : '', r.qr === false ? 'Non' : 'Oui', anos, r.nbSig, r.nbRevue, g(r.geoDebut, 'lat'), g(r.geoDebut, 'lon'), g(r.geoDebut, 'acc'), [...new Set(r.alertes || [])].join(' ; '), r.dev]);
  });
  const A3 = [['Parking', 'Barrière', 'Ouverte le', 'Par', 'Motif', 'Commentaire', 'Saisie de l’ouverture', 'Refermée le', 'Par', 'Commentaire', 'Saisie de la fermeture', 'Durée (min)', 'En cours']];
  barr.forEach(b => {
    const fin = b.fermee ? new Date(b.fermee.quand) : new Date();
    A3.push([siteNom(b.site), b.barriere, fmtDT(b.ouverte.quand), b.ouverte.agent || '', b.ouverte.motif || '', b.ouverte.comment || '', fmtDT(b.ouverte.ts), b.fermee ? fmtDT(b.fermee.quand) : '', b.fermee ? b.fermee.agent || '' : '', b.fermee ? b.fermee.comment || '' : '', b.fermee ? fmtDT(b.fermee.ts) : '', Math.round((fin - new Date(b.ouverte.quand)) / 60000), b.fermee ? '' : 'Oui']);
  });
  const TYPES = { ronde_debut: 'Début de ronde', ronde_fin: 'Fin de ronde', checklist: 'Contrôles', comptage: 'Comptage', signalement: 'Signalement', revue: 'Constat de suivi', barriere: 'Barrière', pm: 'Police municipale' };
  const A4 = [['Appareil', 'N°', 'Type', 'Jour déclaré', 'Horodatage appareil', 'Reçu par le relais', 'Parking', 'Agent', 'Détail', 'Latitude', 'Longitude', 'Précision (m)', 'Âge position (s)', 'Empreinte']];
  evs.forEach(e => {
    const ev = e.ev; let det = '';
    if (ev.t === 'signalement') det = `${ev.data.ref} ${ev.data.cat}${ev.data.sub ? ' / ' + ev.data.sub : ''}${ev.data.desc ? ' — ' + ev.data.desc : ''}${ev.data.horsRonde ? ' (hors ronde)' : ''}`;
    if (ev.t === 'revue') det = `${ev.data.ref} : ${VERDICTS[ev.data.verdict]}${ev.data.comment ? ' — ' + ev.data.comment : ''}`;
    if (ev.t === 'checklist') det = Object.entries(ev.data.items || {}).map(([k, v]) => `${k} : ${v}`).join(' ; ');
    if (ev.t === 'comptage') det = `${ev.data.total} véhicule(s)`;
    if (ev.t === 'ronde_debut') det = ev.data.qr ? 'QR scanné' : 'sans QR';
    if (ev.t === 'ronde_fin') det = `${ev.data.nbSig} signalement(s)`;
    if (ev.t === 'barriere') det = `${ev.data.barriere} ${ev.data.action} — déclaré ${fmtDT(ev.data.quand)}${ev.data.motif ? ' — ' + ev.data.motif : ''}`;
    if (ev.t === 'pm') det = `${ev.data.ref} : ${PM_STATUTS[ev.data.statut]}${ev.data.obs ? ' — ' + ev.data.obs : ''}`;
    const geo = e.dec && e.dec.geo;
    A4.push([ev.dev, ev.seq, TYPES[ev.t] || ev.t, fmtJour(ev.jour), e.dec ? fmtDT(e.dec.ts) : 'illisible', fmtDT(e.at), siteNom(ev.site), ev.agent || '', det, g(geo, 'lat'), g(geo, 'lon'), g(geo, 'acc'), g(geo, 'age'), e.hashOk ? 'OK' : 'INVALIDE']);
  });
  hideBusy();
  await shareOrDownload(xlsxBlob([
    ['Signalements', A1, [11, 12, 22, 30, 45, 12, 16, 20, 24, 7, 7, 8, 8, 16, 12, 11, 11, 9, 7, 16, 16, 16, 14, 11, 40, 30, 40]],
    ['Rondes', A2, [12, 12, 16, 16, 10, 10, 10, 40, 12, 10, 11, 11, 10, 40, 10]],
    ['Barrières', A3, [12, 12, 16, 12, 22, 25, 16, 16, 12, 25, 16, 10, 8]],
    ['Journal', A4, [10, 6, 16, 11, 16, 16, 12, 12, 55, 11, 11, 10, 10, 10]]
  ]), `rondes_parkings_complet_${fileStamp()}.xlsx`, 'Rondes parkings');
}
/* ---------- Archive complète ---------- */
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
  const dump = { format: 'rondes-cachan-archive', v: 3, app: APP_VERSION, cree: new Date().toISOString(), cfg: state.cfg, sigs: await dbAll('sigs'), rondes: await dbAll('rondes'), barrieres: (await kvGet('barrieres')) || [], journal: (await dbAll('entries')).map(e => ({ ...e })) };
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
  const qrBloc = (titre, texte, role) => {
    const show = h('div', { class: 'qr' });
    return h('div', { class: 'card' }, h('h2', null, titre), h('p', { class: 'muted small' }, texte),
      h('button', { onclick: () => { show.innerHTML = qrSvg(appBaseURL() + '#join=' + packJoin(state.conn, role), 'L'); } }, 'Afficher le QR'), show,
      h('p', { class: 'muted small' }, 'Ce QR donne accès aux signalements de l’équipe : ne le diffusez pas.'));
  };
  return page('QR codes', { back: true },
    first ? h('div', { class: 'banner ok' }, 'Relais activé. Étapes suivantes : saisir les prénoms des agents (Listes, catégories et coordonnées), imprimer les QR des parkings, configurer les téléphones.') : null,
    qrBloc('Téléphones des agents', 'À scanner une fois avec l’appareil photo de chaque téléphone agents. Les modifications de listes leur parviennent ensuite seules.', 'agent'),
    qrBloc('Police municipale', 'À scanner avec l’appareil du chef de service PM : accès aux signalements de véhicules (plaques, emplacements, photos), observations, traité / non traité.', 'pm'),
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
  cfg = normalizeCfg(cfg);
  state.cfg = cfg; await kvSet('cfg', cfg);
  await queueDec('d', { type: 'cfg', ts: new Date().toISOString(), cfg });
  syncNow();
}

/* ---------- Listes, catégories et coordonnées ---------- */
VIEWS.supConfig = async () => {
  const c = JSON.parse(JSON.stringify(state.cfg));
  const ta = (lines, hmin = 140) => h('textarea', { value: lines.join('\n'), style: `min-height:${hmin}px` });
  const tAgents = ta(c.agents), tDests = ta(c.dests), tMotifs = ta(c.motifs);
  const tTaxo = h('textarea', { value: taxoToText(c.cats), style: 'min-height:420px;font-size:15px' });
  const sitesInputs = c.sites.map(s => ({ nom: h('input', { type: 'text', value: s.nom }), barr: h('input', { type: 'text', value: s.barrieres.join(', ') }) }));
  const tel = h('input', { type: 'tel', value: c.urgenceTel || '', placeholder: '01 …' });
  const mail = h('input', { type: 'email', value: c.urgenceMail || '', placeholder: 'prenom.nom@ville-cachan.fr' });
  const lines = t => t.value.split('\n').map(x => x.trim()).filter(Boolean);
  return page('Listes et catégories', { back: true },
    h('div', { class: 'banner info' }, 'Les modifications sont transmises automatiquement à tous les appareils.'),
    h('div', { class: 'card' }, h('h2', null, 'Parkings et barrières'),
      ...sitesInputs.map(si => h('div', { style: 'margin-bottom:10px' }, si.nom, h('label', { class: 'f small' }, 'Barrières (séparées par des virgules)'), si.barr))),
    h('label', { class: 'f' }, 'Agents (un par ligne)'), tAgents,
    h('label', { class: 'f' }, 'Destinataires (un par ligne)'), tDests,
    h('label', { class: 'f' }, 'Catégories et sous-catégories'),
    h('p', { class: 'muted small' }, 'Une catégorie par ligne, suivie si besoin de « | » et des destinataires proposés par défaut. Puis ses sous-catégories, une par ligne commençant par « - ». Un « ! » en fin de ligne marque une sous-catégorie urgente par défaut. Chaque catégorie est une ligne de la check-list de ronde. « Autre (préciser) » est ajouté automatiquement.'),
    tTaxo,
    h('label', { class: 'f' }, 'Motifs d’ouverture de barrière (un par ligne)'), tMotifs,
    h('label', { class: 'f' }, 'Téléphone d’urgence affiché aux agents'), tel,
    h('label', { class: 'f' }, 'Courriel d’urgence'), mail,
    h('button', {
      class: 'ok', onclick: safe(async () => {
        const cats = parseTaxo(tTaxo.value);
        if (!cats.length) throw new Error('Aucune catégorie.');
        const vides = cats.filter(x => !x.subs.length).map(x => x.nom);
        if (vides.length) throw new Error('Catégorie sans sous-catégorie : ' + vides.join(', '));
        const dests = lines(tDests);
        const inconnus = [...new Set(cats.flatMap(x => x.dests).filter(x => !dests.includes(x)))];
        if (inconnus.length) throw new Error('Destinataire inconnu dans les catégories : ' + inconnus.join(', '));
        if (!lines(tAgents).length || !dests.length) throw new Error('Listes vides.');
        sitesInputs.forEach((si, k) => {
          if (si.nom.value.trim()) c.sites[k].nom = si.nom.value.trim();
          const b = si.barr.value.split(',').map(x => x.trim()).filter(Boolean);
          if (b.length) c.sites[k].barrieres = b;
        });
        Object.assign(c, { agents: lines(tAgents), dests, cats, motifs: lines(tMotifs), urgenceTel: tel.value.trim(), urgenceMail: mail.value.trim(), cfgId: randCode(8) });
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
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).then(r => r.update()).catch(e => console.warn('SW', e));
  state.mode = (await kvGet('mode')) || null;
  state.conn = (await kvGet('conn')) || null;
  state.cfg = normalizeCfg((await kvGet('cfg')) || null);
  state.priv = state.mode === 'superviseur' ? await kvGet('privKey') : null;
  if (state.mode && !state.conn) { state.mode = null; }
  if (state.mode === 'agent' && (await kvGet('ronde'))) Geo.start();
  const hash = location.hash;
  if (hash && /^#(join|site|cfg)=/.test(hash)) {
    history.replaceState(null, '', location.pathname + location.search);
    await handleLink(hash, false);
    if (state.view === 'home') render();
  } else render();
  if (state.mode) syncNow().then(refreshIfHome);
}
window.addEventListener('hashchange', () => {
  const hash = location.hash;
  if (hash && /^#(join|site|cfg)=/.test(hash)) { history.replaceState(null, '', location.pathname + location.search); handleLink(hash, false); }
});
window.addEventListener('online', () => { if (state.mode) syncNow().then(refreshIfHome); });
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && state.mode) syncNow().then(refreshIfHome); });
setInterval(() => { if (document.visibilityState === 'visible' && state.mode) syncNow().then(refreshIfHome); }, SYNC_PERIODE_MS);
boot();
