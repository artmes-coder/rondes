/* Rondes parkings — Ville de Cachan, DPMS — version 2.1
 * Application web hors ligne (PWA).
 * - L'outil appartient aux agents : rondes, comptage des véhicules, signalements (catégories et
 *   sous-catégories, plusieurs par catégorie, en ronde ou hors ronde), barrières laissées ouvertes.
 * - Police municipale : suivi des signalements de véhicules (observations, traité / non traité).
 * - Superviseurs : consultation, interventions ponctuelles, exports.
 * - Le relais (Google Apps Script) conserve un journal chiffré commun. Clé d'équipe : agents, PM,
 *   superviseurs. Clé superviseur : horodatage exact, positions, notes internes. Le relais ne lit rien.
 */
'use strict';

const APP_VERSION = '2.12.0';
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
/* Lecture-modification-écriture d'une clé en UNE transaction : deux traitements simultanés
   (saisie de l'utilisateur et relève en cours) ne peuvent plus s'écraser mutuellement. */
async function kvUpdate(k, fn) {
  const d = await openDB(); const t = d.transaction('kv', 'readwrite'); const o = t.objectStore('kv');
  let res;
  o.get(k).onsuccess = e => { res = fn(e.target.result); o.put(res, k); };
  await txDone(t); return res;
}
const idFile = () => uuid();
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
// « relais occupé » : un autre appareil utilise le relais ; rien n'a été traité, on peut réessayer sans risque
async function relayCall(url, body, timeoutMs = 60000) {
  for (let essai = 0; ; essai++) {
    try { return await relayCall1(url, body, timeoutMs); }
    catch (e) {
      if (!/occup/i.test(e.code || '') || essai >= 3) throw e;
      await new Promise(r => setTimeout(r, 3000 * (essai + 1) + Math.random() * 2000));
    }
  }
}
async function relayCall1(url, body, timeoutMs) {
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
        : j.err === 'relais non initialisé' ? 'Le relais a été remis à zéro. Cet appareil doit être reconfiguré : Réglages → Réinitialiser cet appareil, puis configuration (superviseur : « activer un relais neuf » ; téléphones : nouveau QR).'
          : j.err === 'accès refusé' ? 'Le relais a été réactivé avec de nouveaux codes. Cet appareil doit être reconfiguré : Réglages → Réinitialiser cet appareil, puis nouveau QR ou « Ajouter cet appareil comme superviseur ».'
            : /occup/i.test(j.err || '') ? 'Le relais est occupé par un autre appareil (envoi ou récupération en cours). Réessayez dans une minute.'
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
/*   - sous-catégorie ! | destinataires propres (facultatifs)              */
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
- Réserve de tickets basse ou vide
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
- Tags ou graffitis | Ateliers, DPMS, Police municipale
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
- Véhicule entré sans ticket
Bornes de recharge électrique | Ateliers
- Borne hors service
- Câble de recharge arraché ou dégradé
- Extincteur à proximité absent
- Chaleur, fumée ou odeur de brûlé !
Accès et circulation | Ateliers
- Rampe ou voie de circulation encombrée
- Verglas ou neige sur la rampe ou les voies !
- Chaussée dégradée (trou, plaque soulevée)
- Portail en panne (ne s’ouvre pas ou ne se ferme pas)
- Bouton-poussoir désactivé ou hors service
- Portail endommagé
Entretien et maintenance | DPMS
- Vérification des extincteurs dépassée
- Entretien des portails, barrières ou portes automatiques dépassé
- Visite de l’ascensoriste dépassée
- Vérification du désenfumage dépassée
- Vérification de l’éclairage de sécurité dépassée
- Maintenance des bornes de recharge dépassée
- Date d’entretien absente ou illisible
Sûreté et présences | Police municipale
- Personnes installées (squat, regroupement)
- Personne endormie ou en difficulté !
- Comportement suspect
- Trace d’effraction | Police municipale, DPMS, Ateliers
- Caméra de vidéoprotection dégradée ou masquée | Police municipale, DPMS
- Vandalisme en cours (appeler le 17) ! | Police municipale, DPMS`;
const CONTROLES_DEFAUT = `[Quotidien]
- Rampes et voies de circulation : dégagées ; en hiver, ni verglas ni neige | Accès et circulation | Rampe ou voie de circulation encombrée ; Verglas ou neige sur la rampe ou les voies ! ; Chaussée dégradée (trou, plaque soulevée)
- ?Portails : fonctionnent, boutons-poussoirs sur « on » | Accès et circulation | Portail en panne (ne s’ouvre pas ou ne se ferme pas) ; Bouton-poussoir désactivé ou hors service ; Portail endommagé
- Barrières et bornes : fonctionnent, en bon état, bornes à ticket en service, réserve de tickets suffisante | Barrières, bornes et caisses | Barrière en panne (bloquée ouverte) ; Barrière en panne (ne s’ouvre pas) ; Lisse de barrière cassée ; Borne d’entrée hors service ; Borne de sortie hors service ; Réserve de tickets basse ou vide
- ?Caisses automatiques : en service, pas de trace d’effraction | Barrières, bornes et caisses | Caisse automatique hors service ; Caisse forcée ou dégradée !
- ?Ascenseurs : en service | Accès piétons et ascenseurs | Ascenseur en panne ; Alarme ou téléphone de cabine d’ascenseur hors service
- Éclairage : lumières et interrupteurs fonctionnent | Éclairage et électricité | Luminaire(s) éteint(s) ; Zone entièrement dans le noir ; Interrupteur hors service ; Luminaire cassé ou pendant
- Blocs de sécurité : allumés | Issues de secours et évacuation | Bloc de sécurité éteint ; Bloc de sécurité arraché ou cassé
- Eau : grilles de drainage dégagées, pas d’eau stagnante, pas de trace d’infiltration (taches, suintements) aux plafonds, parois et sols | Bâtiment et infiltrations | Grille de drainage ou avaloir bouché ; Eau stagnante ; Infiltration (taches, suintements) ; Inondation !
- Extincteurs : tous en place (aucun manquant), non utilisés, accès libre | Sécurité incendie | Extincteur absent ou déplacé ; Extincteur utilisé, vide ou scellé rompu ; Accès à un extincteur encombré
- Issues de secours : balisées, portes fermées et non verrouillées | Issues de secours et évacuation | Issue de secours encombrée ! ; Issue de secours verrouillée ou bloquée ! ; Balisage de l’issue absent ou éteint ; Barre antipanique ou porte d’issue défectueuse
- Portes coupe-feu : fermées, non calées ouvertes | Sécurité incendie | Porte coupe-feu calée ouverte ; Porte coupe-feu qui ne se ferme pas
- Plans d’évacuation et consignes : en place, lisibles, non dégradés | Issues de secours et évacuation | Plan d’évacuation absent ou arraché ; Plan ou consignes dégradés ou illisibles
- Propreté : déchets, dépôts, urine, seringues, tags | Propreté et hygiène
- Sûreté : présences, squat, effraction, caméra masquée ou arrachée | Sûreté et présences
- Véhicules : épave, véhicule gênant, sur place PMR, ventouse, entré sans ticket | Véhicules | Épave ou véhicule hors d’usage ; Véhicule mal stationné ou gênant ; Véhicule sur place PMR sans carte ; Véhicule ventouse (stationnement prolongé) ; Véhicule entré sans ticket ; Véhicule fracturé ou vitre brisée ; Personne dans un véhicule
[Mensuel]
- Extincteurs (un par un) : goupille et scellé présents, support fixé, panneau au-dessus | Sécurité incendie | Goupille ou scellé absent ; Support d’extincteur arraché ; Panneau de l’extincteur absent
- Bac d’absorbant : rempli, pelle présente | Sécurité incendie | Bac d’absorbant vide ou absent ; Pelle absente
- ?Déclencheurs manuels d’alarme : intacts, non masqués, accessibles (sans les actionner) | Sécurité incendie | Déclencheur cassé, masqué ou enclenché ; Déclencheur inaccessible
- Blocs d’éclairage de sécurité (un par un) : voyant de charge allumé, bloc non arraché | Issues de secours et évacuation | Voyant de charge éteint ; Bloc de sécurité arraché ou cassé
- Ferme-portes des portes coupe-feu : la porte lâchée se referme complètement seule | Sécurité incendie | Porte coupe-feu qui ne se referme pas seule ; Ferme-porte cassé ou décroché
- ?Commandes de désenfumage : accessibles, boîtier fermé et intact, signalées (sans les actionner) | Sécurité incendie | Boîtier de commande ouvert ou dégradé ; Commande inaccessible ; Signalétique de la commande absente
- Bouches de ventilation : dégagées, rien stationné ni stocké devant | Ventilation et qualité de l’air | Bouche de ventilation obstruée ; Bouche de ventilation dégradée
- Signalétique d’évacuation : panneaux en place et visibles | Issues de secours et évacuation | Panneau d’évacuation absent ou masqué ; Panneau d’évacuation éteint ou cassé
- ?Interphones et appel d’urgence : essai d’appel depuis chaque borne et depuis la cabine d’ascenseur | Barrières, bornes et caisses | Interphone de borne hors service ; Téléphone de cabine d’ascenseur hors service ; Appel sans réponse du centre
- Signalisation routière : hauteur maximale à l’entrée, gabarit, marquage au sol, places PMR, numérotation | Signalisation et marquage
- ?Bornes de recharge : en service, câbles intacts, extincteur à proximité | Bornes de recharge électrique
- Issues côté extérieur : rien stationné ni entreposé devant les sorties et les accès pompiers | Issues de secours et évacuation | Véhicule stationné devant une sortie ; Objets entreposés devant une sortie ; Accès pompiers encombré !
[Trimestriel]
- Entretien des extincteurs : sur chaque extincteur, étiquette de vérification de moins d’un an | Entretien et maintenance | Vérification des extincteurs dépassée ; Étiquette absente ou illisible
- ?Entretien des portails, barrières et portes automatiques : dernière visite d’entretien datée de moins de 6 mois | Entretien et maintenance | Entretien des portails, barrières ou portes automatiques dépassé ; Date d’entretien absente ou illisible
- ?Entretien des ascenseurs : passage de l’ascensoriste noté, dernière visite de moins de 6 semaines | Entretien et maintenance | Visite de l’ascensoriste dépassée ; Date d’entretien absente ou illisible
- ?Entretien du désenfumage : dernière vérification datée, échéance indiquée non dépassée | Entretien et maintenance | Vérification du désenfumage dépassée ; Date d’entretien absente ou illisible
- Entretien de l’éclairage de sécurité : dernière vérification datée, échéance indiquée non dépassée | Entretien et maintenance | Vérification de l’éclairage de sécurité dépassée ; Date d’entretien absente ou illisible
- ?Entretien des bornes de recharge : dernière maintenance datée, échéance indiquée non dépassée | Entretien et maintenance | Maintenance des bornes de recharge dépassée ; Date d’entretien absente ou illisible
- ?Colonnes sèches : prises accessibles, capots et bouchons présents, signalétique en place | Sécurité incendie | Prise de colonne sèche inaccessible ; Capot ou bouchon manquant ; Signalétique absente
- ?Locaux techniques : portes fermées à clé, signalées, aucun stockage | Bâtiment et infiltrations | Local technique ouvert ou non verrouillé ; Stockage dans un local technique ; Signalétique absente
- Clés, badges et télécommandes d’exploitation : inventaire | Autre | Clé, badge ou télécommande manquant`;
const CONTROLES_DEFAUT_V16 = `[Quotidien]
- Rampes et voies de circulation : dégagées ; en hiver, ni verglas ni neige | Accès et circulation
- ?Portails : fonctionnent, boutons-poussoirs sur « on » | Accès et circulation
- Barrières et bornes : fonctionnent, en bon état, bornes à ticket en service, réserve de tickets suffisante | Barrières, bornes et caisses
- ?Caisses automatiques : en service, pas de trace d’effraction | Barrières, bornes et caisses
- ?Ascenseurs : en service | Accès piétons et ascenseurs
- Éclairage : lumières et interrupteurs fonctionnent | Éclairage et électricité
- Blocs de sécurité : allumés | Issues de secours et évacuation
- Eau : grilles de drainage dégagées, pas d’eau stagnante, pas de trace d’infiltration (taches, suintements) aux plafonds, parois et sols | Bâtiment et infiltrations
- Extincteurs : tous en place (aucun manquant), non utilisés, accès libre | Sécurité incendie
- Issues de secours : balisées, portes fermées et non verrouillées | Issues de secours et évacuation
- Portes coupe-feu : fermées, non calées ouvertes | Sécurité incendie
- Plans d’évacuation et consignes : en place, lisibles, non dégradés | Issues de secours et évacuation
- Propreté : déchets, dépôts, urine, seringues, tags | Propreté et hygiène
- Sûreté : présences, squat, effraction, caméra masquée ou arrachée | Sûreté et présences
- Véhicules : épave, véhicule gênant, sur place PMR, ventouse, entré sans ticket | Véhicules
[Mensuel]
- Extincteurs (un par un) : goupille et scellé présents, support fixé, panneau au-dessus | Sécurité incendie
- Bac d’absorbant : rempli, pelle présente | Sécurité incendie
- ?Déclencheurs manuels d’alarme : intacts, non masqués, accessibles (sans les actionner) | Sécurité incendie
- Blocs d’éclairage de sécurité (un par un) : voyant de charge allumé, bloc non arraché | Issues de secours et évacuation
- Ferme-portes des portes coupe-feu : la porte lâchée se referme complètement seule | Sécurité incendie
- ?Commandes de désenfumage : accessibles, boîtier fermé et intact, signalées (sans les actionner) | Sécurité incendie
- Bouches de ventilation : dégagées, rien stationné ni stocké devant | Ventilation et qualité de l’air
- Signalétique d’évacuation : panneaux en place et visibles | Issues de secours et évacuation
- ?Interphones et appel d’urgence : essai d’appel depuis chaque borne et depuis la cabine d’ascenseur | Barrières, bornes et caisses
- Signalisation routière : hauteur maximale à l’entrée, gabarit, marquage au sol, places PMR, numérotation | Signalisation et marquage
- ?Bornes de recharge : en service, câbles intacts, extincteur à proximité | Bornes de recharge électrique
- Issues côté extérieur : rien stationné ni entreposé devant les sorties et les accès pompiers | Issues de secours et évacuation
[Trimestriel]
- Entretien des extincteurs : sur chaque extincteur, étiquette de vérification de moins d’un an | Entretien et maintenance
- ?Entretien des portails, barrières et portes automatiques : dernière visite d’entretien datée de moins de 6 mois | Entretien et maintenance
- ?Entretien des ascenseurs : passage de l’ascensoriste noté, dernière visite de moins de 6 semaines | Entretien et maintenance
- ?Entretien du désenfumage : dernière vérification datée, échéance indiquée non dépassée | Entretien et maintenance
- Entretien de l’éclairage de sécurité : dernière vérification datée, échéance indiquée non dépassée | Entretien et maintenance
- ?Entretien des bornes de recharge : dernière maintenance datée, échéance indiquée non dépassée | Entretien et maintenance
- ?Colonnes sèches : prises accessibles, capots et bouchons présents, signalétique en place | Sécurité incendie
- ?Locaux techniques : portes fermées à clé, signalées, aucun stockage | Bâtiment et infiltrations
- Clés, badges et télécommandes d’exploitation : inventaire | Autre`;
const CONTROLES_DEFAUT_V15 = `[Quotidien]
- Barrières et bornes : entrée et sortie fonctionnent, lisses intactes, bornes de ticket en service | Barrières, bornes et caisses
- ?Caisses automatiques : en service, écran allumé, pas de trace d’effraction | Barrières, bornes et caisses
- Issues de secours : dégagées, portes fermées mais pas verrouillées | Issues de secours et évacuation
- Portes coupe-feu : fermées, pas calées ouvertes | Sécurité incendie
- Éclairage : pas de zone éteinte, blocs de sécurité allumés | Éclairage et électricité
- Extincteurs : à leur place, non utilisés, accès libre | Sécurité incendie
- Propreté : déchets, dépôts, urine, seringues, tags, taches d’huile, matières combustibles stockées | Propreté et hygiène
- Eau : fuite, flaque importante, avaloir bouché | Bâtiment et infiltrations
- ?Ascenseurs : en service | Accès piétons et ascenseurs
- Sûreté : présences, squat, trace d’effraction, caméra masquée ou arrachée | Sûreté et présences
- Véhicules : épave, véhicule gênant ou sur place PMR, véhicule suspect ou présent depuis longtemps | Véhicules
[Mensuel]
- Extincteurs (un par un) : étiquette de vérification à jour, goupille et scellé, support fixé, panneau au-dessus | Sécurité incendie
- Bac d’absorbant : rempli, pelle présente | Sécurité incendie
- Déclencheurs manuels d’alarme : intacts, non masqués, accessibles (sans les actionner) | Sécurité incendie
- Blocs d’éclairage de sécurité (un par un) : voyant de charge allumé, bloc non arraché | Issues de secours et évacuation
- Ferme-portes des portes coupe-feu : la porte lâchée se referme complètement seule | Sécurité incendie
- ?Commandes de désenfumage : accessibles, boîtier fermé et intact, signalées | Sécurité incendie
- Bouches de ventilation : dégagées, rien stationné ni stocké devant | Ventilation et qualité de l’air
- Signalétique d’évacuation : panneaux en place et visibles | Issues de secours et évacuation
- Plans et consignes : affichés, lisibles, non arrachés | Sécurité incendie
- ?Interphones et appel d’urgence : essai d’appel depuis chaque borne et depuis la cabine d’ascenseur | Barrières, bornes et caisses
- Signalisation routière : hauteur maximale à l’entrée, gabarit, marquage au sol, places PMR, numérotation | Signalisation et marquage
- ?Bornes de recharge : en service, câbles intacts, extincteur à proximité | Bornes de recharge électrique
- Issues côté extérieur : rien stationné ni entreposé devant les sorties et les accès pompiers | Issues de secours et évacuation
[Trimestriel]
- ?Colonnes sèches : prises accessibles, capots et bouchons présents, signalétique en place | Sécurité incendie
- Inventaire des extincteurs : nombre et emplacements conformes au plan, au moins un par niveau à chaque issue | Sécurité incendie
- Plans d’évacuation et d’intervention : conformes à l’état des lieux, affichés aux bons endroits | Sécurité incendie
- ?Locaux techniques : portes fermées à clé, signalées, aucun stockage | Bâtiment et infiltrations
- Dates d’entretien affichées : ascenseurs, portes automatiques, extincteurs ; relever et signaler les dates dépassées | Autre
- Clés, badges et télécommandes d’exploitation : inventaire | Autre`;
const NIVEAUX = { quotidien: { titre: 'Quotidien', lbl: 'Contrôle quotidien', jours: 1 }, mensuel: { titre: 'Mensuel', lbl: 'Contrôle mensuel', jours: 30 }, trimestriel: { titre: 'Trimestriel', lbl: 'Contrôle trimestriel', jours: 90 } };
function slug(t) { return t.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40); }
function parseControles(text) {
  const out = []; let niv = null;
  for (const raw of text.split('\n')) {
    const l = raw.trim(); if (!l) continue;
    const m = /^\[(.+)\]$/.exec(l);
    if (m) { const k = slug(m[1]); niv = NIVEAUX[k] ? k : null; continue; }
    if (!niv || !/^[-•*]/.test(l)) continue;
    let t = l.replace(/^[-•*]\s*/, ''), opt = false;
    if (t.startsWith('?')) { opt = true; t = t.slice(1).trim(); }
    const [corps, cat, quoi] = t.split('|').map(x => x.trim());
    const i = corps.indexOf(' : ');
    const lbl = (i >= 0 ? corps.slice(0, i) : corps).trim(), detail = i >= 0 ? corps.slice(i + 3).trim() : '';
    const subs = (quoi || '').split(';').map(x => x.trim()).filter(Boolean).map(x => x.endsWith('!') ? { lbl: x.slice(0, -1).trim(), urgent: true } : { lbl: x, urgent: false });
    if (lbl) out.push({ id: niv[0] + '-' + slug(lbl), niveau: niv, lbl, detail, cat: cat || CAT_AUTRE, opt, ...(subs.length ? { subs } : {}) });
  }
  return out;
}
function controlesToText(list) {
  return Object.keys(NIVEAUX).map(n => [`[${NIVEAUX[n].titre}]`, ...list.filter(c => c.niveau === n).map(c => `- ${c.opt ? '?' : ''}${c.lbl}${c.detail ? ' : ' + c.detail : ''}${c.cat ? ' | ' + c.cat : ''}${(c.subs || []).length ? ' | ' + c.subs.map(x => x.lbl + (x.urgent ? ' !' : '')).join(' ; ') : ''}`)].join('\n')).join('\n');
}
const MOTIFS_DEFAUT = ['Panne de la barrière', 'Panne de borne ou de caisse', 'Intervention technique', 'Forte affluence', 'Consigne de la hiérarchie', 'Autre'];
const CAT_AUTRE = 'Autre';
const JOURS = ['L', 'M', 'M', 'J', 'V', 'S', 'D'], JOURS_LONGS = ['lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi', 'dimanche'];
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
      const [corps, ds] = line.replace(/^[-•*]\s*/, '').split('|');
      let lbl = corps.trim(), urgent = false;
      if (lbl.endsWith('!')) { urgent = true; lbl = lbl.slice(0, -1).trim(); }
      if (lbl) cats[cats.length - 1].subs.push({ lbl, urgent, dests: (ds || '').split(',').map(x => x.trim()).filter(Boolean) });
    } else {
      const [nom, d] = line.split('|');
      cats.push({ nom: nom.trim(), dests: (d || '').split(',').map(x => x.trim()).filter(Boolean), subs: [] });
    }
  }
  return cats.filter(c => c.nom && c.nom !== CAT_AUTRE);
}
function taxoToText(cats) {
  return cats.map(c => [c.nom + (c.dests.length ? ' | ' + c.dests.join(', ') : ''), ...c.subs.map(s => '- ' + s.lbl + (s.urgent ? ' !' : '') + ((s.dests || []).length ? ' | ' + s.dests.join(', ') : ''))].join('\n')).join('\n');
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
  if (!Array.isArray(n.objetsContact) || !n.objetsContact.length) n.objetsContact = OBJETS_DEFAUT.slice();
  n.gps = n.gps === true;
  if (!Array.isArray(n.controles) || !n.controles.length || [CONTROLES_DEFAUT_V15, CONTROLES_DEFAUT_V16].some(t => JSON.stringify(n.controles) === JSON.stringify(parseControles(t)))) n.controles = parseControles(CONTROLES_DEFAUT);
  // Catégories utilisées par les points de contrôle mais absentes de la liste : reprises de la liste par défaut
  const taxo = parseTaxo(TAXO_DEFAUT);
  n.cats.forEach(c => (c.subs || []).forEach(sc => {
    if (Array.isArray(sc.dests)) return;
    const dc = taxo.find(x => x.nom === c.nom), ds = dc && dc.subs.find(x => x.lbl === sc.lbl);
    sc.dests = ds ? ds.dests.slice() : [];
  }));
  [...new Set(n.controles.map(c => c.cat).filter(Boolean))].forEach(nom => {
    if (nom !== CAT_AUTRE && !n.cats.some(c => c.nom === nom)) { const d = taxo.find(c => c.nom === nom); if (d) n.cats.push(d); }
  });
  n.contacts = n.contacts && typeof n.contacts === 'object' ? n.contacts : {};
  n.rondeMatin = { actif: true, heure: '20:00', jours: [1, 2, 3, 4, 5, 6, 7], ...(n.rondeMatin || {}) };
  if (!n.rondeMatin.v2) { if (n.rondeMatin.heure === '12:00') n.rondeMatin.heure = '20:00'; n.rondeMatin.v2 = true; }   // ancienne valeur par défaut : fin de service des agents à 20 h
  n.synthese = { actif: true, jour: 1, heure: '08:00', ...(n.synthese || {}) };
  n.planning = { ...PLANNING_DEFAUT, ...(n.planning || {}) };
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
    if (!(state.cfg && state.cfg.gps) || this.watchId !== null || !('geolocation' in navigator)) return;
    this.err = null;
    this.watchId = navigator.geolocation.watchPosition(
      p => { this.last = { lat: +p.coords.latitude.toFixed(6), lon: +p.coords.longitude.toFixed(6), acc: Math.round(p.coords.accuracy), ts: p.timestamp }; this.err = null; },
      e => { this.err = ({ 1: 'refusée', 2: 'indisponible', 3: 'délai dépassé' })[e.code] || 'erreur'; },
      { enableHighAccuracy: true, maximumAge: 30000, timeout: 30000 });
  },
  stop() { if (this.watchId !== null) navigator.geolocation.clearWatch(this.watchId); this.watchId = null; this.last = null; },
  async waitFirst(ms) { const t0 = Date.now(); while (this.watchId !== null && !this.last && !this.err && Date.now() - t0 < ms) await new Promise(r => setTimeout(r, 200)); },
  snap() {
    if (!(state.cfg && state.cfg.gps)) return { off: true };
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
  if (view !== 'reglages' && view !== 'scan') state.reglagesOk = false;
  state.view = view; state.params = params;
  if (!opts.noPush && view !== 'home') history.pushState({ v: view }, '');
  render();
  if (!opts.keepScroll) window.scrollTo(0, 0);
}
let _renderSeq = 0;
async function render() {
  if (_majPrete && state.view === 'home' && appliquerMiseAJour()) return;
  const root = document.getElementById('app');
  const fn = VIEWS[state.view] || VIEWS.home;
  const my = ++_renderSeq;
  document.body.classList.toggle('large', state.mode === 'superviseur');
  document.body.classList.toggle('agent', state.mode === 'agent');
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
const TYPES_DEC = new Set(['cfg', 'action', 'msg', 'demande_close', 'vapid', 'contact', 'contact_suivi', 'ronde_close', 'registre', 'ctl_reactiver', 'sig', 'suppr', 'annul_suppr', 'planning']);
const TYPES_EV = new Set(['signalement', 'revue', 'action', 'pm', 'barriere', 'ronde_debut', 'ronde_fin', 'ronde_annulee', 'ronde_agent', 'comptage', 'controle', 'msg', 'checklist', 'contact', 'contact_suivi']);
async function fold() {
  const sup = state.mode === 'superviseur';
  const { entries, decQueue } = await (async () => {
    const db = await openDB(); const t = db.transaction(['entries', 'kv']);
    const [a, b] = await Promise.all([reqP(t.objectStore('entries').getAll()), reqP(t.objectStore('kv').get('decQueue'))]);
    return { entries: a, decQueue: b || [] };
  })();
  const items = [];
  for (const e of entries) items.push({ ord: e.n * 100000 + (e.idx || 0), e });
  if (!sup) for (const ev of await dbAll('events')) items.push({ ord: ORD_LOCAL + ev.seq, e: { kind: 'ev', ev, local: true } });
  if (sup) decQueue.forEach((q, i) => items.push({ ord: 2 * ORD_LOCAL + i, e: { kind: q.kind === 'n' ? 'note' : 'dec', dec: q.d, local: true } }));
  items.sort((a, b) => a.ord - b.ord);

  // Suppressions décidées par un superviseur (annulables depuis la corbeille)
  const suppr = new Map(), annul = new Set();
  for (const { e } of items) if (e.kind === 'dec') {
    if (e.dec.type === 'suppr') suppr.set(e.dec.id, e.dec);
    if (e.dec.type === 'annul_suppr') annul.add(e.dec.cible);
  }
  const del = { sigs: new Set(), rids: new Set() };
  for (const [id, d] of suppr) if (!annul.has(id)) { (d.sigs || []).forEach(x => del.sigs.add(x)); (d.rids || []).forEach(x => del.rids.add(x)); }
  const annulees = new Set();
  for (const { e } of items) if (e.kind === 'ev' && e.ev.t === 'ronde_annulee' && e.ev.rid) annulees.add(e.ev.rid);
  const evSupprime = ev => (ev.t === 'signalement' && del.sigs.has(ev.data.id)) || (['revue', 'action', 'pm', 'msg'].includes(ev.t) && del.sigs.has(ev.data.sig)) || (ev.rid && del.rids.has(ev.rid) && ev.t !== 'signalement' && ev.t !== 'revue');

  const sigs = new Map(), rondes = new Map(), barr = new Map(), devs = {}, chains = {}, ctl = {}, ctlHist = [], regs = [], mcs = new Map();
  const rondesCloses = new Map();   // rid → décision du DPMS « ronde clôturée »
  const planning = { jours: {}, notes: {}, pubs: [] }, planMotifs = {};
  let cfg = null, vapid = null;
  const alertes = [];
  // Précédence causale : une saisie d'agent ne change le statut (clos / ouvert) que si l'appareil
  // connaissait le dernier changement de statut (même appareil, ou changement antérieur à son curseur « vu »).
  // Sinon (décision du DPMS ou d'un autre appareil qu'il ignorait) : le constat est conservé, le statut ne bouge pas.
  const nDe = e => e.local ? Infinity : (e.n || 0);
  const decVues = new Set();
  let inconnu = false;
  for (const { e } of items) if ((e.kind === 'dec' && e.dec && e.dec.type && !TYPES_DEC.has(e.dec.type)) || (e.kind === 'ev' && e.ev && !TYPES_EV.has(e.ev.t))) inconnu = true;
  if (inconnu !== !!(await kvGet('majRequise'))) { await kvSet('majRequise', inconnu); if (inconnu) chercherMiseAJour(true); }
  const connu = (x, ev) => ev.vu == null || x.statutDev === ev.dev || (x.statutN || 0) <= ev.vu;
  const marquer = (x, e, qui) => { x.statutN = nDe(e); x.statutDev = qui; };
  for (const { e } of items) {
    if (e.kind === 'ev') {
      const ev = e.ev, dec = e.dec || null, ts = dec && dec.ts;
      e.supprime = evSupprime(ev);
      if (sup) {
        (chains[ev.dev] = chains[ev.dev] || []).push(e);
        const dv = devs[ev.dev] = devs[ev.dev] || { seq: 0, at: null, role: ev.role || 'agent' };
        if (ev.seq > dv.seq) { dv.seq = ev.seq; dv.at = e.at; }
      }
      if (e.supprime) continue;
      if (ev.t === 'signalement') {
        if (!sigs.has(ev.data.id)) sigs.set(ev.data.id, {
          id: ev.data.id, ref: ev.data.ref, site: ev.site, cat: ev.data.cat, sub: ev.data.sub || '', desc: ev.data.desc,
          plaque: ev.data.plaque || '', emplacement: ev.data.emplacement || '', vehicule: ev.data.vehicule || '',
          dests: ev.data.dests, urgent: ev.data.urgent, horsRonde: !!ev.data.horsRonde, statut: 'ouvert', jour: ev.jour, agent: ev.agent,
          thumb: e.local && typeof ev.data.thumb === 'string' ? ev.data.thumb : null,
          photos: (ev.photos || []).map(p => ({ id: p.id, sha: p.sha, n: e.n })), ts, geo: dec && dec.geo, dev: ev.dev, seq: ev.seq,
          suivi: [], journalDPMS: [], notes: '', notesAgents: '', pm: null, actions: [], messages: [], controle: ev.data.controle || null, rid: ev.rid || null, mc: ev.data.mc || null,
          statutN: nDe(e), statutDev: ev.dev
        });
      } else if (ev.t === 'contact') { mcAjout(mcs, ev.data, ev.agent); const c = mcs.get(ev.data.id); if (c && c.statutN == null) marquer(c, e, ev.dev);
      } else if (ev.t === 'contact_suivi') { mcSuivi(mcs, ev.data, ev.agent, { connu: x => connu(x, ev), marquer: x => marquer(x, e, ev.dev) });
      } else if (ev.t === 'ronde_agent' && ev.rid) {   // rondier corrigé : saisies de la ronde réattribuées
        const { ancien, nouveau } = ev.data;
        for (const s of sigs.values()) {
          if (s.rid === ev.rid && s.agent === ancien) s.agent = nouveau;
          (s.suivi || []).forEach(v => { if (v.rid === ev.rid && v.agent === ancien) v.agent = nouveau; });
        }
        ctlHist.forEach(x => { if (x.rid === ev.rid && x.agent === ancien) x.agent = nouveau; });
        Object.values(ctl).forEach(c => Object.keys(NIVEAUX).forEach(n => { if (c[n] && c[n].rid === ev.rid && c[n].agent === ancien) c[n].agent = nouveau; }));
      } else if (ev.t === 'msg') {
        const s = sigs.get(ev.data.sig);
        if (s) s.messages.push({ id: `${ev.dev}:${ev.seq}`, role: 'agent', de: ev.agent || 'Agent', quand: ev.data.quand || ts || ev.jour, texte: ev.data.texte || '', fait: !!ev.data.fait });
      } else if (ev.t === 'revue') {
        const s = sigs.get(ev.data.sig);
        if (s) {
          const change = (s.statut === 'clos') !== (ev.data.verdict === 'resolu');
          const ok = !change || connu(s, ev);   // pas de changement de statut : toujours appliqué
          s.suivi.push({ rid: ev.rid || null, jour: ev.jour, ts, agent: ev.agent, verdict: ev.data.verdict, comment: ev.data.comment, photos: (ev.photos || []).map(p => ({ id: p.id, sha: p.sha, n: e.n })), geo: dec && dec.geo, sansEffet: ok ? undefined : true });
          if (change && ok) { applyVerdict(s, ev.data.verdict, ts || ev.jour, ev.agent); marquer(s, e, ev.dev); }
          else if (!change && ev.data.verdict === 'aggrave') s.aggrave = true;
          else if (sup && (s.statut === 'clos') !== (ev.data.verdict === 'resolu')) alertes.push({ key: `tardif:${ev.dev}:${ev.seq}`, at: e.at, msg: `${s.ref} (${siteNom(s.site)}) — constat « ${VERDICTS[ev.data.verdict]} » de ${ev.agent} fait sans connaître la dernière décision (${s.statut === 'clos' ? 'clôture' : 'réouverture'}) : statut maintenu, à vérifier.` });
        } else if (sup) alertes.push({ key: `orph:${ev.dev}:${ev.seq}`, at: e.at, msg: `${siteNom(ev.site)}, ${ev.agent}, ${fmtJour(ev.jour)} — constat sur un signalement inconnu (${ev.data.ref})` });
      } else if (ev.t === 'action') {
        const s = sigs.get(ev.data.sig);
        if (s) s.actions.push({ type: ev.data.type, dest: ev.data.dest, texte: ev.data.texte, prevu: ev.data.prevu, quand: ev.data.quand || null, jour: ev.jour, par: ev.agent });
      } else if (ev.t === 'pm') {
        const s = sigs.get(ev.data.sig);
        if (s) {
          s.pm = s.pm || { statut: null, hist: [] };
          s.pm.hist.push({ jour: ev.jour, ts, statut: ev.data.statut, obs: ev.data.obs || '', agent: ev.agent });
          if (ev.data.statut !== 'obs') s.pm.statut = ev.data.statut;
          if (connu(s, ev)) {
            if (ev.data.statut === 'traite') { s.statut = 'clos'; s.closLe = ts || ev.jour; s.closPar = 'Police municipale'; marquer(s, e, ev.dev); }
            else if (ev.data.statut === 'non_traite' && s.closPar === 'Police municipale') { s.statut = 'ouvert'; s.closLe = null; s.closPar = null; marquer(s, e, ev.dev); }
          }
        }
      } else if (ev.t === 'controle') {
        const c = ctl[ev.site] = ctl[ev.site] || { na: new Set() };
        const quand = ev.data.quand || ev.jour + 'T12:00:00';
        if (!ev.data.partiel && (!c[ev.data.niveau] || c[ev.data.niveau].at <= quand)) c[ev.data.niveau] = { at: quand, agent: ev.agent, rid: ev.rid || null };
        Object.entries(ev.data.items || {}).forEach(([k, v]) => { if (v === 'NA' && !(ev.vu != null && c.retabli && (c.retabli[k] || 0) > ev.vu)) c.na.add(k); });
        ctlHist.push({ site: ev.site, niveau: ev.data.niveau, at: quand, ts, agent: ev.agent, items: ev.data.items, libelles: ev.data.libelles, rid: ev.rid, partiel: !!ev.data.partiel });
      } else if (ev.t === 'barriere') {
        const d = ev.data;
        const b = barr.get(d.bid) || { bid: d.bid, site: ev.site, barriere: d.barriere, ouverte: null, fermee: null };
        const rec = { quand: d.quand, agent: ev.agent, motif: d.motif || '', comment: d.comment || '', ts, dev: ev.dev };
        if (d.action === 'ouverte' && !b.ouverte) b.ouverte = rec;
        if (d.action === 'fermee' && !b.fermee) b.fermee = rec;
        if (b.ouverte) barr.set(d.bid, b);
      }
      if (sup && ev.rid && !del.rids.has(ev.rid) && !annulees.has(ev.rid)) {
        const r = rondes.get(ev.rid) || { rid: ev.rid, site: ev.site, agent: ev.agent, dev: ev.dev, jour: ev.jour, nbSig: 0, nbRevue: 0, compte: null, alertes: [] };
        if (ev.t === 'ronde_debut') { r.debut = ts; r.qr = ev.data.qr; r.geoDebut = dec && dec.geo; }
        if (ev.t === 'ronde_agent') { r.agent = ev.data.nouveau; (r.changements = r.changements || []).push(`${ev.data.ancien} → ${ev.data.nouveau}`); }
        if (ev.t === 'ronde_fin') { r.fin = ts; r.closDPMS = false; if (r.compte == null && ev.data.compte != null) r.compte = ev.data.compte; }
        if (ev.t === 'checklist') r.checklist = ev.data.items;
        if (ev.t === 'comptage') r.compte = ev.data.total;
        if (ev.t === 'controle') { r.controles = r.controles || []; r.controles.push({ niveau: ev.data.niveau, nc: resumeControle(ev.data.items, ev.data.libelles).nc }); }
        if (ev.t === 'signalement') r.nbSig++;
        if (ev.t === 'revue') r.nbRevue++;
        rondes.set(ev.rid, r);
      }
    } else if (e.kind === 'dec') {
      const d = e.dec;
      if (d.id && d.type !== 'cfg') { if (decVues.has(d.id)) continue; decVues.add(d.id); }   // même décision reçue deux fois
      if (d.type === 'cfg') { if (validCfg(normalizeCfg(d.cfg))) cfg = normalizeCfg(d.cfg); }
      else if (d.type === 'action') {
        const s = sigs.get(d.action.sig);
        if (s) s.actions.push({ ...d.action, quand: d.action.quand || d.ts, jour: (d.ts || '').slice(0, 10), par: d.par || 'DPMS' });
      }
      else if (d.type === 'msg') { const s = sigs.get(d.sig); if (s) s.messages.push({ id: d.id, role: 'dpms', de: 'DPMS', quand: d.ts, texte: d.texte || '', demande: !!d.demande }); }
      else if (d.type === 'demande_close') { const s = sigs.get(d.sig); if (s) s.messages.push({ id: d.id, role: 'dpms', de: 'DPMS', quand: d.ts, texte: '', cloture: true }); }
      else if (d.type === 'vapid') { if (!vapid && d.pub && d.priv) vapid = { pub: d.pub, priv: d.priv }; }
      else if (d.type === 'contact') { mcAjout(mcs, d.contact, d.par || 'DPMS'); const c = mcs.get(d.contact && d.contact.id); if (c && c.statutN == null) marquer(c, e, 'dpms'); }
      else if (d.type === 'contact_suivi') mcSuivi(mcs, d.suivi, d.par || 'DPMS', { connu: () => true, marquer: x => marquer(x, e, 'dpms') });
      else if (d.type === 'planning' && d.du && d.au && d.jours) {
        for (let iso = d.du; iso <= d.au; iso = ajJ(iso, 1)) planning.jours[iso] = d.jours[iso] || {};
        planning.notes[d.mois] = d.notes || '';
        planning.pubs.push({ id: d.id, ts: d.ts, mois: d.mois, du: d.du, au: d.au, chg: d.chg || [] });
      }
      else if (d.type === 'ronde_close' && d.rid) {
        if (!rondesCloses.has(d.rid)) rondesCloses.set(d.rid, d);
        const r = sup && rondes.get(d.rid);
        if (r && !r.fin) { r.fin = d.fin || d.ts; r.closDPMS = true; }   // une fin déjà reçue de l'agent prime
      }
      else if (d.type === 'registre') regs.push({ site: d.site, at: d.ts, note: d.note || '', par: d.par || 'DPMS' });
      else if (d.type === 'ctl_reactiver') { const c = ctl[d.site] = ctl[d.site] || { na: new Set() }; c.na.delete(d.item); (c.retabli = c.retabli || {})[d.item] = nDe(e); }
      else if (d.type === 'sig') {
        const s = sigs.get(d.sig);
        if (s) { Object.assign(s, d.patch); if (d.patch && 'statut' in d.patch) marquer(s, e, 'dpms'); s.journalDPMS.push({ ts: d.ts, action: d.action, par: d.par || 'DPMS' }); }
      }
    } else if (e.kind === 'note' && sup) {
      if (e.dec.type === 'planning_motifs' && e.dec.du && e.dec.au) {   // motifs d'absence et lignes internes : superviseurs seulement
        for (let iso = e.dec.du; iso <= e.dec.au; iso = ajJ(iso, 1)) planMotifs[iso] = { ...((e.dec.motifs || {})[iso] || {}), ...((e.dec.internes || {})[iso] || {}) };
      } else {
        const s = sigs.get(e.dec.sig);
        if (s) s.notes = e.dec.text;
      }
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
          if (e.supprime || (ev.rid && del.rids.has(ev.rid))) return;
          alertes.push({ key: `${dev}:${ev.seq}:${m}`, at: e.at, msg: `${lieu} — saisie n° ${ev.seq} (${dev}) : ${m}` });
          if (ev.rid && rondes.has(ev.rid)) rondes.get(ev.rid).alertes.push(m);
        };
        if (!e.hashOk) add('empreinte invalide (saisie modifiée)');
        if (prev === null) { if (ev.seq !== 1) add(`saisies 1 à ${ev.seq - 1} manquantes`); else if (ev.prev !== '0'.repeat(64)) add('chaînage rompu'); }
        else if (ev.seq !== prev.ev.seq + 1) add(`saisies ${prev.ev.seq + 1} à ${ev.seq - 1} manquantes`);
        else if (ev.prev !== prev.ev.hash) add('chaînage rompu avec la saisie précédente');
        if (!e.dec) add('horodatage et position illisibles');
        else {
          (e.dec.photoAges || []).forEach(a => { if (a.ageMin != null && a.ageMin > 15) add(`photo prise ${a.ageMin} min avant la saisie`); });
          if (ev.t === 'ronde_debut' && e.dec.geo && e.dec.geo.err) add('position : ' + e.dec.geo.err);
        }
        prev = e;
      }
    }
    for (const r of rondes.values()) if (r.closDPMS) r.alertes.push('clôturée par le DPMS (fin de ronde non faite par l’agent)');
    for (const r of rondes.values()) if (r.debut && !r.fin && Date.now() - new Date(r.debut).getTime() > 6 * 3600000) r.alertes.push('fin de ronde non reçue');
    await dbReplaceAll('rondes', [...rondes.values()]);
    await kvSet('alertes', alertes.sort((a, b) => (b.at || '').localeCompare(a.at || '')));
    await kvSet('devs', devs);
  }
  for (const s of sigs.values()) {   // fil d'échanges : ordre chronologique, demande d'action en attente
    s.messages = (s.messages || []).sort((a, b) => (a.quand || '').localeCompare(b.quand || ''));
    let dem = null;
    for (const m of s.messages) { if (m.demande) dem = m; else if (m.fait || m.cloture) dem = null; }
    s.demande = dem ? dem.id : null;
  }
  await dbReplaceAll('sigs', [...sigs.values()]);
  for (const x of sigs.values()) if (x.mc && mcs.has(x.mc)) mcs.get(x.mc).sigs.push(x.id);
  await kvSet('mainCourante', [...mcs.values()].sort((a, b) => b.quand.localeCompare(a.quand)));
  await kvSet('rondesCloses', [...rondesCloses.keys()]);
  await kvSet('planning', planning);
  if (sup) await kvSet('planningMotifs', planMotifs);
  if (vapid) await kvSet('vapid', vapid);
  await kvSet('barrieres', [...barr.values()]);
  await kvSet('supprimes', { sigs: [...del.sigs], rids: [...del.rids] });
  await kvSet('controles', Object.fromEntries(Object.entries(ctl).map(([k, v]) => [k, { ...v, na: [...v.na], retabli: undefined }])));
  if (sup) await kvSet('controlesHist', ctlHist.sort((a, b) => b.at.localeCompare(a.at)));
  if (sup) await kvSet('registres', regs.sort((a, b) => b.at.localeCompare(a.at)));
  if (sup) await kvSet('corbeille', [...suppr.values()].map(d => ({ ...d, annule: annul.has(d.id) })).sort((a, b) => b.ts.localeCompare(a.ts)));
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
      vu: (await kvGet('cursor')) || 0,   // dernier élément du journal connu de l'appareil (ordre causal)
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
// Téléphone partagé : hors ronde, aucun nom n'est présélectionné ; l'agent choisit le sien à chaque saisie.
async function agentPrefere() {
  const r = await kvGet('ronde');
  return r ? r.agent : null;
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
      maintenirPush();
      majBadge();
      return r;
    } catch (e) { await kvSet('syncErr', { at: new Date().toISOString(), msg: e.message }); return { err: e.message }; }
    finally { state.sync.busy = false; _syncP = null; }
  })();
  return _syncP;
}
let _notifSup = [], _notifAgents = [];
const NOTIF_TYPES = {
  superviseur: [['urgent', 'Signalements urgents'], ['sig', 'Autres nouveaux signalements'], ['aggrave', 'Signalements aggravés'], ['clos', 'Signalements clos'],
    ['action', 'Traitement (appels, courriels, interventions)'], ['pm', 'Suites de la police municipale'], ['msg', 'Réponses des agents'],
    ['barriere_ouverte', 'Barrière ouverte'], ['barriere_fermee', 'Barrière refermée'], ['ronde_fin', 'Ronde terminée'], ['ronde_annulee', 'Ronde annulée'],
    ['contact', 'Main courante (usagers, demandes, réponses)'],
    ['alerte_ronde', 'Aucune ronde dans la journée'], ['ronde_longue', 'Ronde de plus d’1 h'], ['synthese', 'Synthèse hebdomadaire'], ['rappel_planning', 'Rappel du 15 : planning du mois suivant']],
  agent: [['msg', 'Messages et demandes d’action du DPMS'], ['planning', 'Planning publié ou modifié']]
};
const GROUPES_NOTIF = { urgent: 'Signalements', barriere_ouverte: 'Barrières', ronde_fin: 'Rondes', contact: 'Main courante', alerte_ronde: 'Alertes programmées (envoyées par le relais)' };
const typesNotif = () => NOTIF_TYPES[state.mode === 'superviseur' ? 'superviseur' : 'agent'];
async function prefsNotif() {
  const p = await kvGet('notifPrefs'), tous = typesNotif().map(x => x[0]);
  if (!p) return tous;
  const connus = (await kvGet('notifConnus')) || tous.filter(t => !['contact', 'planning', 'rappel_planning'].includes(t));   // types ajoutés après le choix : cochés d'office
  const nouveaux = tous.filter(t => !connus.includes(t));
  if (nouveaux.length) { p.push(...nouveaux); await kvSet('notifPrefs', p); await kvSet('notifConnus', tous); }
  return p;
}
/* Ronde clôturée par le DPMS (agent qui a oublié de terminer) : le téléphone qui la porte l'arrête,
   envoie ce qui a été contrôlé (contrôle quotidien partiel), sans fin de ronde de l'agent. */
async function appliquerClotureDPMS() {
  const ronde = await kvGet('ronde');
  if (!ronde || !((await kvGet('rondesCloses')) || []).includes(ronde.rid)) return;
  const st = await etatControles();
  const sigsRonde = (await dbAll('sigs')).filter(s => ronde.sigsCrees.includes(s.id));
  const pts = pointsSite(st, ronde.site, 'quotidien');
  const rep = (ronde.ctl || {}).quotidien || {};
  if (pts.some(c => rep[c.id] || ncDe(sigsRonde, c.id).length)) {
    const items = {}; pts.forEach(c => { items[c.id] = ncDe(sigsRonde, c.id).length ? 'NC' : (rep[c.id] || 'NF'); });
    await appendEvent('controle', ronde, { niveau: 'quotidien', items, libelles: Object.fromEntries(pts.map(c => [c.id, c.lbl])), quand: new Date().toISOString(), partiel: true });
  }
  await kvDel('ronde');
  Object.keys(revueDraft).forEach(k => delete revueDraft[k]);
  Geo.stop();
  await kvSet('rondeCloseDPMS', { site: ronde.site, agent: ronde.agent, debut: ronde.debut, at: new Date().toISOString() });
  await fold();
  if (['ronde', 'revue', 'controle'].includes(state.view) || (state.view === 'signalement' && state.params && state.params.from === 'ronde')) { sigDraft = null; go('home'); }
  else if (state.view === 'home') render();
  setTimeout(() => syncNow(), 500);
}
/* Alertes tenues par le relais : « une ronde a commencé ici aujourd'hui » et « ronde pas finie au bout d'1 h » */
async function opsPourRelais(batch) {
  const ops = [];
  const finies = new Set(batch.filter(ev => ev.t === 'ronde_fin' || ev.t === 'ronde_annulee').map(ev => ev.rid));
  for (const ev of batch) {
    if (ev.t === 'ronde_debut') {
      ops.push({ op: 'pointer', site: ev.site });
      if (!finies.has(ev.rid)) ops.push({ op: 'differer', rid: ev.rid, site: ev.site, agent: ev.agent, debut: (ev.data && ev.data.debut) || new Date().toISOString() });
    }
    if ((ev.t === 'ronde_fin' || ev.t === 'ronde_annulee') && ev.rid) ops.push({ op: 'annuler', rid: ev.rid });
  }
  if (ops.length) await kvUpdate('opsRelais', l => [...(l || []), ...ops.map(o => ({ ...o, oid: idFile() }))].slice(-50));
}
let _opsEnCours = false;
async function traiterOpsRelais() {
  if (_opsEnCours) return; _opsEnCours = true;
  try {
    for (let garde = 0; garde < 100; garde++) {
      const ops = (await kvGet('opsRelais')) || [];
      if (!ops.length) break;
      const o = ops[0];
      try {
        if (o.op === 'pointer') await rel('pointer', { jetons: [await jeton('site', o.site)] }, 20000);
        else if (o.op === 'annuler') await rel('annulerDiffere', { id: await jeton('ronde', o.rid) }, 20000);
        else if (o.op === 'differer') {
          const r = await rel('abos', { cible: 'sup' }, 20000);
          const messages = await chiffrerPour(r.abos, 'ronde_longue', { type: 'ronde_longue', titre: `Ronde de plus d’1 h — ${siteNom(o.site)}`, corps: `${o.agent || ''} : commencée à ${fmtHeure(o.debut)}, toujours pas terminée.` });
          if (Object.keys(messages).length) await rel('differer', { id: await jeton('ronde', o.rid), apres: new Date(new Date(o.debut).getTime() + 3600000).toISOString(), messages }, 20000);
        }
      } catch (e) { if (!/opération inconnue|invalide/.test(e.message)) break; }   // relais ancien : opération ignorée
      await kvUpdate('opsRelais', l => (l || []).filter(x => (o.oid ? x.oid !== o.oid : JSON.stringify(x) !== JSON.stringify(o))));
    }
  } finally { _opsEnCours = false; }
}
async function notifPourSup(ev) {
  const lieu = siteNom(ev.site), qui = ev.agent || '';
  const sig = ev.data && ev.data.sig ? await dbGet('sigs', ev.data.sig) : null;
  if (ev.data && ev.data.sig && !sig && ['revue', 'action', 'pm', 'msg'].includes(ev.t)) return null;   // signalement supprimé par le DPMS
  if (ev.t === 'signalement' && !(await dbGet('sigs', ev.data.id))) return null;
  const quoi = sig ? `${sig.ref} ${sigTitre(sig)}` : (ev.data && ev.data.ref) || '';
  if (ev.t === 'revue' && ev.data.verdict === 'resolu') return { type: 'clos', titre: `Signalement clos — ${lieu}`, corps: `${quoi} — ${qui}${ev.data.comment ? ' : ' + ev.data.comment : ''}` };
  if (ev.t === 'revue' && ev.data.verdict === 'aggrave') return { type: 'aggrave', titre: `Signalement aggravé — ${lieu}`, corps: `${quoi} — ${qui}${ev.data.comment ? ' : ' + ev.data.comment : ''}` };
  if (ev.t === 'action') return { type: 'action', titre: `${ACTIONS[ev.data.type] || 'Traitement'} — ${lieu}`, corps: `${quoi}${ev.data.dest ? ' — ' + ev.data.dest : ''}${ev.data.prevu ? ' — le ' + fmtDT(ev.data.prevu) : ''}${ev.data.texte ? ' : ' + ev.data.texte : ''} (${qui})` };
  if (ev.t === 'pm') return { type: 'pm', titre: `Police municipale — ${PM_STATUTS[ev.data.statut] || ''}`, corps: `${quoi}${ev.data.obs ? ' : ' + ev.data.obs : ''}` };
  if (ev.t === 'barriere' && ev.data.action === 'fermee') return { type: 'barriere_fermee', titre: `Barrière refermée — ${lieu}`, corps: `${ev.data.barriere} — ${qui}` };
  if (ev.t === 'ronde_annulee') return { type: 'ronde_annulee', titre: `Ronde annulée — ${lieu}`, corps: qui };
  if (ev.t === 'signalement') return { type: ev.data.urgent ? 'urgent' : 'sig', titre: `${ev.data.urgent ? 'URGENT — ' : ''}Signalement — ${lieu}`, corps: `${ev.data.cat}${ev.data.sub ? ' — ' + ev.data.sub : ''}${ev.data.horsRonde ? ' (hors ronde)' : ''} — ${qui}` };
  if (ev.t === 'barriere' && ev.data.action === 'ouverte') return { type: 'barriere_ouverte', titre: `Barrière ouverte — ${lieu}`, corps: `${ev.data.barriere} : ${ev.data.motif || ''} — ${qui}` };
  if (ev.t === 'ronde_fin') return { type: 'ronde_fin', titre: `Ronde terminée — ${lieu}`, corps: `${qui} : ${ev.data.compte} véhicule(s), ${ev.data.nbSig} signalement(s)` };
  if (ev.t === 'contact') return { type: 'contact', titre: `Main courante — ${ev.data.site ? siteNom(ev.data.site) : 'général'}`, corps: `${CANAUX[ev.data.canal] || ''} : ${ev.data.objet} — ${ev.data.nom}. ${SUITES[ev.data.suite] || ''}${ev.data.service ? ' : ' + ev.data.service : ''} (${qui})` };
  if (ev.t === 'contact_suivi' && ev.data.type === 'reponse') return { type: 'contact', titre: `Réponse à l’usager — ${ev.data.ref}`, corps: `${ev.data.nom} : ${ev.data.texte}${ev.data.regle ? ' (réglée)' : ' (attente)'} — ${qui}` };
  if (ev.t === 'msg') return { type: 'msg', titre: `Réponse d’un agent — ${ev.data.ref}`, corps: `${qui} : ${ev.data.fait ? '[action faite] ' : ''}${ev.data.texte || ''}` };
  return null;
}
async function agentSync() {
  const dev = await kvGet('dev');
  let sent = 0;
  const envoyes = [];
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
    envoyes.push(...batch);   // notifications calculées après la relève
    await opsPourRelais(batch);
    for (let i = 0; i < batch.length; i++) await storeEvEntry(batch[i], r.n, i, r.at, { media: !!media });
    const d = await openDB(); const t = d.transaction(['events', 'photos'], 'readwrite');
    batch.forEach(ev => { t.objectStore('events').delete(ev.seq); ev.photos.forEach(p => t.objectStore('photos').delete(p.id)); });
    await txDone(t);
    sent += batch.length;
  }
  const got = await pullAll(['p', 'd']);
  if (sent || got) await fold();
  await appliquerClotureDPMS();
  for (const ev of envoyes) { const n = await notifPourSup(ev); if (n) _notifSup.push(n); }
  if (_notifSup.length) { const l = _notifSup; _notifSup = []; notifier('sup', l); }
  traiterOpsRelais();
  return { sent, got };
}
async function pullAll(kinds) {
  let after = (await kvGet('cursor')) || 0, got = 0;
  for (let guard = 0; guard < 400; guard++) {
    const r = await rel('pull', { after, kinds, max: 60 }, 180000);
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

/* ---------- Mise à jour de l'application ----------
   Le téléphone ou l'ordinateur peut rester ouvert des jours : la présence d'une nouvelle version est vérifiée
   à chaque retour sur l'application et toutes les 15 min ; elle est installée dès que l'écran ne contient
   aucune saisie en cours (accueil, listes), jamais au milieu d'un signalement. */
let _swReg = null, _majPrete = false, _majVerifiee = 0;
function chercherMiseAJour(force) {
  if (!_swReg || (!force && Date.now() - _majVerifiee < 15 * 60000)) return;
  _majVerifiee = Date.now();
  _swReg.update().catch(() => { });
}
function saisieEnCours() {
  if (VUES_SAISIE.includes(state.view) || ['contact', 'changerRondier', 'supConfig'].includes(state.view)) return true;
  if (state.params && state.params.mode) return true;
  if (document.getElementById('busy') || document.querySelector('.demande, .confirm-contact, .lightbox')) return true;
  return [...document.querySelectorAll('main textarea, main input[type=text], main input[type=search], main input[type=email], main input[type=tel]')].some(x => x.value && x.value.trim() && x.defaultValue !== x.value);
}
function appliquerMiseAJour() {
  if (!_majPrete || saisieEnCours()) return false;
  _majPrete = false; location.reload(); return true;
}
if ('serviceWorker' in navigator) {
  const avaitControleur = !!navigator.serviceWorker.controller;   // premier chargement : pas une mise à jour
  navigator.serviceWorker.addEventListener('message', e => {
    if (e.data && e.data.type === 'maj') { if (e.ports && e.ports[0]) e.ports[0].postMessage('ok'); _majPrete = true; appliquerMiseAJour(); }
  });
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (avaitControleur) { _majPrete = true; appliquerMiseAJour(); } });
}
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') { chercherMiseAJour(); appliquerMiseAJour(); } });
setInterval(() => { chercherMiseAJour(); appliquerMiseAJour(); }, 60000);

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
      h('p', { class: 'muted' }, 'Sur un appareil superviseur, ouvrez « Configurer les appareils » puis scannez le QR correspondant avec cet appareil (ou ouvrez le lien copié).'),
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
  const node = page('Scanner', { back: true }, wrap, msg);
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
  const m = /^#(join|site|cfg|sup|ptype|dests)=(.+)$/.exec(hash);
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
  if (l.kind === 'dests') {   // lien « destinataires » préparé par Claude : pré-remplit l'écran, le superviseur vérifie et enregistre
    let p = null; try { p = JSON.parse(td.decode(fflate.inflateSync(b64u.dec(l.val)))); } catch (e) { p = null; }
    if (!p || !Array.isArray(p.dests)) { toast('Lien illisible.', true); go('home'); return; }
    if (state.mode !== 'superviseur') { toast('Ce lien s’ouvre sur l’application superviseur.', true); go('home'); return; }
    go('supConfig', { prefillDests: p }); return;
  }
  if (l.kind === 'ptype') {   // lien « planning-type et règles » préparé par Claude : pré-remplit l'écran, le superviseur vérifie et enregistre
    let p = null; try { p = JSON.parse(td.decode(fflate.inflateSync(b64u.dec(l.val)))); } catch (e) { p = null; }
    if (!p || typeof p.type !== 'string') { toast('Lien de planning illisible.', true); go('home'); return; }
    if (state.mode !== 'superviseur') { toast('Ce lien s’ouvre sur l’application superviseur.', true); go('home'); return; }
    go('planningParam', { prefill: { type: p.type, regles: typeof p.regles === 'string' ? p.regles : null, ancre: /^\d{4}-\d{2}-\d{2}$/.test(p.ancre || '') ? p.ancre : null } }); return;
  }
  if (l.kind === 'sup') {   // QR « autre superviseur » : adresse du relais seulement ; la phrase de passe reste demandée
    let u = ''; try { u = td.decode(b64u.dec(l.val)); } catch (e) { u = ''; }
    if (!/^https?:\/\/[^\s]+$/.test(u)) { toast('QR superviseur illisible.', true); go('home'); return; }
    if (state.mode) { toast(state.mode === 'superviseur' ? 'Cet appareil est déjà superviseur.' : 'Cet appareil est déjà configuré : Réglages → Réinitialiser cet appareil, puis scanner à nouveau.', true); go('home'); return; }
    go('supJoin', { url: u }); return;
  }
  if (l.kind === 'site') { toast('Les QR de parking ne sont plus utilisés : appuyez sur « Commencer ma ronde ».'); go('home'); return; }
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
const ICO = {
  ronde: '<svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="5" y="3" width="14" height="18" rx="2"/><path d="M9 3v2h6V3"/><path d="m9 12 2 2 4-4"/></svg>',
  alerte: '<svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3 2 20h20L12 3z"/><path d="M12 10v4"/><path d="M12 17.5v.01"/></svg>',
  barriere: '<svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 21V9"/><circle cx="4" cy="7" r="2"/><path d="M6 8l15-4"/><path d="M10 6.9v2M14 5.8v2M18 4.8v2"/></svg>',
  usager: '<svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="9" cy="7" r="4"/><path d="M2 21v-1a6 6 0 0 1 6-6h2a6 6 0 0 1 6 6v1"/><path d="M16 3h6v5h-3l-3 2z"/></svg>',
  ok: '<svg viewBox="0 0 24 24" width="64" height="64" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="m7 12 3.5 3.5L17 9"/></svg>'
};
const gros = (ico, txt, cls, onclick) => h('button', { class: 'tuile ' + cls, onclick }, h('span', { class: 'tuile-ico', html: ICO[ico] }), h('span', null, txt));

/* Fenêtre de confirmation de l'application (à la place des fenêtres du navigateur) */
function demander({ titre, texte, oui = 'Oui', non = 'Non', danger = false }) {
  return new Promise(res => {
    const fin = v => { box.remove(); res(v); };
    const box = h('div', { class: 'confirm-contact demande' }, h('div', { class: 'sheet' },
      h('h3', null, titre), texte ? h('p', { style: 'white-space:pre-line' }, texte) : null,
      h('button', { class: (danger ? 'danger' : 'ok') + ' big', onclick: () => fin(true) }, oui),
      h('button', { class: 'sec', onclick: () => fin(false) }, non)));
    document.body.append(box);
  });
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
  const ouvertes = await barrieresOuvertes();
  const cfg = state.cfg;
  const closeDPMS = await kvGet('rondeCloseDPMS');
  const nbEnCours = (await dbAll('sigs')).filter(s => s.statut === 'ouvert').length;
  const majRequise = await kvGet('majRequise');
  return page('Rondes parkings', null,
    majRequise ? bandeauMaj() : null,
    closeDPMS ? h('div', { class: 'card', style: 'border:2px solid var(--warn);background:var(--warn-bg)' },
      h('b', null, `Ronde clôturée par le DPMS`),
      h('p', { style: 'margin:4px 0' }, `${siteNom(closeDPMS.site)} — ${closeDPMS.agent}, commencée à ${fmtHeure(closeDPMS.debut)}. Elle n’avait pas été terminée. Les contrôles déjà faits ont été envoyés.`),
      h('button', { class: 'sec', onclick: safe(async () => { await kvDel('rondeCloseDPMS'); render(); }) }, 'J’ai compris')) : null,
    await carteMessagesAccueil(),
    await carteATransmettre(),
    await cartePlanningAgents(),
    ronde ? h('div', { class: 'card encours' },
      h('div', { class: 'small' }, 'Ronde en cours'),
      h('div', { class: 'grand' }, `${siteNom(ronde.site)} — ${ronde.agent}`),
      h('div', { class: 'muted small' }, `commencée à ${fmtHeure(ronde.debut)}`),
      gros('ronde', 'Reprendre la ronde', 'accent', () => resumeRonde(ronde)),
      lienAnnuler())
      : gros('ronde', 'Commencer ma ronde', 'principal', () => go('choixParking')),
    gros('alerte', 'Signalement hors ronde', 'sec', () => go('signalement', { from: 'hr' })),
    gros('barriere', 'Barrière ouverte', 'sec', () => go('barriere', {})),
    gros('usager', 'Usager ou demande reçue', 'sec', () => { mcDraft = null; go('contact'); }),
    ouvertes.length ? h('div', { class: 'card', style: 'border:2px solid var(--bad)' }, h('h2', null, 'Barrières ouvertes'), ...ouvertes.map(b => barriereLigne(b, true))) : null,
    nbEnCours ? h('button', { class: 'sec', style: 'margin-top:18px', onclick: () => go('suivi') }, `Suivi des signalements en cours (${nbEnCours})`) : null,
    h('button', { class: 'sec', style: 'margin-top:10px', onclick: () => go('plus') }, 'Plus : usagers, échanges, réglages…'),
    cfg.urgenceTel ? h('a', { class: 'btn danger', style: 'margin-top:10px', href: 'tel:' + cfg.urgenceTel }, 'Urgence : appeler le DPMS') : null,
    h('img', { class: 'logo-ville', src: 'logo-cachan.png', alt: 'Ville de Cachan — parkings' }),
    h('p', { class: 'muted small foot' }, enAttente ? `${enAttente} saisie(s) en attente d’envoi (envoi automatique dès qu’il y a du réseau)` : 'Tout est envoyé au DPMS.', ` · version ${APP_VERSION}`));
}
VIEWS.plus = async () => {
  const lastSync = await kvGet('lastSync');
  const syncErr = await kvGet('syncErr');
  const enAttente = (await dbAll('events')).length;
  const cfg = state.cfg;
  const enCours = (await dbAll('sigs')).filter(s => s.statut === 'ouvert');
  const aTransmettre = enCours.filter(s => suiviEtat(s) === 'À transmettre').length;
  const nbMC = (await mainCourante()).filter(c => c.statut === 'ouvert').length;
  return page('Plus', { back: true },
    h('button', { class: 'sec', onclick: () => go('suivi') }, `Suivi des signalements en cours (${enCours.length}${aTransmettre ? `, dont ${aTransmettre} à transmettre` : ''})`),
    h('button', { class: 'sec', onclick: () => go('messagesAgents') }, 'Échanges avec le DPMS'),
    h('button', { class: 'sec', onclick: () => go('mainCourante', {}) }, `Suivi des usagers et demandes${nbMC ? ` — ${nbMC} à traiter` : ''}`),
    await carteControlesAFaire(),
    h('div', { class: 'card' }, h('h2', null, 'Signalements en cours par parking'),
      ...cfg.sites.map(s => h('div', { class: 'stat' }, h('span', null, s.nom), h('b', null, enCours.filter(x => x.site === s.id).length)))),
    syncCard(enAttente, lastSync, syncErr),
    await carteAppareil(),
    h('div', { class: 'card' }, h('h2', null, 'Extrait Excel'),
      h('div', { class: 'row' },
        h('button', { class: 'sec', onclick: safe(() => agentExcel(7)) }, '7 jours'),
        h('button', { class: 'sec', onclick: safe(() => agentExcel(31)) }, '31 jours'),
        h('button', { class: 'sec', onclick: safe(() => agentExcel(HISTO_JOURS)) }, 'Tout'))),
    cfg.urgenceMail ? h('a', { class: 'btn sec', href: 'mailto:' + cfg.urgenceMail + '?subject=' + encodeURIComponent('URGENT — parking') }, 'Écrire au DPMS') : null,
    h('p', { class: 'muted small' }, 'En cas de danger immédiat : 17 ou 112.'),
    h('button', { class: 'link', onclick: () => go('reglages') }, 'Réglages'),
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
};

async function carteControlesAFaire() {
  const st = await etatControles();
  const lignes = state.cfg.sites.map(s => {
    const dus = Object.keys(NIVEAUX).filter(n => echeance(st, s.id, n).du);
    return dus.length ? h('div', { class: 'stat' }, h('span', null, s.nom), h('span', { class: 'small' }, ...dus.map(n => h('span', { class: 'badge ' + (n === 'quotidien' ? 'b-info' : 'b-warn') }, n)))) : null;
  }).filter(Boolean);
  return lignes.length ? h('div', { class: 'card' }, h('h2', null, 'Contrôles à faire'), ...lignes, h('p', { class: 'muted small' }, 'Ils sont proposés pendant la ronde du parking concerné.')) : null;
}
VIEWS.choixParking = async () => page('Commencer ma ronde', { back: true },
  h('p', { class: 'question' }, 'Quel parking ?'),
  ...(state.cfg ? state.cfg.sites : []).map(s => h('button', { class: 'big sec choix', onclick: safe(() => startRondeFlow(s.id)) }, s.nom)));

async function startRondeFlow(siteId) {
  const cur = await kvGet('ronde');
  if (cur) {
    if (cur.site === siteId) { resumeRonde(cur); return; }
    if (!(await demander({ titre: 'Une ronde est déjà en cours', texte: `Ronde à ${siteNom(cur.site)} (${cur.agent}). La terminer et commencer à ${siteNom(siteId)} ?`, oui: 'Oui, la terminer', non: 'Non' }))) { go('home'); return; }
    await finishRonde(cur, true);
  }
  go('choixAgent', { site: siteId });
}
VIEWS.choixAgent = async ({ site }) => page(`Ronde : ${siteNom(site)}`, { back: true },
  h('p', { class: 'question' }, 'Qui fait la ronde ?'),
  ...state.cfg.agents.map(a => h('button', { class: 'big sec choix', onclick: safe(() => beginRonde(site, a)) }, a)));

async function beginRonde(site, agent) {
  await kvDel('rondeCloseDPMS');
  showBusy('Démarrage de la ronde…');
  Geo.start();
  if (navigator.onLine) await Promise.race([syncNow().catch(() => null), new Promise(r => setTimeout(r, 6000))]);   // clôtures et suppressions du DPMS
  await Geo.waitFirst(4000);
  const st = await etatControles();
  const ouverts = (await dbAll('sigs')).filter(s => s.site === site && s.statut === 'ouvert').length;
  const etapes = [];
  if (ouverts) etapes.push({ id: 'revue', lbl: 'Signalements en cours' });
  etapes.push({ id: 'comptage', lbl: 'Comptage des véhicules' }, { id: 'quotidien', lbl: 'Contrôle quotidien' });
  ['mensuel', 'trimestriel'].forEach(n => { if (echeance(st, site, n).du && pointsSite(st, site, n).length) etapes.push({ id: n, lbl: NIVEAUX[n].lbl }); });
  etapes.push({ id: 'fin', lbl: 'Fin de ronde' });
  const ronde = { rid: uuid(), site, agent, debut: new Date().toISOString(), etapes, step: etapes[0].id, pos: {}, revus: [], reporte: {}, lance: {},
    checklist: {}, sigsCrees: [], compte: 0, compteValide: false, quotidienDu: true, ctl: {}, ctlEmis: {} };
  await appendEvent('ronde_debut', ronde, { debut: ronde.debut });
  await kvSet('ronde', ronde);
  await kvSet('dernierAgent', { nom: agent, at: Date.now() });
  hideBusy();
  resumeRonde(ronde);
}
function bandeauRondier(ronde, boutons = true) {
  return h('div', { class: 'rondier' + (boutons ? '' : ' compact') },
    h('div', null, h('div', { class: 'small' }, `Rondier — ${siteNom(ronde.site)}, depuis ${fmtHeure(ronde.debut)}`), h('b', null, ronde.agent)),
    boutons ? h('div', { class: 'rondier-btns' },
      h('button', { class: 'sec', onclick: () => go('changerRondier') }, 'Changer le rondier'),
      h('button', { class: 'danger', onclick: safe(annulerRonde) }, 'Annuler la ronde')) : null);
}
VIEWS.changerRondier = async () => {
  const ronde = await kvGet('ronde');
  if (!ronde) return agentHome();
  const retour = () => go('ronde', {}, { noPush: true });
  return page('Changer le rondier', { back: retour },
    h('div', { class: 'banner info' }, `Ronde en cours à ${siteNom(ronde.site)}, commencée à ${fmtHeure(ronde.debut)} au nom de ${ronde.agent}. Ce qui a déjà été saisi passe au nom du rondier choisi ; la ronde continue.`),
    h('p', { class: 'question' }, 'Qui fait la ronde ?'),
    ...state.cfg.agents.filter(a => a !== ronde.agent).map(a => h('button', { class: 'big sec choix', onclick: safe(async () => {
      const r = await kvGet('ronde');
      await appendEvent('ronde_agent', r, { ancien: r.agent, nouveau: a });
      r.agent = a; await kvSet('ronde', r);
      await fold();
      toast(`Rondier : ${a}.`);
      retour();
    }) }, a)));
};
function resumeRonde(ronde) {
  Geo.start();
  go('ronde');
}

/* ---------- Contrôles : quotidien (à chaque ronde), mensuel (30 j), trimestriel (90 j) ---------- */
async function etatControles() { return (await kvGet('controles')) || {}; }
function joursDepuis(iso) { return Math.round((new Date(localDate() + 'T12:00:00') - new Date(localDate(new Date(iso)) + 'T12:00:00')) / 86400000); }
function echeance(st, site, niveau) {
  const d = st[site] && st[site][niveau];
  if (!d) return { du: true, dernier: null, depuis: null };
  const j = joursDepuis(d.at);
  return { du: niveau === 'quotidien' ? j >= 1 : j >= NIVEAUX[niveau].jours, dernier: d, depuis: j };
}
function pointsSite(st, site, niveau) {
  const na = new Set((st[site] && st[site].na) || []);
  return state.cfg.controles.filter(c => c.niveau === niveau && !na.has(c.id));
}
function catPour(c) { return state.cfg.cats.some(x => x.nom === c.cat) ? c.cat : CAT_AUTRE; }
function ncDe(sigsRonde, id) { return sigsRonde.filter(s => s.controle && s.controle.id === id); }
function restantControle(ronde, niveau, points, sigsRonde) {
  const rep = (ronde.ctl || {})[niveau] || {};
  return points.filter(c => !ncDe(sigsRonde, c.id).length && !rep[c.id]).length;
}
async function emettreControle(ronde, niveau, points, sigsRonde) {
  const rep = (ronde.ctl || {})[niveau] || {};
  const items = {};
  points.forEach(c => { items[c.id] = ncDe(sigsRonde, c.id).length ? 'NC' : (rep[c.id] || 'NF'); });
  await appendEvent('controle', ronde, { niveau, items, libelles: Object.fromEntries(points.map(c => [c.id, c.lbl])), quand: new Date().toISOString() });
  ronde.ctlEmis = { ...(ronde.ctlEmis || {}), [niveau]: true };
  await kvSet('ronde', ronde);
}
function resumeControle(items, libelles) {
  const par = v => Object.entries(items || {}).filter(([, x]) => x === v).map(([k]) => (libelles || {})[k] || k);
  return { c: par('C').length, nc: par('NC'), na: par('NA'), nf: par('NF') };
}

/* ---------- Ronde : parcours guidé, une étape à la fois ---------- */
const revueDraft = {};
function ancienneRonde(ronde) {   // ronde commencée avec une version précédente
  if (ronde.etapes) return;
  ronde.etapes = [{ id: 'comptage', lbl: 'Comptage des véhicules' }, { id: 'quotidien', lbl: 'Contrôle quotidien' }, { id: 'fin', lbl: 'Fin de ronde' }];
  ronde.step = ronde.compteValide ? 'quotidien' : 'comptage';
  ronde.pos = {}; ronde.revus = []; ronde.reporte = {}; ronde.lance = {};
}
function enteteRonde(ronde) {
  const et = ronde.etapes, i = Math.max(0, et.findIndex(e => e.id === ronde.step));
  return [bandeauRondier(ronde),
    h('div', { class: 'etapes' },
      h('div', { class: 'etapes-txt' }, `Étape ${i + 1} sur ${et.length} : `, h('b', null, et[i].lbl)),
      h('div', { class: 'etapes-barre' }, ...et.map((e, k) => h('span', { class: k < i ? 'fait' : k === i ? 'en-cours' : '' }))))];
}
async function avancer(ronde, vers) {
  const et = ronde.etapes, i = et.findIndex(e => e.id === ronde.step);
  ronde.step = vers || (et[i + 1] || et[et.length - 1]).id;
  await kvSet('ronde', ronde);
  go('ronde', {}, { noPush: true });
}
VIEWS.revue = (p) => VIEWS.ronde(p);
VIEWS.controle = (p) => VIEWS.ronde(p);
VIEWS.ronde = async () => {
  const ronde = await kvGet('ronde');
  if (!ronde) return agentHome();
  ancienneRonde(ronde);
  if (ronde.step === 'revue') return etapeRevue(ronde);
  if (ronde.step === 'comptage') return etapeComptage(ronde);
  if (ronde.step === 'quotidien' || ronde.step === 'mensuel' || ronde.step === 'trimestriel') return etapeControle(ronde, ronde.step);
  return etapeFin(ronde);
};
const pageRonde = (ronde, ...contenu) => page(`Ronde : ${siteNom(ronde.site)}`, { back: () => go('home') }, ...enteteRonde(ronde), ...contenu);

async function etapeRevue(ronde) {
  const open = (await dbAll('sigs')).filter(s => s.site === ronde.site && s.statut === 'ouvert' && !ronde.sigsCrees.includes(s.id))
    .sort((a, b) => (a.jour || '').localeCompare(b.jour || ''));
  const vus = new Set(ronde.revus || []);
  const total = open.length + [...vus].filter(id => !open.some(s => s.id === id)).length;
  const s = open.find(x => !vus.has(x.id));
  if (!s) return avancer(ronde);
  const k = total - open.filter(x => !vus.has(x.id)).length + 1;
  const d = revueDraft[s.id] || (revueDraft[s.id] = { comment: '', photos: [], ouvert: false });
  const blocPhotos = await photoBlocks(s.photos || [], { rapide: true, grande: true, vignette: await thumbOf(s) });
  const thumbs = h('div', { class: 'thumbs' });
  const drawThumbs = () => thumbs.replaceChildren(...d.photos.map((p, i) =>
    h('div', { class: 't' }, h('img', { src: p.thumb, alt: '' }), h('button', { onclick: () => { d.photos.splice(i, 1); drawThumbs(); } }, '×'))));
  drawThumbs();
  const plus = h('div', { style: d.ouvert ? '' : 'display:none' },
    h('textarea', { placeholder: 'Commentaire (facultatif)', value: d.comment, oninput: e => d.comment = e.target.value }),
    thumbs,
    h('button', { class: 'sec', onclick: photoInput(async f => { showBusy('Photo…'); d.photos.push(await processPhoto(f)); hideBusy(); drawThumbs(); }) }, 'Ajouter une photo'));
  const repondre = verdict => safe(async () => {
    if (verdict === 'resolu' && !(await demander({ titre: 'Problème résolu ?', texte: 'Le signalement sera clos.', oui: 'Oui, résolu', non: 'Non' }))) return;
    showBusy('Enregistrement…');
    await savePhotos(d.photos);
    await appendEvent('revue', ronde, { sig: s.id, ref: s.ref, verdict, comment: d.comment.trim() },
      { photoAges: d.photos.map(p => ({ id: p.id, ageMin: p.ageMin })) }, d.photos);
    delete revueDraft[s.id];
    const r = await kvGet('ronde');
    r.revus = [...(r.revus || []), s.id]; r.revueFaite = (r.revueFaite || 0) + 1;
    await kvSet('ronde', r);
    await fold();
    hideBusy();
    go('ronde', {}, { noPush: true });
  });
  return pageRonde(ronde,
    h('p', { class: 'compte-etape' }, `Signalement ${k} sur ${total}`),
    h('div', { class: 'card' },
      h('h3', null, sigTitre(s), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('p', { class: 'muted small' }, `${s.ref} — signalé le ${fmtJour(s.jour)}${s.agent ? ' par ' + s.agent : ''}`),
      s.desc ? h('p', null, s.desc) : null,
      s.plaque || s.emplacement ? h('p', null, s.plaque ? [h('b', null, 'Plaque : '), s.plaque, ' '] : null, s.emplacement ? [h('b', null, 'Emplacement : '), s.emplacement] : null) : null,
      blocPhotos.length ? blocPhotos : h('p', { class: 'muted small' }, 'Pas de photo.'),
      s.notesAgents ? h('div', { class: 'banner info' }, 'DPMS : ' + s.notesAgents) : null,
      bandeauDemande(s),
      (s.messages || []).length ? h('button', { class: 'sec', onclick: () => go('echanges', { id: s.id }) }, icoCourrier(await nonLus(s), !!s.demande), 'Échanges avec le DPMS') : null,
      pmBanner(s)),
    h('p', { class: 'question' }, 'Aujourd’hui, ce problème est :'),
    h('button', { class: 'ok big rep', onclick: repondre('resolu') }, 'Résolu'),
    h('button', { class: 'sec big rep', onclick: repondre('encours') }, 'Toujours en cours'),
    h('button', { class: 'danger big rep', onclick: repondre('aggrave') }, 'Aggravé'),
    h('button', { class: 'link', onclick: () => { d.ouvert = !d.ouvert; plus.style.display = d.ouvert ? '' : 'none'; } }, 'Ajouter un commentaire ou une photo'),
    plus);
}

async function etapeComptage(ronde) {
  if (ronde.compte == null) ronde.compte = 0;
  const nb = h('div', { class: 'compteur' }, String(ronde.compte));
  const lbl = h('span');
  const maj = async v => { ronde.compte = Math.max(0, v); nb.textContent = ronde.compte; lbl.textContent = `Valider : ${ronde.compte} véhicule${ronde.compte > 1 ? 's' : ''}`; await kvSet('ronde', ronde); };
  lbl.textContent = `Valider : ${ronde.compte} véhicule${ronde.compte > 1 ? 's' : ''}`;
  const saisie = h('input', { type: 'number', min: '0', inputmode: 'numeric', value: String(ronde.compte) });
  const corr = h('div', { style: 'display:none' }, h('label', { class: 'f' }, 'Nombre de véhicules'), saisie,
    h('button', { class: 'sec', onclick: safe(async () => { if (/^\d+$/.test(saisie.value.trim())) { await maj(parseInt(saisie.value, 10)); corr.style.display = 'none'; } else toast('Nombre invalide.', true); }) }, 'OK'));
  return pageRonde(ronde,
    h('p', { class: 'question' }, 'Comptez les véhicules garés : touchez « +1 » pour chaque véhicule.'),
    h('div', { class: 'card' }, nb,
      h('button', { class: 'plus1', onclick: safe(async () => { await maj(ronde.compte + 1); if (navigator.vibrate) navigator.vibrate(30); }) }, '+1 véhicule'),
      h('div', { class: 'row' },
        h('button', { class: 'sec', onclick: safe(() => maj(ronde.compte - 1)) }, '−1'),
        h('button', { class: 'sec', onclick: () => { corr.style.display = ''; saisie.value = String(ronde.compte); saisie.focus(); } }, 'Corriger')),
      corr),
    h('button', { class: 'ok big', onclick: safe(async () => {
      if (!(await demander({ titre: 'Valider le comptage ?', texte: `${ronde.compte} véhicule${ronde.compte > 1 ? 's' : ''} compté${ronde.compte > 1 ? 's' : ''}.`, oui: 'Oui, valider', non: 'Non, corriger' }))) return;
      const r = await kvGet('ronde'); r.compte = ronde.compte; r.compteValide = true;
      await appendEvent('comptage', r, { total: r.compte });
      await avancer(r);
    }) }, lbl));
}

function prochainPoint(pts, etat, depuis) {
  for (let k = depuis; k < pts.length; k++) if (!etat(pts[k])) return k;
  for (let k = 0; k < Math.min(depuis, pts.length); k++) if (!etat(pts[k])) return k;
  return -1;
}
async function etapeControle(ronde, niveau) {
  const st = await etatControles();
  const pts = pointsSite(st, ronde.site, niveau);
  const sigsRonde = (await dbAll('sigs')).filter(s => ronde.sigsCrees.includes(s.id));
  ronde.ctl = ronde.ctl || {}; ronde.pos = ronde.pos || {}; ronde.lance = ronde.lance || {}; ronde.reporte = ronde.reporte || {};
  const rep = ronde.ctl[niveau] = ronde.ctl[niveau] || {};
  const etat = c => ncDe(sigsRonde, c.id).length ? 'NC' : rep[c.id];
  const periodique = niveau !== 'quotidien';

  if (periodique && (ronde.ctlEmis || {})[niveau]) return avancer(ronde);
  if (periodique && !ronde.lance[niveau]) {   // proposition du contrôle mensuel ou trimestriel
    const e = echeance(st, ronde.site, niveau);
    const fait = pts.filter(etat).length;
    return pageRonde(ronde,
      h('div', { class: 'card', style: 'border:2px solid var(--accent)' },
        h('h3', null, `Le ${NIVEAUX[niveau].lbl.toLowerCase()} est à faire dans ce parking`),
        h('p', null, e.dernier ? `Dernier : il y a ${e.depuis} jours (${e.dernier.agent}).` : 'Il n’a jamais été fait dans ce parking.'),
        h('p', null, `${pts.length} points à contrôler${fait ? ` (${fait} déjà faits)` : ''}.`)),
      h('button', { class: 'accent big', onclick: safe(async () => { ronde.lance[niveau] = true; ronde.reporte[niveau] = false; await kvSet('ronde', ronde); go('ronde', {}, { noPush: true }); }) }, fait ? 'Reprendre maintenant' : 'Faire maintenant'),
      h('button', { class: 'sec big', onclick: safe(async () => { ronde.reporte[niveau] = true; await avancer(ronde); }) }, 'Plus tard'));
  }

  // Liste déroulante des thèmes : l'agent ouvre le thème qu'il veut, dans l'ordre qu'il veut,
  // et peut revenir sur un thème déjà vu pour signaler à la volée.
  ronde.ouvert = ronde.ouvert || {};
  const faits = pts.filter(etat).length, reste = pts.length - faits;
  const suivantNonFait = depuis => { const k = prochainPoint(pts, c => c.id === depuis ? 'x' : etat(c), Math.max(0, pts.findIndex(c => c.id === depuis))); return k < 0 ? null : pts[k].id; };
  if (ronde.ouvert[niveau] === undefined) ronde.ouvert[niveau] = reste ? pts[prochainPoint(pts, etat, 0)].id : null;
  const ouvert = ronde.ouvert[niveau];
  const ouvrir = async (id, defiler = true) => { const r = await kvGet('ronde'); r.ouvert = r.ouvert || {}; r.ouvert[niveau] = id; await kvSet('ronde', r); _defilerVers = defiler && id ? 'pt-' + id : null; go('ronde', {}, { noPush: true }); };
  const repondre = (c, val) => safe(async () => { rep[c.id] = val; ronde.ouvert[niveau] = suivantNonFait(c.id); await kvSet('ronde', ronde); _defilerVers = ronde.ouvert[niveau] ? 'pt-' + ronde.ouvert[niveau] : 'fin-controle'; go('ronde', {}, { noPush: true }); });
  const signaler = c => safe(async () => { ronde.ouvert[niveau] = suivantNonFait(c.id); await kvSet('ronde', ronde); _defilerVers = 'pt-' + c.id; go('signalement', { from: 'ronde', cat: catPour(c), ctl: { id: c.id, niveau, lbl: c.lbl } }); });
  const carte = c => {
    const v = etat(c), nc = ncDe(sigsRonde, c.id), ouv = ouvert === c.id;
    const cls = v === 'NC' ? 'nc' : v === 'C' ? 'c' : v === 'NA' ? 'na' : 'afaire';
    const badge = v === 'NC' ? `Non conforme${nc.length > 1 ? ` (${nc.length})` : ''}` : v === 'C' ? 'Conforme' : v === 'NA' ? 'N’existe pas ici' : 'À vérifier';
    return h('div', { class: `theme ${cls}${ouv ? ' ouvert' : ''}`, id: 'pt-' + c.id },
      h('button', { class: 'theme-tete', onclick: () => ouvrir(ouv ? null : c.id, !ouv) },
        h('span', { class: 'theme-pastille' }, v === 'NC' ? '!' : v === 'C' ? '✓' : v === 'NA' ? '–' : ''),
        h('span', { class: 'theme-lbl' }, c.lbl),
        h('span', { class: 'theme-badge' }, badge),
        h('span', { class: 'theme-chevron' }, ouv ? '▲' : '▼')),
      ouv ? h('div', { class: 'theme-corps' },
        c.detail ? h('p', { class: 'point-detail' }, c.detail) : null,
        nc.length ? h('div', { class: 'banner bad' }, `${nc.length} signalement${nc.length > 1 ? 's' : ''} : ${nc.map(x => x.sub || x.cat).join(', ')}`) : null,
        nc.length ? null : h('button', { class: 'ok big' + (v === 'C' ? ' choisi' : ''), onclick: repondre(c, 'C') }, v === 'C' ? '✓ Conforme' : 'Conforme'),
        h('button', { class: 'danger big', onclick: signaler(c) }, nc.length ? 'Signaler un autre problème ici' : 'Non conforme : signaler'),
        c.opt && !nc.length ? h('button', { class: 'link', onclick: safe(async () => {
          if (!(await demander({ titre: 'Équipement absent ?', texte: `« ${c.lbl} » n’existe pas dans ce parking. Ce point ne sera plus proposé ici.`, oui: 'Oui, il n’existe pas ici', non: 'Non' }))) return;
          await repondre(c, 'NA')();
        }) }, 'N’existe pas dans ce parking') : null) : null);
  };
  const valider = safe(async () => {
    if (reste) { const id = pts[prochainPoint(pts, etat, 0)].id; toast(`Encore ${reste} thème${reste > 1 ? 's' : ''} à vérifier.`, true); return ouvrir(id); }
    if (periodique) { showBusy('Enregistrement…'); await emettreControle(ronde, niveau, pts, sigsRonde); await fold(); hideBusy(); toast(`${NIVEAUX[niveau].lbl} enregistré.`); }
    delete ronde.ouvert[niveau];
    return avancer(ronde);
  });
  if (_defilerVers) { const cible = _defilerVers; _defilerVers = null; setTimeout(() => { const el = document.getElementById(cible); if (el) el.scrollIntoView({ block: 'start', behavior: 'smooth' }); }, 60); }
  return pageRonde(ronde,
    h('p', { class: 'compte-etape' }, `${faits} sur ${pts.length} vérifié${faits > 1 ? 's' : ''}`, reste ? h('span', { class: 'muted small' }, ` · reste ${reste}`) : h('span', { class: 'muted small' }, ' · tout est vérifié')),
    h('p', { class: 'muted small', style: 'margin:0 0 8px' }, 'Touchez un thème pour l’ouvrir. Vous pouvez y revenir à tout moment pour signaler un problème.'),
    h('div', { class: 'themes' }, ...pts.map(carte)),
    h('button', { class: (reste ? 'sec' : 'ok') + ' big', id: 'fin-controle', style: 'margin-top:14px', onclick: valider },
      reste ? `Encore ${reste} thème${reste > 1 ? 's' : ''} à vérifier` : (periodique ? `Valider le ${NIVEAUX[niveau].lbl.toLowerCase()}` : 'Contrôle terminé : étape suivante')),
    h('button', { class: 'sec', onclick: () => go('signalement', { from: 'ronde' }) }, '+ Autre signalement (hors liste)'),
    periodique ? h('button', { class: 'link', onclick: safe(async () => { ronde.lance[niveau] = false; ronde.reporte[niveau] = true; await avancer(ronde); }) }, 'Finir plus tard') : null);
}
let _defilerVers = null;
async function etapeFin(ronde) {
  const st = await etatControles();
  const sigsRonde = (await dbAll('sigs')).filter(s => ronde.sigsCrees.includes(s.id));
  const ptsQ = pointsSite(st, ronde.site, 'quotidien');
  const repQ = (ronde.ctl || {}).quotidien || {};
  const ncQ = ptsQ.filter(c => ncDe(sigsRonde, c.id).length).length, cQ = ptsQ.filter(c => !ncDe(sigsRonde, c.id).length && repQ[c.id] === 'C').length;
  const resteQ = restantControle(ronde, 'quotidien', ptsQ, sigsRonde);
  const libres = sigsRonde.filter(s => !s.controle);
  const periodiques = ['mensuel', 'trimestriel'].filter(n => ronde.etapes.some(e => e.id === n));
  const ligne = (t, v, btn) => h('div', { class: 'stat' }, h('span', null, t), h('span', { style: 'display:flex;gap:8px;align-items:center' }, h('b', null, v), btn || null));
  const revenir = vers => h('button', { class: 'link mini', style: 'width:auto;margin:0;min-height:36px', onclick: safe(async () => { if (vers === 'comptage') ronde.compteValide = false; else if (vers !== 'quotidien') { ronde.lance[vers] = true; ronde.reporte[vers] = false; } await avancer(ronde, vers); }) }, vers === 'comptage' ? 'modifier' : vers === 'quotidien' ? 'revoir' : 'reprendre');
  return pageRonde(ronde,
    h('p', { class: 'question' }, 'Autre chose à signaler dans ce parking ?'),
    h('div', { class: 'row' },
      h('button', { class: 'sec big', onclick: () => go('signalement', { from: 'ronde' }) }, '+ Signalement'),
      h('button', { class: 'sec big', onclick: () => go('barriere', { site: ronde.site, from: 'ronde' }) }, 'Barrière ouverte')),
    h('div', { class: 'card' }, h('h2', null, 'Récapitulatif'),
      ligne('Véhicules comptés', ronde.compteValide ? String(ronde.compte) : 'non validé', revenir('comptage')),
      ligne('Contrôle quotidien', resteQ ? `${resteQ} thème(s) restant(s)` : `${cQ} conforme(s), ${ncQ} non conforme(s)`, revenir('quotidien')),
      ...periodiques.map(n => ligne(NIVEAUX[n].lbl, (ronde.ctlEmis || {})[n] ? 'fait' : 'reporté', (ronde.ctlEmis || {})[n] ? null : revenir(n))),
      ligne('Autres signalements', String(libres.length)),
      libres.length ? h('p', { class: 'muted small' }, libres.map(s => sigTitre(s)).join(' · ')) : null),
    h('button', { class: 'ok big', onclick: safe(async () => {
      if (!ronde.compteValide) return avancer(ronde, 'comptage');
      if (resteQ) { toast(`Contrôle quotidien : encore ${resteQ} thème(s) à vérifier.`, true); return avancer(ronde, 'quotidien'); }
      await finishRonde(ronde);
    }) }, 'Terminer la ronde'));
}
VIEWS.rondeFinie = async ({ site, agent }) => {
  const enAttente = (await dbAll('events')).length;
  const aTrans = await sigsATransmettre();
  return page('Ronde terminée', null,
    h('div', { class: 'fini' }, h('span', { html: ICO.ok }), h('h2', null, 'Ronde terminée'), h('p', null, `${siteNom(site)} — ${agent}`),
      h('p', { class: enAttente ? 'muted' : '' }, enAttente ? (navigator.onLine ? 'Envoi au DPMS en cours…' : 'Pas de réseau : la ronde partira automatiquement dès qu’il y en aura.') : 'Ronde envoyée au DPMS.')),
    aTrans.length ? h('div', { class: 'card a-transmettre' }, h('div', { class: 'at-titre' }, `Maintenant : ${aTrans.length} signalement${aTrans.length > 1 ? 's' : ''} à transmettre`),
      h('p', { class: 'small', style: 'margin:4px 0 8px' }, 'Appelez ou écrivez aux services concernés, puis indiquez-le dans l’application.'),
      h('button', { class: 'danger big', onclick: () => go('transmettre', { id: aTrans[0].id }, { noPush: true }) }, 'Transmettre maintenant')) : null,
    h('button', { class: (aTrans.length ? 'sec' : '') + ' big', onclick: () => go('home') }, aTrans.length ? 'Plus tard (retour à l’accueil)' : 'Retour à l’accueil'));
};
async function annulerRonde() {
  const ronde = await kvGet('ronde');
  if (!ronde) { go('home'); return; }
  const gardes = [];
  const n = ronde.sigsCrees.length, r = ronde.revueFaite;
  if (n) gardes.push(n > 1 ? `les ${n} signalements faits pendant la ronde` : 'le signalement fait pendant la ronde');
  ['mensuel', 'trimestriel'].forEach(n => { if ((ronde.ctlEmis || {})[n]) gardes.push(`le ${NIVEAUX[n].lbl.toLowerCase()}`); });
  if (r) gardes.push(r > 1 ? `l’état des ${r} signalements en cours` : 'l’état du signalement en cours');
  if (!(await demander({ titre: 'Annuler la ronde ?', danger: true, oui: 'Oui, annuler la ronde', non: 'Non, continuer',
    texte: `${siteNom(ronde.site)}, ${ronde.agent}.\nElle n’apparaîtra pas dans les rondes. Le comptage et les contrôles non validés sont effacés.`
      + (gardes.length ? `\n\nDéjà enregistré et conservé : ${gardes.join(', ')}.` : '') }))) return;
  showBusy('Annulation…');
  await appendEvent('ronde_annulee', ronde, { nbSig: ronde.sigsCrees.length });
  await kvDel('ronde');
  await kvDel('dernierAgent');
  Object.keys(revueDraft).forEach(k => delete revueDraft[k]);
  Geo.stop();
  hideBusy();
  toast('Ronde annulée.');
  go('home');
  syncNow().then(() => refreshIfHome());
}
const lienAnnuler = () => h('button', { class: 'link', style: 'color:var(--bad)', onclick: safe(annulerRonde) }, 'Annuler la ronde');

async function finishRonde(ronde, silent) {
  showBusy('Clôture de la ronde…');
  const st = await etatControles();
  const sigsRonde = (await dbAll('sigs')).filter(s => ronde.sigsCrees.includes(s.id));
  if (ronde.quotidienDu && !(ronde.ctlEmis || {}).quotidien) await emettreControle(ronde, 'quotidien', pointsSite(st, ronde.site, 'quotidien'), sigsRonde);
  if (!ronde.compteValide) await appendEvent('comptage', ronde, { total: ronde.compte || 0, nonValide: true });
  await appendEvent('ronde_fin', ronde, { nbSig: ronde.sigsCrees.length, compte: ronde.compte || 0 });
  await kvDel('ronde');
  Geo.stop();
  hideBusy();
  if (!silent) {
    go('rondeFinie', { site: ronde.site, agent: ronde.agent }, { noPush: true });
    syncNow().then(() => { if (state.view === 'rondeFinie') render(); refreshIfHome(); });
  }
}

/* ---------- Signalement (en ronde ou hors ronde) ---------- */
let sigDraft = null;
/* Signalement en étapes : (parking, agent) → catégorie → « Quoi ? » → photo → précisions → destinataires et envoi */
VIEWS.signalement = async ({ cat, from, ctl, mc }) => {
  const ronde = from === 'ronde' ? await kvGet('ronde') : null;
  if (from === 'ronde' && !ronde) return agentHome();
  const cfg = state.cfg;
  const cle = `${from}|${cat || ''}|${ctl ? ctl.id : ''}|${mc ? mc.id : ''}`;
  if (!sigDraft || sigDraft.cle !== cle) {
    sigDraft = { cle, site: ronde ? ronde.site : (mc && mc.site) || null, agent: ronde ? ronde.agent : (mc && mc.agent) || await agentPrefere(),
      cat: cat || null, sub: null, desc: (mc && mc.desc) || '', plaque: '', emplacement: '', vehicule: '', photos: [], dests: [], urgent: false, destsTouche: false, etape: null };
    if (cat) appliquerCat(sigDraft, cat);
  }
  if (from !== 'ronde') Geo.start();
  const d = sigDraft;
  const point = ctl ? (cfg.controles || []).find(x => x.id === ctl.id) : null;
  const subsPoint = () => point && (point.subs || []).length && d.cat === catPour(point) ? point.subs : null;
  const listeQuoi = () => { const c = catByNom(d.cat); return subsPoint() || (c ? c.subs : null); };
  const descOblig = () => d.sub === SUB_AUTRE || (d.cat === CAT_AUTRE && !(subsPoint() && d.sub));
  // Photo obligatoire, sauf quand une personne est en cause (sécurité de l'agent, image des personnes)
  const photoFacultative = () => d.cat === 'Sûreté et présences' || /personne|comportement suspect|vandalisme en cours/i.test(d.sub || '');
  const etapes = () => [
    ...(ronde ? [] : [...(mc && mc.site ? [] : ['parking']), ...(mc && mc.agent ? [] : ['agent'])]),
    ...(ctl ? [] : ['cat']),
    ...(listeQuoi() ? ['quoi'] : []),
    'photo', 'details', 'envoi'];
  if (!d.etape || !etapes().includes(d.etape)) d.etape = etapes()[0];
  const ets = etapes(), k = ets.indexOf(d.etape);
  const aller = e => { d.etape = e; go('signalement', { cat, from, ctl, mc }, { noPush: true }); };
  const suivante = () => aller(etapes()[etapes().indexOf(d.etape) + 1]);
  const retour = () => { sigDraft = null; if (from === 'ronde') go('ronde', {}, { noPush: true }); else { geoStopSiLibre(); go('home'); } };
  const back = async () => {
    if (k > 0) return aller(ets[k - 1]);
    if (d.photos.length || d.desc) { if (!(await demander({ titre: 'Abandonner ce signalement ?', oui: 'Oui, abandonner', non: 'Non, continuer', danger: true }))) return; }
    retour();
  };
  const choixListe = (opts, sel, fn) => opts.map(o => h('button', { class: 'choix' + ((typeof o === 'object' ? o.val : o) === sel ? ' on' : ''), onclick: () => fn(typeof o === 'object' ? o.val : o) }, typeof o === 'object' ? o.lbl : o));
  const appliquerQuoi = v => {
    d.sub = v;
    const c = catByNom(d.cat), liste = listeQuoi() || [];
    const sp = liste.find(x => x.lbl === v), sc = c && c.subs.find(x => x.lbl === v);
    if (sp && sp.urgent) { d.urgent = true; d.urgentAuto = true; } else if (d.urgentAuto) { d.urgent = false; d.urgentAuto = false; }
    if (!d.destsTouche) {   // destinataires propres à la sous-catégorie, sinon ceux de la catégorie
      const ds = ((sc && sc.dests && sc.dests.length) ? sc.dests : (c && c.dests) || []).filter(x => cfg.dests.includes(x)).filter(destPourSite(d.site));
      d.dests.splice(0, d.dests.length, ...ds);
    }
  };

  let corps;
  if (d.etape === 'parking') {
    corps = [h('p', { class: 'question' }, 'Dans quel parking ?'), ...choixListe(cfg.sites.map(x => ({ val: x.id, lbl: x.nom })), d.site, v => { if (v !== d.site) { d.site = v; if (!d.destsTouche) d.dests.splice(0, d.dests.length, ...d.dests.filter(destPourSite(v))); } suivante(); })];
  } else if (d.etape === 'agent') {
    corps = [h('p', { class: 'question' }, 'Qui fait le signalement ?'), ...choixListe(cfg.agents, d.agent, v => { d.agent = v; suivante(); })];
  } else if (d.etape === 'cat') {
    corps = [h('p', { class: 'question' }, 'De quoi s’agit-il ?'), ...choixListe([...cfg.cats.map(c => c.nom), CAT_AUTRE], d.cat, v => { if (v !== d.cat) appliquerCat(d, v); suivante(); })];
  } else if (d.etape === 'quoi') {
    corps = [h('p', { class: 'question' }, 'Qu’est-ce qui ne va pas ?'), ...choixListe([...listeQuoi().map(x => x.lbl), SUB_AUTRE], d.sub, v => { appliquerQuoi(v); suivante(); })];
  } else if (d.etape === 'photo') {
    const thumbs = h('div', { class: 'thumbs grandes' }, ...d.photos.map((p, n) =>
      h('div', { class: 't' }, h('img', { src: p.thumb, alt: '' }), h('button', { onclick: () => { d.photos.splice(n, 1); aller('photo'); } }, '×'))));
    corps = [h('p', { class: 'question' }, d.photos.length ? `${d.photos.length} photo${d.photos.length > 1 ? 's' : ''}` : photoFacultative() ? 'Prenez une photo si vous le pouvez sans danger.' : 'Prenez une photo du problème (obligatoire).'),
      thumbs,
      d.photos.length < 4 ? h('button', { class: (d.photos.length ? 'sec' : 'accent') + ' big', onclick: photoInput(async f => { showBusy('Photo…'); d.photos.push(await processPhoto(f)); hideBusy(); aller('photo'); }) }, d.photos.length ? 'Ajouter une autre photo' : 'Prendre une photo') : null,
      d.photos.length ? h('button', { class: 'ok big', onclick: suivante }, 'Suivant')
        : photoFacultative() ? h('button', { class: 'sec', onclick: suivante }, 'Pas de photo possible (personne présente, danger)') : null];
  } else if (d.etape === 'details') {
    const veh = isVehicule(d.cat);
    corps = [h('p', { class: 'question' }, descOblig() ? 'Décrivez le problème (obligatoire)' : 'Précisions (facultatif)'),
      h('textarea', { placeholder: 'Où exactement (niveau, place, porte…), ce qui ne va pas. Le micro du clavier permet de dicter.', oninput: e => d.desc = e.target.value, value: d.desc }),
      veh ? [
        h('label', { class: 'f' }, 'Plaque d’immatriculation'),
        h('input', { type: 'text', value: d.plaque, placeholder: 'AB-123-CD', autocapitalize: 'characters', oninput: e => d.plaque = e.target.value }),
        h('label', { class: 'f' }, 'Emplacement (numéro de place, niveau)'),
        h('input', { type: 'text', value: d.emplacement, placeholder: 'Ex. : place 112, niveau -1', oninput: e => d.emplacement = e.target.value }),
        h('label', { class: 'f' }, 'Véhicule (marque, modèle, couleur)'),
        h('input', { type: 'text', value: d.vehicule, placeholder: 'Ex. : Renault Clio grise', oninput: e => d.vehicule = e.target.value })] : null,
      h('button', { class: 'ok big', onclick: () => { if (descOblig() && !d.desc.trim()) { toast('Décrivez le problème.', true); return; } suivante(); } }, 'Suivant')];
  } else {   // envoi
    const cases = h('div');
    const drawCases = () => cases.replaceChildren(...cfg.dests.filter(destPourSite(d.site)).map(x => {
      const on = d.dests.includes(x);
      return h('button', { class: 'case' + (on ? ' on' : ''), onclick: () => { const n = d.dests.indexOf(x); if (n >= 0) d.dests.splice(n, 1); else d.dests.push(x); d.destsTouche = true; drawCases(); } },
        h('span', { class: 'coche' }, on ? '✓' : ''), x);
    }));
    drawCases();
    const urg = h('button', { class: 'case urgent' + (d.urgent ? ' on' : ''), onclick: () => { d.urgent = !d.urgent; d.urgentAuto = false; urg.className = 'case urgent' + (d.urgent ? ' on' : ''); urg.firstChild.textContent = d.urgent ? '✓' : ''; } },
      h('span', { class: 'coche' }, d.urgent ? '✓' : ''), 'Urgent');
    corps = [
      h('div', { class: 'card resume' },
        h('b', null, d.cat === CAT_AUTRE && !d.sub ? 'Autre' : `${d.cat}${d.sub ? ' — ' + d.sub : ''}`),
        d.desc ? h('div', null, d.desc) : null,
        d.plaque ? h('div', null, 'Plaque : ', h('span', { class: 'plaque' }, d.plaque.toUpperCase())) : null,
        h('div', { class: 'muted small' }, `${siteNom(d.site)} — ${d.agent} — ${d.photos.length} photo${d.photos.length > 1 ? 's' : ''}`)),
      h('p', { class: 'question' }, 'À signaler à :'), h('p', { class: 'muted small', style: 'margin-top:-6px' }, 'Un ou plusieurs : touchez pour cocher ou décocher.'),
      cases,
      h('p', { class: 'question' }, 'Priorité :'), urg,
      h('button', { class: 'ok big', style: 'margin-top:16px', onclick: save(false) }, 'Enregistrer le signalement'),
      h('button', { class: 'sec', onclick: save(true) }, ctl ? 'Enregistrer et signaler une autre non-conformité ici' : 'Enregistrer et faire un autre signalement')];
  }

  function save(encore) {
    return safe(async () => {
      if (!d.site) throw new Error('Choisissez le parking.');
      if (!d.agent) throw new Error('Indiquez qui fait le signalement.');
      if (!d.cat) throw new Error('Choisissez une catégorie.');
      if ((d.cat !== CAT_AUTRE || subsPoint()) && !d.sub) throw new Error('Choisissez ce qui ne va pas (« Quoi ? »).');
      if (descOblig() && !d.desc.trim()) throw new Error('Décrivez le problème.');
      if (!d.photos.length && !photoFacultative()) { aller('photo'); throw new Error('Une photo est obligatoire.'); }
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
          dests: [...d.dests], urgent: d.urgent, horsRonde: !ronde, thumb: d.photos[0] ? d.photos[0].thumb : null, controle: ctl || null, mc: mc ? mc.id : null
        },
        { photoAges: d.photos.map(p => ({ id: p.id, ageMin: p.ageMin })) }, d.photos);
      if (ronde) { const r = await kvGet('ronde'); r.sigsCrees.push(id); await kvSet('ronde', r); }
      else { await kvSet('dernierAgent', { nom: d.agent, at: Date.now() }); state.dernierSite = d.site; }
      await fold();
      const urgent = d.urgent, catNom = d.cat, siteId = d.site, agent = d.agent;
      sigDraft = null;
      hideBusy();
      if (urgent || !ronde) syncNow().then(refreshIfHome);   // urgent ou hors ronde : envoi immédiat
      if (encore) {
        toast(`Signalement ${ref} enregistré.`);
        if (!ronde) { sigDraft = { cle, site: siteId, agent, cat: null, sub: null, desc: '', plaque: '', emplacement: '', vehicule: '', photos: [], dests: [], urgent: false, destsTouche: false, etape: 'cat' }; }
        go('signalement', { from, cat: ctl ? catNom : null, ctl, mc }, { noPush: true });
      } else if (ronde) { toast(`Signalement ${ref} enregistré.`); go('ronde', {}, { noPush: true }); }
      else go('sigEnregistre', { ref }, { noPush: true });
    });
  }

  return page(ctl ? 'Non conforme' : ronde ? 'Nouveau signalement' : 'Signalement hors ronde', { back },
    ronde ? bandeauRondier(ronde, false) : null,
    ctl ? h('div', { class: 'banner warn' }, `${NIVEAUX[ctl.niveau].lbl} — ${ctl.lbl}`) : null,
    h('div', { class: 'etapes' }, h('div', { class: 'etapes-txt' }, `Étape ${k + 1} sur ${ets.length}`),
      h('div', { class: 'etapes-barre' }, ...ets.map((e, n) => h('span', { class: n < k ? 'fait' : n === k ? 'en-cours' : '' })))),
    ...corps,
    k > 0 ? h('button', { class: 'link', onclick: back }, '◀ Étape précédente') : null);
};
VIEWS.sigEnregistre = async ({ ref }) => {
  const s = (await dbAll('sigs')).find(x => x.ref === ref);
  const reste = s ? destsATransmettre(s) : [];
  const tel = dn => ((state.cfg.contacts || {})[dn] || {}).tel;
  return page('Signalement enregistré', null,
  h('div', { class: 'fini' }, h('span', { html: ICO.ok }), h('h2', null, 'Signalement enregistré'), h('p', null, `Référence ${ref}`),
    h('p', { class: 'muted' }, navigator.onLine ? 'Le DPMS le reçoit par l’application.' : 'Pas de réseau : il partira automatiquement dès qu’il y en aura.')),
  reste.length ? h('div', { class: 'card a-transmettre' },
    h('div', { class: 'at-titre' }, 'Maintenant : prévenir le service'),
    h('p', { style: 'margin:4px 0 8px' }, reste.map(dn => tel(dn) ? `${dn} (${tel(dn)})` : dn).join(' · ')),
    h('p', { class: 'small', style: 'margin:0 0 8px' }, 'Appelez ou envoyez un courriel, puis indiquez-le dans l’application.'),
    h('button', { class: 'danger big', onclick: () => go('transmettre', { id: s.id }, { noPush: true }) }, 'Transmettre maintenant'),
    h('button', { class: 'link', onclick: () => go('home') }, 'Plus tard : il reste en rouge sur l’accueil')) : h('button', { class: 'big', onclick: () => go('home') }, 'Retour à l’accueil'),
  h('button', { class: 'sec', onclick: () => { sigDraft = null; go('signalement', { from: 'hr' }, { noPush: true }); } }, 'Faire un autre signalement'));
};
/* Destinataire propre à un parking : « Designa (Dumotel) » n'est proposé que pour les signalements de Dumotel */
function destSite(nom) {
  const m = /\(([^)]+)\)\s*$/.exec(nom || ''); if (!m) return null;
  const s = (state.cfg.sites || []).find(x => sansAccent(x.nom) === sansAccent(m[1]));
  return s ? s.id : null;
}
const destPourSite = site => x => { const ds = destSite(x); return !ds || !site || ds === site; };
function appliquerCat(d, nom) {
  d.cat = nom; d.sub = null;
  const c = catByNom(nom);
  if (!d.destsTouche) { d.dests.splice(0, d.dests.length, ...((c && c.dests) || []).filter(x => state.cfg.dests.includes(x)).filter(destPourSite(d.site))); }
}

/* ---------- Barrière ouverte / refermée ---------- */
let barrDraft = null;
/* Barrière ouverte / refermée, en étapes : (parking) → barrière → motif → depuis quand → (qui) → enregistrer */
VIEWS.barriere = async ({ site, bid, from, agent }) => {
  const cfg = state.cfg;
  const toutes = (await kvGet('barrieres')) || [];
  const b = bid ? toutes.find(x => x.bid === bid) : null;
  if (bid && (!b || b.fermee)) { toast('Cette barrière est déjà signalée refermée.'); return agentHome(); }
  const ronde = from === 'ronde' ? await kvGet('ronde') : null;
  const cle = (bid || 'new|' + (site || '')) + '|' + (agent || '');
  if (!barrDraft || barrDraft.cle !== cle) {
    barrDraft = { cle, site: site || (b && b.site) || null, barriere: null, motif: null, comment: '', agent: ronde ? ronde.agent : agent || await agentPrefere(), maintenant: null, quand: localDT(), etape: null };
  }
  Geo.start();
  const d = barrDraft;
  const etapes = () => b ? ['quand', ...(ronde ? [] : ['agent']), 'ok']
    : [...(site ? [] : ['parking']), 'barriere', 'motif', 'quand', ...(ronde || agent ? [] : ['agent']), 'ok'];
  if (!d.etape || !etapes().includes(d.etape)) d.etape = etapes()[0];
  const ets = etapes(), k = ets.indexOf(d.etape);
  const aller = e => { d.etape = e; go('barriere', { site, bid, from, agent }, { noPush: true }); };
  const suivante = () => aller(ets[k + 1]);
  const quitter = () => { barrDraft = null; if (from === 'ronde') go('ronde', {}, { noPush: true }); else { geoStopSiLibre(); go('home'); } };
  const back = () => k > 0 ? aller(ets[k - 1]) : quitter();
  const choixListe = (opts, sel, fn) => opts.map(o => h('button', { class: 'choix' + ((typeof o === 'object' ? o.val : o) === sel ? ' on' : ''), onclick: () => fn(typeof o === 'object' ? o.val : o) }, typeof o === 'object' ? o.lbl : o));

  const save = safe(async () => {
    const quand = d.maintenant ? new Date() : new Date(d.quand);
    if (isNaN(quand)) throw new Error('Heure invalide.');
    if (quand.getTime() > Date.now() + 120000) throw new Error('L’heure ne peut pas être dans le futur.');
    if (quand.getTime() < Date.now() - 7 * 86400000) throw new Error('Au-delà de 7 jours, prévenez le DPMS.');
    if (!d.agent) throw new Error('Indiquez qui fait la déclaration.');
    if (b) {
      if (quand < new Date(b.ouverte.quand)) throw new Error(`La fermeture ne peut pas précéder l’ouverture (${fmtDT(b.ouverte.quand)}).`);
      await appendEvent('barriere', { site: b.site, agent: d.agent }, { bid: b.bid, action: 'fermee', barriere: b.barriere, quand: quand.toISOString(), comment: d.comment.trim(), retro: !d.maintenant });
    } else {
      if (!d.site) throw new Error('Choisissez le parking.');
      if (!d.barriere) throw new Error('Choisissez la barrière.');
      if (toutes.some(x => x.site === d.site && x.barriere === d.barriere && !x.fermee)) throw new Error('Cette barrière est déjà signalée ouverte.');
      if (!d.motif) throw new Error('Indiquez pourquoi elle est ouverte.');
      await appendEvent('barriere', { site: d.site, agent: d.agent }, { bid: uuid(), action: 'ouverte', barriere: d.barriere, quand: quand.toISOString(), motif: d.motif, comment: d.comment.trim(), retro: !d.maintenant });
    }
    await kvSet('dernierAgent', { nom: d.agent, at: Date.now() });
    await fold();
    const msg = b ? `Barrière ${b.barriere} (${siteNom(b.site)}) : fermeture enregistrée.` : `Barrière ${d.barriere} (${siteNom(d.site)}) : ouverture enregistrée.`;
    barrDraft = null;
    syncNow().then(refreshIfHome);
    if (from === 'ronde') { toast(msg); go('ronde', {}, { noPush: true }); }
    else go('barriereOk', { msg, ouverte: !b }, { noPush: true });
  });

  let corps;
  if (d.etape === 'parking') {
    corps = [h('p', { class: 'question' }, 'Dans quel parking ?'), ...choixListe(cfg.sites.map(x => ({ val: x.id, lbl: x.nom })), d.site, v => { if (v !== d.site) d.barriere = null; d.site = v; suivante(); })];
  } else if (d.etape === 'barriere') {
    const s = siteById(d.site || site);
    const ouv = toutes.filter(x => x.site === s.id && !x.fermee).map(x => x.barriere);
    corps = [h('p', { class: 'question' }, 'Quelle barrière ?'),
      ...s.barrieres.map(x => ouv.includes(x)
        ? h('button', { class: 'choix', disabled: true }, `${x} (déjà signalée ouverte)`)
        : h('button', { class: 'choix' + (d.barriere === x ? ' on' : ''), onclick: () => { d.barriere = x; d.site = s.id; suivante(); } }, x))];
  } else if (d.etape === 'motif') {
    corps = [h('p', { class: 'question' }, 'Pourquoi est-elle ouverte ?'), ...choixListe(cfg.motifs, d.motif, v => { d.motif = v; suivante(); })];
  } else if (d.etape === 'quand') {
    const plusTot = h('div', { style: d.maintenant === false ? '' : 'display:none' },
      h('label', { class: 'f' }, b ? 'Refermée à' : 'Ouverte depuis'),
      h('input', { type: 'datetime-local', value: d.quand, max: localDT(new Date(Date.now() + 60000)), oninput: e => d.quand = e.target.value }),
      h('button', { class: 'ok big', onclick: suivante }, 'Suivant'));
    corps = [b ? h('div', { class: 'card' }, barriereLigne(b, false)) : null,
      h('p', { class: 'question' }, b ? 'Quand a-t-elle été refermée ?' : 'Depuis quand est-elle ouverte ?'),
      h('button', { class: 'choix' + (d.maintenant === true ? ' on' : ''), onclick: () => { d.maintenant = true; suivante(); } }, b ? 'Maintenant' : 'Maintenant (je viens de l’ouvrir)'),
      h('button', { class: 'choix' + (d.maintenant === false ? ' on' : ''), onclick: () => { d.maintenant = false; plusTot.style.display = ''; } }, 'Plus tôt (oubli de déclaration)'),
      plusTot];
  } else if (d.etape === 'agent') {
    corps = [h('p', { class: 'question' }, 'Qui fait la déclaration ?'), ...choixListe(cfg.agents, d.agent, v => { d.agent = v; suivante(); })];
  } else {
    const quand = d.maintenant ? 'maintenant' : fmtDT(new Date(d.quand));
    corps = [
      h('div', { class: 'card resume' },
        h('b', null, b ? `Barrière ${b.barriere} — ${siteNom(b.site)}` : `Barrière ${d.barriere} — ${siteNom(d.site)}`),
        h('div', null, b ? `Refermée : ${quand}` : `Ouverte : ${quand}`),
        !b ? h('div', null, `Motif : ${d.motif}`) : null,
        h('div', { class: 'muted small' }, `Déclarée par ${d.agent}`)),
      h('p', { class: 'question' }, 'Commentaire (facultatif)'),
      h('input', { type: 'text', value: d.comment, oninput: e => d.comment = e.target.value }),
      h('button', { class: (b ? 'ok' : 'danger') + ' big', style: 'margin-top:16px', onclick: save }, b ? 'Enregistrer la fermeture' : 'Enregistrer l’ouverture')];
  }
  return page(b ? 'Barrière refermée' : 'Barrière ouverte', { back },
    ronde ? bandeauRondier(ronde, false) : null,
    ets.length > 1 ? h('div', { class: 'etapes' }, h('div', { class: 'etapes-txt' }, `Étape ${k + 1} sur ${ets.length}`),
      h('div', { class: 'etapes-barre' }, ...ets.map((e, n) => h('span', { class: n < k ? 'fait' : n === k ? 'en-cours' : '' })))) : null,
    ...corps.filter(Boolean),
    k > 0 ? h('button', { class: 'link', onclick: back }, '◀ Étape précédente') : null);
};
VIEWS.barriereOk = async ({ msg, ouverte }) => page(ouverte ? 'Barrière ouverte' : 'Barrière refermée', null,
  h('div', { class: 'fini' }, h('span', { html: ICO.ok }), h('h2', null, ouverte ? 'Ouverture enregistrée' : 'Fermeture enregistrée'), h('p', null, msg),
    ouverte ? h('p', { class: 'muted' }, 'Elle reste affichée en rouge sur l’accueil jusqu’à « Barrière refermée ».') : null),
  h('button', { class: 'big', onclick: () => go('home') }, 'Retour à l’accueil'));

/* ====================================================================== */
/* MAIN COURANTE : contacts des usagers et demandes reçues                 */
/* (agents et DPMS ; heure du contact, réponse apportée, délai)           */
/* ====================================================================== */
const OBJETS_DEFAUT = ['Renseignement', 'Réclamation', 'Paiement ou ticket', 'Badge d’abonné', 'Véhicule bloqué', 'Objet perdu', 'Objet trouvé', 'Incivilité', 'Autre'];
const CANAUX = { place: 'Usager sur place', appel: 'Appel téléphonique reçu', courriel: 'Courriel reçu', service: 'Demande d’un service ou d’un élu', autre: 'Autre' };
const SUITES = { regle: 'Réponse donnée, demande réglée', barriere: 'Barrière ouverte pour l’usager', signalement: 'Problème à signaler (signalement)', service: 'Orienté vers un autre service', dpms: 'À traiter par le DPMS (réponse à apporter)' };
const SUITES_CLOSES = ['regle', 'barriere', 'service'];
const MODES_REP = { place: 'Sur place', telephone: 'Par téléphone', courriel: 'Par courriel', courrier: 'Par courrier' };
const MODE_CANAL = { place: 'place', appel: 'telephone', courriel: 'courriel', service: 'courriel', autre: 'place' };
const MAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
let mcDraft = null;
const lieuMC = c => c.site ? siteNom(c.site) : 'Général';
/* Reconstitution (appelée par fold) */
function mcAjout(mcs, c, par) {
  if (!c || !c.id || mcs.has(c.id)) return;
  const clos = SUITES_CLOSES.includes(c.suite);
  const rep = clos ? [{ quand: c.quand, mode: MODE_CANAL[c.canal] || 'place', texte: c.suite === 'service' ? `Orienté vers : ${c.service}${c.reponse ? ' — ' + c.reponse : ''}` : c.suite === 'barriere' ? `Barrière ouverte${c.reponse ? ' — ' + c.reponse : ''}` : c.reponse || 'Réponse donnée', par, regle: true }] : [];
  mcs.set(c.id, { ...c, par, statut: clos ? 'clos' : 'ouvert', closLe: clos ? c.quand : null, reponduLe: clos ? c.quand : null, hist: rep.map(r => ({ ...r, type: 'reponse' })), sigs: [], annule: false });
}
function mcSuivi(mcs, s, par, causal = { connu: () => true, marquer: () => { } }) {
  const c = s && mcs.get(s.mc);
  if (!c) return;
  const quand = s.quand || new Date().toISOString();
  const ok = causal.connu(c);
  c.hist.push({ type: s.type, quand, mode: s.mode || null, texte: s.texte || '', par: par || 'DPMS', regle: !!s.regle, sansEffet: ok ? undefined : true });
  if (s.type === 'reponse') { if (!c.reponduLe || quand < c.reponduLe) c.reponduLe = quand; if (s.regle && ok) { c.statut = 'clos'; c.closLe = quand; causal.marquer(c); } }
  if (s.type === 'clos' && ok) { c.statut = 'clos'; c.closLe = quand; causal.marquer(c); }
  if (s.type === 'rouvert' && ok) { c.statut = 'ouvert'; c.closLe = null; causal.marquer(c); }
  if (s.type === 'annule') c.annule = true;
}
async function mainCourante() { return ((await kvGet('mainCourante')) || []).filter(c => !c.annule); }
function delaiMC(c) { return c.reponduLe ? Math.max(0, Math.round((new Date(c.reponduLe) - new Date(c.quand)) / 60000)) : null; }
function badgeMC(c) {
  if (c.annule) return h('span', { class: 'badge' }, 'Annulée');
  if (c.statut === 'ouvert') { const j = joursDepuis(c.quand); return h('span', { class: 'badge b-bad' }, c.reponduLe ? 'Réponse d’attente' : `À traiter${j ? ` (${j} j)` : ''}`); }
  const dl = delaiMC(c);
  return h('span', { class: 'badge b-ok' }, `Réglée${dl ? ' en ' + dureeTxt(dl) : ''}`);
}
async function enregistrerSuiviMC(c, s, agent) {
  const quand = s.quand || new Date().toISOString();
  const data = { ...s, mc: c.id, ref: c.ref, nom: c.nom, quand };
  if (state.mode === 'superviseur') await queueDec('d', { type: 'contact_suivi', id: uuid(), ts: new Date().toISOString(), par: 'DPMS', suivi: data });
  else {
    if (!agent) throw new Error('Indiquez votre nom.');
    await appendEvent('contact_suivi', { site: c.site, agent }, data);
  }
  await fold(); syncNow();
}
/* Courriel ou appel à l'usager depuis la fiche : confirmation au retour dans l'application */
function courrielUsager(c) {
  const q = new Date(c.quand);
  const sujet = `Votre demande du ${fmtJour(localDate(q))} — ${c.objet} (réf. ${c.ref})`;
  const corps = `Bonjour,\n\nNous faisons suite à votre demande du ${fmtDT(c.quand)}${c.site ? ` concernant le parking ${siteNom(c.site)}` : ''} :\n« ${c.demande} »\n\n\n\nCordialement,\n\nDPMS — Ville de Cachan`;
  return `mailto:${encodeURIComponent(c.mail)}?subject=${encodeURIComponent(sujet)}&body=${encodeURIComponent(corps)}`;
}
async function lancerContactMC(c, type) {
  await kvSet('contactAttente', { mc: c.id, ref: c.ref, type, at: Date.now() });
  setTimeout(confirmerContact, 2500);
}
async function confirmerMC(att) {
  const c = ((await kvGet('mainCourante')) || []).find(x => x.id === att.mc);
  if (!c) { await kvDel('contactAttente'); return; }
  _confirmOuverte = true;
  const equipe = state.mode !== 'superviseur';
  const d = { par: equipe ? await agentPrefere() : null, texte: '' };
  const fermer = async () => { await kvDel('contactAttente'); box.remove(); _confirmOuverte = false; };
  const enreg = (type, regle, texteDefaut) => safe(async () => {
    if (equipe && !d.par) throw new Error('Indiquez qui a répondu.');
    await enregistrerSuiviMC(c, { type, mode: att.type === 'tel' ? 'telephone' : 'courriel', regle, texte: d.texte.trim() || texteDefaut }, d.par);
    await fermer();
    toast(type === 'reponse' ? (regle ? 'Réponse enregistrée, demande réglée.' : 'Réponse d’attente enregistrée.') : 'Appel sans réponse noté.');
    if (equipe && type === 'reponse') go('home'); else render();
  });
  const box = h('div', { class: 'confirm-contact' }, h('div', { class: 'sheet' },
    h('h3', null, att.type === 'tel' ? 'Appel' : 'Courriel', ` — ${c.ref}`),
    h('p', { class: 'muted small' }, `${c.nom} — ${c.objet}`),
    h('p', { style: 'font-weight:600' }, att.type === 'tel' ? `Avez-vous pu joindre ${c.nom} ?` : `Avez-vous envoyé la réponse à ${c.nom} ?`),
    equipe ? [h('label', { class: 'f' }, 'Qui ?'), chips(state.cfg.agents, d.par, v => d.par = v)] : null,
    h('textarea', { placeholder: 'Réponse apportée (résumé)', oninput: e => d.texte = e.target.value }),
    h('button', { class: 'ok', onclick: enreg('reponse', true, att.type === 'tel' ? 'Réponse donnée par téléphone' : 'Réponse envoyée par courriel') }, 'Oui, demande réglée'),
    h('button', { class: 'sec', onclick: enreg('reponse', false, att.type === 'tel' ? 'Réponse d’attente par téléphone' : 'Réponse d’attente par courriel') }, 'Oui, réponse d’attente (pas encore réglée)'),
    att.type === 'tel' ? h('button', { class: 'sec', onclick: enreg('note', false, 'Appel sans réponse') }, 'Pas de réponse (à rappeler)') : null,
    h('button', { class: 'link', onclick: safe(fermer) }, att.type === 'tel' ? 'Non, je n’ai pas appelé' : 'Non, pas envoyé')));
  document.body.append(box);
}

/* Saisie guidée, une question par écran */
VIEWS.contact = async () => {
  const sup = state.mode === 'superviseur', cfg = state.cfg;
  const ronde = sup ? null : await kvGet('ronde');
  if (!mcDraft) mcDraft = { site: ronde ? ronde.site : undefined, canal: null, objet: null, nom: '', mail: '', tel: '', sansMail: false, demande: '', maintenant: null, quand: localDT(), agent: ronde ? ronde.agent : null, suite: null, service: '', texte: '', etape: null };
  if (!sup) Geo.start();
  const d = mcDraft;
  const etapes = () => ['parking', 'canal', 'objet', 'usager', 'demande', 'quand', ...(sup || ronde ? [] : ['agent']), 'suite', 'ok'];
  if (!d.etape || !etapes().includes(d.etape)) d.etape = etapes()[0];
  const ets = etapes(), k = ets.indexOf(d.etape);
  const aller = e => { d.etape = e; go('contact', {}, { noPush: true }); };
  const suivante = () => aller(ets[k + 1]);
  const quitter = () => { mcDraft = null; if (!sup) geoStopSiLibre(); go(sup ? 'mainCourante' : 'home'); };
  const back = async () => {
    if (k > 0) return aller(ets[k - 1]);
    if (d.nom || d.demande) { if (!(await demander({ titre: 'Abandonner cette saisie ?', oui: 'Oui, abandonner', non: 'Non, continuer', danger: true }))) return; }
    quitter();
  };
  const choixListe = (opts, sel, fn) => opts.map(o => h('button', { class: 'choix' + ((typeof o === 'object' ? o.val : o) === sel ? ' on' : ''), onclick: () => fn(typeof o === 'object' ? o.val : o) }, typeof o === 'object' ? o.lbl : o));
  const service = d.canal === 'service';

  const save = safe(async () => {
    const quand = d.maintenant ? new Date() : new Date(d.quand);
    if (isNaN(quand)) throw new Error('Heure invalide.');
    if (quand.getTime() > Date.now() + 120000) throw new Error('L’heure ne peut pas être dans le futur.');
    if (!d.nom.trim()) throw new Error('Indiquez le nom.');
    if (!d.sansMail && !MAIL_RE.test(d.mail.trim())) throw new Error('Indiquez un courriel valide.');
    if (d.sansMail && !d.tel.trim()) throw new Error('Sans courriel, indiquez un téléphone.');
    if (d.suite === 'regle' && !d.texte.trim()) throw new Error('Indiquez la réponse donnée.');
    if (d.suite === 'service' && !d.service.trim()) throw new Error('Indiquez le service.');
    if (!sup && !d.agent) throw new Error('Indiquez qui enregistre.');
    const c = {
      id: uuid(), ref: 'MC-' + randCode(5), quand: quand.toISOString(), saisi: new Date().toISOString(), site: d.site || null, canal: d.canal, objet: d.objet,
      nom: d.nom.trim(), mail: d.sansMail ? '' : d.mail.trim(), tel: d.tel.trim(), demande: d.demande.trim(), suite: d.suite, service: d.service.trim(), reponse: d.texte.trim()
    };
    if (sup) await queueDec('d', { type: 'contact', id: uuid(), ts: new Date().toISOString(), par: 'DPMS', contact: c });
    else { await appendEvent('contact', { site: c.site, agent: d.agent }, c); await kvSet('dernierAgent', { nom: d.agent, at: Date.now() }); }
    await fold();
    const agent = d.agent;
    mcDraft = null;
    syncNow().then(refreshIfHome);
    if (c.suite === 'barriere') { toast(`${c.ref} enregistré. Déclarez maintenant l’ouverture de la barrière.`); barrDraft = null; go('barriere', { site: c.site || undefined, agent }, { noPush: true }); }
    else if (c.suite === 'signalement') { toast(`${c.ref} enregistré. Faites maintenant le signalement.`); sigDraft = null; go('signalement', { from: 'hr', mc: { id: c.id, site: c.site, agent, desc: c.demande } }, { noPush: true }); }
    else go('contactOk', { ref: c.ref, id: c.id, ouvert: !SUITES_CLOSES.includes(c.suite) }, { noPush: true });
  });

  let corps;
  if (d.etape === 'parking') {
    corps = [h('p', { class: 'question' }, 'Quel parking est concerné ?'),
      ...choixListe([...cfg.sites.map(x => ({ val: x.id, lbl: x.nom })), { val: null, lbl: 'Aucun en particulier' }], d.site, v => { d.site = v; suivante(); })];
  } else if (d.etape === 'canal') {
    corps = [h('p', { class: 'question' }, 'Comment la demande est-elle arrivée ?'), ...choixListe(Object.entries(CANAUX).map(([val, lbl]) => ({ val, lbl })), d.canal, v => { d.canal = v; suivante(); })];
  } else if (d.etape === 'objet') {
    corps = [h('p', { class: 'question' }, 'Quel est l’objet de la demande ?'), ...choixListe(cfg.objetsContact, d.objet, v => { d.objet = v; suivante(); })];
  } else if (d.etape === 'usager') {
    const telBox = h('input', { type: 'tel', inputmode: 'tel', value: d.tel, placeholder: '06 12 34 56 78', oninput: e => d.tel = e.target.value });
    corps = [h('p', { class: 'question' }, service ? 'Quel service ou quel élu ?' : 'Qui est l’usager ?'),
      h('label', { class: 'f' }, service ? 'Nom (et service)' : 'Nom et prénom'),
      h('input', { type: 'text', value: d.nom, autocomplete: 'off', oninput: e => d.nom = e.target.value }),
      h('label', { class: 'f' }, 'Courriel (pour la réponse)'),
      h('input', { type: 'email', inputmode: 'email', value: d.mail, disabled: d.sansMail ? true : null, placeholder: 'nom@exemple.fr', oninput: e => d.mail = e.target.value }),
      h('button', { class: 'case' + (d.sansMail ? ' on' : ''), onclick: () => { d.sansMail = !d.sansMail; render(); } }, h('span', { class: 'coche' }, d.sansMail ? '✓' : ''), 'Pas de courriel'),
      h('label', { class: 'f' }, d.sansMail ? 'Téléphone (obligatoire sans courriel)' : 'Téléphone (facultatif)'), telBox,
      h('button', { class: 'ok big', style: 'margin-top:16px', onclick: safe(() => {
        if (!d.nom.trim()) throw new Error('Indiquez le nom.');
        if (!d.sansMail && !MAIL_RE.test(d.mail.trim())) throw new Error('Indiquez un courriel valide, ou cochez « Pas de courriel ».');
        if (d.sansMail && !d.tel.trim()) throw new Error('Sans courriel, indiquez un téléphone.');
        suivante();
      }) }, 'Suivant')];
  } else if (d.etape === 'demande') {
    corps = [h('p', { class: 'question' }, 'Quelle est la demande ?'),
      h('textarea', { value: d.demande, placeholder: 'Ce que la personne demande ou signale', style: 'min-height:140px', oninput: e => d.demande = e.target.value }),
      h('button', { class: 'ok big', onclick: safe(() => { if (d.demande.trim().length < 3) throw new Error('Décrivez la demande.'); suivante(); }) }, 'Suivant')];
  } else if (d.etape === 'quand') {
    const plusTot = h('div', { style: d.maintenant === false ? '' : 'display:none' },
      h('label', { class: 'f' }, 'Reçue le'),
      h('input', { type: 'datetime-local', value: d.quand, max: localDT(new Date(Date.now() + 60000)), oninput: e => d.quand = e.target.value }),
      h('button', { class: 'ok big', onclick: suivante }, 'Suivant'));
    corps = [h('p', { class: 'question' }, 'Quand la demande a-t-elle été reçue ?'),
      h('button', { class: 'choix' + (d.maintenant === true ? ' on' : ''), onclick: () => { d.maintenant = true; suivante(); } }, 'À l’instant'),
      h('button', { class: 'choix' + (d.maintenant === false ? ' on' : ''), onclick: () => { d.maintenant = false; plusTot.style.display = ''; } }, 'Plus tôt'),
      plusTot];
  } else if (d.etape === 'agent') {
    corps = [h('p', { class: 'question' }, 'Qui enregistre ?'), ...choixListe(cfg.agents, d.agent, v => { d.agent = v; suivante(); })];
  } else if (d.etape === 'suite') {
    const opts = sup ? ['regle', 'service', 'dpms'] : ['regle', 'barriere', 'signalement', 'service', 'dpms'];
    corps = [h('p', { class: 'question' }, 'Quelle suite a été donnée ?'), ...choixListe(opts.map(val => ({ val, lbl: val === 'dpms' && sup ? 'À traiter (réponse à apporter)' : SUITES[val] })), d.suite, v => { d.suite = v; suivante(); })];
  } else {
    const quand = d.maintenant ? 'à l’instant' : (isNaN(new Date(d.quand)) ? d.quand : fmtDT(new Date(d.quand).toISOString()));
    corps = [
      h('div', { class: 'card resume' },
        h('b', null, `${d.objet} — ${d.site ? siteNom(d.site) : 'général'}`),
        h('div', null, `${CANAUX[d.canal]}, ${quand}`),
        h('div', null, `${d.nom}${d.mail && !d.sansMail ? ' — ' + d.mail : ''}${d.tel ? ' — ' + d.tel : ''}`),
        h('div', { class: 'muted small' }, d.demande),
        h('div', { class: 'muted small' }, `Suite : ${SUITES[d.suite]}${sup ? '' : ` — enregistré par ${d.agent}`}`)),
      d.suite === 'regle' ? [h('p', { class: 'question' }, 'Quelle réponse a été donnée ?'), h('textarea', { value: d.texte, oninput: e => d.texte = e.target.value })] : null,
      d.suite === 'service' ? [h('p', { class: 'question' }, 'Vers quel service ?'), h('input', { type: 'text', value: d.service, placeholder: 'Ex. : stationnement, services techniques, police municipale', oninput: e => d.service = e.target.value }),
        h('label', { class: 'f' }, 'Précision (facultatif)'), h('input', { type: 'text', value: d.texte, oninput: e => d.texte = e.target.value })] : null,
      ['barriere', 'signalement', 'dpms'].includes(d.suite) ? [h('label', { class: 'f' }, 'Commentaire (facultatif)'), h('input', { type: 'text', value: d.texte, oninput: e => d.texte = e.target.value })] : null,
      h('button', { class: 'ok big', style: 'margin-top:16px', onclick: save },
        d.suite === 'barriere' ? 'Enregistrer, puis déclarer la barrière' : d.suite === 'signalement' ? 'Enregistrer, puis faire le signalement' : 'Enregistrer')];
  }
  return page('Usager ou demande reçue', { back },
    h('div', { class: 'etapes' }, h('div', { class: 'etapes-txt' }, `Étape ${k + 1} sur ${ets.length}`),
      h('div', { class: 'etapes-barre' }, ...ets.map((e, n) => h('span', { class: n < k ? 'fait' : n === k ? 'en-cours' : '' })))),
    ...corps.flat().filter(Boolean),
    k > 0 ? h('button', { class: 'link', onclick: back }, '◀ Étape précédente') : null);
};
VIEWS.contactOk = async ({ ref, id, ouvert }) => page('Usagers et demandes', null,
  h('div', { class: 'fini' }, h('span', { html: ICO.ok }), h('h2', null, 'Demande enregistrée'), h('p', null, `Référence ${ref}`),
    ouvert ? h('p', { class: 'muted' }, 'Demande à traiter : la réponse apportée à la personne s’enregistre depuis « Plus » → « Suivi des usagers et demandes ».') : null),
  h('button', { class: 'big', onclick: () => go('home') }, 'Retour à l’accueil'),
  h('button', { class: 'sec', onclick: () => go('contactFiche', { id }) }, 'Voir la fiche'));

/* Liste */
VIEWS.mainCourante = async ({ f } = {}) => {
  const tous = await mainCourante();
  f = f || (tous.some(c => c.statut === 'ouvert') ? 'atraiter' : 'semaine');
  const auj = localDate(), il7 = Date.now() - 7 * 86400000;
  const filtres = { atraiter: c => c.statut === 'ouvert', jour: c => localDate(new Date(c.quand)) === auj, semaine: c => new Date(c.quand).getTime() >= il7, tout: () => true };
  const liste = tous.filter(filtres[f]).sort((a, b) => f === 'atraiter' ? a.quand.localeCompare(b.quand) : b.quand.localeCompare(a.quand));
  return page('Suivi des usagers et demandes', { back: true },
    h('p', { class: 'muted small', style: 'margin:0 0 10px' }, 'Les demandes déjà enregistrées. Pour en noter une nouvelle : bouton ci-dessous, ou « Usager ou demande reçue » sur l’accueil.'),
    h('button', { class: 'ok big', onclick: () => { mcDraft = null; go('contact'); } }, 'Nouvelle entrée'),
    chips([{ val: 'atraiter', lbl: `À traiter (${tous.filter(filtres.atraiter).length})` }, { val: 'jour', lbl: 'Aujourd’hui' }, { val: 'semaine', lbl: '7 jours' }, { val: 'tout', lbl: 'Tout' }], f, v => go('mainCourante', { f: v }, { noPush: true })),
    liste.length ? liste.map(c => h('button', { class: 'list-item', onclick: () => go('contactFiche', { id: c.id }) },
      h('div', { class: 'l1' }, `${c.objet} — ${c.nom}`, badgeMC(c)),
      h('div', { class: 'muted small' }, `${c.ref} · ${fmtDT(c.quand)} · ${lieuMC(c)} · ${CANAUX[c.canal] || ''}`),
      h('div', { class: 'small' }, c.demande.length > 120 ? c.demande.slice(0, 119) + '…' : c.demande)))
      : h('p', { class: 'muted' }, 'Aucune entrée.'));
};

/* Fiche */
VIEWS.contactFiche = async ({ id, mode }) => {
  const c = ((await kvGet('mainCourante')) || []).find(x => x.id === id);
  if (!c) { toast('Entrée introuvable.', true); return VIEWS.mainCourante({}); }
  const sup = state.mode === 'superviseur', equipe = !sup;
  const sigs = [];
  for (const sid of c.sigs) { const s = await dbGet('sigs', sid); if (s) sigs.push(s); }
  const d = { par: equipe ? await agentPrefere() : null, texte: '', mode: MODE_CANAL[c.canal] || 'place' };
  const fiche = () => go('contactFiche', { id }, { noPush: true });
  let bas;
  if (mode === 'reponse' || mode === 'note') {
    const enreg = regle => safe(async () => {
      if (!d.texte.trim()) throw new Error(mode === 'note' ? 'Écrivez la note.' : 'Indiquez la réponse apportée.');
      await enregistrerSuiviMC(c, { type: mode, mode: mode === 'reponse' ? d.mode : null, texte: d.texte.trim(), regle }, d.par);
      toast(mode === 'note' ? 'Note ajoutée.' : regle ? 'Réponse enregistrée, demande réglée.' : 'Réponse d’attente enregistrée.');
      fiche();
    });
    bas = h('div', { class: 'card' }, h('h2', null, mode === 'note' ? 'Ajouter une note' : 'Réponse apportée à la personne'),
      mode === 'reponse' ? [h('label', { class: 'f' }, 'Comment ?'), chips(Object.entries(MODES_REP).map(([val, lbl]) => ({ val, lbl })), d.mode, v => d.mode = v)] : null,
      h('textarea', { placeholder: mode === 'note' ? 'Note interne' : 'Ce qui a été répondu', oninput: e => d.texte = e.target.value }),
      equipe ? [h('label', { class: 'f' }, 'Qui ?'), chips(state.cfg.agents, d.par, v => d.par = v)] : null,
      mode === 'reponse' ? [h('button', { class: 'ok', onclick: enreg(true) }, 'Enregistrer : demande réglée'), h('button', { class: 'sec', onclick: enreg(false) }, 'Enregistrer : réponse d’attente')]
        : h('button', { class: 'ok', onclick: enreg(false) }, 'Enregistrer la note'),
      h('button', { class: 'link', onclick: fiche }, 'Annuler'));
  } else {
    bas = h('div', { class: 'card' }, h('h2', null, 'Que voulez-vous faire ?'),
      c.mail ? h('a', { class: 'btn ok', href: courrielUsager(c), onclick: () => { lancerContactMC(c, 'mail'); } }, 'Répondre par courriel') : null,
      c.tel ? h('a', { class: 'btn sec', href: 'tel:' + c.tel.replace(/\s+/g, ''), onclick: () => { lancerContactMC(c, 'tel'); } }, 'Appeler ' + c.nom) : null,
      h('button', { class: 'sec', onclick: () => go('contactFiche', { id, mode: 'reponse' }, { noPush: true }) }, 'Enregistrer une réponse déjà donnée'),
      h('button', { class: 'sec', onclick: () => go('contactFiche', { id, mode: 'note' }, { noPush: true }) }, 'Ajouter une note'),
      c.statut === 'ouvert'
        ? h('button', { class: 'link', onclick: safe(async () => {
          const motif = prompt('Clore sans réponse — motif :', '');
          if (motif === null) return;
          let par = d.par;
          if (equipe && !par) { par = prompt('Votre nom :', ''); if (!par) return; }
          await enregistrerSuiviMC(c, { type: 'clos', texte: motif.trim() || 'Clos sans réponse' }, par); toast('Demande close.'); fiche();
        }) }, 'Clore sans réponse')
        : h('button', { class: 'link', onclick: safe(async () => {
          let par = d.par;
          if (equipe && !par) { par = prompt('Votre nom :', ''); if (!par) return; }
          await enregistrerSuiviMC(c, { type: 'rouvert', texte: 'Rouverte' }, par); toast('Demande rouverte.'); fiche();
        }) }, 'Rouvrir'),
      sup ? h('button', { class: 'link', style: 'color:var(--bad)', onclick: safe(async () => {
        if (!(await demander({ titre: 'Annuler cette entrée ?', texte: 'À utiliser pour une saisie erronée : l’entrée disparaît des listes, de la synthèse et des exports (elle reste dans le journal).', oui: 'Oui, annuler l’entrée', non: 'Non', danger: true }))) return;
        await enregistrerSuiviMC(c, { type: 'annule', texte: 'Entrée annulée' }); toast('Entrée annulée.'); go('mainCourante', {}, { noPush: true });
      }) }, 'Annuler l’entrée (saisie erronée)') : null);
  }
  const TYPES_HIST = { reponse: 'Réponse', note: 'Note', clos: 'Close', rouvert: 'Rouverte', annule: 'Annulée' };
  const dl = delaiMC(c);
  return page(`Main courante ${c.ref}`, { back: true },
    h('div', { class: 'card' },
      h('div', { class: 'l1', style: 'display:flex;justify-content:space-between;gap:8px;align-items:center' }, h('b', null, `${c.objet} — ${lieuMC(c)}`), badgeMC(c)),
      h('dl', { class: 'kv' },
        h('dt', null, 'Reçue'), h('dd', null, `${fmtDT(c.quand)} — ${CANAUX[c.canal] || ''}`),
        h('dt', null, c.canal === 'service' ? 'Service ou élu' : 'Usager'), h('dd', null, c.nom),
        c.mail ? [h('dt', null, 'Courriel'), h('dd', null, h('a', { href: 'mailto:' + c.mail }, c.mail))] : null,
        c.tel ? [h('dt', null, 'Téléphone'), h('dd', null, h('a', { href: 'tel:' + c.tel.replace(/\s+/g, '') }, c.tel))] : null,
        h('dt', null, 'Demande'), h('dd', { style: 'white-space:pre-wrap' }, c.demande),
        h('dt', null, 'Suite'), h('dd', null, SUITES[c.suite] + (c.service ? ' : ' + c.service : '')),
        h('dt', null, 'Enregistrée par'), h('dd', null, `${c.par}${c.saisi && c.saisi.slice(0, 16) !== c.quand.slice(0, 16) ? ` le ${fmtDT(c.saisi)}` : ''}`),
        c.reponduLe ? [h('dt', null, 'Première réponse'), h('dd', null, `${fmtDT(c.reponduLe)}${dl ? ` (délai : ${dureeTxt(dl)})` : ''}`)] : null)),
    sigs.length ? h('div', { class: 'card' }, h('h2', null, 'Signalement lié'), ...sigs.map(s => h('button', { class: 'list-item', onclick: () => go(sup ? 'supSig' : 'suiviSig', { id: s.id }) },
      h('div', { class: 'l1' }, `${s.ref} — ${sigTitre(s)}`, suiviBadge(s))))) : null,
    c.hist.length ? h('div', { class: 'card' }, h('h2', null, 'Historique'), ...c.hist.slice().sort((a, b) => a.quand.localeCompare(b.quand)).map(x => h('div', { class: 'stat', style: 'display:block' },
      h('div', null, h('b', null, `${TYPES_HIST[x.type] || x.type}${x.type === 'reponse' ? (x.regle ? ' (réglée)' : ' (attente)') : ''}`), ` — ${fmtDT(x.quand)} — ${x.par}${x.mode ? ' — ' + (MODES_REP[x.mode] || '').toLowerCase() : ''}`),
      x.texte ? h('div', { class: 'small' }, x.texte) : null))) : null,
    bas);
};
async function carteMainCourante() {
  const l = await mainCourante(), auj = localDate();
  const ouverts = l.filter(c => c.statut === 'ouvert');
  return h('div', { class: 'card' }, h('h2', null, 'Usagers et demandes (main courante)'),
    h('button', { class: 'list-item', onclick: () => go('mainCourante', { f: 'atraiter' }) }, h('div', { class: 'stat' }, h('span', null, 'Demandes à traiter', ouverts.some(c => joursDepuis(c.quand) >= 2) ? h('span', { class: 'badge b-bad' }, 'en attente') : null), h('b', null, ouverts.length))),
    h('button', { class: 'list-item', onclick: () => go('mainCourante', { f: 'jour' }) }, h('div', { class: 'stat' }, h('span', null, 'Contacts aujourd’hui'), h('b', null, l.filter(c => localDate(new Date(c.quand)) === auj).length))),
    h('div', { class: 'row' }, h('button', { class: 'sec', onclick: () => { mcDraft = null; go('contact'); } }, 'Nouvelle entrée'), h('button', { class: 'sec', onclick: () => go('mainCourante', {}) }, 'Toute la main courante')));
}
/* ====================================================================== */
/* PLANNING MENSUEL                                                       */
/* Superviseur : grille du mois (semaines complètes), pré-remplie par le  */
/* planning-type, absences saisies, contrôles, prompt pour Claude et      */
/* import de sa réponse, publication. Agents : onglet Planning.           */
/* ====================================================================== */
const ABSENCES = { CA: 'Congés', MAL: 'Maladie', FOR: 'Formation', ABS: 'Absence' };
/* Jours fériés (art. L3133-1 du code du travail) : 1er janvier, lundi de Pâques, 1er et 8 mai, Ascension,
   lundi de Pentecôte, 14 juillet, 15 août, 1er et 11 novembre, 25 décembre. Parkings fermés ces jours-là
   (règle « jours fériés travaillés : … » pour des exceptions). */
function joursFeries(an) {
  const a = an % 19, b = Math.floor(an / 100), c = an % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3);
  const hh = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - hh - k) % 7, m = Math.floor((a + 11 * hh + 22 * l) / 451);
  const mois = Math.floor((hh + l - 7 * m + 114) / 31), jour = ((hh + l - 7 * m + 114) % 31) + 1;
  const paques = localDate(new Date(an, mois - 1, jour));
  const f2 = (mm, jj) => `${an}-${String(mm).padStart(2, '0')}-${String(jj).padStart(2, '0')}`;
  return new Map([[f2(1, 1), 'Jour de l’an'], [ajJ(paques, 1), 'Lundi de Pâques'], [f2(5, 1), 'Fête du travail'], [f2(5, 8), 'Victoire 1945'],
    [ajJ(paques, 39), 'Ascension'], [ajJ(paques, 50), 'Lundi de Pentecôte'], [f2(7, 14), 'Fête nationale'], [f2(8, 15), 'Assomption'],
    [f2(11, 1), 'Toussaint'], [f2(11, 11), 'Armistice 1918'], [f2(12, 25), 'Noël']]);
}
function ferieFerme(iso) {   // nom du jour férié si les parkings sont fermés ce jour-là, sinon null
  const nom = joursFeries(Number(iso.slice(0, 4))).get(iso);
  if (!nom) return null;
  const r = reglesPlanning();
  if (r.feriesOuverts) return null;
  return r.feriesTravailles.some(x => sansAccent(nom).includes(sansAccent(x)) || sansAccent(x).includes(sansAccent(nom))) ? null : nom;
}
const PLANNING_DEFAUT = {
  ancre: '2026-09-14',
  vacations: [{ code: 'MAT', lbl: 'Matin', debut: '08:00', fin: '15:00' }, { code: 'MID', lbl: 'Midi', debut: '10:30', fin: '17:30' }, { code: 'SOI', lbl: 'Soir', debut: '13:00', fin: '20:00' }],
  type: '',
  regles: 'couverture : 08:00-20:00\nheures par quinzaine : 70'
};
const NOMS_JOURS = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
const JC = ['D', 'L', 'Ma', 'Me', 'J', 'V', 'S'];
const sansAccent = t => String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
const dISO = d => localDate(d);
const dDe = iso => { const [y, m, j] = iso.split('-').map(Number); return new Date(y, m - 1, j); };
const ajJ = (iso, n) => { const d = dDe(iso); d.setDate(d.getDate() + n); return dISO(d); };
const ecartJ = (a, b) => Math.round((dDe(b) - dDe(a)) / 86400000);
const jourCourt = iso => { const d = dDe(iso); return `${JC[d.getDay()]} ${d.getDate()}`; };
const jourLong = iso => { const d = dDe(iso); return `${NOMS_JOURS[d.getDay()]} ${d.getDate()} ${MOIS_LONGS[d.getMonth()]}`; };
const nomMois = m => { const [y, mm] = m.split('-').map(Number); return `${MOIS_LONGS[mm - 1]} ${y}`; };
const moisDe = iso => iso.slice(0, 7);
const moisSuivant = (m, n = 1) => { const [y, mm] = m.split('-').map(Number); const d = new Date(y, mm - 1 + n, 1); return dISO(d).slice(0, 7); };
function periodeMois(m) {   // semaines dont le lundi tombe dans le mois : du 1er lundi au dimanche qui suit le dernier lundi
  const [y, mm] = m.split('-').map(Number);
  const p = new Date(y, mm - 1, 1), d = new Date(y, mm, 0);
  p.setDate(p.getDate() + (8 - p.getDay()) % 7); d.setDate(d.getDate() - ((d.getDay() + 6) % 7) + 6);
  const jours = []; for (let x = new Date(p); x <= d; x.setDate(x.getDate() + 1)) jours.push(dISO(x));
  return { du: jours[0], au: jours[jours.length - 1], jours };
}
const minDe = hm => { const [a, b] = String(hm).split(':').map(Number); return a * 60 + (b || 0); };
const hDe = min => `${String(Math.floor(min / 60)).padStart(2, '0')}:${String(min % 60).padStart(2, '0')}`;

/* Configuration : planning-type et règles (texte modifiable par le superviseur) */
function cfgPlanning() { return { ...PLANNING_DEFAUT, ...((state.cfg && state.cfg.planning) || {}) }; }
function vacationsP() { return cfgPlanning().vacations; }
const estVacation = code => vacationsP().some(v => v.code === code);
const libelleCode = (code, pourAgents) => {
  const v = vacationsP().find(x => x.code === code);
  if (v) return `${v.lbl} ${v.debut.replace(':', 'h').replace('h00', 'h')}-${v.fin.replace(':', 'h').replace('h00', 'h')}`;
  if (code === 'R') return 'Repos';
  if (code === 'F') return pourAgents ? 'Férié' : 'Férié (fermé)';
  if (ABSENCES[code]) return pourAgents ? 'Absent' : ABSENCES[code];
  return '';
};
const CODES_OK = () => new Set([...vacationsP().map(v => v.code), 'R', 'F', ...Object.keys(ABSENCES)]);
function trouverNom(token, noms) {
  const t = sansAccent(token).replace(/\s+/g, ' ');
  if (!t) return null;
  return noms.find(n => sansAccent(n) === t) || noms.find(n => sansAccent(n).startsWith(t) || t.startsWith(sansAccent(n).split(' ')[0]) && sansAccent(n).split(' ')[0].length >= 3) || null;
}
function reglesPlanning() {
  const r = { couverture: [480, 1200], heures: 70, incompatibles: [], interdits: [], internes: [], feriesOuverts: false, feriesTravailles: [] };
  for (const raw of (cfgPlanning().regles || '').split('\n')) {
    const i = raw.indexOf(':'); if (i < 0) continue;
    const k = sansAccent(raw.slice(0, i)), v = raw.slice(i + 1).trim();
    const noms = v.split(/[,;]/).map(x => x.trim()).filter(Boolean);
    if (k.startsWith('couverture')) { const m = /(\d{1,2}[:h]\d{2})\s*-\s*(\d{1,2}[:h]\d{2})/.exec(v); if (m) r.couverture = [minDe(m[1].replace('h', ':')), minDe(m[2].replace('h', ':'))]; }
    else if (k.startsWith('heures')) { const n = parseFloat(v.replace(',', '.')); if (n > 0) r.heures = n; }
    else if (k.startsWith('jamais sur la meme vacation') || k.startsWith('jamais ensemble')) { if (noms.length > 1) r.incompatibles.push(noms); }
    else if (k.startsWith('jamais le ')) { const j = NOMS_JOURS.indexOf(k.slice(10).trim()); if (j >= 0) noms.forEach(n => r.interdits.push({ nom: n, jour: j })); }
    else if (k.startsWith('lignes internes') || k.startsWith('ligne interne')) r.internes.push(...noms);
    else if (k.startsWith('jours feries travailles')) r.feriesTravailles.push(...noms);
    else if (k.startsWith('jours feries')) r.feriesOuverts = /ouvert/.test(sansAccent(v));
  }
  return r;
}
function typePlanning() {   // { nom: [14 codes] }, ordre conservé
  const out = [];
  for (const raw of (cfgPlanning().type || '').split('\n')) {
    const i = raw.indexOf(':'); if (i < 0) continue;
    const nom = raw.slice(0, i).trim(); if (!nom) continue;
    const codes = raw.slice(i + 1).replace(/\|/g, ' ').split(/\s+/).map(x => x.trim().toUpperCase()).filter(Boolean).map(x => x === '-' || x === '—' ? 'R' : x);
    if (codes.length === 14) out.push({ nom, codes });
  }
  return out;
}
function lignesPlanning() {   // agents du planning-type (ou liste des agents), puis lignes internes
  const t = typePlanning(), r = reglesPlanning();
  const agents = t.length ? t.map(x => x.nom) : (state.cfg.agents || []).slice();
  return { agents, internes: r.internes.filter(n => !agents.includes(n)) };
}
function codeType(nom, iso) {
  const t = typePlanning().find(x => x.nom === nom);
  if (!t) return '';
  const k = ((ecartJ(cfgPlanning().ancre, iso) % 14) + 14) % 14;
  return t.codes[k] || '';
}

/* Contrôles */
function verifierPlanning(grille, jours) {
  const regles = reglesPlanning(), { agents, internes } = lignesPlanning(), vacs = vacationsP();
  const pb = [], parJour = {};
  const ajout = (iso, msg, niveau = 'bad') => { pb.push({ iso, msg, niveau }); (parJour[iso] = parJour[iso] || []).push(niveau); };
  const vac = c => vacs.find(v => v.code === c);
  for (const iso of jours) {
    const g = grille[iso] || {}, wd = dDe(iso).getDay();
    const ferie = ferieFerme(iso);
    if (ferie) {   // parkings fermés : pas de couverture exigée ; quelqu'un de service est signalé
      for (const n of [...agents, ...internes]) if (vac(g[n])) ajout(iso, `${n} est de service un jour férié (${ferie}, parkings fermés)`);
      continue;
    }
    // couverture continue
    const iv = [...agents, ...internes].map(n => vac(g[n])).filter(Boolean).map(v => [minDe(v.debut), minDe(v.fin)]).sort((a, b) => a[0] - b[0]);
    let t = regles.couverture[0]; const trous = [];
    for (const [a, b] of iv) { if (a > t) trous.push([t, Math.min(a, regles.couverture[1])]); t = Math.max(t, b); if (t >= regles.couverture[1]) break; }
    if (t < regles.couverture[1]) trous.push([t, regles.couverture[1]]);
    trous.filter(([a, b]) => b > a).forEach(([a, b]) => ajout(iso, `personne de ${hDe(a)} à ${hDe(b)}`));
    // jamais sur la même vacation
    for (const groupe of regles.incompatibles) {
      const ns = groupe.map(x => trouverNom(x, agents)).filter(Boolean);
      for (let i = 0; i < ns.length; i++) for (let j = i + 1; j < ns.length; j++) if (vac(g[ns[i]]) && g[ns[i]] === g[ns[j]]) ajout(iso, `${ns[i]} et ${ns[j]} sur la même vacation (${vac(g[ns[i]]).lbl.toLowerCase()})`);
    }
    // jours interdits
    for (const x of regles.interdits) { const n = trouverNom(x.nom, agents); if (n && x.jour === wd && vac(g[n])) ajout(iso, `${n} ne travaille jamais le ${NOMS_JOURS[wd]}`); }
    // repos du planning-type travaillé (information)
    for (const n of agents) if (vac(g[n]) && codeType(n, iso) === 'R') ajout(iso, `${n} travaille un jour de repos du planning-type`, 'info');
  }
  // heures par quinzaine (une absence un jour normalement travaillé compte comme travaillée)
  const heures = {};
  const anc = cfgPlanning().ancre;
  const debQ = iso => ajJ(iso, -(((ecartJ(anc, iso) % 14) + 14) % 14));
  const quinz = [...new Set(jours.map(debQ))].filter(q => jours.includes(q) && jours.includes(ajJ(q, 13)));
  for (const n of agents) {
    heures[n] = quinz.map(q => {
      let min = 0;
      for (let i = 0; i < 14; i++) {
        const iso = ajJ(q, i), c = (grille[iso] || {})[n], v = vac(c);
        if (v) min += minDe(v.fin) - minDe(v.debut);
        else if (ABSENCES[c] || c === 'F') { const vt = vac(codeType(n, iso)); if (vt) min += minDe(vt.fin) - minDe(vt.debut); }   // absence ou férié un jour normalement travaillé
      }
      const h = Math.round(min / 6) / 10;
      if (h !== regles.heures) pb.push({ iso: q, msg: `${n} : ${String(h).replace('.', ',')} h sur la quinzaine du ${jourCourt(q)} (attendu ${regles.heures} h)`, niveau: 'bad', quinzaine: true });
      return { du: q, h };
    });
  }
  return { pb, parJour, heures, quinz };
}

/* État : publié (journal), brouillon (local au superviseur) */
async function planningPublie() { return (await kvGet('planning')) || { jours: {}, notes: {}, pubs: [] }; }
async function grilleMois(m) {
  const pub = await planningPublie(), motifs = state.mode === 'superviseur' ? ((await kvGet('planningMotifs')) || {}) : {};
  const br = ((await kvGet('planningBrouillons')) || {})[m];
  const { jours } = periodeMois(m), { agents, internes } = lignesPlanning();
  const base = {}, grille = {};
  for (const iso of jours) {
    base[iso] = {}; grille[iso] = {};
    for (const n of [...agents, ...internes]) {
      const p = (motifs[iso] || {})[n] || (pub.jours[iso] || {})[n];
      base[iso][n] = p || '';
      grille[iso][n] = (br && br.cells[iso] && br.cells[iso][n] != null) ? br.cells[iso][n] : (p || (ferieFerme(iso) ? 'F' : codeType(n, iso)) || '');
    }
  }
  const publieMois = pub.pubs.filter(x => x.mois === m).slice(-1)[0] || null;
  return { jours, agents, internes, base, grille, brouillon: br || null, publieMois, notes: br && br.notes != null ? br.notes : (pub.notes[m] || ''), pub };
}
async function majBrouillon(m, fn) {
  await kvUpdate('planningBrouillons', all => {
    const a = all || {}; const b = a[m] || { cells: {}, notes: null, claude: [] };
    fn(b); b.maj = new Date().toISOString(); a[m] = b; return a;
  });
}
const normCode = c => c && c !== 'R' ? c : 'R';   // case vide = repos
function differences(G, avant, jours, lignes) {
  const out = [];
  for (const iso of jours) for (const n of lignes) { const a = normCode((avant[iso] || {})[n]), b = normCode((G[iso] || {})[n]); if (a !== b) out.push({ iso, n, a, b }); }
  return out;
}

/* Prompt pour Claude */
function promptClaude(m, G, consigne) {
  const { jours, agents, internes } = G, cp = cfgPlanning(), regles = reglesPlanning(), v = verifierPlanning(G.grille, jours);
  const L = [...agents, ...internes], larg = Math.max(...L.map(n => n.length), 8);
  const ligne = (n, codes) => n.padEnd(larg) + ' : ' + codes.map(c => (c || 'R').padEnd(3)).join(' ');
  const absences = [];
  for (const n of agents) {   // périodes consécutives regroupées
    let deb = null, code = null;
    jours.concat([null]).forEach((iso, k) => {
      const c = iso ? G.grille[iso][n] : null;
      if (deb && c !== code) { const fin = jours[k - 1]; absences.push(`${n} : ${ABSENCES[code].toLowerCase()} ${deb === fin ? 'le ' + jourLong(deb) : `du ${jourLong(deb)} au ${jourLong(fin)}`}`); deb = null; }
      if (iso && ABSENCES[c] && !deb) { deb = iso; code = c; }
    });
  }
  return [
    `Tu m'aides à établir le planning des agents de surveillance des parkings municipaux. Planning de ${nomMois(m)} (les semaines dont le lundi tombe dans le mois) : du ${jourLong(jours[0])} au ${jourLong(jours[jours.length - 1])} (${jours.length} jours).`,
    '',
    'VACATIONS ET CODES',
    ...cp.vacations.map(x => `${x.code} = ${x.lbl.toLowerCase()} ${x.debut}-${x.fin}`),
    'R = repos ; F = jour férié (parkings fermés) ; CA = congés ; MAL = maladie ; FOR = formation ; ABS = autre absence',
    '',
    'RÈGLES',
    `- Une présence continue de ${hDe(regles.couverture[0])} à ${hDe(regles.couverture[1])} chaque jour.`,
    `- ${regles.heures} h par agent et par quinzaine ; une absence ou un jour férié un jour normalement travaillé compte comme travaillé. Quinzaines : cycle de deux semaines commençant le ${jourLong(cp.ancre)}.`,
    ...regles.incompatibles.map(g => `- ${g.join(' et ')} : jamais sur la même vacation (de préférence matin contre soir).`),
    ...regles.interdits.map(x => `- ${x.nom} ne travaille jamais le ${NOMS_JOURS[x.jour]}.`),
    ...(internes.length ? [`- Ligne(s) interne(s), non montrée(s) aux agents : ${internes.join(', ')}.`] : []),
    (() => { const f = jours.filter(ferieFerme); return f.length ? `- Jours fériés : parkings fermés, personne ne travaille (code F pour tous) : ${f.map(iso => `${jourLong(iso)} (${ferieFerme(iso)})`).join(', ')}. Un jour férié normalement travaillé compte comme travaillé.` : '- Aucun jour férié sur la période.'; })(),
    '- Les absences (CA, MAL, FOR, ABS) et les jours fériés (F) sont imposés : ne les modifie jamais.',
    '- Respecte au maximum les jours de repos du planning-type et change le moins de cases possible.',
    '',
    `PLANNING-TYPE (2 semaines ; la semaine 1 commence le ${jourLong(cp.ancre)})`,
    ...typePlanning().map(t => `${t.nom.padEnd(larg)} : ${t.codes.slice(0, 7).join(' ')} | ${t.codes.slice(7).join(' ')}`),
    '',
    'ABSENCES SAISIES',
    ...(absences.length ? absences : ['aucune']),
    '',
    'PROJET ACTUEL (planning-type et absences)',
    ''.padEnd(larg) + '   ' + jours.map(iso => jourCourt(iso).replace(' ', '').padEnd(3)).join(' '),
    ...L.map(n => ligne(n, jours.map(iso => G.grille[iso][n]))),
    '',
    'CONTRÔLES DU PROJET',
    ...(v.pb.filter(p => p.niveau === 'bad').length ? v.pb.filter(p => p.niveau === 'bad').map(p => `- ${p.quinzaine ? '' : jourLong(p.iso) + ' : '}${p.msg}`) : ['- aucun problème détecté']),
    '',
    'DEMANDE',
    consigne && consigne.trim() ? consigne.trim() : 'Propose le planning définitif en corrigeant les problèmes ci-dessus, en tenant compte des absences.',
    '',
    'FORMAT DE RÉPONSE OBLIGATOIRE',
    `Explique brièvement tes choix (cases modifiées et pourquoi, points restant à arbitrer), puis termine ta réponse par le bloc ci-dessous, dans un bloc de code : une ligne par agent, exactement ${jours.length} codes séparés par des espaces, du ${jourCourt(jours[0])} au ${jourCourt(jours[jours.length - 1])}, uniquement les codes listés plus haut.`,
    '```',
    `=== PLANNING-CLAUDE ${m} ===`,
    ...L.map(n => `${n} : …`),
    '=== FIN ===',
    '```'
  ].join('\n');
}
function lireReponseClaude(texte, G) {
  const t = String(texte || '').replace(/\r/g, '');
  const d = t.search(/=+\s*PLANNING-CLAUDE/i);
  if (d < 0) throw new Error('Bloc « === PLANNING-CLAUDE … === » introuvable : copiez toute la réponse de Claude, ou au moins son bloc final.');
  const f = t.slice(d).search(/=+\s*FIN\s*=+/i);
  const bloc = t.slice(d, f < 0 ? undefined : d + f).split('\n');
  const mois = (/PLANNING-CLAUDE\s+(\d{4}-\d{2})/i.exec(bloc[0]) || [])[1] || null;
  const L = [...G.agents, ...G.internes], ok = CODES_OK(), cells = {}, erreurs = [], vus = new Set();
  for (const raw of bloc.slice(1)) {
    const i = raw.indexOf(':'); if (i < 0) continue;
    const nom = trouverNom(raw.slice(0, i).replace(/[`*]/g, ''), L);
    if (!nom) { if (raw.trim() && !/^`/.test(raw.trim())) erreurs.push(`ligne ignorée (agent inconnu) : « ${raw.slice(0, i).trim()} »`); continue; }
    const codes = raw.slice(i + 1).replace(/\|/g, ' ').split(/\s+/).map(x => x.trim().toUpperCase().replace(/[^A-Z]/g, '')).filter(Boolean).map(x => x === 'REPOS' ? 'R' : x);
    if (codes.length !== G.jours.length) { erreurs.push(`${nom} : ${codes.length} codes au lieu de ${G.jours.length} — ligne non appliquée`); continue; }
    const mauvais = codes.filter(c => !ok.has(c));
    if (mauvais.length) { erreurs.push(`${nom} : codes inconnus (${[...new Set(mauvais)].join(', ')}) — ligne non appliquée`); continue; }
    vus.add(nom);
    G.jours.forEach((iso, k) => { (cells[iso] = cells[iso] || {})[nom] = codes[k]; });
  }
  if (!vus.size) throw new Error('Aucune ligne exploitable dans le bloc de Claude.' + (erreurs.length ? ' ' + erreurs[0] : ''));
  // absences imposées : jamais remplacées par la proposition
  for (const iso of G.jours) for (const n of vus) {
    const avant = G.grille[iso][n];
    if (ABSENCES[avant] && cells[iso][n] !== avant) { erreurs.push(`${n}, ${jourLong(iso)} : absence (${ABSENCES[avant].toLowerCase()}) conservée`); cells[iso][n] = avant; }
    if (avant === 'F' && cells[iso][n] !== 'F' && ferieFerme(iso)) { erreurs.push(`${n}, ${jourLong(iso)} : jour férié conservé`); cells[iso][n] = 'F'; }
  }
  return { cells, erreurs, mois, lignes: [...vus] };
}

/* Vue superviseur */
let _planSel = null;   // { iso, n } case en cours de choix
VIEWS.planning = async ({ mois } = {}) => {
  if (state.mode !== 'superviseur') return VIEWS.planningAgents({});
  const auj = localDate();
  const m = mois || (Number(auj.slice(8)) >= 15 ? moisSuivant(moisDe(auj)) : moisDe(auj));
  const G = await grilleMois(m);
  const recharger = () => go('planning', { mois: m }, { noPush: true, keepScroll: true });
  const lignes = [...G.agents, ...G.internes];
  if (!typePlanning().length) {
    return page('Planning', { back: true },
      h('div', { class: 'banner warn' }, 'Le planning-type n’est pas encore saisi : il sert à pré-remplir chaque mois et à contrôler les repos.'),
      h('button', { class: 'ok big', onclick: () => go('planningParam') }, 'Saisir le planning-type et les règles'));
  }
  const v = verifierPlanning(G.grille, G.jours);
  const diffs = differences(G.grille, G.base, G.jours, lignes);
  const importes = new Set((G.brouillon && G.brouillon.claude) || []);
  const vac = c => vacationsP().find(x => x.code === c);
  const choisir = (iso, n) => {
    const actuel = G.grille[iso][n];
    const opts = [...vacationsP().map(x => ({ c: x.code, l: `${x.lbl} ${x.debut}-${x.fin}` })), { c: 'R', l: 'Repos' }, { c: 'F', l: 'Férié (fermé)' }, ...Object.entries(ABSENCES).map(([c, l]) => ({ c, l }))];
    const fermer = () => box.remove();
    const poser = c => safe(async () => { await majBrouillon(m, b => { (b.cells[iso] = b.cells[iso] || {})[n] = c; b.claude = (b.claude || []).filter(k => k !== iso + '|' + n); }); fermer(); recharger(); });
    const box = h('div', { class: 'confirm-contact', onclick: e => { if (e.target === box) fermer(); } }, h('div', { class: 'sheet' },
      h('h3', null, `${n} — ${jourLong(iso)}`),
      h('p', { class: 'muted small' }, `Planning-type : ${libelleCode(codeType(n, iso)) || 'repos'}${ferieFerme(iso) ? ` · ${ferieFerme(iso)} : férié, parkings fermés` : ''}`),
      h('div', { class: 'plan-choix' }, ...opts.map(o => h('button', { class: `pc pc-${o.c}${o.c === actuel ? ' on' : ''}`, onclick: poser(o.c) }, h('b', null, o.c), ' ', o.l))),
      h('button', { class: 'sec', onclick: poser(ferieFerme(iso) ? 'F' : (codeType(n, iso) || 'R')) }, ferieFerme(iso) ? 'Revenir à « férié »' : 'Revenir au planning-type'),
      h('button', { class: 'link', onclick: fermer }, 'Fermer')));
    document.body.append(box);
  };
  const thJour = iso => {
    const d = dDe(iso), hors = moisDe(iso) !== m, pbj = (v.parJour[iso] || []).includes('bad');
    const fer = ferieFerme(iso);
    return h('th', { class: `pj${hors ? ' hors' : ''}${d.getDay() === 0 || d.getDay() === 6 ? ' we' : ''}${pbj ? ' pb' : ''}${iso === auj ? ' auj' : ''}${fer ? ' ferie' : ''}`, title: [fer ? `${fer} (férié, parkings fermés)` : null, ...v.pb.filter(p => p.iso === iso && !p.quinzaine).map(p => p.msg)].filter(Boolean).join('\n') || null },
      h('div', null, JC[d.getDay()]), h('div', null, d.getDate()));
  };
  const td = (iso, n) => {
    const c = G.grille[iso][n], base = G.base[iso][n], modif = normCode(c) !== normCode(base) && G.publieMois, cl = importes.has(iso + '|' + n);
    return h('td', { class: `pc pc-${c || 'vide'}${moisDe(iso) !== m ? ' hors' : ''}${modif ? ' modif' : ''}${cl ? ' claude' : ''}`, onclick: () => choisir(iso, n), title: `${n} — ${jourLong(iso)} : ${libelleCode(c) || '—'}` }, c === 'R' ? '' : c);
  };
  const semaines = []; for (let i = 0; i < G.jours.length; i += 7) semaines.push(G.jours.slice(i, i + 7));
  const table = h('div', { class: 'plan-scroll' }, h('table', { class: 'plan' },
    h('thead', null, h('tr', null, h('th', { class: 'nom' }, ''), ...G.jours.map(thJour))),
    h('tbody', null, ...lignes.map(n => h('tr', { class: G.internes.includes(n) ? 'interne' : null },
      h('th', { class: 'nom' }, n, G.internes.includes(n) ? h('span', { class: 'muted small' }, ' (interne)') : null), ...G.jours.map(iso => td(iso, n)))))));
  // absences : saisie rapide
  const ab = { n: G.agents[0], du: G.jours.find(x => moisDe(x) === m), au: G.jours.find(x => moisDe(x) === m), c: 'CA' };
  const formAbs = h('div', { class: 'card' }, h('h2', null, '1. Absences connues'),
    h('div', { class: 'row' },
      h('select', { onchange: e => ab.n = e.target.value }, ...G.agents.map(n => h('option', { value: n }, n))),
      h('input', { type: 'date', value: ab.du, min: G.jours[0], max: G.jours[G.jours.length - 1], onchange: e => ab.du = e.target.value }),
      h('input', { type: 'date', value: ab.au, min: G.jours[0], max: G.jours[G.jours.length - 1], onchange: e => ab.au = e.target.value })),
    chips(Object.entries(ABSENCES).map(([val, lbl]) => ({ val, lbl })), ab.c, x => ab.c = x),
    h('button', { class: 'sec', onclick: safe(async () => {
      if (!ab.du || !ab.au || ab.au < ab.du) throw new Error('Dates invalides.');
      await majBrouillon(m, b => { for (let iso = ab.du; iso <= ab.au; iso = ajJ(iso, 1)) if (G.jours.includes(iso) && !ferieFerme(iso)) (b.cells[iso] = b.cells[iso] || {})[ab.n] = ab.c; });   // les jours fériés restent fériés
      toast(`${ABSENCES[ab.c]} : ${ab.n}, du ${jourCourt(ab.du)} au ${jourCourt(ab.au)}.`); recharger();
    }) }, 'Ajouter l’absence'),
    h('p', { class: 'muted small' }, 'Ou touchez directement une case du tableau. Les motifs (congés, maladie…) ne sont visibles que des superviseurs ; les agents voient « Absent ».'));
  // Claude
  const consigne = h('textarea', { placeholder: 'Consigne particulière (facultatif). Ex. : Nassiha demande à échanger son samedi 21 avec Saïd.', style: 'min-height:70px' });
  const reponse = h('textarea', { placeholder: 'Collez ici la réponse de Claude (au moins le bloc === PLANNING-CLAUDE … === FIN ===)', style: 'min-height:90px' });
  const sansPb = !v.pb.some(p => p.niveau === 'bad');
  const carteClaude = h('div', { class: 'card' }, h('h2', null, '2. Proposition de Claude (mois compliqué)'),
    sansPb ? h('div', { class: 'banner ok', style: 'margin-bottom:8px' }, 'Aucun problème détecté : pas besoin de Claude, vous pouvez publier directement (étape 3).') : h('p', { class: 'small', style: 'color:var(--bad);margin:0 0 6px' }, 'Des problèmes restent à arbitrer : vous pouvez les corriger à la main dans le tableau, ou demander une proposition à Claude.'),
    consigne,
    h('button', { class: 'sec', onclick: safe(async () => {
      const txt = promptClaude(m, G, consigne.value);
      let copie = false; try { await navigator.clipboard.writeText(txt); copie = true; } catch (e) { }
      const fermer = () => box.remove();
      const box = h('div', { class: 'confirm-contact', onclick: e => { if (e.target === box) fermer(); } }, h('div', { class: 'sheet', style: 'max-width:900px' },
        h('h3', null, copie ? 'Prompt copié' : 'Prompt prêt à copier'),
        h('p', { class: 'small' }, 'Ouvrez votre conversation « planning » dans Claude, collez, envoyez. Puis copiez la réponse de Claude et revenez ici : « Appliquer la réponse de Claude ».'),
        h('textarea', { readonly: true, value: txt, style: 'min-height:320px;font-family:ui-monospace,Consolas,monospace;font-size:12px' }),
        h('button', { class: 'ok', onclick: safe(async () => { await navigator.clipboard.writeText(txt); toast('Copié.'); }) }, 'Copier'),
        h('button', { class: 'link', onclick: fermer }, 'Fermer')));
      document.body.append(box);
    }) }, 'Préparer le prompt pour Claude'),
    reponse,
    h('button', { class: 'sec', onclick: safe(async () => {
      if (!reponse.value.trim()) { try { reponse.value = await navigator.clipboard.readText(); } catch (e) { } }
      const r = lireReponseClaude(reponse.value, G);
      if (r.mois && r.mois !== m && !(await demander({ titre: 'Autre mois ?', texte: `La réponse de Claude concerne ${nomMois(r.mois)}, vous êtes sur ${nomMois(m)}. L’appliquer quand même ?`, oui: 'Appliquer', non: 'Annuler' }))) return;
      const changes = [];
      for (const iso of G.jours) for (const n of r.lignes) if (normCode(r.cells[iso][n]) !== normCode(G.grille[iso][n])) changes.push(iso + '|' + n);
      await majBrouillon(m, b => { for (const iso of G.jours) for (const n of r.lignes) (b.cells[iso] = b.cells[iso] || {})[n] = r.cells[iso][n]; b.claude = changes; });
      if (r.erreurs.length) toast(`Proposition appliquée (${changes.length} case(s) modifiée(s)). ${r.erreurs.length} remarque(s) : ${r.erreurs.slice(0, 2).join(' ; ')}`, true);
      else toast(`Proposition de Claude appliquée : ${changes.length} case(s) modifiée(s), surlignées dans le tableau.`);
      recharger();
    }) }, 'Appliquer la réponse de Claude'));
  // publication
  const notes = h('textarea', { value: G.notes || '', placeholder: 'Notes affichées aux agents sous le planning (ex. : les parkings sont ouverts et gardés).', style: 'min-height:60px', onchange: safe(async e => { const val = e.target.value; await majBrouillon(m, b => { b.notes = val; }); }) });
  const nbBad = v.pb.filter(p => p.niveau === 'bad').length;
  const publier = safe(async () => {
    const notesVal = notes.value.trim();
    const avantPub = G.publieMois ? differences(G.grille, G.base, G.jours, G.agents) : [];
    if (!(await demander({ titre: G.publieMois ? 'Publier les modifications ?' : `Publier le planning de ${nomMois(m)} ?`, oui: 'Publier', non: 'Annuler',
      texte: (nbBad ? `Attention : ${nbBad} problème(s) restent signalés par les contrôles.\n\n` : '') + (G.publieMois ? `${avantPub.length} case(s) modifiée(s) pour les agents.` : 'Les agents le verront dans leur onglet Planning et recevront une notification.') }))) return;
    const jours = {}, motifs = {}, internes = {};
    for (const iso of G.jours) {
      jours[iso] = {};
      for (const n of G.agents) { const c = G.grille[iso][n] || 'R'; jours[iso][n] = ABSENCES[c] ? 'ABS' : c; if (ABSENCES[c]) (motifs[iso] = motifs[iso] || {})[n] = c; }
      for (const n of G.internes) { const c = G.grille[iso][n]; if (c) (internes[iso] = internes[iso] || {})[n] = c; }
    }
    const lib = c => c === 'ABS' ? 'absent' : (libelleCode(c, true) || 'repos').toLowerCase();
    const chg = G.publieMois ? differences(jours, Object.fromEntries(G.jours.map(iso => [iso, Object.fromEntries(G.agents.map(n => [n, (G.pub.jours[iso] || {})[n] || '']))])), G.jours, G.agents) : [];
    const notif = G.publieMois
      ? { titre: `Planning de ${nomMois(m)} modifié`, corps: chg.length ? chg.slice(0, 6).map(x => `${jourCourt(x.iso)} : ${x.n} ${lib(x.b)}`).join(' · ') + (chg.length > 6 ? ` … (${chg.length} changements)` : '') : 'Notes mises à jour.' }
      : { titre: `Planning de ${nomMois(m)} publié`, corps: 'Consultez-le dans l’application : bouton Planning.' };
    const id = uuid(), ts = new Date().toISOString();
    await queueDec('d', { type: 'planning', id, ts, mois: m, du: G.jours[0], au: G.jours[G.jours.length - 1], jours, notes: notesVal, chg: chg.map(x => ({ iso: x.iso, n: x.n, b: x.b })), notif });
    await queueDec('n', { type: 'planning_motifs', id, ts, du: G.jours[0], au: G.jours[G.jours.length - 1], motifs, internes });
    await kvUpdate('planningBrouillons', a => { const x = a || {}; delete x[m]; return x; });
    await fold(); syncNow();
    toast(`Planning de ${nomMois(m)} publié.`); recharger();
  });
  const statut = G.publieMois
    ? (diffs.length || (G.brouillon && G.brouillon.notes != null && G.brouillon.notes !== (G.pub.notes[m] || '')) ? h('span', { class: 'badge b-warn' }, `publié le ${fmtDT(G.publieMois.ts)} — ${diffs.length} modification(s) non publiée(s)`) : h('span', { class: 'badge b-ok' }, `publié le ${fmtDT(G.publieMois.ts)}`))
    : h('span', { class: 'badge b-bad' }, 'pas encore publié');
  const heures = h('div', { class: 'card' }, h('h2', null, 'Heures par quinzaine'),
    h('table', { class: 'plan-h' }, h('tr', null, h('th', null, ''), ...v.quinz.map(q => h('th', null, `du ${jourCourt(q)}`))),
      ...G.agents.map(n => h('tr', null, h('th', null, n), ...(v.heures[n] || []).map(x => h('td', { class: x.h === reglesPlanning().heures ? 'ok' : 'bad' }, `${String(x.h).replace('.', ',')} h`))))));
  const listePb = v.pb.filter(p => p.niveau === 'bad');
  const infos = v.pb.filter(p => p.niveau === 'info');
  return page('Planning', { back: true },
    h('div', { class: 'card plan-tete' },
      h('div', { class: 'row', style: 'align-items:center' },
        h('button', { class: 'sec', style: 'width:auto', onclick: () => go('planning', { mois: moisSuivant(m, -1) }, { noPush: true }) }, '‹'),
        h('h2', { style: 'margin:0;flex:1;text-align:center;text-transform:none;font-size:22px' }, nomMois(m)),
        h('button', { class: 'sec', style: 'width:auto', onclick: () => go('planning', { mois: moisSuivant(m, 1) }, { noPush: true }) }, '›')),
      h('p', { class: 'muted small', style: 'text-align:center;margin:4px 0' }, `Du ${jourLong(G.jours[0])} au ${jourLong(G.jours[G.jours.length - 1])} (semaines commençant en ${MOIS_LONGS[Number(m.slice(5)) - 1]}) · `, statut)),
    formAbs,
    table,
    h('div', { class: 'plan-legende' }, ...vacationsP().map(x => h('span', { class: `pc pc-${x.code}` }, `${x.code} ${x.debut}-${x.fin}`)), h('span', { class: 'pc pc-R' }, 'Repos (vide)'), h('span', { class: 'pc pc-F' }, 'F Férié (fermé)'), ...Object.entries(ABSENCES).map(([c, l]) => h('span', { class: `pc pc-${c}` }, `${c} ${l}`)), h('span', { class: 'pc claude-ex' }, 'proposé par Claude'), h('span', { class: 'pc modif-ex' }, 'modifié depuis la publication')),
    h('div', { class: 'card', style: listePb.length ? 'border-color:var(--bad)' : 'border-color:var(--ok)' }, h('h2', null, `Contrôles${listePb.length ? ` : ${listePb.length} problème(s)` : ' : rien à signaler'}`),
      ...listePb.map(p => h('div', { class: 'small', style: 'color:var(--bad)' }, `${p.quinzaine ? '' : jourLong(p.iso) + ' : '}${p.msg}`)),
      infos.length ? h('details', null, h('summary', { class: 'small muted' }, `${infos.length} information(s) : jours de repos du planning-type travaillés`), ...infos.map(p => h('div', { class: 'small muted' }, `${jourLong(p.iso)} : ${p.msg}`))) : null),
    heures,
    carteClaude,
    h('div', { class: 'card' }, h('h2', null, '3. Publication'), notes,
      h('button', { class: 'ok big', onclick: publier }, G.publieMois ? 'Publier les modifications aux agents' : `Publier le planning de ${nomMois(m)}`),
      G.brouillon ? h('button', { class: 'link', onclick: safe(async () => { if (!(await demander({ titre: 'Abandonner le brouillon ?', texte: 'Le tableau revient au planning publié (ou au planning-type).', oui: 'Oui, abandonner', non: 'Non', danger: true }))) return; await kvUpdate('planningBrouillons', a => { const x = a || {}; delete x[m]; return x; }); recharger(); }) }, 'Abandonner les modifications non publiées') : null),
    h('div', { class: 'row' },
      h('button', { class: 'sec', onclick: safe(() => exporterPlanning(m, G)) }, 'Excel'),
      G.publieMois ? h('button', { class: 'sec', onclick: safe(() => telechargerPlanningPDF(m)) }, 'PDF (version agents)') : null,
      h('button', { class: 'sec', onclick: () => go('planningParam') }, 'Planning-type et règles')));
};
async function exporterPlanning(m, G) {
  const L = [...G.agents, ...G.internes];
  const aoa = [['Agent', ...G.jours.map(iso => jourCourt(iso))], ...L.map(n => [n, ...G.jours.map(iso => G.grille[iso][n] === 'R' ? '' : G.grille[iso][n] || '')])];
  aoa.push([], ['Codes', ...vacationsP().map(v => `${v.code} ${v.debut}-${v.fin}`), 'F férié (fermé)', 'CA congés', 'MAL maladie', 'FOR formation', 'ABS absence']);
  await shareOrDownload(xlsxBlob([[`Planning ${m}`, aoa, [18, ...G.jours.map(() => 6)]]]), `planning_${m}.xlsx`, `Planning ${nomMois(m)}`);
}

/* Paramètres : planning-type, règles, horaires */
VIEWS.planningParam = async ({ prefill } = {}) => {
  const c = JSON.parse(JSON.stringify(state.cfg)); c.planning = { ...PLANNING_DEFAUT, ...(c.planning || {}) };
  if (prefill) { c.planning.type = prefill.type; if (prefill.regles) c.planning.regles = prefill.regles; if (prefill.ancre) c.planning.ancre = prefill.ancre; }
  const tType = h('textarea', { value: c.planning.type, style: 'min-height:160px;font-family:ui-monospace,Consolas,monospace;font-size:14px', placeholder: 'Agent 1 : R MAT MAT MAT MAT SOI MAT | R MAT MAT MAT MAT R R' });
  const tRegles = h('textarea', { value: c.planning.regles, style: 'min-height:140px' });
  const ancre = h('input', { type: 'date', value: c.planning.ancre });
  const vacs = c.planning.vacations.map(v => ({ v, lbl: h('input', { type: 'text', value: v.lbl }), d: h('input', { type: 'time', value: v.debut }), f: h('input', { type: 'time', value: v.fin }) }));
  return page('Planning-type et règles', { back: () => go('planning') },
    prefill ? h('div', { class: 'banner warn' }, 'Planning-type et règles reçus par lien : vérifiez, puis « Enregistrer » en bas de page.') : null,
    h('div', { class: 'card' }, h('h2', null, 'Planning-type (cycle de 2 semaines)'),
      h('p', { class: 'muted small' }, 'Une ligne par agent : nom, deux-points, 7 codes de la semaine 1 (lundi → dimanche), « | », 7 codes de la semaine 2. R = repos. Les noms doivent correspondre à ceux de la liste des agents.'),
      tType,
      h('label', { class: 'f' }, 'Lundi d’une semaine 1 (début du cycle et des quinzaines)'), ancre),
    h('div', { class: 'card' }, h('h2', null, 'Règles contrôlées'),
      h('p', { class: 'muted small' }, 'Une règle par ligne : « couverture : 08:00-20:00 » · « heures par quinzaine : 70 » · « jamais sur la même vacation : Nom1, Nom2 » · « jamais le dimanche : Nom » · « lignes internes : Vacataire » (lignes du tableau non montrées aux agents).'),
      tRegles),
    h('div', { class: 'card' }, h('h2', null, 'Vacations'),
      ...vacs.map(x => h('div', { class: 'row', style: 'align-items:center' }, h('b', { style: 'width:52px;flex:none' }, x.v.code), x.lbl, x.d, x.f))),
    h('button', { class: 'ok big', onclick: safe(async () => {
      c.planning.type = tType.value.trim(); c.planning.regles = tRegles.value.trim(); c.planning.ancre = ancre.value || c.planning.ancre;
      c.planning.vacations = vacs.map(x => ({ code: x.v.code, lbl: x.lbl.value.trim() || x.v.lbl, debut: x.d.value || x.v.debut, fin: x.f.value || x.v.fin }));
      const anc = dDe(c.planning.ancre); if (anc.getDay() !== 1) throw new Error('La date de début de cycle doit être un lundi.');
      const ancien = state.cfg.planning; state.cfg.planning = c.planning;
      const t = typePlanning(), lignesBrutes = c.planning.type.split('\n').filter(l => l.trim());
      state.cfg.planning = ancien;
      if (t.length !== lignesBrutes.length) throw new Error('Planning-type : chaque ligne doit contenir un nom, « : », puis 14 codes.');
      const ok = new Set([...c.planning.vacations.map(v => v.code), 'R']);
      const mauvais = t.flatMap(x => x.codes).filter(x => !ok.has(x));
      if (mauvais.length) throw new Error(`Codes inconnus dans le planning-type : ${[...new Set(mauvais)].join(', ')}`);
      c.cfgId = randCode(8);
      await publishCfg(c);
      toast('Planning-type et règles enregistrés.'); go('planning');
    }) }, 'Enregistrer'));
};

/* Agents : onglet Planning (lecture seule, sans motif d'absence) */
async function equipeDuJour(iso) {
  const pub = await planningPublie(), j = pub.jours[iso];
  if (!j) return null;
  const parVac = vacationsP().map(v => ({ v, noms: Object.entries(j).filter(([, c]) => c === v.code).map(([n]) => n) }));
  const vals = Object.values(j);
  return { parVac, ferie: vals.length && vals.every(c => c === 'F') ? (joursFeries(Number(iso.slice(0, 4))).get(iso) || 'Jour férié') : null, absents: Object.entries(j).filter(([, c]) => c === 'ABS').map(([n]) => n), repos: Object.entries(j).filter(([, c]) => c === 'R' || !c || c === 'F').map(([n]) => n) };
}
VIEWS.planningAgents = async ({ semaine, agent } = {}) => {
  const pub = await planningPublie(), auj = localDate();
  const lundi = semaine || ajJ(auj, -((dDe(auj).getDay() + 6) % 7));
  const derniere = pub.pubs.slice(-1)[0];
  await kvSet('planningVu', derniere ? derniere.id : null);
  const nav = p => go('planningAgents', { semaine: lundi, agent, ...p }, { noPush: true });
  const noms = [...new Set(Object.values(pub.jours).flatMap(j => Object.keys(j)))];
  const jours = []; for (let i = 0; i < 7; i++) jours.push(ajJ(lundi, i));
  const carteJour = async iso => {
    const e = await equipeDuJour(iso);
    return h('div', { class: 'card pa-jour' + (iso === auj ? ' auj' : '') },
      h('h3', null, jourLong(iso).replace(/^./, x => x.toUpperCase()), iso === auj ? h('span', { class: 'badge b-info' }, 'aujourd’hui') : null),
      !e ? h('p', { class: 'muted' }, 'Pas encore de planning publié.') : e.ferie ? h('div', { class: 'pa-l pc-F' }, h('b', null, `${e.ferie} : jour férié, parkings fermés`)) : [
        ...e.parVac.map(x => h('div', { class: `pa-l pc-${x.v.code}` }, h('span', { class: 'pa-v' }, `${x.v.lbl} ${x.v.debut.replace(':00', 'h').replace(':', 'h')}-${x.v.fin.replace(':00', 'h').replace(':', 'h')}`), h('b', null, x.noms.join(', ') || '—'))),
        e.absents.length ? h('div', { class: 'small muted' }, `Absent : ${e.absents.join(', ')}`) : null,
        e.repos.length ? h('div', { class: 'small muted' }, `Repos : ${e.repos.join(', ')}`) : null]);
  };
  let corps;
  if (agent) {
    const m = moisDe(lundi), { jours: js } = periodeMois(m);
    corps = [h('div', { class: 'card' }, h('h2', null, `${agent} — ${nomMois(m)}`),
      ...js.filter(iso => moisDe(iso) === m).map(iso => { const c = (pub.jours[iso] || {})[agent]; return h('div', { class: 'stat' + (iso === auj ? ' auj' : '') }, h('span', null, jourLong(iso)), h('b', { class: c ? `pa-c pc-${c}` : '' }, c ? (c === 'ABS' ? 'Absent' : c === 'F' ? 'Férié' : libelleCode(c, true) || 'Repos') : '—')); }))];
  } else corps = await Promise.all(jours.map(carteJour));
  const m = moisDe(lundi);
  const moisPub = [...new Set(pub.pubs.map(x => x.mois))].sort();
  const moisPdf = moisPub.includes(m) ? m : moisPub.filter(x => x <= m).pop() || moisPub[0];
  return page('Planning', { back: true },
    moisPdf ? h('button', { class: 'sec', onclick: safe(() => telechargerPlanningPDF(moisPdf)) }, `Télécharger le planning de ${nomMois(moisPdf)} (PDF)`) : null,
    derniere ? h('div', { class: 'card', style: 'border-color:var(--navy)' }, h('b', null, `Mise à jour du ${fmtDT(derniere.ts)}`),
      derniere.chg && derniere.chg.length ? h('div', { class: 'small' }, derniere.chg.slice(0, 8).map(x => `${jourCourt(x.iso)} : ${x.n} ${x.b === 'ABS' ? 'absent' : (libelleCode(x.b, true) || 'repos').toLowerCase()}`).join(' · ')) : null) : null,
    h('div', { class: 'row', style: 'align-items:center' },
      h('button', { class: 'sec', style: 'width:auto', onclick: () => nav({ semaine: ajJ(lundi, agent ? -28 : -7) }) }, '‹'),
      h('b', { style: 'flex:1;text-align:center' }, agent ? nomMois(m) : `${dDe(lundi).getDate()} → ${jourLong(ajJ(lundi, 6)).replace(/^\S+ /, '')}`),
      h('button', { class: 'sec', style: 'width:auto', onclick: () => nav({ semaine: ajJ(lundi, agent ? 35 : 7) }) }, '›')),
    noms.length ? chips([{ val: '', lbl: 'Toute l’équipe' }, ...noms.map(n => ({ val: n, lbl: n }))], agent || '', x => nav({ agent: x || undefined })) : null,
    ...corps,
    pub.notes[m] ? h('div', { class: 'card' }, h('h2', null, 'Notes'), h('p', { style: 'white-space:pre-wrap' }, pub.notes[m])) : null);
};
/* PDF du planning publié, au format du tableau papier (une grille par semaine) — sans motif d'absence */
async function pdfPlanning(m) {
  const pub = await planningPublie();
  const p = pub.pubs.filter(x => x.mois === m).slice(-1)[0];
  if (!p) throw new Error(`Le planning de ${nomMois(m)} n'est pas encore publié.`);
  const { jsPDF } = await chargerScript('lib/jspdf.umd.min.js', 'jspdf');
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  const txt = t => String(t == null ? '' : t).replace(/[\u202f\u00a0]/g, ' ').replace(/[^\x00-\xff‘’“”–—•…€Œœ]/g, '?');
  const { jours } = periodeMois(m);
  const noms = [...new Set(jours.flatMap(iso => Object.keys(pub.jours[iso] || {})))];
  const ordre = typePlanning().map(t => t.nom); noms.sort((a, b) => (ordre.indexOf(a) + 1 || 99) - (ordre.indexOf(b) + 1 || 99));
  const M = 12, W = 210, colN = 44, colJ = (W - 2 * M - colN) / 7, hl = 7;
  const jourNum = d => d.getDate() === 1 ? '1er' : String(d.getDate());
  const datePetite = iso => { const d = dDe(iso); return `${jourNum(d)} ${MOIS[d.getMonth()]}`; };
  const dateLongue = iso => { const d = dDe(iso); return `${jourNum(d)} ${MOIS_LONGS[d.getMonth()]}`; };
  const STY = { MAT: [[0, 112, 192], [255, 255, 255], 'MATIN'], MID: [[112, 48, 160], [255, 255, 255], 'MIDI'], SOI: [[220, 0, 0], [255, 255, 255], 'SOIR'],
    R: [[242, 242, 242], [128, 128, 128], 'REPOS'], ABS: [[255, 255, 255], [90, 96, 105], 'ABSENT'], F: [[230, 233, 238], [90, 96, 105], 'FÉRIÉ'] };
  const vacs = vacationsP();
  let y = 16;
  doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.setTextColor(0, 0, 0);
  doc.text(txt('PARKINGS MUNICIPAUX — PLANNING DES AGENTS DE SURVEILLANCE'), M, y); y += 6;
  doc.setFontSize(11); doc.text(txt(`Du ${NOMS_JOURS[dDe(jours[0]).getDay()]} ${dateLongue(jours[0])} au ${NOMS_JOURS[dDe(jours[jours.length - 1]).getDay()]} ${dateLongue(jours[jours.length - 1])} ${jours[jours.length - 1].slice(0, 4)}`), M, y); y += 7;
  const cell = (x, yy, w, hh, fond, couleur, texte, gras, taille) => {
    if (fond) { doc.setFillColor(...fond); doc.rect(x, yy, w, hh, 'F'); }
    doc.setDrawColor(190, 190, 190); doc.setLineWidth(0.2); doc.rect(x, yy, w, hh);
    if (texte) { doc.setFont('helvetica', gras ? 'bold' : 'normal'); doc.setFontSize(taille || 9); doc.setTextColor(...(couleur || [0, 0, 0])); doc.text(txt(texte), x + w / 2, yy + hh / 2 + 1.3, { align: 'center' }); }
  };
  for (let i = 0; i < jours.length; i += 7) {
    const sem = jours.slice(i, i + 7);
    const hBloc = hl * (3 + noms.length) + 5;
    if (y + hBloc > 275) { doc.addPage(); y = 16; }
    const d0 = dDe(sem[0]), d6 = dDe(sem[6]);
    const titreSem = d0.getMonth() === d6.getMonth() ? `Semaine du ${jourNum(d0)} au ${dateLongue(sem[6])}` : `Semaine du ${dateLongue(sem[0])} au ${dateLongue(sem[6])}`;
    cell(M + colN, y, colJ * 7, hl, [217, 217, 217], [0, 0, 0], titreSem, true, 10); y += hl;
    sem.forEach((iso, k) => cell(M + colN + k * colJ, y, colJ, hl - 1, null, [0, 0, 0], NOMS_JOURS[dDe(iso).getDay()].replace(/^./, c => c.toUpperCase()), true, 9)); y += hl - 1;
    sem.forEach((iso, k) => cell(M + colN + k * colJ, y, colJ, hl - 1, ferieFerme(iso) ? [230, 233, 238] : null, [0, 0, 0], datePetite(iso), true, 10)); y += hl - 1;
    for (const n of noms) {
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(0, 0, 0);
      doc.text(txt(n), M + 1, y + hl / 2 + 1.3);
      sem.forEach((iso, k) => { const c = (pub.jours[iso] || {})[n] || 'R'; const st = STY[c] || STY.R; cell(M + colN + k * colJ, y, colJ, hl, st[0], st[1], st[2], c !== 'R' && c !== 'F' && c !== 'ABS', c === 'R' ? 8.5 : 10); });
      y += hl;
    }
    y += 5;
  }
  if (y + 30 > 285) { doc.addPage(); y = 16; }
  doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(0, 0, 0); doc.text('LÉGENDE', M, y + 4); y += 7;
  const leg = [...vacs.map(v => [v.code, `${v.debut.replace(':', 'h')} – ${v.fin.replace(':', 'h')}`]), ['R', 'Jour de repos'], ['F', 'Jour férié : parkings fermés'], ['ABS', 'Absent']];
  for (const [c, l] of leg) { const st = STY[c] || STY.R; cell(M + colN, y, colJ, 6, st[0], st[1], st[2], c !== 'R' && c !== 'F' && c !== 'ABS', 9); doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor(0, 0, 0); doc.text(txt(l), M + colN + colJ + 3, y + 4.3); y += 6.5; }
  if (pub.notes[m]) { y += 3; doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.splitTextToSize(txt(pub.notes[m]), W - 2 * M).forEach(l => { if (y > 285) { doc.addPage(); y = 16; } doc.text(l, M, y); y += 5; }); }
  doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(120, 120, 120);
  doc.text(txt(`Planning validé le ${fmtDT(p.ts)}`), M, 292);
  return { blob: doc.output('blob'), nom: `planning_${m}.pdf` };
}
async function telechargerPlanningPDF(m) {
  showBusy('Préparation du PDF…');
  try { const { blob, nom } = await pdfPlanning(m); hideBusy(); await shareOrDownload(blob, nom, `Planning ${nomMois(m)}`); }
  finally { hideBusy(); }
}
async function cartePlanningAgents() {
  const pub = await planningPublie();
  if (!pub.pubs.length) return null;
  const e = await equipeDuJour(localDate());
  const nouveau = (await kvGet('planningVu')) !== pub.pubs[pub.pubs.length - 1].id;
  return h('button', { class: 'list-item', style: nouveau ? 'border:2px solid var(--accent)' : null, onclick: () => go('planningAgents', {}) },
    h('div', { class: 'l1' }, 'Planning', nouveau ? h('span', { class: 'badge b-warn' }, 'mis à jour') : null),
    e ? h('div', { class: 'small' }, e.ferie ? `${e.ferie} : jour férié, parkings fermés` : e.parVac.map(x => `${x.v.lbl} : ${x.noms.join(', ') || '—'}`).join(' · ')) : h('div', { class: 'muted small' }, 'Pas de planning aujourd’hui.'));
}
async function cartePlanningSup() {
  const auj = localDate(), m = moisDe(auj), ms = moisSuivant(m), pub = await planningPublie();
  const br = (await kvGet('planningBrouillons')) || {};
  const etat = mm => { const p = pub.pubs.filter(x => x.mois === mm).slice(-1)[0]; return p ? `publié le ${fmtJour(p.ts.slice(0, 10))}` : br[mm] ? 'brouillon en cours' : 'pas encore préparé'; };
  const aFaire = Number(auj.slice(8)) >= 15 && !pub.pubs.some(x => x.mois === ms);
  return h('div', { class: 'card', style: aFaire ? 'border-color:var(--warn)' : null }, h('h2', null, 'Planning des agents'),
    aFaire ? h('p', { style: 'color:var(--warn);font-weight:700;margin:0 0 6px' }, `Planning de ${nomMois(ms)} à préparer.`) : null,
    h('button', { class: 'list-item', onclick: () => go('planning', { mois: m }) }, h('div', { class: 'stat' }, h('span', null, nomMois(m)), h('span', { class: 'muted small' }, etat(m)))),
    h('button', { class: 'list-item', onclick: () => go('planning', { mois: ms }) }, h('div', { class: 'stat' }, h('span', null, nomMois(ms)), h('span', { class: 'muted small' }, etat(ms)))));
}
/* Rappel du 15 : notification programmée sur le relais (alerte différée), refaite chaque mois */
async function programmerRappelPlanning(abos) {
  const auj = new Date(), m = moisDe(localDate());
  const pub = await planningPublie();
  let cible = moisSuivant(m), quand = new Date(auj.getFullYear(), auj.getMonth(), 15, 9, 0);
  if (quand <= auj || pub.pubs.some(x => x.mois === cible)) { cible = moisSuivant(m, 2); quand = new Date(auj.getFullYear(), auj.getMonth() + 1, 15, 9, 0); }
  const id = await jeton('planning', cible);
  const empreinte = JSON.stringify([cible, (abos || []).map(a => [a.e, a.prefs])]);
  const prev = await kvGet('rappelPlanning');
  if (prev && prev.empreinte === empreinte) return;
  const messages = await chiffrerPour(abos, 'rappel_planning', { type: 'rappel_planning', titre: `Planning de ${nomMois(cible)} à préparer`, corps: 'Saisissez les absences connues, puis préparez la proposition avec Claude.', lien: '#planning' });
  if (prev && prev.id && prev.id !== id) { try { await rel('annulerDiffere', { id: prev.id }, 15000); } catch (e) { } }
  if (Object.keys(messages).length) await rel('differer', { id, apres: quand.toISOString(), messages }, 20000);
  await kvSet('rappelPlanning', { empreinte, id, cible, at: Date.now() });
}
/* ====================================================================== */
/* TRAITEMENT DES SIGNALEMENTS (dispatch vers les services et prestataires) */
/* ====================================================================== */
const ACTIONS = {
  tel: 'Appel ou message téléphonique', mail: 'Courriel envoyé', autre: 'Transmis autrement', relance: 'Relance',
  reponse: 'Réponse reçue', prevue: 'Intervention prévue', faite: 'Intervention réalisée', note: 'Note', echec: 'Appel sans réponse'
};
const SUIVIS = ['À transmettre', 'En attente de réponse', 'Réponse reçue', 'Intervention prévue', 'Intervention réalisée'];
function suiviEtat(s) {
  const a = (s.actions || []).filter(x => x.type !== 'note' && x.type !== 'echec');
  const reste = destsATransmettre(s).length;
  if (reste && !a.some(x => ['reponse', 'prevue', 'faite'].includes(x.type))) return 'À transmettre';   // tant qu'un service n'a pas été prévenu
  if (!a.length) return 'En attente de réponse';
  return { reponse: 'Réponse reçue', prevue: 'Intervention prévue', faite: 'Intervention réalisée' }[a[a.length - 1].type] || 'En attente de réponse';
}
function suiviBadge(s) {
  if (s.statut === 'clos') return h('span', { class: 'badge b-ok' }, 'Clos');
  const e = suiviEtat(s);
  const cls = e === 'À transmettre' ? 'b-bad' : e === 'Intervention réalisée' ? 'b-ok' : e === 'Intervention prévue' ? 'b-info' : 'b-warn';
  const p = (s.actions || []).filter(x => x.type === 'prevue').pop();
  return h('span', { class: 'badge ' + cls }, e + (e === 'Intervention prévue' && p && p.prevu ? ' le ' + fmtDT(p.prevu) : ''));
}
function texteSignalement(s) {
  return [`Signalement ${s.ref} — parking ${siteNom(s.site)}`, sigTitre(s), s.desc ? 'Détail : ' + s.desc : null,
    s.plaque ? 'Plaque : ' + s.plaque : null, s.emplacement ? 'Emplacement : ' + s.emplacement : null, s.vehicule ? 'Véhicule : ' + s.vehicule : null,
    `Signalé le ${fmtJour(s.jour)}${s.agent ? ' par ' + s.agent : ''}.`, '', 'Ville de Cachan — surveillance des parkings'].filter(x => x !== null).join('\n');
}
/* Photos d'une liste (téléchargées depuis le relais au besoin) */
async function photosDe(list, opts = {}) {
  const miss = new Set();
  for (const p of list || []) if (p.n && !(await dbGet('photos', p.id))) miss.add(p.n);
  const erreurs = [];
  if (miss.size && opts.rapide && !navigator.onLine) erreurs.push('pas de réseau');
  else for (const n of miss) { try { await fetchMedia(n, opts.rapide ? 20000 : 180000); } catch (e) { erreurs.push(e.message); } }
  const out = [];
  for (const p of list || []) { const c = await dbGet('photos', p.id); if (c) out.push({ ...c, attendu: p.sha }); }
  return { photos: out, erreurs };
}
async function partagerSig(s) {
  const txt = texteSignalement(s);
  const { photos } = await photosDe([...(s.photos || [])]);
  const files = photos.map((p, i) => new File([p.blob], `${s.ref}-${i + 1}.jpg`, { type: 'image/jpeg' }));
  const data = files.length ? { title: `Signalement ${s.ref}`, text: txt, files } : { title: `Signalement ${s.ref}`, text: txt };
  if (navigator.share && (!navigator.canShare || navigator.canShare(data))) {
    toast('Choisissez votre messagerie (Gmail, Outlook…).');
    try { await navigator.share(data); return true; } catch (e) { if (e.name === 'AbortError') return false; }
  }
  try { await navigator.clipboard.writeText(txt); } catch (e) { }
  files.forEach(f => download(f, f.name));
  toast('Texte copié' + (files.length ? ' et photos téléchargées' : '') + ' : collez-les dans votre courriel.');
  return true;
}
function journalActions(s) {
  const a = s.actions || [];
  if (!a.length) return [h('p', { class: 'muted' }, 'Aucune action enregistrée.')];
  return a.slice().reverse().map(x => h('div', { class: 'stat', style: 'display:block' },
    h('div', null, h('b', null, ACTIONS[x.type] || x.type), x.dest ? ' — ' + x.dest : '', x.type === 'prevue' && x.prevu ? ` — le ${fmtDT(x.prevu)}` : ''),
    x.texte ? h('div', { class: 'small' }, x.texte) : null,
    h('div', { class: 'muted small' }, `${x.quand ? fmtDT(x.quand) : fmtJour(x.jour)}${x.par ? ' — ' + x.par : ''}`)));
}
function contactsBloc(s) {
  const dests = s.dests || [];
  const sujet = `[Parking ${siteNom(s.site)}] ${sigTitre(s)} — ${s.ref}`;
  const corps = encodeURIComponent(texteSignalement(s));
  const avecMail = dests.filter(dn => ((state.cfg.contacts || {})[dn] || {}).mail);
  return h('div', { class: 'card' }, h('h2', null, 'Transmettre'),
    ...dests.map(dn => {
      const c = (state.cfg.contacts || {})[dn] || {};
      return h('div', { class: 'stat', style: 'gap:8px;flex-wrap:wrap' }, h('b', null, dn),
        h('span', { style: 'display:flex;gap:6px;flex-wrap:wrap' },
          c.tel ? h('a', { class: 'btn sec mini', href: 'tel:' + c.tel.replace(/\s/g, ''), onclick: () => contactLance(s, 'tel', [dn]) }, 'Appeler ' + c.tel) : null,
          c.mail ? h('a', { class: 'btn sec mini', href: `mailto:${c.mail}?subject=${encodeURIComponent(sujet)}&body=${corps}`, onclick: () => contactLance(s, 'mail', [dn]) }, 'Courriel') : null,
          !c.tel && !c.mail ? h('span', { class: 'muted small' }, 'coordonnées non renseignées') : null));
    }),
    avecMail.length > 1 ? h('a', { class: 'btn sec', href: `mailto:${avecMail.map(dn => state.cfg.contacts[dn].mail).join(',')}?subject=${encodeURIComponent(sujet)}&body=${corps}`, onclick: () => contactLance(s, 'mail', avecMail) }, `Un seul courriel à tous (${avecMail.join(', ')})`) : null,
    h('button', { class: 'sec', onclick: safe(async () => { if (await partagerSig(s)) contactLance(s, 'mail', null); }) }, 'Courriel avec les photos'),
    h('p', { class: 'muted small' }, 'Après l’appel ou l’envoi, l’application vous demande de le confirmer : l’action est alors enregistrée.'));
}

/* Confirmation au retour d'un appel ou d'un courriel : « avez-vous appelé… ? » */
async function contactLance(s, type, dests) {
  await kvSet('contactAttente', { sig: s.id, ref: s.ref, type, dests, at: Date.now() });
  setTimeout(confirmerContact, 2500);
}
let _confirmOuverte = false;
async function confirmerContact() {
  if (_confirmOuverte || document.visibilityState !== 'visible' || !state.cfg) return;
  const c = await kvGet('contactAttente');
  if (!c) return;
  if (Date.now() - c.at > 12 * 3600000) { await kvDel('contactAttente'); return; }
  if (c.mc) return confirmerMC(c);
  const s = await dbGet('sigs', c.sig);
  if (!s) { await kvDel('contactAttente'); return; }
  _confirmOuverte = true;
  const equipe = state.mode !== 'superviseur';
  const choisis = c.dests ? [...c.dests] : [];
  const d = { par: equipe ? await agentPrefere() : null, texte: '' };
  const qui = c.dests ? c.dests.join(', ') : null;
  const question = c.type === 'tel' ? `Avez-vous pu joindre ${qui} ?` : qui ? `Avez-vous envoyé le courriel à ${qui} ?` : 'Avez-vous envoyé le courriel avec les photos ? À qui ?';
  const fermer = async () => { await kvDel('contactAttente'); box.remove(); _confirmOuverte = false; };
  const enregistrer = (type, suffixe) => safe(async () => {
    if (!choisis.length) throw new Error('Indiquez à qui.');
    if (equipe && !d.par) throw new Error('Indiquez qui a appelé ou écrit.');
    const texte = [suffixe, d.texte.trim()].filter(Boolean).join(' — ');
    for (const dest of choisis) {
      const a = { sig: s.id, ref: s.ref, type, dest, texte, prevu: null, quand: new Date().toISOString() };
      if (equipe) await appendEvent('action', { site: s.site, agent: d.par }, a);
      else await queueDec('d', { type: 'action', id: uuid(), ts: new Date().toISOString(), par: 'DPMS', action: a });
    }
    await fermer(); await fold(); syncNow();
    toast('Action enregistrée dans le traitement.');
    // agents : appel passé ou courriel envoyé → retour à l'accueil ; sans réponse → on reste sur la fiche pour relancer
    if (equipe && !suffixe) go('home'); else render();
  });
  const box = h('div', { class: 'confirm-contact' }, h('div', { class: 'sheet' },
    h('h3', null, c.type === 'tel' ? 'Appel' : 'Courriel', ` — ${s.ref}`),
    h('p', { class: 'muted small' }, sigTitre(s)),
    h('p', { style: 'font-weight:600' }, question),
    c.dests ? null : chips(s.dests || [], choisis, () => { }, { multi: true }),
    equipe ? [h('label', { class: 'f' }, 'Qui ?'), chips(state.cfg.agents, d.par, v => d.par = v)] : null,
    h('textarea', { placeholder: c.type === 'tel' ? 'Ce qui a été dit ou convenu (facultatif)' : 'Précisions (facultatif)', oninput: e => d.texte = e.target.value }),
    c.type === 'tel'
      ? [h('button', { class: 'ok', onclick: enregistrer('tel') }, 'Oui, appel passé'), h('button', { class: 'sec', onclick: enregistrer('tel', 'pas de réponse') }, 'Pas de réponse (à relancer)')]
      : h('button', { class: 'ok', onclick: enregistrer('mail') }, 'Oui, courriel envoyé'),
    h('button', { class: 'link', onclick: safe(fermer) }, c.type === 'tel' ? 'Non, je n’ai pas appelé' : 'Non, pas envoyé')));
  document.body.append(box);
}
/* Formulaire d'action ; agents = liste des noms (équipe) ou null (superviseur) */
function actionForm(s, { onSave, onClore, agents, parDefaut }) {
  const d = { type: null, dest: (s.dests || [])[0] || null, texte: '', prevu: localDT(new Date(Date.now() + 86400000)), par: parDefaut || null };
  const prevuBox = h('div', { style: 'display:none' }, h('label', { class: 'f' }, 'Prévue le'),
    h('input', { type: 'datetime-local', value: d.prevu, oninput: e => d.prevu = e.target.value }));
  const destsOpts = [...new Set([...(s.dests || []), ...state.cfg.dests])];
  const typeBox = chips(Object.entries(ACTIONS).map(([val, lbl]) => ({ val, lbl })), d.type, v => { d.type = v; prevuBox.style.display = v === 'prevue' ? '' : 'none'; });
  const destBox = chips(destsOpts, d.dest, v => d.dest = v);
  const texte = h('textarea', { placeholder: 'Ex. : demande envoyée à M. X (ticket 1234) ; réponse : passage jeudi matin', oninput: e => d.texte = e.target.value });
  const box = h('div', { class: 'card' }, h('h2', null, 'Ajouter une action'),
    h('label', { class: 'f' }, 'Action'), typeBox,
    h('label', { class: 'f' }, 'Auprès de'), destBox,
    prevuBox,
    h('label', { class: 'f' }, 'Détail'), texte,
    agents ? [h('label', { class: 'f' }, 'Par'), chips(agents, d.par, v => d.par = v)] : null,
    h('button', {
      class: 'ok', onclick: safe(async () => {
        if (!d.type) throw new Error('Choisissez l’action.');
        if (d.type !== 'note' && !d.dest) throw new Error('Indiquez auprès de qui.');
        if (d.type === 'note' && !d.texte.trim()) throw new Error('Saisissez la note.');
        if (agents && !d.par) throw new Error('Indiquez qui enregistre l’action.');
        const prevu = d.type === 'prevue' ? new Date(d.prevu) : null;
        if (prevu && isNaN(prevu)) throw new Error('Date d’intervention invalide.');
        await onSave({ sig: s.id, ref: s.ref, type: d.type, dest: d.type === 'note' ? '' : d.dest, texte: d.texte.trim(), prevu: prevu ? prevu.toISOString() : null, quand: new Date().toISOString() }, d.par);
      })
    }, 'Enregistrer l’action'),
    s.statut === 'ouvert' && onClore ? h('button', {
      class: 'link', onclick: safe(async () => {
        if (agents && !d.par) throw new Error('Indiquez d’abord qui clôt (« Par »).');
        if (!confirm('Le problème est résolu : clore le signalement ?')) return;
        await onClore(d.par, d.texte.trim());
      })
    }, 'Problème résolu : clore le signalement') : null);
  box.preset = (type, dest) => { d.type = type; typeBox.set(type); if (dest) { d.dest = dest; destBox.set(dest); } prevuBox.style.display = type === 'prevue' ? '' : 'none'; setTimeout(() => { box.scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 300); };
  return box;
}
function galerie(photos, grande) {
  if (!photos.length) return null;
  const urls = photos.map(p => URL.createObjectURL(p.blob));
  return h('div', null,
    photos.some(p => p.attendu && p.sha !== p.attendu) ? h('div', { class: 'banner bad' }, 'Une photo a été modifiée après la saisie (empreinte différente).') : null,
    h('div', { class: 'gallery' + (grande ? ' grande' : '') }, ...urls.map((u, i) => h('img', { src: u, alt: 'Photo', onclick: () => lightbox(urls, i) }))));
}
function lightbox(urls, i) {
  const img = h('img', { alt: 'Photo' });
  const cpt = h('div', { class: 'lb-cpt' });
  const show = k => { i = (k + urls.length) % urls.length; img.src = urls[i]; cpt.textContent = urls.length > 1 ? `${i + 1} / ${urls.length}` : ''; };
  const key = e => { if (e.key === 'Escape') close(); if (e.key === 'ArrowRight') show(i + 1); if (e.key === 'ArrowLeft') show(i - 1); };
  const close = () => { box.remove(); document.removeEventListener('keydown', key); };
  const box = h('div', { class: 'lightbox', onclick: e => { if (e.target === box) close(); } }, img, cpt,
    urls.length > 1 ? [h('button', { class: 'lb-nav lb-prev', onclick: () => show(i - 1), 'aria-label': 'Précédente' }, '‹'), h('button', { class: 'lb-nav lb-next', onclick: () => show(i + 1), 'aria-label': 'Suivante' }, '›')] : null,
    h('button', { class: 'lb-close', onclick: close, 'aria-label': 'Fermer' }, '×'));
  document.addEventListener('keydown', key);
  document.body.append(box);
  show(i);
}

/* ====================================================================== */
/* ÉCHANGES DPMS ↔ AGENTS (un fil par signalement)                        */
/* ====================================================================== */
const ICO_COURRIER = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/></svg>';
function icoCourrier(n, alerte) { return h('span', { class: 'courrier' + (alerte ? ' alerte' : ''), title: 'Échanges', html: ICO_COURRIER + (n ? `<b>${Number(n)}</b>` : '') }); }
const cleLus = () => state.mode === 'superviseur' ? 'msgLusSup' : 'msgLus';
const roleAutre = () => state.mode === 'superviseur' ? 'agent' : 'dpms';
async function nonLus(s, lus) {
  lus = lus || new Set((await kvGet(cleLus())) || []);
  return (s.messages || []).filter(m => m.role === roleAutre() && !m.cloture && !lus.has(m.id)).length;
}
async function marquerLus(s) {
  const lus = new Set((await kvGet(cleLus())) || []);
  const avant = lus.size;
  (s.messages || []).forEach(m => lus.add(m.id));
  if (lus.size !== avant) await kvSet(cleLus(), [...lus]);
}
async function etatEchanges() {
  const lus = new Set((await kvGet(cleLus())) || []);
  const out = [];
  for (const s of await dbAll('sigs')) {
    if (!(s.messages || []).length) continue;
    out.push({ s, n: await nonLus(s, lus), demande: !!s.demande, dernier: s.messages[s.messages.length - 1] });
  }
  return out.sort((a, b) => (b.n > 0) - (a.n > 0) || b.demande - a.demande || (b.dernier.quand || '').localeCompare(a.dernier.quand || ''));
}
function filMessages(s) {
  const ms = s.messages || [];
  if (!ms.length) return h('p', { class: 'muted small' }, 'Aucun échange pour l’instant.');
  const moi = m => (state.mode === 'superviseur') === (m.role === 'dpms');
  return h('div', { class: 'fil' }, ...ms.map(m => m.cloture
    ? h('div', { class: 'fil-sys' }, `${fmtDT(m.quand)} — demande d’action close par le DPMS`)
    : h('div', { class: 'bulle' + (moi(m) ? ' moi' : '') + (m.demande ? ' demande' : '') },
      h('div', { class: 'bulle-tete' }, h('b', null, m.de), ` · ${fmtDT(m.quand)} `,
        m.demande ? h('span', { class: 'badge b-warn' }, 'Demande d’action') : null, m.fait ? h('span', { class: 'badge b-ok' }, 'Action faite') : null),
      m.texte ? h('div', { class: 'bulle-txt' }, m.texte) : null)));
}
async function formReponseAgent(s, apres) {
  const d = { par: await agentPrefere(), texte: '' };
  const ta = h('textarea', { placeholder: 'Votre réponse au DPMS (ce qui a été fait, ce qui reste à faire…)', oninput: e => d.texte = e.target.value });
  const envoyer = fait => safe(async () => {
    if (!d.par) throw new Error('Indiquez qui répond.');
    if (!fait && !d.texte.trim()) throw new Error('Écrivez votre réponse.');
    if (fait && !confirm('Indiquer au DPMS que l’action demandée est faite ?')) return;
    await appendEvent('msg', { site: s.site, agent: d.par }, { sig: s.id, ref: s.ref, texte: d.texte.trim(), fait: !!fait, quand: new Date().toISOString() });
    await fold();
    await marquerLus(await dbGet('sigs', s.id));
    toast(fait ? 'Action signalée comme faite au DPMS.' : 'Réponse envoyée au DPMS.');
    syncNow().then(refreshIfHome);
    apres();
  });
  return h('div', { class: 'reponse' },
    h('label', { class: 'f' }, 'Qui répond ?'), chips(state.cfg.agents, d.par, v => d.par = v), ta,
    h('div', { class: 'row' },
      h('button', { class: 'ok', onclick: envoyer(false) }, 'Répondre'),
      s.demande ? h('button', { class: 'accent', onclick: envoyer(true) }, 'Action faite') : null));
}
function bandeauDemande(s) {
  if (!s.demande) return null;
  const m = s.messages.find(x => x.id === s.demande);
  return h('div', { class: 'banner warn', style: 'display:flex;gap:8px;align-items:flex-start' }, icoCourrier(0, true),
    h('div', null, h('b', null, 'Demande d’action du DPMS'), m && m.texte ? h('div', null, m.texte) : null));
}
async function carteEchangesAgents(s, apres) {
  await marquerLus(s);
  return h('div', { class: 'card' }, h('h2', null, 'Échanges avec le DPMS'),
    filMessages(s), await formReponseAgent(s, apres));
}
VIEWS.echanges = async ({ id }) => {
  const s = await dbGet('sigs', id);
  if (!s) return page('Échanges', { back: true }, h('p', null, 'Signalement introuvable.'));
  const th = await thumbOf(s);
  return page('Échanges', { back: true },
    h('div', { class: 'card', style: 'display:flex;gap:10px;align-items:flex-start' },
      th ? h('img', { src: th, alt: '', class: 'vignette', onclick: () => lightbox([th], 0) }) : null,
      h('div', null, h('div', { class: 'l1' }, sigTitre(s)), h('div', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ref} — ${depuisJours(s.jour)}${s.emplacement ? ' — ' + s.emplacement : ''}`),
        s.desc ? h('div', { class: 'small' }, s.desc) : null,
        h('button', { class: 'link', style: 'width:auto;margin:0;min-height:32px', onclick: () => go('suiviSig', { id: s.id }) }, 'Voir la fiche complète'))),
    bandeauDemande(s),
    await carteEchangesAgents(s, () => go('echanges', { id }, { noPush: true })));
};
VIEWS.messagesAgents = async () => {
  const l = await etatEchanges();
  return page('Messages du DPMS', { back: true },
    l.length ? l.map(({ s, n, demande, dernier }) => h('button', { class: 'list-item' + (n || demande ? ' msg-new' : ''), onclick: () => go('echanges', { id: s.id }) },
      h('div', { style: 'display:flex;gap:10px;align-items:flex-start' }, icoCourrier(n, n > 0 || demande),
        h('div', { style: 'flex:1;min-width:0' },
          h('div', { class: 'l1' }, sigTitre(s), demande ? h('span', { class: 'badge b-warn' }, 'Action demandée') : null),
          h('div', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ref}`),
          h('div', { class: 'small' }, `${dernier.de} : ${dernier.cloture ? 'demande close' : dernier.texte || (dernier.fait ? 'action faite' : '')}`)))))
      : h('p', { class: 'muted' }, 'Aucun échange.'));
};
async function carteMessagesAccueil() {
  const l = await etatEchanges();
  const nonlus = l.reduce((a, x) => a + x.n, 0), dem = l.filter(x => x.demande).length;
  if (!nonlus && !dem) return null;
  const sup = state.mode === 'superviseur';
  return h('button', { class: 'list-item msg-new', onclick: () => sup ? go('supSigs', { f: 'echanges' }) : go('messagesAgents') },
    h('div', { style: 'display:flex;gap:12px;align-items:center' }, icoCourrier(nonlus, true),
      h('div', null, h('div', { class: 'l1' }, sup ? 'Échanges avec les agents' : 'Messages du DPMS'),
        h('div', { class: 'small' }, [nonlus ? `${nonlus} ${sup ? 'réponse(s) non lue(s)' : 'message(s) non lu(s)'}` : null, dem ? `${dem} demande(s) d’action en attente` : null].filter(Boolean).join(' · ')))));
}
async function carteEchangesSup(s, recharger) {
  await marquerLus(s);
  const ta = h('textarea', { placeholder: 'Question ou consigne aux agents (ex. : quelles suites ont été données ?)' });
  const envoyer = demande => safe(async () => {
    const texte = ta.value.trim();
    if (!texte) throw new Error('Écrivez le message.');
    await queueDec('d', { type: 'msg', id: uuid(), ts: new Date().toISOString(), sig: s.id, ref: s.ref, texte, demande: !!demande, par: 'DPMS' });
    await fold(); syncNow();
    toast(demande ? 'Demande d’action envoyée aux agents.' : 'Message envoyé aux agents.');
    recharger();
  });
  return h('div', { class: 'card' }, h('h2', null, 'Échanges avec les agents'),
    filMessages(s), ta,
    h('div', { class: 'row' },
      h('button', { class: 'sec', onclick: envoyer(false) }, 'Envoyer'),
      h('button', { class: 'accent', onclick: envoyer(true) }, 'Demande d’action')),
    s.demande ? h('button', { class: 'link', onclick: safe(async () => {
      await queueDec('d', { type: 'demande_close', id: uuid(), ts: new Date().toISOString(), sig: s.id, par: 'DPMS' });
      await fold(); syncNow(); toast('Demande close.'); recharger();
    }) }, 'Clore la demande sans réponse') : null);
}

/* Notifications sur l'appareil (Web Push, sans contenu : « nouveau message ») */
async function majBadge() {
  try { if (!navigator.setAppBadge || !state.cfg) return; const l = await etatEchanges(); const n = l.reduce((a, x) => a + x.n, 0) + l.filter(x => x.demande && state.mode !== 'superviseur').length; n ? navigator.setAppBadge(n) : navigator.clearAppBadge(); } catch (e) { }
}
const PUSH_AUDS = ['https://fcm.googleapis.com', 'https://updates.push.services.mozilla.com', 'https://web.push.apple.com'];
const pushDispo = () => 'serviceWorker' in navigator && 'PushManager' in window && 'Notification' in window;
async function assurerVapid() {   // superviseur : clés de notification de l'équipe (la première publiée fait foi)
  if (await kvGet('vapid')) return false;
  if (((await kvGet('decQueue')) || []).some(q => q.d && q.d.type === 'vapid')) return false;
  const k = await crypto.subtle.generateKey({ name: 'ECDSA', namedCurve: 'P-256' }, true, ['sign', 'verify']);
  const pub = b64u.enc(new Uint8Array(await crypto.subtle.exportKey('raw', k.publicKey)));
  const priv = await crypto.subtle.exportKey('jwk', k.privateKey);
  await queueDec('d', { type: 'vapid', id: uuid(), ts: new Date().toISOString(), pub, priv });
  return true;
}
async function jwtVapid(aud, priv) {
  const key = await crypto.subtle.importKey('jwk', { kty: priv.kty, crv: priv.crv, x: priv.x, y: priv.y, d: priv.d }, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['sign']);
  const enc = o => b64u.enc(te.encode(JSON.stringify(o)));
  const t = enc({ typ: 'JWT', alg: 'ES256' }) + '.' + enc({ aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: appBaseURL() });
  return t + '.' + b64u.enc(new Uint8Array(await crypto.subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, key, te.encode(t))));
}
/* Texte de la notification chiffré pour l'appareil destinataire (RFC 8291, aes128gcm) :
   le relais et le service de notification ne voient qu'un bloc chiffré */
async function chiffrerPush(abo, texte) {
  const uaPub = b64u.dec(abo.p256dh), auth = b64u.dec(abo.auth);
  const eph = await crypto.subtle.generateKey({ name: 'ECDH', namedCurve: 'P-256' }, true, ['deriveBits']);
  const asPub = new Uint8Array(await crypto.subtle.exportKey('raw', eph.publicKey));
  const uaKey = await crypto.subtle.importKey('raw', uaPub, { name: 'ECDH', namedCurve: 'P-256' }, false, []);
  const secret = await crypto.subtle.deriveBits({ name: 'ECDH', public: uaKey }, eph.privateKey, 256);
  const hkdf = async (sel, ikm, info, n) => new Uint8Array(await crypto.subtle.deriveBits({ name: 'HKDF', hash: 'SHA-256', salt: sel, info },
    await crypto.subtle.importKey('raw', ikm, 'HKDF', false, ['deriveBits']), n * 8));
  const cat = (...l) => { const o = new Uint8Array(l.reduce((t, x) => t + x.length, 0)); let i = 0; l.forEach(x => { o.set(x, i); i += x.length; }); return o; };
  const ikm = await hkdf(auth, new Uint8Array(secret), cat(te.encode('WebPush: info\0'), uaPub, asPub), 32);
  const sel = crypto.getRandomValues(new Uint8Array(16));
  const cek = await hkdf(sel, ikm, te.encode('Content-Encoding: aes128gcm\0'), 16);
  const nonce = await hkdf(sel, ikm, te.encode('Content-Encoding: nonce\0'), 12);
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv: nonce },
    await crypto.subtle.importKey('raw', cek, 'AES-GCM', false, ['encrypt']), cat(te.encode(texte), new Uint8Array([2]))));
  const rs = new Uint8Array([0, 0, 16, 0]);
  const tout = cat(sel, rs, new Uint8Array([asPub.length]), asPub, ct);
  let bin = ''; tout.forEach(x => bin += String.fromCharCode(x));
  return btoa(bin);
}
function composerNotif(items) {
  const coupe = (t, n) => t.length > n ? t.slice(0, n - 1) + '…' : t;
  return items.length === 1 ? { ...items[0], corps: coupe(items[0].corps || '', 600) }
    : { type: 'plusieurs', titre: `Rondes parkings : ${items.length} nouveautés`, corps: coupe(items.map(i => `${i.titre} — ${i.corps}`).join('\n'), 1500) };
}
async function chiffrerPour(abos, type, message) {
  const out = {};
  for (const a of abos || []) if (!a.prefs || a.prefs.includes(type)) { try { out[a.e] = await chiffrerPush(a, JSON.stringify(message)); } catch (e) { } }
  return out;
}
async function jeton(type, id) {   // identifiant anonyme pour le relais (il ne peut pas retrouver le parking ni la ronde)
  const hsh = new Uint8Array(await crypto.subtle.digest('SHA-256', te.encode(`${state.conn.team}|${type}|${id}`)));
  return b64u.enc(hsh).slice(0, 32);
}
async function notifier(cible, items) {
  try {
    const v = await kvGet('vapid'); if (!v) return;
    const jwts = {}; for (const a of PUSH_AUDS) jwts[a] = await jwtVapid(a, v.priv);
    let corps = null;
    if (items && items.length) {
      try {
        const r = await rel('abos', { cible }, 20000);
        corps = {};
        for (const a of r.abos || []) {
          const retenus = items.filter(i => !a.prefs || a.prefs.includes(i.type));   // préférences de l'appareil destinataire
          if (!retenus.length) continue;
          try { corps[a.e] = await chiffrerPush(a, JSON.stringify(composerNotif(retenus))); } catch (e) { }
        }
      } catch (e) { corps = null; }
    }
    await rel('notifier', { cible, jwts, k: v.pub, corps, seulsChiffres: !!corps }, 30000);
  } catch (e) { console.warn('notification non envoyée :', e.message); }
}
async function abonnerPush(v, force) {
  const reg = await navigator.serviceWorker.ready;
  let sub = await reg.pushManager.getSubscription();
  const cle = sub && sub.options && sub.options.applicationServerKey ? b64u.enc(new Uint8Array(sub.options.applicationServerKey)) : null;
  if (sub && cle !== v.pub) { await sub.unsubscribe(); sub = null; }
  if (!sub) sub = await reg.pushManager.subscribe({ userVisibleOnly: true, applicationServerKey: b64u.dec(v.pub) });
  const prev = await kvGet('pushAbo');
  const cles = (sub.toJSON && sub.toJSON().keys) || {};
  const prefs = await prefsNotif(), pj = prefs.join(',');
  if (force || !prev || !prev.cles || prev.prefs !== pj || prev.endpoint !== sub.endpoint || prev.pub !== v.pub || Date.now() - prev.at > 86400000) {
    try { await rel('abonner', { endpoint: sub.endpoint, role: state.mode === 'superviseur' ? 'sup' : 'agents', p256dh: cles.p256dh, auth: cles.auth, prefs }, 30000); }
    catch (e) { if (/opération inconnue/.test(e.message)) throw new Error('le relais doit être mis à jour (version 4) pour les notifications'); throw e; }
    await kvSet('pushAbo', { endpoint: sub.endpoint, pub: v.pub, at: Date.now(), cles: !!cles.p256dh, prefs: pj });
  }
}
async function maintenirPush() {
  try {
    if (!pushDispo() || Notification.permission !== 'granted' || !['agent', 'superviseur'].includes(state.mode)) return;
    const v = await kvGet('vapid');
    if (v && (await kvGet('pushAbo'))) await abonnerPush(v);
  } catch (e) { console.warn(e.message); }
}
/* Réglages : choix des notifications reçues sur cet appareil */
async function carteNotifPrefs() {
  if (!['agent', 'superviseur'].includes(state.mode)) return null;
  const prefs = await prefsNotif();
  const abo = await kvGet('pushAbo');
  const actives = pushDispo() && Notification.permission === 'granted' && abo;
  const ligne = ([type, lbl]) => h('button', {
    class: 'case' + (prefs.includes(type) ? ' on' : ''), onclick: safe(async () => {
      const i = prefs.indexOf(type); if (i >= 0) prefs.splice(i, 1); else prefs.push(type);
      await kvSet('notifPrefs', [...prefs]); await kvSet('notifConnus', typesNotif().map(x => x[0]));
      render();
      const v = await kvGet('vapid');
      if (actives && v) {
        try {
          await abonnerPush(v);
          if (state.mode === 'superviseur') await programmerAlertes(true);
          toast('Choix enregistré.');
        } catch (e) { toast('Choix enregistré sur l’appareil ; il sera transmis au relais à la prochaine connexion.', true); }
      } else toast('Choix enregistré.');
    })
  }, h('span', { class: 'coche' }, prefs.includes(type) ? '✓' : ''), lbl);
  return h('div', { class: 'card' }, h('h2', null, 'Notifications reçues sur cet appareil'),
    actives ? null : h('p', { class: 'muted small' }, 'Les notifications ne sont pas encore activées sur cet appareil (carte « Cet appareil » de l’accueil). Vos choix seront appliqués dès leur activation.'),
    h('div', { class: 'cases' }, ...typesNotif().flatMap(x => [GROUPES_NOTIF[x[0]] && state.mode === 'superviseur' ? h('p', { class: 'muted small', style: 'margin:14px 0 0;font-weight:700' }, GROUPES_NOTIF[x[0]]) : null, ligne(x)])),
    h('div', { class: 'row' },
      h('button', { class: 'sec', onclick: safe(async () => { await kvSet('notifPrefs', typesNotif().map(x => x[0])); await kvSet('notifConnus', typesNotif().map(x => x[0])); await maintenirPush(); if (state.mode === 'superviseur') programmerAlertes(true).catch(() => { }); render(); }) }, 'Tout cocher'),
      h('button', { class: 'sec', onclick: safe(async () => { await kvSet('notifPrefs', []); await kvSet('notifConnus', typesNotif().map(x => x[0])); await maintenirPush(); if (state.mode === 'superviseur') programmerAlertes(true).catch(() => { }); render(); }) }, 'Tout décocher')));
}
/* Installation sur l'écran d'accueil et notifications : carte « Cet appareil », affichée tant qu'il reste quelque chose à faire */
let _installEvt = null;
window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); _installEvt = e; if (['home', 'plus'].includes(state.view)) render(); });
window.addEventListener('appinstalled', () => { _installEvt = null; toast('Application installée sur l’écran d’accueil.'); if (['home', 'plus'].includes(state.view)) render(); });
const estInstallee = () => (window.matchMedia && matchMedia('(display-mode: standalone)').matches) || navigator.standalone === true;
async function carteAppareil() {
  if (!['agent', 'superviseur'].includes(state.mode)) return null;
  const lignes = [];
  if (!estInstallee()) {
    if (_installEvt) lignes.push(h('button', { class: 'accent', onclick: safe(async () => { const e = _installEvt; _installEvt = null; e.prompt(); await e.userChoice; render(); }) }, 'Installer l’application sur cet appareil'));
    else lignes.push(h('p', { class: 'small', style: 'margin:4px 0' }, h('b', null, 'Installer l’application : '),
      'menu ⋮ de Chrome (en haut à droite) → « Installer l’application » ou « Ajouter à l’écran d’accueil ». ',
      'Si ce menu n’existe pas ou ne propose pas l’installation, la page est ouverte dans un navigateur intégré (Google Lens, messagerie…) : menu ⋮ → « Ouvrir dans Chrome », puis recommencer.'));
  }
  if (!pushDispo()) lignes.push(h('p', { class: 'small', style: 'margin:4px 0' }, h('b', null, 'Notifications : '), 'indisponibles dans ce navigateur. Ouvrez l’application dans Chrome (de préférence installée).'));
  else { const c = await cartePush(); if (c) lignes.push(c); }
  if (!lignes.length) return null;
  return h('div', { class: 'card', style: 'border:2px solid var(--accent)' }, h('h2', null, 'Cet appareil'), ...lignes);
}
async function cartePush() {
  if (!pushDispo() || !['agent', 'superviseur'].includes(state.mode)) return null;
  if (Notification.permission === 'granted' && (await kvGet('pushAbo'))) return null;
  if (Notification.permission === 'denied') return h('p', { class: 'muted small' }, 'Notifications bloquées sur cet appareil : les autoriser pour ce site dans les réglages de Chrome.');
  return h('button', { class: 'sec', onclick: safe(async () => {
    const v = await kvGet('vapid');
    if (!v) throw new Error('Notifications pas encore prêtes : un superviseur doit d’abord ouvrir son application connectée au réseau.');
    if ((await Notification.requestPermission()) !== 'granted') throw new Error('Notifications refusées sur cet appareil.');
    showBusy('Activation…');
    try { await abonnerPush(v, true); } finally { hideBusy(); }
    toast('Notifications activées sur cet appareil.'); render();
  }) }, 'Activer les notifications sur cet appareil');
}

/* ---------- Vues de traitement (équipe : téléphone agents ou PC de la loge) ---------- */
/* Liste à traiter, regroupée par étape */
/* ---------- Transmission aux services : appel ou courriel faits « à l'extérieur »
   (téléphone de service, ordinateur de la loge), puis indiqués dans l'application ---------- */
const DEST_AUTO_RE = /\bdpms\b/i;   // le DPMS reçoit tout par l'application : rien à transmettre
const TYPES_TRANSMIS = ['tel', 'mail', 'autre', 'relance'];
function destsATransmettre(s) {
  return (s.dests || []).filter(dn => !DEST_AUTO_RE.test(dn) && !(s.actions || []).some(a => a.dest === dn && TYPES_TRANSMIS.includes(a.type)));
}
async function sigsATransmettre() {
  return (await dbAll('sigs')).filter(s => s.statut === 'ouvert' && suiviEtat(s) === 'À transmettre')
    .sort((a, b) => (b.urgent - a.urgent) || (a.ts || a.jour || '').localeCompare(b.ts || b.jour || ''));
}
async function copier(texte, quoi) {
  try { await navigator.clipboard.writeText(texte); toast(`${quoi} copié.`); }
  catch (e) { prompt(`Copiez ${quoi.toLowerCase()} :`, texte); }
}
let _transPar = null;
VIEWS.transmettre = async ({ id } = {}) => {
  const liste = await sigsATransmettre();
  let s = id ? await dbGet('sigs', id) : liste[0];
  if (!s) return page('Transmettre', { back: () => go('home') }, h('div', { class: 'fini' }, h('span', { html: ICO.ok }), h('h2', null, 'Tout est transmis'), h('p', null, 'Aucun signalement n’attend d’être transmis.')), h('button', { class: 'big', onclick: () => go('home') }, 'Retour à l’accueil'));
  const ronde = await kvGet('ronde');
  if (!_transPar) _transPar = ronde ? ronde.agent : null;
  const rang = liste.findIndex(x => x.id === s.id), total = liste.length;
  const reste = destsATransmettre(s);
  const suivant = liste.find(x => x.id !== s.id);
  const recharger = () => go('transmettre', { id: s.id }, { noPush: true, keepScroll: true });
  const sujet = `[Parking ${siteNom(s.site)}] ${sigTitre(s)} — ${s.ref}`;
  const texte = texteSignalement(s);
  const precisions = {};
  const enregistrer = (dn, type) => safe(async () => {
    if (!_transPar) { toast('Indiquez d’abord qui transmet (en haut).', true); document.querySelector('.qui-transmet') && document.querySelector('.qui-transmet').scrollIntoView({ behavior: 'smooth' }); return; }
    const p = (precisions[dn] && precisions[dn].value || '').trim();
    await appendEvent('action', { site: s.site, agent: _transPar }, { sig: s.id, ref: s.ref, type, dest: dn, texte: type === 'echec' ? ('Pas de réponse' + (p ? ' — ' + p : '')) : p, prevu: null, quand: new Date().toISOString() });
    await fold(); syncNow();
    const maj = await dbGet('sigs', s.id);
    toast(type === 'echec' ? `Noté : ${dn} pas joint. Réessayez plus tard.` : type === 'relance' ? `Relance de ${dn} notée.` : `Noté : ${dn} ${type === 'tel' ? 'appelé' : type === 'mail' ? 'prévenu par courriel' : 'prévenu'}.`);
    if (type !== 'echec' && type !== 'relance' && !destsATransmettre(maj).length) go('transmisOk', { id: s.id }, { noPush: true });
    else recharger();
  });
  const carteDest = dn => {
    const c = (state.cfg.contacts || {})[dn] || {};
    const fait = (s.actions || []).filter(a => a.dest === dn && TYPES_TRANSMIS.includes(a.type)).pop();
    const echec = (s.actions || []).filter(a => a.dest === dn && a.type === 'echec').pop();
    precisions[dn] = h('input', { type: 'text', placeholder: 'Précision (facultatif) : nom de la personne jointe, n° de ticket…' });
    return h('div', { class: 'card dest' + (fait ? ' fait' : '') },
      h('div', { class: 'dest-nom' }, dn, fait ? h('span', { class: 'badge b-ok' }, `${ACTIONS[fait.type].toLowerCase()} ${fmtQuand(fait.quand || fait.jour)}${fait.par ? ' — ' + fait.par : ''}`) : echec ? h('span', { class: 'badge b-warn' }, `pas joint ${fmtQuand(echec.quand || echec.jour)}`) : h('span', { class: 'badge b-bad' }, 'à prévenir')),
      c.tel ? h('div', { class: 'dest-l' }, h('span', { class: 'muted small' }, 'Téléphone'), h('a', { class: 'dest-tel', href: 'tel:' + c.tel.replace(/\s/g, '') }, c.tel)) : null,
      c.mail ? h('div', { class: 'dest-l' }, h('span', { class: 'muted small' }, 'Courriel'), h('span', { class: 'dest-mail' }, c.mail), h('button', { class: 'sec mini', onclick: () => copier(c.mail, 'Adresse') }, 'Copier l’adresse')) : null,
      !c.tel && !c.mail ? h('p', { class: 'muted small' }, 'Coordonnées non renseignées : demandez-les au DPMS.') : null,
      fait ? [precisions[dn], h('div', { class: 'dest-btns' }, h('button', { class: 'sec', onclick: enregistrer(dn, 'relance') }, 'J’ai relancé ce service'))] : [
        precisions[dn],
        h('p', { class: 'small', style: 'margin:10px 0 4px;font-weight:600' }, 'Une fois fait, touchez :'),
        h('div', { class: 'dest-btns' },
          h('button', { class: 'ok', onclick: enregistrer(dn, 'tel') }, '✓ J’ai appelé'),
          h('button', { class: 'ok', onclick: enregistrer(dn, 'mail') }, '✓ J’ai envoyé le courriel')),
        h('div', { class: 'dest-btns' },
          h('button', { class: 'sec', onclick: enregistrer(dn, 'echec') }, 'Pas de réponse'),
          h('button', { class: 'sec', onclick: enregistrer(dn, 'autre') }, 'Prévenu autrement'))]);
  };
  const th = await thumbOf(s);
  return page('Transmettre', { back: () => go('home') },
    total > 1 && rang >= 0 ? h('p', { class: 'compte-etape' }, `Signalement ${rang + 1} sur ${total} à transmettre`) : null,
    h('div', { class: 'card' + (s.urgent ? ' urgent-bord' : '') },
      h('div', { style: 'display:flex;gap:12px;align-items:flex-start' },
        th ? h('img', { src: th, alt: '', class: 'vignette' }) : null,
        h('div', { style: 'flex:1;min-width:0' },
          h('b', { style: 'font-size:19px' }, sigTitre(s)), s.urgent ? h('span', { class: 'badge b-bad' }, 'URGENT') : null,
          h('div', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ref} — signalé ${depuisJours(s.jour)}${s.agent ? ' par ' + s.agent : ''}`),
          s.desc ? h('div', { class: 'small' }, s.desc) : null))),
    h('div', { class: 'banner warn', style: 'font-size:16px' }, h('b', null, 'À faire : '), reste.length ? `prévenir ${reste.join(', ')}` : 'rien de plus',
      ' — par téléphone (téléphone de service) ou par courriel (ordinateur de la loge), puis l’indiquer ici.'),
    h('div', { class: 'card qui-transmet' }, h('p', { class: 'question', style: 'margin-top:0' }, 'Qui transmet ?'), chips(state.cfg.agents, _transPar, v => { _transPar = v; })),
    ...(s.dests || []).filter(dn => !DEST_AUTO_RE.test(dn)).map(carteDest),
    h('details', { class: 'card' }, h('summary', null, 'Texte à dire ou à envoyer (copier)'),
      h('p', { class: 'small', style: 'margin:8px 0 4px' }, h('b', null, 'Objet : '), sujet),
      h('pre', { class: 'texte-sig' }, texte),
      h('div', { class: 'row' },
        h('button', { class: 'sec', onclick: () => copier(sujet, 'Objet') }, 'Copier l’objet'),
        h('button', { class: 'sec', onclick: () => copier(texte, 'Texte') }, 'Copier le texte')),
      (s.photos || []).length ? h('button', { class: 'sec', onclick: safe(async () => { showBusy('Photos…'); try { const { photos } = await photosDe(s.photos); hideBusy(); photos.forEach((p, i) => download(p.blob, `${s.ref}-${i + 1}.jpg`)); toast(`${photos.length} photo(s) téléchargée(s) : joignez-les au courriel.`); } finally { hideBusy(); } }) }, 'Télécharger les photos (pour le courriel)') : null),
    suivant ? h('button', { class: 'link', onclick: () => go('transmettre', { id: suivant.id }, { noPush: true }) }, 'Passer au signalement suivant (celui-ci restera à transmettre)') : null,
    h('button', { class: 'link', onclick: () => go('suiviSig', { id: s.id }) }, 'Voir la fiche complète'));
};
VIEWS.transmisOk = async ({ id }) => {
  const s = await dbGet('sigs', id);
  const liste = await sigsATransmettre();
  return page('Transmis', null,
    h('div', { class: 'fini' }, h('span', { html: ICO.ok }), h('h2', null, 'Signalement transmis'), h('p', null, s ? `${sigTitre(s)} — ${s.ref}` : ''),
      h('p', { class: 'muted' }, 'Il passe en « en attente de réponse ». Quand le service répond ou intervient, indiquez-le dans « Suivi des signalements ».')),
    liste.length ? h('button', { class: 'danger big', onclick: () => go('transmettre', { id: liste[0].id }, { noPush: true }) }, `Transmettre le suivant (${liste.length} restant${liste.length > 1 ? 's' : ''})`) : null,
    h('button', { class: (liste.length ? 'sec' : 'ok') + ' big', onclick: () => go('home') }, 'Retour à l’accueil'));
};
async function carteATransmettre() {
  const liste = await sigsATransmettre();
  if (!liste.length) return null;
  const vieux = liste.filter(s => joursDepuis(s.ts || s.jour) >= 1).length;
  return h('div', { class: 'card a-transmettre' },
    h('div', { class: 'at-titre' }, `${liste.length} signalement${liste.length > 1 ? 's' : ''} à transmettre`),
    h('p', { class: 'small', style: 'margin:2px 0 8px' }, vieux ? `dont ${vieux} depuis plus d’un jour. ` : '', 'Appelez ou écrivez au service concerné, puis indiquez-le dans l’application.'),
    ...liste.slice(0, 3).map(s => h('button', { class: 'list-item', onclick: () => go('transmettre', { id: s.id }) },
      h('div', { class: 'l1' }, sigTitre(s), s.urgent ? h('span', { class: 'badge b-bad' }, 'URGENT') : null),
      h('div', { class: 'muted small' }, `${siteNom(s.site)} — ${depuisJours(s.jour)} — à prévenir : ${destsATransmettre(s).join(', ')}`))),
    liste.length > 3 ? h('p', { class: 'muted small' }, `… et ${liste.length - 3} autre(s).`) : null,
    h('button', { class: 'danger big', onclick: () => go('transmettre', { id: liste[0].id }) }, 'Transmettre maintenant'));
}
const PROCHAINE = { 'À transmettre': 'À transmettre aux services', 'En attente de réponse': 'Relancer si pas de réponse', 'Réponse reçue': 'Suivre l’intervention', 'Intervention prévue': 'Vérifier le passage', 'Intervention réalisée': 'À clore si le problème est réglé' };
VIEWS.suivi = async () => {
  const enCours = (await dbAll('sigs')).filter(s => s.statut === 'ouvert').sort((a, b) => (b.urgent - a.urgent) || (a.jour || '').localeCompare(b.jour || ''));
  const groupes = [];
  for (const e of SUIVIS) {
    const l = enCours.filter(s => suiviEtat(s) === e);
    if (!l.length) continue;
    const items = [];
    for (const s of l) {
      const th = await thumbOf(s);
      const nl = await nonLus(s);
      items.push(h('button', { class: 'list-item', onclick: () => go(e === 'À transmettre' ? 'transmettre' : 'suiviSig', { id: s.id }) },
        h('div', { style: 'display:flex;gap:12px;align-items:center' },
          th ? h('img', { src: th, alt: '', class: 'vignette' }) : h('div', { class: 'vignette vide' }),
          h('div', { style: 'flex:1;min-width:0' },
            h('div', { class: 'l1' }, sigTitre(s), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, (s.messages || []).length ? icoCourrier(nl, nl > 0 || !!s.demande) : null),
            h('div', { class: 'muted small' }, `${siteNom(s.site)} — ${depuisJours(s.jour)} — ${(s.dests || []).join(', ')}`)))));
    }
    groupes.push(h('h2', { class: 'groupe' + (e === 'À transmettre' ? ' rouge' : '') }, `${PROCHAINE[e]} (${l.length})`), ...items);
  }
  return page('Suivi des signalements', { back: true },
    h('p', { class: 'muted small', style: 'margin:0 0 8px' }, 'Touchez un signalement pour indiquer du nouveau : réponse du service, intervention prévue ou faite, problème réglé.'),
    groupes.length ? groupes : h('p', { class: 'muted' }, 'Aucun signalement en cours.'));
};

/* Fiche de traitement : un menu « Que voulez-vous faire ? », puis un écran par action */
const ACTIONS_EQUIPE = {
  reponse: { titre: 'Une réponse a été reçue', qui: 'Réponse de qui ?', detail: 'Ce qui a été répondu', type: 'reponse' },
  prevue: { titre: 'Une intervention est prévue', qui: 'Par qui ?', detail: 'Précisions (facultatif)', type: 'prevue', date: true },
  faite: { titre: 'L’intervention a été réalisée', qui: 'Par qui ?', detail: 'Précisions (facultatif)', type: 'faite' },
  autre: { titre: 'Transmis autrement', qui: 'À qui ?', detail: 'Comment (en main propre, à la loge…)', type: 'autre' },
  note: { titre: 'Ajouter une note', detail: 'Note', type: 'note', detailOblig: true },
  clore: { titre: 'Problème résolu : clore le signalement', detail: 'Commentaire (facultatif)' }
};
const actionDraft = {};
VIEWS.suiviSig = async ({ id, mode }) => {
  const s = await dbGet('sigs', id);
  if (!s) return page('Signalement', { back: true }, h('p', null, 'Introuvable.'));
  const menu = () => go('suiviSig', { id }, { noPush: true });
  const ouvrir = m => go('suiviSig', { id, mode: m });
  const resume = async (grande) => {
    const blocPhotos = grande ? await photoBlocks(s.photos || [], { rapide: true, grande: true, vignette: await thumbOf(s) }) : [];
    const der = (s.actions || []).filter(a => a.type !== 'note').slice(-1)[0];
    return h('div', { class: 'card' },
      h('h3', null, sigTitre(s), ' ', suiviBadge(s), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null),
      h('p', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ref} — signalé le ${fmtJour(s.jour)}${s.agent ? ' par ' + s.agent : ''}`),
      s.desc ? h('p', null, s.desc) : null,
      s.plaque || s.emplacement ? h('p', null, s.plaque ? [h('span', { class: 'plaque' }, s.plaque), ' '] : null, s.emplacement || '') : null,
      blocPhotos,
      h('p', { class: 'small' }, h('b', null, 'À signaler à : '), (s.dests || []).join(', ')),
      der ? h('p', { class: 'small' }, h('b', null, 'Dernière action : '), `${ACTIONS[der.type]}${der.dest ? ' — ' + der.dest : ''}, ${fmtQuand(der.quand || der.jour)}${der.par ? ' (' + der.par + ')' : ''}`) : null,
      pmBanner(s));
  };

  if (!mode) {
    return page(s.ref, { back: () => go('suivi', {}, { noPush: true }) },
      bandeauDemande(s),
      await resume(true),
      (s.messages || []).length ? h('button', { class: 'sec', onclick: () => go('echanges', { id: s.id }) }, icoCourrier(await nonLus(s), !!s.demande), 'Échanges avec le DPMS') : null,
      s.statut === 'clos' ? h('div', { class: 'banner ok' }, 'Ce signalement est clos.') : [
        h('p', { class: 'question' }, 'Que voulez-vous faire ?'),
        h('button', { class: 'choix' + (suiviEtat(s) === 'À transmettre' ? ' rouge' : ''), onclick: () => go('transmettre', { id: s.id }) }, suiviEtat(s) === 'À transmettre' ? 'Transmettre au service (appel ou courriel)' : 'Relancer le service (appel ou courriel)'),
        h('button', { class: 'choix', onclick: () => ouvrir('reponse') }, ACTIONS_EQUIPE.reponse.titre),
        h('button', { class: 'choix', onclick: () => ouvrir('prevue') }, ACTIONS_EQUIPE.prevue.titre),
        h('button', { class: 'choix', onclick: () => ouvrir('faite') }, ACTIONS_EQUIPE.faite.titre),
        h('button', { class: 'choix vert', onclick: () => ouvrir('clore') }, ACTIONS_EQUIPE.clore.titre),
        h('button', { class: 'link', onclick: () => ouvrir('note') }, 'Ajouter une note')],
      h('details', { class: 'card historique' }, h('summary', null, `Historique (${(s.actions || []).length + (s.suivi || []).length})`),
        ...journalActions(s),
        (s.suivi || []).length ? [h('p', { class: 'small', style: 'font-weight:600;margin:10px 0 2px' }, 'Constats en ronde'),
          ...s.suivi.map(v => h('p', { class: 'small', style: 'margin:2px 0' }, `${fmtJour(v.jour)} — ${v.agent || ''} : ${VERDICTS[v.verdict]}${v.comment ? ' — ' + v.comment : ''}`))] : null));
  }

  if (mode === 'transmettre') return VIEWS.transmettre({ id });

  const def = ACTIONS_EQUIPE[mode];
  const cle = id + '|' + mode;
  const d = actionDraft[cle] || (actionDraft[cle] = { dest: (s.dests || []).length === 1 ? s.dests[0] : null, par: await agentPrefere(), texte: '', prevu: localDT(new Date(new Date().setHours(9, 0, 0, 0) + 86400000)) });
  const choix = (opts, sel, fn) => h('div', null, ...opts.map(o => h('button', { class: 'choix' + (o === sel ? ' on' : ''), onclick: () => { fn(o); go('suiviSig', { id, mode }, { noPush: true, keepScroll: true }); } }, o)));
  const enregistrer = safe(async () => {
    if (def.qui && !d.dest) throw new Error('Indiquez ' + def.qui.toLowerCase().replace(' ?', '') + '.');
    if (def.detailOblig && !d.texte.trim()) throw new Error('Écrivez la note.');
    if (!d.par) throw new Error('Indiquez qui enregistre.');
    if (mode === 'clore') {
      if (!(await demander({ titre: 'Clore le signalement ?', texte: 'Le problème est résolu : le signalement sera clos pour tous.', oui: 'Oui, clore', non: 'Non' }))) return;
      await appendEvent('revue', { site: s.site, agent: d.par }, { sig: s.id, ref: s.ref, verdict: 'resolu', comment: d.texte.trim() || 'Clos depuis le traitement' });
      delete actionDraft[cle];
      await fold(); syncNow(); toast('Signalement clos.'); go('suivi', {}, { noPush: true });
      return;
    }
    const prevu = def.date ? new Date(d.prevu) : null;
    if (prevu && isNaN(prevu)) throw new Error('Date invalide.');
    await appendEvent('action', { site: s.site, agent: d.par }, { sig: s.id, ref: s.ref, type: def.type, dest: def.qui ? d.dest : '', texte: d.texte.trim(), prevu: prevu ? prevu.toISOString() : null, quand: new Date().toISOString() });
    delete actionDraft[cle];
    await fold(); syncNow();
    toast('Action enregistrée.');
    if (mode === 'faite' && (await demander({ titre: 'Le problème est-il réglé ?', texte: 'Si oui, le signalement sera clos.', oui: 'Oui, clore le signalement', non: 'Non, le laisser en cours' }))) {
      await appendEvent('revue', { site: s.site, agent: d.par }, { sig: s.id, ref: s.ref, verdict: 'resolu', comment: 'Intervention réalisée' });
      await fold(); syncNow(); toast('Signalement clos.'); go('suivi', {}, { noPush: true }); return;
    }
    menu();
  });
  const dests = [...new Set([...(s.dests || []), ...state.cfg.dests])];
  return page(def.titre, { back: menu },
    h('p', { class: 'muted small' }, `${sigTitre(s)} — ${siteNom(s.site)} — ${s.ref}`),
    def.qui ? [h('p', { class: 'question' }, def.qui), choix(dests, d.dest, v => d.dest = v)] : null,
    def.date ? [h('p', { class: 'question' }, 'Prévue le'), h('input', { type: 'datetime-local', value: d.prevu, oninput: e => d.prevu = e.target.value })] : null,
    h('p', { class: 'question' }, def.detail), h('textarea', { value: d.texte, oninput: e => d.texte = e.target.value }),
    h('p', { class: 'question' }, 'Qui enregistre ?'), choix(state.cfg.agents, d.par, v => d.par = v),
    h('button', { class: (mode === 'clore' ? 'ok' : 'ok') + ' big', style: 'margin-top:16px', onclick: enregistrer }, mode === 'clore' ? 'Clore le signalement' : 'Enregistrer'));
};

/* ---------- Journal lisible (agents, PM) ---------- */
async function journalEquipe() {
  const sp = (await kvGet('supprimes')) || { sigs: [], rids: [] };
  const ds = new Set(sp.sigs), dr = new Set(sp.rids);
  const garde = ev => !((ev.t === 'signalement' && ds.has(ev.data.id)) || (['revue', 'action', 'pm', 'msg'].includes(ev.t) && ds.has(ev.data.sig)) || (ev.rid && dr.has(ev.rid) && !['signalement', 'revue'].includes(ev.t)));
  return (await journalEquipeBrut()).filter(garde);
}
async function journalEquipeBrut() {
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
    else if (ev.t === 'controle') { const r = resumeControle(ev.data.items, ev.data.libelles); aoa.push([...base, NIVEAUX[ev.data.niveau].lbl, '', '', '', r.nc.length ? 'Non conforme : ' + r.nc.join(', ') : 'Tout conforme', '', '', '', '', `${r.c} conforme(s)`, r.na.length ? 'N’existe pas : ' + r.na.join(', ') : '']); }
    else if (ev.t === 'comptage') aoa.push([...base, 'Comptage des véhicules', '', '', '', '', '', '', '', '', ev.data.total, '']);
    else if (ev.t === 'barriere') aoa.push([...base, ev.data.action === 'ouverte' ? 'Barrière ouverte' : 'Barrière refermée', '', '', ev.data.barriere, ev.data.motif || '', '', '', '', '', fmtDT(ev.data.quand), ev.data.comment || '']);
    else if (ev.t === 'action') { const s = refs.get(ev.data.sig) || {}; aoa.push([...base, 'Traitement : ' + ACTIONS[ev.data.type], ev.data.ref, s.cat || '', s.sub || '', s.desc || '', s.plaque || '', s.emplacement || '', ev.data.dest || '', '', ev.data.prevu ? 'prévue le ' + fmtDT(ev.data.prevu) : '', ev.data.texte || '']); }
    else if (ev.t === 'msg') { const s = refs.get(ev.data.sig) || {}; aoa.push([...base, 'Message au DPMS', ev.data.ref, s.cat || '', s.sub || '', s.desc || '', s.plaque || '', s.emplacement || '', '', '', ev.data.fait ? 'Action faite' : '', ev.data.texte || '']); }
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
async function hashMdp(mdp, sel) {
  const k = await crypto.subtle.importKey('raw', te.encode(mdp), 'PBKDF2', false, ['deriveBits']);
  return b64u.enc(await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt: b64u.dec(sel), iterations: 150000 }, k, 256));
}
async function reglagesVerrou() {
  const reg = await kvGet('mdpReglages');
  const p1 = h('input', { type: 'password', autocomplete: 'new-password' });
  const p2 = h('input', { type: 'password', autocomplete: 'new-password' });
  if (!reg) {
    const definir = safe(async () => {
      if (p1.value.length < 4) throw new Error('4 caractères minimum.');
      if (p1.value !== p2.value) throw new Error('Les deux saisies diffèrent.');
      const sel = b64u.enc(crypto.getRandomValues(new Uint8Array(16)));
      await kvSet('mdpReglages', { sel, hash: await hashMdp(p1.value, sel) });
      state.reglagesOk = true; toast('Mot de passe enregistré.'); render();
    });
    p2.onkeydown = e => { if (e.key === 'Enter') definir(); };
    return page('Réglages', { back: true },
      h('div', { class: 'banner info' }, 'Première ouverture : choisissez le mot de passe qui protégera les réglages de cet appareil.'),
      h('label', { class: 'f' }, 'Mot de passe'), p1, h('label', { class: 'f' }, 'Confirmation'), p2,
      h('button', { class: 'ok', onclick: definir }, 'Définir le mot de passe'));
  }
  const ouvrir = safe(async () => {
    if ((await hashMdp(p1.value, reg.sel)) !== reg.hash) { p1.value = ''; throw new Error('Mot de passe incorrect.'); }
    state.reglagesOk = true; render();
  });
  p1.onkeydown = e => { if (e.key === 'Enter') ouvrir(); };
  setTimeout(() => p1.focus(), 50);
  return page('Réglages', { back: true }, h('label', { class: 'f' }, 'Mot de passe des réglages'), p1, h('button', { class: 'ok', onclick: ouvrir }, 'Ouvrir'));
}
VIEWS.reglages = async () => {
  if (estEquipe() && !state.reglagesOk) return reglagesVerrou();
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
    await carteNotifPrefs(),
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
    u0 ? h('div', { class: 'banner info' }, 'Adresse du relais reçue par le QR. Saisissez la phrase de passe superviseur.')
      : h('p', { class: 'muted' }, 'L’adresse du relais figure dans les Réglages de tout appareil déjà configuré, ou dans le QR « Autre superviseur » (Configurer les appareils).'),
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
  await kvUpdate('decQueue', q => [...(q || []), { qid: idFile(), kind, d }]);
}
async function supSync() {
  const vault = await kvGet('vaultPending');
  if (vault) { await rel('putVault', { data: vault }); await kvDel('vaultPending'); }
  await viderFileSup();
  const got = await pullAll(['p', 'd', 'n']);
  await fold();
  if (await assurerVapid()) { await viderFileSup(); await fold(); }
  try { await programmerAlertes(); } catch (e) { console.warn('alertes non programmées :', e.message); }
  try { if (await kvGet('vapid')) { const r = await rel('abos', { cible: 'sup' }, 20000); await programmerRappelPlanning(r.abos); } } catch (e) { console.warn('rappel du planning non programmé :', e.message); }
  if (_notifAgents.length) { const l = _notifAgents; _notifAgents = []; notifier('agents', l); }
  return { got };
}
async function programmerAlertes(force) {
  const v = await kvGet('vapid'); const cfg = state.cfg;
  if (!v || !cfg) return;
  const r = await rel('abos', { cible: 'sup' }, 20000);
  const rm = cfg.rondeMatin, sy = cfg.synthese;
  const empreinte = JSON.stringify({ a: (r.abos || []).map(a => [a.e, a.prefs]), rm, sy, s: cfg.sites.map(x => [x.id, x.nom]), k: v.pub, an: new Date().getFullYear(), fr: cfg.planning && cfg.planning.regles });
  const prev = await kvGet('programme');
  if (!force && prev && prev.empreinte === empreinte && Date.now() - prev.at < 6 * 86400000) return;
  if (!force && prev && prev.ancien && Date.now() - prev.at < 6 * 3600000) return;   // relais pas encore mis à jour
  const quotidiennes = [], hebdo = [];
  if (rm.actif !== false && cfg.sites.length <= 6) {
    const jetons = []; for (const x of cfg.sites) jetons.push(await jeton('site', x.id));
    const messages = {};
    for (let m = 1; m < (1 << cfg.sites.length); m++) {   // un message chiffré par combinaison de parkings sans ronde
      const idx = cfg.sites.map((x, i) => i).filter(i => m & (1 << i));
      messages[idx.join(',')] = await chiffrerPour(r.abos, 'alerte_ronde', { type: 'alerte_ronde', titre: 'Aucune ronde aujourd’hui', corps: `${idx.map(i => cfg.sites[i].nom).join(', ')} — aucune ronde avant ${rm.heure.replace(':', ' h ')}` });
    }
    const an = new Date().getFullYear(), exclure = [...joursFeries(an).keys(), ...joursFeries(an + 1).keys()].filter(ferieFerme);   // relais version 7 : pas d'alerte les jours fériés
    quotidiennes.push({ id: 'ronde-matin', heure: rm.heure, jours: rm.jours, jetons, messages, exclure });
  }
  if (sy.actif !== false) hebdo.push({ id: 'synthese', jour: sy.jour, heure: sy.heure,
    messages: await chiffrerPour(r.abos, 'synthese', { type: 'synthese', titre: 'Synthèse de la semaine', corps: 'Rondes, contrôles, signalements, barrières : touchez pour ouvrir le PDF.', lien: '#synthese' }) });
  try { await rel('programmer', { vapid: { d: v.priv.d, pub: v.pub, sub: appBaseURL() }, quotidiennes, hebdo }, 30000); }
  catch (e) { if (/opération inconnue/.test(e.message)) { await kvSet('programme', { ancien: true, at: Date.now() }); return; } throw e; }
  await kvSet('programme', { empreinte, at: Date.now() });
}
async function viderFileSup() {
  for (let garde = 0; garde < 500; garde++) {
    // relu à chaque tour : les décisions prises pendant l'envoi restent dans la file
    let q = (await kvGet('decQueue')) || [];
    if (!q.length) break;
    if (q.some(x => !x.qid)) q = await kvUpdate('decQueue', l => (l || []).map(x => x.qid ? x : { ...x, qid: idFile() }));   // anciennes versions
    const it = q[0];
    const data = it.kind === 'n'
      ? b64u.enc(await sealBytesFor(await getPubKey(), fflate.deflateSync(te.encode(JSON.stringify(it.d)))))
      : await encJSON(it.d);
    const r = await rel('putDec', { kind: it.kind, data });
    if (it.d && it.d.type === 'planning' && it.d.notif) _notifAgents.push({ type: 'planning', titre: it.d.notif.titre, corps: it.d.notif.corps, lien: '#planning' });
    if (it.d && it.d.type === 'msg') _notifAgents.push({ type: 'msg', titre: `${it.d.demande ? 'Demande d’action' : 'Message'} du DPMS — ${it.d.ref || ''}`, corps: it.d.texte || '' });
    // déplacement atomique : la décision passe de la file au journal local dans la même transaction
    const db = await openDB(); const t = db.transaction(['entries', 'kv'], 'readwrite');
    t.objectStore('entries').put({ k: `${it.kind}:${r.n}`, kind: it.kind === 'n' ? 'note' : 'dec', n: r.n, at: r.at, dec: it.d });
    const kv = t.objectStore('kv');
    kv.get('decQueue').onsuccess = e => kv.put((e.target.result || []).filter(x => x.qid !== it.qid), 'decQueue');
    await txDone(t);
  }
}

function bandeauMaj() {
  return h('div', { class: 'banner warn' }, 'Une version plus récente de l’application est utilisée par un autre appareil : certaines informations ne s’affichent pas ici. ',
    h('button', { class: 'sec', style: 'margin-top:6px', onclick: () => { chercherMiseAJour(true); setTimeout(() => location.reload(), 2500); } }, 'Mettre à jour maintenant'));
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
  const majRequise = await kvGet('majRequise');
  return page('Rondes parkings', null,
    majRequise ? bandeauMaj() : null,
    h('div', { class: 'card' },
      h('div', { class: 'stat' },
        h('span', null, state.sync.busy ? 'Relève en cours…' : lastSync ? `Relevé ${fmtQuand(lastSync)}` : 'Jamais relevé'),
        h('button', { class: 'chip', style: 'width:auto', disabled: state.sync.busy, onclick: safe(async () => { const r = await syncNow(); toast(r && r.err ? 'Relève impossible : ' + r.err : r && r.offline ? 'Pas de réseau.' : r && r.got ? `${r.got} élément(s) reçu(s).` : 'Rien de nouveau.', !!(r && (r.err || r.offline))); render(); }) }, 'Actualiser')),
      syncErr ? h('p', { class: 'muted small' }, 'Dernier essai : ' + syncErr.msg) : null,
      queue.length ? h('p', { class: 'muted small' }, `${queue.length} intervention(s) en attente d’envoi.`) : null),
    alertes.length ? h('button', { class: 'list-item', style: 'border-color:var(--warn)', onclick: () => go('supAlertes') },
      h('div', { class: 'l1' }, `${alertes.length} point(s) d’attention`, h('span', { class: 'badge b-warn' }, 'à lire')),
      h('div', { class: 'muted small' }, alertes[0].msg)) : null,
    await carteRondesJour(),
    await cartePlanningSup(),
    await carteMessagesAccueil(),
    await carteAppareil(),
    await alerteRegistres(),
    barrOuv.length ? h('div', { class: 'card', style: 'border-color:var(--bad)' }, h('h2', null, 'Barrières ouvertes en ce moment'), ...barrOuv.map(b => barriereLigne(b, false))) : null,
    h('button', { class: 'big', onclick: () => go('supStats') }, 'Statistiques et graphiques'),
    h('div', { class: 'card' }, h('h2', null, 'Signalements'),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'ouvert' }) }, h('div', { class: 'stat' }, h('span', null, 'En cours', urg ? h('span', { class: 'badge b-bad' }, urg + ' urgent(s)') : null), h('b', null, ouverts.length))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'a_transmettre' }) }, h('div', { class: 'stat' }, h('span', null, 'À transmettre aux services'), h('b', null, ouverts.filter(s => suiviEtat(s) === 'À transmettre').length))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'clos' }) }, h('div', { class: 'stat' }, h('span', null, 'Clos'), h('b', null, sigs.length - ouverts.length))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'vehicules' }) }, h('div', { class: 'stat' }, h('span', null, 'Véhicules en attente de la PM'), h('b', null, vehPM)))),
    await carteMainCourante(),
    await carteControlesSup(),
    h('div', { class: 'card' }, h('h2', null, 'Dernières rondes'),
      rondes.length ? rondes.map(rondeLigne) : h('p', { class: 'muted' }, 'Aucune ronde reçue.'),
      h('div', { class: 'row' },
        rondes.length ? h('button', { class: 'sec', onclick: () => go('supRondes') }, 'Toutes les rondes') : null,
        h('button', { class: 'sec', onclick: () => go('supBarrieres') }, 'Barrières'))),
    h('div', { class: 'card' }, h('h2', null, 'Exports'),
      h('button', { class: 'sec', onclick: () => go('synthese') }, 'Synthèse hebdomadaire (PDF)'),
      h('button', { class: 'sec', onclick: safe(supExcel) }, 'Excel complet (heures et positions)'),
      h('button', { class: 'sec', onclick: safe(supArchive) }, 'Archive complète avec photos (.zip)')),
    h('div', { class: 'card' }, h('h2', null, 'Paramétrage'),
      h('button', { class: 'sec', onclick: () => go('supQR') }, 'Configurer les appareils (QR et liens)'),
      h('button', { class: 'sec', onclick: () => go('supConfig') }, 'Listes, catégories et coordonnées'),
      h('button', { class: 'sec', onclick: () => go('supRelais') }, 'Relais et appareils'),
      h('button', { class: 'sec', onclick: () => go('supCorbeille') }, 'Corbeille (éléments supprimés)')),
    Object.keys(devs).length ? h('div', { class: 'card' }, h('h2', null, 'Appareils de l’équipe'),
      ...Object.entries(devs).map(([d, v]) => h('div', { class: 'stat' }, h('span', { style: 'white-space:nowrap;margin-right:8px' }, d + (v.role === 'pm' ? ' (PM)' : '')), h('span', { class: 'muted small', style: 'text-align:right' }, `${v.seq} saisies — dernier envoi ${fmtQuand(v.at)}`)))) : null,
    h('button', { class: 'link', onclick: () => go('reglages') }, 'Réglages'),
    h('img', { class: 'logo-ville', src: 'logo-cachan.png', alt: 'Ville de Cachan — parkings' }),
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
}
/* Une ronde par jour et par parking ; ronde ouverte depuis plus d'1 h */
const jourIso = (d = new Date()) => (d.getDay() || 7);
const minutesDe = hm => { const [a, b] = String(hm || '20:00').split(':').map(Number); return a * 60 + (b || 0); };
async function carteRondesJour() {
  const rm = state.cfg.rondeMatin || {}, auj = localDate(), now = new Date();
  const rondes = await dbAll('rondes');
  const duJour = rondes.filter(r => r.debut && localDate(new Date(r.debut)) === auj);
  const longues = rondes.filter(r => r.debut && !r.fin && now - new Date(r.debut) > 3600000 && now - new Date(r.debut) < 86400000);
  const attendu = rm.actif !== false && (rm.jours || []).includes(jourIso(now)) && !ferieFerme(localDate(now));
  const limite = attendu && now.getHours() * 60 + now.getMinutes() >= minutesDe(rm.heure);
  const lignes = state.cfg.sites.map(x => {
    const l = duJour.filter(r => r.site === x.id).sort((a, b) => a.debut.localeCompare(b.debut));
    const badge = l.length ? h('span', { class: 'badge b-ok' }, `${l.length > 1 ? l.length + ' rondes, ' : ''}${fmtHeure(l[0].debut)}`)
      : !attendu ? h('span', { class: 'badge' }, 'pas de ronde prévue')
        : limite ? h('span', { class: 'badge b-bad' }, 'aucune ronde') : h('span', { class: 'badge b-warn' }, `pas encore (limite ${rm.heure.replace(':', ' h ')})`);
    return h('div', { class: 'stat' }, h('span', null, x.nom), badge);
  });
  const manquants = state.cfg.sites.filter(x => !duJour.some(r => r.site === x.id));
  const alerte = (limite && manquants.length) || longues.length;
  return h('div', { class: 'card', style: alerte ? 'border-color:var(--bad)' : null },
    h('h2', null, 'Rondes du jour'),
    limite && manquants.length ? h('p', { style: 'color:var(--bad);font-weight:700;margin:0 0 6px' }, `Aucune ronde aujourd’hui à ${manquants.map(x => x.nom).join(', ')}.`) : null,
    ...longues.map(r => h('div', { style: 'margin:0 0 8px' }, h('p', { style: 'color:var(--bad);font-weight:700;margin:0 0 4px' }, `Ronde de plus d’1 h — ${siteNom(r.site)} : ${r.agent}, commencée à ${fmtHeure(r.debut)}, toujours pas terminée.`),
      h('button', { class: 'sec', style: 'width:auto;margin:0', onclick: safe(() => cloturerRonde(r)) }, 'Clôturer la ronde'))),
    ...lignes);
}
/* Registres de sécurité : contrôle mensuel par le DPMS (30 jours glissants) */
async function etatRegistres() {
  const r = (await kvGet('registres')) || [];
  return state.cfg.sites.map(s => {
    const d = r.find(x => x.site === s.id) || null;
    const j = d ? joursDepuis(d.at) : null;
    return { site: s.id, nom: s.nom, dernier: d, depuis: j, du: !d || j >= 30, hist: r.filter(x => x.site === s.id) };
  });
}
async function registreControle(site) {
  const note = prompt(`Registre de sécurité — ${siteNom(site)}\nObservations (facultatif) :`, '');
  if (note === null) return false;
  await queueDec('d', { type: 'registre', id: uuid(), ts: new Date().toISOString(), site, note: note.trim(), par: 'DPMS' });
  await fold(); syncNow(); toast('Contrôle du registre enregistré.');
  return true;
}
async function alerteRegistres() {
  const dus = (await etatRegistres()).filter(r => r.du);
  if (!dus.length) return null;
  return h('button', { class: 'list-item', style: 'border-color:var(--warn)', onclick: () => go('supControles') },
    h('div', { class: 'l1' }, 'Registres de sécurité à contrôler', h('span', { class: 'badge b-warn' }, 'mensuel')),
    h('div', { class: 'muted small' }, dus.map(r => `${r.nom} (${r.dernier ? `dernier contrôle il y a ${r.depuis} jours` : 'jamais contrôlé'})`).join(' · ')));
}
async function carteControlesSup() {
  const st = await etatControles();
  const etat = (site, n) => {
    const e = echeance(st, site, n);
    if (!e.dernier) return h('span', { class: 'badge b-bad' }, `${n} : jamais`);
    if (n === 'quotidien') return h('span', { class: 'badge ' + (e.depuis <= 1 ? 'b-ok' : 'b-bad') }, `${n} : ${e.depuis === 0 ? 'aujourd’hui' : e.depuis === 1 ? 'hier' : `il y a ${e.depuis} j`}`);
    return h('span', { class: 'badge ' + (e.du ? 'b-bad' : 'b-ok') }, `${n} : ${e.du ? `en retard (${e.depuis} j)` : e.depuis === 0 ? 'aujourd’hui' : `il y a ${e.depuis} j`}`);
  };
  const regs = await etatRegistres();
  const etatReg = r => h('span', { class: 'badge ' + (r.du ? 'b-bad' : 'b-ok') }, `registre : ${!r.dernier ? 'à contrôler' : r.du ? `à contrôler (${r.depuis} j)` : r.depuis === 0 ? 'aujourd’hui' : `il y a ${r.depuis} j`}`);
  return h('div', { class: 'card' }, h('h2', null, 'Contrôles'),
    ...state.cfg.sites.map(s => h('div', { class: 'stat', style: 'flex-wrap:wrap;gap:6px' }, h('b', null, s.nom), h('span', { style: 'display:flex;gap:4px;flex-wrap:wrap' }, ...Object.keys(NIVEAUX).map(n => etat(s.id, n)), etatReg(regs.find(r => r.site === s.id))))),
    h('button', { class: 'sec', onclick: () => go('supControles') }, 'Détail des contrôles'));
}
VIEWS.supControles = async () => {
  const st = await etatControles();
  const hist = (await kvGet('controlesHist')) || [];
  const tous = state.cfg.controles;
  const regs = await etatRegistres();
  return page('Contrôles', { back: true },
    ...state.cfg.sites.map(s => {
      const na = ((st[s.id] || {}).na || []).map(id => tous.find(c => c.id === id)).filter(Boolean);
      const rg = regs.find(r => r.site === s.id);
      return h('div', { class: 'card' }, h('h2', null, s.nom),
        h('div', { class: 'stat', style: 'display:block' }, h('b', null, 'Registre de sécurité (DPMS, mensuel)'), ' ',
          rg.du ? h('span', { class: 'badge b-bad' }, 'à contrôler') : h('span', { class: 'badge b-ok' }, 'à jour'),
          h('div', { class: 'muted small' }, rg.dernier ? `Dernier : ${fmtDT(rg.dernier.at)}${rg.dernier.note ? ' — ' + rg.dernier.note : ''} — prochain avant le ${fmtJour(localDate(new Date(new Date(rg.dernier.at).getTime() + 30 * 86400000)))}` : 'Jamais contrôlé'),
          h('button', { class: 'sec', style: 'margin-top:6px', onclick: safe(async () => { if (await registreControle(s.id)) go('supControles', {}, { noPush: true, keepScroll: true }); }) }, 'Registre contrôlé')),
        ...Object.keys(NIVEAUX).map(n => {
          const e = echeance(st, s.id, n);
          const proch = e.dernier && n !== 'quotidien' ? new Date(new Date(e.dernier.at).getTime() + NIVEAUX[n].jours * 86400000) : null;
          return h('div', { class: 'stat', style: 'display:block' }, h('b', null, NIVEAUX[n].lbl), ' ',
            e.du ? h('span', { class: 'badge b-bad' }, n === 'quotidien' ? 'à faire aujourd’hui' : 'à faire') : h('span', { class: 'badge b-ok' }, 'à jour'),
            h('div', { class: 'muted small' }, e.dernier ? `Dernier : ${fmtDT(e.dernier.at)} par ${e.dernier.agent}${proch ? ` — prochain avant le ${fmtJour(localDate(proch))}` : ''}` : 'Jamais fait'));
        }),
        na.length ? h('div', { style: 'margin-top:8px' }, h('p', { class: 'small', style: 'font-weight:600;margin:4px 0' }, 'Points qui n’existent pas dans ce parking'),
          ...na.map(c => h('div', { class: 'stat' }, h('span', { class: 'small' }, `${c.lbl} (${NIVEAUX[c.niveau].titre.toLowerCase()})`),
            h('button', { class: 'link mini', onclick: safe(async () => { await queueDec('d', { type: 'ctl_reactiver', id: uuid(), ts: new Date().toISOString(), site: s.id, item: c.id }); await fold(); syncNow(); toast('Point rétabli.'); go('supControles', {}, { noPush: true, keepScroll: true }); }) }, 'Rétablir')))) : null);
    }),
    h('div', { class: 'card' }, h('h2', null, 'Historique'),
      hist.length ? hist.slice(0, 150).map(x => {
        const r = resumeControle(x.items, x.libelles);
        return h('div', { class: 'stat', style: 'display:block' },
          h('div', null, h('b', null, `${siteNom(x.site)} — ${NIVEAUX[x.niveau].lbl}`), ` — ${fmtDT(x.ts || x.at)} — ${x.agent}`, x.partiel ? h('span', { class: 'badge b-warn' }, 'partiel : ronde clôturée par le DPMS') : null),
          h('div', { class: 'small' }, `${r.c} conforme(s)`, r.nc.length ? h('span', { style: 'color:var(--bad)' }, ` · non conforme : ${r.nc.join(', ')}`) : null,
            r.na.length ? ` · n’existe pas : ${r.na.join(', ')}` : '', r.nf.length ? ` · non faits : ${r.nf.join(', ')}` : ''));
      }) : h('p', { class: 'muted' }, 'Aucun contrôle enregistré.')));
};
async function cloturerRonde(r) {
  const depuis = r.debut ? dureeTxt(Math.round((Date.now() - new Date(r.debut)) / 60000)) : '?';
  if (!(await demander({ titre: 'Clôturer cette ronde ?', oui: 'Oui, clôturer', non: 'Non',
    texte: `${siteNom(r.site)} — ${r.agent}, commencée le ${fmtDT(r.debut)} (il y a ${depuis}).\n\nÀ sa prochaine connexion, le téléphone qui porte la ronde l’arrête et envoie les contrôles déjà faits. Si l’agent termine sa ronde avant, c’est sa fin de ronde qui compte.` }))) return;
  await queueDec('d', { type: 'ronde_close', id: uuid(), ts: new Date().toISOString(), rid: r.rid, par: 'DPMS' });
  try { await rel('annulerDiffere', { id: await jeton('ronde', r.rid) }, 15000); } catch (e) { }   // plus d'alerte « ronde de plus d'1 h »
  await fold(); syncNow();
  toast('Ronde clôturée.'); render();
}
function rondeLigne(r) {
  const dur = r.debut && r.fin && !r.closDPMS ? Math.round((new Date(r.fin) - new Date(r.debut)) / 60000) : null;
  const anos = r.checklist ? Object.entries(r.checklist).filter(([, v]) => v !== 'RAS').map(([k, v]) => v === 'Anomalie' ? k : `${k} (${v.toLowerCase()})`) : [];
  return h('div', { class: 'stat', style: 'display:block' },
    h('div', null, h('b', null, `${siteNom(r.site)} — ${r.agent}`), !r.fin ? h('span', { class: 'badge b-warn' }, 'fin non reçue') : null, r.closDPMS ? h('span', { class: 'badge b-warn' }, 'clôturée par le DPMS') : null),
    !r.fin && state.mode === 'superviseur' ? h('button', { class: 'sec', style: 'width:auto;margin:6px 0', onclick: safe(() => cloturerRonde(r)) }, 'Clôturer la ronde') : null,
    h('div', { class: 'muted small' }, `${fmtDT(r.debut)}${r.fin ? ' → ' + fmtHeure(r.fin) : ''}${dur != null ? ` (${dur} min)` : ''} — ${r.compte != null ? r.compte + ' véhicule(s), ' : ''}${r.nbSig} signalement(s), ${r.nbRevue} constat(s)`),
    anos.length ? h('div', { class: 'small' }, 'Anomalies : ' + anos.join(', ')) : null,
    (r.controles || []).length ? h('div', { class: 'small' }, (r.controles || []).map(c => `${NIVEAUX[c.niveau].lbl}${c.nc.length ? ' — non conforme : ' + c.nc.join(', ') : ' — conforme'}`).join(' ; ')) : null);
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

/* ---------- Liste et fiche des signalements (sur PC : liste et fiche côte à côte) ---------- */
const ecranLarge = () => window.innerWidth >= 1000;
VIEWS.supSigs = async ({ f = 'ouvert', site = '', cat = '', q = '', sel = null }) => {
  const all = await dbAll('sigs');
  const lus = new Set((await kvGet(cleLus())) || []);
  const nl = new Map(); for (const x of all) nl.set(x.id, await nonLus(x, lus));
  const tests = { ouvert: s => s.statut === 'ouvert', a_transmettre: s => s.statut === 'ouvert' && suiviEtat(s) === 'À transmettre', clos: s => s.statut === 'clos', vehicules: s => isVehicule(s.cat), echanges: s => (s.messages || []).length > 0, tous: () => true };
  const qq = (q || '').toLowerCase();
  const list = all.filter(s => tests[f](s) && (!site || s.site === site) && (!cat || s.cat === cat)
    && (!qq || [s.ref, s.cat, s.sub, s.desc, s.plaque, s.emplacement, s.agent, s.vehicule].join(' ').toLowerCase().includes(qq)))
    .sort(f === 'echanges'
      ? (a, b) => (nl.get(b.id) > 0) - (nl.get(a.id) > 0) || !!b.demande - !!a.demande || ((b.messages.slice(-1)[0] || {}).quand || '').localeCompare((a.messages.slice(-1)[0] || {}).quand || '')
      : (a, b) => (b.urgent - a.urgent) || (b.ts || b.jour || '').localeCompare(a.ts || a.jour || ''));
  const nav = p => go('supSigs', { f, site, cat, q, sel, ...p }, { noPush: true, keepScroll: true });
  const large = ecranLarge();
  const filtres = h('div', { class: 'card' },
    h('div', { class: 'chips' }, ...[['ouvert', 'En cours'], ['a_transmettre', 'À transmettre'], ['clos', 'Clos'], ['vehicules', 'Véhicules'], ['echanges', 'Échanges'], ['tous', 'Tous']].map(([k, l]) =>
      h('button', { class: 'chip' + (f === k ? ' on' : ''), onclick: () => nav({ f: k, sel: null }) }, `${l} (${all.filter(tests[k]).length})`))),
    h('div', { class: 'filtres' },
      h('select', { onchange: e => nav({ site: e.target.value, sel: null }) }, h('option', { value: '' }, 'Tous les parkings'), ...state.cfg.sites.map(s => h('option', { value: s.id, selected: s.id === site ? true : null }, s.nom))),
      h('select', { onchange: e => nav({ cat: e.target.value, sel: null }) }, h('option', { value: '' }, 'Toutes les catégories'), ...[...state.cfg.cats.map(c => c.nom), CAT_AUTRE].map(c => h('option', { value: c, selected: c === cat ? true : null }, c))),
      h('input', { type: 'search', value: q, placeholder: 'Rechercher (plaque, réf., texte…)', onchange: e => nav({ q: e.target.value, sel: null }) })),
    list.length ? h('button', { class: 'link', style: 'color:var(--bad)', onclick: safe(async () => {
      if (prompt(`Supprimer les ${list.length} signalement(s) affiché(s) par ces filtres ?\nIls disparaîtront de tous les appareils, des statistiques et des exports (annulable depuis la corbeille).\nTapez SUPPRIMER pour confirmer.`) !== 'SUPPRIMER') return;
      await supprimer({ sigs: list.map(x => x.id), quoi: `${list.length} signalement(s) (filtre : ${{ ouvert: 'en cours', a_transmettre: 'à transmettre', clos: 'clos', vehicules: 'véhicules', echanges: 'échanges', tous: 'tous' }[f]}${site ? ', ' + siteNom(site) : ''}${cat ? ', ' + cat : ''}${q ? ', « ' + q + ' »' : ''})` });
      toast(`${list.length} signalement(s) supprimé(s).`); nav({ sel: null });
    }) }, `Supprimer les ${list.length} signalement(s) affiché(s)…`) : null);
  const items = [];
  for (const s of list) {
    const th = await thumbOf(s);
    items.push(h('button', { class: 'list-item' + (sel === s.id ? ' sel' : ''), onclick: () => large ? nav({ sel: s.id }) : go('supSig', { id: s.id }) },
      h('div', { style: 'display:flex;gap:10px;align-items:flex-start' },
        th ? h('img', { src: th, alt: '', class: 'vignette' }) : h('div', { class: 'vignette vide' }),
        h('div', { style: 'flex:1;min-width:0' },
          h('div', { class: 'l1' }, sigTitre(s)),
          h('div', { class: 'small' }, statutBadge(s.statut), s.statut === 'ouvert' ? suiviBadge(s) : null, s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, isVehicule(s.cat) ? pmBadge(s) : null,
            (s.messages || []).length ? icoCourrier(nl.get(s.id), nl.get(s.id) > 0) : null, s.demande ? h('span', { class: 'badge b-warn' }, 'Action demandée') : null),
          h('div', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ts ? fmtDT(s.ts) : fmtJour(s.jour)} — ${s.agent || ''}${s.horsRonde ? ' (hors ronde)' : ''}`),
          s.plaque ? h('div', { class: 'small' }, h('span', { class: 'plaque' }, s.plaque), s.emplacement ? ' ' + s.emplacement : '') : null))));
  }
  const vide = h('p', { class: 'muted' }, 'Aucun signalement.');
  if (!large) return page('Signalements', { back: true }, filtres, items.length ? items : vide);
  const choisi = sel ? await dbGet('sigs', sel) : null;
  return page('Signalements', { back: true }, filtres,
    h('div', { class: 'split' },
      h('div', { class: 'pane-list' }, items.length ? items : vide),
      h('div', { class: 'pane-detail' }, choisi ? await supSigDetail(choisi, () => nav({})) : h('div', { class: 'card muted' }, 'Sélectionnez un signalement dans la liste.'))));
};
function geoLink(geo) {
  if (!geo || geo.err || geo.lat == null) return h('span', { class: 'muted' }, geo && geo.err ? 'position ' + geo.err : '—');
  return h('a', { href: `https://www.openstreetmap.org/?mlat=${geo.lat}&mlon=${geo.lon}#map=19/${geo.lat}/${geo.lon}`, target: '_blank', rel: 'noopener' },
    `${geo.lat}, ${geo.lon} (± ${geo.acc} m)`);
}
const aGeo = g => g && !g.err && g.lat != null;
/* Photos : téléchargées à la demande depuis le relais, vérifiées, conservées sur l'appareil */
const _mediaEnCours = new Map();   // un seul téléchargement par lot, même si plusieurs écrans le demandent
function fetchMedia(n, delai = 180000) {
  if (_mediaEnCours.has(n)) return _mediaEnCours.get(n);
  const p = fetchMedia_(n, delai).finally(() => _mediaEnCours.delete(n));
  _mediaEnCours.set(n, p);
  return p;
}
async function fetchMedia_(n, delai) {
  let r;
  try { r = await rel('media', { n }, delai); }
  catch (e) { if (/interdit/.test(e.message)) throw new Error('le relais doit être mis à jour (version 3) pour afficher les photos sur cet appareil'); throw e; }
  const z = fflate.unzipSync(await teamDec(b64u.dec(r.data)));
  const out = [];
  for (const [name, bytes] of Object.entries(z)) out.push({ id: name.replace(/\.jpg$/, ''), blob: new Blob([bytes], { type: 'image/jpeg' }), sha: await sha256hex(bytes) });
  await dbPutMany('photos', out);
  return out;
}
/* Photos sans bloquer l'affichage : déjà sur l'appareil → affichées tout de suite ;
   sinon la fiche s'affiche immédiatement (vignette si connue) et les photos arrivent ensuite */
async function photosEnCache(list) { for (const p of list || []) if (!(await dbGet('photos', p.id))) return false; return true; }
function photosDiffere(list, opts = {}) {
  const attente = opts.vignette
    ? h('div', { class: 'gallery' + (opts.grande ? ' grande' : '') }, h('img', { src: opts.vignette, alt: 'Photo du signalement', onclick: () => lightbox([opts.vignette], 0) }))
    : null;
  const box = h('div', { class: 'photos-diff' }, attente, h('p', { class: 'muted small' }, 'Chargement des photos…'));
  photosDe(list, opts).then(({ photos, erreurs }) => {
    box.replaceChildren(...[photos.length ? galerie(photos, opts.grande) : attente, erreurs.length ? h('p', { class: 'muted small' }, 'Photos indisponibles : ' + erreurs[0]) : null].filter(Boolean));
  }).catch(e => box.replaceChildren(...[attente, h('p', { class: 'muted small' }, 'Photos indisponibles : ' + e.message)].filter(Boolean)));
  return box;
}
async function photoBlocks(list, opts = {}) {
  if (!(list || []).length) return opts.vignette ? [h('div', { class: 'gallery' + (opts.grande ? ' grande' : '') }, h('img', { src: opts.vignette, alt: 'Photo du signalement', onclick: () => lightbox([opts.vignette], 0) }))] : [];
  if (await photosEnCache(list)) {
    const { photos, erreurs } = await photosDe(list);
    return [galerie(photos, opts.grande), erreurs.length ? h('p', { class: 'muted small' }, 'Photos indisponibles : ' + erreurs[0]) : null];
  }
  return [photosDiffere(list, opts)];
}
async function supDecision(s, action, patch) {
  await queueDec('d', { type: 'sig', id: uuid(), ts: new Date().toISOString(), par: 'DPMS', action, sig: s.id, patch });
  await fold();
  syncNow();
}
/* Contenu de la fiche (pleine page sur téléphone, panneau de droite sur PC) */
async function supSigDetail(s, recharger) {
  const done = (msg) => { toast(msg); recharger(); };
  const notes = h('textarea', { value: s.notes || '', placeholder: 'Notes internes DPMS (lisibles par les seuls superviseurs)' });
  const suivi = [];
  for (const v of s.suivi || []) {
    suivi.push(h('div', { class: 'stat', style: 'display:block' },
      h('div', null, h('b', null, VERDICTS[v.verdict] || v.verdict), ` — ${v.ts ? fmtDT(v.ts) : fmtJour(v.jour)} — ${v.agent || ''}`, v.sansEffet ? h('span', { class: 'badge b-warn' }, 'sans effet sur le statut : décision du DPMS non connue de l’agent') : null),
      v.comment ? h('div', { class: 'small' }, v.comment) : null,
      aGeo(v.geo) ? h('div', { class: 'small' }, 'Position : ', geoLink(v.geo)) : null,
      ...(await photoBlocks(v.photos))));
  }
  const form = actionForm(s, {
    agents: null,
    onSave: async (a) => { await queueDec('d', { type: 'action', id: uuid(), ts: new Date().toISOString(), par: 'DPMS', action: a }); await fold(); syncNow(); done('Action enregistrée.'); },
    onClore: async () => { await supDecision(s, 'clôture', { statut: 'clos', closLe: new Date().toISOString(), closPar: 'DPMS' }); done('Clos.'); }
  });
  return h('div', null,
    h('div', { class: 'card' },
      h('h3', null, sigTitre(s), ' ', statutBadge(s.statut), s.statut === 'ouvert' ? suiviBadge(s) : null, s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, s.aggrave ? h('span', { class: 'badge b-bad' }, 'Aggravé') : null),
      h('dl', { class: 'kv' },
        h('dt', null, 'Référence'), h('dd', null, s.ref),
        h('dt', null, 'Parking'), h('dd', null, siteNom(s.site)),
        h('dt', null, 'Signalé'), h('dd', null, s.ts ? fmtDT(s.ts) : fmtJour(s.jour), s.agent ? ' par ' + s.agent : '', s.horsRonde ? ' (hors ronde)' : ''),
        aGeo(s.geo) ? [h('dt', null, 'Position'), h('dd', null, geoLink(s.geo))] : null,
        h('dt', null, 'À signaler à'), h('dd', null, (s.dests || []).join(', ')),
        s.plaque ? [h('dt', null, 'Plaque'), h('dd', null, h('span', { class: 'plaque' }, s.plaque))] : null,
        s.emplacement ? [h('dt', null, 'Emplacement'), h('dd', null, s.emplacement)] : null,
        s.vehicule ? [h('dt', null, 'Véhicule'), h('dd', null, s.vehicule)] : null,
        s.statut === 'clos' ? [h('dt', null, 'Clos'), h('dd', null, `${s.closLe && s.closLe.length > 10 ? fmtDT(s.closLe) : fmtJour(s.closLe)}${s.closPar ? ' par ' + s.closPar : ''}`)] : null,
        s.notesAgents ? [h('dt', null, 'Message aux agents'), h('dd', null, s.notesAgents)] : null),
      s.desc ? h('p', null, s.desc) : null,
      ...(await photoBlocks(s.photos, { vignette: await thumbOf(s) }))),
    await carteEchangesSup(s, recharger),
    h('div', { class: 'card' }, h('h2', null, 'Traitement'), ...journalActions(s)),
    s.pm && s.pm.hist.length ? h('div', { class: 'card' }, h('h2', null, 'Police municipale'),
      ...s.pm.hist.map(x => h('p', { class: 'small' }, h('b', null, `${x.ts ? fmtDT(x.ts) : fmtJour(x.jour)} — ${PM_STATUTS[x.statut]}`), x.obs ? ' : ' + x.obs : ''))) : null,
    suivi.length ? h('div', { class: 'card' }, h('h2', null, 'Constats des agents'), ...suivi) : null,
    s.statut === 'ouvert' ? contactsBloc(s) : null,
    form,
    h('div', { class: 'card' }, h('h2', null, 'Intervenir'),
      h('p', { class: 'muted small' }, 'Toute intervention est transmise automatiquement aux agents, à la PM et aux autres superviseurs.'),
      s.statut === 'clos' ? h('button', { class: 'sec', onclick: safe(async () => { await supDecision(s, 'réouverture', { statut: 'ouvert', closLe: null, closPar: null }); done('Rouvert.'); }) }, 'Rouvrir') : null,
      h('button', { class: 'sec', onclick: () => { state.retour = { view: state.view, params: { ...state.params } }; go('supSig', { id: s.id, edit: true }); } }, 'Modifier le signalement'),
      h('button', { class: 'danger', onclick: safe(async () => {
        if (!confirm(`Supprimer le signalement ${s.ref} (${sigTitre(s)}) ?\nIl disparaîtra de tous les appareils, des statistiques et des exports. Annulable depuis la corbeille.`)) return;
        await supprimer({ sigs: [s.id], quoi: `Signalement ${s.ref} — ${sigTitre(s)} (${siteNom(s.site)}, ${fmtJour(s.jour)})` });
        toast('Signalement supprimé.');
        if (state.view === 'supSigs') go('supSigs', { ...state.params, sel: null }, { noPush: true }); else go('supSigs', { f: 'ouvert' }, { noPush: true });
      }) }, 'Supprimer ce signalement'),
      s.plaque && s.statut === 'clos' ? h('button', { class: 'link', onclick: safe(async () => { if (confirm('Effacer la plaque de ce signalement clos ?')) { await supDecision(s, 'effacement de la plaque', { plaque: '' }); done('Plaque effacée.'); } }) }, 'Effacer la plaque (dossier traité)') : null),
    h('div', { class: 'card' }, h('h2', null, 'Notes internes'), notes,
      h('button', { class: 'sec', onclick: safe(async () => { await queueDec('n', { type: 'note', ts: new Date().toISOString(), sig: s.id, text: notes.value }); await fold(); syncNow(); toast('Notes enregistrées.'); }) }, 'Enregistrer les notes')),
    (s.journalDPMS || []).length ? h('div', { class: 'card' }, h('h2', null, 'Interventions DPMS'),
      ...s.journalDPMS.map(j => h('p', { class: 'small' }, `${fmtDT(j.ts)} — ${j.action}`))) : null);
}
VIEWS.supSig = async ({ id, edit }) => {
  const s = await dbGet('sigs', id);
  if (!s) return page('Signalement', { back: true }, h('p', null, 'Introuvable.'));
  if (edit) {
    const catSel = h('select', null, ...[...state.cfg.cats.map(c => c.nom), CAT_AUTRE].map(c => h('option', { value: c, selected: c === s.cat ? true : null }, c)));
    const sub = h('input', { type: 'text', value: s.sub || '' });
    const desc = h('textarea', { value: s.desc || '' });
    const plaque = h('input', { type: 'text', value: s.plaque || '' });
    const empl = h('input', { type: 'text', value: s.emplacement || '' });
    const dests = [...(s.dests || [])];
    let urgent = !!s.urgent;
    const urg = h('button', { class: 'chip' + (urgent ? ' on-bad' : ''), onclick: e => { urgent = !urgent; e.target.className = 'chip' + (urgent ? ' on-bad' : ''); } }, 'Urgent');
    const retour = () => { const r = state.retour || { view: 'supSig', params: { id } }; state.retour = null; go(r.view, r.params, { noPush: true }); };
    return page(`Modifier ${s.ref}`, { back: retour },
      h('label', { class: 'f' }, 'Catégorie'), catSel,
      h('label', { class: 'f' }, 'Sous-catégorie'), sub,
      h('label', { class: 'f' }, 'Description'), desc,
      h('label', { class: 'f' }, 'Plaque'), plaque,
      h('label', { class: 'f' }, 'Emplacement'), empl,
      h('label', { class: 'f' }, 'À signaler à'), chips(state.cfg.dests, dests, () => { }, { multi: true }),
      h('label', { class: 'f' }, 'Priorité'), h('div', { class: 'chips' }, urg),
      h('button', {
        class: 'ok', onclick: safe(async () => {
          await supDecision(s, 'modification', { cat: catSel.value, sub: sub.value.trim(), desc: desc.value.trim(), plaque: plaque.value.trim().toUpperCase(), emplacement: empl.value.trim(), dests: [...dests], urgent });
          toast('Modifié. Transmis aux agents.'); retour();
        })
      }, 'Enregistrer et transmettre aux agents'));
  }
  return page(`Signalement ${s.ref}`, { back: true }, await supSigDetail(s, () => go('supSig', { id }, { noPush: true, keepScroll: true })));
};

/* ====================================================================== */
/* STATISTIQUES ET GRAPHIQUES (superviseur)                               */
/* Palette catégorielle validée (ordre fixe) : bleu, orange, aqua.        */
/* ====================================================================== */
const SERIES = ['#2a78d6', '#eb6834', '#1baf7a'];
const INDICS = {
  sig_n: { lbl: 'Nombre de signalements', src: 'sig' },
  sig_statut: { lbl: 'Signalements en cours / clos', src: 'sig' },
  sig_delai: { lbl: 'Délai moyen de clôture (jours)', src: 'sig' },
  act_n: { lbl: 'Actions de traitement', src: 'act' },
  ctl_n: { lbl: 'Contrôles réalisés', src: 'ctl' },
  ctl_nc: { lbl: 'Non-conformités relevées', src: 'ctlnc' },
  ronde_n: { lbl: 'Nombre de rondes', src: 'ronde' },
  ronde_veh: { lbl: 'Véhicules comptés (moyenne par ronde)', src: 'ronde' },
  barr_n: { lbl: 'Ouvertures de barrière', src: 'barr' },
  barr_h: { lbl: 'Durée d’ouverture des barrières (heures)', src: 'barr' }
};
const AXES = {
  cat: { lbl: 'Catégorie', src: ['sig', 'act'], key: r => r.cat },
  sub: { lbl: 'Sous-catégorie', src: ['sig'], key: r => r.sub ? `${r.sub}` : `${r.cat} (sans précision)` },
  agent: { lbl: 'Agent', src: ['sig', 'act', 'ronde', 'barr', 'ctl', 'ctlnc'], key: r => r.agent || '?' },
  site: { lbl: 'Parking', src: ['sig', 'act', 'ronde', 'barr', 'ctl', 'ctlnc'], key: r => siteNom(r.site) },
  point: { lbl: 'Point de contrôle', src: ['ctlnc'], key: r => r.point },
  niveau: { lbl: 'Type de contrôle', src: ['ctl', 'ctlnc'], key: r => NIVEAUX[r.niveau].lbl },
  dest: { lbl: 'Destinataire', src: ['sig', 'act'], key: r => r.dests && r.dests.length ? r.dests : ['(aucun)'] },
  type: { lbl: 'Type d’action', src: ['act'], key: r => ACTIONS[r.type] || r.type },
  barriere: { lbl: 'Barrière', src: ['barr'], key: r => `${siteNom(r.site)} — ${r.barriere}` },
  motif: { lbl: 'Motif d’ouverture', src: ['barr'], key: r => r.motif || '?' },
  jour: { lbl: 'Jour', src: ['sig', 'act', 'ronde', 'barr', 'ctl', 'ctlnc'], temps: 'jour' },
  semaine: { lbl: 'Semaine', src: ['sig', 'act', 'ronde', 'barr', 'ctl', 'ctlnc'], temps: 'semaine' },
  mois: { lbl: 'Mois', src: ['sig', 'act', 'ronde', 'barr', 'ctl', 'ctlnc'], temps: 'mois' }
};
const TABLEAU_DEFAUT = [
  { i: 'sig_statut', a: 'cat' }, { i: 'sig_n', a: 'agent' }, { i: 'sig_n', a: 'semaine' },
  { i: 'ctl_nc', a: 'point' }, { i: 'sig_delai', a: 'cat' }, { i: 'barr_h', a: 'barriere' }, { i: 'ronde_veh', a: 'site' }
];
const MOIS = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];
function dSig(s) { return new Date(s.ts || ((s.jour || '').slice(0, 10) + 'T12:00:00')); }
function dClos(s) { return s.closLe ? new Date(s.closLe.length > 10 ? s.closLe : s.closLe + 'T12:00:00') : null; }
function cleTemps(d, mode) {
  if (mode === 'mois') return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
  if (mode === 'semaine') { const l = new Date(d); l.setHours(12, 0, 0, 0); l.setDate(l.getDate() - ((l.getDay() + 6) % 7)); return localDate(l); }
  return localDate(d);
}
function lblTemps(k, mode) {
  if (mode === 'mois') { const [y, m] = k.split('-'); return `${MOIS[+m - 1]} ${y}`; }
  const [y, m, d] = k.split('-'); return mode === 'semaine' ? `sem. du ${d}/${m}` : `${d}/${m}`;
}
function suiteTemps(min, max, mode) {
  const out = []; const d = new Date(min); d.setHours(12, 0, 0, 0);
  if (mode === 'mois') d.setDate(1);
  if (mode === 'semaine') d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
  for (let g = 0; d <= max && g < 800; g++) {
    out.push(cleTemps(d, mode));
    if (mode === 'mois') d.setMonth(d.getMonth() + 1); else d.setDate(d.getDate() + (mode === 'semaine' ? 7 : 1));
  }
  return [...new Set(out)];
}
async function donneesStats({ du, au, site }) {
  const dans = d => d && !isNaN(d) && (!du || d >= du) && (!au || d <= au);
  const okSite = r => !site || r.site === site;
  const sigs = (await dbAll('sigs')).filter(okSite);
  const sig = sigs.map(s => {
    const c = dClos(s), o = dSig(s);
    return { date: o, cat: s.cat, sub: s.sub, agent: s.agent, site: s.site, dests: s.dests, statut: s.statut, delai: s.statut === 'clos' && c ? Math.max(0, (c - o) / 86400000) : null };
  }).filter(r => dans(r.date));
  const act = sigs.flatMap(s => (s.actions || []).map(a => ({ date: new Date(a.quand || a.jour + 'T12:00:00'), type: a.type, dests: a.dest ? [a.dest] : [], agent: a.par, site: s.site, cat: s.cat }))).filter(r => dans(r.date));
  const ronde = (await dbAll('rondes')).filter(okSite).map(r => ({ date: new Date(r.debut || r.jour + 'T12:00:00'), agent: r.agent, site: r.site, compte: r.compte })).filter(r => dans(r.date));
  const barr = ((await kvGet('barrieres')) || []).filter(okSite).map(b => {
    const o = new Date(b.ouverte.quand), f = b.fermee ? new Date(b.fermee.quand) : new Date();
    return { date: o, site: b.site, barriere: b.barriere, agent: b.ouverte.agent, motif: b.ouverte.motif, heures: Math.max(0, (f - o) / 3600000), enCours: !b.fermee };
  }).filter(r => dans(r.date));
  const ctl = ((await kvGet('controlesHist')) || []).filter(okSite).map(x => ({ date: new Date(x.ts || x.at), site: x.site, agent: x.agent, niveau: x.niveau, items: x.items, libelles: x.libelles })).filter(r => dans(r.date));
  const ctlnc = ctl.flatMap(x => resumeControle(x.items, x.libelles).nc.map(point => ({ ...x, point })));
  return { sig, act, ronde, barr, ctl, ctlnc };
}
const moyenne = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : null;
const arrondi = (v, n = 1) => v == null ? null : Math.round(v * 10 ** n) / 10 ** n;
function agreger(spec, D) {
  const ind = INDICS[spec.i], ax = AXES[spec.a];
  const recs = D[ind.src];
  const groupes = new Map();
  for (const r of recs) {
    let ks = ax.temps ? cleTemps(r.date, ax.temps) : ax.key(r);
    if (!Array.isArray(ks)) ks = [ks];
    for (const k of ks) { if (!groupes.has(k)) groupes.set(k, []); groupes.get(k).push(r); }
  }
  let cles = [...groupes.keys()];
  if (ax.temps && recs.length) {
    const ds = recs.map(r => r.date.getTime());
    cles = suiteTemps(new Date(Math.min(...ds)), new Date(Math.max(...ds)), ax.temps);
  }
  const val = (rs) => {
    rs = rs || [];
    switch (spec.i) {
      case 'sig_statut': return [rs.filter(r => r.statut === 'ouvert').length, rs.filter(r => r.statut === 'clos').length];
      case 'sig_delai': return [arrondi(moyenne(rs.map(r => r.delai).filter(x => x != null)))];
      case 'ronde_veh': return [arrondi(moyenne(rs.map(r => r.compte).filter(x => x != null)))];
      case 'barr_h': return [arrondi(rs.reduce((s, r) => s + r.heures, 0))];
      default: return [rs.length];
    }
  };
  let lignes = cles.map(k => ({ k, lbl: ax.temps ? lblTemps(k, ax.temps) : k, v: val(groupes.get(k)) }));
  let note = '';
  if (!ax.temps) {
    lignes = lignes.filter(l => l.v.some(x => x != null && x !== 0));
    lignes.sort((a, b) => b.v.reduce((s, x) => s + (x || 0), 0) - a.v.reduce((s, x) => s + (x || 0), 0));
    if (lignes.length > 15) { note = `${lignes.length - 15} autres valeurs non affichées`; lignes = lignes.slice(0, 15); }
  }
  const series = spec.i === 'sig_statut' ? ['En cours', 'Clos'] : [ind.lbl];
  return { lignes, series, note, temps: !!ax.temps, titre: `${ind.lbl} par ${ax.lbl.toLowerCase()}`, axe: ax.lbl, unite: spec.i === 'sig_delai' ? ' j' : spec.i === 'barr_h' ? ' h' : '' };
}
let _charts = [];
function detruireCharts() { _charts.forEach(c => { try { c.destroy(); } catch (e) { } }); _charts = []; }
function carteGraphique(spec, D, onRetirer) {
  const g = agreger(spec, D);
  const carte = h('div', { class: 'card chart-card' }, h('h3', null, g.titre));
  if (!g.lignes.length) { carte.append(h('p', { class: 'muted' }, 'Aucune donnée sur la période.')); }
  else {
    const empile = g.series.length > 1;
    const ligne = g.temps && g.lignes.length > 12 && !empile;
    const horizontal = !g.temps;
    const hauteur = horizontal ? Math.max(160, 44 + g.lignes.length * (empile ? 30 : 26)) : 260;
    const canvas = h('canvas');
    const zone = h('div', { style: `position:relative;height:${hauteur}px` }, canvas);
    const table = h('table', { class: 'tableau', style: 'display:none' },
      h('thead', null, h('tr', null, h('th', null, g.axe), ...g.series.map(s => h('th', null, s)))),
      h('tbody', null, ...g.lignes.map(l => h('tr', null, h('td', null, l.lbl), ...l.v.map(x => h('td', null, x == null ? '—' : String(x).replace('.', ',') + g.unite))))));
    carte.append(zone, table);
    const datasets = g.series.map((nom, j) => ({
      label: nom, data: g.lignes.map(l => l.v[j]), backgroundColor: SERIES[j], borderColor: ligne ? SERIES[j] : '#ffffff',
      borderWidth: ligne ? 2 : (empile ? { top: 0, bottom: 0, left: 0, right: 2 } : 0), borderRadius: empile ? 0 : 4, borderSkipped: 'start',
      maxBarThickness: 22, pointRadius: 4, pointBackgroundColor: SERIES[j], pointBorderColor: '#ffffff', pointBorderWidth: 2, tension: 0, spanGaps: true
    }));
    const grille = { color: '#e8ebf0', drawTicks: false }, ticks = { color: '#5d6b7c', padding: 6, font: { size: 12 } };
    const entier = ['sig_n', 'sig_statut', 'act_n', 'ronde_n', 'barr_n', 'ctl_n', 'ctl_nc'].includes(spec.i);
    const ticksValeur = { ...ticks, precision: entier ? 0 : undefined, callback: v => String(v).replace('.', ',') };
    const chart = new Chart(canvas, {
      type: ligne ? 'line' : 'bar',
      data: { labels: g.lignes.map(l => l.lbl), datasets },
      options: {
        indexAxis: horizontal ? 'y' : 'x', responsive: true, maintainAspectRatio: false, animation: false,
        interaction: { mode: horizontal ? 'y' : 'index', intersect: false },
        plugins: {
          legend: { display: g.series.length > 1, position: 'top', align: 'start', labels: { boxWidth: 12, boxHeight: 12, color: '#1b2430' } },
          tooltip: { callbacks: { label: c => `${c.dataset.label} : ${c.parsed[horizontal ? 'x' : 'y'] == null ? '—' : String(c.parsed[horizontal ? 'x' : 'y']).replace('.', ',') + g.unite}` } }
        },
        scales: {
          x: { stacked: empile, grid: horizontal ? grille : { display: false }, border: { color: '#cfd5de' }, ticks: horizontal ? ticksValeur : { ...ticks, maxRotation: 0, autoSkip: true }, beginAtZero: true },
          y: { stacked: empile, grid: horizontal ? { display: false } : grille, border: { color: '#cfd5de' }, ticks: horizontal ? { ...ticks, autoSkip: false } : ticksValeur, beginAtZero: true }
        }
      }
    });
    _charts.push(chart);
    carte.append(h('div', { class: 'chart-actions' },
      g.note ? h('span', { class: 'muted small' }, g.note) : null,
      h('button', { class: 'link mini', onclick: e => { const vis = table.style.display !== 'none'; table.style.display = vis ? 'none' : ''; zone.style.display = vis ? '' : 'none'; e.target.textContent = vis ? 'Tableau' : 'Graphique'; } }, 'Tableau'),
      h('button', { class: 'link mini', onclick: () => { const a = document.createElement('a'); a.href = chart.toBase64Image('image/png', 1); a.download = g.titre.replace(/[^\wÀ-ÿ]+/g, '_') + '.png'; a.click(); } }, 'Image'),
      h('button', { class: 'link mini', onclick: safe(() => shareOrDownload(xlsxBlob([['Données', [[g.axe, ...g.series], ...g.lignes.map(l => [l.lbl, ...l.v])], [30, ...g.series.map(() => 16)]]]), g.titre.replace(/[^\wÀ-ÿ]+/g, '_') + '.xlsx', g.titre)) }, 'Excel'),
      onRetirer ? h('button', { class: 'link mini', onclick: onRetirer }, 'Retirer') : null));
  }
  return carte;
}
VIEWS.supStats = async (p = {}) => {
  detruireCharts();
  state.cleanup = detruireCharts;
  const per = p.per || '30';
  const maintenant = new Date();
  let du = null, au = null;
  if (per === 'perso') { du = p.du ? new Date(p.du + 'T00:00:00') : null; au = p.au ? new Date(p.au + 'T23:59:59') : null; }
  else if (per !== 'tout') { du = new Date(maintenant.getTime() - (+per) * 86400000); }
  const D = await donneesStats({ du, au, site: p.site || '' });
  const nav = q => go('supStats', { ...p, ...q }, { noPush: true, keepScroll: true });
  const clos = D.sig.filter(r => r.statut === 'clos');
  const delai = moyenne(clos.map(r => r.delai).filter(x => x != null));
  const veh = moyenne(D.ronde.map(r => r.compte).filter(x => x != null));
  const hBarr = D.barr.reduce((s, r) => s + r.heures, 0);
  const tuile = (v, l, sous) => h('div', { class: 'kpi' }, h('b', null, v), h('span', null, l), sous ? h('small', null, sous) : null);
  const perso = (await kvGet('statsPerso')) || [];

  // Constructeur de graphique
  const choix = { i: 'sig_n', a: 'cat' };
  const selAxe = h('select');
  const majAxes = () => { const src = INDICS[choix.i].src; const ok = Object.entries(AXES).filter(([, a]) => a.src.includes(src)); if (!ok.some(([k]) => k === choix.a)) choix.a = ok[0][0]; selAxe.replaceChildren(...ok.map(([k, a]) => h('option', { value: k, selected: k === choix.a ? true : null }, a.lbl))); };
  selAxe.onchange = e => choix.a = e.target.value;
  const selInd = h('select', { onchange: e => { choix.i = e.target.value; majAxes(); } }, ...Object.entries(INDICS).map(([k, x]) => h('option', { value: k }, x.lbl)));
  majAxes();

  const grille = h('div', { class: 'charts' });
  const node = page('Statistiques', { back: true },
    h('div', { class: 'card' },
      h('div', { class: 'chips' }, ...[['7', '7 jours'], ['30', '30 jours'], ['90', '3 mois'], ['365', '12 mois'], ['tout', 'Tout'], ['perso', 'Dates…']].map(([k, l]) =>
        h('button', { class: 'chip' + (per === k ? ' on' : ''), onclick: () => nav({ per: k }) }, l))),
      h('div', { class: 'filtres' },
        per === 'perso' ? [h('input', { type: 'date', value: p.du || '', onchange: e => nav({ du: e.target.value }) }), h('input', { type: 'date', value: p.au || '', onchange: e => nav({ au: e.target.value }) })] : null,
        h('select', { onchange: e => nav({ site: e.target.value }) }, h('option', { value: '' }, 'Tous les parkings'), ...state.cfg.sites.map(s => h('option', { value: s.id, selected: s.id === p.site ? true : null }, s.nom))))),
    h('div', { class: 'kpis' },
      tuile(D.sig.length, 'signalements'),
      tuile(D.sig.length - clos.length, 'en cours'),
      tuile(clos.length, 'clos', D.sig.length ? `${Math.round(clos.length / D.sig.length * 100)} %` : null),
      tuile(delai == null ? '—' : String(arrondi(delai)).replace('.', ',') + ' j', 'délai moyen de clôture'),
      tuile(D.ronde.length, 'rondes'),
      tuile(D.ctl.length, 'contrôles réalisés', `${D.ctlnc.length} non-conformité(s)`),
      tuile(veh == null ? '—' : String(arrondi(veh)).replace('.', ','), 'véhicules par ronde (moy.)'),
      tuile(D.barr.length, 'ouvertures de barrière', D.barr.some(r => r.enCours) ? 'dont en cours' : null),
      tuile(String(arrondi(hBarr)).replace('.', ',') + ' h', 'barrières ouvertes (cumul)')),
    h('div', { class: 'card' }, h('h2', null, 'Créer un graphique'),
      h('div', { class: 'filtres' }, selInd, h('span', { class: 'muted' }, 'par'), selAxe,
        h('button', { class: 'ok', style: 'width:auto;margin:0', onclick: safe(async () => { const l = (await kvGet('statsPerso')) || []; l.unshift({ ...choix }); await kvSet('statsPerso', l.slice(0, 30)); nav({}); }) }, 'Ajouter'))),
    grille);
  perso.forEach((spec, k) => grille.append(carteGraphique(spec, D, safe(async () => { const l = (await kvGet('statsPerso')) || []; l.splice(k, 1); await kvSet('statsPerso', l); nav({}); }))));
  TABLEAU_DEFAUT.forEach(spec => grille.append(carteGraphique(spec, D, null)));
  return node;
};

/* ---------- Historique des rondes ---------- */
VIEWS.supRondes = async () => {
  const toutes = (await dbAll('rondes')).sort((a, b) => (b.debut || '').localeCompare(a.debut || ''));
  const rs = toutes.slice(0, 300);
  const recharger = () => go('supRondes', {}, { noPush: true, keepScroll: true });
  return page('Rondes', { back: true },
    toutes.length ? h('div', { class: 'card' },
      h('p', { class: 'muted small' }, 'Supprimer une ronde efface ses horaires, son comptage et ses contrôles. Les signalements faits pendant la ronde sont conservés (ils se suppriment à part). Annulable depuis la corbeille.'),
      h('button', { class: 'danger', onclick: safe(async () => {
        if (prompt(`Supprimer les ${toutes.length} ronde(s) ?\nTapez SUPPRIMER pour confirmer.`) !== 'SUPPRIMER') return;
        await supprimer({ rids: toutes.map(r => r.rid), quoi: `Toutes les rondes (${toutes.length})` });
        toast('Rondes supprimées.'); recharger();
      }) }, `Supprimer toutes les rondes (${toutes.length})`)) : null,
    rs.length ? rs.map(r => h('div', { class: 'card' }, rondeLigne(r), aGeo(r.geoDebut) ? h('div', { class: 'small' }, 'Position au départ : ', geoLink(r.geoDebut)) : null,
      r.alertes && r.alertes.length ? h('div', { class: 'small', style: 'color:var(--warn)' }, [...new Set(r.alertes)].join(' ; ')) : null,
      h('button', { class: 'link', style: 'color:var(--bad);text-align:left;justify-content:flex-start', onclick: safe(async () => {
        if (!confirm(`Supprimer la ronde de ${r.agent} à ${siteNom(r.site)} du ${fmtDT(r.debut)} ?`)) return;
        await supprimer({ rids: [r.rid], quoi: `Ronde ${siteNom(r.site)} — ${r.agent} — ${fmtDT(r.debut)}` });
        toast('Ronde supprimée.'); recharger();
      }) }, 'Supprimer cette ronde')))
      : h('p', { class: 'muted' }, 'Aucune ronde reçue.'));
};
async function supprimer({ sigs = [], rids = [], quoi }) {
  await queueDec('d', { type: 'suppr', id: uuid(), ts: new Date().toISOString(), sigs, rids, quoi });
  await fold();
  syncNow();
}
VIEWS.supCorbeille = async () => {
  const l = (await kvGet('corbeille')) || [];
  return page('Corbeille', { back: true },
    h('p', { class: 'muted small' }, 'Éléments supprimés, du plus récent au plus ancien. « Restaurer » les fait réapparaître sur tous les appareils.'),
    l.length ? l.map(d => h('div', { class: 'card', style: d.annule ? 'opacity:.6' : '' },
      h('div', null, h('b', null, d.quoi || 'Suppression'), d.annule ? h('span', { class: 'badge b-ok' }, 'restauré') : null),
      h('div', { class: 'muted small' }, `Supprimé le ${fmtDT(d.ts)}`),
      d.annule ? null : h('button', { class: 'sec', onclick: safe(async () => {
        await queueDec('d', { type: 'annul_suppr', id: uuid(), ts: new Date().toISOString(), cible: d.id });
        await fold(); syncNow(); toast('Restauré.'); go('supCorbeille', {}, { noPush: true });
      }) }, 'Restaurer'))) : h('p', { class: 'muted' }, 'La corbeille est vide.'));
};

/* ---------- Excel complet ---------- */
async function supExcel() {
  showBusy('Construction du classeur…');
  const sigs = (await dbAll('sigs')).sort((a, b) => (a.ts || a.jour || '').localeCompare(b.ts || b.jour || ''));
  const rondes = (await dbAll('rondes')).sort((a, b) => (a.debut || '').localeCompare(b.debut || ''));
  const barr = ((await kvGet('barrieres')) || []).sort((a, b) => a.ouverte.quand.localeCompare(b.ouverte.quand));
  const sp = (await kvGet('supprimes')) || { sigs: [], rids: [] };
  const ds = new Set(sp.sigs), dr = new Set(sp.rids);
  const evs = (await dbAll('entries')).filter(e => e.kind === 'ev').filter(({ ev }) => !((ev.t === 'signalement' && ds.has(ev.data.id)) || (['revue', 'action', 'pm', 'msg'].includes(ev.t) && ds.has(ev.data.sig)) || (ev.rid && dr.has(ev.rid) && !['signalement', 'revue'].includes(ev.t)))).sort((a, b) => a.k.localeCompare(b.k));
  const g = (geo, k) => geo && !geo.err && geo[k] != null ? geo[k] : '';
  const A1 = [['Réf.', 'Parking', 'Catégorie', 'Sous-catégorie', 'Description', 'Plaque', 'Emplacement', 'Véhicule', 'Destinataires', 'Urgent', 'Aggravé', 'Hors ronde', 'Statut', 'Signalé le', 'Agent', 'Latitude', 'Longitude', 'Précision (m)', 'Photos', 'Dernier constat', 'Date dernier constat', 'Clos le', 'Clos par', 'Traitement', 'Actions', 'Suivi PM', 'Observations PM', 'Message aux agents', 'Échanges', 'Action demandée en attente', 'Notes DPMS']];
  sigs.forEach(s => {
    const last = (s.suivi || []).slice(-1)[0];
    A1.push([s.ref, siteNom(s.site), s.cat, s.sub || '', s.desc, s.plaque || '', s.emplacement || '', s.vehicule || '', (s.dests || []).join(', '), s.urgent ? 'Oui' : '', s.aggrave ? 'Oui' : '', s.horsRonde ? 'Oui' : '', STATUTS[s.statut] || s.statut, s.ts ? fmtDT(s.ts) : fmtJour(s.jour), s.agent || '',
      g(s.geo, 'lat'), g(s.geo, 'lon'), g(s.geo, 'acc'), (s.photos || []).length, last ? VERDICTS[last.verdict] : '', last ? (last.ts ? fmtDT(last.ts) : fmtJour(last.jour)) : '',
      s.closLe ? (s.closLe.length > 10 ? fmtDT(s.closLe) : fmtJour(s.closLe)) : '', s.closPar || '',
      s.statut === 'clos' ? 'Clos' : suiviEtat(s), (s.actions || []).map(a => `${a.quand ? fmtDT(a.quand) : fmtJour(a.jour)} ${ACTIONS[a.type]}${a.dest ? ' — ' + a.dest : ''}${a.texte ? ' : ' + a.texte : ''}`).join(' | '),
      isVehicule(s.cat) || s.pm ? pmEtat(s) : '', s.pm ? s.pm.hist.map(x => `${fmtJour(x.jour)} ${PM_STATUTS[x.statut]}${x.obs ? ' : ' + x.obs : ''}`).join(' | ') : '', s.notesAgents || '',
      (s.messages || []).map(m => `${fmtDT(m.quand)} ${m.de}${m.demande ? ' [demande d’action]' : ''}${m.fait ? ' [action faite]' : ''}${m.cloture ? ' [demande close]' : ''}${m.texte ? ' : ' + m.texte : ''}`).join(' | '), s.demande ? 'Oui' : '', s.notes || '']);
  });
  const A2 = [['Parking', 'Agent', 'Début', 'Fin', 'Durée (min)', 'Véhicules comptés', 'Non-conformités', 'Signalements', 'Constats', 'Lat. départ', 'Lon. départ', 'Précision (m)', 'Points d’attention', 'Téléphone']];
  rondes.forEach(r => {
    const dur = r.debut && r.fin && !r.closDPMS ? Math.round((new Date(r.fin) - new Date(r.debut)) / 60000) : '';
    const anos = [...(r.checklist ? Object.entries(r.checklist).filter(([, v]) => v !== 'RAS').map(([k, v]) => `${k} (${v})`) : []), ...(r.controles || []).flatMap(c => c.nc.map(x => `${x} (${NIVEAUX[c.niveau].lbl.toLowerCase()})`))].join(', ');
    A2.push([siteNom(r.site), r.agent, fmtDT(r.debut), fmtDT(r.fin), dur, r.compte != null ? r.compte : '', anos, r.nbSig, r.nbRevue, g(r.geoDebut, 'lat'), g(r.geoDebut, 'lon'), g(r.geoDebut, 'acc'), [...new Set(r.alertes || [])].join(' ; '), r.dev]);
  });
  const A3 = [['Parking', 'Barrière', 'Ouverte le', 'Par', 'Motif', 'Commentaire', 'Saisie de l’ouverture', 'Refermée le', 'Par', 'Commentaire', 'Saisie de la fermeture', 'Durée (min)', 'En cours']];
  barr.forEach(b => {
    const fin = b.fermee ? new Date(b.fermee.quand) : new Date();
    A3.push([siteNom(b.site), b.barriere, fmtDT(b.ouverte.quand), b.ouverte.agent || '', b.ouverte.motif || '', b.ouverte.comment || '', fmtDT(b.ouverte.ts), b.fermee ? fmtDT(b.fermee.quand) : '', b.fermee ? b.fermee.agent || '' : '', b.fermee ? b.fermee.comment || '' : '', b.fermee ? fmtDT(b.fermee.ts) : '', Math.round((fin - new Date(b.ouverte.quand)) / 60000), b.fermee ? '' : 'Oui']);
  });
  const TYPES = { controle: 'Contrôle', action: 'Traitement', ronde_debut: 'Début de ronde', ronde_fin: 'Fin de ronde', ronde_annulee: 'Ronde annulée', ronde_agent: 'Changement de rondier', msg: 'Message au DPMS', checklist: 'Contrôles', comptage: 'Comptage', signalement: 'Signalement', revue: 'Constat de suivi', barriere: 'Barrière', pm: 'Police municipale' };
  const A4 = [['Appareil', 'N°', 'Type', 'Jour déclaré', 'Horodatage appareil', 'Reçu par le relais', 'Parking', 'Agent', 'Détail', 'Latitude', 'Longitude', 'Précision (m)', 'Âge position (s)', 'Empreinte']];
  evs.forEach(e => {
    const ev = e.ev; let det = '';
    if (ev.t === 'signalement') det = `${ev.data.ref} ${ev.data.cat}${ev.data.sub ? ' / ' + ev.data.sub : ''}${ev.data.desc ? ' — ' + ev.data.desc : ''}${ev.data.horsRonde ? ' (hors ronde)' : ''}`;
    if (ev.t === 'revue') det = `${ev.data.ref} : ${VERDICTS[ev.data.verdict]}${ev.data.comment ? ' — ' + ev.data.comment : ''}`;
    if (ev.t === 'checklist') det = Object.entries(ev.data.items || {}).map(([k, v]) => `${k} : ${v}`).join(' ; ');
    if (ev.t === 'comptage') det = `${ev.data.total} véhicule(s)`;
    if (ev.t === 'controle') { const r = resumeControle(ev.data.items, ev.data.libelles); det = `${NIVEAUX[ev.data.niveau].lbl} : ${r.c} conforme(s)${r.nc.length ? ', non conforme : ' + r.nc.join(', ') : ''}`; }
    if (ev.t === 'action') det = `${ev.data.ref} : ${ACTIONS[ev.data.type]}${ev.data.dest ? ' — ' + ev.data.dest : ''}${ev.data.texte ? ' : ' + ev.data.texte : ''}`;
    if (ev.t === 'ronde_fin') det = `${ev.data.nbSig} signalement(s)`;
    if (ev.t === 'ronde_agent') det = `${ev.data.ancien} → ${ev.data.nouveau}`;
    if (ev.t === 'msg') det = `${ev.data.ref} : ${ev.data.fait ? '[action faite] ' : ''}${ev.data.texte || ''}`;
    if (ev.t === 'ronde_annulee') det = `annulée par l’agent${ev.data.nbSig ? ` (${ev.data.nbSig} signalement(s) conservé(s))` : ''}`;
    if (ev.t === 'barriere') det = `${ev.data.barriere} ${ev.data.action} — déclaré ${fmtDT(ev.data.quand)}${ev.data.motif ? ' — ' + ev.data.motif : ''}`;
    if (ev.t === 'pm') det = `${ev.data.ref} : ${PM_STATUTS[ev.data.statut]}${ev.data.obs ? ' — ' + ev.data.obs : ''}`;
    const geo = e.dec && e.dec.geo;
    A4.push([ev.dev, ev.seq, TYPES[ev.t] || ev.t, fmtJour(ev.jour), e.dec ? fmtDT(e.dec.ts) : 'illisible', fmtDT(e.at), siteNom(ev.site), ev.agent || '', det, g(geo, 'lat'), g(geo, 'lon'), g(geo, 'acc'), g(geo, 'age'), e.hashOk ? 'OK' : 'INVALIDE']);
  });
  hideBusy();
  await shareOrDownload(xlsxBlob([
    ['Signalements', A1, [11, 12, 22, 30, 45, 12, 16, 20, 24, 7, 7, 8, 8, 16, 12, 11, 11, 9, 7, 16, 16, 16, 14, 20, 60, 11, 40, 30, 60, 10, 40]],
    ['Rondes', A2, [12, 12, 16, 16, 10, 10, 40, 12, 10, 11, 11, 10, 40, 10]],
    ['Traitement', [['Réf.', 'Parking', 'Signalement', 'Date', 'Action', 'Auprès de', 'Détail', 'Intervention prévue le', 'Par'],
      ...sigs.flatMap(s => (s.actions || []).map(a => [s.ref, siteNom(s.site), sigTitre(s), a.quand ? fmtDT(a.quand) : fmtJour(a.jour), ACTIONS[a.type], a.dest || '', a.texte || '', a.prevu ? fmtDT(a.prevu) : '', a.par || '']))],
      [11, 12, 40, 16, 26, 16, 50, 18, 12]],
    ['Contrôles', [['Date', 'Parking', 'Contrôle', 'Agent', 'Conformes', 'Non conformes', 'N’existe pas ici', 'Non faits', 'Observations'],
      ...[...((await kvGet('controlesHist')) || []).map(x => { const r = resumeControle(x.items, x.libelles); return { at: x.ts || x.at, row: [fmtDT(x.ts || x.at), siteNom(x.site), NIVEAUX[x.niveau].lbl, x.agent, r.c, r.nc.join(', '), r.na.join(', '), r.nf.join(', '), x.partiel ? 'Partiel : ronde clôturée par le DPMS' : ''] }; }),
        ...((await kvGet('registres')) || []).map(x => ({ at: x.at, row: [fmtDT(x.at), siteNom(x.site), 'Registre de sécurité', x.par, '', '', '', '', x.note] }))]
        .sort((a, b) => a.at.localeCompare(b.at)).map(x => x.row)],
      [16, 12, 20, 12, 10, 50, 30, 30, 40]],
    ['Barrières', A3, [12, 12, 16, 12, 22, 25, 16, 16, 12, 25, 16, 10, 8]],
    ['Main courante', [['Réf.', 'Reçue le', 'Parking', 'Canal', 'Objet', 'Nom', 'Courriel', 'Téléphone', 'Demande', 'Suite donnée', 'Service', 'Enregistrée par', 'Saisie le', 'Statut', 'Première réponse', 'Délai de réponse (min)', 'Historique', 'Signalement lié'],
      ...(await mainCourante()).slice().reverse().map(c => [c.ref, fmtDT(c.quand), lieuMC(c), CANAUX[c.canal] || '', c.objet, c.nom, c.mail, c.tel, c.demande, SUITES[c.suite] || '', c.service || '', c.par, fmtDT(c.saisi), c.statut === 'clos' ? 'Réglée / close' : 'À traiter',
        c.reponduLe ? fmtDT(c.reponduLe) : '', delaiMC(c) == null ? '' : delaiMC(c), c.hist.map(x => `${fmtDT(x.quand)} ${x.type}${x.mode ? ' (' + (MODES_REP[x.mode] || x.mode) + ')' : ''} ${x.par}${x.texte ? ' : ' + x.texte : ''}`).join(' | '), c.sigs.length ? sigs.filter(x => c.sigs.includes(x.id)).map(x => x.ref).join(', ') : ''])],
      [11, 16, 12, 20, 18, 20, 24, 14, 50, 30, 18, 14, 16, 12, 16, 10, 60, 14]],
    ['Journal', A4, [10, 6, 16, 11, 16, 16, 12, 12, 55, 11, 11, 10, 10, 10]]
  ]), `rondes_parkings_complet_${fileStamp()}.xlsx`, 'Rondes parkings');
}
/* ---------- Synthèse hebdomadaire (PDF) ---------- */
function chargerScript(src, globale) {
  if (window[globale]) return Promise.resolve(window[globale]);
  return new Promise((ok, ko) => {
    const s = document.createElement('script'); s.src = src;
    s.onload = () => window[globale] ? ok(window[globale]) : ko(new Error('bibliothèque PDF illisible'));
    s.onerror = () => ko(new Error('bibliothèque PDF non chargée (réseau ?)'));
    document.head.appendChild(s);
  });
}
function periodeSemaine(decalage) {   // decalage 0 = semaine en cours, -1 = semaine précédente (lundi → dimanche)
  const d = new Date(); d.setHours(0, 0, 0, 0); d.setDate(d.getDate() - (jourIso(d) - 1) + 7 * decalage);
  const fin = new Date(d); fin.setDate(fin.getDate() + 7);
  const jours = []; for (let i = 0; i < 7; i++) { const x = new Date(d); x.setDate(x.getDate() + i); jours.push(x); }
  return { debut: d, fin, jours };
}
const JOUR_COURT = ['dim.', 'lun.', 'mar.', 'mer.', 'jeu.', 'ven.', 'sam.'];
const MOIS_LONGS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
const dateLongue = d => `${JOURS_LONGS[jourIso(d) - 1]} ${d.getDate()} ${MOIS_LONGS[d.getMonth()]} ${d.getFullYear()}`;
function libellePeriode(p) {
  const der = p.jours[6];
  return `du ${dateLongue(p.debut)} au ${dateLongue(der)}`;
}
const dureeTxt = min => min == null ? '—' : min < 60 ? `${min} min` : `${Math.floor(min / 60)} h ${String(min % 60).padStart(2, '0')}`;
async function donneesSynthese(p) {
  const dans = iso => { if (!iso) return false; const d = new Date(iso.length === 10 ? iso + 'T12:00:00' : iso); return d >= p.debut && d < p.fin; };
  const cfg = state.cfg, rm = cfg.rondeMatin || {};
  const maintenant = new Date(), finEff = p.fin < maintenant ? p.fin : maintenant;
  const rondes = (await dbAll('rondes')).filter(r => dans(r.debut));
  const sigs = await dbAll('sigs');
  const barr = ((await kvGet('barrieres')) || []);
  const hist = ((await kvGet('controlesHist')) || []).filter(x => dans(x.ts || x.at));
  const st = await etatControles();
  const regs = await etatRegistres();
  const joursAttendus = p.jours.filter(j => (rm.jours || [1, 2, 3, 4, 5, 6, 7]).includes(jourIso(j)) && j < finEff && !ferieFerme(localDate(j)));
  const parSite = cfg.sites.map(x => {
    const rs = rondes.filter(r => r.site === x.id);
    const durees = rs.filter(r => r.fin && !r.closDPMS).map(r => Math.round((new Date(r.fin) - new Date(r.debut)) / 60000));
    const comptes = rs.filter(r => r.compte != null).map(r => r.compte);
    const sans = rm.actif === false ? [] : joursAttendus.filter(j => !rs.some(r => localDate(new Date(r.debut)) === localDate(j)));
    const ctl = hist.filter(c => c.site === x.id);
    const nc = {};
    ctl.forEach(c => resumeControle(c.items, c.libelles).nc.forEach(l => { nc[l] = (nc[l] || 0) + 1; }));
    const b = barr.filter(y => y.site === x.id && (dans(y.ouverte.quand) || (new Date(y.ouverte.quand) < p.debut && (!y.fermee || new Date(y.fermee.quand) >= p.debut))));
    const dureeB = b.reduce((t, y) => {
      const a = Math.max(new Date(y.ouverte.quand), p.debut), z = Math.min(y.fermee ? new Date(y.fermee.quand) : maintenant, finEff);
      return t + Math.max(0, Math.round((z - a) / 60000));
    }, 0);
    return {
      site: x, rondes: rs.length, nonTerminees: rs.filter(r => !r.fin || r.closDPMS).length, sans,
      dureeMoy: durees.length ? Math.round(durees.reduce((a, b) => a + b, 0) / durees.length) : null,
      vehMoy: comptes.length ? Math.round(comptes.reduce((a, b) => a + b, 0) / comptes.length) : null,
      ctlQuot: ctl.filter(c => c.niveau === 'quotidien').length,
      nc, mensuel: echeance(st, x.id, 'mensuel'), trimestriel: echeance(st, x.id, 'trimestriel'), registre: regs.find(r => r.site === x.id),
      barrieres: b.filter(y => dans(y.ouverte.quand)).length, dureeB, barrOuvertes: barr.filter(y => y.site === x.id && !y.fermee)
    };
  });
  const cree = s => s.ts || s.jour;
  const nouveaux = sigs.filter(s => dans(cree(s)));
  const clos = sigs.filter(s => s.statut === 'clos' && dans(s.closLe));
  const ouverts = sigs.filter(s => s.statut !== 'clos');
  const age = s => joursDepuis(cree(s));
  const parCat = {};
  nouveaux.forEach(s => { const k = s.cat; parCat[k] = parCat[k] || { n: 0, u: 0 }; parCat[k].n++; if (s.urgent) parCat[k].u++; });
  const parEtat = {};
  ouverts.forEach(s => { const e = suiviEtat(s); parEtat[e] = (parEtat[e] || 0) + 1; });
  const mcTous = await mainCourante();
  const mcSem = mcTous.filter(c => dans(c.quand));
  const parObjet = {};
  mcSem.forEach(c => { const o = parObjet[c.objet] = parObjet[c.objet] || { n: 0, r: 0 }; o.n++; if (c.statut === 'clos') o.r++; });
  const delais = mcSem.filter(c => !SUITES_CLOSES.includes(c.suite) && c.reponduLe).map(delaiMC);
  const mc = { sem: mcSem, parObjet, parCanal: Object.keys(CANAUX).map(k => [CANAUX[k], mcSem.filter(c => c.canal === k).length]).filter(x => x[1]),
    surPlace: mcSem.filter(c => SUITES_CLOSES.includes(c.suite)).length, attente: mcTous.filter(c => c.statut === 'ouvert').sort((a, b) => a.quand.localeCompare(b.quand)),
    delaiMoy: delais.length ? Math.round(delais.reduce((a, b) => a + b, 0) / delais.length) : null, nbDelais: delais.length };
  return {
    mc, p, parSite, nouveaux, clos, ouverts, parCat, parEtat, joursAttendus,
    urgentsOuverts: ouverts.filter(s => s.urgent),
    anciens: ouverts.filter(s => age(s) > 30).sort((a, b) => cree(a).localeCompare(cree(b))),
    demandes: ouverts.filter(s => s.demande),
    age, edite: maintenant, releve: await kvGet('lastSync')
  };
}
async function pdfSynthese(D) {
  const { jsPDF } = await chargerScript('lib/jspdf.umd.min.js', 'jspdf');
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  const W = 210, M = 16, L = W - 2 * M, BAS = 282;
  const txt = t => String(t == null ? '' : t).replace(/[  ]/g, ' ').replace(/→/g, '->').replace(/[^\x00-\xff‘’“”–—•…€Œœ]/g, '?');
  let y = M;
  const saut = n => { if (y + n > BAS) { doc.addPage(); y = M; } };
  const couleur = (r, g, b) => doc.setTextColor(r, g, b);
  const titre = t => { saut(32); y += 3; doc.setFont('helvetica', 'bold'); doc.setFontSize(13); couleur(29, 79, 145); doc.text(txt(t), M, y); y += 2; doc.setDrawColor(29, 79, 145); doc.setLineWidth(0.5); doc.line(M, y, W - M, y); y += 6; couleur(20, 20, 20); };
  const para = (t, o = {}) => {
    doc.setFont('helvetica', o.gras ? 'bold' : 'normal'); doc.setFontSize(o.taille || 10);
    if (o.rouge) couleur(179, 38, 30); else if (o.gris) couleur(90, 96, 105); else couleur(20, 20, 20);
    const l = doc.splitTextToSize(txt(t), (o.largeur || L) - (o.retrait || 0));
    for (const x of l) { saut(5); doc.text(x, M + (o.retrait || 0), y); y += (o.taille || 10) * 0.45; }
    y += 1; couleur(20, 20, 20);
  };
  const tableau = (entetes, lignes, largeurs, o = {}) => {
    const tot = largeurs.reduce((a, b) => a + b, 0), ws = largeurs.map(w => w / tot * L);
    const ligne = (cells, gras, fond) => {
      doc.setFont('helvetica', gras ? 'bold' : 'normal'); doc.setFontSize(9);
      const parts = cells.map((c, i) => doc.splitTextToSize(txt(c), ws[i] - 2));
      const hgt = Math.max(...parts.map(x => x.length)) * 4 + 2;
      if (y + hgt > BAS) { doc.addPage(); y = M; if (!gras) { ligne(entetes, true, [238, 243, 250]); doc.setFont('helvetica', 'normal'); doc.setFontSize(9); } }
      if (fond) { doc.setFillColor(...fond); doc.rect(M, y - 3.6, L, hgt, 'F'); }
      let x = M;
      parts.forEach((pp, i) => { pp.forEach((s, k) => doc.text(s, x + 1, y + k * 4)); x += ws[i]; });
      y += hgt;
      doc.setDrawColor(201, 206, 214); doc.setLineWidth(0.2); doc.line(M, y - 3.6, W - M, y - 3.6);
    };
    saut(16);
    ligne(entetes, true, [238, 243, 250]);
    if (!lignes.length) ligne([o.vide || 'Aucun élément.', ...entetes.slice(1).map(() => '')], false);
    lignes.forEach(l => ligne(l, false));
    y += 3;
  };
  // en-tête
  try {
    const img = await new Promise((ok, ko) => { const i = new Image(); i.onload = () => ok(i); i.onerror = ko; i.src = 'logo-cachan.png'; });
    const cv = document.createElement('canvas'); cv.width = 360; cv.height = Math.round(360 * img.naturalHeight / img.naturalWidth);
    const cx = cv.getContext('2d'); cx.fillStyle = '#fff'; cx.fillRect(0, 0, cv.width, cv.height); cx.drawImage(img, 0, 0, cv.width, cv.height);
    doc.addImage(cv.toDataURL('image/jpeg', 0.88), 'JPEG', W - M - 24, M - 5, 24, 24 * cv.height / cv.width);
  } catch (e) { }
  doc.setFont('helvetica', 'bold'); doc.setFontSize(10); couleur(29, 79, 145); doc.text(txt('VILLE DE CACHAN — DPMS'), M, y); y += 8;
  doc.setFontSize(19); couleur(20, 20, 20); doc.text(txt('Parkings — synthèse hebdomadaire'), M, y); y += 7;
  para(`Semaine ${libellePeriode(D.p)}`, { taille: 12, gras: true, largeur: L - 28 });
  para(`Éditée le ${fmtDT(D.edite.toISOString())} — données relevées ${D.releve ? 'le ' + fmtDT(D.releve) : '(jamais relevées)'}. Parkings : ${state.cfg.sites.map(x => x.nom).join(', ')}.`, { gris: true, taille: 9, largeur: L - 28 });
  y += 2;
  // en bref
  titre('En bref');
  const totR = D.parSite.reduce((a, x) => a + x.rondes, 0), totSans = D.parSite.reduce((a, x) => a + x.sans.length, 0);
  const attendues = D.joursAttendus.length * D.parSite.length;
  const urgN = D.nouveaux.filter(s => s.urgent).length;
  const totB = D.parSite.reduce((a, x) => a + x.barrieres, 0), totBd = D.parSite.reduce((a, x) => a + x.dureeB, 0);
  const encoreOuv = D.parSite.flatMap(x => x.barrOuvertes);
  [`Rondes : ${totR} réalisée(s)${state.cfg.rondeMatin.actif !== false ? ` ; ${totSans} journée(s)-parking sans ronde sur ${attendues} attendue(s)` : ''}.`,
   `Signalements : ${D.nouveaux.length} nouveau(x)${urgN ? ` dont ${urgN} urgent(s)` : ''}, ${D.clos.length} clos dans la semaine ; ${D.ouverts.length} en cours à ce jour, dont ${D.parEtat['À transmettre'] || 0} à transmettre et ${D.anciens.length} ouvert(s) depuis plus de 30 jours.`,
   `Barrières : ${totB} ouverture(s), ${dureeTxt(totBd)} d’ouverture cumulée${encoreOuv.length ? ` ; ${encoreOuv.length} encore ouverte(s) à l’édition` : ''}.`,
   `Main courante : ${D.mc.sem.length} contact(s) ou demande(s), dont ${D.mc.surPlace} réglé(s) immédiatement ; ${D.mc.attente.length} en attente de réponse à ce jour${D.mc.delaiMoy != null ? ` ; délai moyen de réponse ${dureeTxt(D.mc.delaiMoy)}` : ''}.`,
   D.demandes.length ? `Demandes d’action aux agents sans réponse : ${D.demandes.length}.` : null]
    .filter(Boolean).forEach(t => para('•  ' + t, { retrait: 2 }));
  // rondes
  titre('Rondes');
  tableau(['Parking', 'Rondes', 'Durée moyenne', 'Véhicules (moy.)', 'Non terminées', 'Jours sans ronde'],
    D.parSite.map(x => [x.site.nom, x.rondes, dureeTxt(x.dureeMoy), x.vehMoy == null ? '—' : x.vehMoy, x.nonTerminees || '—', x.sans.length ? x.sans.map(j => `${JOUR_COURT[j.getDay()]} ${j.getDate()}`).join(', ') : 'aucun']),
    [22, 12, 18, 16, 15, 40]);
  // contrôles
  titre('Contrôles');
  tableau(['Parking', 'Contrôles quotidiens', 'Mensuel', 'Trimestriel', 'Registre de sécurité (DPMS)'],
    D.parSite.map(x => {
      const e = (v, n) => !v.dernier ? 'jamais fait' : `${fmtJour(localDate(new Date(v.dernier.at)))}${v.du ? ' — EN RETARD' : ''}`;
      return [x.site.nom, x.ctlQuot, e(x.mensuel), e(x.trimestriel), x.registre && x.registre.dernier ? `${fmtJour(localDate(new Date(x.registre.dernier.at)))}${x.registre.du ? ' — à faire' : ''}` : 'jamais contrôlé'];
    }), [22, 18, 24, 24, 30]);
  const ncl = D.parSite.flatMap(x => Object.entries(x.nc).map(([l, n]) => [x.site.nom, l, n])).sort((a, b) => b[2] - a[2]);
  para('Points relevés non conformes dans la semaine :', { gras: true });
  tableau(['Parking', 'Point de contrôle', 'Nombre de rondes'], ncl.slice(0, 15), [22, 60, 18], { vide: 'Aucune non-conformité.' });
  // signalements
  titre('Signalements');
  tableau(['Catégorie', 'Nouveaux', 'dont urgents'], Object.entries(D.parCat).sort((a, b) => b[1].n - a[1].n).map(([c, v]) => [c, v.n, v.u || '—']), [60, 20, 20], { vide: 'Aucun nouveau signalement.' });
  para('Signalements en cours, par état du traitement :', { gras: true });
  tableau(['État', 'Nombre'], Object.entries(D.parEtat).sort((a, b) => b[1] - a[1]), [60, 20], { vide: 'Aucun signalement en cours.' });
  const ligneSig = s => [s.ref, siteNom(s.site), sigTitre(s), `${D.age(s)} j`, suiviEtat(s)];
  if (D.urgentsOuverts.length) { para('Urgents encore ouverts :', { gras: true, rouge: true }); tableau(['Réf.', 'Parking', 'Signalement', 'Âge', 'Traitement'], D.urgentsOuverts.map(ligneSig), [17, 14, 41, 7, 21]); }
  if (D.anciens.length) { para('Ouverts depuis plus de 30 jours :', { gras: true }); tableau(['Réf.', 'Parking', 'Signalement', 'Âge', 'Traitement'], D.anciens.slice(0, 20).map(ligneSig), [17, 14, 41, 7, 21]); if (D.anciens.length > 20) para(`… et ${D.anciens.length - 20} autre(s) (voir l’Excel complet).`, { gris: true, taille: 9 }); }
  if (D.demandes.length) { para('Demandes d’action aux agents en attente :', { gras: true }); tableau(['Réf.', 'Parking', 'Signalement', 'Demandée le'], D.demandes.map(s => { const m = s.messages.find(x => x.id === s.demande); return [s.ref, siteNom(s.site), sigTitre(s), m ? fmtDT(m.quand) : '']; }), [17, 14, 47, 22]); }
  // main courante
  titre('Main courante (usagers et demandes reçues)');
  para(`${D.mc.sem.length} entrée(s) dans la semaine${D.mc.parCanal.length ? ' : ' + D.mc.parCanal.map(([l, n]) => `${l.toLowerCase()} ${n}`).join(', ') : ''}. ${D.mc.surPlace} réglée(s) immédiatement${D.mc.nbDelais ? ` ; ${D.mc.nbDelais} réponse(s) apportée(s) ensuite, délai moyen ${dureeTxt(D.mc.delaiMoy)}` : ''}.`);
  tableau(['Objet', 'Entrées', 'dont réglées'], Object.entries(D.mc.parObjet).sort((a, b) => b[1].n - a[1].n).map(([o, v]) => [o, v.n, v.r]), [60, 20, 20], { vide: 'Aucune entrée dans la semaine.' });
  if (D.mc.attente.length) {
    para('En attente de réponse :', { gras: true, rouge: D.mc.attente.some(c => joursDepuis(c.quand) >= 7) });
    tableau(['Réf.', 'Reçue le', 'Parking', 'Objet', 'Nom', 'Âge'], D.mc.attente.map(c => [c.ref, fmtDT(c.quand), lieuMC(c), c.objet, c.nom, `${joursDepuis(c.quand)} j`]), [14, 18, 14, 20, 24, 8]);
  }
  // barrières
  titre('Barrières');
  tableau(['Parking', 'Ouvertures', 'Durée cumulée', 'Encore ouvertes'],
    D.parSite.map(x => [x.site.nom, x.barrieres, dureeTxt(x.dureeB), x.barrOuvertes.length ? x.barrOuvertes.map(b => `${b.barriere} (depuis ${fmtDT(b.ouverte.quand)})`).join(', ') : '—']), [22, 14, 18, 46]);
  // pied de page
  const n = doc.getNumberOfPages();
  for (let i = 1; i <= n; i++) {
    doc.setPage(i); doc.setFont('helvetica', 'normal'); doc.setFontSize(8); couleur(90, 96, 105);
    doc.text(txt(`Ville de Cachan — DPMS — rondes parkings — semaine ${libellePeriode(D.p)}`), M, 291);
    doc.text(`${i} / ${n}`, W - M, 291, { align: 'right' });
  }
  return doc.output('blob');
}
VIEWS.synthese = async (params = {}) => {
  if (state.mode !== 'superviseur') return VIEWS.home();
  const dec = params.dec != null ? params.dec : (new Date().getDay() === 1 && new Date().getHours() < 12 ? -1 : (jourIso() === state.cfg.synthese.jour ? -1 : 0));
  const p = periodeSemaine(dec);
  const D = await donneesSynthese(p);
  const totR = D.parSite.reduce((a, x) => a + x.rondes, 0);
  const nomFichier = `synthese_parkings_${localDate(p.debut)}.pdf`;
  const lastSync = await kvGet('lastSync');
  return page('Synthèse hebdomadaire', { back: true },
    h('div', { class: 'card' },
      chips([{ val: -1, lbl: 'Semaine dernière' }, { val: 0, lbl: 'Semaine en cours' }], dec, v => go('synthese', { dec: v })),
      h('p', { style: 'margin:10px 0 4px;font-weight:700' }, `Semaine ${libellePeriode(p)}`),
      h('p', { class: 'muted small' }, `Données relevées ${lastSync ? fmtQuand(lastSync) : '— jamais'}.`),
      h('div', { class: 'stat' }, h('span', null, 'Rondes réalisées'), h('b', null, totR)),
      h('div', { class: 'stat' }, h('span', null, 'Journées-parking sans ronde'), h('b', null, D.parSite.reduce((a, x) => a + x.sans.length, 0))),
      h('div', { class: 'stat' }, h('span', null, 'Nouveaux signalements'), h('b', null, D.nouveaux.length)),
      h('div', { class: 'stat' }, h('span', null, 'Signalements clos'), h('b', null, D.clos.length)),
      h('div', { class: 'stat' }, h('span', null, 'En cours (dont + de 30 jours)'), h('b', null, `${D.ouverts.length} (${D.anciens.length})`)),
      h('div', { class: 'stat' }, h('span', null, 'Ouvertures de barrière'), h('b', null, D.parSite.reduce((a, x) => a + x.barrieres, 0))),
      h('div', { class: 'stat' }, h('span', null, 'Main courante (en attente de réponse)'), h('b', null, `${D.mc.sem.length} (${D.mc.attente.length})`))),
    h('button', { class: 'ok big', onclick: safe(async () => { showBusy('Préparation du PDF…'); try { const b = await pdfSynthese(await donneesSynthese(p)); hideBusy(); await shareOrDownload(b, nomFichier, 'Synthèse hebdomadaire des parkings'); } finally { hideBusy(); } }) }, 'Ouvrir / partager le PDF'),
    h('button', { class: 'sec', onclick: safe(async () => { const r = await syncNow(); toast(r && r.err ? 'Relève impossible : ' + r.err : 'Données à jour.', !!(r && r.err)); render(); }) }, 'Actualiser les données'),
    h('p', { class: 'muted small' }, 'Le PDF est fabriqué sur ce téléphone à partir des données relevées : le relais ne peut pas lire les données, il ne peut donc pas l’envoyer lui-même par courriel.'));
};
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

/* ---------- QR et liens de configuration ---------- */
function qrSvg(text, ecc = 'M') {
  const q = qrcode(0, ecc); q.addData(text, 'Byte'); q.make();
  return q.createSvgTag({ cellSize: 4, margin: 4, scalable: true });
}
VIEWS.supQR = async ({ first } = {}) => {
  const bloc = (titre, texte, role) => {
    const lien = appBaseURL() + '#join=' + packJoin(state.conn, role);
    const show = h('div', { class: 'qr' });
    return h('div', { class: 'card' }, h('h2', null, titre), h('p', { class: 'muted small' }, texte),
      h('div', { class: 'row' },
        h('button', { onclick: () => { show.innerHTML = qrSvg(lien, 'L'); } }, 'Afficher le QR'),
        h('button', { class: 'sec', onclick: safe(async () => { try { await navigator.clipboard.writeText(lien); toast('Lien copié : ouvrez-le dans le navigateur de l’appareil à configurer.'); } catch (e) { prompt('Copiez ce lien :', lien); } }) }, 'Copier le lien (PC)')),
      show,
      h('p', { class: 'muted small' }, 'Ce QR et ce lien donnent accès aux signalements de l’équipe : ne les diffusez pas.'));
  };
  return page('Configurer les appareils', { back: true },
    first ? h('div', { class: 'banner ok' }, 'Relais activé. Étapes suivantes : saisir les prénoms des agents et les coordonnées des destinataires (Listes, catégories et coordonnées), puis configurer les téléphones.') : null,
    bloc('Téléphones des agents et PC de la loge', 'Téléphone : scanner le QR avec l’appareil photo. PC sans caméra : copier le lien et l’ouvrir dans le navigateur du PC. Les modifications de listes parviennent ensuite seules.', 'agent'),
    bloc('Police municipale', 'Appareil du chef de service PM : signalements de véhicules (plaques, emplacements, photos), observations, traité / non traité.', 'pm'),
    blocSup());
};
function blocSup() {
  const lien = appBaseURL() + '#sup=' + b64u.enc(te.encode(state.conn.url));
  const show = h('div', { class: 'qr' });
  return h('div', { class: 'card' }, h('h2', null, 'Autre superviseur'),
    h('p', { class: 'muted small' }, 'Téléphone ou ordinateur d’un autre superviseur : scanner le QR (ou ouvrir le lien). L’adresse du relais est remplie d’office ; la phrase de passe superviseur reste demandée, elle n’est pas dans le QR.'),
    h('div', { class: 'row' },
      h('button', { onclick: () => { show.innerHTML = qrSvg(lien, 'M'); } }, 'Afficher le QR'),
      h('button', { class: 'sec', onclick: safe(async () => { try { await navigator.clipboard.writeText(lien); toast('Lien copié.'); } catch (e) { prompt('Copiez ce lien :', lien); } }) }, 'Copier le lien (PC)')),
    show);
}
async function publishCfg(cfg) {
  cfg = normalizeCfg(cfg);
  state.cfg = cfg; await kvSet('cfg', cfg);
  await queueDec('d', { type: 'cfg', ts: new Date().toISOString(), cfg });
  syncNow();
}

/* ---------- Listes, catégories et coordonnées ---------- */
VIEWS.supConfig = async ({ prefillDests } = {}) => {
  const c = JSON.parse(JSON.stringify(state.cfg));
  let taxoTexte = taxoToText(c.cats);
  if (prefillDests) {   // fusion : un destinataire existant (même nom, hors parenthèse) est mis à jour, les autres sont ajoutés
    const base = n => sansAccent(String(n).replace(/\s*\([^)]*\)\s*$/, ''));
    const renomLigne = (l, a, b) => { const i = l.indexOf('|'); if (i < 0) return l; const ds = l.slice(i + 1).split(',').map(x => x.trim()); return ds.includes(a) ? l.slice(0, i + 1) + ' ' + [...new Set(ds.map(x => x === a ? b : x))].join(', ') : l; };
    for (const [a, b] of Object.entries(prefillDests.renommer || {})) {
      const k = c.dests.findIndex(x => sansAccent(x) === sansAccent(a));
      if (k < 0 || c.dests.some(x => sansAccent(x) === sansAccent(b))) continue;
      const ancien = c.dests[k]; c.dests[k] = b; c.contacts[b] = c.contacts[ancien] || {}; delete c.contacts[ancien];
      taxoTexte = taxoTexte.split('\n').map(l => renomLigne(l, ancien, b)).join('\n');
    }
    for (const [nom, tel, mail] of prefillDests.dests) {
      const ex = c.dests.find(x => sansAccent(x) === sansAccent(nom)) || (prefillDests.remplacer && prefillDests.remplacer[nom] ? null : c.dests.find(x => base(x) === base(nom) && !/\(/.test(x)));
      if (ex) { const k = c.contacts[ex] || {}; delete c.contacts[ex]; c.dests[c.dests.indexOf(ex)] = nom; c.contacts[nom] = { tel: tel || k.tel || '', mail: mail || k.mail || '' }; }
      else { c.dests.push(nom); c.contacts[nom] = { tel: tel || '', mail: mail || '' }; }
    }
    for (const [ancien, nouveaux] of Object.entries(prefillDests.remplacer || {})) {   // dans les catégories : un nom remplacé par plusieurs (un par parking)
      taxoTexte = taxoTexte.split('\n').map(l => { const i = l.indexOf('|'); if (i < 0) return l; const ds = l.slice(i + 1).split(',').map(x => x.trim()); if (!ds.includes(ancien)) return l; return l.slice(0, i + 1) + ' ' + [...new Set(ds.flatMap(x => x === ancien ? nouveaux : [x]))].join(', '); }).join('\n');
      if (!prefillDests.dests.some(x => x[0] === ancien)) { const k = c.dests.indexOf(ancien); if (k >= 0 && !taxoTexte.includes(ancien)) { c.dests.splice(k, 1); delete c.contacts[ancien]; } }
    }
  }
  const ta = (lines, hmin = 140) => h('textarea', { value: lines.join('\n'), style: `min-height:${hmin}px` });
  const tAgents = ta(c.agents), tMotifs = ta(c.motifs), tObjets = ta(c.objetsContact);
  const tDests = ta(c.dests.map(d => { const k = c.contacts[d] || {}; return [d, k.tel || '', k.mail || ''].join(' | ').replace(/( \| )+$/, ''); }), 170);
  const tTaxo = h('textarea', { value: taxoTexte, style: 'min-height:420px;font-size:15px' });
  const tCtl = h('textarea', { value: controlesToText(c.controles), style: 'min-height:420px;font-size:15px' });
  const sitesInputs = c.sites.map(s => ({ nom: h('input', { type: 'text', value: s.nom }), barr: h('input', { type: 'text', value: s.barrieres.join(', ') }) }));
  const tel = h('input', { type: 'tel', value: c.urgenceTel || '', placeholder: '01 …' });
  const mail = h('input', { type: 'email', value: c.urgenceMail || '', placeholder: 'prenom.nom@ville-cachan.fr' });
  const hMatin = h('input', { type: 'time', value: c.rondeMatin.heure });
  const jSynth = h('select', null, ...JOURS_LONGS.map((j, i) => h('option', { value: String(i + 1), selected: c.synthese.jour === i + 1 ? true : null }, j)));
  const hSynth = h('input', { type: 'time', value: c.synthese.heure });
  const lines = t => t.value.split('\n').map(x => x.trim()).filter(Boolean);
  return page('Listes et catégories', { back: true },
    prefillDests ? h('div', { class: 'banner warn' }, 'Destinataires reçus par lien : vérifiez « Destinataires et coordonnées » et les catégories, puis « Enregistrer » en bas de page.') : null,
    h('div', { class: 'banner info' }, 'Les modifications sont transmises automatiquement à tous les appareils.'),
    h('div', { class: 'card' }, h('h2', null, 'Parkings et barrières'),
      ...sitesInputs.map(si => h('div', { style: 'margin-bottom:10px' }, si.nom, h('label', { class: 'f small' }, 'Barrières (séparées par des virgules)'), si.barr))),
    h('label', { class: 'f' }, 'Agents (un par ligne)'), tAgents,
    h('label', { class: 'f' }, 'Destinataires et coordonnées'),
    h('p', { class: 'muted small' }, 'Une ligne par destinataire : nom | téléphone | courriel. Les coordonnées servent aux boutons « Appeler » et « Courriel » du traitement. Un nom suivi d’un parking entre parenthèses, par exemple « Designa (Dumotel) », n’est proposé que pour les signalements de ce parking.'),
    tDests,
    h('label', { class: 'f' }, 'Points de contrôle'),
    h('p', { class: 'muted small' }, 'Trois rubriques : [Quotidien] (une fois par jour et par parking), [Mensuel] (tous les 30 jours), [Trimestriel] (tous les 90 jours). Une ligne par point : « - Libellé : ce qu’il faut regarder | catégorie du signalement en cas de non-conformité ». Un « ? » devant le libellé permet à l’agent d’indiquer que ce point n’existe pas dans le parking.'),
    tCtl,
    h('label', { class: 'f' }, 'Catégories et sous-catégories des signalements'),
    h('p', { class: 'muted small' }, 'Une catégorie par ligne, suivie si besoin de « | » et des destinataires proposés par défaut. Puis ses sous-catégories, une par ligne commençant par « - ». Un « ! » en fin de ligne marque une sous-catégorie urgente par défaut. Une sous-catégorie peut avoir ses propres destinataires : « - Tags ou graffitis | Ateliers, DPMS, Police municipale ». Chaque catégorie est une ligne de contrôle de la ronde. « Autre (préciser) » est ajouté automatiquement.'),
    tTaxo,
    h('label', { class: 'f' }, 'Motifs d’ouverture de barrière (un par ligne)'), tMotifs,
    h('label', { class: 'f' }, 'Main courante : objets des demandes (un par ligne)'), tObjets,
    h('label', { class: 'f' }, 'Téléphone d’urgence affiché aux agents'), tel,
    h('label', { class: 'f' }, 'Courriel d’urgence'), mail,
    h('label', { class: 'f' }, 'Géolocalisation GPS des saisies'),
    chips([{ val: false, lbl: 'Désactivée' }, { val: true, lbl: 'Activée' }], c.gps, v => c.gps = v),
    h('label', { class: 'f' }, 'Une ronde par jour et par parking : alerte si aucune ronde avant (fin de service)'),
    chips([{ val: true, lbl: 'Activée' }, { val: false, lbl: 'Désactivée' }], c.rondeMatin.actif !== false, v => c.rondeMatin.actif = v),
    hMatin, h('p', { class: 'muted small' }, 'Jours où une ronde est attendue :'),
    chips(JOURS.map((j, i) => ({ val: i + 1, lbl: j })), c.rondeMatin.jours, () => { }, { multi: true }),
    h('label', { class: 'f' }, 'Synthèse hebdomadaire (notification, puis PDF)'),
    chips([{ val: true, lbl: 'Activée' }, { val: false, lbl: 'Désactivée' }], c.synthese.actif !== false, v => c.synthese.actif = v),
    h('div', { class: 'row' }, jSynth, hSynth),
    h('button', {
      class: 'ok', onclick: safe(async () => {
        const cats = parseTaxo(tTaxo.value);
        const controles = parseControles(tCtl.value);
        if (!controles.some(x => x.niveau === 'quotidien')) throw new Error('Aucun point de contrôle quotidien.');
        if (!cats.length) throw new Error('Aucune catégorie.');
        const vides = cats.filter(x => !x.subs.length).map(x => x.nom);
        if (vides.length) throw new Error('Catégorie sans sous-catégorie : ' + vides.join(', '));
        const dl = lines(tDests).map(l => l.split('|').map(x => x.trim()));
        const dests = dl.map(x => x[0]).filter(Boolean);
        const contacts = Object.fromEntries(dl.filter(x => x[0]).map(x => [x[0], { tel: x[1] || '', mail: x[2] || '' }]));
        const inconnus = [...new Set(cats.flatMap(x => x.dests).filter(x => !dests.includes(x)))];
        if (inconnus.length) throw new Error('Destinataire inconnu dans les catégories : ' + inconnus.join(', '));
        if (!lines(tAgents).length || !dests.length) throw new Error('Listes vides.');
        sitesInputs.forEach((si, k) => {
          if (si.nom.value.trim()) c.sites[k].nom = si.nom.value.trim();
          const b = si.barr.value.split(',').map(x => x.trim()).filter(Boolean);
          if (b.length) c.sites[k].barrieres = b;
        });
        Object.assign(c, { agents: lines(tAgents), dests, contacts, cats, controles, motifs: lines(tMotifs), objetsContact: lines(tObjets), urgenceTel: tel.value.trim(), urgenceMail: mail.value.trim(), cfgId: randCode(8),
          rondeMatin: { ...c.rondeMatin, heure: hMatin.value || '20:00', jours: [...c.rondeMatin.jours].sort() }, synthese: { ...c.synthese, jour: parseInt(jSynth.value, 10), heure: hSynth.value || '08:00' } });
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
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).then(r => { _swReg = r; return r.update(); }).catch(e => console.warn('SW', e));
  state.mode = (await kvGet('mode')) || null;
  state.conn = (await kvGet('conn')) || null;
  state.cfg = normalizeCfg((await kvGet('cfg')) || null);
  state.priv = state.mode === 'superviseur' ? await kvGet('privKey') : null;
  if (state.mode && !state.conn) { state.mode = null; }
  if (state.mode === 'agent' && (await kvGet('ronde'))) Geo.start();
  const hash = location.hash;
  if (hash && /^#(join|site|cfg|sup|ptype|dests)=/.test(hash)) {
    history.replaceState(null, '', location.pathname + location.search);
    await handleLink(hash, false);
    if (state.view === 'home') render();
  } else if (hash === '#synthese' && state.mode === 'superviseur') {
    history.replaceState(null, '', location.pathname + location.search);
    go('synthese');
  } else if (hash === '#planning' && state.mode) {
    history.replaceState(null, '', location.pathname + location.search);
    go(state.mode === 'superviseur' ? 'planning' : 'planningAgents');
  } else render();
  if (state.mode) syncNow().then(refreshIfHome);
}
window.addEventListener('hashchange', () => {
  const hash = location.hash;
  if (hash && /^#(join|site|cfg|sup|ptype|dests)=/.test(hash)) { history.replaceState(null, '', location.pathname + location.search); handleLink(hash, false); }
  else if (hash === '#synthese' && state.mode === 'superviseur') { history.replaceState(null, '', location.pathname + location.search); go('synthese'); }
  else if (hash === '#planning' && state.mode) { history.replaceState(null, '', location.pathname + location.search); go(state.mode === 'superviseur' ? 'planning' : 'planningAgents'); }
});
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') setTimeout(confirmerContact, 600); });
setTimeout(confirmerContact, 1500);
if ('serviceWorker' in navigator) navigator.serviceWorker.addEventListener('message', e => {
  if (e.data === 'lien:#planning' && state.mode) { go(state.mode === 'superviseur' ? 'planning' : 'planningAgents'); syncNow().then(() => render()); return; }
  if (e.data === 'lien:#synthese' && state.mode === 'superviseur') { go('synthese'); syncNow().then(() => { if (state.view === 'synthese') render(); }); return; }
  if ((e.data === 'sync' || String(e.data).startsWith('lien:')) && state.mode) syncNow().then(refreshIfHome);
});
window.addEventListener('online', () => { if (state.mode) syncNow().then(refreshIfHome); });
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && state.mode) syncNow().then(refreshIfHome); });
setInterval(() => { if (document.visibilityState === 'visible' && state.mode) syncNow().then(refreshIfHome); }, SYNC_PERIODE_MS);
boot();
