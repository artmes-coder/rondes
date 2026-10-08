/* Rondes parkings — Ville de Cachan, DPMS — version 1.5
 * Application web hors ligne (PWA).
 * - L'outil appartient aux agents : rondes, comptage des véhicules, signalements (catégories et
 *   sous-catégories, plusieurs par catégorie, en ronde ou hors ronde), barrières laissées ouvertes.
 * - Police municipale : suivi des signalements de véhicules (observations, traité / non traité).
 * - Superviseurs : consultation, interventions ponctuelles, exports.
 * - Le relais (Google Apps Script) conserve un journal chiffré commun. Clé d'équipe : agents, PM,
 *   superviseurs. Clé superviseur : horodatage exact, positions, notes internes. Le relais ne lit rien.
 */
'use strict';

const APP_VERSION = '1.5.0';
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
        : j.err === 'relais non initialisé' ? 'Le relais a été remis à zéro. Cet appareil doit être reconfiguré : Réglages → Réinitialiser cet appareil, puis configuration (superviseur : « activer un relais neuf » ; téléphones : nouveau QR).'
          : j.err === 'accès refusé' ? 'Le relais a été réactivé avec de nouveaux codes. Cet appareil doit être reconfiguré : Réglages → Réinitialiser cet appareil, puis nouveau QR ou « Ajouter cet appareil comme superviseur ».'
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
const CONTROLES_DEFAUT = `[Quotidien]
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
    const [corps, cat] = t.split('|').map(x => x.trim());
    const i = corps.indexOf(' : ');
    const lbl = (i >= 0 ? corps.slice(0, i) : corps).trim(), detail = i >= 0 ? corps.slice(i + 3).trim() : '';
    if (lbl) out.push({ id: niv[0] + '-' + slug(lbl), niveau: niv, lbl, detail, cat: cat || CAT_AUTRE, opt });
  }
  return out;
}
function controlesToText(list) {
  return Object.keys(NIVEAUX).map(n => [`[${NIVEAUX[n].titre}]`, ...list.filter(c => c.niveau === n).map(c => `- ${c.opt ? '?' : ''}${c.lbl}${c.detail ? ' : ' + c.detail : ''}${c.cat ? ' | ' + c.cat : ''}`)].join('\n')).join('\n');
}
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
  n.gps = n.gps === true;
  if (!Array.isArray(n.controles) || !n.controles.length) n.controles = parseControles(CONTROLES_DEFAUT);
  n.contacts = n.contacts && typeof n.contacts === 'object' ? n.contacts : {};
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
  const root = document.getElementById('app');
  const fn = VIEWS[state.view] || VIEWS.home;
  const my = ++_renderSeq;
  document.body.classList.toggle('large', state.mode === 'superviseur');
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

  // Suppressions décidées par un superviseur (annulables depuis la corbeille)
  const suppr = new Map(), annul = new Set();
  for (const { e } of items) if (e.kind === 'dec') {
    if (e.dec.type === 'suppr') suppr.set(e.dec.id, e.dec);
    if (e.dec.type === 'annul_suppr') annul.add(e.dec.cible);
  }
  const del = { sigs: new Set(), rids: new Set() };
  for (const [id, d] of suppr) if (!annul.has(id)) { (d.sigs || []).forEach(x => del.sigs.add(x)); (d.rids || []).forEach(x => del.rids.add(x)); }
  const evSupprime = ev => (ev.t === 'signalement' && del.sigs.has(ev.data.id)) || (['revue', 'action', 'pm'].includes(ev.t) && del.sigs.has(ev.data.sig)) || (ev.rid && del.rids.has(ev.rid) && ev.t !== 'signalement' && ev.t !== 'revue');

  const sigs = new Map(), rondes = new Map(), barr = new Map(), devs = {}, chains = {}, ctl = {}, ctlHist = [];
  let cfg = null;
  const alertes = [];
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
          suivi: [], journalDPMS: [], notes: '', notesAgents: '', pm: null, actions: [], controle: ev.data.controle || null
        });
      } else if (ev.t === 'revue') {
        const s = sigs.get(ev.data.sig);
        if (s) {
          s.suivi.push({ jour: ev.jour, ts, agent: ev.agent, verdict: ev.data.verdict, comment: ev.data.comment, photos: (ev.photos || []).map(p => ({ id: p.id, sha: p.sha, n: e.n })), geo: dec && dec.geo });
          applyVerdict(s, ev.data.verdict, ts || ev.jour, ev.agent);
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
          if (ev.data.statut === 'traite') { s.statut = 'clos'; s.closLe = ts || ev.jour; s.closPar = 'Police municipale'; }
          else if (ev.data.statut === 'non_traite' && s.closPar === 'Police municipale') { s.statut = 'ouvert'; s.closLe = null; s.closPar = null; }
        }
      } else if (ev.t === 'controle') {
        const c = ctl[ev.site] = ctl[ev.site] || { na: new Set() };
        const quand = ev.data.quand || ev.jour + 'T12:00:00';
        if (!c[ev.data.niveau] || c[ev.data.niveau].at <= quand) c[ev.data.niveau] = { at: quand, agent: ev.agent };
        Object.entries(ev.data.items || {}).forEach(([k, v]) => { if (v === 'NA') c.na.add(k); });
        ctlHist.push({ site: ev.site, niveau: ev.data.niveau, at: quand, ts, agent: ev.agent, items: ev.data.items, libelles: ev.data.libelles, rid: ev.rid });
      } else if (ev.t === 'barriere') {
        const d = ev.data;
        const b = barr.get(d.bid) || { bid: d.bid, site: ev.site, barriere: d.barriere, ouverte: null, fermee: null };
        const rec = { quand: d.quand, agent: ev.agent, motif: d.motif || '', comment: d.comment || '', ts, dev: ev.dev };
        if (d.action === 'ouverte' && !b.ouverte) b.ouverte = rec;
        if (d.action === 'fermee' && !b.fermee) b.fermee = rec;
        if (b.ouverte) barr.set(d.bid, b);
      }
      if (sup && ev.rid && !del.rids.has(ev.rid)) {
        const r = rondes.get(ev.rid) || { rid: ev.rid, site: ev.site, agent: ev.agent, dev: ev.dev, jour: ev.jour, nbSig: 0, nbRevue: 0, compte: null, alertes: [] };
        if (ev.t === 'ronde_debut') { r.debut = ts; r.qr = ev.data.qr; r.geoDebut = dec && dec.geo; }
        if (ev.t === 'ronde_fin') { r.fin = ts; if (r.compte == null && ev.data.compte != null) r.compte = ev.data.compte; }
        if (ev.t === 'checklist') r.checklist = ev.data.items;
        if (ev.t === 'comptage') r.compte = ev.data.total;
        if (ev.t === 'controle') { r.controles = r.controles || []; r.controles.push({ niveau: ev.data.niveau, nc: resumeControle(ev.data.items, ev.data.libelles).nc }); }
        if (ev.t === 'signalement') r.nbSig++;
        if (ev.t === 'revue') r.nbRevue++;
        rondes.set(ev.rid, r);
      }
    } else if (e.kind === 'dec') {
      const d = e.dec;
      if (d.type === 'cfg') { if (validCfg(normalizeCfg(d.cfg))) cfg = normalizeCfg(d.cfg); }
      else if (d.type === 'action') {
        const s = sigs.get(d.action.sig);
        if (s) s.actions.push({ ...d.action, quand: d.action.quand || d.ts, jour: (d.ts || '').slice(0, 10), par: d.par || 'DPMS' });
      }
      else if (d.type === 'ctl_reactiver') { if (ctl[d.site]) ctl[d.site].na.delete(d.item); }
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
    for (const r of rondes.values()) if (r.debut && !r.fin && Date.now() - new Date(r.debut).getTime() > 6 * 3600000) r.alertes.push('fin de ronde non reçue');
    await dbReplaceAll('rondes', [...rondes.values()]);
    await kvSet('alertes', alertes.sort((a, b) => (b.at || '').localeCompare(a.at || '')));
    await kvSet('devs', devs);
  }
  await dbReplaceAll('sigs', [...sigs.values()]);
  await kvSet('barrieres', [...barr.values()]);
  await kvSet('supprimes', { sigs: [...del.sigs], rids: [...del.rids] });
  await kvSet('controles', Object.fromEntries(Object.entries(ctl).map(([k, v]) => [k, { ...v, na: [...v.na] }])));
  if (sup) await kvSet('controlesHist', ctlHist.sort((a, b) => b.at.localeCompare(a.at)));
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
  const enCours = sigs.filter(s => s.statut === 'ouvert');
  const perSite = cfg.sites.map(s => {
    const o = enCours.filter(x => x.site === s.id).length;
    return h('div', { class: 'stat' }, h('span', null, s.nom), h('span', null, h('b', null, o), h('span', { class: 'muted small' }, o > 1 ? ' en cours' : ' en cours')));
  });
  const parSuivi = SUIVIS.map(e => [e, enCours.filter(s => suiviEtat(s) === e).length]).filter(([, n]) => n);
  return page('Rondes parkings', null,
    ronde ? h('div', { class: 'card', style: 'border:2px solid var(--accent)' },
      h('h3', null, `Ronde en cours : ${siteNom(ronde.site)}`),
      h('p', { class: 'muted' }, `${ronde.agent} — commencée à ${fmtHeure(ronde.debut)}`),
      h('button', { class: 'accent big', onclick: () => resumeRonde(ronde) }, 'Reprendre la ronde'))
      : h('button', { class: 'big', onclick: () => go('choixParking') }, 'Commencer ma ronde'),
    await carteControlesAFaire(),
    h('div', { class: 'row' },
      h('button', { class: 'sec', onclick: () => go('signalement', { from: 'hr' }) }, 'Signalement hors ronde'),
      h('button', { class: 'sec', onclick: () => go('barriere', {}) }, 'Barrière ouverte')),
    ouvertes.length ? h('div', { class: 'card', style: 'border-color:var(--bad)' }, h('h2', null, 'Barrières ouvertes'), ...ouvertes.map(b => barriereLigne(b, true))) : null,
    h('div', { class: 'card' }, h('h2', null, 'Traitement des signalements'),
      parSuivi.length ? parSuivi.map(([e, n]) => h('button', { class: 'list-item', onclick: () => go('suivi', { f: e }) }, h('div', { class: 'stat' }, h('span', null, e), h('b', null, n))))
        : h('p', { class: 'muted' }, 'Aucun signalement en cours.'),
      h('button', { class: 'sec', onclick: () => go('suivi', { f: 'tous' }) }, 'Tous les signalements en cours')),
    h('div', { class: 'card' }, h('h2', null, 'Par parking'), ...perSite),
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

async function carteControlesAFaire() {
  const st = await etatControles();
  const lignes = state.cfg.sites.map(s => {
    const dus = Object.keys(NIVEAUX).filter(n => echeance(st, s.id, n).du);
    return dus.length ? h('div', { class: 'stat' }, h('span', null, s.nom), h('span', { class: 'small' }, ...dus.map(n => h('span', { class: 'badge ' + (n === 'quotidien' ? 'b-info' : 'b-warn') }, n)))) : null;
  }).filter(Boolean);
  return lignes.length ? h('div', { class: 'card' }, h('h2', null, 'Contrôles à faire'), ...lignes, h('p', { class: 'muted small' }, 'Ils sont proposés pendant la ronde du parking concerné.')) : null;
}
VIEWS.choixParking = async () => page('Commencer ma ronde', { back: true },
  h('p', null, 'Quel parking ?'),
  ...(state.cfg ? state.cfg.sites : []).map(s => h('button', { class: 'big sec', onclick: () => startRondeFlow(s.id) }, s.nom)));

async function startRondeFlow(siteId) {
  const cur = await kvGet('ronde');
  if (cur) {
    if (cur.site === siteId) { resumeRonde(cur); return; }
    if (!confirm(`Une ronde est en cours à ${siteNom(cur.site)}. La terminer et démarrer à ${siteNom(siteId)} ?`)) { go('home'); return; }
    await finishRonde(cur, true);
  }
  go('choixAgent', { site: siteId });
}
VIEWS.choixAgent = async ({ site }) => page(`Ronde : ${siteNom(site)}`, { back: true },
  h('p', null, 'Qui fait la ronde ?'),
  ...state.cfg.agents.map(a => h('button', { class: 'big sec', onclick: safe(() => beginRonde(site, a)) }, a)));

async function beginRonde(site, agent) {
  showBusy('Démarrage de la ronde…');
  Geo.start();
  await Geo.waitFirst(4000);
  const ronde = { rid: uuid(), site, agent, debut: new Date().toISOString(), step: 'revue', checklist: {}, sigsCrees: [], compte: 0, compteValide: false,
    quotidienDu: echeance(await etatControles(), site, 'quotidien').du, ctl: {}, ctlEmis: {} };
  await appendEvent('ronde_debut', ronde, {});
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

/* ---------- Contrôles : quotidien (une fois par jour), mensuel (30 j), trimestriel (90 j) ---------- */
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
function lignesControle(ronde, niveau, points, sigsRonde, rafraichir) {
  ronde.ctl = ronde.ctl || {};
  const rep = ronde.ctl[niveau] = ronde.ctl[niveau] || {};
  return points.map(c => {
    const nc = ncDe(sigsRonde, c.id);
    const v = nc.length ? 'NC' : rep[c.id];
    const etat = v === 'NC' ? h('span', { class: 'badge b-bad' }, nc.length > 1 ? `non conforme (${nc.length})` : 'non conforme')
      : v === 'C' ? h('span', { class: 'badge b-ok' }, 'conforme')
        : v === 'NA' ? h('span', { class: 'badge b-info' }, 'n’existe pas ici') : h('span', { class: 'badge b-warn' }, 'à contrôler');
    return h('div', { class: 'ctl' },
      h('div', { class: 'ctl-txt' }, h('b', null, c.lbl), ' ', etat, c.detail ? h('div', { class: 'muted small' }, c.detail) : null),
      h('div', { class: 'ctl-btns' },
        nc.length || v === 'NA' ? null : h('button', {
          class: v === 'C' ? 'sel-ras' : '', onclick: safe(async () => {
            if (v === 'C') delete rep[c.id]; else rep[c.id] = 'C';
            await kvSet('ronde', ronde); rafraichir();
          })
        }, 'Conforme'),
        v === 'NA' ? h('button', { class: 'sec', onclick: safe(async () => { delete rep[c.id]; await kvSet('ronde', ronde); rafraichir(); }) }, 'Annuler') :
          h('button', { class: nc.length ? 'sel-ano' : '', onclick: () => go('signalement', { from: 'ronde', cat: catPour(c), ctl: { id: c.id, niveau, lbl: c.lbl } }) }, nc.length ? '+ Non conforme' : 'Non conforme'),
        c.opt && !nc.length && v !== 'NA' ? h('button', {
          class: 'link mini', onclick: safe(async () => {
            if (!confirm(`« ${c.lbl} » n’existe pas dans ce parking ?\nCe point ne sera plus proposé ici.`)) return;
            rep[c.id] = 'NA'; await kvSet('ronde', ronde); rafraichir();
          })
        }, 'N’existe pas ici') : null));
  });
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

/* ---------- Ronde ---------- */
VIEWS.ronde = async () => {
  const ronde = await kvGet('ronde');
  if (!ronde) return agentHome();
  if (ronde.compte == null) ronde.compte = 0;
  ronde.ctlEmis = ronde.ctlEmis || {};
  const st = await etatControles();
  const sigsRonde = (await dbAll('sigs')).filter(s => ronde.sigsCrees.includes(s.id));
  const rafraichir = () => go('ronde', {}, { noPush: true, keepScroll: true });

  // Comptage des véhicules : à valider ; une fois validé, il se replie
  let compteur;
  if (ronde.compteValide) {
    compteur = h('div', { class: 'card', style: 'display:flex;align-items:center;justify-content:space-between;gap:8px' },
      h('span', null, h('b', null, `${ronde.compte} véhicule${ronde.compte > 1 ? 's' : ''} compté${ronde.compte > 1 ? 's' : ''}`), ' ', h('span', { class: 'badge b-ok' }, 'validé')),
      h('button', { class: 'link', style: 'width:auto;margin:0;min-height:40px', onclick: safe(async () => { ronde.compteValide = false; await kvSet('ronde', ronde); rafraichir(); }) }, 'modifier'));
  } else {
    const nb = h('div', { class: 'compteur' }, String(ronde.compte));
    const majCompte = async (v) => { ronde.compte = Math.max(0, v); nb.textContent = ronde.compte; await kvSet('ronde', ronde); };
    compteur = h('div', { class: 'card' },
      h('h2', null, 'Comptage des véhicules'),
      nb,
      h('button', { class: 'plus1', onclick: safe(async () => { await majCompte(ronde.compte + 1); if (navigator.vibrate) navigator.vibrate(30); }) }, '+1 véhicule'),
      h('div', { class: 'row' },
        h('button', { class: 'sec', onclick: safe(() => majCompte(ronde.compte - 1)) }, '−1'),
        h('button', { class: 'sec', onclick: safe(async () => { const v = prompt('Nombre de véhicules :', ronde.compte); if (v !== null && /^\d+$/.test(v.trim())) await majCompte(parseInt(v, 10)); }) }, 'Corriger')),
      h('button', {
        class: 'ok', onclick: safe(async () => {
          if (!confirm(`Valider le comptage : ${ronde.compte} véhicule${ronde.compte > 1 ? 's' : ''} ?`)) return;
          ronde.compteValide = true; await kvSet('ronde', ronde);
          await appendEvent('comptage', ronde, { total: ronde.compte });
          rafraichir();
        })
      }, 'Valider le comptage'));
  }

  // Contrôle quotidien (le premier de la journée dans ce parking)
  const ptsQ = pointsSite(st, ronde.site, 'quotidien');
  let quotidien;
  if (ronde.quotidienDu) {
    const reste = restantControle(ronde, 'quotidien', ptsQ, sigsRonde);
    quotidien = h('div', { class: 'card' }, h('h2', null, `Contrôle quotidien (${ptsQ.length - reste}/${ptsQ.length})`),
      ...lignesControle(ronde, 'quotidien', ptsQ, sigsRonde, rafraichir));
  } else {
    const e = echeance(st, ronde.site, 'quotidien');
    quotidien = h('div', { class: 'card muted small' }, `Contrôle quotidien déjà fait aujourd’hui${e.dernier ? ` par ${e.dernier.agent} à ${fmtHeure(e.dernier.at)}` : ''}.`);
  }
  // Contrôles mensuel et trimestriel : proposés quand ils arrivent à échéance
  const periodiques = ['mensuel', 'trimestriel'].map(niv => {
    if (ronde.ctlEmis[niv]) return h('div', { class: 'card small' }, h('span', { class: 'badge b-ok' }, 'fait'), ` ${NIVEAUX[niv].lbl} enregistré.`);
    const e = echeance(st, ronde.site, niv);
    if (!e.du) return null;
    const pts = pointsSite(st, ronde.site, niv);
    const fait = pts.length - restantControle(ronde, niv, pts, sigsRonde);
    return h('div', { class: 'card', style: 'border-color:var(--accent)' },
      h('h2', null, `${NIVEAUX[niv].lbl} à faire`),
      h('p', { class: 'muted small' }, e.dernier ? `Dernier : il y a ${e.depuis} jours (${e.dernier.agent}).` : 'Jamais fait dans ce parking.', fait ? ` ${fait}/${pts.length} points déjà contrôlés.` : ''),
      h('button', { class: 'accent', onclick: () => go('controle', { niveau: niv }) }, fait ? 'Reprendre' : 'Faire maintenant'));
  });

  const libres = sigsRonde.filter(s => !s.controle);
  return page(`Ronde : ${siteNom(ronde.site)}`, { back: () => go('home') },
    h('p', { class: 'muted' }, `${ronde.agent} — commencée à ${fmtHeure(ronde.debut)}`),
    compteur,
    quotidien,
    ...periodiques,
    h('div', { class: 'row' },
      h('button', { class: 'sec', onclick: () => go('signalement', { from: 'ronde' }) }, `+ Signalement${libres.length ? ` (${libres.length})` : ''}`),
      h('button', { class: 'sec', onclick: () => go('barriere', { site: ronde.site, from: 'ronde' }) }, 'Barrière ouverte')),
    h('button', {
      class: 'ok big', onclick: safe(async () => {
        if (!ronde.compteValide) { toast('Validez d’abord le comptage des véhicules.', true); return; }
        const reste = ronde.quotidienDu ? restantControle(ronde, 'quotidien', ptsQ, sigsRonde) : 0;
        if (reste) { toast(`Contrôle quotidien : encore ${reste} point(s) à contrôler.`, true); return; }
        if (!confirm('Terminer la ronde ?')) return;
        await finishRonde(ronde);
      })
    }, 'Terminer la ronde'));
};
VIEWS.controle = async ({ niveau }) => {
  const ronde = await kvGet('ronde');
  if (!ronde) return agentHome();
  const st = await etatControles();
  const pts = pointsSite(st, ronde.site, niveau);
  const sigsRonde = (await dbAll('sigs')).filter(s => ronde.sigsCrees.includes(s.id));
  const reste = restantControle(ronde, niveau, pts, sigsRonde);
  return page(`${NIVEAUX[niveau].lbl} : ${siteNom(ronde.site)}`, { back: () => go('ronde', {}, { noPush: true }) },
    h('div', { class: 'card' }, h('h2', null, `${pts.length - reste}/${pts.length} points contrôlés`),
      ...lignesControle(ronde, niveau, pts, sigsRonde, () => go('controle', { niveau }, { noPush: true, keepScroll: true }))),
    h('button', {
      class: 'ok big', onclick: safe(async () => {
        if (reste) { toast(`Encore ${reste} point(s) à contrôler.`, true); return; }
        await emettreControle(ronde, niveau, pts, sigsRonde);
        await fold();
        toast(`${NIVEAUX[niveau].lbl} enregistré.`);
        go('ronde', {}, { noPush: true });
      })
    }, `Valider le ${NIVEAUX[niveau].lbl.toLowerCase()}`),
    h('button', { class: 'link', onclick: () => go('ronde', {}, { noPush: true }) }, 'Plus tard (les réponses sont gardées)'));
};
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
VIEWS.signalement = async ({ cat, from, ctl }) => {
  const ronde = from === 'ronde' ? await kvGet('ronde') : null;
  if (from === 'ronde' && !ronde) return agentHome();
  const cfg = state.cfg;
  const cle = `${from}|${cat || ''}|${ctl ? ctl.id : ''}`;
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
  const retour = () => { sigDraft = null; if (ctl && ctl.niveau !== 'quotidien') go('controle', { niveau: ctl.niveau }, { noPush: true }); else if (from === 'ronde') go('ronde', {}, { noPush: true }); else { geoStopSiLibre(); go('home'); } };
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
        dests: [...d.dests], urgent: d.urgent, horsRonde: !ronde, thumb: d.photos[0] ? d.photos[0].thumb : null, controle: ctl || null
      },
      { photoAges: d.photos.map(p => ({ id: p.id, ageMin: p.ageMin })) }, d.photos);
    if (ronde) { const r = await kvGet('ronde'); r.sigsCrees.push(id); await kvSet('ronde', r); }
    else { await kvSet('dernierAgent', { nom: d.agent, at: Date.now() }); state.dernierSite = d.site; }
    await fold();
    const urgent = d.urgent, catNom = d.cat;
    sigDraft = null;
    hideBusy();
    toast(`Signalement ${ref} enregistré.`);
    if (encore) go('signalement', { from, cat: catNom, ctl }, { noPush: true });
    else retour();
    if (urgent || !ronde) syncNow().then(refreshIfHome);   // urgent ou hors ronde : envoi immédiat
  });

  return page(ctl ? 'Non conforme' : ronde ? 'Nouveau signalement' : 'Signalement hors ronde', { back },
    ctl ? h('div', { class: 'banner warn' }, `${NIVEAUX[ctl.niveau].lbl} — ${ctl.lbl}`) : null,
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

/* ====================================================================== */
/* TRAITEMENT DES SIGNALEMENTS (dispatch vers les services et prestataires) */
/* ====================================================================== */
const ACTIONS = {
  tel: 'Appel ou message téléphonique', mail: 'Courriel envoyé', autre: 'Transmis autrement', relance: 'Relance',
  reponse: 'Réponse reçue', prevue: 'Intervention prévue', faite: 'Intervention réalisée', note: 'Note'
};
const SUIVIS = ['À transmettre', 'En attente de réponse', 'Réponse reçue', 'Intervention prévue', 'Intervention réalisée'];
function suiviEtat(s) {
  const a = (s.actions || []).filter(x => x.type !== 'note');
  if (!a.length) return 'À transmettre';
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
async function photosDe(list) {
  const miss = new Set();
  for (const p of list || []) if (p.n && !(await dbGet('photos', p.id))) miss.add(p.n);
  const erreurs = [];
  for (const n of miss) { try { await fetchMedia(n); } catch (e) { erreurs.push(e.message); } }
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
    try { await navigator.share(data); return; } catch (e) { if (e.name === 'AbortError') return; }
  }
  try { await navigator.clipboard.writeText(txt); } catch (e) { }
  files.forEach(f => download(f, f.name));
  toast('Texte copié' + (files.length ? ' et photos téléchargées' : '') + ' : collez-les dans votre message.');
}
function journalActions(s) {
  const a = s.actions || [];
  if (!a.length) return [h('p', { class: 'muted' }, 'Aucune action enregistrée.')];
  return a.slice().reverse().map(x => h('div', { class: 'stat', style: 'display:block' },
    h('div', null, h('b', null, ACTIONS[x.type] || x.type), x.dest ? ' — ' + x.dest : '', x.type === 'prevue' && x.prevu ? ` — le ${fmtDT(x.prevu)}` : ''),
    x.texte ? h('div', { class: 'small' }, x.texte) : null,
    h('div', { class: 'muted small' }, `${x.quand ? fmtDT(x.quand) : fmtJour(x.jour)}${x.par ? ' — ' + x.par : ''}`)));
}
function contactsBloc(s, form) {
  const dests = s.dests || [];
  const sujet = `[Parking ${siteNom(s.site)}] ${sigTitre(s)} — ${s.ref}`;
  return h('div', { class: 'card' }, h('h2', null, 'Transmettre'),
    ...dests.map(dn => {
      const c = (state.cfg.contacts || {})[dn] || {};
      return h('div', { class: 'stat', style: 'gap:8px;flex-wrap:wrap' }, h('b', null, dn),
        h('span', { style: 'display:flex;gap:6px;flex-wrap:wrap' },
          c.tel ? h('a', { class: 'btn sec mini', href: 'tel:' + c.tel.replace(/\s/g, ''), onclick: () => form.preset('tel', dn) }, 'Appeler ' + c.tel) : null,
          c.mail ? h('a', { class: 'btn sec mini', href: `mailto:${c.mail}?subject=${encodeURIComponent(sujet)}&body=${encodeURIComponent(texteSignalement(s))}`, onclick: () => form.preset('mail', dn) }, 'Courriel') : null,
          !c.tel && !c.mail ? h('span', { class: 'muted small' }, 'coordonnées non renseignées') : null));
    }),
    h('button', { class: 'sec', onclick: safe(async () => { await partagerSig(s); form.preset('autre', dests[0] || null); }) }, 'Partager avec les photos (WhatsApp, courriel…)'),
    h('p', { class: 'muted small' }, 'Après l’envoi, complétez et enregistrez l’action ci-dessous.'));
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
function galerie(photos) {
  if (!photos.length) return null;
  const urls = photos.map(p => URL.createObjectURL(p.blob));
  return h('div', null,
    photos.some(p => p.attendu && p.sha !== p.attendu) ? h('div', { class: 'banner bad' }, 'Une photo a été modifiée après la saisie (empreinte différente).') : null,
    h('div', { class: 'gallery' }, ...urls.map((u, i) => h('img', { src: u, alt: 'Photo', onclick: () => lightbox(urls, i) }))));
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

/* ---------- Vues de traitement (équipe : téléphone agents ou PC de la loge) ---------- */
VIEWS.suivi = async ({ f = 'À transmettre' }) => {
  const enCours = (await dbAll('sigs')).filter(s => s.statut === 'ouvert');
  const list = enCours.filter(s => f === 'tous' || suiviEtat(s) === f).sort((a, b) => (b.urgent - a.urgent) || (a.jour || '').localeCompare(b.jour || ''));
  const cards = [];
  for (const s of list) {
    const th = await thumbOf(s);
    cards.push(h('button', { class: 'list-item', onclick: () => go('suiviSig', { id: s.id }) },
      h('div', { style: 'display:flex;gap:10px;align-items:flex-start' },
        th ? h('img', { src: th, alt: '', class: 'vignette' }) : null,
        h('div', { style: 'flex:1;min-width:0' },
          h('div', { class: 'l1' }, sigTitre(s), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null),
          h('div', { class: 'small' }, suiviBadge(s), ' ', (s.dests || []).join(', ')),
          h('div', { class: 'muted small' }, `${siteNom(s.site)} — ${s.ref} — ${depuisJours(s.jour)}${s.plaque ? ' — ' + s.plaque : ''}`)))));
  }
  return page('Traitement', { back: true },
    h('div', { class: 'chips', style: 'margin-bottom:12px' }, ...[...SUIVIS, 'tous'].map(k =>
      h('button', { class: 'chip' + (f === k ? ' on' : ''), onclick: () => go('suivi', { f: k }, { noPush: true }) }, `${k === 'tous' ? 'Tous' : k} (${enCours.filter(s => k === 'tous' || suiviEtat(s) === k).length})`))),
    cards.length ? cards : h('p', { class: 'muted' }, 'Aucun signalement.'));
};
VIEWS.suiviSig = async ({ id }) => {
  const s = await dbGet('sigs', id);
  if (!s) return page('Signalement', { back: true }, h('p', null, 'Introuvable.'));
  const recharger = () => go('suiviSig', { id }, { noPush: true });
  const form = actionForm(s, {
    agents: state.cfg.agents, parDefaut: await agentPrefere(),
    onSave: async (a, par) => {
      await appendEvent('action', { site: s.site, agent: par }, a);
      await kvSet('dernierAgent', { nom: par, at: Date.now() });
      await fold(); syncNow(); toast('Action enregistrée.'); recharger();
    },
    onClore: async (par, texte) => {
      await appendEvent('revue', { site: s.site, agent: par }, { sig: s.id, ref: s.ref, verdict: 'resolu', comment: texte || 'Clos depuis le traitement' });
      await fold(); syncNow(); toast('Signalement clos.'); go('suivi', { f: 'tous' }, { noPush: true });
    }
  });
  const { photos, erreurs } = await photosDe([...(s.photos || []), ...(s.suivi || []).flatMap(v => v.photos || [])]);
  return page(s.ref, { back: () => go('suivi', { f: 'tous' }, { noPush: true }) },
    h('div', { class: 'card' },
      h('h3', null, sigTitre(s), ' ', suiviBadge(s), s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null),
      h('p', { class: 'muted small' }, `${siteNom(s.site)} — signalé le ${fmtJour(s.jour)}${s.agent ? ' par ' + s.agent : ''}${s.horsRonde ? ' (hors ronde)' : ''} — à signaler à : ${(s.dests || []).join(', ')}`),
      s.desc ? h('p', null, s.desc) : null,
      s.plaque || s.emplacement ? h('p', null, s.plaque ? [h('span', { class: 'plaque' }, s.plaque), ' '] : null, s.emplacement || '') : null,
      galerie(photos),
      erreurs.length ? h('p', { class: 'muted small' }, 'Photos indisponibles : ' + erreurs[0]) : null,
      s.notesAgents ? h('div', { class: 'banner info' }, 'DPMS : ' + s.notesAgents) : null,
      pmBanner(s),
      (s.suivi || []).length ? h('div', null, h('p', { class: 'small', style: 'font-weight:600;margin:10px 0 2px' }, 'Constats en ronde'),
        ...s.suivi.map(v => h('p', { class: 'small', style: 'margin:2px 0' }, `${fmtJour(v.jour)} — ${v.agent || ''} : ${VERDICTS[v.verdict]}${v.comment ? ' — ' + v.comment : ''}`))) : null),
    contactsBloc(s, form),
    h('div', { class: 'card' }, h('h2', null, 'Actions entreprises'), ...journalActions(s)),
    form);
};

/* ---------- Journal lisible (agents, PM) ---------- */
async function journalEquipe() {
  const sp = (await kvGet('supprimes')) || { sigs: [], rids: [] };
  const ds = new Set(sp.sigs), dr = new Set(sp.rids);
  const garde = ev => !((ev.t === 'signalement' && ds.has(ev.data.id)) || (['revue', 'action', 'pm'].includes(ev.t) && ds.has(ev.data.sig)) || (ev.rid && dr.has(ev.rid) && !['signalement', 'revue'].includes(ev.t)));
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
    h('button', { class: 'big', onclick: () => go('supStats') }, 'Statistiques et graphiques'),
    h('div', { class: 'card' }, h('h2', null, 'Signalements'),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'ouvert' }) }, h('div', { class: 'stat' }, h('span', null, 'En cours', urg ? h('span', { class: 'badge b-bad' }, urg + ' urgent(s)') : null), h('b', null, ouverts.length))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'a_transmettre' }) }, h('div', { class: 'stat' }, h('span', null, 'À transmettre aux services'), h('b', null, ouverts.filter(s => suiviEtat(s) === 'À transmettre').length))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'clos' }) }, h('div', { class: 'stat' }, h('span', null, 'Clos'), h('b', null, sigs.length - ouverts.length))),
      h('button', { class: 'list-item', onclick: () => go('supSigs', { f: 'vehicules' }) }, h('div', { class: 'stat' }, h('span', null, 'Véhicules en attente de la PM'), h('b', null, vehPM)))),
    await carteControlesSup(),
    h('div', { class: 'card' }, h('h2', null, 'Dernières rondes'),
      rondes.length ? rondes.map(rondeLigne) : h('p', { class: 'muted' }, 'Aucune ronde reçue.'),
      h('div', { class: 'row' },
        rondes.length ? h('button', { class: 'sec', onclick: () => go('supRondes') }, 'Toutes les rondes') : null,
        h('button', { class: 'sec', onclick: () => go('supBarrieres') }, 'Barrières'))),
    h('div', { class: 'card' }, h('h2', null, 'Exports'),
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
    h('p', { class: 'muted small foot' }, `Version ${APP_VERSION}`));
}
async function carteControlesSup() {
  const st = await etatControles();
  const etat = (site, n) => {
    const e = echeance(st, site, n);
    if (!e.dernier) return h('span', { class: 'badge b-bad' }, `${n} : jamais`);
    if (n === 'quotidien') return h('span', { class: 'badge ' + (e.depuis <= 1 ? 'b-ok' : 'b-bad') }, `${n} : ${e.depuis === 0 ? 'aujourd’hui' : e.depuis === 1 ? 'hier' : `il y a ${e.depuis} j`}`);
    return h('span', { class: 'badge ' + (e.du ? 'b-bad' : 'b-ok') }, `${n} : ${e.du ? `en retard (${e.depuis} j)` : e.depuis === 0 ? 'aujourd’hui' : `il y a ${e.depuis} j`}`);
  };
  return h('div', { class: 'card' }, h('h2', null, 'Contrôles'),
    ...state.cfg.sites.map(s => h('div', { class: 'stat', style: 'flex-wrap:wrap;gap:6px' }, h('b', null, s.nom), h('span', { style: 'display:flex;gap:4px;flex-wrap:wrap' }, ...Object.keys(NIVEAUX).map(n => etat(s.id, n))))),
    h('button', { class: 'sec', onclick: () => go('supControles') }, 'Détail des contrôles'));
}
VIEWS.supControles = async () => {
  const st = await etatControles();
  const hist = (await kvGet('controlesHist')) || [];
  const tous = state.cfg.controles;
  return page('Contrôles', { back: true },
    ...state.cfg.sites.map(s => {
      const na = ((st[s.id] || {}).na || []).map(id => tous.find(c => c.id === id)).filter(Boolean);
      return h('div', { class: 'card' }, h('h2', null, s.nom),
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
          h('div', null, h('b', null, `${siteNom(x.site)} — ${NIVEAUX[x.niveau].lbl}`), ` — ${fmtDT(x.ts || x.at)} — ${x.agent}`),
          h('div', { class: 'small' }, `${r.c} conforme(s)`, r.nc.length ? h('span', { style: 'color:var(--bad)' }, ` · non conforme : ${r.nc.join(', ')}`) : null,
            r.na.length ? ` · n’existe pas : ${r.na.join(', ')}` : '', r.nf.length ? ` · non faits : ${r.nf.join(', ')}` : ''));
      }) : h('p', { class: 'muted' }, 'Aucun contrôle enregistré.')));
};
function rondeLigne(r) {
  const dur = r.debut && r.fin ? Math.round((new Date(r.fin) - new Date(r.debut)) / 60000) : null;
  const anos = r.checklist ? Object.entries(r.checklist).filter(([, v]) => v !== 'RAS').map(([k, v]) => v === 'Anomalie' ? k : `${k} (${v.toLowerCase()})`) : [];
  return h('div', { class: 'stat', style: 'display:block' },
    h('div', null, h('b', null, `${siteNom(r.site)} — ${r.agent}`), !r.fin ? h('span', { class: 'badge b-warn' }, 'fin non reçue') : null),
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
  const tests = { ouvert: s => s.statut === 'ouvert', a_transmettre: s => s.statut === 'ouvert' && suiviEtat(s) === 'À transmettre', clos: s => s.statut === 'clos', vehicules: s => isVehicule(s.cat), tous: () => true };
  const qq = (q || '').toLowerCase();
  const list = all.filter(s => tests[f](s) && (!site || s.site === site) && (!cat || s.cat === cat)
    && (!qq || [s.ref, s.cat, s.sub, s.desc, s.plaque, s.emplacement, s.agent, s.vehicule].join(' ').toLowerCase().includes(qq)))
    .sort((a, b) => (b.urgent - a.urgent) || (b.ts || b.jour || '').localeCompare(a.ts || a.jour || ''));
  const nav = p => go('supSigs', { f, site, cat, q, sel, ...p }, { noPush: true, keepScroll: true });
  const large = ecranLarge();
  const filtres = h('div', { class: 'card' },
    h('div', { class: 'chips' }, ...[['ouvert', 'En cours'], ['a_transmettre', 'À transmettre'], ['clos', 'Clos'], ['vehicules', 'Véhicules'], ['tous', 'Tous']].map(([k, l]) =>
      h('button', { class: 'chip' + (f === k ? ' on' : ''), onclick: () => nav({ f: k, sel: null }) }, `${l} (${all.filter(tests[k]).length})`))),
    h('div', { class: 'filtres' },
      h('select', { onchange: e => nav({ site: e.target.value, sel: null }) }, h('option', { value: '' }, 'Tous les parkings'), ...state.cfg.sites.map(s => h('option', { value: s.id, selected: s.id === site ? true : null }, s.nom))),
      h('select', { onchange: e => nav({ cat: e.target.value, sel: null }) }, h('option', { value: '' }, 'Toutes les catégories'), ...[...state.cfg.cats.map(c => c.nom), CAT_AUTRE].map(c => h('option', { value: c, selected: c === cat ? true : null }, c))),
      h('input', { type: 'search', value: q, placeholder: 'Rechercher (plaque, réf., texte…)', onchange: e => nav({ q: e.target.value, sel: null }) })),
    list.length ? h('button', { class: 'link', style: 'color:var(--bad)', onclick: safe(async () => {
      if (prompt(`Supprimer les ${list.length} signalement(s) affiché(s) par ces filtres ?\nIls disparaîtront de tous les appareils, des statistiques et des exports (annulable depuis la corbeille).\nTapez SUPPRIMER pour confirmer.`) !== 'SUPPRIMER') return;
      await supprimer({ sigs: list.map(x => x.id), quoi: `${list.length} signalement(s) (filtre : ${{ ouvert: 'en cours', a_transmettre: 'à transmettre', clos: 'clos', vehicules: 'véhicules', tous: 'tous' }[f]}${site ? ', ' + siteNom(site) : ''}${cat ? ', ' + cat : ''}${q ? ', « ' + q + ' »' : ''})` });
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
          h('div', { class: 'small' }, statutBadge(s.statut), s.statut === 'ouvert' ? suiviBadge(s) : null, s.urgent ? h('span', { class: 'badge b-bad' }, 'Urgent') : null, isVehicule(s.cat) ? pmBadge(s) : null),
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
  const { photos, erreurs } = await photosDe(list);
  return [galerie(photos), erreurs.length ? h('p', { class: 'muted small' }, 'Photos indisponibles : ' + erreurs[0]) : null];
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
      h('div', null, h('b', null, VERDICTS[v.verdict] || v.verdict), ` — ${v.ts ? fmtDT(v.ts) : fmtJour(v.jour)} — ${v.agent || ''}`),
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
      ...(await photoBlocks(s.photos))),
    h('div', { class: 'card' }, h('h2', null, 'Traitement'), ...journalActions(s)),
    s.pm && s.pm.hist.length ? h('div', { class: 'card' }, h('h2', null, 'Police municipale'),
      ...s.pm.hist.map(x => h('p', { class: 'small' }, h('b', null, `${x.ts ? fmtDT(x.ts) : fmtJour(x.jour)} — ${PM_STATUTS[x.statut]}`), x.obs ? ' : ' + x.obs : ''))) : null,
    suivi.length ? h('div', { class: 'card' }, h('h2', null, 'Constats des agents'), ...suivi) : null,
    s.statut === 'ouvert' ? contactsBloc(s, form) : null,
    form,
    h('div', { class: 'card' }, h('h2', null, 'Intervenir'),
      h('p', { class: 'muted small' }, 'Toute intervention est transmise automatiquement aux agents, à la PM et aux autres superviseurs.'),
      s.statut === 'clos' ? h('button', { class: 'sec', onclick: safe(async () => { await supDecision(s, 'réouverture', { statut: 'ouvert', closLe: null, closPar: null }); done('Rouvert.'); }) }, 'Rouvrir') : null,
      h('button', { class: 'sec', onclick: () => { state.retour = { view: state.view, params: { ...state.params } }; go('supSig', { id: s.id, edit: true }); } }, 'Modifier ou écrire aux agents'),
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
    const msg = h('input', { type: 'text', value: s.notesAgents || '', placeholder: 'Ex. : intervention Ateliers prévue mardi' });
    const retour = () => { const r = state.retour || { view: 'supSig', params: { id } }; state.retour = null; go(r.view, r.params, { noPush: true }); };
    return page(`Modifier ${s.ref}`, { back: retour },
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
  const evs = (await dbAll('entries')).filter(e => e.kind === 'ev').filter(({ ev }) => !((ev.t === 'signalement' && ds.has(ev.data.id)) || (['revue', 'action', 'pm'].includes(ev.t) && ds.has(ev.data.sig)) || (ev.rid && dr.has(ev.rid) && !['signalement', 'revue'].includes(ev.t)))).sort((a, b) => a.k.localeCompare(b.k));
  const g = (geo, k) => geo && !geo.err && geo[k] != null ? geo[k] : '';
  const A1 = [['Réf.', 'Parking', 'Catégorie', 'Sous-catégorie', 'Description', 'Plaque', 'Emplacement', 'Véhicule', 'Destinataires', 'Urgent', 'Aggravé', 'Hors ronde', 'Statut', 'Signalé le', 'Agent', 'Latitude', 'Longitude', 'Précision (m)', 'Photos', 'Dernier constat', 'Date dernier constat', 'Clos le', 'Clos par', 'Traitement', 'Actions', 'Suivi PM', 'Observations PM', 'Message aux agents', 'Notes DPMS']];
  sigs.forEach(s => {
    const last = (s.suivi || []).slice(-1)[0];
    A1.push([s.ref, siteNom(s.site), s.cat, s.sub || '', s.desc, s.plaque || '', s.emplacement || '', s.vehicule || '', (s.dests || []).join(', '), s.urgent ? 'Oui' : '', s.aggrave ? 'Oui' : '', s.horsRonde ? 'Oui' : '', STATUTS[s.statut] || s.statut, s.ts ? fmtDT(s.ts) : fmtJour(s.jour), s.agent || '',
      g(s.geo, 'lat'), g(s.geo, 'lon'), g(s.geo, 'acc'), (s.photos || []).length, last ? VERDICTS[last.verdict] : '', last ? (last.ts ? fmtDT(last.ts) : fmtJour(last.jour)) : '',
      s.closLe ? (s.closLe.length > 10 ? fmtDT(s.closLe) : fmtJour(s.closLe)) : '', s.closPar || '',
      s.statut === 'clos' ? 'Clos' : suiviEtat(s), (s.actions || []).map(a => `${a.quand ? fmtDT(a.quand) : fmtJour(a.jour)} ${ACTIONS[a.type]}${a.dest ? ' — ' + a.dest : ''}${a.texte ? ' : ' + a.texte : ''}`).join(' | '),
      isVehicule(s.cat) || s.pm ? pmEtat(s) : '', s.pm ? s.pm.hist.map(x => `${fmtJour(x.jour)} ${PM_STATUTS[x.statut]}${x.obs ? ' : ' + x.obs : ''}`).join(' | ') : '', s.notesAgents || '', s.notes || '']);
  });
  const A2 = [['Parking', 'Agent', 'Début', 'Fin', 'Durée (min)', 'Véhicules comptés', 'Non-conformités', 'Signalements', 'Constats', 'Lat. départ', 'Lon. départ', 'Précision (m)', 'Points d’attention', 'Téléphone']];
  rondes.forEach(r => {
    const dur = r.debut && r.fin ? Math.round((new Date(r.fin) - new Date(r.debut)) / 60000) : '';
    const anos = [...(r.checklist ? Object.entries(r.checklist).filter(([, v]) => v !== 'RAS').map(([k, v]) => `${k} (${v})`) : []), ...(r.controles || []).flatMap(c => c.nc.map(x => `${x} (${NIVEAUX[c.niveau].lbl.toLowerCase()})`))].join(', ');
    A2.push([siteNom(r.site), r.agent, fmtDT(r.debut), fmtDT(r.fin), dur, r.compte != null ? r.compte : '', anos, r.nbSig, r.nbRevue, g(r.geoDebut, 'lat'), g(r.geoDebut, 'lon'), g(r.geoDebut, 'acc'), [...new Set(r.alertes || [])].join(' ; '), r.dev]);
  });
  const A3 = [['Parking', 'Barrière', 'Ouverte le', 'Par', 'Motif', 'Commentaire', 'Saisie de l’ouverture', 'Refermée le', 'Par', 'Commentaire', 'Saisie de la fermeture', 'Durée (min)', 'En cours']];
  barr.forEach(b => {
    const fin = b.fermee ? new Date(b.fermee.quand) : new Date();
    A3.push([siteNom(b.site), b.barriere, fmtDT(b.ouverte.quand), b.ouverte.agent || '', b.ouverte.motif || '', b.ouverte.comment || '', fmtDT(b.ouverte.ts), b.fermee ? fmtDT(b.fermee.quand) : '', b.fermee ? b.fermee.agent || '' : '', b.fermee ? b.fermee.comment || '' : '', b.fermee ? fmtDT(b.fermee.ts) : '', Math.round((fin - new Date(b.ouverte.quand)) / 60000), b.fermee ? '' : 'Oui']);
  });
  const TYPES = { controle: 'Contrôle', action: 'Traitement', ronde_debut: 'Début de ronde', ronde_fin: 'Fin de ronde', checklist: 'Contrôles', comptage: 'Comptage', signalement: 'Signalement', revue: 'Constat de suivi', barriere: 'Barrière', pm: 'Police municipale' };
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
    if (ev.t === 'barriere') det = `${ev.data.barriere} ${ev.data.action} — déclaré ${fmtDT(ev.data.quand)}${ev.data.motif ? ' — ' + ev.data.motif : ''}`;
    if (ev.t === 'pm') det = `${ev.data.ref} : ${PM_STATUTS[ev.data.statut]}${ev.data.obs ? ' — ' + ev.data.obs : ''}`;
    const geo = e.dec && e.dec.geo;
    A4.push([ev.dev, ev.seq, TYPES[ev.t] || ev.t, fmtJour(ev.jour), e.dec ? fmtDT(e.dec.ts) : 'illisible', fmtDT(e.at), siteNom(ev.site), ev.agent || '', det, g(geo, 'lat'), g(geo, 'lon'), g(geo, 'acc'), g(geo, 'age'), e.hashOk ? 'OK' : 'INVALIDE']);
  });
  hideBusy();
  await shareOrDownload(xlsxBlob([
    ['Signalements', A1, [11, 12, 22, 30, 45, 12, 16, 20, 24, 7, 7, 8, 8, 16, 12, 11, 11, 9, 7, 16, 16, 16, 14, 20, 60, 11, 40, 30, 40]],
    ['Rondes', A2, [12, 12, 16, 16, 10, 10, 40, 12, 10, 11, 11, 10, 40, 10]],
    ['Traitement', [['Réf.', 'Parking', 'Signalement', 'Date', 'Action', 'Auprès de', 'Détail', 'Intervention prévue le', 'Par'],
      ...sigs.flatMap(s => (s.actions || []).map(a => [s.ref, siteNom(s.site), sigTitre(s), a.quand ? fmtDT(a.quand) : fmtJour(a.jour), ACTIONS[a.type], a.dest || '', a.texte || '', a.prevu ? fmtDT(a.prevu) : '', a.par || '']))],
      [11, 12, 40, 16, 26, 16, 50, 18, 12]],
    ['Contrôles', [['Date', 'Parking', 'Contrôle', 'Agent', 'Conformes', 'Non conformes', 'N’existe pas ici', 'Non faits'],
      ...((await kvGet('controlesHist')) || []).slice().reverse().map(x => { const r = resumeControle(x.items, x.libelles); return [fmtDT(x.ts || x.at), siteNom(x.site), NIVEAUX[x.niveau].lbl, x.agent, r.c, r.nc.join(', '), r.na.join(', '), r.nf.join(', ')]; })],
      [16, 12, 20, 12, 10, 50, 30, 30]],
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
    bloc('Police municipale', 'Appareil du chef de service PM : signalements de véhicules (plaques, emplacements, photos), observations, traité / non traité.', 'pm'));
};
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
  const tAgents = ta(c.agents), tMotifs = ta(c.motifs);
  const tDests = ta(c.dests.map(d => { const k = c.contacts[d] || {}; return [d, k.tel || '', k.mail || ''].join(' | ').replace(/( \| )+$/, ''); }), 170);
  const tTaxo = h('textarea', { value: taxoToText(c.cats), style: 'min-height:420px;font-size:15px' });
  const tCtl = h('textarea', { value: controlesToText(c.controles), style: 'min-height:420px;font-size:15px' });
  const sitesInputs = c.sites.map(s => ({ nom: h('input', { type: 'text', value: s.nom }), barr: h('input', { type: 'text', value: s.barrieres.join(', ') }) }));
  const tel = h('input', { type: 'tel', value: c.urgenceTel || '', placeholder: '01 …' });
  const mail = h('input', { type: 'email', value: c.urgenceMail || '', placeholder: 'prenom.nom@ville-cachan.fr' });
  const lines = t => t.value.split('\n').map(x => x.trim()).filter(Boolean);
  return page('Listes et catégories', { back: true },
    h('div', { class: 'banner info' }, 'Les modifications sont transmises automatiquement à tous les appareils.'),
    h('div', { class: 'card' }, h('h2', null, 'Parkings et barrières'),
      ...sitesInputs.map(si => h('div', { style: 'margin-bottom:10px' }, si.nom, h('label', { class: 'f small' }, 'Barrières (séparées par des virgules)'), si.barr))),
    h('label', { class: 'f' }, 'Agents (un par ligne)'), tAgents,
    h('label', { class: 'f' }, 'Destinataires et coordonnées'),
    h('p', { class: 'muted small' }, 'Une ligne par destinataire : nom | téléphone | courriel. Les coordonnées servent aux boutons « Appeler » et « Courriel » du traitement.'),
    tDests,
    h('label', { class: 'f' }, 'Points de contrôle'),
    h('p', { class: 'muted small' }, 'Trois rubriques : [Quotidien] (une fois par jour et par parking), [Mensuel] (tous les 30 jours), [Trimestriel] (tous les 90 jours). Une ligne par point : « - Libellé : ce qu’il faut regarder | catégorie du signalement en cas de non-conformité ». Un « ? » devant le libellé permet à l’agent d’indiquer que ce point n’existe pas dans le parking.'),
    tCtl,
    h('label', { class: 'f' }, 'Catégories et sous-catégories des signalements'),
    h('p', { class: 'muted small' }, 'Une catégorie par ligne, suivie si besoin de « | » et des destinataires proposés par défaut. Puis ses sous-catégories, une par ligne commençant par « - ». Un « ! » en fin de ligne marque une sous-catégorie urgente par défaut. Chaque catégorie est une ligne de contrôle de la ronde. « Autre (préciser) » est ajouté automatiquement.'),
    tTaxo,
    h('label', { class: 'f' }, 'Motifs d’ouverture de barrière (un par ligne)'), tMotifs,
    h('label', { class: 'f' }, 'Téléphone d’urgence affiché aux agents'), tel,
    h('label', { class: 'f' }, 'Courriel d’urgence'), mail,
    h('label', { class: 'f' }, 'Géolocalisation GPS des saisies'),
    chips([{ val: false, lbl: 'Désactivée' }, { val: true, lbl: 'Activée' }], c.gps, v => c.gps = v),
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
        Object.assign(c, { agents: lines(tAgents), dests, contacts, cats, controles, motifs: lines(tMotifs), urgenceTel: tel.value.trim(), urgenceMail: mail.value.trim(), cfgId: randCode(8) });
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
