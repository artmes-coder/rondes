/**
 * Relais chiffré — Rondes parkings, Ville de Cachan (DPMS) — version 6
 *
 * Journal partagé entre les téléphones des agents et les appareils superviseurs.
 * Tout ce qui est stocké ici est chiffré sur les appareils avant l'envoi :
 * ce script et le compte Google qui l'héberge ne peuvent rien lire.
 *
 * Contenu du dossier Drive « Relais rondes parkings (chiffré) » :
 *   p_<n>_<tél>.bin  saisies d'une ronde (chiffrées pour l'équipe)
 *   m_<n>_<tél>.bin  photos de ces saisies (chiffrées pour l'équipe)
 *   d_<n>.bin        interventions et configuration du superviseur (chiffrées pour l'équipe)
 *   n_<n>.bin        notes internes du superviseur (chiffrées pour les seuls superviseurs)
 *   coffre.bin       clés du superviseur, chiffrées par sa phrase de passe
 *   index.json       liste des fichiers (numéros, types, tailles) — aucun contenu
 *   abonnements.json adresses de notification des appareils et leurs clés publiques ;
 *                    le texte des notifications est chiffré par l'appareil qui l'envoie,
 *                    pour l'appareil qui le reçoit : ni ce script ni Google ne peuvent le lire
 *
 * REPARTIR DE ZÉRO : dans la barre d'outils de l'éditeur, choisir la fonction
 * « reinitialiserRelais » puis cliquer sur « Exécuter ». Tout le contenu du relais
 * est mis à la corbeille et les codes d'accès sont effacés.
 *
 * NOTIFICATIONS (version 4) : après la mise à jour, choisir la fonction
 * « autoriserNotifications » puis « Exécuter », et accepter l'autorisation
 * « se connecter à un service externe ».
 *
 * ALERTES PROGRAMMÉES (version 6 : ronde du matin manquante, ronde de plus d'une heure,
 * synthèse du lundi) : après la mise à jour, choisir la fonction « installerAlertes »
 * puis « Exécuter », et accepter l'autorisation demandée. Le relais vérifie alors
 * toutes les 5 minutes s'il doit envoyer une alerte. Il ne connaît que des jetons
 * anonymes (« une ronde a commencé aujourd'hui pour le jeton X ») et des messages
 * déjà chiffrés pour chaque appareil : il ne sait ni de quel parking il s'agit,
 * ni ce que disent les messages.
 */
var DOSSIER = 'Relais rondes parkings (chiffré)';
var TAILLE_MAX = 30000000;      // caractères par dépôt
var PULL_OCTETS = 8000000;      // volume maximal renvoyé par relève
var PULL_NOMBRE = 150;          // nombre maximal d'éléments renvoyés par relève

function doGet() {
  return json_({ ok: true, service: 'relais-rondes', v: 6 });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, err: 'requête illisible' }); }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (err) { return json_({ ok: false, err: 'relais occupé, réessayer' }); }
  try {
    return json_(traiter_(req));
  } catch (err) {
    return json_({ ok: false, err: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

/** À exécuter une fois depuis l'éditeur après la mise à jour en version 4 (autorisation d'envoi des notifications). */
function autoriserNotifications() {
  UrlFetchApp.fetch('https://fcm.googleapis.com/', { muteHttpExceptions: true });
  Logger.log('Notifications autorisées.');
}

/** À exécuter une fois depuis l'éditeur (version 6) : vérification des alertes toutes les 5 minutes. */
function installerAlertes() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'verifierAlertes') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('verifierAlertes').timeBased().everyMinutes(5).create();
  Logger.log('Alertes programmées : vérification toutes les 5 minutes.');
}

/** À exécuter depuis l'éditeur pour repartir de zéro. */
function reinitialiserRelais() {
  var props = PropertiesService.getScriptProperties();
  var fid = props.getProperty('FOLDER_ID');
  if (fid) { try { DriveApp.getFolderById(fid).setTrashed(true); } catch (e) { } }
  var it = DriveApp.getFoldersByName(DOSSIER);
  while (it.hasNext()) it.next().setTrashed(true);
  props.deleteAllProperties();
  Logger.log('Relais réinitialisé : il peut être activé par un nouveau superviseur.');
}

function traiter_(req) {
  var props = PropertiesService.getScriptProperties();
  var supTok = props.getProperty('SUP_TOKEN');
  var agTok = props.getProperty('AGENT_TOKEN');

  if (req.op === 'init') {
    if (supTok) return { ok: false, err: 'deja-initialise' };
    if (!jetonValide_(req.supTok) || !jetonValide_(req.agentTok) || req.supTok === req.agentTok) return { ok: false, err: 'codes d’accès invalides' };
    props.setProperties({ SUP_TOKEN: req.supTok, AGENT_TOKEN: req.agentTok, SEQ: '0' });
    dossier_();
    ecrireIndex_([]);
    return { ok: true };
  }
  if (!supTok) return { ok: false, err: 'relais non initialisé' };

  if (req.op === 'getVault') { // protégé par la phrase de passe, déchiffré uniquement sur l'appareil
    var v = premier_(dossier_(), 'coffre.bin');
    return v ? { ok: true, data: v.getBlob().getDataAsString() } : { ok: false, err: 'aucun superviseur enregistré' };
  }

  var estSup = req.tok === supTok, estAgent = req.tok === agTok;
  if (!estSup && !estAgent) return { ok: false, err: 'accès refusé' };
  var d = dossier_();

  switch (req.op) {
    case 'ping':
      return { ok: true, role: estSup ? 'superviseur' : 'agents', seq: parseInt(props.getProperty('SEQ'), 10) || 0 };

    case 'putVault':
      if (!estSup || !chaine_(req.data)) return { ok: false, err: 'interdit' };
      remplacer_(d, 'coffre.bin', req.data);
      return { ok: true };

    case 'put': { // téléphone agents : saisies (+ photos)
      if (!estAgent) return { ok: false, err: 'interdit' };
      var dev = String(req.dev || '').replace(/[^A-Za-z0-9-]/g, '').slice(0, 24);
      if (!dev || !chaine_(req.data) || (req.media != null && !chaine_(req.media))) return { ok: false, err: 'dépôt invalide' };
      var n = suivant_(props), at = new Date().toISOString(), idx = lireIndex_();
      var fp = d.createFile('p_' + pad_(n) + '_' + dev + '.bin', req.data, MimeType.PLAIN_TEXT);
      var ent = { n: n, k: 'p', dev: dev, id: fp.getId(), size: req.data.length, at: at };
      if (req.media) {
        var fm = d.createFile('m_' + pad_(n) + '_' + dev + '.bin', req.media, MimeType.PLAIN_TEXT);
        ent.mid = fm.getId(); ent.msize = req.media.length;
      }
      idx.push(ent); ecrireIndex_(idx);
      return { ok: true, n: n, at: at };
    }

    case 'putDec': { // superviseur : intervention, configuration (d) ou note interne (n)
      if (!estSup) return { ok: false, err: 'interdit' };
      if ((req.kind !== 'd' && req.kind !== 'n') || !chaine_(req.data)) return { ok: false, err: 'dépôt invalide' };
      var n2 = suivant_(props), at2 = new Date().toISOString(), idx2 = lireIndex_();
      var f2 = d.createFile(req.kind + '_' + pad_(n2) + '.bin', req.data, MimeType.PLAIN_TEXT);
      idx2.push({ n: n2, k: req.kind, id: f2.getId(), size: req.data.length, at: at2 });
      ecrireIndex_(idx2);
      return { ok: true, n: n2, at: at2 };
    }

    case 'pull': { // relève incrémentale
      var permis = estSup ? { p: 1, d: 1, n: 1 } : { p: 1, d: 1 };
      var kinds = Array.isArray(req.kinds) ? req.kinds : ['p', 'd'];
      var after = parseInt(req.after, 10) || 0;
      var liste = lireIndex_().filter(function (x) { return x.n > after && permis[x.k] && kinds.indexOf(x.k) >= 0; });
      liste.sort(function (a, b) { return a.n - b.n; });
      var out = [], total = 0, last = after, more = false;
      for (var i = 0; i < liste.length; i++) {
        var x = liste[i];
        if (out.length && (out.length >= PULL_NOMBRE || total + x.size > PULL_OCTETS)) { more = true; break; }
        var data;
        try { data = DriveApp.getFileById(x.id).getBlob().getDataAsString(); } catch (err) { data = null; }
        out.push({ n: x.n, k: x.k, dev: x.dev || null, at: x.at, media: !!x.mid, data: data });
        total += x.size; last = x.n;
      }
      return { ok: true, entries: out, last: last, more: more };
    }

    case 'abonner': { // adresse de notification d'un appareil (agents ou superviseur)
      var ep = String(req.endpoint || '');
      if (!pushAutorise_(ep) || (req.role !== 'agents' && req.role !== 'sup')) return { ok: false, err: 'abonnement invalide' };
      if (req.role === 'sup' && !estSup) return { ok: false, err: 'interdit' };
      var abos = lireAbos_().filter(function (a) { return a.e !== ep; });
      var cle = function (x) { return typeof x === 'string' && /^[A-Za-z0-9_-]{16,120}$/.test(x) ? x : null; };
      var prefs = Array.isArray(req.prefs) ? req.prefs.filter(function (x) { return typeof x === 'string' && /^[a-z_]{2,30}$/.test(x); }).slice(0, 30) : null;
      abos.push({ e: ep, r: req.role, p256dh: cle(req.p256dh), auth: cle(req.auth), prefs: prefs, at: new Date().toISOString() });
      ecrireAbos_(abos.slice(-40));
      return { ok: true };
    }

    case 'abos': { // clés publiques des appareils à notifier (pour chiffrer le texte sur l'appareil émetteur)
      if (req.cible !== 'agents' && req.cible !== 'sup') return { ok: false, err: 'cible invalide' };
      return { ok: true, abos: lireAbos_().filter(function (a) { return a.r === req.cible && a.p256dh && a.auth; })
        .map(function (a) { return { e: a.e, p256dh: a.p256dh, auth: a.auth, prefs: a.prefs || null }; }) };
    }

    case 'notifier': { // « nouveau message » vers les appareils agents ou superviseurs (sans contenu)
      if ((req.cible !== 'agents' && req.cible !== 'sup') || !chaine_(req.k) || !req.jwts || typeof req.jwts !== 'object') return { ok: false, err: 'notification invalide' };
      var garder = [], envoyes = 0, liste2 = lireAbos_();
      for (var j = 0; j < liste2.length; j++) {
        var a = liste2[j];
        var jwt = a.r === req.cible ? req.jwts[origine_(a.e)] : null;
        if (!jwt || typeof jwt !== 'string' || (req.seulsChiffres && a.p256dh && a.auth && !(req.corps && req.corps[a.e]))) { garder.push(a); continue; }
        var code = 0, entetes = { Authorization: 'vapid t=' + jwt + ', k=' + req.k, TTL: '86400', Urgency: 'high' };
        var corps = req.corps && typeof req.corps[a.e] === 'string' && req.corps[a.e].length < 8000 ? req.corps[a.e] : null;
        if (corps) {   // texte chiffré pour cet appareil (RFC 8291)
          try {
            var chiffre = { Authorization: entetes.Authorization, TTL: entetes.TTL, Urgency: entetes.Urgency, 'Content-Encoding': 'aes128gcm' };
            code = UrlFetchApp.fetch(a.e, { method: 'post', contentType: 'application/octet-stream', payload: Utilities.base64Decode(corps), muteHttpExceptions: true, headers: chiffre }).getResponseCode();
          } catch (err) { code = 0; }
        }
        if (!corps || (code >= 400 && code !== 404 && code !== 410) || code === 0) {   // sans texte : « nouveau message »
          try {
            code = UrlFetchApp.fetch(a.e, { method: 'post', payload: '', muteHttpExceptions: true, headers: entetes }).getResponseCode();
          } catch (err) { code = 0; }
        }
        if (code === 404 || code === 410) continue;   // appareil désabonné : adresse oubliée
        if (code >= 200 && code < 300) envoyes++;
        garder.push(a);
      }
      ecrireAbos_(garder);
      return { ok: true, envoyes: envoyes };
    }

    case 'pointer': { // jetons anonymes « une ronde a commencé aujourd'hui »
      if (!Array.isArray(req.jetons)) return { ok: false, err: 'pointage invalide' };
      var pt = lireJson_('pointages.json', {}), jour = jourParis_();
      req.jetons.slice(0, 20).forEach(function (j) { if (jetonValide_(j)) pt[j] = jour; });
      ecrireJson_('pointages.json', pt);
      return { ok: true };
    }

    case 'programmer': { // superviseur : alertes du jour et de la semaine, messages déjà chiffrés par appareil
      if (!estSup) return { ok: false, err: 'interdit' };
      var al = lireJson_('alertes.json', {});
      if (req.vapid && req.vapid.d && req.vapid.pub) al.vapid = { d: String(req.vapid.d), pub: String(req.vapid.pub), sub: String(req.vapid.sub || 'mailto:dpms@localhost') };
      if (Array.isArray(req.quotidiennes)) al.quotidiennes = req.quotidiennes.slice(0, 10);
      if (Array.isArray(req.hebdo)) al.hebdo = req.hebdo.slice(0, 10);
      al.envoyees = al.envoyees || {};
      ecrireJson_('alertes.json', al);
      return { ok: true };
    }

    case 'differer': { // alerte à envoyer plus tard sauf annulation (ronde de plus d'une heure)
      if (!jetonValide_(req.id) || !req.messages || typeof req.messages !== 'object' || !req.apres) return { ok: false, err: 'alerte invalide' };
      var al2 = lireJson_('alertes.json', {});
      al2.differees = (al2.differees || []).filter(function (x) { return x.id !== req.id; });
      al2.differees.push({ id: req.id, apres: String(req.apres), messages: req.messages });
      al2.differees = al2.differees.slice(-30);
      ecrireJson_('alertes.json', al2);
      return { ok: true };
    }

    case 'annulerDiffere': {
      var al3 = lireJson_('alertes.json', {});
      al3.differees = (al3.differees || []).filter(function (x) { return x.id !== req.id; });
      ecrireJson_('alertes.json', al3);
      return { ok: true };
    }

    case 'media': { // photos d'un dépôt (équipe et superviseurs ; elles restent chiffrées)
      var m = lireIndex_().filter(function (x) { return x.n === parseInt(req.n, 10) && x.mid; })[0];
      if (!m) return { ok: false, err: 'photos introuvables' };
      return { ok: true, data: DriveApp.getFileById(m.mid).getBlob().getDataAsString() };
    }
  }
  return { ok: false, err: 'opération inconnue' };
}

function dossier_() {
  var props = PropertiesService.getScriptProperties();
  var fid = props.getProperty('FOLDER_ID');
  if (fid) { try { var f = DriveApp.getFolderById(fid); if (!f.isTrashed()) return f; } catch (e) { } }
  var nf = DriveApp.createFolder(DOSSIER);
  props.setProperty('FOLDER_ID', nf.getId());
  return nf;
}
function lireIndex_() {
  var f = premier_(dossier_(), 'index.json');
  return f ? JSON.parse(f.getBlob().getDataAsString() || '[]') : [];
}
function ecrireIndex_(idx) {
  var d = dossier_(), f = premier_(d, 'index.json');
  if (f) f.setContent(JSON.stringify(idx)); else d.createFile('index.json', JSON.stringify(idx), MimeType.PLAIN_TEXT);
}
function suivant_(props) {
  var n = (parseInt(props.getProperty('SEQ'), 10) || 0) + 1;
  props.setProperty('SEQ', String(n));
  return n;
}
function premier_(d, nom) {
  var it = d.getFilesByName(nom);
  return it.hasNext() ? it.next() : null;
}
function remplacer_(d, nom, data) {
  var it = d.getFilesByName(nom);
  while (it.hasNext()) it.next().setTrashed(true);
  d.createFile(nom, data, MimeType.PLAIN_TEXT);
}
/* ---------- Alertes programmées (déclencheur toutes les 5 minutes) ---------- */
function verifierAlertes() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) { return; }
  try {
    var al = lireJson_('alertes.json', {});
    if (!al.vapid) return;
    var pt = lireJson_('pointages.json', {}), jour = jourParis_(), heure = Utilities.formatDate(new Date(), 'Europe/Paris', 'HH:mm');
    var jourSem = parseInt(Utilities.formatDate(new Date(), 'Europe/Paris', 'u'), 10);   // 1 = lundi
    al.envoyees = al.envoyees || {};
    var change = false;
    (al.quotidiennes || []).forEach(function (q) {
      if (al.envoyees[q.id] === jour || heure < q.heure || (q.jours && q.jours.indexOf(jourSem) < 0)) return;
      var manquants = [];
      (q.jetons || []).forEach(function (j, i) { if (pt[j] !== jour) manquants.push(i); });
      if (manquants.length && q.messages && q.messages[manquants.join(',')]) envoyerChiffres_(al.vapid, q.messages[manquants.join(',')]);
      al.envoyees[q.id] = jour; change = true;
    });
    (al.hebdo || []).forEach(function (w) {
      if (al.envoyees[w.id] === jour || jourSem !== w.jour || heure < w.heure) return;
      envoyerChiffres_(al.vapid, w.messages || {});
      al.envoyees[w.id] = jour; change = true;
    });
    var maintenant = new Date().toISOString(), restantes = [];
    (al.differees || []).forEach(function (d) {
      if (d.apres <= maintenant) { envoyerChiffres_(al.vapid, d.messages || {}); change = true; }
      else restantes.push(d);
    });
    al.differees = restantes;
    if (change) ecrireJson_('alertes.json', al);
  } finally { lock.releaseLock(); }
}
function envoyerChiffres_(vapid, messages) {
  var jwts = {};
  Object.keys(messages).forEach(function (ep) {
    var aud = origine_(ep);
    if (!pushAutorise_(ep) || typeof messages[ep] !== 'string') return;
    try {
      if (!jwts[aud]) jwts[aud] = jwtVapid_(aud, vapid);
      var code = UrlFetchApp.fetch(ep, { method: 'post', contentType: 'application/octet-stream', payload: Utilities.base64Decode(messages[ep]), muteHttpExceptions: true,
        headers: { Authorization: 'vapid t=' + jwts[aud] + ', k=' + vapid.pub, TTL: '86400', Urgency: 'high', 'Content-Encoding': 'aes128gcm' } }).getResponseCode();
      if (code === 404 || code === 410) ecrireAbos_(lireAbos_().filter(function (a) { return a.e !== ep; }));
    } catch (e) { }
  });
}
function jourParis_() { return Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyy-MM-dd'); }
function lireJson_(nom, defaut) {
  var f = premier_(dossier_(), nom);
  try { return f ? JSON.parse(f.getBlob().getDataAsString() || 'null') || defaut : defaut; } catch (e) { return defaut; }
}
function ecrireJson_(nom, o) { remplacer_(dossier_(), nom, JSON.stringify(o)); }

/* ---------- Signature ES256 (P-256) des jetons VAPID, en JavaScript pur ---------- */
var EC_P = BigInt('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff');
var EC_N = BigInt('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551');
var EC_A = EC_P - BigInt(3);
var EC_G = [BigInt('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'), BigInt('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')];
function mod_(a, m) { var r = a % m; return r < BigInt(0) ? r + m : r; }
function puiss_(b, e, m) { var r = BigInt(1); b = mod_(b, m); while (e > BigInt(0)) { if (e & BigInt(1)) r = r * b % m; b = b * b % m; e >>= BigInt(1); } return r; }
function inv_(a, m) { return puiss_(mod_(a, m), m - BigInt(2), m); }
function ajoute_(P, Q) {
  if (!P) return Q; if (!Q) return P;
  var l;
  if (P[0] === Q[0]) {
    if (mod_(P[1] + Q[1], EC_P) === BigInt(0)) return null;
    l = mod_((BigInt(3) * P[0] * P[0] + EC_A) * inv_(BigInt(2) * P[1], EC_P), EC_P);
  } else l = mod_((Q[1] - P[1]) * inv_(Q[0] - P[0], EC_P), EC_P);
  var x = mod_(l * l - P[0] - Q[0], EC_P);
  return [x, mod_(l * (P[0] - x) - P[1], EC_P)];
}
function multiplie_(k, P) { var R = null; while (k > BigInt(0)) { if (k & BigInt(1)) R = ajoute_(R, P); P = ajoute_(P, P); k >>= BigInt(1); } return R; }
function octets_(b) { return b.map(function (x) { return x < 0 ? x + 256 : x; }); }
function signes_(b) { return b.map(function (x) { return x > 127 ? x - 256 : x; }); }
function versEntier_(b) { var h = '0x'; octets_(b).forEach(function (x) { h += ('0' + x.toString(16)).slice(-2); }); return BigInt(h); }
function versOctets_(n) { var h = n.toString(16); while (h.length < 64) h = '0' + h; var o = []; for (var i = 0; i < 64; i += 2) o.push(parseInt(h.substr(i, 2), 16)); return o; }
function hmac_(cle, msg) { return octets_(Utilities.computeHmacSha256Signature(signes_(msg), signes_(cle))); }
function signerES256_(texte, dB64u) {
  var d = versEntier_(Utilities.base64DecodeWebSafe(dB64u.replace(/=+$/, '') + '==='.slice((dB64u.replace(/=+$/, '').length + 3) % 4)));
  var h = octets_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texte, Utilities.Charset.UTF_8));
  var z = versEntier_(h), x = versOctets_(d), h1 = versOctets_(mod_(z, EC_N));
  var V = [], K = [], i, k;
  for (i = 0; i < 32; i++) { V.push(1); K.push(0); }
  K = hmac_(K, V.concat([0], x, h1)); V = hmac_(K, V);
  K = hmac_(K, V.concat([1], x, h1)); V = hmac_(K, V);
  for (;;) {   // nonce déterministe (RFC 6979)
    V = hmac_(K, V); k = versEntier_(V);
    if (k > BigInt(0) && k < EC_N) {
      var r = mod_(multiplie_(k, EC_G)[0], EC_N);
      var s = mod_(inv_(k, EC_N) * (z + r * d), EC_N);
      if (r > BigInt(0) && s > BigInt(0)) return versOctets_(r).concat(versOctets_(s));
    }
    K = hmac_(K, V.concat([0])); V = hmac_(K, V);
  }
}
function b64u_(o) { return Utilities.base64EncodeWebSafe(signes_(o)).replace(/=+$/, ''); }
function jwtVapid_(aud, vapid) {
  var enc = function (o) { return Utilities.base64EncodeWebSafe(JSON.stringify(o), Utilities.Charset.UTF_8).replace(/=+$/, ''); };
  var t = enc({ typ: 'JWT', alg: 'ES256' }) + '.' + enc({ aud: aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: vapid.sub });
  return t + '.' + b64u_(signerES256_(t, vapid.d));
}

function lireAbos_() {
  var f = premier_(dossier_(), 'abonnements.json');
  try { return f ? JSON.parse(f.getBlob().getDataAsString() || '[]') : []; } catch (e) { return []; }
}
function ecrireAbos_(l) { remplacer_(dossier_(), 'abonnements.json', JSON.stringify(l)); }
function pushAutorise_(ep) {
  return ep.length < 1000 && /^https:\/\/(fcm\.googleapis\.com|updates\.push\.services\.mozilla\.com|[a-z0-9.-]+\.push\.apple\.com|[a-z0-9.-]+\.notify\.windows\.com)\//.test(ep);
}
function origine_(ep) { var m = /^https:\/\/[^\/]+/.exec(ep); return m ? m[0] : ''; }
function chaine_(s) { return typeof s === 'string' && s.length > 0 && s.length <= TAILLE_MAX; }
function pad_(n) { return ('0000000000' + (parseInt(n, 10) || 0)).slice(-10); }
function jetonValide_(t) { return typeof t === 'string' && /^[A-Za-z0-9_-]{20,64}$/.test(t); }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
