/**
 * Relais chiffré — Rondes parkings, Ville de Cachan (DPMS) — version 3
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
 *
 * REPARTIR DE ZÉRO : dans la barre d'outils de l'éditeur, choisir la fonction
 * « reinitialiserRelais » puis cliquer sur « Exécuter ». Tout le contenu du relais
 * est mis à la corbeille et les codes d'accès sont effacés.
 */
var DOSSIER = 'Relais rondes parkings (chiffré)';
var TAILLE_MAX = 30000000;      // caractères par dépôt
var PULL_OCTETS = 8000000;      // volume maximal renvoyé par relève
var PULL_NOMBRE = 150;          // nombre maximal d'éléments renvoyés par relève

function doGet() {
  return json_({ ok: true, service: 'relais-rondes', v: 3 });
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
function chaine_(s) { return typeof s === 'string' && s.length > 0 && s.length <= TAILLE_MAX; }
function pad_(n) { return ('0000000000' + (parseInt(n, 10) || 0)).slice(-10); }
function jetonValide_(t) { return typeof t === 'string' && /^[A-Za-z0-9]{20,64}$/.test(t); }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
