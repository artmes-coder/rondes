/**
 * Relais chiffré — Rondes parkings, Ville de Cachan (DPMS)
 *
 * Boîte aux lettres entre le téléphone des agents et le téléphone superviseur.
 * Tout ce qui transite ici est chiffré sur les téléphones avant l'envoi :
 * ce script et le compte Google qui l'héberge ne peuvent rien lire.
 *
 * Installation : voir README (section « Transmission automatique »).
 * Les codes d'accès sont enregistrés au premier appel du téléphone superviseur.
 * Pour les réinitialiser : Paramètres du projet > Propriétés du script > supprimer
 * SUP_TOKEN et AGENT_TOKEN, puis réactiver le relais depuis le téléphone superviseur.
 */
var DOSSIER = 'Relais rondes parkings (chiffré)';
var TAILLE_MAX = 30000000; // caractères par dépôt

function doGet() {
  return json_({ ok: true, service: 'relais-rondes', v: 1 });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, err: 'requête illisible' }); }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, err: 'relais occupé, réessayer' });
  try {
    return json_(traiter_(req));
  } catch (err) {
    return json_({ ok: false, err: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

function traiter_(req) {
  var props = PropertiesService.getScriptProperties();
  var supTok = props.getProperty('SUP_TOKEN');
  var agTok = props.getProperty('AGENT_TOKEN');

  if (req.op === 'init') {
    if (supTok) return req.supTok === supTok ? { ok: true, deja: true } : { ok: false, err: 'relais déjà initialisé par un autre superviseur' };
    if (!jetonValide_(req.supTok) || !jetonValide_(req.agentTok) || req.supTok === req.agentTok) return { ok: false, err: 'codes d’accès invalides' };
    props.setProperties({ SUP_TOKEN: req.supTok, AGENT_TOKEN: req.agentTok });
    dossier_();
    return { ok: true };
  }
  if (!supTok) return { ok: false, err: 'relais non initialisé' };
  var estSup = req.tok === supTok, estAgent = req.tok === agTok;
  if (!estSup && !estAgent) return { ok: false, err: 'accès refusé' };

  var dev = String(req.dev || '').replace(/[^A-Za-z0-9-]/g, '').slice(0, 24);
  var d = dossier_();

  switch (req.op) {
    case 'ping':
      return { ok: true, role: estSup ? 'superviseur' : 'agents' };

    case 'put': // téléphone agents : dépôt d'un paquet de saisies chiffré
      if (!estAgent || !dev) return { ok: false, err: 'interdit' };
      if (typeof req.data !== 'string' || req.data.length > TAILLE_MAX) return { ok: false, err: 'taille' };
      var nom = 'p_' + dev + '_' + pad_(req.to) + '.bin';
      remplacer_(d, nom, req.data);
      return { ok: true, nom: nom };

    case 'getEtat': // téléphone agents : mise à jour déposée par le superviseur
      if (!dev) return { ok: false, err: 'téléphone non identifié' };
      var f = premier_(d, 'e_' + dev + '.bin');
      return f ? { ok: true, data: f.getBlob().getDataAsString(), ver: f.getId() } : { ok: true, data: null };

    case 'list': // superviseur : paquets en attente
      if (!estSup) return { ok: false, err: 'interdit' };
      var out = [], it = d.getFiles();
      while (it.hasNext()) {
        var x = it.next();
        if (x.getName().indexOf('p_') === 0) out.push({ id: x.getId(), nom: x.getName(), taille: x.getSize() });
      }
      out.sort(function (a, b) { return a.nom < b.nom ? -1 : a.nom > b.nom ? 1 : 0; });
      return { ok: true, files: out };

    case 'get':
      if (!estSup) return { ok: false, err: 'interdit' };
      var g = fichier_(d, req.id);
      return g ? { ok: true, data: g.getBlob().getDataAsString() } : { ok: false, err: 'introuvable' };

    case 'del':
      if (!estSup) return { ok: false, err: 'interdit' };
      var h = fichier_(d, req.id);
      if (h) h.setTrashed(true);
      return { ok: true };

    case 'putEtat': // superviseur : mise à jour pour un téléphone agents
      if (!estSup || !dev) return { ok: false, err: 'interdit' };
      if (typeof req.data !== 'string' || req.data.length > TAILLE_MAX) return { ok: false, err: 'taille' };
      remplacer_(d, 'e_' + dev + '.bin', req.data);
      return { ok: true };
  }
  return { ok: false, err: 'opération inconnue' };
}

function dossier_() {
  var it = DriveApp.getFoldersByName(DOSSIER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(DOSSIER);
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
function fichier_(d, id) { // uniquement un fichier du dossier du relais
  var it = d.getFiles();
  while (it.hasNext()) { var x = it.next(); if (x.getId() === id) return x; }
  return null;
}
function pad_(n) { return ('0000000000' + (parseInt(n, 10) || 0)).slice(-10); }
function jetonValide_(t) { return typeof t === 'string' && /^[A-Za-z0-9]{20,64}$/.test(t); }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
