/**
 * ARE Fluides – serveur (Google Apps Script)
 * Livraison 1 : connexion, comptes, chantiers.
 *
 * Ce script est rattaché au Google Sheets « ARE Fluides - Base ».
 * Les mots de passe sont stockés hachés (SHA-256 + sel) dans l'onglet caché « _Comptes ».
 */

var APP = { name: 'ARE Fluides', version: 'L3', sessionSeconds: 6 * 60 * 60, hashRounds: 300 };
var SH = {
  params: 'Paramètres', ops: 'Opérateurs', chantiers: 'Chantiers', mouv: 'Mouvements', bout: 'Bouteilles',
  outils: 'Outillage', dist: 'Distributeurs', fluides: 'Fluides', machines: 'Machines', comptes: '_Comptes'
};

/* ------------------------------------------------------------------ Page web */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.name)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // permet l'affichage dans l'appli installable (fluides.toumany.com)
}

/** Point d'entrée unique appelé par l'interface. */
function api(req) {
  try {
    req = req || {}; OP_NAMES_ = null;
    var action = req.action, data = req.data || {};
    var pub = { status: status_, setup: setup_, login: login_, resume: resume_ };
    if (pub[action]) return ok_(pub[action](data));
    var user = auth_(req.token);
    var priv = {
      logout: function (d) { CacheService.getScriptCache().remove('s_' + req.token); forgetDevice_(d.remember); return true; },
      me: function () { return user; },
      listChantiers: listChantiers_,
      getChantier: function (d) { return getChantier_(d.id); },
      saveChantier: function (d) { return saveChantier_(d, user); },
      deleteChantier: function (d) { return deleteChantier_(d, user); },
      listFluides: listFluides_,
      listMachines: listMachines_,
      extractReleve: function (d) { return extractReleve_(d); },
      mesContext: function (d) { return mesContext_(d, user); },
      uploadPhoto: function (d) { return uploadPhoto_(d, user); },
      readPlate: function (d) { return readPlate_(d); },
      cerfaTemplate: function () { return cerfaTemplate_(); },
      saveManip: function (d) { return saveManip_(d, user); },
      saveCerfa: function (d) { return saveCerfa_(d, user); },
      cerfaData: function (d) { return cerfaData_(d, user); },
      buildReport: function (d) { return buildReport_(d, user); },
      updateReport: function (d) { return updateReport_(d, user); },
      deleteReport: function (d) { return deleteReport_(d, user); },
      updateMes: function (d) { return updateMes_(d, user); },
      deleteManip: function (d) { return deleteManip_(d, user); },
      deleteMes: function (d) { return deleteMes_(d, user); },
      changePassword: function (d) { return changePassword_(user, d); },
      listOperateurs: function () { admin_(user); return listOperateurs_(); },
      saveOperateur: function (d) { admin_(user); return saveOperateur_(d, user); },
      setOperateurActif: function (d) { admin_(user); return setOperateurActif_(d, user); },
      deleteOperateur: function (d) { admin_(user); return deleteOperateur_(d, user); },
      disconnectOperateur: function (d) { admin_(user); return disconnectOperateur_(d, user); },
      // Livraison 3
      profileSummary: function () { return profileSummary_(user); },
      updateProfile: function (d) { return updateProfile_(user, d); },
      uploadAttestation: function (d) { return uploadAttestation_(user, d); },
      listBouteilles: function () { return listBouteilles_(); },
      bouteille: function (d) { return bouteilleDetail_(d); },
      saveBouteille: function (d) { admin_(user); return saveBouteille_(d, user); },
      setBouteilleStatut: function (d) { admin_(user); return setBouteilleStatut_(d, user); },
      deleteBouteille: function (d) { admin_(user); return deleteBouteille_(d, user); },
      listDetecteurs: function () { return listDetecteurs_(); },
      saveDetecteur: function (d) { admin_(user); return saveDetecteur_(d, user); },
      listDistributeurs: function () { return distributeurs_(); },
      saveDistributeur: function (d) { admin_(user); return saveDistributeur_(d, user); },
      declaration: function (d) { admin_(user); return declaration_(d); },
      // Livraison 4
      reglages: function () { return reglages_(); },
      saveReglages: function (d) { admin_(user); return saveReglages_(d); },
      machinesBase: function () { return machinesBase_(); },
      validateMachine: function (d) { admin_(user); return validateMachine_(d, user); },
      registreApercu: function (d) { admin_(user); return registreApercu_(d); },
      exportRegistre: function (d) { admin_(user); return exportRegistre_(d); },
      alertes: function () { return alertes_(); },
      peseeControle: function (d) { return peseeControle_(d, user); },
      envoyerClient: function (d) { return envoyerClient_(d, user); },
      readCerfaPapier: function (d) { return readCerfaPapier_(d); },
      saveCerfaPapier: function (d) { return saveCerfaPapier_(d, user); },
      exportDeclaration: function (d) { admin_(user); return d.format === 'xlsx' ? exportDeclarationXlsx_(d) : exportDeclaration_(d); }
    };
    if (!priv[action]) throw new Error('Action inconnue : ' + action);
    return ok_(priv[action](data));
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

function ok_(data) { return { ok: true, data: data }; }

/* ------------------------------------------------------------------ Tableaux */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Onglet introuvable dans le tableur : ' + name);
  return sh;
}

/** Lit un onglet en objets {entête: valeur}, dates converties en jj/mm/aaaa. */
function readTable_(name) {
  var sh = sheet_(name);
  var values = sh.getDataRange().getValues();
  var headers = values.shift().map(function (h) { return String(h).trim(); });
  var tz = Session.getScriptTimeZone();
  var rows = [];
  values.forEach(function (r, i) {
    if (r.every(function (c) { return c === '' || c === null; })) return;
    var o = { _row: i + 2 };
    headers.forEach(function (h, j) {
      var v = r[j];
      if (Object.prototype.toString.call(v) === '[object Date]') v = Utilities.formatDate(v, tz, 'dd/MM/yyyy');
      o[h] = v === null ? '' : v;
    });
    rows.push(o);
  });
  return { sh: sh, headers: headers, rows: rows };
}

/** Écrit un objet dans une ligne (création si row absent). Seules les colonnes connues sont écrites. */
function writeRow_(t, obj, row) {
  var target = row || (t.sh.getLastRow() + 1);
  var current = row ? t.sh.getRange(row, 1, 1, t.headers.length).getValues()[0] : t.headers.map(function () { return ''; });
  var formulas = row ? t.sh.getRange(row, 1, 1, t.headers.length).getFormulas()[0] : null;
  var out = t.headers.map(function (h, j) {
    if (formulas && formulas[j]) return formulas[j];           // ne jamais écraser une formule
    if (Object.prototype.hasOwnProperty.call(obj, h)) return toCell_(obj[h]);
    return current[j];
  });
  t.sh.getRange(target, 1, 1, t.headers.length).setValues([out]);
  return target;
}

/** jj/mm/aaaa -> Date ; nombres "1,5" -> 1.5 ; sinon texte. */
function toCell_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v !== 'string') return v;
  var s = v.trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  // Nombre décimal (1.5 ou 1,5) : écrit en nombre, sinon le tableur français lit « 1.5 » comme une date (1er mai)
  if (/^-?\d+[.,]\d+$/.test(s)) return Number(s.replace(',', '.'));
  return s;
}

function columnLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function nextId_(rows, prefix, key) {
  var max = 0;
  rows.forEach(function (r) {
    var m = String(r[key] || '').match(new RegExp('^' + prefix + '(\\d+)$'));
    if (m) max = Math.max(max, +m[1]);
  });
  return prefix + ('000' + (max + 1)).slice(-3);
}

/* ------------------------------------------------------------------ Comptes */

function comptes_() {
  var sh = ss_().getSheetByName(SH.comptes);
  if (!sh) {
    sh = ss_().insertSheet(SH.comptes);
    sh.getRange(1, 1, 1, 4).setValues([['E-mail', 'Sel', 'Empreinte', 'Mis à jour']]);
    sh.hideSheet();
  }
  return readTable_(SH.comptes);
}

function hash_(password, salt) {
  var h = salt + '|' + password;
  for (var i = 0; i < APP.hashRounds; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + '|' + salt, Utilities.Charset.UTF_8));
  }
  return h;
}

function setPassword_(email, password) {
  if (!password || String(password).length < 8) throw new Error('Le mot de passe doit faire au moins 8 caractères.');
  var t = comptes_();
  var key = String(email).trim().toLowerCase();
  var salt = Utilities.getUuid();
  var rec = { 'E-mail': key, 'Sel': salt, 'Empreinte': hash_(password, salt), 'Mis à jour': new Date() };
  var existing = t.rows.filter(function (r) { return String(r['E-mail']).toLowerCase() === key; })[0];
  writeRow_(t, rec, existing ? existing._row : null);
}

function checkPassword_(email, password) {
  var key = String(email).trim().toLowerCase();
  var rec = comptes_().rows.filter(function (r) { return String(r['E-mail']).toLowerCase() === key; })[0];
  if (!rec) return false;
  return hash_(password, rec['Sel']) === rec['Empreinte'];
}

function renamePassword_(oldEmail, newEmail) {
  var c = comptes_();
  var rec = c.rows.filter(function (r) { return String(r['E-mail']).toLowerCase() === oldEmail; })[0];
  if (rec) writeRow_(c, { 'E-mail': newEmail, 'Mis à jour': new Date() }, rec._row);
}

function hasPassword_(email) {
  var key = String(email || '').trim().toLowerCase();
  return !!key && comptes_().rows.some(function (r) { return String(r['E-mail']).toLowerCase() === key; });
}

function operatorPublic_(o) {
  return {
    id: o['ID'], prenom: o['Prénom'] || '', nom: o['Nom'] || '', email: o['E-mail (identifiant)'] || '', email2: o['E-mail secondaire'] || '',
    role: String(o['Rôle'] || '').toLowerCase() === 'admin' ? 'admin' : 'operateur',
    attestation: o['N° attestation d’aptitude'] ? String(o['N° attestation d’aptitude']) : '',
    attestationChargee: !!String(o['Attestation (fichier)'] || '').trim() || (/^oui/i.test(String(o['Attestation chargée'] || '')) && !/recharger/i.test(String(o['Attestation chargée'] || ''))),
    attestationDate: (String(o['Attestation chargée'] || '').match(/(\d{2}\/\d{2}\/\d{4})/) || [])[1] || '',
    actif: String(o['Actif'] === undefined ? 'oui' : (o['Actif'] || 'oui')).toLowerCase() !== 'non'
  };
}

/** Ajoute une colonne à un onglet si elle n'existe pas encore (mise à niveau du tableur). */
function ensureColumn_(name, header, fill) {
  var sh = sheet_(name);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  if (headers.indexOf(header) >= 0) return;
  var col = headers.length + 1;
  sh.getRange(1, col).setValue(header);
  var last = sh.getLastRow();
  if (fill !== undefined && last > 1) {
    var vals = []; for (var i = 2; i <= last; i++) vals.push([fill]);
    sh.getRange(2, col, last - 1, 1).setValues(vals);
  }
}

function setOperateurActif_(d, user) {
  if (d.id === user.id) throw new Error('Tu ne peux pas désactiver ton propre compte.');
  ensureColumn_(SH.ops, 'Actif', 'oui');
  var t = readTable_(SH.ops);
  var row = t.rows.filter(function (r) { return r['ID'] === d.id; })[0];
  if (!row) throw new Error('Compte introuvable.');
  writeRow_(t, { 'Actif': d.actif ? 'oui' : 'non' }, row._row);
  if (!d.actif) { revokeRemember_(d.id); cutSessions_(d.id); }
  return listOperateurs_();
}

function deleteOperateur_(d, user) {
  if (d.id === user.id) throw new Error('Tu ne peux pas supprimer ton propre compte.');
  var t = readTable_(SH.ops);
  var row = t.rows.filter(function (r) { return r['ID'] === d.id; })[0];
  if (!row) throw new Error('Compte introuvable.');
  var name = ((row['Prénom'] || '') + ' ' + (row['Nom'] || '')).trim().toLowerCase();
  var email = String(row['E-mail (identifiant)'] || '').toLowerCase();
  var used = name && readTable_(SH.mouv).rows.some(function (m) {
    var o = String(m['Opérateur'] || '').toLowerCase();
    return o === name || o === String(row['ID']).toLowerCase() || (email && o === email);
  });
  if (used) throw new Error('Ce compte apparaît dans le registre des mouvements : il ne peut pas être supprimé, désactive-le plutôt.');
  t.sh.deleteRow(row._row);
  revokeRemember_(d.id);
  if (email) {
    var c = comptes_();
    var rec = c.rows.filter(function (r) { return String(r['E-mail']).toLowerCase() === email; })[0];
    if (rec) c.sh.deleteRow(rec._row);
  }
  return listOperateurs_();
}

function status_() {
  var ops = readTable_(SH.ops).rows.map(operatorPublic_);
  var adminReady = ops.some(function (o) { return o.role === 'admin' && o.actif && o.email && hasPassword_(o.email); });
  var owner = '';
  try { owner = Session.getEffectiveUser().getEmail(); } catch (e) { /* pas d'accès */ }
  var adminRow = ops.filter(function (o) { return o.role === 'admin'; })[0];
  return { needsSetup: !adminReady, ownerEmail: owner, adminPrenom: adminRow ? adminRow.prenom : '', version: APP.version };
}

/** Première connexion : crée le mot de passe de l'admin. Réservé au propriétaire du compte Google. */
function setup_(d) {
  if (!status_().needsSetup) throw new Error('Le compte administrateur existe déjà.');
  var owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  var email = String(d.email || '').trim().toLowerCase();
  if (!owner || email !== owner) throw new Error('Pour la première connexion, utilise l’adresse Gmail propriétaire de l’appli.');
  var t = readTable_(SH.ops);
  var adminRow = t.rows.filter(function (r) { return String(r['Rôle']).toLowerCase() === 'admin'; })[0];
  var rec = { 'Prénom': d.prenom || (adminRow ? adminRow['Prénom'] : ''), 'Nom': d.nom || (adminRow ? adminRow['Nom'] : ''), 'E-mail (identifiant)': email, 'Rôle': 'admin' };
  if (!adminRow) rec['ID'] = nextId_(t.rows, 'OP', 'ID');
  writeRow_(t, rec, adminRow ? adminRow._row : null);
  setPassword_(email, d.password);
  return login_({ email: email, password: d.password, remember: d.remember !== false });
}

function login_(d) {
  var email = String(d.email || '').trim().toLowerCase();
  Utilities.sleep(300); // ralentit les essais en rafale
  // on accepte l'e-mail principal ou l'e-mail secondaire ; le mot de passe est rattaché à l'e-mail principal
  var op = readTable_(SH.ops).rows.map(operatorPublic_).filter(function (o) {
    return email && (o.email.toLowerCase() === email || String(o.email2).toLowerCase() === email);
  })[0];
  if (!op || !checkPassword_(op.email, d.password || '')) throw new Error('E-mail ou mot de passe incorrect.');
  if (!op.actif) throw new Error('Ce compte est désactivé. Contacte l’administrateur.');
  var out = issueSession_(op);
  if (d.remember) out.remember = createRemember_(op.id);
  return out;
}

function issueSession_(op) {
  var token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  CacheService.getScriptCache().put('s_' + token, JSON.stringify({ id: op.id, email: op.email, at: Date.now() }), APP.sessionSeconds);
  return { token: token, user: op };
}

/** Coupe les sessions en cours d'un opérateur : toute session ouverte avant cet instant est refusée. */
function cutSessions_(opId) {
  PropertiesService.getScriptProperties().setProperty('cut_' + opId, String(Date.now()));
}

/** Admin : déconnecte un opérateur de tous ses appareils. */
function disconnectOperateur_(d, user) {
  if (d.id === user.id) throw new Error('Pour toi-même, utilise « Se déconnecter » dans ton profil.');
  cutSessions_(d.id);
  revokeRemember_(d.id);
  return listOperateurs_();
}

/* ------------------------------------------------------------------ « Rester connecté » (30 jours glissants) */

var REMEMBER_DAYS = 30;

function sessions_() {
  var sh = ss_().getSheetByName('_Sessions');
  if (!sh) {
    sh = ss_().insertSheet('_Sessions');
    sh.getRange(1, 1, 1, 5).setValues([['Empreinte', 'ID opérateur', 'Créé le (ms)', 'Expire le (ms)', 'Dernière utilisation (ms)']]);
    sh.hideSheet();
  }
  return readTable_('_Sessions');
}

function tokenHash_(token) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'remember|' + token, Utilities.Charset.UTF_8));
}

function createRemember_(opId) {
  var t = sessions_();
  var now = Date.now();
  // ménage : retire les jetons expirés
  t.rows.filter(function (r) { return Number(r['Expire le (ms)']) < now; }).reverse().forEach(function (r) { t.sh.deleteRow(r._row); });
  t = sessions_();
  var token = Utilities.getUuid() + Utilities.getUuid();
  writeRow_(t, { 'Empreinte': tokenHash_(token), 'ID opérateur': opId, 'Créé le (ms)': now, 'Expire le (ms)': now + REMEMBER_DAYS * 864e5, 'Dernière utilisation (ms)': now });
  return token;
}

/** Reconnexion automatique : valide le jeton, prolonge de 30 jours, ouvre une session. */
function resume_(d) {
  if (!d.remember) throw new Error('SESSION');
  var t = sessions_();
  var h = tokenHash_(d.remember), now = Date.now();
  var rec = t.rows.filter(function (r) { return r['Empreinte'] === h; })[0];
  if (!rec || Number(rec['Expire le (ms)']) < now) throw new Error('SESSION');
  var op = readTable_(SH.ops).rows.map(operatorPublic_).filter(function (o) { return o.id === rec['ID opérateur']; })[0];
  if (!op || !op.actif) { t.sh.deleteRow(rec._row); throw new Error('SESSION'); }
  writeRow_(t, { 'Expire le (ms)': now + REMEMBER_DAYS * 864e5, 'Dernière utilisation (ms)': now }, rec._row);
  return issueSession_(op);
}

/** Supprime les jetons « rester connecté » d'un opérateur (sauf, éventuellement, celui de l'appareil en cours). */
function revokeRemember_(opId, keepToken) {
  var t = sessions_();
  var keep = keepToken ? tokenHash_(keepToken) : null;
  t.rows.filter(function (r) { return r['ID opérateur'] === opId && r['Empreinte'] !== keep; })
    .reverse().forEach(function (r) { t.sh.deleteRow(r._row); });
}

function forgetDevice_(token) {
  if (!token) return;
  var t = sessions_(), h = tokenHash_(token);
  t.rows.filter(function (r) { return r['Empreinte'] === h; }).reverse().forEach(function (r) { t.sh.deleteRow(r._row); });
}

function auth_(token) {
  if (!token) throw new Error('SESSION');
  var raw = CacheService.getScriptCache().get('s_' + token);
  if (!raw) throw new Error('SESSION');
  var s = JSON.parse(raw);
  var op = readTable_(SH.ops).rows.map(operatorPublic_).filter(function (o) { return o.id === s.id; })[0];
  if (!op || !op.actif || op.email.toLowerCase() !== s.email.toLowerCase()) throw new Error('SESSION');
  var cut = Number(PropertiesService.getScriptProperties().getProperty('cut_' + op.id) || 0);
  if (cut && (s.at || 0) <= cut) { CacheService.getScriptCache().remove('s_' + token); throw new Error('SESSION'); }
  return op;
}

function admin_(user) { if (user.role !== 'admin') throw new Error('Réservé à l’administrateur.'); }

function changePassword_(user, d) {
  if (!checkPassword_(user.email, d.oldPassword || '')) throw new Error('Mot de passe actuel incorrect.');
  setPassword_(user.email, d.newPassword);
  revokeRemember_(user.id, d.remember); // déconnecte les autres appareils, garde celui-ci
  return true;
}

function listOperateurs_() {
  var now = Date.now();
  var sess = sessions_().rows.filter(function (r) { return Number(r['Expire le (ms)']) >= now; });
  var tz = Session.getScriptTimeZone();
  return readTable_(SH.ops).rows.map(function (o) {
    var p = operatorPublic_(o);
    p.hasPassword = hasPassword_(p.email);
    var mine = sess.filter(function (r) { return r['ID opérateur'] === p.id; });
    p.appareils = mine.length;
    var last = mine.reduce(function (m, r) { return Math.max(m, Number(r['Dernière utilisation (ms)']) || 0); }, 0);
    p.derniereUtilisation = last ? Utilities.formatDate(new Date(last), tz, 'dd/MM/yyyy HH:mm') : '';
    return p;
  });
}

function saveOperateur_(d, user) {
  ensureColumn_(SH.ops, 'E-mail secondaire');
  var t = readTable_(SH.ops);
  var email = String(d.email || '').trim().toLowerCase();
  var email2 = String(d.email2 || '').trim().toLowerCase();
  if (!email || email.indexOf('@') < 1) throw new Error('Adresse e-mail invalide.');
  if (email2 && email2.indexOf('@') < 1) throw new Error('Adresse e-mail secondaire invalide.');
  if (email2 && email2 === email) throw new Error('L’e-mail secondaire doit être différent de l’e-mail principal.');
  var taken = function (addr) {
    return addr && t.rows.some(function (r) {
      return r['ID'] !== d.id && [r['E-mail (identifiant)'], r['E-mail secondaire']].some(function (x) { return String(x || '').toLowerCase() === addr; });
    });
  };
  if (taken(email) || taken(email2)) throw new Error('Cette adresse e-mail est déjà utilisée par un autre compte.');
  var row = d.id ? t.rows.filter(function (r) { return r['ID'] === d.id; })[0] : null;
  if (row && row['ID'] === user.id && d.role !== 'admin') throw new Error('Tu ne peux pas retirer ton propre rôle admin.');
  var oldEmail = row ? String(row['E-mail (identifiant)'] || '').toLowerCase() : '';
  if (!row && !d.password) throw new Error('Définis un mot de passe pour ce nouveau compte.');
  if (d.password && String(d.password).length < 8) throw new Error('Le mot de passe doit faire au moins 8 caractères.');
  var rec = {
    'Prénom': d.prenom || '', 'Nom': d.nom || '', 'E-mail (identifiant)': email, 'E-mail secondaire': email2,
    'Rôle': d.role === 'admin' ? 'admin' : 'opérateur', 'N° attestation d’aptitude': d.attestation || ''
  };
  if (!row) rec['ID'] = nextId_(t.rows, 'OP', 'ID');
  writeRow_(t, rec, row ? row._row : null);
  if (d.password) {
    setPassword_(email, d.password);
    if (row && row['ID'] !== user.id) { revokeRemember_(row['ID']); cutSessions_(row['ID']); } // mot de passe réinitialisé : appareils déconnectés
    else if (row) revokeRemember_(row['ID'], d.remember);
  }
  else if (row && oldEmail && oldEmail !== email) renamePassword_(oldEmail, email); // garde le mot de passe si l'e-mail principal change
  return listOperateurs_();
}

/* ------------------------------------------------------------------ Chantiers */

var CH_FIELDS = {
  id: 'ID', projet: 'N° projet', dateMes: 'Date MES', type: 'Type d’installation', prenom: 'Prénom', nom: 'Nom',
  adresse: 'Adresse', cp: 'Code postal', ville: 'Ville', tel: 'Téléphone', email: 'E-mail',
  marque: 'Marque UE', modele: 'Modèle UE', serieUE: 'N° série UE', serieUI: 'N° série UI', fluide: 'Fluide',
  inflam: 'Inflammable (si fluide inconnu)', longueur: 'Longueur liaisons (m)', precharge: 'Précharge (m)', gpm: 'Charge (g/m)', chargeUsine: 'Charge usine (kg)',
  commentaires: 'Commentaires'
};

function gwpMap_() {
  var m = {};
  try { listFluides_().forEach(function (f) { if (f.gwp) m[normFluide_(f.code)] = f.gwp; }); } catch (e) {}
  return m;
}
function normFluide_(s) { return String(s || '').toUpperCase().replace(/\s+/g, '').replace(/^R-?/, 'R'); }

function chantierFromRow_(r, gwps) {
  var o = {};
  Object.keys(CH_FIELDS).forEach(function (k) { o[k] = r[CH_FIELDS[k]] === undefined ? '' : String(r[CH_FIELDS[k]]); });
  var g = r['GWP'];
  // Ancienne formule en erreur (#ERROR!, #N/A…) ou vide : on reprend la valeur de l'onglet Fluides
  if (g === '' || g === 'saisir' || /^#/.test(String(g)) || !isFinite(Number(g))) g = (gwps || gwpMap_())[normFluide_(o.fluide)] || '';
  o.gwp = g === '' ? '' : String(g).replace('.', ',');
  o.creePar = String(r['Créé par'] || '');
  // Responsable du chantier : désigné par l'admin, sinon celui qui l'a créé. Il a les droits du créateur.
  o.responsable = String(r['Responsable'] || '') || o.creePar;
  o.responsableNom = opNames_()[o.responsable] || '';
  o.docsEnvoyes = String(r['Documents envoyés'] || '');
  o.client = (o.prenom ? o.prenom + ' ' : '') + o.nom;
  o.misEnService = !!o.dateMes;
  // Année du chantier : année de mise en service, sinon année de création, sinon année en cours (chantier à mettre en service)
  var an = function (v) { var m = String(v || '').match(/(\d{4})/); return m ? m[1] : ''; };
  o.annee = an(o.dateMes) || an(r['Créé le']) || String(new Date().getFullYear());
  return o;
}

var OP_NAMES_ = null;   // noms des comptes, lus une fois par appel
function opNames_() {
  if (OP_NAMES_) return OP_NAMES_;
  OP_NAMES_ = {};
  try { readTable_(SH.ops).rows.forEach(function (o) { OP_NAMES_[o['ID']] = ((o['Prénom'] || '') + ' ' + (o['Nom'] || '')).trim(); }); } catch (e) {}
  return OP_NAMES_;
}

function listChantiers_() {
  var gwps = gwpMap_();
  return readTable_(SH.chantiers).rows.map(function (r) { return chantierFromRow_(r, gwps); }).reverse();
}

function getChantier_(id) {
  var r = readTable_(SH.chantiers).rows.filter(function (x) { return x['ID'] === id; })[0];
  if (!r) throw new Error('Chantier introuvable.');
  return chantierFromRow_(r);
}

function saveChantier_(d, user) {
  if (!String(d.nom || '').trim()) throw new Error('Le nom du client est obligatoire.');
  if (d.dateMes && !/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(String(d.dateMes).trim())) throw new Error('Date de mise en service au format jj/mm/aaaa.');
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    ensureColumn_(SH.chantiers, 'Charge usine (kg)');
    ensureColumn_(SH.chantiers, 'Créé par');
    ensureColumn_(SH.chantiers, 'Créé le');
    ensureColumn_(SH.chantiers, 'Responsable');
    var t = readTable_(SH.chantiers);
    var row = d.id ? t.rows.filter(function (r) { return r['ID'] === d.id; })[0] : null;
    if (d.id && !row) throw new Error('Chantier introuvable.');
    if (row && user && user.role !== 'admin' && (String(row['Responsable'] || '') || String(row['Créé par'] || '')) !== user.id) throw new Error('Tu ne peux modifier que les chantiers dont tu es responsable.');
    var rec = {};
    Object.keys(CH_FIELDS).forEach(function (k) { if (k !== 'id' && d[k] !== undefined) rec[CH_FIELDS[k]] = d[k]; });
    if (!row) {
      rec['ID'] = nextId_(t.rows.concat(corbeille_().rows), 'C', 'ID');   // jamais réutiliser l'ID d'un chantier supprimé
      if (user) { rec['Créé par'] = user.id; rec['Responsable'] = user.id; }
      rec['Créé le'] = new Date();
    }
    // Seul l'admin désigne ou change le responsable (un opérateur actif ou lui-même)
    if (user && user.role === 'admin' && d.responsable) {
      var op = readTable_(SH.ops).rows.filter(function (o) { return o['ID'] === d.responsable; })[0];
      if (!op) throw new Error('Responsable inconnu.');
      if (String(op['Actif'] || 'oui').toLowerCase() === 'non') throw new Error('Ce compte est désactivé : choisis un autre responsable.');
      rec['Responsable'] = d.responsable;
    }
    var r = writeRow_(t, rec, row ? row._row : null);
    // GWP : valeur saisie à la main si fournie, sinon valeur de l'onglet Fluides (écrite en nombre, sans formule)
    var gwpCol = t.headers.indexOf('GWP') + 1;
    if (gwpCol > 0) {
      var manual = Number(String(d.gwpManuel || '').replace(',', '.'));
      var gwp = d.gwpManuel && isFinite(manual) ? manual : (gwpMap_()[normFluide_(d.fluide)] || '');
      t.sh.getRange(r, gwpCol).setValue(gwp);
    }
    SpreadsheetApp.flush();
    var id = row ? d.id : rec['ID'];
    var out = getChantier_(id);
  } finally { lock.releaseLock(); }
  // Base machines : ajout d'un nouveau modèle ou mise à jour de ses données, si l'opérateur l'a choisi
  if (d.machine === 'add' || d.machine === 'update') {
    try { out.machineMsg = upsertMachine_(d, user, id); } catch (e) { out.machineMsg = 'Base machines non mise à jour : ' + e.message; }
  }
  return out;
}

/* ---------- Base machines alimentée depuis l'appli */

var MACH_COLS = ['Marque', 'Modèle', 'Type d’appareil', 'Gamme', 'Type d’installation', 'Alimentation', 'Liaisons', 'Fluide', 'Charge usine (kg)',
  'Longueur préchargée (m)', 'Charge supplémentaire (g/m)', 'Longueur max UE-UI (m)', 'Source', 'Remarque'];
var MACH_DATA = [['fluide', 'Fluide', 'fluide'], ['chargeUsine', 'Charge usine (kg)', 'charge usine', 'kg'], ['precharge', 'Longueur préchargée (m)', 'longueur préchargée', 'm'], ['gpm', 'Charge supplémentaire (g/m)', 'charge supplémentaire', 'g/m']];

/** Ajoute le modèle à l'onglet Machines, ou met à jour ses données. L'admin écrit directement ; un opérateur ajoute un modèle
 *  (marqué « à vérifier ») ou propose une modification, que l'admin valide dans l'onglet Machines. */
function upsertMachine_(d, user, chantierId) {
  var modele = String(d.modele || '').trim(), marque = String(d.marque || '').trim();
  if (!modele || !marque) return '';
  ensureSheet_(SH.machines, MACH_COLS);
  ensureColumn_(SH.machines, 'Source'); ensureColumn_(SH.machines, 'À vérifier'); ensureColumn_(SH.machines, 'Propositions (JSON)');
  var t = readTable_(SH.machines), k = normModele_(modele);
  var row = t.rows.filter(function (r) { return normModele_(r['Modèle'] || r['Modèle UE']) === k; })[0];
  var admin = user && user.role === 'admin';
  var qui = user ? ((user.prenom || '') + ' ' + (user.nom || '')).trim() : 'l’appli';
  var quand = today_();
  var val = function (key) {
    var v = String(d[key] === undefined || d[key] === null ? '' : d[key]).trim();
    if (key === 'fluide') return normFluide_(v);
    var n = num_(v); return isFinite(n) ? n : '';
  };
  var same = function (a, b) {
    if (typeof a === 'number' || typeof b === 'number') { var x = num_(a), y = num_(b); return isFinite(x) && isFinite(y) && Math.abs(x - y) < 1e-6; }
    return normFluide_(a) === normFluide_(b);
  };
  var vide = function (v) { return v === '' || v === null || v === undefined || /^n\.?a\.?$/i.test(String(v).trim()); };
  if (!row) {
    var pre = val('precharge'), gpm = val('gpm');
    var rec = { 'Marque': marque, 'Modèle': modele, 'Type d’appareil': 'Unité extérieure', 'Type d’installation': String(d.type || ''),
      'Liaisons': (pre !== '' && pre > 0) || (gpm !== '' && gpm > 0) ? 'Frigorifique' : '',
      'Fluide': val('fluide'), 'Charge usine (kg)': val('chargeUsine'), 'Longueur préchargée (m)': pre, 'Charge supplémentaire (g/m)': gpm,
      'Source': 'Ajouté depuis l’appli par ' + qui + ' le ' + quand + (chantierId ? ' (chantier ' + chantierId + ')' : ''),
      'À vérifier': admin ? '' : 'Nouveau modèle ajouté par ' + qui + ' le ' + quand,
      'Propositions (JSON)': admin ? '' : JSON.stringify([{ nouveau: true, par: qui, date: quand }]) };
    writeRow_(t, rec, null);
    return admin ? 'Modèle ' + modele + ' ajouté à la base machines.' : 'Modèle ' + modele + ' ajouté à la base machines (l’administrateur vérifiera ses données).';
  }
  var changes = [], up = {};
  MACH_DATA.forEach(function (f) {
    var v = val(f[0]); if (v === '') return;
    var old = row[f[1]];
    if (!vide(old) && same(old, v)) return;
    changes.push(f[2] + ' ' + (typeof v === 'number' ? fr_(v, f[0] === 'gpm' ? 0 : 2) : v) + (f[3] ? ' ' + f[3] : '') + (vide(old) ? ' (vide dans la base)' : ' au lieu de ' + (typeof v === 'number' ? fr_(num_(old), f[0] === 'gpm' ? 0 : 2) : old)));
    up[f[1]] = v;
  });
  if (!changes.length) return '';
  if (admin) {
    up['Source'] = (String(row['Source'] || '') + (row['Source'] ? ' · ' : '') + 'mis à jour par ' + qui + ' le ' + quand).slice(0, 1000);
    writeRow_(t, up, row._row);
    return 'Base machines mise à jour pour ' + row['Modèle'] + ' : ' + changes.join(', ') + '.';
  }
  var props = []; try { props = JSON.parse(row['Propositions (JSON)'] || '[]'); } catch (e) { props = []; }
  props.push({ par: qui, date: quand, up: up, texte: changes.join(', ') });
  writeRow_(t, { 'À vérifier': (String(row['À vérifier'] || '') + (row['À vérifier'] ? ' · ' : '') + qui + ' propose le ' + quand + ' : ' + changes.join(', ')).slice(0, 1000),
    'Propositions (JSON)': JSON.stringify(props).slice(0, 20000) }, row._row);
  return 'Proposition envoyée : l’administrateur validera la mise à jour de ' + row['Modèle'] + ' dans la base machines.';
}

var CORBEILLE = '_Chantiers supprimés';

/** Onglet caché qui garde une copie de chaque chantier supprimé (avec la date et l'auteur). */
function corbeille_() {
  var sh = ss_().getSheetByName(CORBEILLE);
  if (!sh) {
    var src = sheet_(SH.chantiers);
    var h = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0].concat(['Supprimé le', 'Supprimé par']);
    sh = ss_().insertSheet(CORBEILLE);
    sh.getRange(1, 1, 1, h.length).setValues([h]);
    sh.hideSheet();
  }
  return readTable_(CORBEILLE);
}

/** Supprime une fiche chantier : seulement sans mise en service ; un opérateur ne supprime que les siennes. */
function deleteChantier_(d, user) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readTable_(SH.chantiers);
    var row = t.rows.filter(function (r) { return r['ID'] === d.id; })[0];
    if (!row) throw new Error('Chantier introuvable.');
    if (String(row['Date MES'] || '').trim()) throw new Error('Ce chantier a une mise en service : il ne peut pas être supprimé.');
    var mouv = [];
    try { mouv = readTable_(SH.mouv).rows; } catch (e) { mouv = []; }
    if (mouv.some(function (m) { return String(m['ID chantier'] || '').trim() === d.id || String(m['Chantier / Client'] || '').trim() === d.id; })) throw new Error('Ce chantier a des opérations sur le fluide : il ne peut pas être supprimé.');
    if (user.role !== 'admin' && (String(row['Responsable'] || '') || String(row['Créé par'] || '')) !== user.id) throw new Error('Tu ne peux supprimer que les chantiers dont tu es responsable.');
    var c = corbeille_();
    var rec = {};
    t.headers.forEach(function (h) { if (h) rec[h] = row[h]; });
    rec['Supprimé le'] = new Date(); rec['Supprimé par'] = (user.prenom + ' ' + user.nom).trim() || user.email;
    writeRow_(c, rec, null);
    t.sh.deleteRow(row._row);
    return true;
  } finally { lock.releaseLock(); }
}

/* ------------------------------------------------------------------ Fluides */

function listFluides_() {
  var t = readTable_(SH.fluides);
  // R290 (propane) des PAC monoblocs : ajouté automatiquement s'il manque (GWP 0,02 d'après le catalogue Panasonic)
  if (!t.rows.some(function (f) { return String(f['Code']).toUpperCase() === 'R290'; })) {
    writeRow_(t, { 'Code': 'R290', 'GWP': 0.02, 'Classe de sécurité': 'A3', 'Inflammable': 'oui',
      'Rubrique déchets (case 12)': '16 05 04* – Gaz en récipients à pression contenant des substances dangereuses', 'Mention ADR (case 12)': 'UN 1978, Propane, 2.1' });
    t = readTable_(SH.fluides);
  }
  return t.rows.map(function (f) {
    return { code: String(f['Code']), gwp: Number(f['GWP']) || 0, classe: f['Classe de sécurité'], inflammable: String(f['Inflammable']).toLowerCase() === 'oui' };
  });
}

/* ------------------------------------------------------------------ Machines (base des PAC) */

/** Onglet « Machines » : tous les appareils (UE, UI, modules, ballons). Absent = liste vide (l'appli reste utilisable). */
function listMachines_() {
  if (!ss_().getSheetByName(SH.machines)) return [];
  // « NA » = non applicable : traité comme vide
  var dec = function (v) { return v === '' || v === null || v === undefined || /^n\.?a\.?$/i.test(String(v).trim()) ? '' : String(v).replace('.', ','); };
  var mod = function (r) { return String(r['Modèle'] || r['Modèle UE'] || '').trim(); };
  // Une seule liste dans l'onglet : on ne propose que les unités extérieures et les monoblocs
  var interieur = /^(Unité intérieure|Module hydraulique|Ballon|Split|Gainable)/i;
  return readTable_(SH.machines).rows.filter(function (r) {
    if (!mod(r)) return false;
    if (/^Aucune/i.test(String(r['Liaisons'] || ''))) return true;   // RAC Solo : appareil complet sans unité extérieure
    return !interieur.test(String(r['Type d’appareil'] || '')) && !interieur.test(String(r['Catégorie'] || ''));
  }).map(function (r) {
    return {
      marque: String(r['Marque'] || '').trim(), modele: mod(r), gamme: String(r['Gamme'] || ''),
      type: TYPES_INSTALL.indexOf(String(r['Type d’installation'])) >= 0 ? String(r['Type d’installation']) : '',
      liaison: String(r['Liaisons'] || ''), fluide: String(r['Fluide'] || '').toUpperCase().replace(/\s+/g, '').replace(/^R-?/, 'R'),
      charge: dec(r['Charge usine (kg)']), precharge: dec(r['Longueur préchargée (m)']), gpm: dec(r['Charge supplémentaire (g/m)']),
      phase: String(r['Alimentation'] || ''), lmax: dec(r['Longueur max UE-UI (m)']),
      remarque: String(r['Remarque'] || ''), appareil: String(r['Type d’appareil'] || '')
    };
  });
}

function normModele_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }

/** Retrouve une machine à partir d'un modèle lu sur une feuille ou une plaque (tolère tirets, espaces et suffixes). */
function findMachine_(modele) {
  var k = normModele_(modele);
  if (k.length < 4) return null;
  var list = listMachines_();
  var exact = list.filter(function (m) { return normModele_(m.modele) === k; })[0];
  if (exact) return exact;
  var near = list.filter(function (m) { var n = normModele_(m.modele); return n.length >= 5 && (k.indexOf(n) === 0 || n.indexOf(k) === 0); });
  return near.length === 1 ? near[0] : null;
}

/* ------------------------------------------------------------------ Lecture de la feuille de relevé (API Claude) */

var CLAUDE = { url: 'https://api.anthropic.com/v1/messages', model: 'claude-sonnet-5-5', version: '2023-06-01' };
var TYPES_INSTALL = ['PAC AE', 'PAC AA Monosplit', 'PAC AA Gainable', 'PAC AA Multisplit', 'PAC AA Multisplit gainable', 'Ballon thermodynamique'];

function claudeKey_() {
  var k = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!k) throw new Error('Lecture automatique pas encore activée : la clé API Claude n’est pas configurée dans le script.');
  return k;
}

function callClaude_(content, maxTokens) {
  var res = UrlFetchApp.fetch(CLAUDE.url, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': claudeKey_(), 'anthropic-version': CLAUDE.version },
    payload: JSON.stringify({ model: CLAUDE.model, max_tokens: maxTokens || 1200, messages: [{ role: 'user', content: content }] })
  });
  var code = res.getResponseCode(), body = {};
  try { body = JSON.parse(res.getContentText()); } catch (e) { /* réponse non JSON */ }
  if (code === 401) throw new Error('Clé API Claude refusée : vérifie la clé enregistrée dans le script.');
  if (code === 400 && /credit/i.test(res.getContentText())) throw new Error('Plus de crédit sur le compte API Claude : recharge-le dans la Console Claude.');
  if (code === 429 || code === 529) throw new Error('Le service de lecture est occupé, réessaie dans une minute.');
  if (code !== 200) throw new Error('Lecture impossible (erreur ' + code + ')' + (body.error && body.error.message ? ' : ' + body.error.message : '.'));
  return (body.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
}

/** Reçoit une photo (ou un PDF) de la première page de la feuille de relevé et renvoie les champs du chantier. */
function extractReleve_(d) {
  var mime = String(d.mime || '');
  var data = String(d.data || '');
  if (!data) throw new Error('Aucune image reçue.');
  if (data.length > 25 * 1024 * 1024) throw new Error('Fichier trop lourd : prends une photo de la première page seulement.');
  var isPdf = mime === 'application/pdf';
  if (!isPdf && ['image/jpeg', 'image/png', 'image/webp', 'image/gif'].indexOf(mime) < 0) throw new Error('Format non pris en charge : envoie une photo (JPEG, PNG) ou un PDF.');
  var block = isPdf
    ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: data } }
    : { type: 'image', source: { type: 'base64', media_type: mime, data: data } };
  var prompt = [
    'Tu lis la feuille de relevé (visite technique) d’une entreprise française qui installe des pompes à chaleur chez des particuliers.',
    isPdf ? 'Ne regarde que la première page du document.' : 'C’est une photo de la première page, parfois remplie à la main.',
    'Extrais les informations suivantes et réponds UNIQUEMENT par un objet JSON, sans texte autour :',
    '{',
    '  "prenom": prénom du client,',
    '  "nom": nom de famille du client (si un couple : "M. et Mme X" -> nom "X"),',
    '  "adresse": numéro et rue du chantier,',
    '  "cp": code postal (5 chiffres),',
    '  "ville": ville,',
    '  "projet": numéro de projet / de dossier / de devis,',
    '  "type": une valeur parmi ' + JSON.stringify(TYPES_INSTALL) + ' (PAC air/eau = "PAC AE"),',
    '  "marque": marque de l’unité extérieure (ex. DAIKIN, VIESSMANN),',
    '  "modele": modèle / référence de l’unité extérieure (ex. ERGA04EV),',
    '  "fluide": fluide frigorigène (ex. R32, R410A),',
    '  "longueur": longueur des liaisons frigorifiques en mètres (nombre),',
    '  "precharge": longueur de précharge en mètres (nombre),',
    '  "gpm": charge complémentaire en grammes par mètre (nombre),',
    '  "puissance": puissance de la PAC en kW si indiquée (nombre),',
    '  "alimentation": "monophasé" ou "triphasé" si indiqué (installation électrique ou machine),',
    '  "ecs": "oui" si la production d’eau chaude sanitaire par la PAC est prévue, "non" si elle est exclue, sinon "",',
    '  "suggestions": jusqu’à 3 modèles de la LISTE DES MACHINES ci-dessous qui correspondent le mieux à l’installation décrite, du plus probable au moins probable, sous la forme [{"modele": "référence exacte de la liste", "raison": "courte justification"}] ; tiens compte de la marque, de la référence écrite (même incomplète ou mal orthographiée), du type d’installation, de la puissance, du monophasé/triphasé, de l’eau chaude sanitaire et du nombre de pièces ou de zones ; [] si rien ne correspond,',
    '  "illisible": liste des champs présents sur la feuille mais que tu n’as pas pu lire avec certitude',
    '}',
    'Règles : mets "" pour toute information absente ou illisible ; ne devine jamais (sauf pour "suggestions", qui sont des propositions) ; recopie les nombres sans unité, avec un point décimal.'
  ].join('\n');
  var machines = [];
  try { machines = listMachines_(); } catch (e) { machines = []; }
  if (machines.length) {
    prompt += '\n\nLISTE DES MACHINES (modèle | marque | gamme | type d’installation | alimentation | liaisons) :\n' +
      machines.map(function (m) { return [m.modele, m.marque, m.gamme, m.type, m.phase, m.liaison].join(' | '); }).join('\n');
  }
  var text = callClaude_([block, { type: 'text', text: prompt }], 1800);
  var i = text.indexOf('{'), j = text.lastIndexOf('}');
  if (i < 0 || j < i) throw new Error('La feuille n’a pas pu être lue : réessaie avec une photo plus nette, bien à plat.');
  var raw;
  try { raw = JSON.parse(text.slice(i, j + 1)); } catch (e) { throw new Error('La feuille n’a pas pu être lue : réessaie avec une photo plus nette.'); }
  var keys = ['prenom', 'nom', 'adresse', 'cp', 'ville', 'projet', 'type', 'marque', 'modele', 'fluide', 'longueur', 'precharge', 'gpm', 'puissance', 'alimentation', 'ecs'];
  var out = {};
  keys.forEach(function (k) { var v = raw[k]; out[k] = v === null || v === undefined ? '' : String(v).trim().slice(0, 200); });
  if (out.type && TYPES_INSTALL.indexOf(out.type) < 0) out.type = '';
  if (out.fluide) out.fluide = out.fluide.toUpperCase().replace(/\s+/g, '').replace(/^R-?/, 'R');
  ['longueur', 'precharge', 'gpm'].forEach(function (k) { out[k] = out[k].replace('.', ','); });
  out.illisible = Array.isArray(raw.illisible) ? raw.illisible.map(String).slice(0, 15) : [];
  // Modèles proposés : ceux choisis par Claude dans la liste, sinon un classement simple
  var byKey = {};
  machines.forEach(function (x) { byKey[normModele_(x.modele)] = x; });
  var sugg = [];
  (Array.isArray(raw.suggestions) ? raw.suggestions : []).forEach(function (sg) {
    var x = sg && byKey[normModele_(sg.modele || sg)];
    if (x && sugg.length < 3 && !sugg.some(function (y) { return y.modele === x.modele; })) sugg.push({ modele: x.modele, marque: x.marque, raison: String(sg.raison || '').slice(0, 160) });
  });
  if (!sugg.length) sugg = scoreMachines_(machines, out).slice(0, 3);
  out.modeleLu = out.modele;
  // Complète avec la base des machines quand le modèle d'UE est reconnu
  var m = out.modele ? findMachine_(out.modele) : null;
  if (m) {
    out.modele = m.modele;
    if (!out.marque) out.marque = m.marque;
    if (!out.fluide) out.fluide = m.fluide;
    if (!out.precharge && m.precharge) out.precharge = m.precharge;
    if (!out.gpm && m.gpm) out.gpm = m.gpm;
    if (!out.type && m.type) out.type = m.type;
    out.machine = true;
    sugg = [{ modele: m.modele, marque: m.marque, raison: 'Référence lue sur la feuille' }].concat(sugg.filter(function (y) { return y.modele !== m.modele; })).slice(0, 3);
  }
  out.suggestions = sugg;
  return out;
}

/** Classement de secours : marque, type, alimentation, ECS et puissance la plus proche. */
function scoreMachines_(machines, info) {
  var marque = String(info.marque || '').toLowerCase(), kw = Number(String(info.puissance || '').replace(',', '.'));
  var tri = /tri/i.test(info.alimentation || ''), mono = /mono/i.test(info.alimentation || '');
  var scored = machines.map(function (m) {
    var sc = 0;
    if (marque && m.marque.toLowerCase().indexOf(marque) < 0 && marque.indexOf(m.marque.toLowerCase()) < 0) return null;
    if (info.type) { if (m.type === info.type) sc += 4; else return null; }
    if (tri && m.phase === 'Triphasé') sc += 2; if (tri && m.phase === 'Monophasé') sc -= 3;
    if (mono && m.phase === 'Monophasé') sc += 1; if (mono && m.phase === 'Triphasé') sc -= 3;
    var p = Number((String(m.modele).match(/(\d{2,3})/) || [])[1]) / 10;   // taille approximative d'après la référence (ex. RXM35 -> 3,5)
    if (kw && p) sc -= Math.abs(p - kw) / 2;
    return { m: m, sc: sc };
  }).filter(Boolean);
  if (!marque && !info.type && !kw) return [];
  scored.sort(function (a, b) { return b.sc - a.sc; });
  return scored.map(function (x) { return { modele: x.m.modele, marque: x.m.marque, raison: 'Proposé d’après la marque, le type et l’alimentation' }; });
}

/** À lancer une fois depuis l'éditeur (▶ Exécuter) : autorise les appels externes et vérifie la clé. */
function testerCleClaude() {
  var t = callClaude_([{ type: 'text', text: 'Réponds simplement : OK' }], 10);
  Logger.log('Clé Claude opérationnelle, réponse : ' + t);
}

/* ================================================================== Livraison 2 : mises en service */

var MES = {
  sheet: 'Mises en service',
  headers: ['ID', 'ID chantier', 'Date', 'Opérateur ID', 'Opérateur', 'N° série UE', 'N° série UI', 'Photos', 'Rapport (lien)', 'Rapport (ID)', 'Signature (ID)', 'Créée le', 'Modifiée le', 'Commentaire rapport'],
  mouvExtra: ['ID chantier', 'ID MES', 'Opérateur ID', 'CERFA (lien)', 'CERFA (ID)', 'Données'],
  trashMes: '_MES supprimées', trashMouv: '_Mouvements supprimés',
  ops: {
    mes: 'Démarrage sans manipulation',
    relache: 'Tirage au vide + ouverture des vannes',
    ajout: 'Ajout de charge',
    recup: 'Récupération de charge'
  },
  photoSlots: ['Plaque signalétique UE', 'Plaque signalétique UI', 'Étiquette fluide posée', 'Unité extérieure posée', 'Unité intérieure posée', 'Autre'],
  photoRequired: ['Plaque signalétique UE', 'Plaque signalétique UI', 'Étiquette fluide posée']
};

/** Crée un onglet s'il manque, sinon ajoute les colonnes qui manquent. */
function ensureSheet_(name, headers, hidden) {
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (hidden) sh.hideSheet();
    return sh;
  }
  headers.forEach(function (h) { ensureColumn_(name, h); });
  return sh;
}

function params_() {
  var p = {};
  try { readTable_(SH.params).rows.forEach(function (r) { p[String(r['Paramètre']).trim()] = r['Valeur']; }); } catch (e) { /* onglet absent */ }
  return p;
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return NaN;
  var n = Number(String(v).replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : NaN;
}
function fr_(n, d) { return isFinite(n) ? (Math.round(n * Math.pow(10, d)) / Math.pow(10, d)).toFixed(d).replace('.', ',') : ''; }
function parseFr_(s) {
  var m = String(s || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  var d = new Date(+m[3], +m[2] - 1, +m[1]);
  return d.getDate() === +m[1] && d.getMonth() === +m[2] - 1 ? d : null;   // refuse le 31/02, le 32/11…
}
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'); }

/** Droits de modification (fiche chantier, mise en service, rapport) : l'administrateur, ou l'opérateur qui a créé le chantier.
 *  Sur le chantier d'un autre, un opérateur peut consulter, ajouter des manipulations et compléter les n° de série absents. */
function canEditMes_(c, mesRow, user) {
  if (user.role === 'admin') return true;
  return !!(c && c.responsable && c.responsable === user.id);
}
var DROITS_MSG = 'Seul le responsable du chantier ou un administrateur peut faire cette modification.';

/* ---------- Outillage, bouteilles, distributeurs */

function outils_() {
  var t; try { t = readTable_(SH.outils); } catch (e) { return []; }
  var now = new Date();
  return t.rows.map(function (r) {
    var last = parseFr_(r['Dernier contrôle']), freq = num_(r['Fréquence (mois)']) || 12, next = null;
    if (last) { next = new Date(last.getTime()); next.setMonth(next.getMonth() + freq); }
    return {
      id: String(r['ID']), type: String(r['Type'] || ''), marque: String(r['Marque'] || ''), ref: String(r['Référence'] || ''),
      serie: String(r['N° de série'] || ''), dernierControle: last ? Utilities.formatDate(last, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      prochainControle: next ? Utilities.formatDate(next, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      ok: !!(next && next >= now),
      label: [r['Marque'], r['Référence']].filter(Boolean).join(' ') + (r['N° de série'] ? ' n° ' + r['N° de série'] : '')
    };
  });
}

/** Date de la dernière opération enregistrée sur chaque bouteille : { ref en minuscules: Date } */
function boutDernieresOps_() {
  var out = {};
  try {
    readTable_(SH.mouv).rows.forEach(function (r) {
      var k = String(r['Bouteille'] || '').trim().toLowerCase(), dt = parseFr_(r['Date']);
      if (k && dt && (!out[k] || dt > out[k])) out[k] = dt;
    });
  } catch (e) {}
  return out;
}

function bouteilles_() {
  var t; try { t = readTable_(SH.bout); } catch (e) { return []; }
  var last = boutDernieresOps_(), fmt = function (d) { return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy') : ''; };
  return t.rows.filter(function (r) { return String(r['Réf bouteille'] || '').trim(); }).map(function (r) {
    var vide = num_(r['Poids à vide (kg)']), pesee = num_(r['Dernière pesée (kg)']);
    return {
      derniereOp: fmt(last[String(r['Réf bouteille']).trim().toLowerCase()]), dateEntree: String(r['Date d’entrée'] || ''),
      ref: String(r['Réf bouteille']).trim(), type: String(r['Type'] || ''), fluide: normFluide_(r['Fluide']),
      vide: isFinite(vide) ? fr_(vide, 2) : '', pesee: isFinite(pesee) ? fr_(pesee, 2) : '',
      reste: isFinite(vide) && isFinite(pesee) ? fr_(pesee - vide, 2) : '',
      statut: String(r['Statut'] || ''), distributeur: String(r['Distributeur (ID)'] || '')
    };
  });
}

function distributeurs_() {
  var t; try { t = readTable_(SH.dist); } catch (e) { return []; }
  return t.rows.filter(function (r) { return String(r['Nom'] || '').trim(); }).map(function (r) {
    return { id: String(r['ID']), nom: String(r['Nom']), siret: String(r['SIRET'] || '').replace(/^'/, ''), adresse: String(r['Adresse du site de reprise'] || '') };
  });
}

function fluideInfo_(code) {
  var k = normFluide_(code), f = null;
  try { f = listFluides_().filter(function (x) { return normFluide_(x.code) === k; })[0] || null; } catch (e) { f = null; }
  return f;
}

/* ---------- Lecture du registre */

function mesRows_(chantierId) {
  ensureSheet_(MES.sheet, MES.headers);
  return readTable_(MES.sheet).rows.filter(function (r) { return String(r['ID chantier']) === chantierId; });
}

function mouvTable_() {
  MES.mouvExtra.forEach(function (h) { ensureColumn_(SH.mouv, h); });
  return readTable_(SH.mouv);
}

function manipPublic_(r) {
  var data = {};
  try { data = JSON.parse(r['Données'] || '{}'); } catch (e) { data = {}; }
  return {
    cerfa: String(r['N° CERFA'] || ''), op: data.op || '', label: MES.ops[data.op] || String(r['Opération'] || ''), date: String(r['Date'] || ''),
    bouteille: String(r['Bouteille'] || ''), quantite: data.quantite || '', chargeTotale: r['Charge totale machine (kg)'] === '' ? '' : fr_(num_(r['Charge totale machine (kg)']), 2),
    lien: String(r['CERFA (lien)'] || ''), operateur: String(r['Opérateur'] || ''), papier: !!data.papier
  };
}

function mesPublic_(row, c, user) {
  if (!row) return null;
  var manips = mouvTable_().rows.filter(function (r) { return String(r['ID MES']) === String(row['ID']); }).map(manipPublic_);
  var photos = []; try { photos = JSON.parse(row['Photos'] || '[]'); } catch (e) { photos = []; }
  return {
    id: String(row['ID']), date: String(row['Date'] || ''), operateur: String(row['Opérateur'] || ''), operateurId: String(row['Opérateur ID'] || ''),
    serieUE: String(row['N° série UE'] || ''), serieUI: String(row['N° série UI'] || ''), photos: photos.length,
    photoList: photos.map(function (ph) { return { slot: ph.slot, id: ph.id, url: 'https://drive.google.com/file/d/' + ph.id + '/view' }; }),
    commentaire: String(row['Commentaire rapport'] || ''),
    rapport: String(row['Rapport (lien)'] || ''), manips: manips, peutModifier: canEditMes_(c, row, user)
  };
}

/** Tout ce que l'écran de mise en service doit connaître pour un chantier. */
function mesContext_(d, user) {
  var c = getChantier_(d.chantierId);
  var row = mesRows_(c.id)[0] || null;
  var outils = outils_();
  return {
    chantier: c, machine: c.modele ? findMachine_(c.modele) : null, fluide: fluideInfo_(c.fluide),
    detecteurs: outils.filter(function (o) { return /tecteur/i.test(o.type); }),
    pompe: (outils.filter(function (o) { return /pompe/i.test(o.type); })[0] || null),
    bouteilles: bouteilles_(), distributeurs: distributeurs_(),
    mes: mesPublic_(row, c, user), operateur: { nom: (user.prenom + ' ' + user.nom).trim(), attestation: user.attestation },
    ops: MES.ops, photoSlots: MES.photoSlots, photoRequired: MES.photoRequired
  };
}

/* ---------- Google Drive */

function folderPath_(names) {
  var f = DriveApp.getRootFolder();
  names.forEach(function (n) {
    var it = f.getFoldersByName(n);
    f = it.hasNext() ? it.next() : f.createFolder(n);
  });
  return f;
}

function chantierFolder_(c) {
  var name = c.id + ' – ' + (c.client || c.nom || 'Client');
  return folderPath_(['ARE Fluides', 'Chantiers', name]);
}

function saveFile_(folder, b64, mime, name) {
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, name);
  var file = folder.createFile(blob);
  return { id: file.getId(), url: file.getUrl(), name: name };
}

function trashFile_(id) { if (!id) return; try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* déjà supprimé */ } }

/** À lancer une fois depuis l'éditeur (▶ Exécuter) : autorise l'accès au Drive et crée le dossier de l'appli. */
function autoriserDrive() {
  var f = folderPath_(['ARE Fluides', 'Chantiers']);
  Logger.log('Drive opérationnel : dossier « ARE Fluides › Chantiers » prêt (' + f.getUrl() + ').');
}

/** Envoi d'une photo (JPEG allégé côté téléphone) dans le dossier du chantier. */
function uploadPhoto_(d, user) {
  var c = getChantier_(d.chantierId);
  if (MES.photoSlots.indexOf(d.slot) < 0) throw new Error('Type de photo inconnu.');
  if (!d.data || String(d.data).length > 12 * 1024 * 1024) throw new Error('Photo manquante ou trop lourde.');
  var folder = chantierFolder_(c);
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss');
  var f = saveFile_(folder, d.data, 'image/jpeg', stamp + ' ' + d.slot + '.jpg');
  return { id: f.id, url: f.url, slot: d.slot };
}

/** Modèle vierge du CERFA 15497*04, gardé dans le Drive (copié une fois depuis le dépôt de l'appli). */
var CERFA_SOURCE = 'https://raw.githubusercontent.com/tbaldearchenergie-lab/are-fluides/main/docs/cerfa_15497-04.pdf';
function cerfaTemplate_() {
  var folder = folderPath_(['ARE Fluides', 'Modèles']);
  var it = folder.getFilesByName('cerfa_15497-04.pdf');
  var file;
  if (it.hasNext()) file = it.next();
  else {
    var res = UrlFetchApp.fetch(CERFA_SOURCE, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('Modèle CERFA introuvable (erreur ' + res.getResponseCode() + ').');
    file = folder.createFile(res.getBlob().setName('cerfa_15497-04.pdf'));
  }
  return Utilities.base64Encode(file.getBlob().getBytes());
}

/* ---------- Numéro de fiche CERFA : F + année + mois + séquence (jamais réutilisée) */

function nextCerfa_(dateFr) {
  var max = 0;
  [SH.mouv, MES.trashMouv].forEach(function (n) {
    if (!ss_().getSheetByName(n)) return;
    readTable_(n).rows.forEach(function (r) {
      var m = String(r['N° CERFA'] || '').match(/^F\d{6}(\d{4})$/);
      if (m) max = Math.max(max, +m[1]);
    });
  });
  var d = parseFr_(dateFr) || new Date();
  return 'F' + Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyyMM') + ('0000' + (max + 1)).slice(-4);
}

/* ---------- Champs du CERFA, construits ici pour pouvoir les refaire à l'identique */

function cerfaFields_(m) {
  var p = params_();
  var t = {}, checks = [];
  var NAT = { assemblage: 'Case_Assemblage', mes: 'Case_MiseService', modif: 'Case_Modif', maintenance: 'Case_Maintenance', ctrlPerio: 'Case_CtrlPerio', ctrlNonPerio: 'Case_CtrlNonPerio', demantel: 'Case_Demantel', autre: 'Case_Autre' };
  var nat = Array.isArray(m.nature) && m.nature.length ? m.nature : ['mes'];
  nat.forEach(function (k) { if (NAT[k] && checks.indexOf(NAT[k]) < 0) checks.push(NAT[k]); });
  if (nat.indexOf('autre') >= 0) t['Autre'] = String(m.natureAutre || '').slice(0, 60);
  t['Fiche_no'] = m.cerfa;
  t['Operateur'] = [p['Raison sociale'], p['Adresse'], p['SIRET'] ? 'SIRET : ' + p['SIRET'] : ''].filter(Boolean).join('\n');
  t['Attestation_no'] = String(p['N° d’attestation de capacité'] || '');
  t['Detenteur'] = [m.client, m.adresse].filter(Boolean).join('\n');
  t['Equipement_ID'] = [m.machine, m.serieUE ? 'n° série UE ' + m.serieUE : ''].filter(Boolean).join(' – ');
  t['Equipement_Fluide'] = String(m.fluide || '').replace(/^R-?/i, '');
  t['Equipement_Charge'] = fr_(m.chargeTotale, 2);
  t['Equipement_teqCO2'] = isFinite(m.gwp) ? fr_(m.chargeTotale * m.gwp / 1000, 3) : '';
  if (m.detecteur) {
    t['Detecteur_ID'] = m.detecteur.label;
    var dc = parseFr_(m.detecteur.dernierControle);
    if (dc) { t['Controle_Jour'] = ('0' + dc.getDate()).slice(-2); t['Controle_Mois'] = ('0' + (dc.getMonth() + 1)).slice(-2); t['Controle_Annee'] = String(dc.getFullYear()); }
  }
  if (m.fuite === 'oui') {
    checks.push('Case_Fuite_Oui');
    (m.fuites || []).slice(0, 3).forEach(function (f, i) {
      if (!f || !String(f.loc || '').trim()) return;
      t['Fuite_Loca_' + (i + 1)] = String(f.loc).trim();
      checks.push(f.rep === 'afaire' ? 'Case_Rep_Fuite' + (i + 1) + '_AFaire' : 'Case_Rep_Fuite' + (i + 1) + '_realisee');
    });
  } else checks.push('Case_Fuite_Non');
  var q = fr_(m.quantite, 2);
  if (m.op === 'ajout') { t['11_Quantite'] = q; t['11_QA'] = q; t['11_QB'] = '0'; t['11_QC'] = '0'; t['11_QDE'] = '0'; t['11_Contenant_ID'] = m.bouteille || ''; }
  else if (m.op === 'recup') {
    t['11_Quantite'] = '0'; t['11_QDE'] = q; t['11_QD'] = q; t['11_QE'] = '0'; t['11_Contenant_ID'] = m.bouteille || '';
    checks.push(m.inflammable ? 'Case_12_UN3161' : 'Case_12_UN1078');
    if (m.distributeur) t['13_Instal'] = [m.distributeur.nom, m.distributeur.siret ? 'SIRET ' + m.distributeur.siret : '', m.distributeur.adresse].filter(Boolean).join(' – ');
  } else { t['11_Quantite'] = '0'; t['11_QDE'] = '0'; }
  t['14_Observations'] = m.observationsCerfa || '';
  t['Sign_Operateur_Nom'] = m.operateur + (m.attestation ? ' – attestation n° ' + m.attestation : '');
  t['Sign_Operateur_Qualite'] = 'Opérateur fluides frigorigènes';
  t['Sign_Operateur_Date'] = m.date;
  return { text: t, checks: checks, radio: { 'Bouton_Oui': '2' } };
}

function observations_(m) {
  var s = [];
  if (m.op === 'relache') {
    var rem = { stable: 'remontée stable sous 1 000 microns', humidite: 'remontée puis stabilisation', fuite: 'remontée sans arrêt' }[m.remontee];
    s.push('Mise en service : tirage au vide' + (m.vide ? ' (' + m.vide + ' microns' + (m.tenue ? ' en ' + m.tenue + ' min' : '') + (rem ? ', ' + rem : '') + ')' : '') + ' et ouverture des vannes, charge usine ' + fr_(m.chargeUsine, 2) + ' kg.');
  } else if (m.op === 'ajout') {
    s.push('Mise en service : ajout de charge ' + fr_(m.quantite, 2) + ' kg (bouteille ' + m.bouteille + ').');
  } else if (m.op === 'recup') {
    s.push('Récupération de ' + fr_(m.quantite, 2) + ' kg de fluide (bouteille ' + m.bouteille + ').');
  }
  if (m.longueur) s.push('Liaisons ' + m.longueur + ' m' + (m.precharge ? ', précharge ' + m.precharge + ' m' : '') + (m.gpm ? ', ' + m.gpm + ' g/m' : '') + '.');
  if (m.observations) s.push(String(m.observations).trim());
  return s.join(' ');
}

/* ---------- Enregistrement d'une opération */

function saveManip_(d, user) {
  if (!MES.ops[d.op]) throw new Error('Choisis l’opération réalisée.');
  var date = String(d.date || '').trim();
  if (!parseFr_(date)) throw new Error('Date de l’intervention au format jj/mm/aaaa.');
  if (parseFr_(date) > new Date()) throw new Error('La date de l’intervention ne peut pas être dans le futur.');
  var c = getChantier_(d.chantierId);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var mesSh = ensureSheet_(MES.sheet, MES.headers);
    var mt = readTable_(MES.sheet);
    var row = mt.rows.filter(function (r) { return String(r['ID chantier']) === c.id; })[0] || null;
    var editeur = canEditMes_(c, row, user);
    if (row && d.op === 'mes') throw new Error('La mise en service est déjà enregistrée : choisis une manipulation de fluide.');

    var fl = fluideInfo_(c.fluide), gwp = fl ? fl.gwp : num_(c.gwp);
    var machine = c.modele ? findMachine_(c.modele) : null;
    var chargeUsine = num_(d.chargeUsine);
    if (!isFinite(chargeUsine)) chargeUsine = num_(c.chargeUsine);
    if (!isFinite(chargeUsine) && machine) chargeUsine = num_(machine.charge);
    // Quantité transvasée : pesées avant / après (bouteille avec ses accessoires) ou saisie directe.
    // Le stock se suit ensuite avec la pesée de la bouteille seule (sans accessoires).
    var avant = NaN, apres = NaN, quantite = 0, pesee = null;
    if (d.op === 'ajout' || d.op === 'recup') {
      if (!d.bouteille) throw new Error('Choisis la bouteille utilisée.');
      if (d.mode === 'direct') {
        quantite = num_(d.quantite);
        if (!(quantite > 0)) throw new Error('Indique la quantité ' + (d.op === 'ajout' ? 'chargée' : 'récupérée') + ' (kg).');
      } else {
        var acA = num_(d.poidsAvant), acB = num_(d.poidsApres);
        if (!isFinite(acA) || !isFinite(acB)) throw new Error('Saisis les deux pesées de la bouteille.');
        quantite = d.op === 'ajout' ? acA - acB : acB - acA;
        if (!(quantite > 0)) throw new Error(d.op === 'ajout' ? 'La bouteille doit être plus légère après la charge.' : 'La bouteille doit être plus lourde après la récupération.');
      }
      quantite = Math.round(quantite * 1000) / 1000;
      var bref = bouteilles_().filter(function (b) { return b.ref === d.bouteille; })[0];
      checkDateBouteille_(bref, date);
      avant = bref ? num_(bref.pesee) : NaN;   // dernière pesée de la bouteille seule
      var attendu = isFinite(avant) ? avant + (d.op === 'ajout' ? -quantite : quantite) : NaN;
      apres = num_(d.poidsSeule);
      if (!isFinite(apres)) apres = isFinite(attendu) ? attendu : (d.mode === 'direct' ? NaN : num_(d.poidsApres));
      pesee = { mode: d.mode === 'direct' ? 'direct' : 'pesee', avecAccessoires: d.mode === 'direct' ? null : { avant: num_(d.poidsAvant), apres: num_(d.poidsApres) },
        seule: isFinite(apres) ? apres : null, attendu: isFinite(attendu) ? Math.round(attendu * 1000) / 1000 : null,
        ecart: isFinite(attendu) && isFinite(apres) ? Math.round((apres - attendu) * 1000) / 1000 : null };
    }
    var detecteur = null;
    if (d.op !== 'mes') {
      detecteur = outils_().filter(function (o) { return o.id === d.detecteur; })[0] || null;
      if (!detecteur) throw new Error('Choisis le détecteur de fuite utilisé.');
      if (!detecteur.ok) throw new Error('Le contrôle de ce détecteur est dépassé : il ne peut pas être utilisé.');
    }
    // Case 13 : site de reprise du distributeur rattaché à la bouteille (renseigné par l'administrateur, pas par l'opérateur)
    var distributeur = null;
    if (d.op === 'recup') {
      var bo = bouteilles_().filter(function (b) { return b.ref === d.bouteille; })[0];
      distributeur = bo && bo.distributeur ? (distributeurs_().filter(function (x) { return x.id === bo.distributeur; })[0] || null) : null;
    }

    var now = new Date(), opNom = (user.prenom + ' ' + user.nom).trim() || user.email;
    // En-tête de la mise en service (créé à la première opération)
    if (!row) {
      var rec = { 'ID': nextId_(mt.rows.concat(ss_().getSheetByName(MES.trashMes) ? readTable_(MES.trashMes).rows : []), 'M', 'ID'), 'ID chantier': c.id, 'Date': date,
        'Opérateur ID': user.id, 'Opérateur': opNom, 'N° série UE': d.serieUE || '', 'N° série UI': d.serieUI || '', 'Photos': '[]', 'Créée le': now,
        'Commentaire rapport': String(d.observations || '').trim().slice(0, 3000) };   // commentaire de l'étape Commentaire, repris dans le rapport
      writeRow_(mt, rec, null);
      mt = readTable_(MES.sheet);
      row = mt.rows.filter(function (r) { return String(r['ID chantier']) === c.id; })[0];
    } else {
      var up = { 'Modifiée le': now };
      if (d.serieUE && (editeur || !String(row['N° série UE'] || '').trim())) up['N° série UE'] = d.serieUE;
      if (d.serieUI && (editeur || !String(row['N° série UI'] || '').trim())) up['N° série UI'] = d.serieUI;
      writeRow_(mt, up, row._row);
    }
    // Photos et signature
    var photos = []; try { photos = JSON.parse(row['Photos'] || '[]'); } catch (e) { photos = []; }
    (d.photos || []).forEach(function (p) { if (p && p.id) photos.push({ slot: p.slot, id: p.id }); });
    var upd = { 'Photos': JSON.stringify(photos) };
    if (d.signature) {
      var sig = saveFile_(chantierFolder_(c), String(d.signature).replace(/^data:image\/png;base64,/, ''), 'image/png', 'Signature ' + row['ID'] + ' ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMdd-HHmmss') + '.png');
      upd['Signature (ID)'] = sig.id;
    }
    writeRow_(readTable_(MES.sheet), upd, row._row);
    // Fiche chantier : date de mise en service et numéros de série
    var ct = readTable_(SH.chantiers), crow = ct.rows.filter(function (r) { return r['ID'] === c.id; })[0];
    var cup = {};
    if (!c.dateMes) cup['Date MES'] = date;
    if (d.serieUE && (editeur || !c.serieUE)) cup['N° série UE'] = d.serieUE;
    if (d.serieUI && (editeur || !c.serieUI)) cup['N° série UI'] = d.serieUI;
    if (isFinite(chargeUsine) && !c.chargeUsine) cup['Charge usine (kg)'] = chargeUsine;
    if (Object.keys(cup).length) writeRow_(ct, cup, crow._row);

    var out = { mesId: String(row['ID']), cerfa: '' };
    if (d.op === 'mes') return out;

    // Charge totale de la machine après l'opération
    var mv = mouvTable_();
    var net = 0;
    mv.rows.filter(function (r) { return String(r['ID MES']) === String(row['ID']); }).forEach(function (r) {
      var x = {}; try { x = JSON.parse(r['Données'] || '{}'); } catch (e) { x = {}; }
      if (x.op === 'ajout') net += num_(x.quantite) || 0;
      if (x.op === 'recup') net -= num_(x.quantite) || 0;
    });
    if (d.op === 'ajout') net += quantite;
    if (d.op === 'recup') net -= quantite;
    var chargeTotale = (isFinite(chargeUsine) ? chargeUsine : 0) + net;

    var cerfa = nextCerfa_(date);
    var m = {
      op: d.op, cerfa: cerfa, date: date, client: c.client, adresse: [c.adresse, [c.cp, c.ville].filter(Boolean).join(' ')].filter(Boolean).join(', '),
      machine: [c.marque, c.modele].filter(Boolean).join(' '), serieUE: d.serieUE || c.serieUE || '', fluide: c.fluide, gwp: gwp,
      inflammable: fl ? fl.inflammable : String(c.inflam).toLowerCase() === 'oui', chargeUsine: chargeUsine, chargeTotale: chargeTotale,
      quantite: quantite, bouteille: d.bouteille || '', distributeur: distributeur, detecteur: detecteur,
      fuite: d.fuite === 'oui' ? 'oui' : 'non', fuites: d.fuites || [], vide: d.vide || '', tenue: d.tenue || '', remontee: d.remontee || '',
      nature: Array.isArray(d.nature) ? d.nature.filter(function (k) { return /^(assemblage|mes|modif|maintenance|ctrlPerio|ctrlNonPerio|demantel|autre)$/.test(k); }) : ['mes'], natureAutre: String(d.natureAutre || '').trim(),
      longueur: c.longueur, precharge: c.precharge, gpm: c.gpm, observations: d.observations || '', pesee: pesee,
      operateur: opNom, attestation: user.attestation || ''
    };
    m.observationsCerfa = observations_(m);
    writeRow_(mv, {
      'Date': date, 'Opération': d.op === 'relache' ? 'Mise en service' : MES.ops[d.op], 'Cadre': 'Mise en service',
      'Chantier / Client': c.client, 'Adresse': m.adresse, 'Machine': m.machine, 'Fluide': c.fluide, 'Bouteille': m.bouteille,
      'Poids avant (kg)': isFinite(avant) ? avant : '', 'Poids après (kg)': isFinite(apres) ? apres : '',
      'Variation bouteille (kg)': quantite ? (d.op === 'ajout' ? -quantite : quantite) : 0,
      'Charge totale machine (kg)': Math.round(chargeTotale * 1000) / 1000,
      't éq. CO2': isFinite(gwp) ? Math.round(chargeTotale * gwp) / 1000 : '',
      'Opérateur': opNom, 'N° CERFA': cerfa, 'Observations': m.observationsCerfa,
      'ID chantier': c.id, 'ID MES': String(row['ID']), 'Opérateur ID': user.id, 'Données': JSON.stringify(m)
    }, null);
    // Stock : nouvelle pesée de la bouteille
    if (m.bouteille && isFinite(apres)) {
      var bt = readTable_(SH.bout), br = bt.rows.filter(function (r) { return String(r['Réf bouteille']).trim() === m.bouteille; })[0];
      if (br) {
        var bup = { 'Dernière pesée (kg)': apres }, bv = num_(br['Poids à vide (kg)']);
        if (isFinite(bv)) bup['Fluide restant (kg)'] = Math.round(Math.max(0, apres - bv) * 1000) / 1000;
        if (!String(br['Fluide'] || '').trim() && c.fluide) bup['Fluide'] = c.fluide;   // bouteille de récupération vide : prend le fluide récupéré
        writeRow_(bt, bup, br._row);
      }
    }
    out.cerfa = cerfa;
    if (d.regulId) marquerRegularisation_(d.regulId, cerfa);
    out.fields = cerfaFields_(m);
    out.fileName = cerfaFileName_(m);
    return out;
  } finally { lock.releaseLock(); }
}

function aammjj_(dateFr) { var d = parseFr_(dateFr) || new Date(); return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyMMdd'); }
function cerfaFileName_(m) { return aammjj_(m.date) + ' CERFA ' + String(m.client || '').replace(/[\\\/:*?"<>|]/g, ' ').trim() + ' ' + m.cerfa + '.pdf'; }

function findMouv_(cerfa) {
  var mv = mouvTable_();
  var r = mv.rows.filter(function (x) { return String(x['N° CERFA']) === String(cerfa); })[0];
  if (!r) throw new Error('Fiche ' + cerfa + ' introuvable dans le registre.');
  return { t: mv, row: r };
}

/** Reçoit le CERFA rempli et signé (PDF fait sur le téléphone) et le range dans le Drive du chantier. */
function saveCerfa_(d, user) {
  var f = findMouv_(d.cerfa);
  var c = getChantier_(String(f.row['ID chantier']));
  var mesRow = mesRows_(c.id)[0];
  // Tout opérateur peut régénérer un CERFA : son contenu vient des données enregistrées (cerfaData), pas de la saisie
  if (!d.pdf) throw new Error('PDF manquant.');
  var m = {}; try { m = JSON.parse(f.row['Données'] || '{}'); } catch (e) { m = {}; }
  trashFile_(String(f.row['CERFA (ID)'] || ''));
  var file = saveFile_(chantierFolder_(c), d.pdf, 'application/pdf', cerfaFileName_(m));
  writeRow_(f.t, { 'CERFA (lien)': file.url, 'CERFA (ID)': file.id }, f.row._row);
  return file;
}

/** Données pour refaire un CERFA à l'identique (même numéro, signature enregistrée). */
function cerfaData_(d, user) {
  var f = findMouv_(d.cerfa);
  var c = getChantier_(String(f.row['ID chantier']));
  var mesRow = mesRows_(c.id)[0];
  var m = {}; try { m = JSON.parse(f.row['Données'] || '{}'); } catch (e) { m = {}; }
  var sig = '';
  if (mesRow && mesRow['Signature (ID)']) { try { sig = Utilities.base64Encode(DriveApp.getFileById(String(mesRow['Signature (ID)'])).getBlob().getBytes()); } catch (e) { sig = ''; } }
  return { cerfa: m.cerfa, fields: cerfaFields_(m), signature: sig, fileName: cerfaFileName_(m) };
}

/* ---------- Rapport d'installation (HTML converti en PDF par Google) */

function buildReport_(d, user) {
  var c = getChantier_(d.chantierId);
  var row = mesRows_(c.id)[0];
  if (!row) throw new Error('Aucune mise en service sur ce chantier.');
  // Le rapport est construit à partir des données enregistrées. Un opérateur qui n'a pas créé le chantier
  // (ajout de n° de série absent) peut le refaire s'il existe, mais pas le créer.
  if (!canEditMes_(c, row, user) && !String(row['Rapport (ID)'] || row['Rapport (lien)'] || '')) return { skipped: true };
  var p = params_(), tz = Session.getScriptTimeZone();
  var manips = mouvTable_().rows.filter(function (r) { return String(r['ID MES']) === String(row['ID']); });
  var photos = []; try { photos = JSON.parse(row['Photos'] || '[]'); } catch (e) { photos = []; }
  var e = function (s) { return String(s === undefined || s === null ? '' : s).replace(/[&<>"]/g, function (x) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[x]; }); };
  var rowsHtml = function (pairs) { return pairs.filter(function (x) { return x[1] !== '' && x[1] !== undefined; }).map(function (x) { return '<tr><th>' + e(x[0]) + '</th><td>' + e(x[1]) + '</td></tr>'; }).join(''); };
  var html = '<html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;font-size:11px;color:#14201A;margin:28px}' +
    'h1{font-size:20px;margin:0}h2{font-size:13px;margin:18px 0 6px;color:#1F6B45;text-transform:uppercase;letter-spacing:.04em}' +
    'table{border-collapse:collapse;width:100%}th{text-align:left;color:#4A5A51;font-weight:normal;width:38%;padding:4px 6px;border-bottom:1px solid #E6EAE6;vertical-align:top}' +
    'td{padding:4px 6px;border-bottom:1px solid #E6EAE6;font-weight:bold}.ph{display:inline-block;width:46%;margin:0 2% 12px 0;vertical-align:top}' +
    '.ph img{width:100%;border:1px solid #C9D2CC}.ph div{font-size:10px;color:#4A5A51;margin-top:2px}.head{border-bottom:3px solid #1F6B45;padding-bottom:8px;margin-bottom:6px}</style></head><body>' +
    '<div class="head"><h1>Rapport d’installation</h1><div>' + e(p['Raison sociale'] || 'Archenergie Renouvelable') + ' · ' + e(p['Adresse'] || '') + '</div></div>' +
    '<h2>Chantier</h2><table>' + rowsHtml([['Client', c.client], ['Adresse', [c.adresse, [c.cp, c.ville].filter(Boolean).join(' ')].filter(Boolean).join(', ')], ['N° projet', c.projet], ['Date de mise en service', row['Date']], ['Opérateur', row['Opérateur']]]) + '</table>' +
    '<h2>Installation</h2><table>' + rowsHtml([['Type d’installation', c.type], ['Unité extérieure', [c.marque, c.modele].filter(Boolean).join(' ')], ['N° de série UE', row['N° série UE'] || c.serieUE], ['N° de série UI', row['N° série UI'] || c.serieUI],
      ['Fluide', c.fluide + (c.gwp ? ' (GWP ' + c.gwp + ')' : '')], ['Charge usine', c.chargeUsine ? c.chargeUsine + ' kg' : ''], ['Liaisons frigorifiques', c.longueur ? c.longueur + ' m' + (c.precharge ? ' · précharge ' + c.precharge + ' m' : '') + (c.gpm ? ' · ' + c.gpm + ' g/m' : '') : '']]) + '</table>' +
    '<h2>Opérations réalisées</h2><table>' + (manips.length ? manips.map(function (r) {
      return '<tr><th>' + e(r['Date']) + ' · ' + e(r['Opération']) + '</th><td>' + e(r['Observations']) + (r['N° CERFA'] ? '<br><span style="font-weight:normal">CERFA ' + e(r['N° CERFA']) + '</span>' : '') + '</td></tr>';
    }).join('') : '<tr><th>' + e(row['Date']) + '</th><td>Démarrage sans manipulation de fluide.</td></tr>') + '</table>';
  if (String(row['Commentaire rapport'] || '').trim()) html += '<h2>Commentaires</h2><p style="white-space:pre-wrap">' + e(row['Commentaire rapport']) + '</p>';
  if (photos.length) {
    html += '<h2>Photos</h2>';
    photos.slice(0, 12).forEach(function (ph) {
      try {
        var b = DriveApp.getFileById(ph.id).getBlob();
        html += '<div class="ph"><img src="data:image/jpeg;base64,' + Utilities.base64Encode(b.getBytes()) + '"><div>' + e(ph.slot) + '</div></div>';
      } catch (x) { /* photo supprimée */ }
    });
  }
  html += '<p style="margin-top:20px;color:#4A5A51">Document établi le ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy') + ' par l’application ARE Fluides.</p></body></html>';
  var name = aammjj_(String(row['Date'])) + ' Rapport installation ' + String(c.client || '').replace(/[\\\/:*?"<>|]/g, ' ').trim() + '.pdf';
  var blob = HtmlService.createHtmlOutput(html).getAs('application/pdf').setName(name);
  trashFile_(String(row['Rapport (ID)'] || ''));
  var file = chantierFolder_(c).createFile(blob);
  writeRow_(readTable_(MES.sheet), { 'Rapport (lien)': file.getUrl(), 'Rapport (ID)': file.getId() }, row._row);
  return { url: file.getUrl(), name: name };
}

/* ---------- Rapport d'installation : modification et suppression */

/** Modifie le rapport (commentaire, photos retirées ou ajoutées) puis le refait. */
function updateReport_(d, user) {
  var c = getChantier_(d.chantierId);
  var row = mesRows_(c.id)[0];
  if (!row) throw new Error('Aucune mise en service sur ce chantier.');
  if (!canEditMes_(c, row, user)) throw new Error(DROITS_MSG);
  var photos = []; try { photos = JSON.parse(row['Photos'] || '[]'); } catch (e) { photos = []; }
  var remove = d.removePhotos || [];
  photos = photos.filter(function (ph) { if (remove.indexOf(ph.id) >= 0) { trashFile_(ph.id); return false; } return true; });
  (d.addPhotos || []).forEach(function (ph) { if (ph && ph.id && MES.photoSlots.indexOf(ph.slot) >= 0) photos.push({ slot: ph.slot, id: ph.id }); });
  writeRow_(readTable_(MES.sheet), { 'Photos': JSON.stringify(photos), 'Commentaire rapport': String(d.commentaire || '').slice(0, 3000), 'Modifiée le': new Date() }, row._row);
  return buildReport_({ chantierId: c.id }, user);
}

/** Supprime le PDF du rapport (la mise en service et les CERFA restent). */
function deleteReport_(d, user) {
  var c = getChantier_(d.chantierId);
  var row = mesRows_(c.id)[0];
  if (!row) throw new Error('Aucune mise en service sur ce chantier.');
  if (!canEditMes_(c, row, user)) throw new Error(DROITS_MSG);
  trashFile_(String(row['Rapport (ID)'] || ''));
  writeRow_(readTable_(MES.sheet), { 'Rapport (lien)': '', 'Rapport (ID)': '', 'Modifiée le': new Date() }, row._row);
  return true;
}

/* ---------- Modification et suppression */

function updateMes_(d, user) {
  var c = getChantier_(d.chantierId);
  var row = mesRows_(c.id)[0];
  if (!row) throw new Error('Aucune mise en service sur ce chantier.');
  if (!canEditMes_(c, row, user)) {
    // Compléter seulement les n° de série absents
    var ue0 = String(row['N° série UE'] || '').trim(), ui0 = String(row['N° série UI'] || '').trim();
    var ue = String(d.serieUE || '').trim(), ui = String(d.serieUI || '').trim();
    if ((ue0 && ue !== ue0) || (ui0 && ui !== ui0) || (d.date && String(d.date).trim() !== String(row['Date']))) throw new Error(DROITS_MSG + ' Tu peux seulement compléter les n° de série absents.');
    if (ue === ue0 && ui === ui0) return { cerfas: [] };
    d = { chantierId: c.id, date: String(row['Date']), serieUE: ue, serieUI: ui };
  }
  var date = String(d.date || '').trim();
  if (!parseFr_(date)) throw new Error('Date au format jj/mm/aaaa.');
  if (parseFr_(date) > new Date()) throw new Error('La date ne peut pas être dans le futur.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    writeRow_(readTable_(MES.sheet), { 'Date': date, 'N° série UE': d.serieUE || '', 'N° série UI': d.serieUI || '', 'Modifiée le': new Date() }, row._row);
    var ct = readTable_(SH.chantiers), crow = ct.rows.filter(function (r) { return r['ID'] === c.id; })[0];
    writeRow_(ct, { 'Date MES': date, 'N° série UE': d.serieUE || '', 'N° série UI': d.serieUI || '' }, crow._row);
    // Les opérations gardent leur numéro de fiche : la date et le n° de série sont repris dans leurs données
    var mv = mouvTable_(), cerfas = [];
    mv.rows.filter(function (r) { return String(r['ID MES']) === String(row['ID']); }).forEach(function (r) {
      var m = {}; try { m = JSON.parse(r['Données'] || '{}'); } catch (e) { m = {}; }
      m.date = date; m.serieUE = d.serieUE || '';
      writeRow_(mv, { 'Date': date, 'Données': JSON.stringify(m) }, r._row);
      cerfas.push(String(r['N° CERFA']));
    });
    return { cerfas: cerfas };
  } finally { lock.releaseLock(); }
}

function archiveRow_(trashName, t, r, user) {
  var src = t.headers.concat(['Supprimé le', 'Supprimé par']);
  ensureSheet_(trashName, src, true);
  var tt = readTable_(trashName), rec = {};
  t.headers.forEach(function (h) { if (h) rec[h] = r[h]; });
  rec['Supprimé le'] = new Date(); rec['Supprimé par'] = (user.prenom + ' ' + user.nom).trim() || user.email;
  writeRow_(tt, rec, null);
}

function deleteManipRow_(mv, r, user) {
  var m = {}; try { m = JSON.parse(r['Données'] || '{}'); } catch (e) { m = {}; }
  // Stock : on remet la pesée d'avant si la bouteille n'a pas servi depuis
  if (m.bouteille) {
    var bt = readTable_(SH.bout), br = bt.rows.filter(function (x) { return String(x['Réf bouteille']).trim() === m.bouteille; })[0];
    if (br && num_(br['Dernière pesée (kg)']) === num_(r['Poids après (kg)']) && isFinite(num_(r['Poids avant (kg)']))) {
      var pv = num_(r['Poids avant (kg)']), tare = num_(br['Poids à vide (kg)']), back = { 'Dernière pesée (kg)': pv };
      if (isFinite(tare)) back['Fluide restant (kg)'] = Math.round(Math.max(0, pv - tare) * 1000) / 1000;
      writeRow_(bt, back, br._row);
    }
  }
  trashFile_(String(r['CERFA (ID)'] || ''));
  archiveRow_(MES.trashMouv, mv, r, user);
  mv.sh.deleteRow(r._row);
}

function deleteManip_(d, user) {
  var f = findMouv_(d.cerfa);
  var c = getChantier_(String(f.row['ID chantier']));
  var row = mesRows_(c.id)[0];
  if (!canEditMes_(c, row, user)) throw new Error(DROITS_MSG);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { deleteManipRow_(f.t, f.row, user); } finally { lock.releaseLock(); }
  return true;
}

function deleteMes_(d, user) {
  var c = getChantier_(d.chantierId);
  var row = mesRows_(c.id)[0];
  if (!row) throw new Error('Aucune mise en service sur ce chantier.');
  if (!canEditMes_(c, row, user)) throw new Error(DROITS_MSG);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    // Les opérations d'abord, de bas en haut pour garder les numéros de ligne
    var mv = mouvTable_();
    mv.rows.filter(function (r) { return String(r['ID MES']) === String(row['ID']); }).reverse().forEach(function (r) { deleteManipRow_(mv, r, user); });
    trashFile_(String(row['Rapport (ID)'] || ''));
    trashFile_(String(row['Signature (ID)'] || ''));
    var phs = []; try { phs = JSON.parse(row['Photos'] || '[]'); } catch (e) { phs = []; }
    phs.forEach(function (ph) { trashFile_(ph.id); });
    var mt = readTable_(MES.sheet);
    archiveRow_(MES.trashMes, mt, mt.rows.filter(function (r) { return String(r['ID']) === String(row['ID']); })[0], user);
    mt.sh.deleteRow(row._row);
    var ct = readTable_(SH.chantiers), crow = ct.rows.filter(function (r) { return r['ID'] === c.id; })[0];
    writeRow_(ct, { 'Date MES': '' }, crow._row);
  } finally { lock.releaseLock(); }
  return getChantier_(c.id);
}

/* ---------- Lecture du n° de série sur la photo de la plaque */

function readPlate_(d) {
  if (!d.data) throw new Error('Aucune photo reçue.');
  var prompt = 'Photo de la plaque signalétique d’une ' + (d.which === 'UI' ? 'unité intérieure' : 'unité extérieure') + ' de pompe à chaleur ou de climatisation.\n' +
    'Réponds UNIQUEMENT par un objet JSON : {"serie": numéro de série (Serial No., N° de série, S/N), "modele": référence du modèle, "marque": marque, "fluide": fluide frigorigène (ex. R32), "charge": charge en fluide en kg (nombre)}.\n' +
    'Mets "" pour ce qui est absent ou illisible. Ne devine jamais : recopie exactement les caractères du numéro de série.';
  var text = callClaude_([{ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: String(d.data) } }, { type: 'text', text: prompt }], 300);
  var i = text.indexOf('{'), j = text.lastIndexOf('}');
  var raw = {}; try { raw = JSON.parse(text.slice(i, j + 1)); } catch (e) { raw = {}; }
  var out = {};
  ['serie', 'modele', 'marque', 'fluide', 'charge'].forEach(function (k) { out[k] = raw[k] === null || raw[k] === undefined ? '' : String(raw[k]).trim().slice(0, 80); });
  if (!out.serie) throw new Error('N° de série illisible : reprends la photo plus près, ou saisis-le à la main.');
  return out;
}

/* ================================================================== Livraison 3 : bouteilles, détecteurs, distributeurs, déclaration, profil */

var L3 = { trashBout: '_Bouteilles supprimées', declFolder: ['ARE Fluides', 'Déclarations'], opsFolder: ['ARE Fluides', 'Opérateurs'] };
var BOUT_TYPES = { neuve: 'Neuve', recup: 'Récupération', transfert: 'Transfert' };
var BOUT_COLS = ['Réf bouteille', 'Type', 'Fluide', 'Poids à vide (kg)', 'Poids plein à l’achat (kg)', 'Fluide initial (kg)', 'Dernière pesée (kg)', 'Fluide restant (kg)',
  'Statut', 'Distributeur (ID)', 'Date de retour', 'N° BSFF', 'Date d’entrée', 'Fournisseur achat', 'Capacité max (kg)'];

function boutKind_(type) { var t = String(type || '').toLowerCase(); return /neuv|neuf/.test(t) ? 'neuve' : /transf/.test(t) ? 'transfert' : 'recup'; }
function boutTable_() { ensureSheet_(SH.bout, BOUT_COLS); ensureColumn_(SH.bout, 'Capacité max (kg)'); return readTable_(SH.bout); }
function round3_(n) { return Math.round(n * 1000) / 1000; }

function bouteilleFromRow_(r, dists) {
  var vide = num_(r['Poids à vide (kg)']), plein = num_(r['Poids plein à l’achat (kg)']), pesee = num_(r['Dernière pesée (kg)']);
  var kind = boutKind_(r['Type']), rendue = /rendu/i.test(String(r['Statut'] || ''));
  var dist = (dists || []).filter(function (x) { return x.id === String(r['Distributeur (ID)'] || ''); })[0];
  return {
    ref: String(r['Réf bouteille']).trim(), numero: boutNumero_(r['Réf bouteille']), kind: kind, type: BOUT_TYPES[kind], fluide: normFluide_(r['Fluide']),
    vide: fr_(vide, 2), plein: kind === 'neuve' ? fr_(plein, 2) : '', initial: kind === 'neuve' && isFinite(plein) && isFinite(vide) ? fr_(plein - vide, 2) : '',
    pesee: fr_(pesee, 2), reste: isFinite(vide) && isFinite(pesee) ? fr_(Math.max(0, pesee - vide), 2) : '',
    resteN: isFinite(vide) && isFinite(pesee) ? Math.max(0, pesee - vide) : 0,
    statut: rendue ? 'rendue' : 'depot', statutLabel: rendue ? 'Rendue' : 'Au dépôt',
    distributeur: String(r['Distributeur (ID)'] || ''), distributeurNom: dist ? dist.nom : '', dateRetour: String(r['Date de retour'] || ''),
    bsff: String(r['N° BSFF'] || ''), dateEntree: String(r['Date d’entrée'] || ''), fournisseur: String(r['Fournisseur achat'] || ''),
    capacite: kind === 'neuve' ? '' : fr_(num_(r['Capacité max (kg)']), 2)
  };
}

function listBouteilles_() {
  var dists = distributeurs_(), last = boutDernieresOps_();
  var list = boutTable_().rows.filter(function (r) { return String(r['Réf bouteille'] || '').trim(); }).map(function (r) {
    var b = bouteilleFromRow_(r, dists), lo = last[b.ref.toLowerCase()];
    b.derniereOp = lo ? Utilities.formatDate(lo, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
    return b;
  });
  // Totaux au dépôt par fluide : [{ fluide, kg }]
  var sum = function (f) {
    var by = {};
    list.filter(f).forEach(function (b) { if (b.resteN > 0) { var k = b.fluide || '?'; by[k] = (by[k] || 0) + b.resteN; } });
    return Object.keys(by).sort().map(function (k) { return { fluide: k === '?' ? 'fluide non renseigné' : k, kg: fr_(by[k], 2) }; });
  };
  var fournisseurs = {};
  list.forEach(function (b) { if (b.fournisseur) fournisseurs[b.fournisseur] = 1; });
  return {
    bouteilles: list, distributeurs: dists, fournisseurs: Object.keys(fournisseurs).sort(),
    totals: { neuf: sum(function (b) { return b.statut === 'depot' && b.kind === 'neuve'; }), recup: sum(function (b) { return b.statut === 'depot' && b.kind !== 'neuve'; }) },
    fluides: listFluides_().map(function (f) { return f.code; })
  };
}

function findBout_(ref) {
  var t = boutTable_(), key = String(ref || '').trim().toLowerCase();
  var row = t.rows.filter(function (r) { return String(r['Réf bouteille'] || '').trim().toLowerCase() === key; })[0];
  return { t: t, row: row };
}

function bouteilleDetail_(d) {
  var f = findBout_(d.ref);
  if (!f.row) throw new Error('Bouteille introuvable.');
  var b = bouteilleFromRow_(f.row, distributeurs_());
  var hist = [];
  readTable_(SH.mouv).rows.forEach(function (r) {
    if (String(r['Bouteille'] || '').trim().toLowerCase() !== b.ref.toLowerCase()) return;
    var v = num_(r['Variation bouteille (kg)']);
    if (!isFinite(v)) v = num_(r['Poids après (kg)']) - num_(r['Poids avant (kg)']);
    hist.push({ date: String(r['Date'] || ''), op: String(r['Opération'] || ''), client: String(r['Chantier / Client'] || ''), cerfa: String(r['N° CERFA'] || ''), qty: isFinite(v) ? (v > 0 ? '+' : '−') + fr_(Math.abs(v), 2) : '' });
  });
  if (b.kind === 'neuve' && b.initial) hist.push({ date: b.dateEntree, op: 'Achat' + (b.fournisseur ? ' · ' + b.fournisseur : ''), client: '', cerfa: '', qty: '+' + b.initial });
  if (b.statut === 'rendue') hist.push({ date: b.dateRetour, op: 'Rendue' + (b.distributeurNom ? ' à ' + b.distributeurNom : ''), client: '', cerfa: '', qty: b.reste ? '−' + b.reste : '' });
  var key = function (h) { var x = parseFr_(h.date); return x ? x.getTime() : 0; };
  hist.sort(function (a, c) { return key(c) - key(a); });
  b.hist = hist;
  b.utilisee = hist.some(function (h) { return h.cerfa; });
  return b;
}

/** Ajout (achat ou bouteille déjà en stock) ou modification des caractéristiques d'une bouteille. Admin. */
/** Numéro d'une bouteille sans son préfixe (RECUP_, TRANS_, R32_…). */
function boutNumero_(ref) { return String(ref || '').trim().replace(/^(RECUP|TRANS|R\d+[A-Z]?)_?(?=[0-9A-Z])/i, ''); }
/** Référence composée : RECUP_numéro, TRANS_numéro, ou fluide_numéro pour une neuve (R32_03482). */
function boutRef_(kind, fluide, numero) {
  var n = boutNumero_(String(numero || '').replace(/\s+/g, '')).toUpperCase();
  if (!n) return '';
  return (kind === 'neuve' ? String(fluide || '').toUpperCase() : kind === 'transfert' ? 'TRANS' : 'RECUP') + '_' + n;
}

function saveBouteille_(d, user) {
  var kind = BOUT_TYPES[d.kind] ? d.kind : '';
  if (!kind) throw new Error('Choisis le type de bouteille.');
  var fluide = normFluide_(d.fluide);
  if (kind === 'neuve' && !fluide) throw new Error('Choisis le fluide de la bouteille neuve.');
  var ref = d.numero !== undefined ? boutRef_(kind, fluide, d.numero) : String(d.ref || '').trim();
  if (!ref) throw new Error('Indique le numéro de la bouteille.');
  if (d.original && d.numero !== undefined) {
    // Bouteille existante dont ni le numéro, ni le type, ni le fluide (neuve) ne changent : on garde sa référence telle quelle
    var o = findBout_(d.original).row;
    if (o && boutNumero_(d.original).toUpperCase() === boutNumero_(String(d.numero).replace(/\s+/g, '')).toUpperCase() && boutKind_(o['Type']) === kind && (kind !== 'neuve' || normFluide_(o['Fluide']) === fluide)) ref = String(d.original).trim();
  }
  var vide = num_(d.vide), plein = num_(d.plein), pesee = num_(d.pesee);
  if (!isFinite(vide) || vide <= 0) throw new Error('Indique le poids à vide (tare) de la bouteille.');
  if (kind === 'neuve') {
    if (!fluide) throw new Error('Choisis le fluide de la bouteille neuve.');
    if (!isFinite(plein) || plein <= vide) throw new Error('Le poids plein doit être supérieur au poids à vide.');
  }
  if (isFinite(pesee) && pesee < vide - 0.05) throw new Error('La pesée actuelle ne peut pas être inférieure au poids à vide.');
  var dateEntree = String(d.dateEntree || '').trim();
  if (dateEntree && !parseFr_(dateEntree)) throw new Error('Date au format jj/mm/aaaa.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var f = findBout_(d.original || ref);
    if (d.original && !f.row) throw new Error('Bouteille introuvable.');
    if (!d.original && f.row) throw new Error('Une bouteille porte déjà cette référence.');
    if (d.original && ref.toLowerCase() !== String(d.original).trim().toLowerCase()) {
      if (findBout_(ref).row) throw new Error('Une bouteille porte déjà cette référence.');
      var used = readTable_(SH.mouv).rows.some(function (r) { return String(r['Bouteille'] || '').trim().toLowerCase() === String(d.original).trim().toLowerCase(); });
      if (used) throw new Error('Cette bouteille figure déjà dans le registre : son numéro, son type et son fluide ne peuvent plus changer.');
    }
    var rec = {
      'Réf bouteille': ref, 'Type': BOUT_TYPES[kind], 'Fluide': fluide, 'Poids à vide (kg)': vide,
      'Poids plein à l’achat (kg)': kind === 'neuve' ? plein : '', 'Fluide initial (kg)': kind === 'neuve' ? round3_(plein - vide) : '',
      'Date d’entrée': dateEntree
    };
    if (d.fournisseur !== undefined) rec['Fournisseur achat'] = String(d.fournisseur || '').trim();
    if (d.capacite !== undefined) { var cap = num_(d.capacite); rec['Capacité max (kg)'] = kind !== 'neuve' && isFinite(cap) && cap > 0 ? cap : ''; }
    if (d.distributeur !== undefined) {
      if (d.distributeur && !distributeurs_().some(function (x) { return x.id === d.distributeur; })) throw new Error('Distributeur inconnu.');
      rec['Distributeur (ID)'] = kind === 'neuve' ? '' : String(d.distributeur || '');
    }
    if (!f.row) {
      var p0 = isFinite(pesee) ? pesee : (kind === 'neuve' ? plein : vide);
      rec['Dernière pesée (kg)'] = p0; rec['Fluide restant (kg)'] = round3_(Math.max(0, p0 - vide)); rec['Statut'] = 'Au dépôt';
    } else if (isFinite(pesee)) {
      rec['Dernière pesée (kg)'] = pesee; rec['Fluide restant (kg)'] = round3_(Math.max(0, pesee - vide));
    } else {
      var pOld = num_(f.row['Dernière pesée (kg)']);
      if (isFinite(pOld)) rec['Fluide restant (kg)'] = round3_(Math.max(0, pOld - vide));
    }
    writeRow_(f.t, rec, f.row ? f.row._row : null);
  } finally { lock.releaseLock(); }
  return bouteilleDetail_({ ref: ref });
}

/** Au dépôt ou rendue au distributeur (avec la date et le BSFF). Admin. */
function setBouteilleStatut_(d, user) {
  var f = findBout_(d.ref);
  if (!f.row) throw new Error('Bouteille introuvable.');
  if (d.statut === 'rendue') {
    if (!d.distributeur || !distributeurs_().some(function (x) { return x.id === d.distributeur; })) throw new Error('Choisis le distributeur à qui la bouteille est rendue.');
    if (!parseFr_(d.dateRetour)) throw new Error('Indique la date de retour (jj/mm/aaaa).');
    if (parseFr_(d.dateRetour) > new Date()) throw new Error('La date de retour ne peut pas être dans le futur.');
    writeRow_(f.t, { 'Statut': 'Rendue', 'Distributeur (ID)': d.distributeur, 'Date de retour': d.dateRetour, 'N° BSFF': String(d.bsff || '').trim() }, f.row._row);
  } else {
    writeRow_(f.t, { 'Statut': 'Au dépôt', 'Date de retour': '', 'N° BSFF': '' }, f.row._row);
  }
  return bouteilleDetail_({ ref: d.ref });
}

/** Suppression d'une bouteille jamais utilisée (erreur de saisie). Copie gardée dans un onglet masqué. Admin. */
function deleteBouteille_(d, user) {
  var f = findBout_(d.ref);
  if (!f.row) throw new Error('Bouteille introuvable.');
  var used = readTable_(SH.mouv).rows.some(function (r) { return String(r['Bouteille'] || '').trim().toLowerCase() === String(d.ref).trim().toLowerCase(); });
  if (used) throw new Error('Cette bouteille figure dans le registre des mouvements : elle ne peut pas être supprimée.');
  archiveRow_(L3.trashBout, f.t, f.row, user);
  f.t.sh.deleteRow(f.row._row);
  return true;
}

/* ---------- Détecteurs de fuite (onglet Outillage) */

function listDetecteurs_() {
  return outils_().filter(function (o) { return /d[ée]tect/i.test(o.type); });
}

function saveDetecteur_(d, user) {
  var marque = String(d.marque || '').trim(), ref = String(d.reference || '').trim();
  if (!marque && !ref) throw new Error('Indique la marque et le modèle du détecteur.');
  var last = parseFr_(d.dernierControle);
  if (!last) throw new Error('Indique la date du dernier contrôle (jj/mm/aaaa).');
  if (last > new Date()) throw new Error('La date du dernier contrôle ne peut pas être dans le futur.');
  var freq = num_(d.frequence); if (!isFinite(freq) || freq <= 0) freq = 12;
  var next = new Date(last.getTime()); next.setMonth(next.getMonth() + freq);
  ensureSheet_(SH.outils, ['ID', 'Type', 'Marque', 'Référence', 'N° de série', 'Dernier contrôle', 'Fréquence (mois)', 'Prochain contrôle', 'Statut']);
  var t = readTable_(SH.outils);
  var row = d.id ? t.rows.filter(function (r) { return String(r['ID']) === d.id; })[0] : null;
  if (d.id && !row) throw new Error('Détecteur introuvable.');
  var rec = {
    'Type': 'Détecteur de fuite', 'Marque': marque, 'Référence': ref, 'N° de série': String(d.serie || '').trim(),
    'Dernier contrôle': d.dernierControle, 'Fréquence (mois)': freq, 'Prochain contrôle': next, 'Statut': next >= new Date() ? 'À jour' : 'À contrôler'
  };
  if (!row) rec['ID'] = nextId_(t.rows, 'T', 'ID');
  writeRow_(t, rec, row ? row._row : null);
  return listDetecteurs_();
}

/* ---------- Distributeurs */

function saveDistributeur_(d, user) {
  var nom = String(d.nom || '').trim();
  if (!nom) throw new Error('Indique le nom du distributeur.');
  var siret = String(d.siret || '').replace(/\s/g, '');
  if (siret && !/^\d{14}$/.test(siret)) throw new Error('Le SIRET fait 14 chiffres.');
  ensureSheet_(SH.dist, ['ID', 'Nom', 'SIRET', 'Adresse du site de reprise']);
  var t = readTable_(SH.dist);
  var row = d.id ? t.rows.filter(function (r) { return String(r['ID']) === d.id; })[0] : null;
  if (d.id && !row) throw new Error('Distributeur introuvable.');
  var rec = { 'Nom': nom, 'SIRET': siret ? "'" + siret : '', 'Adresse du site de reprise': String(d.adresse || '').trim() };
  if (!row) rec['ID'] = nextId_(t.rows, 'D', 'ID');
  writeRow_(t, rec, row ? row._row : null);
  return distributeurs_();
}

/* ---------- Profil et attestation */

function updateProfile_(user, d) {
  var t = readTable_(SH.ops), row = t.rows.filter(function (r) { return r['ID'] === user.id; })[0];
  if (!row) throw new Error('Compte introuvable.');
  var prenom = String(d.prenom || '').trim(), nom = String(d.nom || '').trim();
  if (!prenom && !nom) throw new Error('Indique ton prénom et ton nom.');
  writeRow_(t, { 'Prénom': prenom, 'Nom': nom, 'N° attestation d’aptitude': String(d.attestation || '').trim() }, row._row);
  return operatorPublic_(readTable_(SH.ops).rows.filter(function (r) { return r['ID'] === user.id; })[0]);
}

/** Fichier de l'attestation : rangé dans le Drive (dossier Opérateurs), jamais renvoyé ni affiché par l'appli. */
function uploadAttestation_(user, d) {
  var id = d.opId && d.opId !== user.id ? (admin_(user), d.opId) : user.id;
  var mime = d.mime === 'application/pdf' ? 'application/pdf' : 'image/jpeg';
  if (!d.data || String(d.data).length > 14 * 1024 * 1024) throw new Error('Fichier manquant ou trop lourd (10 Mo maximum).');
  ensureColumn_(SH.ops, 'Attestation (fichier)');
  ensureColumn_(SH.ops, 'Attestation chargée');
  var t = readTable_(SH.ops), row = t.rows.filter(function (r) { return r['ID'] === id; })[0];
  if (!row) throw new Error('Compte introuvable.');
  var name = 'Attestation ' + ((row['Prénom'] || '') + ' ' + (row['Nom'] || '')).trim() + ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') + (mime === 'application/pdf' ? '.pdf' : '.jpg');
  var file = saveFile_(folderPath_(L3.opsFolder), d.data, mime, name);
  trashFile_(String(row['Attestation (fichier)'] || ''));
  var date = today_();
  var up = { 'Attestation (fichier)': file.id, 'Attestation chargée': 'oui (' + date + ')' };
  // Lecture du n° d'attestation par Claude : seul ce numéro est gardé (rien d'autre du document n'est conservé ni affiché)
  var numero = '';
  try { numero = readAttestationNumero_(d.data, mime); } catch (e) { numero = ''; }
  if (numero) up['N° attestation d’aptitude'] = numero;
  writeRow_(t, up, row._row);
  return { date: date, numero: numero };
}

function readAttestationNumero_(data, mime) {
  if (!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY')) return '';
  var block = mime === 'application/pdf' ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: String(data) } }
    : { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: String(data) } };
  var prompt = 'Document : attestation d’aptitude à la manipulation des fluides frigorigènes (France).\n' +
    'Réponds UNIQUEMENT par un objet JSON : {"numero": "numéro de l’attestation d’aptitude"}.\n' +
    'Ne renvoie aucune autre information du document. Ne renvoie jamais un numéro de sécurité sociale, de certificat d’organisme ou de SIRET. Mets "" si le numéro de l’attestation est absent ou illisible.';
  var text = callClaude_([block, { type: 'text', text: prompt }], 100);
  var i = text.indexOf('{'), j = text.lastIndexOf('}');
  var raw = {}; try { raw = JSON.parse(text.slice(i, j + 1)); } catch (e) { raw = {}; }
  var n = String(raw.numero || '').trim().slice(0, 40);
  if (!/^[A-Za-z0-9][A-Za-z0-9\-\/ .]{2,39}$/.test(n)) return '';
  if (/^[12]\s?\d{2}\s?\d{2}/.test(n) && n.replace(/\D/g, '').length >= 13) return '';   // ressemble à un n° de sécurité sociale : ignoré
  return n;
}

/** À lancer une fois depuis l'éditeur Apps Script après la mise à jour : accepte les nouvelles autorisations (export Excel). */
function autoriserExcel() {
  var ss = SpreadsheetApp.create('ARE Fluides - test autorisation');
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  Logger.log('Export Excel opérationnel (jeton ' + (ScriptApp.getOAuthToken() ? 'OK' : 'absent') + ').');
}

function profileSummary_(user) {
  var det = listDetecteurs_();
  var out = { annee: new Date().getFullYear() - 1, detecteursAFaire: det.filter(function (x) { return !x.ok; }).length,
    bouteillesSansTare: boutTable_().rows.filter(function (r) { return String(r['Réf bouteille'] || '').trim() && !/rendu/i.test(String(r['Statut'] || '')) && !isFinite(num_(r['Poids à vide (kg)'])); }).length };
  try { out.alertes = alertes_().map(function (x) { return { kind: x.kind, text: x.text, grave: !!x.grave }; }); } catch (e) { out.alertes = []; }
  out.envoiClient = paramOn_('Envoi des documents au client');
  try { out.interventions = mesInterventions_(user, 5); } catch (e) { out.interventions = []; }
  if (user.role !== 'admin') return out;
  out.machinesAVerifier = ss_().getSheetByName(SH.machines) ? readTable_(SH.machines).rows.filter(function (r) { return String(r['À vérifier'] || '').trim(); }).length : 0;
  out.comptes = readTable_(SH.ops).rows.length;
  out.detecteurs = det.length;
  out.distributeurs = distributeurs_().length;
  return out;
}

/** Dernières interventions de l'utilisateur (mises en service, ajouts, récupérations, CERFA papier…), les plus récentes d'abord. */
function mesInterventions_(user, n) {
  if (!ss_().getSheetByName(SH.mouv)) return [];
  var rows = readTable_(SH.mouv).rows.filter(function (r) { return String(r['Opérateur ID'] || '') === String(user.id) && String(r['ID chantier'] || ''); });
  rows.sort(function (a, b) { var da = parseFr_(a['Date']), db = parseFr_(b['Date']); return ((db ? db.getTime() : 0) - (da ? da.getTime() : 0)) || (b._row - a._row); });
  return rows.slice(0, n).map(function (r) {
    var m = manipPublic_(r), data = {}; try { data = JSON.parse(r['Données'] || '{}'); } catch (e) { data = {}; }
    m.chantierId = String(r['ID chantier']);
    if ((data.nature || []).indexOf('mes') >= 0) m.label = 'Mise en service';
    return m;
  });
}

/* ---------- Déclaration annuelle (mêmes rubriques et même ordre que le site SOCOTEC) */

function declaration_(d) {
  var year = parseInt(d.annee, 10) || (new Date().getFullYear() - 1);
  var start = new Date(year, 0, 1), end = new Date(year, 11, 31, 23, 59, 59), prevEnd = new Date(year - 1, 11, 31, 23, 59, 59);
  var dists = distributeurs_();
  var bt = boutTable_().rows.filter(function (r) { return String(r['Réf bouteille'] || '').trim(); });
  var B = {};
  bt.forEach(function (r) {
    var ref = String(r['Réf bouteille']).trim().toLowerCase();
    B[ref] = { ref: String(r['Réf bouteille']).trim(), kind: boutKind_(r['Type']), fluide: normFluide_(r['Fluide']), vide: num_(r['Poids à vide (kg)']), plein: num_(r['Poids plein à l’achat (kg)']),
      entree: parseFr_(r['Date d’entrée']), rendue: /rendu/i.test(String(r['Statut'] || '')), retour: parseFr_(r['Date de retour']), dist: String(r['Distributeur (ID)'] || ''), mv: [] };
  });
  var flows = [];
  readTable_(SH.mouv).rows.forEach(function (r) {
    var b = B[String(r['Bouteille'] || '').trim().toLowerCase()];
    var date = parseFr_(r['Date']);
    if (!b || !date) return;
    var avant = num_(r['Poids avant (kg)']), apres = num_(r['Poids après (kg)']), v = num_(r['Variation bouteille (kg)']);
    if (!isFinite(v) && isFinite(avant) && isFinite(apres)) v = apres - avant;
    if (!isFinite(v) || v === 0) return;
    var nat = ''; try { nat = (JSON.parse(r['Données'] || '{}').nature || []).join(','); } catch (e) { nat = ''; }
    var m = { date: date, v: v, apres: apres, fluide: normFluide_(r['Fluide']) || b.fluide, cadre: String(r['Cadre'] || ''), nature: nat };
    b.mv.push(m);
    if (!b.fluide) b.fluide = m.fluide;
    flows.push({ b: b, m: m });
  });
  Object.keys(B).forEach(function (k) { B[k].mv.sort(function (a, c) { return a.date - c.date; }); });
  // Contenu d'une bouteille à une date : dernière pesée connue moins la tare
  var content = function (b, at) {
    if (b.entree && b.entree > at) return null;
    if (b.rendue && b.retour && b.retour <= at) return null;
    // Poids de départ : pesée avant de la 1re opération, sinon plein (neuve) ou tare (récupération) ; puis chaque opération jusqu'à la date
    var first = b.mv[0], w = first && isFinite(first.apres) ? first.apres - first.v : (b.kind === 'neuve' ? b.plein : b.vide);
    b.mv.forEach(function (m) { if (m.date <= at) w = isFinite(m.apres) ? m.apres : w + m.v; });
    return isFinite(w) && isFinite(b.vide) ? Math.max(0, w - b.vide) : 0;
  };
  var stock = function (fl, neuf, at) {
    return Object.keys(B).reduce(function (a, k) {
      var b = B[k]; if (b.fluide !== fl || (b.kind === 'neuve') !== neuf) return a;
      var c = content(b, at); return a + (c || 0);
    }, 0);
  };
  var inY = function (dt) { return dt && dt >= start && dt <= end; };
  var fluids = {};
  Object.keys(B).forEach(function (k) { var b = B[k]; if (b.fluide) fluids[b.fluide] = 1; });
  var out = [];
  Object.keys(fluids).sort().forEach(function (fl) {
    var neufB = Object.keys(B).map(function (k) { return B[k]; }).filter(function (b) { return b.fluide === fl && b.kind === 'neuve'; });
    var usB = Object.keys(B).map(function (k) { return B[k]; }).filter(function (b) { return b.fluide === fl && b.kind !== 'neuve'; });
    var n1 = stock(fl, true, prevEnd), n2 = stock(fl, true, end), u1 = stock(fl, false, prevEnd), u2 = stock(fl, false, end);
    var acquis = neufB.filter(function (b) { return inY(b.entree); }).reduce(function (a, b) { return a + (isFinite(b.plein - b.vide) ? b.plein - b.vide : 0); }, 0);
    var fy = flows.filter(function (x) { return x.m.fluide === fl && inY(x.m.date); });
    var charge = 0, chargeMes = 0, recup = 0, recupHU = 0, recycle = 0;
    fy.forEach(function (x) {
      if (x.b.kind === 'neuve' && x.m.v < 0) { charge += -x.m.v; if (/mise en service/i.test(x.m.cadre)) chargeMes += -x.m.v; }
      if (x.b.kind !== 'neuve' && x.m.v > 0) { recup += x.m.v; if (/d[ée]mant/i.test(x.m.cadre) || /demantel/.test(x.m.nature)) recupHU += x.m.v; }
      if (x.b.kind !== 'neuve' && x.m.v < 0) recycle += -x.m.v;
    });
    var remis = {};
    usB.filter(function (b) { return b.rendue && inY(b.retour); }).forEach(function (b) {
      var at = new Date(b.retour.getTime() - 1000);
      var c = content(b, at) || 0;
      remis[b.dist] = (remis[b.dist] || 0) + c;
    });
    var remisTotal = Object.keys(remis).reduce(function (a, k) { return a + remis[k]; }, 0);
    if (!n1 && !n2 && !u1 && !u2 && !acquis && !charge && !recup && !remisTotal && !recycle) return;
    var k1 = function (x) { return fr_(x, 2); };
    var pct = function (e, base) { return base > 0 ? fr_(e / base * 100, 1) + ' %' : '0,0 %'; };
    var ecN = n1 + acquis - charge - n2, ecU = u1 + recup - recycle - remisTotal - u2;
    var remisRows = Object.keys(remis).length ? Object.keys(remis).map(function (id) {
      var x = dists.filter(function (z) { return z.id === id; })[0];
      return { label: 'Quantité remise aux fournisseurs' + (x ? ' : ' + x.nom + (x.siret ? ', SIRET ' + x.siret : '') : ' : distributeur non renseigné'), value: k1(remis[id]) };
    }) : [{ label: 'Quantité remise aux fournisseurs', value: k1(0) }];
    out.push({
      code: fl, n1: k1(n1), n2: k1(n2), u1: k1(u1), u2: k1(u2),
      neuf: [
        { label: 'Stock de fluides neufs détenu au 1er janvier', value: k1(n1) },
        { label: 'Quantité acquise sur le territoire national', value: k1(acquis) },
        { label: 'Quantité acquise depuis un pays de l’UE', value: k1(0) },
        { label: 'Quantité acquise depuis un pays hors de l’UE', value: k1(0) },
        { label: 'Quantité chargée', value: k1(charge) },
        { label: 'dont chargé en mise en service d’équipements', value: k1(chargeMes), sub: true },
        { label: 'dont chargé en maintenance d’équipements', value: k1(charge - chargeMes), sub: true },
        { label: 'Quantité cédée à d’autres opérateurs', value: k1(0) },
        { label: 'Stock de fluides neufs détenu au 31 décembre (pesée)', value: k1(n2) },
        { label: 'Écart', value: pct(ecN, n1 + acquis), ecart: k1(ecN) }
      ],
      usage: [
        { label: 'Stock de fluides usagés détenu au 1er janvier', value: k1(u1) },
        { label: 'Quantité récupérée', value: k1(recup) },
        { label: 'dont sur des équipements hors d’usage', value: k1(recupHU), sub: true },
        { label: 'dont récupéré en maintenance d’équipements', value: k1(recup - recupHU), sub: true },
        { label: 'Quantité recyclée', value: k1(recycle) }
      ].concat(remisRows).concat([
        { label: 'Quantité regénérée', value: k1(0) },
        { label: 'Quantité détruite', value: k1(0) },
        { label: 'Stock de fluides usagés détenu au 31 décembre', value: k1(u2) },
        { label: 'Écart', value: pct(ecU, u1 + recup), ecart: k1(ecU) }
      ])
    });
  });
  var years = {};
  flows.forEach(function (x) { years[x.m.date.getFullYear()] = 1; });
  Object.keys(B).forEach(function (k) { if (B[k].entree) years[B[k].entree.getFullYear()] = 1; });
  years[new Date().getFullYear()] = 1; years[year] = 1;
  return { annee: year, fluides: out, annees: Object.keys(years).map(Number).sort() };
}

function exportDeclaration_(d) {
  var x = declaration_(d), p = params_();
  var e = function (s) { return String(s === undefined || s === null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var rows = function (list) { return list.map(function (r) { return '<tr' + (r.sub ? ' class="sub"' : '') + '><td>' + e(r.label) + '</td><td class="v">' + e(r.value) + '</td></tr>'; }).join(''); };
  var html = '<html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;font-size:11px;color:#14201A}h1{font-size:18px;margin:0 0 4px}h2{font-size:14px;margin:18px 0 6px;border-bottom:2px solid #1F6B45;padding-bottom:3px}' +
    'table{border-collapse:collapse;width:100%;margin-bottom:8px}td,th{border:1px solid #D5DCD7;padding:4px 6px;text-align:left}th{background:#E2F0E7}.v{text-align:right;width:90px}.sub td:first-child{padding-left:18px;color:#4A5A51}.note{color:#4A5A51}</style></head><body>' +
    '<h1>Déclaration annuelle des fluides frigorigènes ' + x.annee + '</h1><div class="note">' + e(p['Raison sociale'] || '') + (p['SIRET'] ? ' · SIRET ' + e(p['SIRET']) : '') +
    ' · du 01/01/' + x.annee + ' au 31/12/' + x.annee + ' · calculé depuis le registre des mouvements et le stock de bouteilles, en kg</div>' +
    '<h2>Vue d’ensemble (kg)</h2><table><tr><th>Code gaz</th><th class="v">Neuf 1er janv.</th><th class="v">Neuf 31 déc.</th><th class="v">Usagé 1er janv.</th><th class="v">Usagé 31 déc.</th></tr>' +
    x.fluides.map(function (f) { return '<tr><td>' + e(f.code) + '</td><td class="v">' + f.n1 + '</td><td class="v">' + f.n2 + '</td><td class="v">' + f.u1 + '</td><td class="v">' + f.u2 + '</td></tr>'; }).join('') + '</table>' +
    x.fluides.map(function (f) { return '<h2>' + e(f.code) + '</h2><table><tr><th colspan="2">Fluide neuf</th></tr>' + rows(f.neuf) + '</table><table><tr><th colspan="2">Fluide usagé</th></tr>' + rows(f.usage) + '</table>'; }).join('') +
    '<p class="note">Généré par ARE Fluides le ' + today_() + '.</p></body></html>';
  var name = 'Déclaration fluides ' + x.annee + '.pdf';
  var folder = folderPath_(L3.declFolder);
  var it = folder.getFilesByName(name); while (it.hasNext()) it.next().setTrashed(true);
  var file = folder.createFile(HtmlService.createHtmlOutput(html).getAs('application/pdf').setName(name));
  return { url: file.getUrl(), name: name };
}


/** Récapitulatif de la déclaration en Excel : classeur temporaire Google Sheets exporté en .xlsx dans Drive › Déclarations. */
function exportDeclarationXlsx_(d) {
  var x = declaration_(d), name = 'Déclaration fluides ' + x.annee;
  var val = function (v) { var n = num_(v); return isFinite(n) ? n : v; };
  var ss = SpreadsheetApp.create(name);
  var sh = ss.getSheets()[0];
  sh.setName('Vue d’ensemble');
  var rows = [['Code gaz', 'Neuf 1er janvier (kg)', 'Neuf 31 décembre (kg)', 'Usagé 1er janvier (kg)', 'Usagé 31 décembre (kg)']].concat(
    x.fluides.map(function (f) { return [f.code, val(f.n1), val(f.n2), val(f.u1), val(f.u2)]; }));
  sh.getRange(1, 1, rows.length, 5).setValues(rows);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold');
  if (rows.length > 1) sh.getRange(2, 2, rows.length - 1, 4).setNumberFormat('0.00');
  sh.setColumnWidth(1, 120);
  x.fluides.forEach(function (f) {
    var t = ss.insertSheet(f.code);
    var r = [['Fluide neuf', 'kg']].concat(f.neuf.map(function (l) { return [(l.sub ? '    ' : '') + l.label, l.ecart !== undefined ? l.value + ' (' + l.ecart + ' kg)' : val(l.value)]; }))
      .concat([['', ''], ['Fluide usagé', 'kg']]).concat(f.usage.map(function (l) { return [(l.sub ? '    ' : '') + l.label, l.ecart !== undefined ? l.value + ' (' + l.ecart + ' kg)' : val(l.value)]; }));
    t.getRange(1, 1, r.length, 2).setValues(r);
    t.getRange(1, 1, 1, 2).setFontWeight('bold');
    t.getRange(f.neuf.length + 3, 1, 1, 2).setFontWeight('bold');
    t.setColumnWidth(1, 420);
  });
  SpreadsheetApp.flush();
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  if (resp.getResponseCode() !== 200) throw new Error('Export Excel impossible (' + resp.getResponseCode() + ').');
  var folder = folderPath_(L3.declFolder);
  var it = folder.getFilesByName(name + '.xlsx'); while (it.hasNext()) it.next().setTrashed(true);
  var file = folder.createFile(resp.getBlob().setName(name + '.xlsx'));
  return { url: file.getUrl(), name: name + '.xlsx' };
}


/** Règle de date : une manipulation ne peut pas être datée d'avant la dernière opération déjà enregistrée sur la bouteille
 *  (ni d'avant son entrée au dépôt), sinon les pesées ne suivraient plus l'ordre chronologique. */
function checkDateBouteille_(b, dateFr) {
  if (!b) return;
  var dt = parseFr_(dateFr), last = parseFr_(b.derniereOp), entree = parseFr_(b.dateEntree);
  if (dt && last && dt < last) throw new Error('La bouteille ' + b.ref + ' a déjà une opération enregistrée le ' + b.derniereOp + ' : la date ne peut pas être antérieure. Choisis une date à partir du ' + b.derniereOp + '.');
  if (dt && entree && dt < entree) throw new Error('La bouteille ' + b.ref + ' est entrée au dépôt le ' + b.dateEntree + ' : la date ne peut pas être antérieure.');
}

/* ================================================================== Livraison 4 */

var L4 = { registreFolder: ['ARE Fluides', 'Registre'], cerfaPapierFolder: ['ARE Fluides', 'CERFA papier'], controles: 'Pesées de contrôle',
  trashMach: '_Machines supprimées', tolPesee: 0.01 };

function setParam_(key, value) {
  ensureSheet_(SH.params, ['Paramètre', 'Valeur']);
  var t = readTable_(SH.params), row = t.rows.filter(function (r) { return String(r['Paramètre']).trim() === key; })[0];
  writeRow_(t, { 'Paramètre': key, 'Valeur': value }, row ? row._row : null);
}
function paramOn_(key) { return /^(oui|true|1)$/i.test(String(params_()[key] || '').trim()); }

/* ---------- Réglages (admin) */
function reglages_() {
  return { envoiClient: paramOn_('Envoi des documents au client'), alertesEmail: paramOn_('E-mails d’alerte'), emailAlertes: String(params_()['E-mail des alertes'] || '') };
}
function saveReglages_(d) {
  if (d.envoiClient !== undefined) setParam_('Envoi des documents au client', d.envoiClient ? 'oui' : 'non');
  if (d.alertesEmail !== undefined) setParam_('E-mails d’alerte', d.alertesEmail ? 'oui' : 'non');
  if (d.emailAlertes !== undefined) {
    var e = String(d.emailAlertes || '').trim();
    if (e && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) throw new Error('Adresse e-mail des alertes invalide.');
    setParam_('E-mail des alertes', e);
  }
  return reglages_();
}

/* ---------- Base machines : vue complète et validation des propositions */
function machinesBase_() {
  if (!ss_().getSheetByName(SH.machines)) return [];
  var na = function (v) { return v === '' || v === null || v === undefined ? '' : String(v).replace('.', ','); };
  return readTable_(SH.machines).rows.filter(function (r) { return String(r['Modèle'] || '').trim(); }).map(function (r) {
    var props = []; try { props = JSON.parse(r['Propositions (JSON)'] || '[]'); } catch (e) { props = []; }
    return {
      row: r._row, marque: String(r['Marque'] || ''), modele: String(r['Modèle'] || ''), appareil: String(r['Type d’appareil'] || ''), gamme: String(r['Gamme'] || ''),
      type: String(r['Type d’installation'] || ''), alimentation: String(r['Alimentation'] || ''), liaisons: String(r['Liaisons'] || ''), fluide: String(r['Fluide'] || ''),
      charge: na(r['Charge usine (kg)']), precharge: na(r['Longueur préchargée (m)']), gpm: na(r['Charge supplémentaire (g/m)']), lmax: na(r['Longueur max UE-UI (m)']),
      source: String(r['Source'] || ''), remarque: String(r['Remarque'] || ''), aVerifier: String(r['À vérifier'] || ''),
      nouveau: props.some(function (p) { return p.nouveau; }), propositions: props.filter(function (p) { return !p.nouveau; }).map(function (p) { return p.par + ' (' + p.date + ') : ' + (p.texte || ''); })
    };
  });
}

function validateMachine_(d, user) {
  var t = readTable_(SH.machines), row = t.rows.filter(function (r) { return r._row === Number(d.row); })[0];
  if (!row || normModele_(row['Modèle']) !== normModele_(d.modele)) throw new Error('Machine introuvable : recharge la liste.');
  var props = []; try { props = JSON.parse(row['Propositions (JSON)'] || '[]'); } catch (e) { props = []; }
  var qui = ((user.prenom || '') + ' ' + (user.nom || '')).trim(), quand = today_();
  var nouveau = props.some(function (p) { return p.nouveau; });
  if (d.action === 'accept') {
    var up = { 'À vérifier': '', 'Propositions (JSON)': '' };
    props.forEach(function (p) { if (p.up) Object.keys(p.up).forEach(function (k) { up[k] = p.up[k]; }); });
    up['Source'] = (String(row['Source'] || '') + ' · validé par ' + qui + ' le ' + quand).slice(0, 1000);
    writeRow_(t, up, row._row);
    return { msg: 'Modifications de ' + row['Modèle'] + ' acceptées.', list: machinesBase_() };
  }
  if (d.action === 'refuse') {
    if (nouveau) {
      archiveRow_(L4.trashMach, t, row, user);
      t.sh.deleteRow(row._row);
      return { msg: 'Modèle ' + row['Modèle'] + ' retiré de la base (copie dans l’onglet masqué « ' + L4.trashMach + ' »).', list: machinesBase_() };
    }
    writeRow_(t, { 'À vérifier': '', 'Propositions (JSON)': '' }, row._row);
    return { msg: 'Proposition refusée : ' + row['Modèle'] + ' reste inchangé.', list: machinesBase_() };
  }
  if (d.action === 'edit') {
    var f = d.fields || {}, M = { marque: 'Marque', modele: 'Modèle', appareil: 'Type d’appareil', gamme: 'Gamme', type: 'Type d’installation', alimentation: 'Alimentation',
      liaisons: 'Liaisons', fluide: 'Fluide', charge: 'Charge usine (kg)', precharge: 'Longueur préchargée (m)', gpm: 'Charge supplémentaire (g/m)', lmax: 'Longueur max UE-UI (m)', remarque: 'Remarque' };
    var nm = String(f.modele || '').trim();
    if (!nm) throw new Error('Le modèle est obligatoire.');
    if (!String(f.marque || '').trim()) throw new Error('La marque est obligatoire.');
    if (t.rows.some(function (r) { return r._row !== row._row && normModele_(r['Modèle']) === normModele_(nm); })) throw new Error('Le modèle ' + nm + ' existe déjà dans la base.');
    ['charge', 'precharge', 'gpm', 'lmax'].forEach(function (k) { var v = String(f[k] == null ? '' : f[k]).trim(); if (v && !/^na$/i.test(v) && !/^\d+([.,]\d+)?$/.test(v)) throw new Error('Valeur numérique attendue : ' + M[k] + '.'); });
    var w = {};
    Object.keys(M).forEach(function (k) { if (f[k] !== undefined && t.headers.indexOf(M[k]) >= 0) w[M[k]] = String(f[k]).trim(); });
    w['Source'] = (String(row['Source'] || '') + ' · modifié par ' + qui + ' le ' + quand).slice(0, 1000);
    writeRow_(t, w, row._row);
    return { msg: 'Modèle ' + nm + ' enregistré.', list: machinesBase_() };
  }
  if (d.action === 'delete') {
    archiveRow_(L4.trashMach, t, row, user);
    t.sh.deleteRow(row._row);
    return { msg: 'Modèle ' + row['Modèle'] + ' supprimé de la base (copie dans l’onglet masqué « ' + L4.trashMach + ' »).', list: machinesBase_() };
  }
  throw new Error('Action inconnue.');
}

/* ---------- Registre des mouvements : export d'une période (PDF ou Excel) */
function registreRows_(du, au) {
  var a = parseFr_(du), b = parseFr_(au);
  if (!a || !b) throw new Error('Indique la période (jj/mm/aaaa).');
  if (a > b) throw new Error('La date de début est après la date de fin.');
  b = new Date(b.getFullYear(), b.getMonth(), b.getDate(), 23, 59, 59);
  return readTable_(SH.mouv).rows.filter(function (r) { var dt = parseFr_(r['Date']); return dt && dt >= a && dt <= b; }).map(function (r) {
    var data = {}; try { data = JSON.parse(r['Données'] || '{}'); } catch (e) { data = {}; }
    var v = num_(r['Variation bouteille (kg)']), q = num_(data.quantite);
    if (!isFinite(q)) q = isFinite(v) ? Math.abs(v) : NaN;
    var dt = parseFr_(r['Date']);
    return { t: dt.getTime(), date: String(r['Date']), operation: String(r['Opération'] || ''), cadre: String(r['Cadre'] || ''), client: String(r['Chantier / Client'] || ''),
      machine: String(r['Machine'] || ''), fluide: String(r['Fluide'] || ''), quantite: isFinite(q) && q ? Math.round(q * 1000) / 1000 : '',
      bouteille: String(r['Bouteille'] || ''), charge: num_(r['Charge totale machine (kg)']), teq: num_(r['t éq. CO2']), operateur: String(r['Opérateur'] || ''), cerfa: String(r['N° CERFA'] || '') };
  }).sort(function (x, y) { return x.t - y.t; });
}
var REG_HEAD = ['Date', 'Opération', 'Chantier / client', 'Machine', 'Fluide', 'Quantité (kg)', 'Bouteille', 'Charge totale (kg)', 't éq. CO₂', 'Opérateur', 'N° CERFA'];
function regLine_(x) { return [x.date, x.operation + (x.cadre && x.cadre !== x.operation ? ' (' + x.cadre + ')' : ''), x.client, x.machine, x.fluide, x.quantite, x.bouteille, isFinite(x.charge) ? x.charge : '', isFinite(x.teq) ? x.teq : '', x.operateur, x.cerfa]; }

function registreApercu_(d) { var l = registreRows_(d.du, d.au); return { n: l.length, kg: fr_(l.reduce(function (s, x) { return s + (x.quantite || 0); }, 0), 2) }; }

function exportRegistre_(d) {
  var rows = registreRows_(d.du, d.au), p = params_();
  var name = 'Registre fluides du ' + String(d.du).replace(/\//g, '-') + ' au ' + String(d.au).replace(/\//g, '-');
  var folder = folderPath_(L4.registreFolder);
  if (d.format === 'xlsx') {
    var ss = SpreadsheetApp.create(name), sh = ss.getSheets()[0];
    sh.setName('Registre');
    var data = [REG_HEAD].concat(rows.map(regLine_));
    sh.getRange(1, 1, data.length, REG_HEAD.length).setValues(data);
    sh.getRange(1, 1, 1, REG_HEAD.length).setFontWeight('bold');
    sh.setColumnWidth(3, 220); sh.setColumnWidth(4, 200);
    SpreadsheetApp.flush();
    var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    if (resp.getResponseCode() !== 200) throw new Error('Export Excel impossible (' + resp.getResponseCode() + ').');
    var it = folder.getFilesByName(name + '.xlsx'); while (it.hasNext()) it.next().setTrashed(true);
    var f = folder.createFile(resp.getBlob().setName(name + '.xlsx'));
    return { url: f.getUrl(), name: name + '.xlsx', n: rows.length };
  }
  var e = function (v) { return String(v === undefined || v === null ? '' : v).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var html = '<html><head><meta charset="utf-8"><style>@page{size:A4 landscape;margin:12mm}body{font-family:Arial,sans-serif;font-size:9px;color:#1B2448}h1{font-size:16px;margin:0 0 4px}' +
    'table{border-collapse:collapse;width:100%}th,td{border:1px solid #D5D9E6;padding:3px 4px;text-align:left;vertical-align:top}th{background:#E6E9F4}.n{text-align:right}.note{color:#4E5670}</style></head><body>' +
    '<h1>Registre des mouvements de fluides frigorigènes</h1><div class="note">' + e(p['Raison sociale'] || '') + (p['SIRET'] ? ' · SIRET ' + e(p['SIRET']) : '') + ' · du ' + e(d.du) + ' au ' + e(d.au) + ' · ' + rows.length + ' opération' + (rows.length > 1 ? 's' : '') + '</div><br>' +
    '<table><tr>' + REG_HEAD.map(function (h, i) { return '<th' + (i >= 5 && i <= 8 && i !== 6 ? ' class="n"' : '') + '>' + e(h) + '</th>'; }).join('') + '</tr>' +
    rows.map(function (x) { return '<tr>' + regLine_(x).map(function (v, i) { return '<td' + (i >= 5 && i <= 8 && i !== 6 ? ' class="n"' : '') + '>' + e(typeof v === 'number' ? String(v).replace('.', ',') : v) + '</td>'; }).join('') + '</tr>'; }).join('') +
    '</table><p class="note">Généré par ARE Fluides le ' + today_() + '.</p></body></html>';
  var it2 = folder.getFilesByName(name + '.pdf'); while (it2.hasNext()) it2.next().setTrashed(true);
  var file = folder.createFile(HtmlService.createHtmlOutput(html).getAs('application/pdf').setName(name + '.pdf'));
  return { url: file.getUrl(), name: name + '.pdf', n: rows.length };
}

/* ---------- Alertes d'échéance (accueil et e-mail) */
var ALERTE_JOURS = 30;
function alertes_() {
  var out = [], now = new Date(), soon = new Date(now.getTime() + ALERTE_JOURS * 864e5);
  listDetecteurs_().forEach(function (x) {
    var nx = parseFr_(x.prochainControle), nom = [x.marque, x.ref].filter(Boolean).join(' ') || 'Détecteur';
    if (!nx) return;
    if (nx < now) out.push({ key: 'det:' + x.id + ':' + x.prochainControle, kind: 'detecteur', grave: true, text: nom + (x.serie ? ' (n° ' + x.serie + ')' : '') + ' : contrôle annuel dépassé depuis le ' + x.prochainControle + '.' });
    else if (nx <= soon) out.push({ key: 'det:' + x.id + ':' + x.prochainControle, kind: 'detecteur', text: nom + (x.serie ? ' (n° ' + x.serie + ')' : '') + ' : contrôle annuel à faire avant le ' + x.prochainControle + '.' });
  });
  boutTable_().rows.forEach(function (r) {
    var ref = String(r['Réf bouteille'] || '').trim();
    if (!ref || /rendu/i.test(String(r['Statut'] || ''))) return;
    var kind = boutKind_(r['Type']), vide = num_(r['Poids à vide (kg)']), pesee = num_(r['Dernière pesée (kg)']), reste = isFinite(vide) && isFinite(pesee) ? Math.max(0, pesee - vide) : NaN;
    if (!isFinite(reste)) return;
    if (kind === 'neuve') {
      var init = num_(r['Fluide initial (kg)']); if (!isFinite(init)) init = num_(r['Poids plein à l’achat (kg)']) - vide;
      if (isFinite(init) && init > 0 && reste <= 0.15 * init) out.push({ key: 'bneuve:' + ref + ':' + Math.round(reste * 10), kind: 'bouteille', text: 'Bouteille ' + ref + ' presque vide : ' + fr_(reste, 2) + ' kg restants. Prévoir une bouteille neuve.' });
    } else {
      var cap = num_(r['Capacité max (kg)']);
      if (isFinite(cap) && cap > 0 && reste >= 0.8 * cap) out.push({ key: 'brecup:' + ref + ':' + Math.round(reste * 10), kind: 'bouteille', grave: reste >= cap, text: 'Bouteille ' + ref + (reste >= cap ? ' pleine' : ' presque pleine') + ' : ' + fr_(reste, 2) + ' kg sur ' + fr_(cap, 2) + ' kg. À rendre au distributeur.' });
    }
  });
  return out;
}

/** À brancher une fois (fonction activerAlertesEmail) : e-mail quotidien des nouvelles alertes bouteilles et détecteurs. */
function envoyerAlertes() {
  if (!paramOn_('E-mails d’alerte')) return 'désactivé';
  var to = String(params_()['E-mail des alertes'] || '').trim();
  if (!to) to = readTable_(SH.ops).rows.filter(function (o) { return String(o['Rôle'] || '').toLowerCase() === 'admin' && String(o['Actif'] || 'oui').toLowerCase() !== 'non'; }).map(function (o) { return o['E-mail (identifiant)']; }).filter(Boolean).join(',');
  if (!to) return 'pas de destinataire';
  var props = PropertiesService.getScriptProperties(), sent = {}; try { sent = JSON.parse(props.getProperty('ALERTES_ENVOYEES') || '{}'); } catch (e) { sent = {}; }
  var now = Date.now(), week = 7 * 864e5, all = alertes_();
  var nouvelles = all.filter(function (a) { return !sent[a.key] || now - sent[a.key] > week; });
  var keep = {}; all.forEach(function (a) { if (sent[a.key]) keep[a.key] = sent[a.key]; });
  if (!nouvelles.length) { props.setProperty('ALERTES_ENVOYEES', JSON.stringify(keep)); return 'rien à envoyer'; }
  var li = function (k, titre) { var l = nouvelles.filter(function (a) { return a.kind === k; }); return l.length ? '<h3 style="margin:16px 0 6px">' + titre + '</h3><ul>' + l.map(function (a) { return '<li' + (a.grave ? ' style="color:#A3330D;font-weight:bold"' : '') + '>' + a.text + '</li>'; }).join('') + '</ul>' : ''; };
  MailApp.sendEmail({ to: to, name: 'ARE Fluides', subject: 'ARE Fluides : ' + nouvelles.length + ' alerte' + (nouvelles.length > 1 ? 's' : '') + ' (bouteilles, détecteurs)',
    htmlBody: '<div style="font-family:Arial,sans-serif;color:#1B2448">' + li('bouteille', 'Bouteilles') + li('detecteur', 'Détecteurs de fuite') + '<p style="color:#4E5670">Ces alertes restent affichées sur l’accueil de l’appli tant qu’elles ne sont pas réglées. Un rappel est renvoyé chaque semaine.</p></div>' });
  nouvelles.forEach(function (a) { keep[a.key] = now; });
  props.setProperty('ALERTES_ENVOYEES', JSON.stringify(keep));
  return nouvelles.length + ' alerte(s) envoyée(s)';
}
/** À lancer une fois depuis l'éditeur : autorise l'envoi d'e-mails et programme l'envoi quotidien des alertes (9 h). */
function activerAlertesEmail() {
  ScriptApp.getProjectTriggers().forEach(function (tr) { if (tr.getHandlerFunction() === 'envoyerAlertes') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('envoyerAlertes').timeBased().everyDays(1).atHour(9).create();
  setParam_('E-mails d’alerte', 'oui');
  Logger.log('Alertes par e-mail programmées chaque jour vers 9 h (' + MailApp.getRemainingDailyQuota() + ' e-mails disponibles aujourd’hui).');
}

/* ---------- Pesée de contrôle d'une bouteille */
function peseeControle_(d, user) {
  var f = findBout_(d.ref);
  if (!f.row) throw new Error('Bouteille introuvable.');
  var mesure = num_(d.pesee), date = String(d.date || '').trim(), dt = parseFr_(date);
  if (!isFinite(mesure) || mesure <= 0) throw new Error('Indique la pesée de la bouteille seule (kg).');
  if (!dt) throw new Error('Date au format jj/mm/aaaa.');
  if (dt > new Date()) throw new Error('La date ne peut pas être dans le futur.');
  var b = bouteilles_().filter(function (x) { return x.ref === String(f.row['Réf bouteille']).trim(); })[0];
  checkDateBouteille_(b, date);
  var attendu = num_(f.row['Dernière pesée (kg)']);
  var ecart = isFinite(attendu) ? Math.round((mesure - attendu) * 1000) / 1000 : 0;
  ensureSheet_(L4.controles, ['Date', 'Bouteille', 'Pesée (kg)', 'Poids attendu (kg)', 'Écart (kg)', 'Opérateur', 'Régularisation']);
  var t = readTable_(L4.controles);
  var conforme = Math.abs(Math.round(ecart * 1000)) <= 10;
  var r = writeRow_(t, { 'Date': date, 'Bouteille': String(f.row['Réf bouteille']).trim(), 'Pesée (kg)': mesure, 'Poids attendu (kg)': isFinite(attendu) ? attendu : '', 'Écart (kg)': ecart,
    'Opérateur': ((user.prenom || '') + ' ' + (user.nom || '')).trim(), 'Régularisation': conforme ? 'conforme' : 'à régulariser' }, null);
  var kind = boutKind_(f.row['Type']);
  // Bouteille plus légère : fluide sorti (ajout de charge) ; plus lourde : fluide entré (récupération)
  var op = ecart < 0 ? 'ajout' : 'recup';
  var possible = conforme || !isFinite(attendu) ? true : (op === 'ajout' ? kind !== 'recup' : kind !== 'neuve');
  return { id: t.sh.getLastRow(), conforme: conforme || !isFinite(attendu), attendu: isFinite(attendu) ? fr_(attendu, 2) : '', ecart: ecart, ecartTxt: (ecart > 0 ? '+' : '') + Math.round(ecart * 1000) + ' g',
    op: op, quantite: fr_(Math.abs(ecart), 2), possible: possible, ref: String(f.row['Réf bouteille']).trim(), date: date, pesee: fr_(mesure, 2) };
}
function marquerRegularisation_(id, cerfa) {
  try { var t = readTable_(L4.controles), row = t.rows.filter(function (r) { return r._row === Number(id); })[0]; if (row) writeRow_(t, { 'Régularisation': 'régularisé (' + cerfa + ')' }, row._row); } catch (e) {}
}

/* ---------- Envoi des documents au client */
function envoyerClient_(d, user) {
  if (!paramOn_('Envoi des documents au client')) throw new Error('L’envoi des documents au client est désactivé (réglage administrateur).');
  var c = getChantier_(d.chantierId);
  if (!canEditMes_(c, null, user)) throw new Error(DROITS_MSG);
  var to = String(d.email || c.email || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) throw new Error('Indique l’e-mail du client sur la fiche chantier.');
  var row = mesRows_(c.id)[0];
  if (!row) throw new Error('Pas encore de mise en service pour ce chantier.');
  var files = [], noms = [];
  var add = function (id, label) { if (!id) return; try { var fl = DriveApp.getFileById(id); files.push(fl.getBlob()); noms.push(label); } catch (e) {} };
  add(String(row['Rapport (ID)'] || ''), 'rapport d’installation');
  mouvTable_().rows.filter(function (r) { return String(r['ID MES']) === String(row['ID']); }).forEach(function (r) { add(String(r['CERFA (ID)'] || ''), 'CERFA ' + r['N° CERFA']); });
  if (!files.length) throw new Error('Aucun document à envoyer (rapport ou CERFA à refaire).');
  var p = params_();
  MailApp.sendEmail({ to: to, name: p['Raison sociale'] || 'Archenergie Renouvelable', subject: 'Documents de votre installation – ' + (c.client || ''),
    htmlBody: '<div style="font-family:Arial,sans-serif;color:#1B2448"><p>Bonjour,</p><p>Vous trouverez ci-joint les documents de votre installation' + (c.marque || c.modele ? ' (' + [c.marque, c.modele].filter(Boolean).join(' ') + ')' : '') + ' : ' + noms.join(', ') + '.</p>' +
      '<p>Le CERFA 15497 (fiche d’intervention sur les fluides frigorigènes) est à conserver avec les documents de l’équipement.</p><p>Cordialement,<br>' + ((user.prenom || '') + ' ' + (user.nom || '')).trim() + '<br>' + (p['Raison sociale'] || 'Archenergie Renouvelable') + '</p></div>',
    attachments: files });
  ensureColumn_(SH.chantiers, 'Documents envoyés');
  var t = readTable_(SH.chantiers), cr = t.rows.filter(function (r) { return r['ID'] === c.id; })[0];
  var trace = 'le ' + today_() + ' à ' + to + ' (' + noms.join(', ') + ')';
  if (cr) writeRow_(t, { 'Documents envoyés': trace }, cr._row);
  return { msg: 'Documents envoyés à ' + to + ' : ' + noms.join(', ') + '.', trace: trace };
}

/* ---------- CERFA papier : lecture de la photo, chantiers proposés, enregistrement */
function readCerfaPapier_(d) {
  if (!d.data) throw new Error('Aucune photo reçue.');
  var block = d.mime === 'application/pdf' ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: String(d.data) } }
    : { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: String(d.data) } };
  var prompt = 'Document : fiche d’intervention CERFA 15497 (équipements contenant des fluides frigorigènes), remplie à la main.\n' +
    'Réponds UNIQUEMENT par un objet JSON : {"numero": "n° de la fiche", "date": "date de l’intervention jj/mm/aaaa", "client": "nom du détenteur", "adresse": "adresse (rue)", "cp": "code postal", "ville": "ville",' +
    ' "marque": "marque de l’équipement", "modele": "modèle", "serie": "n° de série", "fluide": "fluide (ex. R32)", "chargeTotale": "charge totale en kg (nombre)",' +
    ' "nature": ["mes"|"assemblage"|"modif"|"maintenance"|"ctrlPerio"|"ctrlNonPerio"|"demantel"|"autre" : cases cochées en case 4],' +
    ' "chargee": "quantité de fluide chargée en kg (case 11, nombre ou vide)", "recuperee": "quantité de fluide récupérée en kg (nombre ou vide)", "fuite": "oui"|"non"|"", "operateur": "nom de l’opérateur"}.\n' +
    'Mets "" pour ce qui est absent ou illisible. Ne devine jamais.';
  var text = callClaude_([block, { type: 'text', text: prompt }], 800);
  var i = text.indexOf('{'), j = text.lastIndexOf('}');
  var raw = {}; try { raw = JSON.parse(text.slice(i, j + 1)); } catch (e) { throw new Error('CERFA illisible : reprends la photo bien à plat, avec toute la page.'); }
  var out = {};
  ['numero', 'date', 'client', 'adresse', 'cp', 'ville', 'marque', 'modele', 'serie', 'fluide', 'chargeTotale', 'chargee', 'recuperee', 'fuite', 'operateur'].forEach(function (k) { out[k] = raw[k] === null || raw[k] === undefined ? '' : String(raw[k]).trim().slice(0, 120); });
  out.nature = Array.isArray(raw.nature) ? raw.nature.filter(function (k) { return /^(assemblage|mes|modif|maintenance|ctrlPerio|ctrlNonPerio|demantel|autre)$/.test(k); }) : [];
  out.fluide = normFluide_(out.fluide);
  ['chargeTotale', 'chargee', 'recuperee'].forEach(function (k) { var n = num_(out[k]); out[k] = isFinite(n) && n > 0 ? fr_(n, 2) : ''; });
  if (out.date && !parseFr_(out.date)) out.date = '';
  out.propositions = chantiersProches_(out);
  return out;
}

/** Chantiers qui ressemblent au CERFA : n° de série, nom du client, adresse, code postal, modèle. */
function chantiersProches_(x) {
  var norm = function (s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, ''); };
  var nc = norm(x.client), na = norm(x.adresse), ns = normModele_(x.serie), nm = normModele_(x.modele);
  return listChantiers_().map(function (c) {
    var sc = 0, why = [];
    if (ns && ns.length >= 5 && (normModele_(c.serieUE) === ns || normModele_(c.serieUI) === ns)) { sc += 10; why.push('même n° de série'); }
    var cc = norm(c.client);
    if (nc && cc && (cc === nc || (nc.length >= 4 && (cc.indexOf(nc) >= 0 || nc.indexOf(cc) >= 0)))) { sc += 5; why.push('même client'); }
    else if (nc && cc && nc.length >= 4) {
      var parts = String(x.client).toLowerCase().split(/\s+/).filter(function (w) { return w.length >= 3; });
      if (parts.some(function (w) { return cc.indexOf(norm(w)) >= 0; })) { sc += 3; why.push('nom proche'); }
    }
    if (na && na.length >= 6 && norm(c.adresse) && (norm(c.adresse).indexOf(na) >= 0 || na.indexOf(norm(c.adresse)) >= 0)) { sc += 3; why.push('même adresse'); }
    if (x.cp && String(c.cp || '') === String(x.cp)) { sc += 1; why.push('même code postal'); }
    if (nm && nm.length >= 5 && normModele_(c.modele) === nm) { sc += 2; why.push('même modèle'); }
    return { c: c, sc: sc, why: why };
  }).filter(function (z) { return z.sc >= 3; }).sort(function (a, b) { return b.sc - a.sc; }).slice(0, 5).map(function (z) {
    return { id: z.c.id, client: z.c.client, adresse: [z.c.adresse, [z.c.cp, z.c.ville].filter(Boolean).join(' ')].filter(Boolean).join(', '), machine: [z.c.marque, z.c.modele].filter(Boolean).join(' '), dateMes: z.c.dateMes, raison: z.why.join(', ') };
  });
}

function saveCerfaPapier_(d, user) {
  var numero = String(d.numero || '').trim();
  if (!numero) throw new Error('Indique le n° du CERFA.');
  var date = String(d.date || '').trim(), dt = parseFr_(date);
  if (!dt) throw new Error('Indique la date du CERFA (jj/mm/aaaa).');
  if (dt > new Date()) throw new Error('La date du CERFA ne peut pas être dans le futur.');
  var mv = mouvTable_();
  if (mv.rows.some(function (r) { return String(r['N° CERFA']).trim().toLowerCase() === numero.toLowerCase(); })) throw new Error('Le CERFA ' + numero + ' est déjà dans le registre.');
  var op = d.op === 'ajout' || d.op === 'recup' ? d.op : 'autre';
  var quantite = 0, avant = NaN, apres = NaN, b = null;
  if (op !== 'autre') {
    quantite = num_(d.quantite);
    if (!(quantite > 0)) throw new Error('Indique la quantité ' + (op === 'ajout' ? 'chargée' : 'récupérée') + ' (kg).');
    if (!d.bouteille) throw new Error('Choisis la bouteille utilisée.');
    b = bouteilles_().filter(function (x) { return x.ref === d.bouteille; })[0];
    if (!b) throw new Error('Bouteille introuvable.');
    checkDateBouteille_(b, date);
    avant = num_(b.pesee);
    apres = num_(d.poidsSeule);
    if (!isFinite(apres) && isFinite(avant)) apres = avant + (op === 'ajout' ? -quantite : quantite);
  }
  // Chantier : existant (choisi parmi les propositions) ou créé à partir du CERFA
  var c;
  if (d.chantierId) c = getChantier_(d.chantierId);
  else {
    if (!String(d.client || '').trim()) throw new Error('Indique le nom du client (détenteur).');
    var parts = String(d.client).trim().split(/\s+/);
    c = saveChantier_({ nom: parts.length > 1 ? parts.slice(1).join(' ') : parts[0], prenom: parts.length > 1 ? parts[0] : '', adresse: d.adresse || '', cp: d.cp || '', ville: d.ville || '',
      marque: d.marque || '', modele: d.modele || '', fluide: d.fluide || '', serieUE: d.serie || '', chargeUsine: '', commentaires: 'Créé depuis le CERFA papier ' + numero }, user);
  }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var photo = null;
    if (d.data) photo = saveFile_(chantierFolder_(c), d.data, d.mime === 'application/pdf' ? 'application/pdf' : 'image/jpeg',
      aammjj_(date) + ' CERFA papier ' + String(c.client || '').replace(/[\\\/:*?"<>|]/g, ' ').trim() + ' ' + numero + (d.mime === 'application/pdf' ? '.pdf' : '.jpg'));
    var rows = mesRows_(c.id), mesRow = rows[0], opNom = ((user.prenom || '') + ' ' + (user.nom || '')).trim() || user.email;
    var nature = Array.isArray(d.nature) ? d.nature : [];
    if (!mesRow) {
      var mt = readTable_(MES.sheet);
      writeRow_(mt, { 'ID': nextId_(mt.rows.concat(readTableSafe_('_MES supprimées')), 'M', 'ID'), 'ID chantier': c.id, 'Date': date, 'Opérateur ID': user.id, 'Opérateur': opNom,
        'N° série UE': d.serie || c.serieUE || '', 'N° série UI': c.serieUI || '', 'Photos': '[]', 'Créée le': new Date(), 'Commentaire rapport': 'Mise en service saisie depuis le CERFA papier ' + numero }, null);
      mesRow = mesRows_(c.id)[0];
      var ct = readTable_(SH.chantiers), cr = ct.rows.filter(function (r) { return r['ID'] === c.id; })[0];
      if (cr && !String(cr['Date MES'] || '').trim()) writeRow_(ct, { 'Date MES': date }, cr._row);
    }
    var fl = fluideInfo_(d.fluide || c.fluide), gwp = fl ? fl.gwp : num_(c.gwp), charge = num_(d.chargeTotale);
    var data = { op: op, papier: true, cerfa: numero, date: date, quantite: quantite || '', nature: nature, fuite: d.fuite || '', operateurPapier: d.operateur || '',
      pesee: op === 'autre' ? null : { mode: 'cerfa', seule: isFinite(apres) ? apres : null, attendu: isFinite(avant) ? Math.round((avant + (op === 'ajout' ? -quantite : quantite)) * 1000) / 1000 : null } };
    writeRow_(mv, {
      'Date': date, 'Opération': op === 'ajout' ? 'Ajout de charge' : op === 'recup' ? 'Récupération' : 'Intervention', 'Cadre': 'CERFA papier',
      'Chantier / Client': c.client, 'Adresse': [c.adresse, [c.cp, c.ville].filter(Boolean).join(' ')].filter(Boolean).join(', '), 'Machine': [c.marque, c.modele].filter(Boolean).join(' '),
      'Fluide': d.fluide || c.fluide, 'Bouteille': b ? b.ref : '', 'Poids avant (kg)': isFinite(avant) ? avant : '', 'Poids après (kg)': isFinite(apres) ? apres : '',
      'Variation bouteille (kg)': quantite ? (op === 'ajout' ? -quantite : quantite) : 0,
      'Charge totale machine (kg)': isFinite(charge) ? charge : '', 't éq. CO2': isFinite(charge) && isFinite(gwp) ? Math.round(charge * gwp) / 1000 : '',
      'Opérateur': d.operateur || opNom, 'N° CERFA': numero, 'Observations': 'CERFA papier saisi par ' + opNom,
      'ID chantier': c.id, 'ID MES': String(mesRow['ID']), 'Opérateur ID': user.id, 'CERFA (lien)': photo ? photo.url : '', 'CERFA (ID)': photo ? photo.id : '', 'Données': JSON.stringify(data)
    }, null);
    if (b && isFinite(apres)) {
      var bt = readTable_(SH.bout), br = bt.rows.filter(function (r) { return String(r['Réf bouteille']).trim() === b.ref; })[0];
      if (br) {
        var bup = { 'Dernière pesée (kg)': apres }, bv = num_(br['Poids à vide (kg)']);
        if (isFinite(bv)) bup['Fluide restant (kg)'] = Math.round(Math.max(0, apres - bv) * 1000) / 1000;
        if (!String(br['Fluide'] || '').trim() && (d.fluide || c.fluide)) bup['Fluide'] = d.fluide || c.fluide;
        writeRow_(bt, bup, br._row);
      }
    }
  } finally { lock.releaseLock(); }
  return { chantier: getChantier_(c.id), msg: 'CERFA ' + numero + ' enregistré sur le chantier ' + c.client + '.' };
}
function readTableSafe_(name) { try { return readTable_(name).rows; } catch (e) { return []; } }
