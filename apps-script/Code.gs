/**
 * ARE Fluides – serveur (Google Apps Script)
 * Livraison 1 : connexion, comptes, chantiers.
 *
 * Ce script est rattaché au Google Sheets « ARE Fluides - Base ».
 * Les mots de passe sont stockés hachés (SHA-256 + sel) dans l'onglet caché « _Comptes ».
 */

var APP = { name: 'ARE Fluides', version: 'L2', sessionSeconds: 6 * 60 * 60, hashRounds: 300 };
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
    req = req || {};
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
      updateMes: function (d) { return updateMes_(d, user); },
      deleteManip: function (d) { return deleteManip_(d, user); },
      deleteMes: function (d) { return deleteMes_(d, user); },
      changePassword: function (d) { return changePassword_(user, d); },
      listOperateurs: function () { admin_(user); return listOperateurs_(); },
      saveOperateur: function (d) { admin_(user); return saveOperateur_(d, user); },
      setOperateurActif: function (d) { admin_(user); return setOperateurActif_(d, user); },
      deleteOperateur: function (d) { admin_(user); return deleteOperateur_(d, user); },
      disconnectOperateur: function (d) { admin_(user); return disconnectOperateur_(d, user); }
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
    attestationChargee: String(o['Attestation chargée'] || '').toLowerCase().indexOf('oui') === 0,
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
  o.client = (o.prenom ? o.prenom + ' ' : '') + o.nom;
  o.misEnService = !!o.dateMes;
  return o;
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
    var t = readTable_(SH.chantiers);
    var row = d.id ? t.rows.filter(function (r) { return r['ID'] === d.id; })[0] : null;
    if (d.id && !row) throw new Error('Chantier introuvable.');
    var rec = {};
    Object.keys(CH_FIELDS).forEach(function (k) { if (k !== 'id' && d[k] !== undefined) rec[CH_FIELDS[k]] = d[k]; });
    if (!row) {
      rec['ID'] = nextId_(t.rows.concat(corbeille_().rows), 'C', 'ID');   // jamais réutiliser l'ID d'un chantier supprimé
      if (user) rec['Créé par'] = user.id;
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
    return getChantier_(id);
  } finally { lock.releaseLock(); }
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
    if (user.role !== 'admin' && String(row['Créé par'] || '') !== user.id) throw new Error('Tu ne peux supprimer que les chantiers que tu as créés.');
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
  headers: ['ID', 'ID chantier', 'Date', 'Opérateur ID', 'Opérateur', 'N° série UE', 'N° série UI', 'Photos', 'Rapport (lien)', 'Rapport (ID)', 'Signature (ID)', 'Créée le', 'Modifiée le'],
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

/** Droits sur une mise en service : admin, créateur du chantier ou opérateur qui l'a faite. */
function canEditMes_(c, mesRow, user) {
  if (user.role === 'admin') return true;
  if (c && c.creePar === user.id) return true;
  return !!(mesRow && String(mesRow['Opérateur ID']) === user.id);
}

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

function bouteilles_() {
  var t; try { t = readTable_(SH.bout); } catch (e) { return []; }
  return t.rows.filter(function (r) { return String(r['Réf bouteille'] || '').trim(); }).map(function (r) {
    var vide = num_(r['Poids à vide (kg)']), pesee = num_(r['Dernière pesée (kg)']);
    return {
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
    return { id: String(r['ID']), nom: String(r['Nom']), siret: String(r['SIRET'] || ''), adresse: String(r['Adresse du site de reprise'] || '') };
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
    lien: String(r['CERFA (lien)'] || ''), operateur: String(r['Opérateur'] || '')
  };
}

function mesPublic_(row, c, user) {
  if (!row) return null;
  var manips = mouvTable_().rows.filter(function (r) { return String(r['ID MES']) === String(row['ID']); }).map(manipPublic_);
  var photos = []; try { photos = JSON.parse(row['Photos'] || '[]'); } catch (e) { photos = []; }
  return {
    id: String(row['ID']), date: String(row['Date'] || ''), operateur: String(row['Opérateur'] || ''), operateurId: String(row['Opérateur ID'] || ''),
    serieUE: String(row['N° série UE'] || ''), serieUI: String(row['N° série UI'] || ''), photos: photos.length,
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
    if (row && !canEditMes_(c, row, user)) throw new Error('Seul l’opérateur qui a fait la mise en service, le créateur du chantier ou un administrateur peut la compléter.');
    if (row && d.op === 'mes') throw new Error('La mise en service est déjà enregistrée : choisis une manipulation de fluide.');

    var fl = fluideInfo_(c.fluide), gwp = fl ? fl.gwp : num_(c.gwp);
    var machine = c.modele ? findMachine_(c.modele) : null;
    var chargeUsine = num_(d.chargeUsine);
    if (!isFinite(chargeUsine)) chargeUsine = num_(c.chargeUsine);
    if (!isFinite(chargeUsine) && machine) chargeUsine = num_(machine.charge);
    var avant = num_(d.poidsAvant), apres = num_(d.poidsApres), quantite = 0;
    if (d.op === 'ajout' || d.op === 'recup') {
      if (!d.bouteille) throw new Error('Choisis la bouteille utilisée.');
      if (!isFinite(avant) || !isFinite(apres)) throw new Error('Saisis les deux pesées de la bouteille.');
      quantite = d.op === 'ajout' ? avant - apres : apres - avant;
      if (!(quantite > 0)) throw new Error(d.op === 'ajout' ? 'La bouteille doit être plus légère après la charge.' : 'La bouteille doit être plus lourde après la récupération.');
    }
    var detecteur = null;
    if (d.op !== 'mes') {
      detecteur = outils_().filter(function (o) { return o.id === d.detecteur; })[0] || null;
      if (!detecteur) throw new Error('Choisis le détecteur de fuite utilisé.');
      if (!detecteur.ok) throw new Error('Le contrôle de ce détecteur est dépassé : il ne peut pas être utilisé.');
    }
    var distributeur = d.op === 'recup' ? (distributeurs_().filter(function (x) { return x.id === d.distributeur; })[0] || null) : null;
    if (d.op === 'recup' && !distributeur) throw new Error('Choisis la destination du fluide récupéré (case 13).');

    var now = new Date(), opNom = (user.prenom + ' ' + user.nom).trim() || user.email;
    // En-tête de la mise en service (créé à la première opération)
    if (!row) {
      var rec = { 'ID': nextId_(mt.rows.concat(ss_().getSheetByName(MES.trashMes) ? readTable_(MES.trashMes).rows : []), 'M', 'ID'), 'ID chantier': c.id, 'Date': date,
        'Opérateur ID': user.id, 'Opérateur': opNom, 'N° série UE': d.serieUE || '', 'N° série UI': d.serieUI || '', 'Photos': '[]', 'Créée le': now };
      writeRow_(mt, rec, null);
      mt = readTable_(MES.sheet);
      row = mt.rows.filter(function (r) { return String(r['ID chantier']) === c.id; })[0];
    } else {
      var up = { 'Modifiée le': now };
      if (d.serieUE) up['N° série UE'] = d.serieUE;
      if (d.serieUI) up['N° série UI'] = d.serieUI;
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
    if (d.serieUE) cup['N° série UE'] = d.serieUE;
    if (d.serieUI) cup['N° série UI'] = d.serieUI;
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
      longueur: c.longueur, precharge: c.precharge, gpm: c.gpm, observations: d.observations || '',
      operateur: opNom, attestation: user.attestation || ''
    };
    m.observationsCerfa = observations_(m);
    writeRow_(mv, {
      'Date': date, 'Opération': d.op === 'relache' ? 'Mise en service' : MES.ops[d.op], 'Cadre': 'Mise en service',
      'Chantier / Client': c.client, 'Adresse': m.adresse, 'Machine': m.machine, 'Fluide': c.fluide, 'Bouteille': m.bouteille,
      'Poids avant (kg)': isFinite(avant) ? avant : '', 'Poids après (kg)': isFinite(apres) ? apres : '',
      'Variation bouteille (kg)': isFinite(avant) && isFinite(apres) ? Math.round((apres - avant) * 1000) / 1000 : 0,
      'Charge totale machine (kg)': Math.round(chargeTotale * 1000) / 1000,
      't éq. CO2': isFinite(gwp) ? Math.round(chargeTotale * gwp) / 1000 : '',
      'Opérateur': opNom, 'N° CERFA': cerfa, 'Observations': m.observationsCerfa,
      'ID chantier': c.id, 'ID MES': String(row['ID']), 'Opérateur ID': user.id, 'Données': JSON.stringify(m)
    }, null);
    // Stock : nouvelle pesée de la bouteille
    if (m.bouteille && isFinite(apres)) {
      var bt = readTable_(SH.bout), br = bt.rows.filter(function (r) { return String(r['Réf bouteille']).trim() === m.bouteille; })[0];
      if (br) writeRow_(bt, { 'Dernière pesée (kg)': apres }, br._row);
    }
    out.cerfa = cerfa;
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
  if (String(f.row['Opérateur ID']) !== user.id && !canEditMes_(c, mesRow, user)) throw new Error('Tu ne peux pas modifier cette fiche.');
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
  if (!canEditMes_(c, mesRow, user)) throw new Error('Tu ne peux pas modifier cette fiche.');
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

/* ---------- Modification et suppression */

function updateMes_(d, user) {
  var c = getChantier_(d.chantierId);
  var row = mesRows_(c.id)[0];
  if (!row) throw new Error('Aucune mise en service sur ce chantier.');
  if (!canEditMes_(c, row, user)) throw new Error('Seul l’opérateur qui a fait la mise en service, le créateur du chantier ou un administrateur peut la modifier.');
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
    if (br && num_(br['Dernière pesée (kg)']) === num_(r['Poids après (kg)']) && isFinite(num_(r['Poids avant (kg)']))) writeRow_(bt, { 'Dernière pesée (kg)': num_(r['Poids avant (kg)']) }, br._row);
  }
  trashFile_(String(r['CERFA (ID)'] || ''));
  archiveRow_(MES.trashMouv, mv, r, user);
  mv.sh.deleteRow(r._row);
}

function deleteManip_(d, user) {
  var f = findMouv_(d.cerfa);
  var c = getChantier_(String(f.row['ID chantier']));
  var row = mesRows_(c.id)[0];
  if (!canEditMes_(c, row, user)) throw new Error('Seul l’opérateur qui a fait la mise en service, le créateur du chantier ou un administrateur peut la modifier.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { deleteManipRow_(f.t, f.row, user); } finally { lock.releaseLock(); }
  return true;
}

function deleteMes_(d, user) {
  var c = getChantier_(d.chantierId);
  var row = mesRows_(c.id)[0];
  if (!row) throw new Error('Aucune mise en service sur ce chantier.');
  if (!canEditMes_(c, row, user)) throw new Error('Seul l’opérateur qui a fait la mise en service, le créateur du chantier ou un administrateur peut la supprimer.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    // Les opérations d'abord, de bas en haut pour garder les numéros de ligne
    var mv = mouvTable_();
    mv.rows.filter(function (r) { return String(r['ID MES']) === String(row['ID']); }).reverse().forEach(function (r) { deleteManipRow_(mv, r, user); });
    trashFile_(String(row['Rapport (ID)'] || ''));
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
