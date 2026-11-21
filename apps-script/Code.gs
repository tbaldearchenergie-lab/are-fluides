/**
 * ARE Fluides – serveur (Google Apps Script)
 * Livraison 1 : connexion, comptes, chantiers.
 *
 * Ce script est rattaché au Google Sheets « ARE Fluides - Base ».
 * Les mots de passe sont stockés hachés (SHA-256 + sel) dans l'onglet caché « _Comptes ».
 */

var APP = { name: 'ARE Fluides', version: 'L1', sessionSeconds: 6 * 60 * 60, hashRounds: 300 };
var SH = {
  params: 'Paramètres', ops: 'Opérateurs', chantiers: 'Chantiers', mouv: 'Mouvements', bout: 'Bouteilles',
  outils: 'Outillage', dist: 'Distributeurs', fluides: 'Fluides', comptes: '_Comptes'
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
      saveChantier: function (d) { return saveChantier_(d); },
      listFluides: listFluides_,
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
  if (cut && (s.at || 0) < cut) { CacheService.getScriptCache().remove('s_' + token); throw new Error('SESSION'); }
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
  inflam: 'Inflammable (si fluide inconnu)', longueur: 'Longueur liaisons (m)', precharge: 'Précharge (m)', gpm: 'Charge (g/m)',
  commentaires: 'Commentaires'
};

function chantierFromRow_(r) {
  var o = {};
  Object.keys(CH_FIELDS).forEach(function (k) { o[k] = r[CH_FIELDS[k]] === undefined ? '' : String(r[CH_FIELDS[k]]); });
  o.gwp = r['GWP'] === '' || r['GWP'] === 'saisir' ? '' : String(r['GWP']);
  o.client = (o.prenom ? o.prenom + ' ' : '') + o.nom;
  o.misEnService = !!o.dateMes;
  return o;
}

function listChantiers_() {
  return readTable_(SH.chantiers).rows.map(chantierFromRow_).reverse();
}

function getChantier_(id) {
  var r = readTable_(SH.chantiers).rows.filter(function (x) { return x['ID'] === id; })[0];
  if (!r) throw new Error('Chantier introuvable.');
  return chantierFromRow_(r);
}

function saveChantier_(d) {
  if (!String(d.nom || '').trim()) throw new Error('Le nom du client est obligatoire.');
  if (d.dateMes && !/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(String(d.dateMes).trim())) throw new Error('Date de mise en service au format jj/mm/aaaa.');
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readTable_(SH.chantiers);
    var row = d.id ? t.rows.filter(function (r) { return r['ID'] === d.id; })[0] : null;
    if (d.id && !row) throw new Error('Chantier introuvable.');
    var rec = {};
    Object.keys(CH_FIELDS).forEach(function (k) { if (k !== 'id' && d[k] !== undefined) rec[CH_FIELDS[k]] = d[k]; });
    if (!row) rec['ID'] = nextId_(t.rows, 'C', 'ID');
    var r = writeRow_(t, rec, row ? row._row : null);
    // GWP : valeur saisie à la main si fournie, sinon formule qui lit l'onglet Fluides
    var gwpCol = t.headers.indexOf('GWP') + 1;
    var fluideCol = t.headers.indexOf('Fluide') + 1;
    if (gwpCol > 0 && fluideCol > 0) {
      var manual = Number(String(d.gwpManuel || '').replace(',', '.'));
      if (d.gwpManuel && isFinite(manual)) {
        t.sh.getRange(r, gwpCol).setValue(manual);
      } else {
        var L = columnLetter_(fluideCol);
        t.sh.getRange(r, gwpCol).setFormula('=IF(' + L + r + '="","",IFERROR(INDEX(Fluides!$B$2:$B$100,MATCH(' + L + r + ',Fluides!$A$2:$A$100,0)),"saisir"))');
      }
    }
    SpreadsheetApp.flush();
    var id = row ? d.id : rec['ID'];
    return getChantier_(id);
  } finally { lock.releaseLock(); }
}

/* ------------------------------------------------------------------ Fluides */

function listFluides_() {
  return readTable_(SH.fluides).rows.map(function (f) {
    return { code: String(f['Code']), gwp: Number(f['GWP']) || 0, classe: f['Classe de sécurité'], inflammable: String(f['Inflammable']).toLowerCase() === 'oui' };
  });
}
