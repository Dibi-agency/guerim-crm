/**
 * Guérim CRM : backend Google Apps Script
 * v2 : passerelles Google Agenda, Gmail et Outlook.
 * v3 : connexion par mot de passe propre au CRM, invitations vers Outlook et Gmail.
 * v3.2 : comptes individuels (identifiant + mot de passe), traçabilité des modifications.
 * v3.3 : chargement allégé (photos à part), synchronisation agenda en arrière-plan.
 * v3.4 : espace administrateur (état du système, tests, journal d'activité).
 * v3.5 : chaque utilisateur choisit son mot de passe à la première connexion.
 * v3.6 : mot de passe oublié (code par e-mail), e-mail de récupération par compte.
 * v3.7 : point d'accès doPost pour l'application installable (PWA hébergée sur GitHub Pages).
 * v3.8 : option « Rester connecté 30 jours » sur l'appareil.
 * v3.9 : posts réseaux sociaux planifiés dans l'agenda.
 * v3.10 : gestion des comptes réservée aux administrateurs.
 * Données stockées dans un Google Sheets privé, créé automatiquement
 * dans le Drive du compte qui déploie l'application.
 */

const APP_VERSION = '3.10.0 (2026-10-09)';

const SHEET_PROFILS = 'Profils';
const SHEET_ECHANGES = 'Echanges';
const SHEET_JOURNAL = 'Journal';
const JOURNAL_FIELDS = ['date', 'niveau', 'utilisateur', 'action', 'detail'];
const JOURNAL_MAX = 3000; // au-delà, les plus anciennes lignes sont effacées

// Identifiants qui voient l'espace Administration. Modifiable uniquement ici, jamais depuis l'interface.
const ADMIN_LOGINS = ['demphis', 'caroline'];

const PROFIL_FIELDS = [
  'id', 'createdAt', 'updatedAt',
  'prenom', 'nom', 'sexe', 'dateNaissance', 'nationalite', 'ville',
  'telephone', 'email',
  'consentement', 'dateConsentement',
  'courant', 'courantAutre', 'statutConversion', 'dateConversion', 'beitDin',
  'source', 'tags', 'notes', 'etape', 'motifNonRetenu',
  'datePreEntretien', 'heurePreEntretien', 'dateEnregistrement', 'heureEnregistrement', 'lieuEnregistrement',
  'numeroEpisode', 'titreEpisode', 'dateDiffusionRadio', 'dateMiseEnLigne', 'lienEpisode',
  'photo', 'syncIds', 'modifiePar'
];

const ECHANGE_FIELDS = [
  'id', 'profilId', 'createdAt', 'date', 'canal', 'resume',
  'prochaineAction', 'dateRelance', 'faite', 'syncIds', 'auteur'
];

const SHEET_POSTS = 'Posts';
const POST_FIELDS = [
  'id', 'createdAt', 'updatedAt', 'date', 'heure', 'reseaux', 'sujet', 'texte',
  'profilId', 'statut', 'auteur', 'syncIds'
];

// Données révélant une conviction religieuse (RGPD, art. 9) : jamais stockées sans consentement,
// jamais envoyées vers un agenda.
const SENSITIVE_FIELDS = ['courant', 'courantAutre', 'statutConversion', 'dateConversion', 'beitDin'];

// Champs conservés après retrait du consentement : uniquement l'historique de production.
const KEEP_ON_WITHDRAW = [
  'id', 'createdAt', 'etape', 'datePreEntretien', 'dateEnregistrement', 'lieuEnregistrement',
  'numeroEpisode', 'titreEpisode', 'dateDiffusionRadio', 'dateMiseEnLigne', 'lienEpisode'
];

const MAX_PHOTO_CHARS = 45000; // une cellule Sheets plafonne à 50 000 caractères

const DEFAULT_SETTINGS = {
  google: 'oui',            // agenda interne « Guérim » du compte qui héberge le CRM
  calName: 'Guérim',
  inviteOutlook: 'non',     // invitations vers l'agenda Outlook de Caroline
  outlookEmail: '',
  inviteGmail: 'non',       // invitations vers l'agenda Google de Caroline
  gmailEmail: '',
  relances: 'oui',
  posts: 'oui',             // posts réseaux envoyés vers les agendas
  expediteur: 'Podcast Guérim',
  replyTo: '',              // les réponses des invités arrivent dans sa vraie boîte
  signature: 'Caroline Amouyal\nPodcast Guérim'
};

const SESSION_SECONDS = 21600;  // 6 h, prolongées à chaque action
const MAX_FAILS = 5;            // tentatives avant blocage
const LOCK_SECONDS = 900;       // blocage de 15 min
const MIN_PASSWORD = 10;
const LONG_SESSION_DAYS = 30; // option « Rester connecté »

/* ---------- Web app ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Guérim CRM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Point d'accès de l'application installable (PWA hébergée hors de Google).
 * Reçoit { fn, args } en JSON et renvoie { ok, result } ou { ok: false, error }.
 * Seules les fonctions listées dans remoteApi_ sont appelables ; chacune vérifie elle-même la session.
 */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = remoteApi_()[req.fn];
    if (!fn) throw new Error('Action inconnue.');
    out = { ok: true, result: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function remoteApi_() {
  return {
    api_status: api_status, api_setup: api_setup, api_login: api_login, api_logout: api_logout,
    api_changePassword: api_changePassword, api_forgot: api_forgot, api_resetWithCode: api_resetWithCode,
    api_setMyEmail: api_setMyEmail, api_listUsers: api_listUsers, api_addUser: api_addUser,
    api_setUserActive: api_setUserActive, api_resetUserPassword: api_resetUserPassword, api_deleteUser: api_deleteUser,
    api_getAll: api_getAll, api_getPhotos: api_getPhotos, api_saveProfil: api_saveProfil, api_syncProfil: api_syncProfil,
    api_deleteProfil: api_deleteProfil, api_retirerConsentement: api_retirerConsentement,
    api_saveEchange: api_saveEchange, api_syncEchange: api_syncEchange, api_deleteEchange: api_deleteEchange,
    api_savePost: api_savePost, api_syncPost: api_syncPost, api_deletePost: api_deletePost,
    api_sendMail: api_sendMail, api_saveSettings: api_saveSettings, api_resyncAll: api_resyncAll,
    api_adminStatus: api_adminStatus, api_adminTest: api_adminTest
  };
}

/** À lancer une fois depuis l'éditeur : crée le classeur, l'agenda, et déclenche les autorisations. */
function installer() {
  const ss = getSpreadsheet_();
  getSheet_(SHEET_PROFILS, PROFIL_FIELDS);
  getSheet_(SHEET_ECHANGES, ECHANGE_FIELDS);
  const cal = getCalendar_();
  GmailApp.getAliases(); // force l'autorisation Gmail
  Logger.log('Classeur : ' + ss.getUrl());
  Logger.log('Agenda : ' + cal.getName());
  if (!Object.keys(users_()).length) {
    Logger.log("Code d'installation, à saisir à la première ouverture de l'application : " + setupCode_(true));
  }
  Logger.log('Service Gmail avancé : ' + (gmailAdvanced_() ? 'actif' : 'inactif, les invitations Outlook partiront en pièce jointe simple'));
}

/* ---------- API appelée par le front ---------- */

function api_getAll(token) {
  const me = guard_(token);
  return {
    version: APP_VERSION,
    settings: getSettings_(),
    gmailAdvanced: gmailAdvanced_(),
    me: me,
    profils: readAll_(SHEET_PROFILS, PROFIL_FIELDS, 'photo'), // photos envoyées ensuite par api_getPhotos
    echanges: readAll_(SHEET_ECHANGES, ECHANGE_FIELDS),
    posts: readAll_(SHEET_POSTS, POST_FIELDS)
  };
}

function api_saveProfil(token, p) {
  const me = guard_(token);
  return withLock_(function () {
    const keepPhoto = !p || p.photo === undefined || p.photo === '__KEEP__';
    const clean = pick_(p, PROFIL_FIELDS);
    const now = new Date().toISOString();
    const prev = clean.id ? findObj_(SHEET_PROFILS, PROFIL_FIELDS, clean.id) : null;
    if (keepPhoto) clean.photo = prev ? prev.photo : '';
    if (!clean.id) { clean.id = Utilities.getUuid(); clean.createdAt = now; }
    clean.updatedAt = now;
    clean.syncIds = prev ? prev.syncIds : ''; // jamais pris du client
    clean.modifiePar = me.name;
    if (!clean.etape) clean.etape = 'contact';
    if (clean.consentement !== 'oui' && clean.consentement !== 'retire') clean.consentement = 'non';
    if (clean.consentement !== 'oui') SENSITIVE_FIELDS.forEach(function (f) { clean[f] = ''; });
    if (clean.courant !== 'autre') clean.courantAutre = '';
    if (!validTime_(clean.heurePreEntretien)) clean.heurePreEntretien = '';
    if (!validTime_(clean.heureEnregistrement)) clean.heureEnregistrement = '';
    if (clean.photo && (clean.photo.indexOf('data:image/') !== 0 || clean.photo.length > MAX_PHOTO_CHARS)) clean.photo = '';
    upsert_(SHEET_PROFILS, PROFIL_FIELDS, clean);
    if (!prev) log_('info', me.name, 'Fiche créée', shortName_(clean));
    return clean; // la synchronisation agenda suit via api_syncProfil
  });
}

function api_deleteProfil(token, id) {
  const me = guard_(token);
  return withLock_(function () {
    const p = findObj_(SHEET_PROFILS, PROFIL_FIELDS, id);
    const errors = [];
    if (p) errors.push.apply(errors, syncItems_(p.syncIds, {}, p.id, false).errors);
    errors.push.apply(errors, clearEchangesOf_(id));
    errors.push.apply(errors, detachPostsOf_(id));
    deleteById_(SHEET_PROFILS, PROFIL_FIELDS, id);
    log_(errors.length ? 'erreur' : 'info', me.name, 'Fiche supprimée', (p ? shortName_(p) : id) + (errors.length ? ' / ' + errors.join(' ') : ''));
    return { ok: true, _warn: errors.join(' ') };
  });
}

function api_retirerConsentement(token, id) {
  const me = guard_(token);
  return withLock_(function () {
    const current = findObj_(SHEET_PROFILS, PROFIL_FIELDS, id);
    if (!current) throw new Error('Profil introuvable.');
    const errors = syncItems_(current.syncIds, {}, id, false).errors; // annule tous les rendez-vous
    errors.push.apply(errors, clearEchangesOf_(id));
    errors.push.apply(errors, detachPostsOf_(id));
    const kept = {};
    PROFIL_FIELDS.forEach(function (f) { kept[f] = KEEP_ON_WITHDRAW.indexOf(f) > -1 ? current[f] : ''; });
    kept.prenom = 'Profil';
    kept.nom = 'retiré';
    kept.consentement = 'retire';
    kept.dateConsentement = today_();
    kept.updatedAt = new Date().toISOString();
    kept.modifiePar = me.name;
    upsert_(SHEET_PROFILS, PROFIL_FIELDS, kept);
    log_(errors.length ? 'erreur' : 'info', me.name, 'Consentement retiré', shortName_(current) + (errors.length ? ' / ' + errors.join(' ') : ''));
    kept._warn = errors.join(' ');
    return kept;
  });
}

function api_saveEchange(token, e) {
  const me = guard_(token);
  return withLock_(function () {
    const clean = pick_(e, ECHANGE_FIELDS);
    if (!clean.profilId) throw new Error('Échange sans profil.');
    const prev = clean.id ? findObj_(SHEET_ECHANGES, ECHANGE_FIELDS, clean.id) : null;
    if (!clean.id) { clean.id = Utilities.getUuid(); clean.createdAt = new Date().toISOString(); }
    clean.syncIds = prev ? prev.syncIds : '';
    clean.auteur = prev ? prev.auteur : me.name;
    if (clean.faite !== 'oui') clean.faite = 'non';
    upsert_(SHEET_ECHANGES, ECHANGE_FIELDS, clean);
    return clean; // la synchronisation agenda suit via api_syncEchange
  });
}

function api_deleteEchange(token, id) {
  const me = guard_(token);
  return withLock_(function () {
    const e = findObj_(SHEET_ECHANGES, ECHANGE_FIELDS, id);
    const errors = e ? syncItems_(e.syncIds, {}, e.id, false).errors : [];
    deleteById_(SHEET_ECHANGES, ECHANGE_FIELDS, id);
    return { ok: true, _warn: errors.join(' ') };
  });
}

/** Photos chargées après l'affichage, pour ne pas ralentir l'ouverture. */
function api_getPhotos(token) {
  guard_(token);
  const sh = getSheet_(SHEET_PROFILS, PROFIL_FIELDS);
  const header = HEADERS_[SHEET_PROFILS];
  const last = sh.getLastRow();
  const out = {};
  if (last < 2) return out;
  const ids = sh.getRange(2, header.indexOf('id') + 1, last - 1, 1).getDisplayValues();
  const photos = sh.getRange(2, header.indexOf('photo') + 1, last - 1, 1).getDisplayValues();
  ids.forEach(function (r, i) { if (r[0] && photos[i][0]) out[r[0]] = photos[i][0]; });
  return out;
}

/** Synchronisation agenda d'une fiche, appelée en arrière-plan juste après l'enregistrement. */
function api_syncProfil(token, id) {
  const me = guard_(token);
  return withLock_(function () {
    const p = findObj_(SHEET_PROFILS, PROFIL_FIELDS, id);
    if (!p) return { syncIds: '', _warn: '' };
    const res = syncItems_(p.syncIds, profilEvents_(p), p.id, false);
    if (res.raw !== p.syncIds) { p.syncIds = res.raw; upsert_(SHEET_PROFILS, PROFIL_FIELDS, p); }
    if (res.errors.length) log_('erreur', me.name, 'Synchronisation agenda', shortName_(p) + ' / ' + res.errors.join(' '));
    return { syncIds: p.syncIds, _warn: res.errors.join(' ') };
  });
}

function api_syncEchange(token, id) {
  const me = guard_(token);
  return withLock_(function () {
    const e = findObj_(SHEET_ECHANGES, ECHANGE_FIELDS, id);
    if (!e) return { syncIds: '', _warn: '' };
    const p = findObj_(SHEET_PROFILS, PROFIL_FIELDS, e.profilId);
    const res = syncItems_(e.syncIds, echangeEvents_(e, p), e.id, false);
    if (res.raw !== e.syncIds) { e.syncIds = res.raw; upsert_(SHEET_ECHANGES, ECHANGE_FIELDS, e); }
    if (res.errors.length) log_('erreur', me.name, 'Synchronisation relance', (p ? shortName_(p) : '') + ' / ' + res.errors.join(' '));
    return { syncIds: e.syncIds, _warn: res.errors.join(' ') };
  });
}

/** Envoie un mail depuis Gmail à la personne de la fiche et l'inscrit dans le journal des échanges. */
function api_sendMail(token, profilId, subject, body, joinInvite) {
  const me = guard_(token);
  return withLock_(function () {
    const p = findObj_(SHEET_PROFILS, PROFIL_FIELDS, profilId);
    if (!p) throw new Error('Profil introuvable.');
    if (!validEmail_(p.email)) throw new Error("Cette fiche n'a pas d'adresse mail valide.");
    subject = String(subject || '').trim();
    body = String(body || '').trim();
    if (!subject || !body) throw new Error("L'objet et le message sont obligatoires.");
    const s = getSettings_();

    let joined = false;
    const enr = joinInvite ? profilEvents_(p).enr : null;
    try {
    if (enr) {
      // Invitation propre à l'invité : titre neutre, aucun lien vers le CRM.
      const d = { date: enr.date, heure: enr.heure, dur: enr.dur, title: 'Enregistrement du podcast Guérim', desc: '', lieu: enr.lieu };
      const ics = buildIcs_(d, p.id + '-enr-invite@guerim-crm', Math.floor(Date.now() / 1000), 'REQUEST', p.email);
      sendIcsMail_(p.email, subject, body, ics, 'REQUEST');
      joined = true;
    } else {
      const o = { name: s.expediteur };
      if (validEmail_(s.replyTo)) o.replyTo = s.replyTo;
      GmailApp.sendEmail(p.email, subject, body, o);
    }
    } catch (err) {
      log_('erreur', me.name, 'Mail non envoyé', shortName_(p) + ' / ' + err.message);
      throw err;
    }
    log_('info', me.name, 'Mail envoyé', shortName_(p) + (joined ? ' (avec invitation)' : ''));

    const e = {
      id: Utilities.getUuid(), profilId: profilId, createdAt: new Date().toISOString(), date: today_(),
      canal: 'mail', resume: 'Mail envoyé : ' + subject + (joined ? ' (invitation agenda jointe)' : '') + '\n\n' + body,
      prochaineAction: '', dateRelance: '', faite: 'non', syncIds: '', auteur: me.name
    };
    upsert_(SHEET_ECHANGES, ECHANGE_FIELDS, e);
    return e;
  });
}

function api_saveSettings(token, input) {
  const me = guard_(token);
  const s = getSettings_();
  const next = {
    google: input.google === 'oui' ? 'oui' : 'non',
    calName: String(input.calName || '').trim() || DEFAULT_SETTINGS.calName,
    inviteOutlook: input.inviteOutlook === 'oui' ? 'oui' : 'non',
    outlookEmail: String(input.outlookEmail || '').trim(),
    inviteGmail: input.inviteGmail === 'oui' ? 'oui' : 'non',
    gmailEmail: String(input.gmailEmail || '').trim(),
    relances: input.relances === 'oui' ? 'oui' : 'non',
    posts: input.posts === 'non' ? 'non' : 'oui',
    expediteur: String(input.expediteur || '').trim() || DEFAULT_SETTINGS.expediteur,
    replyTo: String(input.replyTo || '').trim(),
    signature: String(input.signature || '').trim()
  };
  if (next.inviteOutlook === 'oui' && !validEmail_(next.outlookEmail)) throw new Error("L'adresse Outlook n'est pas valide.");
  if (next.inviteGmail === 'oui' && !validEmail_(next.gmailEmail)) throw new Error("L'adresse Gmail n'est pas valide.");
  if (next.replyTo && !validEmail_(next.replyTo)) throw new Error("L'adresse de réponse n'est pas valide.");
  PropertiesService.getScriptProperties().setProperty('SETTINGS', JSON.stringify(next));
  if (next.calName !== s.calName && next.google === 'oui') {
    try { getCalendar_().setName(next.calName); } catch (err) { /* agenda recréé au besoin */ }
  }
  const res = resyncAll_(false);
  log_(res.errors.length ? 'erreur' : 'info', me.name, 'Réglages enregistrés', res.count + ' élément(s) synchronisé(s)' + (res.errors.length ? ' / ' + res.errors.join(' ') : ''));
  return { settings: next, count: res.count, errors: res.errors };
}

function api_resyncAll(token, force) {
  const me = guard_(token);
  const res = resyncAll_(force);
  log_(res.errors.length ? 'erreur' : 'info', me.name, 'Resynchronisation complète', res.count + ' élément(s)' + (res.errors.length ? ' / ' + res.errors.join(' ') : ''));
  return res;
}

/** Recalcule tous les rendez-vous. force = true reconstruit tout, même ce qui n'a pas changé. */
function resyncAll_(force) {
  return withLock_(function () {
    const errors = [];
    let count = 0;
    const profils = readAll_(SHEET_PROFILS, PROFIL_FIELDS);
    const byId = {};
    profils.forEach(function (p) {
      byId[p.id] = p;
      const res = syncItems_(p.syncIds, profilEvents_(p), p.id, !!force);
      if (res.raw !== p.syncIds) { p.syncIds = res.raw; upsert_(SHEET_PROFILS, PROFIL_FIELDS, p); count++; }
      errors.push.apply(errors, res.errors);
    });
    readAll_(SHEET_ECHANGES, ECHANGE_FIELDS).forEach(function (e) {
      const res = syncItems_(e.syncIds, echangeEvents_(e, byId[e.profilId]), e.id, !!force);
      if (res.raw !== e.syncIds) { e.syncIds = res.raw; upsert_(SHEET_ECHANGES, ECHANGE_FIELDS, e); count++; }
      errors.push.apply(errors, res.errors);
    });
    readAll_(SHEET_POSTS, POST_FIELDS).forEach(function (t) {
      const res = syncItems_(t.syncIds, postEvents_(t, byId[t.profilId]), t.id, !!force);
      if (res.raw !== t.syncIds) { t.syncIds = res.raw; upsert_(SHEET_POSTS, POST_FIELDS, t); count++; }
      errors.push.apply(errors, res.errors);
    });
    return { count: count, errors: unique_(errors) };
  });
}

/* ---------- Administration ---------- */

function adminGuard_(token) {
  const me = guard_(token);
  if (!me.admin) throw new Error("Réservé à l'administrateur.");
  return me;
}

function api_adminStatus(token) {
  adminGuard_(token);
  const t0 = Date.now();
  const ss = getSpreadsheet_();
  const count = function (name) { const sh = ss.getSheetByName(name); return sh ? Math.max(sh.getLastRow() - 1, 0) : 0; };
  let quota = null;
  try { quota = MailApp.getRemainingDailyQuota(); } catch (err) { quota = null; }
  let agenda = '';
  try {
    const id = PropertiesService.getScriptProperties().getProperty('CAL_ID');
    const cal = id ? CalendarApp.getCalendarById(id) : null;
    agenda = cal ? cal.getName() : '';
  } catch (err) { agenda = ''; }
  const journal = readJournal_(300);
  const since = Date.now() - 86400000;
  const recent = journal.filter(function (j) { return new Date(j.date).getTime() >= since; });
  const s = getSettings_();
  return {
    version: APP_VERSION,
    compte: Session.getEffectiveUser().getEmail(),
    classeurUrl: ss.getUrl(),
    fuseau: Session.getScriptTimeZone(),
    profils: count(SHEET_PROFILS),
    echanges: count(SHEET_ECHANGES),
    utilisateurs: Object.keys(users_()).length,
    quotaMail: quota,
    gmailAdvanced: gmailAdvanced_(),
    agenda: agenda,
    agendaInterne: s.google === 'oui',
    invitations: inviteTargets_(s),
    erreurs24h: recent.filter(function (j) { return j.niveau === 'erreur'; }).length,
    refus24h: recent.filter(function (j) { return j.action === 'Connexion refusée'; }).length,
    connexions24h: recent.filter(function (j) { return j.action === 'Connexion'; }).length,
    journal: journal,
    ms: Date.now() - t0
  };
}

/** Tests ciblés. what : classeur | agenda | gmail | invitation. */
function api_adminTest(token, what, address) {
  const me = adminGuard_(token);
  const t0 = Date.now();
  const owner = Session.getEffectiveUser().getEmail();
  let detail = '';
  try {
    if (what === 'classeur') {
      const n = readAll_(SHEET_PROFILS, PROFIL_FIELDS, 'photo').length;
      detail = n + ' fiche(s) lue(s)';
    } else if (what === 'agenda') {
      const cal = getCalendar_();
      const start = new Date(Date.now() + 3600000);
      const ev = cal.createEvent('Test CRM Guérim', start, new Date(start.getTime() + 900000));
      ev.deleteEvent();
      detail = 'Événement créé puis supprimé dans « ' + cal.getName() + ' »';
    } else if (what === 'gmail') {
      GmailApp.sendEmail(owner, 'Test CRM Guérim', 'Envoi de test depuis le CRM, le ' + new Date().toLocaleString('fr-FR') + '.', { name: getSettings_().expediteur });
      detail = 'Mail de test envoyé à ' + owner;
    } else if (what === 'invitation') {
      const to = validEmail_(address) ? address : owner;
      const tomorrow = Utilities.formatDate(new Date(Date.now() + 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      sendInvite_(to, { date: tomorrow, heure: '09:00', dur: 15, title: 'Test CRM Guérim', desc: 'Invitation de test, à supprimer.', lieu: '' },
        'test-' + Date.now() + '@guerim-crm', 1, 'REQUEST');
      detail = 'Invitation de test envoyée à ' + to + ' pour demain 9 h';
    } else {
      throw new Error('Test inconnu.');
    }
  } catch (err) {
    log_('erreur', me.name, 'Test ' + what, err.message);
    return { ok: false, ms: Date.now() - t0, detail: err.message };
  }
  log_('info', me.name, 'Test ' + what, detail);
  return { ok: true, ms: Date.now() - t0, detail: detail };
}

/** Écrit une ligne dans l'onglet Journal du classeur. Ne bloque jamais l'action en cours. */
function log_(niveau, utilisateur, action, detail) {
  try {
    const ss = getSpreadsheet_();
    let sh = ss.getSheetByName(SHEET_JOURNAL);
    if (!sh) {
      sh = ss.insertSheet(SHEET_JOURNAL);
      sh.getRange(1, 1, 1, JOURNAL_FIELDS.length).setValues([JOURNAL_FIELDS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date().toISOString(), niveau, escape_(utilisateur), escape_(action), escape_(String(detail || '').slice(0, 500))]);
    if (sh.getLastRow() > JOURNAL_MAX + 200) sh.deleteRows(2, 200);
  } catch (err) { /* le journal ne doit jamais faire échouer l'action */ }
}

function readJournal_(n) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_JOURNAL);
  if (!sh || sh.getLastRow() < 2) return [];
  const last = sh.getLastRow();
  const count = Math.min(n, last - 1);
  return sh.getRange(last - count + 1, 1, count, JOURNAL_FIELDS.length).getDisplayValues().reverse().map(function (r) {
    return { date: r[0], niveau: r[1], utilisateur: unescape_(r[2]), action: unescape_(r[3]), detail: unescape_(r[4]) };
  });
}

/* ---------- Comptes et connexion ---------- */

const LOGIN_RE = /^[a-z0-9._-]{3,30}$/;

/** Accessible sans session : indique si le premier compte reste à créer. */
function api_status() {
  return { needsSetup: !Object.keys(users_()).length, version: APP_VERSION };
}

/** Première ouverture : code d'installation (journal de la fonction installer) + premier compte. */
function api_setup(code, login, name, password) {
  return withLock_(function () {
    const users = users_();
    if (Object.keys(users).length) throw new Error('Le premier compte existe déjà.');
    checkLock_('setup');
    const expected = setupCode_(false);
    if (!expected || String(code || '').trim().toUpperCase() !== expected) {
      failed_('setup');
      throw new Error("Code d'installation incorrect. Lance la fonction installer dans l'éditeur pour l'afficher.");
    }
    login = normLogin_(login);
    users[login] = newUser_(name, password, '');
    users[login].lastLogin = new Date().toISOString();
    saveUsers_(users);
    PropertiesService.getScriptProperties().deleteProperty('SETUP_CODE');
    log_('info', users[login].name, 'Premier compte créé', login);
    return { token: newSession_(login) };
  });
}

function api_login(login, password, remember) {
  login = String(login || '').trim().toLowerCase();
  checkLock_(login);
  const users = users_();
  const u = users[login];
  if (!u || !u.active || hash_(u.salt, String(password || '')) !== u.hash) {
    failed_(login);
    log_('alerte', '', 'Connexion refusée', 'identifiant saisi : ' + login.slice(0, 30));
    throw new Error('Identifiant ou mot de passe incorrect.');
  }
  CacheService.getScriptCache().remove('auth_fails_' + login);
  withLock_(function () {
    const fresh = users_();
    if (fresh[login]) { fresh[login].lastLogin = new Date().toISOString(); saveUsers_(fresh); }
  });
  log_('info', u.name, remember ? 'Connexion (30 jours)' : 'Connexion', login);
  const token = newSession_(login);
  if (remember) saveLongSession_(token, login);
  return { token: token, mustChange: !!u.mustChange };
}

function api_logout(token) {
  if (token) {
    CacheService.getScriptCache().remove('sess_' + token);
    PropertiesService.getScriptProperties().deleteProperty('lsess_' + token);
  }
  return true;
}

function api_changePassword(token, current, next) {
  const me = guard_(token, true);
  return withLock_(function () {
    const users = users_();
    const u = users[me.login];
    if (hash_(u.salt, String(current || '')) !== u.hash) throw new Error('Le mot de passe actuel est incorrect.');
    if (String(next || '') === String(current || '')) throw new Error("Choisis un mot de passe différent de l'actuel.");
    const wasTemp = !!u.mustChange;
    Object.assign(u, cred_(next));
    delete u.mustChange;
    saveUsers_(users);
    dropLongSessions_(me.login, token); // les autres appareils devront se reconnecter
    log_('info', me.name, wasTemp ? 'Mot de passe personnel choisi' : 'Mot de passe changé', me.login);
    return true;
  });
}

function api_listUsers(token) {
  const me = adminGuard_(token);
  const users = users_();
  return Object.keys(users).map(function (login) {
    const u = users[login];
    return { login: login, name: u.name, email: u.email || '', admin: ADMIN_LOGINS.indexOf(login) > -1, active: u.active, lastLogin: u.lastLogin || '', mustChange: !!u.mustChange, createdAt: u.createdAt || '', createdBy: u.createdBy || '', isMe: login === me.login };
  }).sort(function (a, b) { return a.name.localeCompare(b.name, 'fr'); });
}

function api_addUser(token, login, name, password, email) {
  const me = adminGuard_(token);
  return withLock_(function () {
    const users = users_();
    login = normLogin_(login);
    if (users[login]) throw new Error('Cet identifiant est déjà pris.');
    email = String(email || '').trim();
    if (email && !validEmail_(email)) throw new Error("L'adresse e-mail de récupération n'est pas valide.");
    users[login] = newUser_(name, password, me.name);
    users[login].mustChange = true;
    if (email) users[login].email = email;
    saveUsers_(users);
    log_('info', me.name, 'Compte créé', login);
    return true;
  });
}

/** Les comptes administrateurs (ADMIN_LOGINS) ne se gèrent pas entre eux : chacun garde la main sur le sien. */
function protectAdmin_(me, login) {
  if (login !== me.login && ADMIN_LOGINS.indexOf(login) > -1) throw new Error("Ce compte est administrateur : seul son titulaire peut le modifier.");
}

function api_setUserActive(token, login, active) {
  const me = adminGuard_(token);
  protectAdmin_(me, login);
  return withLock_(function () {
    const users = users_();
    if (!users[login]) throw new Error('Compte introuvable.');
    if (login === me.login && !active) throw new Error('Tu ne peux pas désactiver ton propre compte.');
    users[login].active = !!active;
    saveUsers_(users);
    if (!active) dropLongSessions_(login, '');
    log_('info', me.name, active ? 'Compte réactivé' : 'Compte désactivé', login);
    return true;
  });
}

function api_resetUserPassword(token, login, password) {
  const me = adminGuard_(token);
  protectAdmin_(me, login);
  return withLock_(function () {
    const users = users_();
    if (!users[login]) throw new Error('Compte introuvable.');
    Object.assign(users[login], cred_(password));
    users[login].mustChange = login !== me.login;
    saveUsers_(users);
    dropLongSessions_(login, login === me.login ? token : '');
    log_('info', me.name, 'Mot de passe réinitialisé', login);
    return true;
  });
}

function api_deleteUser(token, login) {
  const me = adminGuard_(token);
  protectAdmin_(me, login);
  return withLock_(function () {
    const users = users_();
    if (login === me.login) throw new Error('Tu ne peux pas supprimer ton propre compte.');
    delete users[login];
    saveUsers_(users);
    dropLongSessions_(login, '');
    log_('info', me.name, 'Compte supprimé', login);
    return true;
  });
}

function api_setMyEmail(token, email) {
  const me = guard_(token);
  email = String(email || '').trim();
  if (email && !validEmail_(email)) throw new Error("L'adresse e-mail n'est pas valide.");
  return withLock_(function () {
    const users = users_();
    if (email) users[me.login].email = email; else delete users[me.login].email;
    saveUsers_(users);
    log_('info', me.name, 'E-mail de récupération modifié', me.login);
    return true;
  });
}

/**
 * Mot de passe oublié, étape 1 : envoie un code à 6 chiffres valable 15 minutes.
 * La réponse est identique que le compte existe ou non, pour ne rien révéler.
 */
function api_forgot(login) {
  login = String(login || '').trim().toLowerCase();
  const generic = { ok: true };
  if (!LOGIN_RE.test(login)) return generic;
  const cache = CacheService.getScriptCache();
  if (cache.get('reset_sent_' + login)) return generic; // un envoi par minute au plus
  const u = users_()[login];
  if (!u || !u.active || !validEmail_(u.email)) {
    log_('alerte', '', 'Mot de passe oublié sans e-mail', 'identifiant saisi : ' + login.slice(0, 30));
    return generic;
  }
  const code = String(Math.floor(100000 + Math.random() * 900000));
  const salt = Utilities.getUuid();
  cache.put('reset_' + login, JSON.stringify({ salt: salt, hash: hash_(salt, code) }), 900);
  cache.put('reset_sent_' + login, '1', 60);
  try {
    GmailApp.sendEmail(u.email, 'Guérim CRM : code de réinitialisation',
      'Bonjour ' + u.name + ',\n\nVoici ton code pour choisir un nouveau mot de passe : ' + code +
      "\n\nIl est valable 15 minutes. Si tu n'as rien demandé, ignore ce message : ton mot de passe actuel reste valable.",
      { name: getSettings_().expediteur });
    log_('info', u.name, 'Code de réinitialisation envoyé', login);
  } catch (err) {
    log_('erreur', u.name, 'Code de réinitialisation non envoyé', err.message);
  }
  return generic;
}

/** Mot de passe oublié, étape 2 : code reçu + nouveau mot de passe, puis connexion directe. */
function api_resetWithCode(login, code, password) {
  login = String(login || '').trim().toLowerCase();
  checkLock_('reset_' + login);
  const cache = CacheService.getScriptCache();
  const st = parseJson_(cache.get('reset_' + login));
  if (!st.salt || hash_(st.salt, String(code || '').trim()) !== st.hash) {
    failed_('reset_' + login);
    throw new Error('Code incorrect ou expiré.');
  }
  return withLock_(function () {
    const users = users_();
    const u = users[login];
    if (!u || !u.active) throw new Error('Code incorrect ou expiré.');
    Object.assign(u, cred_(password));
    delete u.mustChange;
    u.lastLogin = new Date().toISOString();
    saveUsers_(users);
    cache.remove('reset_' + login);
    dropLongSessions_(login, '');
    log_('info', u.name, 'Mot de passe réinitialisé par code', login);
    return { token: newSession_(login) };
  });
}

/**
 * Secours, à lancer depuis l'éditeur si plus personne ne peut se connecter :
 * efface tous les comptes et affiche un nouveau code d'installation pour recréer le premier.
 */
function reinitialiserAcces() {
  PropertiesService.getScriptProperties().deleteProperty('USERS');
  Logger.log("Comptes effacés. Nouveau code d'installation : " + setupCode_(true));
}

function guard_(token, allowTemp) {
  const cache = CacheService.getScriptCache();
  let login = token ? cache.get('sess_' + token) : null;
  if (!login && token) login = readLongSession_(token); // session de 30 jours, au-delà des 6 h du cache
  const u = login ? users_()[login] : null;
  if (!u || !u.active) {
    if (token) { cache.remove('sess_' + token); PropertiesService.getScriptProperties().deleteProperty('lsess_' + token); }
    throw new Error('SESSION_EXPIREE');
  }
  cache.put('sess_' + token, login, SESSION_SECONDS);
  if (u.mustChange && !allowTemp) throw new Error('MDP_A_CHANGER');
  return { login: login, name: u.name, email: u.email || '', admin: ADMIN_LOGINS.indexOf(login) > -1 };
}

/** Session longue : conservée dans les propriétés du script, valable 30 jours après la connexion. */
function saveLongSession_(token, login) {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  // Ménage : supprime les sessions longues expirées.
  const all = props.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('lsess_') === 0) { const v = parseJson_(all[k]); if (!v.exp || v.exp < now) props.deleteProperty(k); }
  });
  props.setProperty('lsess_' + token, JSON.stringify({ login: login, exp: now + LONG_SESSION_DAYS * 86400000 }));
}

function readLongSession_(token) {
  const props = PropertiesService.getScriptProperties();
  const v = parseJson_(props.getProperty('lsess_' + token));
  if (!v.login) return null;
  if (!v.exp || v.exp < Date.now()) { props.deleteProperty('lsess_' + token); return null; }
  return v.login;
}

/** Ferme toutes les sessions longues d'un compte (changement ou réinitialisation de mot de passe). */
function dropLongSessions_(login, keepToken) {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('lsess_') !== 0 || k === 'lsess_' + keepToken) return;
    if (parseJson_(all[k]).login === login) props.deleteProperty(k);
  });
}

function newSession_(login) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  CacheService.getScriptCache().put('sess_' + token, login, SESSION_SECONDS);
  return token;
}

function users_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('USERS');
  if (raw) return parseJson_(raw);
  // Reprise de la v3 : l'ancien mot de passe unique devient le compte « caroline ».
  const salt = props.getProperty('PWD_SALT'), hash = props.getProperty('PWD_HASH');
  if (salt && hash) {
    const users = { caroline: { name: 'Caroline', salt: salt, hash: hash, active: true, createdAt: new Date().toISOString(), createdBy: '' } };
    saveUsers_(users);
    props.deleteProperty('PWD_SALT');
    props.deleteProperty('PWD_HASH');
    return users;
  }
  return {};
}

function saveUsers_(users) {
  PropertiesService.getScriptProperties().setProperty('USERS', JSON.stringify(users));
}

function newUser_(name, password, createdBy) {
  name = String(name || '').trim();
  if (!name) throw new Error('Indique le nom de la personne.');
  return Object.assign({ name: name, active: true, createdAt: new Date().toISOString(), createdBy: createdBy || '' }, cred_(password));
}

function normLogin_(login) {
  login = String(login || '').trim().toLowerCase();
  if (!LOGIN_RE.test(login)) throw new Error("L'identifiant doit faire 3 à 30 caractères : lettres minuscules sans accent, chiffres, point, tiret ou tiret bas.");
  return login;
}

function cred_(pwd) {
  pwd = String(pwd || '');
  if (pwd.length < MIN_PASSWORD) throw new Error('Le mot de passe doit contenir au moins ' + MIN_PASSWORD + ' caractères.');
  const salt = Utilities.getUuid();
  return { salt: salt, hash: hash_(salt, pwd) };
}

function hash_(salt, pwd) {
  let h = salt + ':' + pwd;
  for (let i = 0; i < 300; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + h, Utilities.Charset.UTF_8)
      .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  }
  return h;
}

function setupCode_(create) {
  const props = PropertiesService.getScriptProperties();
  let code = props.getProperty('SETUP_CODE');
  if (!code && create) {
    code = Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
    props.setProperty('SETUP_CODE', code);
  }
  return code || '';
}

function checkLock_(key) {
  if (CacheService.getScriptCache().get('auth_lock_' + key)) throw new Error('Trop de tentatives. Réessaie dans 15 minutes.');
}

function failed_(key) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('auth_fails_' + key) || 0) + 1;
  if (n >= MAX_FAILS) { cache.put('auth_lock_' + key, '1', LOCK_SECONDS); cache.remove('auth_fails_' + key); }
  else cache.put('auth_fails_' + key, String(n), LOCK_SECONDS);
}

/* ---------- Événements d'agenda ---------- */

function profilEvents_(p) {
  if (!p || p.consentement === 'retire') return {};
  const who = shortName_(p);
  const ep = p.numeroEpisode ? 'épisode ' + p.numeroEpisode : '';
  const desc = [
    p.numeroEpisode ? 'Épisode ' + p.numeroEpisode + (p.titreEpisode ? ' : ' + p.titreEpisode : '') : '',
    'Fiche : ' + appUrl_()
  ].filter(String).join('\n');
  const out = {};
  if (validDate_(p.datePreEntretien)) {
    out.pre = { date: p.datePreEntretien, heure: validTime_(p.heurePreEntretien) ? p.heurePreEntretien : '', dur: 30,
      title: 'Guérim : pré-entretien, ' + who, desc: desc, lieu: '' };
  }
  if (validDate_(p.dateEnregistrement)) {
    out.enr = { date: p.dateEnregistrement, heure: validTime_(p.heureEnregistrement) ? p.heureEnregistrement : '', dur: 90,
      title: 'Guérim : enregistrement, ' + who, desc: desc, lieu: p.lieuEnregistrement || '' };
  }
  if (validDate_(p.dateDiffusionRadio)) {
    out.radio = { date: p.dateDiffusionRadio, heure: '', dur: 0,
      title: 'Guérim : diffusion radio, ' + (ep ? ep + ' (' + who + ')' : who), desc: desc, lieu: '' };
  }
  if (validDate_(p.dateMiseEnLigne)) {
    out.web = { date: p.dateMiseEnLigne, heure: '', dur: 0,
      title: 'Guérim : mise en ligne, ' + (ep ? ep + ' (' + who + ')' : who), desc: desc, lieu: '' };
  }
  return out;
}

const RESEAUX = { instagram: 'Instagram', facebook: 'Facebook', linkedin: 'LinkedIn', tiktok: 'TikTok', youtube: 'YouTube', x: 'X' };

/** Post réseaux : rappel dans les agendas tant qu'il n'est pas publié. Titre neutre, prénom et initiale seulement. */
function postEvents_(t, p) {
  const s = getSettings_();
  if (s.posts === 'non' || t.statut === 'publie' || !validDate_(t.date)) return {};
  const nets = String(t.reseaux || '').split(',').map(function (k) { return RESEAUX[k] || ''; }).filter(String).join(', ');
  const who = p && p.consentement !== 'retire' ? shortName_(p) : '';
  return {
    post: { date: t.date, heure: validTime_(t.heure) ? t.heure : '', dur: 15,
      title: 'Guérim : post ' + (nets || 'réseaux') + (t.sujet ? ', ' + t.sujet : '') + (who ? ' (' + who + ')' : ''),
      desc: (t.texte ? String(t.texte).slice(0, 1500) + '\n\n' : '') + 'CRM : ' + appUrl_(), lieu: '' }
  };
}

function api_savePost(token, t) {
  const me = guard_(token);
  return withLock_(function () {
    const clean = pick_(t, POST_FIELDS);
    if (!validDate_(clean.date)) throw new Error('Indique la date de publication.');
    if (!clean.sujet && !clean.texte) throw new Error('Indique au moins un sujet ou un texte.');
    if (!validTime_(clean.heure)) clean.heure = '';
    clean.reseaux = clean.reseaux.split(',').filter(function (k) { return RESEAUX[k]; }).join(',');
    clean.statut = clean.statut === 'publie' ? 'publie' : 'a_faire';
    const prev = clean.id ? findObj_(SHEET_POSTS, POST_FIELDS, clean.id) : null;
    const now = new Date().toISOString();
    if (!clean.id) { clean.id = Utilities.getUuid(); clean.createdAt = now; }
    clean.updatedAt = now;
    clean.auteur = prev ? prev.auteur : me.name;
    clean.syncIds = prev ? prev.syncIds : '';
    upsert_(SHEET_POSTS, POST_FIELDS, clean);
    return clean; // la synchronisation agenda suit via api_syncPost
  });
}

function api_syncPost(token, id) {
  const me = guard_(token);
  return withLock_(function () {
    const t = findObj_(SHEET_POSTS, POST_FIELDS, id);
    if (!t) return { syncIds: '', _warn: '' };
    const p = t.profilId ? findObj_(SHEET_PROFILS, PROFIL_FIELDS, t.profilId) : null;
    const res = syncItems_(t.syncIds, postEvents_(t, p), t.id, false);
    if (res.raw !== t.syncIds) { t.syncIds = res.raw; upsert_(SHEET_POSTS, POST_FIELDS, t); }
    if (res.errors.length) log_('erreur', me.name, 'Synchronisation post', (t.sujet || t.date) + ' / ' + res.errors.join(' '));
    return { syncIds: t.syncIds, _warn: res.errors.join(' ') };
  });
}

/** Retrait ou suppression d'un invité : ses posts restent, mais sans lien ni nom dans les agendas. */
function detachPostsOf_(profilId) {
  const errors = [];
  readAll_(SHEET_POSTS, POST_FIELDS).forEach(function (t) {
    if (t.profilId !== profilId) return;
    t.profilId = '';
    const res = syncItems_(t.syncIds, postEvents_(t, null), t.id, false);
    t.syncIds = res.raw;
    errors.push.apply(errors, res.errors);
    upsert_(SHEET_POSTS, POST_FIELDS, t);
  });
  return errors;
}

function api_deletePost(token, id) {
  guard_(token);
  return withLock_(function () {
    const t = findObj_(SHEET_POSTS, POST_FIELDS, id);
    const errors = t ? syncItems_(t.syncIds, {}, t.id, false).errors : [];
    deleteById_(SHEET_POSTS, POST_FIELDS, id);
    return { ok: true, _warn: errors.join(' ') };
  });
}

function echangeEvents_(e, p) {
  const s = getSettings_();
  if (s.relances !== 'oui' || !p || p.consentement === 'retire' || e.faite === 'oui' || !validDate_(e.dateRelance)) return {};
  return {
    rel: { date: e.dateRelance, heure: '', dur: 0,
      title: 'Guérim : relance, ' + shortName_(p) + (e.prochaineAction ? ' (' + e.prochaineAction + ')' : ''),
      desc: 'Fiche : ' + appUrl_(), lieu: '' }
  };
}

/**
 * Aligne Google Agenda et Outlook sur les événements voulus.
 * raw : état mémorisé (JSON), defs : événements voulus par clé, uidBase : identifiant stable.
 */
function syncItems_(raw, defs, uidBase, force) {
  const s = getSettings_();
  const ids = parseJson_(raw);
  const errors = [];
  const allTargets = inviteTargets_(s);
  unique_(Object.keys(ids).concat(Object.keys(defs))).forEach(function (k) {
    const d = defs[k] || null;
    const cur = ids[k] || {};
    const sig = d ? JSON.stringify(d) : '';
    const gWant = s.google === 'oui' && !!d;
    const targets = d ? allTargets : [];
    const prevT = Array.isArray(cur.o) ? cur.o : (cur.o ? [cur.o] : []);
    if (!force && sig === (cur.sig || '') && gWant === !!cur.g && sameList_(targets, prevT)) return;

    const uid = uidBase + '-' + k + '@guerim-crm';
    try {
      if (cur.g) { deleteGoogleEvent_(cur.g); delete cur.g; }
      if (gWant) cur.g = createGoogleEvent_(d);
    } catch (err) { errors.push('Agenda Guérim : ' + err.message); }

    let sent = prevT.slice();
    try {
      const removed = prevT.filter(function (t) { return targets.indexOf(t) < 0; });
      if (removed.length && cur.last) {
        cur.seq = (cur.seq || 0) + 1;
        removed.forEach(function (t) { sendInvite_(t, cur.last, uid, cur.seq, 'CANCEL'); sent = sent.filter(function (x) { return x !== t; }); });
      }
      if (targets.length) {
        cur.seq = (cur.seq || 0) + 1;
        targets.forEach(function (t) { sendInvite_(t, d, uid, cur.seq, 'REQUEST'); if (sent.indexOf(t) < 0) sent.push(t); });
        cur.last = d;
      }
      sent = targets.slice();
    } catch (err) { errors.push('Invitation agenda : ' + err.message); }
    cur.o = sent;
    if (!sent.length) delete cur.last;

    cur.sig = sig;
    if (!d && !cur.g && !cur.o.length) delete ids[k]; else ids[k] = cur;
  });
  const out = JSON.stringify(ids);
  return { raw: out === '{}' ? '' : out, errors: unique_(errors) };
}

function inviteTargets_(s) {
  const t = [];
  if (s.inviteOutlook === 'oui' && validEmail_(s.outlookEmail)) t.push(s.outlookEmail.toLowerCase());
  if (s.inviteGmail === 'oui' && validEmail_(s.gmailEmail)) t.push(s.gmailEmail.toLowerCase());
  return unique_(t);
}

function sameList_(a, b) {
  return a.length === b.length && a.every(function (x) { return b.indexOf(x) > -1; });
}

function getCalendar_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('CAL_ID');
  let cal = id ? CalendarApp.getCalendarById(id) : null;
  if (!cal) {
    cal = CalendarApp.createCalendar(getSettings_().calName || DEFAULT_SETTINGS.calName);
    props.setProperty('CAL_ID', cal.getId());
  }
  return cal;
}

function createGoogleEvent_(d) {
  const cal = getCalendar_();
  const tz = Session.getScriptTimeZone();
  const opts = { description: d.desc || '', location: d.lieu || '' };
  let ev;
  if (d.heure) {
    const start = Utilities.parseDate(d.date + ' ' + d.heure, tz, 'yyyy-MM-dd HH:mm');
    ev = cal.createEvent(d.title, start, new Date(start.getTime() + d.dur * 60000), opts);
  } else {
    ev = cal.createAllDayEvent(d.title, Utilities.parseDate(d.date + ' 12:00', tz, 'yyyy-MM-dd HH:mm'), opts);
  }
  return ev.getId();
}

function deleteGoogleEvent_(id) {
  const ev = getCalendar_().getEventById(id);
  if (ev) ev.deleteEvent();
}

/* ---------- Invitations iCalendar (Outlook et invités) ---------- */

function sendInvite_(to, d, uid, seq, method) {
  const ics = buildIcs_(d, uid, seq, method, to);
  const subject = (method === 'CANCEL' ? 'Annulé : ' : '') + d.title;
  const text = [d.title, humanDate_(d), d.lieu ? 'Lieu : ' + d.lieu : '', d.desc || ''].filter(String).join('\n');
  sendIcsMail_(to, subject, text, ics, method);
}

function buildIcs_(d, uid, seq, method, attendee) {
  const tz = Session.getScriptTimeZone();
  const s = getSettings_();
  const me = Session.getEffectiveUser().getEmail();
  const L = ['BEGIN:VCALENDAR', 'PRODID:-//Guerim CRM//FR', 'VERSION:2.0', 'CALSCALE:GREGORIAN', 'METHOD:' + method,
    'BEGIN:VEVENT', 'UID:' + uid, 'SEQUENCE:' + seq, 'DTSTAMP:' + utc_(new Date())];
  if (d.heure) {
    const st = Utilities.parseDate(d.date + ' ' + d.heure, tz, 'yyyy-MM-dd HH:mm');
    L.push('DTSTART:' + utc_(st), 'DTEND:' + utc_(new Date(st.getTime() + d.dur * 60000)));
  } else {
    const st = Utilities.parseDate(d.date + ' 12:00', tz, 'yyyy-MM-dd HH:mm');
    L.push('DTSTART;VALUE=DATE:' + Utilities.formatDate(st, tz, 'yyyyMMdd'),
      'DTEND;VALUE=DATE:' + Utilities.formatDate(new Date(st.getTime() + 86400000), tz, 'yyyyMMdd'));
  }
  L.push('SUMMARY:' + icsText_(d.title));
  if (d.desc) L.push('DESCRIPTION:' + icsText_(d.desc));
  if (d.lieu) L.push('LOCATION:' + icsText_(d.lieu));
  L.push('ORGANIZER;CN="' + String(s.expediteur).replace(/"/g, '') + '":mailto:' + me,
    'ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:' + attendee,
    'STATUS:' + (method === 'CANCEL' ? 'CANCELLED' : 'CONFIRMED'),
    'TRANSP:' + (d.heure ? 'OPAQUE' : 'TRANSPARENT'),
    'END:VEVENT', 'END:VCALENDAR');
  return L.map(fold_).join('\r\n') + '\r\n';
}

/**
 * Avec le service Gmail avancé : message multipart contenant une vraie partie text/calendar,
 * qu'Outlook et Gmail affichent comme une invitation (boutons Accepter / Refuser).
 * Sans lui : repli sur une pièce jointe .ics, à ouvrir pour l'ajouter à l'agenda.
 */
function sendIcsMail_(to, subject, text, ics, method) {
  const s = getSettings_();
  if (!gmailAdvanced_()) {
    const o = { name: s.expediteur, attachments: [Utilities.newBlob(ics, 'text/calendar', 'invitation.ics')] };
    if (validEmail_(s.replyTo)) o.replyTo = s.replyTo;
    GmailApp.sendEmail(to, subject, text, o);
    return;
  }
  const me = Session.getEffectiveUser().getEmail();
  const b = 'guerim' + Utilities.getUuid().replace(/-/g, '');
  const mime = [
    'From: ' + encHeader_(s.expediteur) + ' <' + me + '>',
    'To: ' + to,
    validEmail_(s.replyTo) ? 'Reply-To: ' + s.replyTo : null,
    'Subject: ' + encHeader_(subject),
    'MIME-Version: 1.0',
    'Content-Type: multipart/mixed; boundary="' + b + '"',
    '',
    '--' + b,
    'Content-Type: multipart/alternative; boundary="alt' + b + '"',
    '',
    '--alt' + b,
    'Content-Type: text/plain; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    b64_(text),
    '--alt' + b,
    'Content-Type: text/calendar; charset=UTF-8; method=' + method,
    'Content-Transfer-Encoding: base64',
    '',
    b64_(ics),
    '--alt' + b + '--',
    '',
    '--' + b,
    'Content-Type: application/ics; name="invitation.ics"',
    'Content-Disposition: attachment; filename="invitation.ics"',
    'Content-Transfer-Encoding: base64',
    '',
    b64_(ics),
    '--' + b + '--',
    ''
  ].filter(function (l) { return l !== null; }).join('\r\n');
  Gmail.Users.Messages.send({ raw: Utilities.base64EncodeWebSafe(Utilities.newBlob(mime).getBytes()) }, 'me');
}

function gmailAdvanced_() {
  return typeof Gmail !== 'undefined' && !!Gmail.Users && !!Gmail.Users.Messages;
}

/* ---------- Stockage ---------- */

let SS_CACHE_ = null;
const HEADERS_ = {};

function getSpreadsheet_() {
  if (SS_CACHE_) return SS_CACHE_;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SS_ID');
  if (id) {
    try { SS_CACHE_ = SpreadsheetApp.openById(id); return SS_CACHE_; } catch (err) { /* recréé plus bas */ }
  }
  const ss = SpreadsheetApp.create('Guérim CRM (données)');
  props.setProperty('SS_ID', ss.getId());
  const first = ss.getSheets()[0];
  first.setName(SHEET_PROFILS);
  SS_CACHE_ = ss;
  return ss;
}

function getSheet_(name, fields) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (!HEADERS_[name]) {
    const cache = CacheService.getScriptCache();
    const key = 'hdr_' + name + '_' + APP_VERSION;
    const cached = cache.get(key);
    if (cached) HEADERS_[name] = JSON.parse(cached);
    else { HEADERS_[name] = ensureHeader_(sh, fields); cache.put(key, JSON.stringify(HEADERS_[name]), 21600); }
  }
  return sh;
}

/** Ajoute les colonnes manquantes en fin de ligne d'en-tête : les classeurs de la v1 restent compatibles. */
function ensureHeader_(sh, fields) {
  let header = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0] : [];
  while (header.length && !header[header.length - 1]) header.pop();
  const missing = fields.filter(function (f) { return header.indexOf(f) < 0; });
  if (missing.length) {
    const col = header.length + 1;
    if (sh.getMaxColumns() < header.length + missing.length) sh.insertColumnsAfter(sh.getMaxColumns(), header.length + missing.length - sh.getMaxColumns());
    sh.getRange(1, col, sh.getMaxRows(), missing.length).setNumberFormat('@');
    sh.getRange(1, col, 1, missing.length).setValues([missing]).setFontWeight('bold');
    header = header.concat(missing);
  }
  if (sh.getFrozenRows() !== 1) sh.setFrozenRows(1);
  return header;
}

/** skip : colonne lourde à ne pas lire (renvoyée vide). */
function readAll_(name, fields, skip) {
  const sh = getSheet_(name, fields);
  const header = HEADERS_[name];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const idIdx = header.indexOf('id');
  const n = last - 1;
  const sk = skip ? header.indexOf(skip) : -1;
  let rows;
  if (sk < 0) {
    rows = sh.getRange(2, 1, n, header.length).getDisplayValues();
  } else {
    const left = sk > 0 ? sh.getRange(2, 1, n, sk).getDisplayValues() : null;
    const right = sk < header.length - 1 ? sh.getRange(2, sk + 2, n, header.length - sk - 1).getDisplayValues() : null;
    rows = [];
    for (let i = 0; i < n; i++) rows.push((left ? left[i] : []).concat([''], right ? right[i] : []));
  }
  return rows
    .filter(function (r) { return r[idIdx]; })
    .map(function (r) {
      const o = {};
      fields.forEach(function (f) { const i = header.indexOf(f); o[f] = i < 0 ? '' : unescape_(r[i]); });
      return o;
    });
}

function findObj_(name, fields, id) {
  if (!id) return null;
  return readAll_(name, fields).filter(function (o) { return o.id === id; })[0] || null;
}

function upsert_(name, fields, obj) {
  const sh = getSheet_(name, fields);
  const header = HEADERS_[name];
  const found = findRow_(sh, header, obj.id);
  const row = found || sh.getLastRow() + 1;
  const existing = found ? sh.getRange(row, 1, 1, header.length).getValues()[0] : null;
  const values = header.map(function (h, i) {
    if (fields.indexOf(h) > -1) return escape_(obj[h]);
    return existing ? existing[i] : ''; // colonnes ajoutées à la main : préservées
  });
  sh.getRange(row, 1, 1, header.length).setNumberFormat('@').setValues([values]);
}

function deleteById_(name, fields, id) {
  const sh = getSheet_(name, fields);
  const row = findRow_(sh, HEADERS_[name], id);
  if (row) sh.deleteRow(row);
}

/** Annule les relances d'agenda d'un profil puis supprime ses échanges. Renvoie les erreurs de synchronisation. */
function clearEchangesOf_(profilId) {
  const errors = [];
  readAll_(SHEET_ECHANGES, ECHANGE_FIELDS).forEach(function (e) {
    if (e.profilId === profilId && e.syncIds) errors.push.apply(errors, syncItems_(e.syncIds, {}, e.id, false).errors);
  });
  const sh = getSheet_(SHEET_ECHANGES, ECHANGE_FIELDS);
  const header = HEADERS_[SHEET_ECHANGES];
  const col = header.indexOf('profilId') + 1;
  const last = sh.getLastRow();
  if (last < 2) return errors;
  const vals = sh.getRange(2, col, last - 1, 1).getDisplayValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] === profilId) sh.deleteRow(i + 2);
  }
  return errors;
}

function findRow_(sh, header, id) {
  if (!id) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, header.indexOf('id') + 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

/* ---------- Réglages ---------- */

function getSettings_() {
  const raw = PropertiesService.getScriptProperties().getProperty('SETTINGS');
  const saved = parseJson_(raw);
  if (saved.outlook && !saved.inviteOutlook) saved.inviteOutlook = saved.outlook; // réglages v2
  delete saved.outlook;
  return Object.assign({}, DEFAULT_SETTINGS, saved);
}

/* ---------- Utilitaires ---------- */

function pick_(src, fields) {
  const o = {};
  fields.forEach(function (f) { o[f] = src && src[f] != null ? String(src[f]).trim() : ''; });
  return o;
}

// Empêche Sheets d'interpréter une saisie commençant par "=" comme une formule.
function escape_(v) {
  v = v == null ? '' : String(v);
  return v.charAt(0) === '=' ? "'" + v : v;
}

function unescape_(v) {
  v = v == null ? '' : String(v);
  return v.charAt(0) === "'" && v.charAt(1) === '=' ? v.slice(1) : v;
}

function parseJson_(raw) {
  if (!raw) return {};
  try { const o = JSON.parse(raw); return o && typeof o === 'object' ? o : {}; } catch (err) { return {}; }
}

function unique_(arr) {
  return arr.filter(function (v, i) { return v && arr.indexOf(v) === i; });
}

function validDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }
function validTime_(s) { return /^([01]\d|2[0-3]):[0-5]\d$/.test(String(s || '')); }
function validEmail_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '')); }

function shortName_(p) {
  return [p.prenom, p.nom ? p.nom.charAt(0) + '.' : ''].filter(String).join(' ') || 'invité';
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function humanDate_(d) {
  const tz = Session.getScriptTimeZone();
  const dt = Utilities.parseDate(d.date + ' ' + (d.heure || '12:00'), tz, 'yyyy-MM-dd HH:mm');
  return Utilities.formatDate(dt, tz, 'dd/MM/yyyy') + (d.heure ? ' à ' + d.heure.replace(':', ' h ') : '');
}

function appUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (err) { return ''; }
}

function utc_(d) { return Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'"); }

function icsText_(s) {
  return String(s || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
}

// Plie les lignes iCalendar à 75 octets (RFC 5545), sans couper un caractère accentué.
function fold_(line) {
  let out = '', cur = '', n = 0;
  for (const ch of line) {
    const c = ch.codePointAt(0);
    const bytes = c < 0x80 ? 1 : c < 0x800 ? 2 : c < 0x10000 ? 3 : 4;
    if (n + bytes > 74) { out += cur + '\r\n '; cur = ''; n = 1; }
    cur += ch; n += bytes;
  }
  return out + cur;
}

function b64_(s) {
  return Utilities.base64Encode(Utilities.newBlob(s).getBytes()).replace(/(.{76})/g, '$1\r\n');
}

function encHeader_(s) {
  return /^[\x20-\x7e]*$/.test(s) ? s : '=?UTF-8?B?' + Utilities.base64Encode(Utilities.newBlob(s).getBytes()) + '?=';
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}
