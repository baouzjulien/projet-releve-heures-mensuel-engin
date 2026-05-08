const SHEETS = {
  conducteurs: 'CONDUCTEURS',
  releves: 'RELEVES_MENSUELS',
  audit: 'AUDIT'
};

function doGet(e) {
  const action = (e.parameter.action || '').trim();
  try {
    if (action === 'login') return json(login(e.parameter.pin, e));
    if (action === 'getRelevesMensuels') return json(getRelevesMensuels(e));
    return json({ success: false, error: 'Action inconnue' });
  } catch (err) {
    audit('ERREUR_GET', null, String(err), e);
    return json({ success: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (payload.action === 'saveReleveMensuel') return json(saveReleveMensuel(payload, e));
    return json({ success: false, error: 'Action inconnue' });
  } catch (err) {
    audit('ERREUR_POST', null, String(err), e);
    return json({ success: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Onglet manquant: ' + name);
  return sh;
}

function rows(name) {
  const sh = sheet(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => Object.fromEntries(headers.map((h, i) => [h, r[i]])));
}

function login(pin, event) {
  const conducteur = rows(SHEETS.conducteurs)
    .find(c => String(c.PIN) === String(pin) && String(c.ACTIF || 'OUI').toUpperCase() !== 'NON');

  if (!conducteur) {
    audit('LOGIN_REFUSE', null, 'PIN refuse: ' + String(pin || ''), event);
    return { success: false };
  }

  const user = {
    id: String(conducteur.ID),
    nom: String(conducteur.NOM || ''),
    prenom: String(conducteur.PRENOM || ''),
    role: String(conducteur.ROLE || 'employe').toLowerCase(),
    email: String(conducteur.EMAIL || '')
  };

  audit('LOGIN_OK', user, 'Connexion reussie', event);
  return { success: true, user };
}

function getRelevesMensuels(event) {
  const user = userFromParams(event.parameter);
  if (!user || user.role !== 'responsable') {
    audit('CONSULTATION_REFUSEE', user, 'Acces responsable refuse', event);
    return { success: false, error: 'Acces reserve au responsable' };
  }

  const mois = String(event.parameter.mois || '');
  const all = rows(SHEETS.releves);
  const releves = all
    .filter(r => !mois || monthToKey(r.MOIS) === mois)
    .map(r => ({
      id: String(r.ID_RELEVE || ''),
      timestamp: dateToText(r.TIMESTAMP),
      mois: monthToKey(r.MOIS),
      idConducteur: String(r.ID_CONDUCTEUR || ''),
      nom: String(r.NOM || ''),
      prenom: String(r.PRENOM || ''),
      role: String(r.ROLE || ''),
      site: String(r.SITE || ''),
      dateReleve: dateToIso(r.DATE_RELEVE),
      immatriculation: String(r.IMMATRICULATION || ''),
      heurePorteur: r.HEURE_PORTEUR || '',
      heureAuxiliaires: r.HEURE_AUXILIAIRES || '',
      kilometrage: r.KILOMETRAGE || '',
      commentaire: String(r.COMMENTAIRE || '')
    }))
    .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));

  audit('CONSULTATION_RESPONSABLE', user, 'Mois: ' + (mois || 'tous') + ', resultats: ' + releves.length, event);
  return { success: true, releves };
}

function saveReleveMensuel(payload, event) {
  const conducteur = payload.conducteur || {};
  const releve = payload.releve || {};
  const required = ['mois', 'site', 'dateReleve', 'immatriculation'];

  required.forEach(key => {
    if (!releve[key]) throw new Error('Champ obligatoire manquant: ' + key);
  });

  const headers = [
    'ID_RELEVE',
    'TIMESTAMP',
    'MOIS',
    'ID_CONDUCTEUR',
    'NOM',
    'PRENOM',
    'ROLE',
    'SITE',
    'DATE_RELEVE',
    'IMMATRICULATION',
    'HEURE_PORTEUR',
    'HEURE_AUXILIAIRES',
    'KILOMETRAGE',
    'COMMENTAIRE'
  ];

  const row = {
    ID_RELEVE: Utilities.getUuid(),
    TIMESTAMP: new Date(),
    MOIS: releve.mois || '',
    ID_CONDUCTEUR: conducteur.id || '',
    NOM: conducteur.nom || '',
    PRENOM: conducteur.prenom || '',
    ROLE: conducteur.role || '',
    SITE: releve.site || '',
    DATE_RELEVE: releve.dateReleve || '',
    IMMATRICULATION: releve.immatriculation || '',
    HEURE_PORTEUR: numberOrBlank(releve.heurePorteur),
    HEURE_AUXILIAIRES: numberOrBlank(releve.heureAuxiliaires),
    KILOMETRAGE: numberOrBlank(releve.kilometrage),
    COMMENTAIRE: releve.commentaire || ''
  };

  const sh = sheet(SHEETS.releves);
  const nextRow = sh.getLastRow() + 1;
  sh.getRange(nextRow, 1, 1, headers.length).setValues([headers.map(h => row[h] ?? '')]);
  sh.getRange(nextRow, 3).setNumberFormat('@').setValue(String(row.MOIS || ''));
  sh.getRange(nextRow, 9).setNumberFormat('@').setValue(String(row.DATE_RELEVE || ''));
  audit('ENVOI_RELEVE', conducteur, `Mois ${row.MOIS}, engin ${row.IMMATRICULATION}, site ${row.SITE}`, event);
  envoyerMailResponsable(row);

  return { success: true, id: row.ID_RELEVE };
}

function envoyerMailResponsable(row) {
  const responsable = rows(SHEETS.conducteurs)
    .find(c => String(c.ROLE || '').toLowerCase() === 'responsable' && String(c.ACTIF || 'OUI').toUpperCase() !== 'NON');

  const email = responsable && String(responsable.EMAIL || '').trim();
  if (!email) {
    audit('EMAIL_NON_ENVOYE', responsable || null, 'Adresse e-mail responsable absente', null);
    return;
  }

  const sujet = `Nouveau releve mensuel engin - ${row.IMMATRICULATION} - ${row.MOIS}`;
  const corps = [
    'Un nouveau releve mensuel engin a ete envoye.',
    '',
    `Chauffeur : ${row.PRENOM} ${row.NOM}`,
    `Mois : ${row.MOIS}`,
    `Site : ${row.SITE}`,
    `Date : ${row.DATE_RELEVE}`,
    `Immatriculation : ${row.IMMATRICULATION}`,
    `Heure porteur : ${row.HEURE_PORTEUR}`,
    `Heure auxiliaires : ${row.HEURE_AUXILIAIRES}`,
    `Kilometrage : ${row.KILOMETRAGE}`,
    `Commentaire : ${row.COMMENTAIRE || '-'}`,
    '',
    `ID releve : ${row.ID_RELEVE}`
  ].join('\n');

  MailApp.sendEmail(email, sujet, corps);
  audit('EMAIL_RESPONSABLE', responsable, 'Email envoye a ' + email + ' pour releve ' + row.ID_RELEVE, null);
}

function audit(action, user, details, event) {
  try {
    const headers = [
      'ID_AUDIT',
      'TIMESTAMP',
      'ACTION',
      'ID_UTILISATEUR',
      'NOM',
      'PRENOM',
      'ROLE',
      'DETAILS',
      'USER_AGENT'
    ];
    const ua = event && event.parameter ? String(event.parameter.userAgent || '') : '';
    const row = {
      ID_AUDIT: Utilities.getUuid(),
      TIMESTAMP: new Date(),
      ACTION: action || '',
      ID_UTILISATEUR: user && user.id ? user.id : '',
      NOM: user && user.nom ? user.nom : '',
      PRENOM: user && user.prenom ? user.prenom : '',
      ROLE: user && user.role ? user.role : '',
      DETAILS: details || '',
      USER_AGENT: ua
    };
    sheet(SHEETS.audit).appendRow(headers.map(h => row[h] ?? ''));
  } catch (err) {
    // L'audit ne doit jamais bloquer l'application.
  }
}

function userFromParams(params) {
  if (!params || !params.idUtilisateur) return null;
  return {
    id: String(params.idUtilisateur || ''),
    nom: String(params.nom || ''),
    prenom: String(params.prenom || ''),
    role: String(params.role || '').toLowerCase()
  };
}

function numberOrBlank(value) {
  if (value === '' || value === null || value === undefined) return '';
  const n = Number(String(value).replace(',', '.'));
  return isNaN(n) ? value : n;
}

function dateToIso(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value).slice(0, 10);
}

function monthToKey(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  const text = String(value);
  const match = text.match(/^(\d{4})-(\d{2})/);
  if (match) return match[1] + '-' + match[2];
  return text.slice(0, 7);
}

function dateToText(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}
