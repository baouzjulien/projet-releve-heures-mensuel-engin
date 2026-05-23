const SHEETS = {
  conducteurs: 'CONDUCTEURS',
  releves: 'RELEVES_MENSUELS',
  audit: 'AUDIT'
};

const RELEVE_HEADERS = [
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
  'COMMENTAIRE',
  'ID_LIGNE',
  'NUM_LIGNE',
  'STATUT',
  'DATE_VALIDATION',
  'ID_RESPONSABLE',
  'NOM_RESPONSABLE',
  'PRENOM_RESPONSABLE'
];

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
    if (payload.action === 'validerReleveMensuel') return json(validerReleveMensuel(payload, e));
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

function ensureReleveHeaders() {
  const sh = sheet(SHEETS.releves);
  const current = sh.getLastRow() ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String) : [];
  const missing = RELEVE_HEADERS.filter(h => !current.includes(h));
  if (!current.length || missing.length) {
    sh.getRange(1, 1, 1, RELEVE_HEADERS.length).setValues([RELEVE_HEADERS]);
  }
  return sh;
}

function findRowsByReleveId(sh, idReleve) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1)
    .map((r, i) => ({
      rowIndex: i + 2,
      row: Object.fromEntries(headers.map((h, col) => [h, r[col]]))
    }))
    .filter(item => String(item.row.ID_RELEVE || '') === String(idReleve));
}

function normaliserLignesReleve(releve) {
  const lignes = Array.isArray(releve.lignes) && releve.lignes.length ? releve.lignes : [releve];
  return lignes.map((ligne, index) => {
    ['site', 'dateReleve', 'immatriculation'].forEach(key => {
      if (!ligne[key]) throw new Error('Champ obligatoire manquant ligne ' + (index + 1) + ': ' + key);
    });
    if (ligne.heurePorteur === '' && ligne.heureAuxiliaires === '' && ligne.kilometrage === '') {
      throw new Error('Heure ou kilometrage obligatoire ligne ' + (index + 1));
    }
    return {
      idLigne: String(ligne.idLigne || Utilities.getUuid()),
      site: String(ligne.site || '').trim(),
      dateReleve: String(ligne.dateReleve || ''),
      immatriculation: String(ligne.immatriculation || '').trim(),
      heurePorteur: ligne.heurePorteur,
      heureAuxiliaires: ligne.heureAuxiliaires,
      kilometrage: ligne.kilometrage,
      commentaire: String(ligne.commentaire || '').trim()
    };
  });
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
      idLigne: String(r.ID_LIGNE || ''),
      numLigne: Number(r.NUM_LIGNE || 1),
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
      commentaire: String(r.COMMENTAIRE || ''),
      statut: String(r.STATUT || 'ENVOYE'),
      dateValidation: dateToText(r.DATE_VALIDATION),
      responsableValidation: [r.PRENOM_RESPONSABLE, r.NOM_RESPONSABLE].filter(Boolean).join(' ')
    }))
    .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));

  audit('CONSULTATION_RESPONSABLE', user, 'Mois: ' + (mois || 'tous') + ', resultats: ' + releves.length, event);
  return { success: true, releves };
}

function saveReleveMensuel(payload, event) {
  const conducteur = payload.conducteur || {};
  const releve = payload.releve || {};
  if (!releve.mois) throw new Error('Champ obligatoire manquant: mois');
  const lignes = normaliserLignesReleve(releve);
  const idReleve = String(releve.idReleve || Utilities.getUuid());
  const sh = ensureReleveHeaders();
  const existing = findRowsByReleveId(sh, idReleve);
  const miseAJour = existing.length > 0;

  if (existing.some(item => String(item.row.STATUT || '').toUpperCase() === 'VALIDE')) {
    audit('MODIFICATION_REFUSEE_RELEVE_VALIDE', conducteur, 'Releve deja valide: ' + idReleve, event);
    throw new Error('Ce releve a deja ete valide par le responsable.');
  }

  const anciensIds = existing.map(item => String(item.row.ID_LIGNE || '')).filter(Boolean);
  const nouveauxIds = lignes.map(ligne => String(ligne.idLigne || '')).filter(Boolean);
  const idsSupprimes = anciensIds.filter(id => !nouveauxIds.includes(id));

  existing.sort((a, b) => b.rowIndex - a.rowIndex).forEach(item => sh.deleteRow(item.rowIndex));

  const timestamp = new Date();
  const rowsToWrite = lignes.map((ligne, index) => ({
    ID_RELEVE: idReleve,
    ID_LIGNE: ligne.idLigne || Utilities.getUuid(),
    NUM_LIGNE: index + 1,
    TIMESTAMP: timestamp,
    MOIS: releve.mois || '',
    ID_CONDUCTEUR: conducteur.id || '',
    NOM: conducteur.nom || '',
    PRENOM: conducteur.prenom || '',
    ROLE: conducteur.role || '',
    SITE: ligne.site || '',
    DATE_RELEVE: ligne.dateReleve || '',
    IMMATRICULATION: ligne.immatriculation || '',
    HEURE_PORTEUR: numberOrBlank(ligne.heurePorteur),
    HEURE_AUXILIAIRES: numberOrBlank(ligne.heureAuxiliaires),
    KILOMETRAGE: numberOrBlank(ligne.kilometrage),
    COMMENTAIRE: ligne.commentaire || '',
    STATUT: 'ENVOYE',
    DATE_VALIDATION: '',
    ID_RESPONSABLE: '',
    NOM_RESPONSABLE: '',
    PRENOM_RESPONSABLE: ''
  }));

  const nextRow = sh.getLastRow() + 1;
  sh.getRange(nextRow, 1, rowsToWrite.length, RELEVE_HEADERS.length)
    .setValues(rowsToWrite.map(row => RELEVE_HEADERS.map(h => row[h] ?? '')));
  sh.getRange(nextRow, 3, rowsToWrite.length, 1).setNumberFormat('@').setValues(rowsToWrite.map(row => [String(row.MOIS || '')]));
  sh.getRange(nextRow, 9, rowsToWrite.length, 1).setNumberFormat('@').setValues(rowsToWrite.map(row => [String(row.DATE_RELEVE || '')]));

  if (idsSupprimes.length) {
    audit('SUPPRESSION_LIGNE', conducteur, 'Releve ' + idReleve + ', lignes supprimees: ' + idsSupprimes.join(', '), event);
  }
  audit(miseAJour ? 'MISE_A_JOUR_RELEVE' : 'ENVOI_RELEVE', conducteur, `Mois ${releve.mois}, lignes ${rowsToWrite.length}`, event);
  envoyerMailResponsable(rowsToWrite[0], rowsToWrite.length, miseAJour);

  return { success: true, id: idReleve, statut: 'ENVOYE', miseAJour: miseAJour };
}

function validerReleveMensuel(payload, event) {
  const responsable = payload.responsable || {};
  if (String(responsable.role || '').toLowerCase() !== 'responsable') {
    audit('VALIDATION_REFUSEE', responsable, 'Acces validation refuse', event);
    throw new Error('Acces reserve au responsable');
  }
  const idReleve = String(payload.idReleve || '');
  if (!idReleve) throw new Error('ID releve manquant');
  const sh = ensureReleveHeaders();
  const existing = findRowsByReleveId(sh, idReleve);
  if (!existing.length) throw new Error('Releve introuvable');
  const dateValidation = new Date();
  existing.forEach(item => {
    sh.getRange(item.rowIndex, RELEVE_HEADERS.indexOf('STATUT') + 1).setValue('VALIDE');
    sh.getRange(item.rowIndex, RELEVE_HEADERS.indexOf('DATE_VALIDATION') + 1).setValue(dateValidation);
    sh.getRange(item.rowIndex, RELEVE_HEADERS.indexOf('ID_RESPONSABLE') + 1).setValue(responsable.id || '');
    sh.getRange(item.rowIndex, RELEVE_HEADERS.indexOf('NOM_RESPONSABLE') + 1).setValue(responsable.nom || '');
    sh.getRange(item.rowIndex, RELEVE_HEADERS.indexOf('PRENOM_RESPONSABLE') + 1).setValue(responsable.prenom || '');
  });
  audit('VALIDATION_RELEVE', responsable, 'Releve valide: ' + idReleve, event);
  return { success: true, id: idReleve, statut: 'VALIDE' };
}

function envoyerMailResponsable(row, nombreLignes, miseAJour) {
  const responsables = rows(SHEETS.conducteurs)
    .filter(c => String(c.ROLE || '').toLowerCase() === 'responsable' && String(c.ACTIF || 'OUI').toUpperCase() !== 'NON');

  const destinataires = responsables
    .map(r => ({ responsable: r, email: String(r.EMAIL || '').trim() }))
    .filter(r => r.email);

  if (!destinataires.length) {
    audit('EMAIL_NON_ENVOYE', null, 'Aucune adresse e-mail responsable renseignee', null);
    return;
  }

  const urlApplication = 'https://baouzjulien.github.io/projet-releve-heures-mensuel-engin/';
  const sujet = `${miseAJour ? 'Mise a jour' : 'Nouveau'} relevé mensuel engin - ${row.IMMATRICULATION} - ${monthToFrenchText(row.MOIS)}`;
  const corps = [
    miseAJour ? 'Un relevé mensuel engin a été mis a jour.' : 'Un nouveau relevé mensuel engin a été envoyé.',
    '',
    `Envoyé le : ${dateTimeToFrenchText(row.TIMESTAMP)}`,
    `Chauffeur : ${row.PRENOM} ${row.NOM}`,
    `Mois : ${monthToFrenchText(row.MOIS)}`,
    `Nombre de lignes : ${nombreLignes || 1}`,
    `Site : ${row.SITE}`,
    `Date : ${dateToFrenchText(row.DATE_RELEVE)}`,
    `Immatriculation : ${row.IMMATRICULATION}`,
    `Heure porteur : ${row.HEURE_PORTEUR}`,
    `Heure auxiliaires : ${row.HEURE_AUXILIAIRES}`,
    `Kilometrage : ${row.KILOMETRAGE}`,
    `Commentaire : ${row.COMMENTAIRE || '-'}`,
    '',
    'Lien vers l application :',
    urlApplication,
    '',
    `ID releve : ${row.ID_RELEVE}`
  ].join('\n');

  destinataires.forEach(({ responsable, email }) => {
    MailApp.sendEmail(email, sujet, corps);
    audit('EMAIL_RESPONSABLE', {
      id: String(responsable.ID || ''),
      nom: String(responsable.NOM || ''),
      prenom: String(responsable.PRENOM || ''),
      role: String(responsable.ROLE || 'responsable').toLowerCase()
    }, 'Email envoye a ' + email + ' pour releve ' + row.ID_RELEVE, null);
  });
}

function envoyerEmailBienvenueResponsable() {
  const responsable = rows(SHEETS.conducteurs)
    .find(c => String(c.ROLE || '').toLowerCase() === 'responsable' && String(c.ACTIF || 'OUI').toUpperCase() !== 'NON');

  if (!responsable) throw new Error('Aucun responsable actif trouve dans CONDUCTEURS');

  const email = String(responsable.EMAIL || '').trim();
  if (!email) throw new Error('Adresse e-mail responsable absente dans la colonne EMAIL');

  const releves = rows(SHEETS.releves).slice(-10).reverse();
  const lignes = releves.length
    ? releves.map(r => [
        '- ',
        monthToKey(r.MOIS),
        ' | ',
        r.IMMATRICULATION || 'Sans immatriculation',
        ' | ',
        r.SITE || 'Sans site',
        ' | ',
        [r.PRENOM, r.NOM].filter(Boolean).join(' ') || 'Sans chauffeur'
      ].join('')).join('\n')
    : 'Aucun releve enregistre pour le moment.';

  const urlApplication = 'https://baouzjulien.github.io/projet-releve-heures-mensuel-engin/';
  const sujet = 'Acces au suivi mensuel des engins HR Occitanie';
  const corps = [
    'Bonjour David,',
    '',
    'L application de suivi mensuel des engins HR Occitanie est disponible.',
    '',
    'Lien d acces :',
    urlApplication,
    '',
    'Tu peux te connecter avec ton code responsable pour consulter les releves par vehicule ou par chauffeur, puis exporter les donnees du mois en CSV ou Excel.',
    '',
    'Derniers releves enregistres :',
    lignes,
    '',
    'Ce message est un premier envoi de presentation. Les prochains releves envoyes par les chauffeurs declencheront automatiquement une notification.'
  ].join('\n');

  MailApp.sendEmail(email, sujet, corps);
  audit('EMAIL_BIENVENUE_RESPONSABLE', {
    id: String(responsable.ID || ''),
    nom: String(responsable.NOM || ''),
    prenom: String(responsable.PRENOM || ''),
    role: String(responsable.ROLE || 'responsable').toLowerCase()
  }, 'Email de bienvenue envoye a ' + email, null);
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

function dateToFrenchText(value) {
  if (!value) return '';
  const date = dateFromIsoText(value);
  if (!date) return String(value);
  return date.getDate() + ' ' + frenchMonthName(date.getMonth()) + ' ' + date.getFullYear();
}

function monthToFrenchText(value) {
  if (!value) return '';
  const text = String(value);
  const match = text.match(/^(\d{4})-(\d{2})/);
  if (!match) return text;
  const date = new Date(Number(match[1]), Number(match[2]) - 1, 1);
  return frenchMonthName(date.getMonth()) + ' ' + date.getFullYear();
}

function dateTimeToFrenchText(value) {
  if (!value) return '';
  const date = Object.prototype.toString.call(value) === '[object Date]' ? value : new Date(String(value).replace(' ', 'T'));
  if (isNaN(date.getTime())) return String(value);
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return dateToFrenchText(date) + ' à ' + date.getHours() + ' h ' + minutes;
}

function frenchMonthName(monthIndex) {
  return [
    'janvier',
    'février',
    'mars',
    'avril',
    'mai',
    'juin',
    'juillet',
    'août',
    'septembre',
    'octobre',
    'novembre',
    'décembre'
  ][monthIndex] || '';
}

function dateFromIsoText(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') return value;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}
