const SHEETS = {
  users: 'UTILISATEURS',
  engines: 'ENGINS',
  records: 'RELEVES',
  validations: 'VALIDATIONS'
};

function doGet(e) {
  const action = (e.parameter.action || '').trim();
  try {
    if (action === 'login') return json(login(e.parameter.pin));
    if (action === 'getConducteurs') return json(getConducteurs());
    if (action === 'getSemaine') return json(getSemaine(e.parameter.idConducteur, e.parameter.semaine));
    return json({ success: false, error: 'Action inconnue' });
  } catch (err) {
    return json({ success: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (payload.action === 'saveReleve') return json(saveReleve(payload));
    if (payload.action === 'signConducteur') return json(signConducteur(payload));
    if (payload.action === 'validateResponsable') return json(validateResponsable(payload));
    return json({ success: false, error: 'Action inconnue' });
  } catch (err) {
    return json({ success: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet(name) {
  const sh = ss().getSheetByName(name);
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

function appendObjects(name, objects, headers) {
  if (!objects.length) return;
  const sh = sheet(name);
  sh.getRange(sh.getLastRow() + 1, 1, objects.length, headers.length)
    .setValues(objects.map(o => headers.map(h => o[h] ?? '')));
}

function login(pin) {
  const user = rows(SHEETS.users).find(u => String(u.PIN) === String(pin) && String(u.ACTIF).toUpperCase() !== 'NON');
  if (!user) return { success: false };
  return {
    success: true,
    user: {
      id: String(user.ID),
      nom: String(user.NOM || ''),
      prenom: String(user.PRENOM || ''),
      role: String(user.ROLE || 'conducteur').toLowerCase()
    }
  };
}

function getConducteurs() {
  const users = rows(SHEETS.users)
    .filter(u => String(u.ROLE || '').toLowerCase() !== 'responsable' && String(u.ACTIF).toUpperCase() !== 'NON')
    .map(u => ({ id: String(u.ID), nom: String(u.NOM || ''), prenom: String(u.PRENOM || ''), role: String(u.ROLE || 'conducteur') }));
  return { success: true, users };
}

function getSemaine(idConducteur, semaine) {
  const lignes = rows(SHEETS.records)
    .filter(r => String(r.ID_CONDUCTEUR) === String(idConducteur) && toIso(r.SEMAINE_DU) === String(semaine))
    .sort((a, b) => Number(a.ORDRE || 0) - Number(b.ORDRE || 0));

  const validation = rows(SHEETS.validations)
    .find(v => String(v.ID_CONDUCTEUR) === String(idConducteur) && toIso(v.SEMAINE_DU) === String(semaine)) || {};

  const first = lignes[0] || {};
  return {
    success: true,
    lignes,
    meta: {
      engin: first.ENGIN || '',
      chantier: first.CHANTIER || ''
    },
    validation: {
      visaConducteur: validation.VISA_CONDUCTEUR || '',
      dateVisaConducteur: validation.DATE_VISA_CONDUCTEUR || '',
      visaResponsable: validation.VISA_RESPONSABLE || '',
      dateVisaResponsable: validation.DATE_VISA_RESPONSABLE || ''
    }
  };
}

function saveReleve(payload) {
  const week = String(payload.semaine || '');
  const id = String(payload.idConducteur || '');
  if (!week || !id) return { success: false, error: 'Conducteur ou semaine manquant' };

  deleteWeekRows(SHEETS.records, id, week);

  const headers = [
    'ID_RELEVE', 'TIMESTAMP', 'ID_CONDUCTEUR', 'NOM', 'PRENOM', 'SEMAINE_DU', 'SEMAINE_AU',
    'ORDRE', 'JOUR', 'DATE', 'ENGIN', 'CHANTIER', 'CONDUCTEUR', 'COMPTEUR_DEBUT',
    'COMPTEUR_FIN', 'TOTAL_HEURES', 'GAZOLE_L', 'OBSERVATIONS'
  ];
  const now = new Date();
  const lignes = (payload.lignes || []).map((l, i) => {
    const debut = number(l.compteurDebut);
    const fin = number(l.compteurFin);
    return {
      ID_RELEVE: Utilities.getUuid(),
      TIMESTAMP: now,
      ID_CONDUCTEUR: id,
      NOM: payload.nom || '',
      PRENOM: payload.prenom || '',
      SEMAINE_DU: week,
      SEMAINE_AU: payload.dateFin || '',
      ORDRE: i + 1,
      JOUR: l.jour || '',
      DATE: l.date || '',
      ENGIN: l.engin || payload.engin || '',
      CHANTIER: l.chantier || payload.chantier || '',
      CONDUCTEUR: l.conducteur || [payload.prenom, payload.nom].filter(Boolean).join(' '),
      COMPTEUR_DEBUT: l.compteurDebut || '',
      COMPTEUR_FIN: l.compteurFin || '',
      TOTAL_HEURES: (isNaN(debut) || isNaN(fin)) ? '' : Math.max(0, fin - debut),
      GAZOLE_L: l.gazole || '',
      OBSERVATIONS: l.observations || ''
    };
  });

  appendObjects(SHEETS.records, lignes, headers);
  upsertValidation(id, week, { ID_CONDUCTEUR: id, SEMAINE_DU: week });
  return { success: true, count: lignes.length };
}

function signConducteur(payload) {
  const id = String(payload.idConducteur || '');
  const week = String(payload.semaine || '');
  upsertValidation(id, week, {
    ID_CONDUCTEUR: id,
    SEMAINE_DU: week,
    VISA_CONDUCTEUR: 'SIGNE',
    DATE_VISA_CONDUCTEUR: new Date()
  });
  return { success: true };
}

function validateResponsable(payload) {
  const id = String(payload.idConducteur || '');
  const week = String(payload.semaine || '');
  upsertValidation(id, week, {
    ID_CONDUCTEUR: id,
    SEMAINE_DU: week,
    VISA_RESPONSABLE: payload.responsable || 'VALIDE',
    DATE_VISA_RESPONSABLE: new Date()
  });
  return { success: true };
}

function deleteWeekRows(sheetName, idConducteur, semaine) {
  const sh = sheet(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0].map(String);
  const idCol = headers.indexOf('ID_CONDUCTEUR');
  const weekCol = headers.indexOf('SEMAINE_DU');
  for (let r = values.length - 1; r >= 1; r--) {
    if (String(values[r][idCol]) === String(idConducteur) && toIso(values[r][weekCol]) === String(semaine)) {
      sh.deleteRow(r + 1);
    }
  }
}

function upsertValidation(idConducteur, semaine, patch) {
  const sh = sheet(SHEETS.validations);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idCol = headers.indexOf('ID_CONDUCTEUR');
  const weekCol = headers.indexOf('SEMAINE_DU');
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(idConducteur) && toIso(values[r][weekCol]) === String(semaine)) {
      headers.forEach((h, i) => {
        if (Object.prototype.hasOwnProperty.call(patch, h)) sh.getRange(r + 1, i + 1).setValue(patch[h]);
      });
      return;
    }
  }
  sh.appendRow(headers.map(h => patch[h] ?? ''));
}

function number(value) {
  if (value === '' || value === null || value === undefined) return NaN;
  return Number(String(value).replace(',', '.'));
}

function toIso(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value).slice(0, 10);
}
