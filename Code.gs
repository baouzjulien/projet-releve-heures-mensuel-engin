const SHEETS = {
  conducteurs: 'CONDUCTEURS',
  releves: 'RELEVES_MENSUELS'
};

function doGet(e) {
  const action = (e.parameter.action || '').trim();
  try {
    if (action === 'login') return json(login(e.parameter.pin));
    return json({ success: false, error: 'Action inconnue' });
  } catch (err) {
    return json({ success: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (payload.action === 'saveReleveMensuel') return json(saveReleveMensuel(payload));
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

function login(pin) {
  const conducteur = rows(SHEETS.conducteurs)
    .find(c => String(c.PIN) === String(pin) && String(c.ACTIF || 'OUI').toUpperCase() !== 'NON');

  if (!conducteur) return { success: false };

  return {
    success: true,
    user: {
      id: String(conducteur.ID),
      nom: String(conducteur.NOM || ''),
      prenom: String(conducteur.PRENOM || ''),
      role: String(conducteur.ROLE || 'employe').toLowerCase()
    }
  };
}

function saveReleveMensuel(payload) {
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

  sheet(SHEETS.releves).appendRow(headers.map(h => row[h] ?? ''));
  return { success: true, id: row.ID_RELEVE };
}

function numberOrBlank(value) {
  if (value === '' || value === null || value === undefined) return '';
  const n = Number(String(value).replace(',', '.'));
  return isNaN(n) ? value : n;
}
