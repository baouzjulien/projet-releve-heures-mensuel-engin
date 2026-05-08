import fs from 'node:fs/promises';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const workbook = Workbook.create();
const users = workbook.worksheets.add('UTILISATEURS');
const engines = workbook.worksheets.add('ENGINS');
const records = workbook.worksheets.add('RELEVES');
const validations = workbook.worksheets.add('VALIDATIONS');

users.getRange('A1:G1').values = [['ID', 'PIN', 'NOM', 'PRENOM', 'ROLE', 'ACTIF', 'COMMENTAIRE']];
users.getRange('A2:G4').values = [
  ['COND001', '1234', 'Dupont', 'Jean', 'conducteur', 'OUI', 'Exemple conducteur'],
  ['COND002', '2345', 'Martin', 'Sophie', 'conducteur', 'OUI', 'Exemple conducteur'],
  ['RESP001', '9999', 'Responsable', 'Travaux', 'responsable', 'OUI', 'Compte responsable']
];

engines.getRange('A1:F1').values = [['ID_ENGIN', 'DESIGNATION', 'IMMATRICULATION', 'TYPE', 'ACTIF', 'COMMENTAIRE']];
engines.getRange('A2:F6').values = [
  ['PELLE001', 'Pelle 8T', '', 'Pelle', 'OUI', 'Exemple'],
  ['CHARGEUR001', 'Chargeur', '', 'Chargeur', 'OUI', 'Exemple'],
  ['MINI001', 'Mini-pelle', '', 'Mini-pelle', 'OUI', 'Exemple'],
  ['CAMION001', 'Camion benne', '', 'Camion', 'OUI', 'Exemple'],
  ['COMPACT001', 'Compacteur', '', 'Compacteur', 'OUI', 'Exemple']
];

records.getRange('A1:R1').values = [[
  'ID_RELEVE', 'TIMESTAMP', 'ID_CONDUCTEUR', 'NOM', 'PRENOM', 'SEMAINE_DU', 'SEMAINE_AU',
  'ORDRE', 'JOUR', 'DATE', 'ENGIN', 'CHANTIER', 'CONDUCTEUR', 'COMPTEUR_DEBUT',
  'COMPTEUR_FIN', 'TOTAL_HEURES', 'GAZOLE_L', 'OBSERVATIONS'
]];

validations.getRange('A1:F1').values = [[
  'ID_CONDUCTEUR', 'SEMAINE_DU', 'VISA_CONDUCTEUR', 'DATE_VISA_CONDUCTEUR',
  'VISA_RESPONSABLE', 'DATE_VISA_RESPONSABLE'
]];

for (const sheet of [users, engines, records, validations]) {
  sheet.getRange('A1:Z1').format.fill = '#1a1a2e';
  sheet.getRange('A1:Z1').format.font = { color: '#ffffff', bold: true };
}

users.getRange('A:G').columnWidthPx = 150;
engines.getRange('A:F').columnWidthPx = 160;
records.getRange('A:R').columnWidthPx = 145;
validations.getRange('A:F').columnWidthPx = 170;

await fs.mkdir('outputs', { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save('outputs/modele_base_heures_engins.xlsx');

const errors = await workbook.inspect({
  kind: 'match',
  searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A',
  options: { useRegex: true, maxResults: 50 },
  summary: 'final formula error scan'
});
console.log(errors.ndjson);
