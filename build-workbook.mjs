import fs from 'node:fs/promises';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const workbook = Workbook.create();
const conducteurs = workbook.worksheets.add('CONDUCTEURS');
const releves = workbook.worksheets.add('RELEVES_MENSUELS');
const audit = workbook.worksheets.add('AUDIT');

conducteurs.getRange('A1:G1').values = [['ID', 'NOM', 'PRENOM', 'PIN', 'ROLE', 'ACTIF', 'EMAIL']];
conducteurs.getRange('A2:G8').values = [
  [1, 'BAOUZ', 'JULIEN', 3960, 'employe', 'OUI', ''],
  [2, 'LANGE', 'DAVID', 5678, 'responsable', 'OUI', ''],
  [3, 'BOUCHEROT', 'REMY', 5412, 'employe', 'OUI', ''],
  [4, 'NOM EMPLOYE 3', 'YVAN', 6875, 'employe', 'OUI', ''],
  [5, 'NOM EMPLOYE 4', 'YACIN', 5143, 'employe', 'OUI', ''],
  [6, 'PHAN', 'HUGO', 8732, 'employe', 'OUI', ''],
  [7, 'FOURNIER', 'GUILLAUME', '', 'responsable', 'OUI', '']
];

releves.getRange('A1:N1').values = [[
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
]];
releves.getRange('O1:U1').values = [[
  'ID_LIGNE',
  'NUM_LIGNE',
  'STATUT',
  'DATE_VALIDATION',
  'ID_RESPONSABLE',
  'NOM_RESPONSABLE',
  'PRENOM_RESPONSABLE'
]];

audit.getRange('A1:I1').values = [[
  'ID_AUDIT',
  'TIMESTAMP',
  'ACTION',
  'ID_UTILISATEUR',
  'NOM',
  'PRENOM',
  'ROLE',
  'DETAILS',
  'USER_AGENT'
]];

for (const sheet of [conducteurs, releves, audit]) {
  sheet.getRange('A1:Z1').format.fill = '#1a1a2e';
  sheet.getRange('A1:Z1').format.font = { color: '#ffffff', bold: true };
}

conducteurs.getRange('A:G').columnWidthPx = 150;
releves.getRange('A:U').columnWidthPx = 165;
audit.getRange('A:I').columnWidthPx = 170;

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
