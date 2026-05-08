import fs from 'node:fs/promises';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const workbook = Workbook.create();
const conducteurs = workbook.worksheets.add('CONDUCTEURS');
const releves = workbook.worksheets.add('RELEVES_MENSUELS');

conducteurs.getRange('A1:F1').values = [['ID', 'NOM', 'PRENOM', 'PIN', 'ROLE', 'ACTIF']];
conducteurs.getRange('A2:F7').values = [
  [1, 'BAOUZ', 'JULIEN', 3960, 'employe', 'OUI'],
  [2, 'LANGE', 'DAVID', 5678, 'responsable', 'OUI'],
  [3, 'BOUCHEROT', 'REMY', 5412, 'employe', 'OUI'],
  [4, 'NOM EMPLOYE 3', 'YVAN', 6875, 'employe', 'OUI'],
  [5, 'NOM EMPLOYE 4', 'YACIN', 5143, 'employe', 'OUI'],
  [6, 'PHAN', 'HUGO', 8732, 'employe', 'OUI']
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

for (const sheet of [conducteurs, releves]) {
  sheet.getRange('A1:Z1').format.fill = '#1a1a2e';
  sheet.getRange('A1:Z1').format.font = { color: '#ffffff', bold: true };
}

conducteurs.getRange('A:F').columnWidthPx = 150;
releves.getRange('A:N').columnWidthPx = 165;

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
