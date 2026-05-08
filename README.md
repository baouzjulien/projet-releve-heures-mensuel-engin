# Releve des heures engins - HR Occitanie

Application PWA construite sur le meme modele que l'application de pointage HR :

- `index.html` : application mobile/desktop avec connexion PIN, saisie, signature, validation responsable, export CSV/Excel.
- `Code.gs` : backend Google Apps Script a coller dans un projet Apps Script lie au fichier Google Sheets de base.
- `modele_base_heures_engins.xlsx` : classeur de depart a importer dans Google Drive / Google Sheets.

## Mise en service

1. Stocker le Google Sheets de base dans le dossier Drive `Projet application relevé heures engins`.
2. Ouvrir Extensions > Apps Script depuis le fichier Google Sheets.
3. Coller le contenu de `Code.gs`.
4. Deployer en application web avec acces "Toute personne disposant du lien".
5. Copier l'URL de deploiement dans `index.html`, constante `API_URL`.
6. Publier les fichiers statiques sur GitHub Pages comme le projet de pointage precedent.

## Emplacement Drive

Dossier cible :
https://drive.google.com/drive/folders/1VNgWteByfafSgLGkzmbfZc7bDrxRxZiU

Google Sheets cree :
https://docs.google.com/spreadsheets/d/1BoARCRDQ97pjTY_-kzYJyBwGPOu1G3Zm4plHwVi3fK4/edit

Logo utilise pour l'application :
`Logo_suivi`, present dans le dossier Drive cible.

URL de reference du modele initial :
https://github.com/baouzjulien/Projet-pointage-des-heures-HR
