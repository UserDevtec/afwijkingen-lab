# Afwijkingen Lab

[![Deploy to GitHub Pages](https://github.com/UserDevtec/afwijkingen-lab/actions/workflows/deploy.yml/badge.svg)](https://github.com/UserDevtec/afwijkingen-lab/actions/workflows/deploy.yml)

Dashboard voor het verwerken van afwijkingenbestanden (dashboard, database, overzicht) in de browser.

Live demo: [https://userdevtec.github.io/turtle-lab](https://github.com/UserDevtec/afwijkingen-lab)

<img width="1845" height="1028" alt="image" src="https://github.com/user-attachments/assets/c69d13cd-3f29-4fe4-b769-e6dc4e179b81" />

## Features
- Uploaden via klik of drag-and-drop van de drie Excel-bestanden.
- Data ophalen uit het overzicht (achterstallig, concept en unieke actiehouders).
- Email concept genereren en kopieren.
- PowerBI export genereren als nieuwe download.
- Logboek met acties en fouten.

## Starten
1. `npm install`
2. `npm run dev`
3. Open de lokale URL die Vite toont.

## Gebruik
1. Upload het dashboard (`Afwijkingen dashboard.xlsm`).
2. Upload de database (`Afwijkingen database.xlsx`).
3. Upload het overzicht (`Afwijkingen overzicht.xlsx`).
4. Gebruik de knoppen:
   - `Data ophalen` vult de resultaten en actiehouders.
   - `Email opstellen` maakt een email concept.
   - `PowerBI data` maakt een nieuwe download: `Afwijkingen database bijgewerkt.xlsx`.

## Opmerkingen
- Alles draait lokaal in de browser; bestanden worden niet geupload.
- Voor PowerBI export worden kolommen gematcht op kolomkop.
