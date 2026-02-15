# LiBeR Verslaggenerator

Webapplicatie voor het genereren van Rabobank MRA (Maatschappelijke Relevantie Analyse) verslagen vanuit audio-opnames van intake gesprekken.

## Wat doet het?

1. **Audio → Transcriptie**: Upload een audio-opname van een intake gesprek. Het systeem transcribeert dit via OpenAI Whisper.
2. **Transcriptie → Secties**: Per sectie van het MRA template wordt via GPT-4o de inhoud gegenereerd vanuit de transcriptie.
3. **Bewerken**: Elke sectie kan handmatig worden aangepast in de Word-achtige editor.
4. **Afbeeldingen**: Upload foto's van post-its, whiteboards en diagrammen en plaats ze op de juiste positie.
5. **Export**: Het volledige verslag wordt samengevoegd als Google Doc met de Rabobank MRA opmaak.

## Verslagstructuur (Rabobank MRA Template)

| Pagina | Secties |
|--------|---------|
| 1 | Header (Opdrachtgever, Vereniging, etc.) + Inleiding + Nulmeting (DNA + Quick Scan) |
| 2 | Positie van de organisatie + 0-meting / Stip op de horizon + Ambitie |
| 3 | Advies + Ondersteuning Rabobank + Voorgestelde expert(s) |
| 4+ | Bijlage (foto's van post-its, whiteboards, aantekeningen) |

## Projectstructuur

```
Liber/
├── frontend/           # Webapplicatie (HTML/CSS/JS)
│   ├── index.html      # Hoofdpagina (SPA)
│   ├── css/
│   │   └── style.css   # Stylesheet
│   ├── js/
│   │   ├── api.js      # API communicatie met Apps Script backend
│   │   ├── app.js      # Hoofdcontroller (navigatie, dashboard, create)
│   │   ├── editor.js   # Document editor (per-sectie generatie)
│   │   └── preview.js  # Verslagpreview (print-ready)
│   └── assets/         # Afbeeldingen en iconen
├── apps-script/        # Google Apps Script backend
│   ├── Code.gs         # Volledige backend (v11)
│   └── appsscript.json # Apps Script manifest
└── README.md
```

## Setup

### 1. Google Apps Script Backend

1. Ga naar [script.google.com](https://script.google.com) en maak een nieuw project aan.
2. Kopieer de inhoud van `apps-script/Code.gs` naar het hoofdbestand.
3. Kopieer `apps-script/appsscript.json` naar het manifest (Bestand → Projectinstellingen → Manifest tonen).
4. Voer `setupAPIKeys()` uit en vul je API keys in:
   - `CLOUDCONVERT_TOKEN` – voor audio conversie/splitting
   - `OPENAI_API_KEY` – voor Whisper transcriptie en GPT generatie
5. Voer `setupPlaceholderSheetConfig()` uit om de Google Sheet configuratie op te slaan.
6. Voer `checkConfiguration()` uit om te verifiëren dat alles correct is ingesteld.
7. Deploy als Web App:
   - **Uitvoeren als**: Jouw account
   - **Toegang**: Iedereen (of alleen jezelf)
8. Kopieer de Web App URL.

### 2. Frontend

1. Open `frontend/index.html` in je browser (of host via GitHub Pages / Netlify / etc.)
2. Bij de eerste keer openen wordt gevraagd om de Apps Script Web App URL in te voeren.
3. Plak de URL uit stap 1.8.

### 3. Google Drive Mappen

Zorg dat de volgende mappen bestaan in Google Drive (pas de IDs aan in `Code.gs`):

| Constante | Doel |
|-----------|------|
| `SOURCE_FOLDER_ID` | Map waar nieuwe audio-opnames worden geplaatst (per submap per vergadering) |
| `TARGET_FOLDER_ID` | Werkmap voor geconverteerde MP3 fragmenten |
| `TRANSCRIPT_FOLDER_ID` | Map voor transcripties en draft documenten |
| `ARCHIVE_FOLDER_ID` | Archief voor originele audiobestanden |
| `REPORTS_FOLDER_ID` | Map voor gegenereerde verslagen |

## Technologie

- **Frontend**: Vanilla HTML/CSS/JavaScript (geen frameworks nodig)
- **Backend**: Google Apps Script
- **AI**: OpenAI Whisper (transcriptie) + GPT-4o (sectie generatie)
- **Audio**: CloudConvert API (metadata, conversie, splitting)
- **Opslag**: Google Drive + Google Sheets + Google Docs
- **Export**: Google Docs met DocumentApp formatting

## API Endpoints

### GET

| Action | Parameters | Beschrijving |
|--------|-----------|-------------|
| `status` | – | API health check |
| `listReports` | – | Lijst alle verslagen |
| `getReport` | `reportId` | Details van een verslag |
| `getSections` | `reportId` | Alle secties van een verslag |
| `getSection` | `reportId`, `sectionId` | Eén sectie |
| `getTemplate` | – | MRA template definitie |
| `getTranscriptStatus` | `reportId` | Transcriptie voortgang |

### POST

| Action | Body | Beschrijving |
|--------|------|-------------|
| `createReport` | `{data: {meetingName, headerFields}}` | Nieuw verslag aanmaken |
| `generateSection` | `{reportId, sectionId}` | Eén sectie genereren |
| `regenerateSection` | `{reportId, sectionId}` | Sectie opnieuw genereren |
| `generateAllSections` | `{reportId}` | Alle secties genereren |
| `updateSection` | `{reportId, sectionId, content}` | Sectie handmatig bijwerken |
| `updateHeader` | `{reportId, fields}` | Header velden bijwerken |
| `assembleReport` | `{reportId, headerFields}` | Verslag samenstellen als Google Doc |
| `insertImage` | `{reportId, imageFileId, position}` | Afbeelding invoegen |
| `insertBijlageImages` | `{reportId, imageFileIds}` | Meerdere bijlage afbeeldingen |

## Licentie

Intern gebruik – Rabobank Kring Metropool Regio Amsterdam / LiBeR
