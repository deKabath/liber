/****************************************************
 * AUDIO TRANSCRIPTIE SCRIPT - VERSIE 11.0 (RESUMABLE)
 * + PLACEHOLDER ANALYSE + PER-SECTIE GENERATIE + DOCX EXPORT
 *
 * Wat doet dit script?
 * FASE 1 – Transcriptie (ongewijzigd t.o.v. v8/v10):
 * - Doorloopt submappen (vergaderingen) in SOURCE_FOLDER_ID
 * - Vindt per submap het eerste audiobestand
 * - Verplaatst origineel naar ARCHIVE_FOLDER_ID
 * - Haalt duur op via CloudConvert metadata
 * - Converteert naar MP3 en splitst in fragmenten indien nodig
 * - Transcribeert fragmenten met OpenAI Whisper, resumable
 * - Slaat deeltranscripties + eindtranscriptie op
 *
 * FASE 2 – Placeholder Analyse (verbeterd in v11):
 * - Analyseert transcriptie met uitgebreide MRA-secties
 * - Nulmeting opgesplitst: Kernwaarden, Sterke eigenschappen,
 *   Aandachtspunten, Quick Scan
 * - Vult placeholders in Google Sheet (wide format)
 *
 * FASE 3 – NIEUW in v11: Per-Sectie Generatie:
 * - generateSection(): genereert 1 specifieke sectie vanuit transcriptie
 * - generateAllSections(): genereert alle secties achter elkaar
 * - Elke sectie krijgt een apart system prompt met de juiste schrijfstijl
 *
 * FASE 4 – NIEUW in v11: Verslag Assembly + Export:
 * - assembleReport(): combineert gegenereerde secties tot 1 document
 * - exportToGoogleDoc(): maakt een Google Doc met MRA-opmaak
 * - Afbeeldingen/bijlagen kunnen op exacte positie geplaatst worden
 *
 * FASE 5 – NIEUW in v11: Web App API (voor frontend):
 * - doGet/doPost endpoints voor de Stitch frontend app
 * - CRUD voor verslagen, secties, afbeeldingen
 * - Status polling voor transcriptie-voortgang
 *
 * Belangrijk:
 * - Dit script is "resumable": Apps Script time-out → later verder.
 ****************************************************/

/****************************************************
 * CONFIG - Script Properties
 ****************************************************/
const PROPS = PropertiesService.getScriptProperties();

const CLOUDCONVERT_TOKEN = PROPS.getProperty('CLOUDCONVERT_TOKEN');
const OPENAI_API_KEY     = PROPS.getProperty('OPENAI_API_KEY');

// Drive folders
const SOURCE_FOLDER_ID     = '1rm4noo1E1C6iTUUmTNX4uWc8Z7Wq-PNl';
const TARGET_FOLDER_ID     = '1n4rBWGdKYamj2xUOiPfA0KfbSJuTZQ38';
const TRANSCRIPT_FOLDER_ID = '1_JPgdA7SUWQ3dy-4k417aDHsIwe2SUWe';
const ARCHIVE_FOLDER_ID    = '1R2qSTI9cV4Yq2iqxrL4eXebcBOMXo-3X';

// NIEUW: map voor gegenereerde verslagen
const REPORTS_FOLDER_ID    = PROPS.getProperty('REPORTS_FOLDER_ID') || TRANSCRIPT_FOLDER_ID;

// CloudConvert API
const CLOUDCONVERT_API = "https://api.cloudconvert.com/v2";

// Fragment limieten
const MAX_DURATION_SEC    = 1800;
const MAX_FILE_SIZE_MB    = 13;
const TARGET_BITRATE_KBPS = 128;

// Whisper settings
const WHISPER_API_URL     = "https://api.openai.com/v1/audio/transcriptions";
const TRANSCRIBE_LANGUAGE = "nl";

// Processed marker
const PROCESSED_PROP = PropertiesService.getScriptProperties();

// Debug
const DEBUG_MODE = true;

// Resumable runtime guard
const RUNTIME_BUDGET_MS = 25 * 60 * 1000;

/****************************************************
 * PLACEHOLDER ANALYSE CONFIG (Rabobank MRA) – UITGEBREID v11
 ****************************************************/
const OPENAI_CHAT_URL = "https://api.openai.com/v1/chat/completions";

const DEFAULT_PLACEHOLDER_SPREADSHEET_ID = "1_hzOEmcALqgGmezTPPqHW_8TxU9UuuTXKgXh76H9nv0";
const PLACEHOLDER_SPREADSHEET_ID = PROPS.getProperty('PLACEHOLDER_SPREADSHEET_ID') || DEFAULT_PLACEHOLDER_SPREADSHEET_ID;

const DEFAULT_PLACEHOLDER_DATA_SHEET_GID = 1000233796;
const PLACEHOLDER_DATA_SHEET_GID = parseInt(PROPS.getProperty('PLACEHOLDER_DATA_SHEET_GID') || String(DEFAULT_PLACEHOLDER_DATA_SHEET_GID), 10);

const PLACEHOLDER_SHEET_NAME     = PROPS.getProperty('PLACEHOLDER_SHEET_NAME') || 'Data';
const PLACEHOLDER_QA_SHEET_NAME  = PROPS.getProperty('PLACEHOLDER_QA_SHEET_NAME') || 'QA';

// OpenAI model voor analyse
const ANALYSIS_MODEL = PROPS.getProperty('ANALYSIS_MODEL') || 'gpt-4o-mini';
// NIEUW: model voor sectie-generatie (krachtiger model aanbevolen)
const GENERATION_MODEL = PROPS.getProperty('GENERATION_MODEL') || 'gpt-4o';

// Chunking
const ANALYSIS_CHUNK_CHARS = parseInt(PROPS.getProperty('ANALYSIS_CHUNK_CHARS') || "12000", 10);
const ANALYSIS_CHUNKS_PER_RUN = parseInt(PROPS.getProperty('ANALYSIS_CHUNKS_PER_RUN') || "2", 10);

/****************************************************
 * RABOBANK MRA TEMPLATE DEFINITIE – v11
 * Gebaseerd op analyse van 6 voorbeeldverslagen
 ****************************************************/

// Placeholder-kolommen (uitgebreid t.o.v. v10)
const PLACEHOLDER_KEYS = [
  '<<naam vereniging>>',
  '<<datum>>',
  '<<contactpersoon>>',
  '<<opgesteld_door>>',
  '<<onderwerp>>',
  '<<Inleiding>>',
  '<<Nulmeting_DNA_kernwaarden>>',
  '<<Nulmeting_DNA_sterke_eigenschappen>>',
  '<<Nulmeting_DNA_aandachtspunten>>',
  '<<Nulmeting_quickscan>>',
  '<<Positie organisatie>>',
  '<<0-meting>>',
  '<<stip_op_horizon>>',
  '<<Ambitie>>',
  '<<Advies>>',
  '<<Ondersteuning_Rabobank>>',
  '<<expert1_naam>>',
  '<<expert1_profiel>>',
  '<<expert2_naam>>',
  '<<expert2_profiel>>',
  '<<Bijlage1>>',
  '<<Bijlage2>>',
  '<<Bijlage3>>'
];

// Template secties met metadata voor generatie
const MRA_SECTIONS = [
  {
    id: 'header',
    title: 'Basisgegevens',
    page: 1,
    fields: ['<<naam vereniging>>', '<<datum>>', '<<contactpersoon>>', '<<opgesteld_door>>', '<<onderwerp>>'],
    generatable: false // handmatig invullen
  },
  {
    id: 'inleiding',
    title: 'Inleiding',
    page: 1,
    fields: ['<<Inleiding>>'],
    generatable: true,
    systemPrompt: `Je schrijft de inleiding van een Rabobank MRA verslag.
MINIMUM LENGTE: 20 woorden. Schrijf minstens 2-3 zinnen.
Formaat: 1-3 korte alinea's.
Begin ALTIJD met: "Op [datum] hebben we een bezoek gebracht aan [type organisatie] [naam] te [plaats]."
Tweede zin: "Aan tafel zaten [aantal] personen: [opsomming rollen]."
Eventueel derde zin over wie namens Rabobank aanschoof.
Schrijfstijl: zakelijk, warm, persoonlijk. Geen opsommingstekens. Nederlands.
Gebruik ALLEEN informatie uit de transcriptie. Verzin NIETS.`
  },
  {
    id: 'nulmeting_dna',
    title: 'Nulmeting – Het DNA',
    page: 1,
    fields: ['<<Nulmeting_DNA_kernwaarden>>', '<<Nulmeting_DNA_sterke_eigenschappen>>', '<<Nulmeting_DNA_aandachtspunten>>'],
    generatable: true,
    systemPrompt: `Je schrijft de DNA-sectie van de Nulmeting in een Rabobank MRA verslag.
MINIMUM LENGTE: 30 woorden.
Begin met de kop "Het DNA werd als volgt beschreven:"
Genereer exact deze 3 sub-items:
- "Kernwaarden (eigen interpretatie nav het gesprek):" gevolgd door 3-5 kernwaarden gescheiden door komma's
- "sterke eigenschappen:" gevolgd door 3-5 sterke punten gescheiden door komma's
- "aandachtspunten:" gevolgd door 3-5 aandachtspunten gescheiden door komma's

De kernwaarden zijn ALTIJD woorden als: sociaal, ongedwongen, saamhorig, familiair, inclusief, laagdrempelig, open, persoonlijk, verbondenheid.
Haal deze uit de sfeer en toon van het gesprek in de transcriptie.
Schrijfstijl: beknopt, geen hele zinnen maar kernbegrippen. Nederlands.`
  },
  {
    id: 'nulmeting_quickscan',
    title: 'Nulmeting – Quick Scan',
    page: 1,
    fields: ['<<Nulmeting_quickscan>>'],
    generatable: true,
    systemPrompt: `Je schrijft de "quick scan" sectie van de Nulmeting in een Rabobank MRA verslag.
MINIMUM LENGTE: 160 woorden. Dit is een uitgebreide sectie.
Begin met: "De door de procesbegeleider gemaakte quick scan:"
Dit zijn observaties over het groepsproces tijdens het gesprek. Focus op:
- Mate van betrokkenheid en openheid
- Of men in stellingen of vragen spreekt
- Of men vanuit ik-vorm of wij-vorm redeneert
- Of men in problemen of kansen denkt
- Of acties persoonlijk/incidenteel of gestructureerd zijn
- Dynamiek: wie nam het woord, luisterde men naar elkaar

Schrijf 5-7 observatie-alinea's. Elke alinea is 2-3 zinnen.
Schrijfstijl: observerend, coachend, direct. Begin veel zinnen met "De aanwezigen..." of "Er is/wordt...".
Nederlands. Gebruik ALLEEN wat uit de transcriptie af te leiden is.`
  },
  {
    id: 'positie',
    title: 'Positie van de organisatie',
    page: 2,
    fields: ['<<Positie organisatie>>'],
    generatable: true,
    systemPrompt: `Je schrijft de sectie "Positie van de vereniging" voor een Rabobank MRA verslag.
MINIMUM LENGTE: 240 woorden. Dit is een van de langste secties van het verslag.
Dit beschrijft de huidige situatie van de organisatie. Behandel (indien relevant):
- Type organisatie, aantal leden, leeftijdsverdeling
- Locatie en accommodatie (staat van onderhoud)
- Activiteiten en frequentie
- Bestuur en organisatiestructuur
- Financiële situatie
- Vrijwilligers en betrokkenheid
- Uitdagingen en knelpunten
- Zichtbaarheid (online/offline)
- Bijzonderheden

FORMAT: Schrijf losse feitelijke zinnen, elke zin op een NIEUWE REGEL.
GEEN opsommingstekens, GEEN bullets, GEEN nummering.
Elke regel is 1-2 zinnen die een feit of observatie beschrijven.
Schrijf minimaal 12-20 regels.
Voorbeeld:
De vereniging telt circa 200 leden waarvan het merendeel 50+.
Het bestuur bestaat uit 5 personen en is op zoek naar verjonging.
De accommodatie is in eigen beheer en verkeert in goede staat.
De financiën staan onder druk, mede door stijgende kosten.
Er is een gebrek aan gecertificeerde trainers.

Schrijfstijl: feitelijk, beschrijvend, zakelijk. Nederlands.
Gebruik ALLEEN informatie uit de transcriptie. Verzin NIETS.`
  },
  {
    id: 'nulmeting_images',
    title: '0-meting / Stip op de horizon',
    page: 2,
    fields: ['<<0-meting>>', '<<stip_op_horizon>>'],
    generatable: false, // dit zijn afbeeldingen die geüpload worden
    allowImages: true,
    layout: 'two-column'
  },
  {
    id: 'ambitie',
    title: 'Ambitie',
    page: 2,
    fields: ['<<Ambitie>>'],
    generatable: true,
    systemPrompt: `Je schrijft de sectie "Ambitie" voor een Rabobank MRA verslag.
MINIMUM LENGTE: 150 woorden. Schrijf een uitgebreide sectie.
Dit beschrijft de toekomstvisie/dromen van de organisatie.
Begin ALTIJD met: "Gekozen is om de stip te zetten op [datum/jaar]." of "Er is een stip gekozen voor [jaar]."
Daarna volgen de concrete ambities als losse zinnen, elk op een nieuwe regel.

FORMAT: Schrijf losse declaratieve zinnen, elke zin op een NIEUWE REGEL.
GEEN opsommingstekens, GEEN bullets, GEEN nummering.
Elke regel is 1-2 zinnen in de tegenwoordige of toekomstige tijd.
Schrijf minimaal 10-15 regels.
Voorbeeld:
De club is financieel gezond.
Er is een helder beleidsplan voor de komende 3 jaar.
De vereniging heeft een actieve jeugdafdeling.
Vrijwilligers worden gewaardeerd en structureel ingezet.
De communicatie is professioneel en bereikt alle leden.

Schrijfstijl: toekomstgericht, positief, concreet. Nederlands.
Gebruik ALLEEN ambities die in de transcriptie benoemd zijn. Verzin NIETS.`
  },
  {
    id: 'advies',
    title: 'Advies',
    page: 3,
    fields: ['<<Advies>>'],
    generatable: true,
    systemPrompt: `Je schrijft de sectie "Advies" voor een Rabobank MRA verslag.
MINIMUM LENGTE: 260 woorden. Dit is de langste sectie van het verslag. Schrijf uitgebreid.
Dit zijn de adviezen van de procesbegeleider.

FORMAT: Schrijf losse alinea's, elke alinea op een NIEUWE REGEL.
GEEN opsommingstekens, GEEN bullets, GEEN nummering.
Elke alinea is 2-4 zinnen. Schrijf minimaal 7-10 alinea's.

VASTE ELEMENTEN die ALTIJD terugkomen (pas de naam van de organisatie aan):
1. Begin ALTIJD met de Einstein quote: "Einstein zei al: als je doet wat je deed dan krijg je wat je kreeg. [Naam] mag durven nieuwe paden te bewandelen en verder te kijken dan de successen en emoties uit het verleden. Ga met de tijd mee en kijk goed om je heen. Neem de tijd om met elkaar hierover in gesprek te gaan en neem de leden hierin stap voor stap mee (zie ontwikkelmodel van Lev Vygotsky). Het ontwikkelen van een "roadmap" is hiervoor een mooi instrument."
2. Noem het principe: "uitspreken – afspreken – aanspreken" als fundament.
3. Verwijs naar Ryan & Deci zelfdeterminatietheorie: autonomie, verbondenheid, competentie. Leg kort uit hoe dit toepasbaar is op de organisatie.
4. Sluit ALTIJD af met: "Zorg dat houding (wat je voorstaat) en gedrag (wat je uiteindelijk in handelen laat zien) congruent zijn. Dan krijgen geformuleerde doelen een krachtige inhoud."

Voeg daarnaast 3-5 specifieke advies-alinea's toe op basis van de transcriptie.
Schrijfstijl: coachend, inspirerend, direct. Nederlands.`
  },
  {
    id: 'ondersteuning',
    title: 'Ondersteuning Rabobank',
    page: 3,
    fields: ['<<Ondersteuning_Rabobank>>'],
    generatable: true,
    systemPrompt: `Je schrijft de sectie "Ondersteuning Rabobank" voor een Rabobank MRA verslag.
MINIMUM LENGTE: 60 woorden.
Dit beschrijft:
1. De initiële hulpvraag/wens van de organisatie
2. De conclusie op basis van het gesprek (wat is de aanbeveling)
3. Eventuele vervolgafspraken

Begin met: "De initiële wens is om aan de slag te gaan met..."
Dan: "Op basis van het gesprek was de conclusie om..."
Schrijf 2-3 alinea's. Kort en bondig.
Schrijfstijl: zakelijk, oplossingsgericht. Nederlands.
Gebruik ALLEEN informatie uit de transcriptie. Verzin NIETS.`
  },
  {
    id: 'experts',
    title: 'Voorgestelde expert(s)',
    page: 3,
    fields: ['<<expert1_naam>>', '<<expert1_profiel>>', '<<expert2_naam>>', '<<expert2_profiel>>'],
    generatable: true,
    systemPrompt: `Je schrijft de sectie "Voorgestelde expert(s)" voor een Rabobank MRA verslag.
MINIMUM LENGTE: 10 woorden.
Als er in de transcriptie een expert wordt voorgesteld, beschrijf dan:
- Naam van de expert
- Korte professionele achtergrond (3-4 zinnen)
- Waarom deze expert geschikt is voor deze opdracht

Als er geen expert benoemd is, schrijf dan:
"Wij komen met een voorstel zodra intern duidelijk is aan welke opdracht men als eerste wilt gaan werken."

Schrijfstijl: professioneel, beknopt. Nederlands.`
  },
  {
    id: 'bijlage',
    title: 'Bijlage',
    page: 4,
    fields: ['<<Bijlage1>>', '<<Bijlage2>>', '<<Bijlage3>>'],
    generatable: false,
    allowImages: true
  }
];

/****************************************************
 * SETUP (1x uitvoeren om keys op te slaan)
 ****************************************************/
function setupAPIKeys() {
  console.log("=== SETUP API KEYS ===");
  console.log("WAARSCHUWING: Verwijder de keys uit deze functie na het uitvoeren!");

  const keys = {
    'CLOUDCONVERT_TOKEN': 'PLAK_HIER_JE_TOKEN',
    'OPENAI_API_KEY': 'PLAK_HIER_JE_KEY'
  };

  for (const [key, value] of Object.entries(keys)) {
    if (value.startsWith('PLAK_HIER')) {
      console.log(`⚠️  ${key}: NIET INGEVULD`);
    } else {
      PROPS.setProperty(key, value);
      console.log(`✓ ${key}: Opgeslagen (${value.substring(0, 10)}...)`);
    }
  }

  console.log("\n✓ Setup voltooid!");
  console.log("BELANGRIJK: Verwijder nu de keys uit deze functie en sla het script op!");
}

function setupPlaceholderSheetConfig() {
  PROPS.setProperty('PLACEHOLDER_SPREADSHEET_ID', DEFAULT_PLACEHOLDER_SPREADSHEET_ID);
  PROPS.setProperty('PLACEHOLDER_DATA_SHEET_GID', String(DEFAULT_PLACEHOLDER_DATA_SHEET_GID));
  PROPS.setProperty('PLACEHOLDER_SHEET_NAME', 'Data');
  PROPS.setProperty('PLACEHOLDER_QA_SHEET_NAME', 'QA');
  PROPS.setProperty('ANALYSIS_MODEL', 'gpt-4o-mini');
  PROPS.setProperty('GENERATION_MODEL', 'gpt-4o');
  PROPS.setProperty('ANALYSIS_CHUNK_CHARS', '12000');
  PROPS.setProperty('ANALYSIS_CHUNKS_PER_RUN', '2');
  console.log("✓ Placeholder sheet config opgeslagen.");
}

/****************************************************
 * HEALTH CHECK
 ****************************************************/
function checkConfiguration() {
  assertConfig_();
  console.log("✓ Config OK");
  console.log("SOURCE_FOLDER_ID:", SOURCE_FOLDER_ID);
  console.log("TARGET_FOLDER_ID:", TARGET_FOLDER_ID);
  console.log("TRANSCRIPT_FOLDER_ID:", TRANSCRIPT_FOLDER_ID);
  console.log("ARCHIVE_FOLDER_ID:", ARCHIVE_FOLDER_ID);
  console.log("REPORTS_FOLDER_ID:", REPORTS_FOLDER_ID);
  console.log("PLACEHOLDER_SPREADSHEET_ID:", PLACEHOLDER_SPREADSHEET_ID);
  console.log("GENERATION_MODEL:", GENERATION_MODEL);
  console.log("Aantal MRA secties:", MRA_SECTIONS.length);
}

/****************************************************
 * ═══════════════════════════════════════════════════
 * FASE 1: TRANSCRIPTIE (ongewijzigd t.o.v. v10)
 * ═══════════════════════════════════════════════════
 ****************************************************/

// [ALLE TRANSCRIPTIE FUNCTIES BLIJVEN IDENTIEK - zie v10]
// checkForNewAudioFiles(), resumePendingWork(), prepareMeetingJob_(),
// finalizeJob_(), findAudioFileInFolder(), moveFileToArchive_(),
// getMetadataViaCloudConvert(), convertDirect(), splitAudioUltraRobust(),
// transcribeWithWhisper_(), etc.
//
// ENIGE WIJZIGING: finalizeJob_() roept nu ook queueSectionGeneration_() aan

function checkForNewAudioFiles() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { console.log("Lock niet verkregen, overslaan"); return; }

  const started = Date.now();
  try {
    console.log("=== START CHECK V11.0 ===");
    assertConfig_();

    const pending = findNextPendingJob_();
    if (pending) {
      console.log("Pending job gevonden → resumePendingWork()");
      scheduleResume_(1);
      return;
    }

    const sourceFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
    const subFolders = sourceFolder.getFolders();

    while (subFolders.hasNext()) {
      if (Date.now() - started > (RUNTIME_BUDGET_MS - 60 * 1000)) {
        scheduleResume_(1);
        return;
      }

      const meetingFolder = subFolders.next();
      const meetingFolderId = meetingFolder.getId();
      const meetingName = meetingFolder.getName();

      if (PROCESSED_PROP.getProperty("folder_" + meetingFolderId) === "done") continue;

      const jobKey = jobKey_(meetingFolderId);
      const jobState = PROPS.getProperty(jobKey + "_state");
      if (jobState && jobState !== "done") continue;

      const audioFile = findAudioFileInFolder(meetingFolder);
      if (!audioFile) continue;

      console.log("→ Audiobestand gevonden:", audioFile.getName());
      prepareMeetingJob_(meetingFolder, meetingName, audioFile);
      scheduleResume_(1);
      return;
    }

    console.log("=== CHECK VOLTOOID: geen nieuwe audio gevonden ===");
  } finally {
    lock.releaseLock();
  }
}

function resumePendingWork() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { console.log("Lock niet verkregen, overslaan"); return; }

  const started = Date.now();
  try {
    assertConfig_();
    const job = findNextPendingJob_();
    if (!job) { cleanupResumeTriggers_(); return; }

    console.log("=== RESUME JOB ===", job.meetingName, "State:", job.state);

    const transcriptFolder = DriveApp.getFolderById(TRANSCRIPT_FOLDER_ID);
    const partsFolder = createOrGetPartsFolder_(transcriptFolder, job.meetingName);

    let docId = job.docId;
    if (!docId) {
      const doc = createTranscriptDoc_(job.meetingName + " - Transcriptie", "", transcriptFolder);
      docId = doc.getId();
      PROPS.setProperty(job.key + "_docId", docId);
    }

    let i = job.nextIndex;
    while (i < job.fragments.length) {
      if (Date.now() - started > (RUNTIME_BUDGET_MS - 90 * 1000)) {
        PROPS.setProperty(job.key + "_nextIndex", String(i));
        PROPS.setProperty(job.key + "_state", "transcribing");
        scheduleResume_(1);
        return;
      }

      const frag = job.fragments[i];
      console.log(`Transcriberen ${i + 1}/${job.fragments.length}: ${frag.name}`);

      const fragFile = DriveApp.getFileById(frag.fileId);
      const text = transcribeWithWhisper_(fragFile) || "";

      const startTC = toTimecode_(frag.startSec);
      const endTC = toTimecode_(frag.endSec);
      const partTitle = `${job.meetingName} - Deel ${frag.index} (${startTC}-${endTC})`;
      const partBody = `Vergadering: ${job.meetingName}\nDeel: ${frag.index}\nStart: ${startTC}\nEind: ${endTC}\n\n${text}\n`;

      createTranscriptTxt(partTitle, partBody, partsFolder);
      appendToDoc_(docId, `\n\n---\nDeel ${frag.index} (${startTC} - ${endTC})\n---\n\n${text}\n`);

      i++;
      PROPS.setProperty(job.key + "_nextIndex", String(i));
      PROPS.setProperty(job.key + "_state", "transcribing");
      Utilities.sleep(900);
    }

    console.log("Alle fragmenten getranscribeerd → finaliseren");
    finalizeJob_(job, partsFolder, docId);

    const next = findNextPendingJob_();
    if (next) scheduleResume_(1);
    else cleanupResumeTriggers_();

    const nextAnalysis = findNextPendingAnalysisJob_();
    if (nextAnalysis) scheduleAnalysisResume_(1);

  } catch (err) {
    console.log("!!! FOUT in resumePendingWork:", err && err.stack ? err.stack : err);
    const job = findNextPendingJob_();
    if (job) {
      PROPS.setProperty(job.key + "_state", "error");
      PROPS.setProperty(job.key + "_error", String(err));
    }
    scheduleResume_(5);
  } finally {
    lock.releaseLock();
  }
}

function prepareMeetingJob_(meetingFolder, meetingName, audioFile) {
  const meetingFolderId = meetingFolder.getId();
  const key = jobKey_(meetingFolderId);

  PROPS.setProperty(key + "_state", "preparing");
  PROPS.setProperty(key + "_meetingName", meetingName);
  PROPS.setProperty(key + "_meetingFolderId", meetingFolderId);
  PROPS.setProperty(key + "_createdAt", new Date().toISOString());
  PROPS.setProperty(key + "_nextIndex", "0");

  const archivedId = moveFileToArchive_(audioFile, meetingFolder);
  PROPS.setProperty(key + "_archivedOriginalFileId", archivedId);

  const metadata = getMetadataViaCloudConvert(DriveApp.getFileById(archivedId));
  if (!metadata) throw new Error("Metadata is leeg.");

  const duration = parseDuration(
    metadata.Duration || metadata.TrackDuration || metadata.duration ||
    (metadata.MediaDuration ? metadata.MediaDuration : null)
  );
  PROPS.setProperty(key + "_durationSec", String(duration));

  const maxFragmentDuration = calculateMaxFragmentDuration();
  let fragments = [];
  let workFolderId = "";

  const archivedFile = DriveApp.getFileById(archivedId);

  if (duration <= maxFragmentDuration) {
    const result = convertDirect(archivedFile);
    workFolderId = result.subFolder.getId();
    fragments = [{ fileId: result.file.getId(), name: result.file.getName(), startSec: 0, endSec: duration, index: 1 }];
  } else {
    const result = splitAudioUltraRobust(archivedFile, duration, maxFragmentDuration);
    workFolderId = result.subFolder.getId();
    fragments = result.fragments.map(f => ({ fileId: f.file.getId(), name: f.file.getName(), startSec: f.startSec, endSec: f.endSec, index: f.index }));
  }

  PROPS.setProperty(key + "_workFolderId", workFolderId);
  PROPS.setProperty(key + "_fragmentsJson", JSON.stringify(fragments));

  const transcriptFolder = DriveApp.getFolderById(TRANSCRIPT_FOLDER_ID);
  const doc = createTranscriptDoc_(meetingName + " - Transcriptie", `Vergadering: ${meetingName}\n\n`, transcriptFolder);
  PROPS.setProperty(key + "_docId", doc.getId());
  PROPS.setProperty(key + "_state", "ready");
  console.log("✓ Job voorbereid. Fragments:", fragments.length);
}

function finalizeJob_(job, partsFolder, docId) {
  const transcriptFolder = DriveApp.getFolderById(TRANSCRIPT_FOLDER_ID);

  const partFiles = partsFolder.getFiles();
  const parts = [];
  while (partFiles.hasNext()) parts.push(partFiles.next());
  parts.sort((a, b) => (a.getName() > b.getName() ? 1 : -1));

  appendToDoc_(docId, "\n\n========================\nDEELTRANSCRIPTIES\n========================\n\n");
  for (const f of parts) {
    const content = f.getBlob().getDataAsString("UTF-8");
    appendToDoc_(docId, `\n--- ${f.getName().replace(/\.txt$/i, "")} ---\n\n${content}\n`);
  }

  const doc = DocumentApp.openById(docId);
  const fullText = doc.getBody().getText();
  const finalTxt = createTranscriptTxt(job.meetingName + " - Transcriptie", fullText, transcriptFolder);

  cleanupMP3FragmentsByIds_(job.fragments.map(x => x.fileId), job.workFolderId);

  PROPS.setProperty(job.key + "_state", "done");
  PROCESSED_PROP.setProperty("folder_" + job.meetingFolderId, "done");
  console.log("✓ Transcriptie job afgerond:", job.meetingName);

  // Queue placeholder analyse
  try {
    queuePlaceholderAnalysis_(job.meetingFolderId, job.meetingName, finalTxt.getId());
    scheduleAnalysisResume_(1);
  } catch (e) {
    console.log("⚠️ Placeholder analyse niet gequeued:", e);
  }

  // NIEUW v11: Queue sectie-generatie
  try {
    queueSectionGeneration_(job.meetingFolderId, job.meetingName, finalTxt.getId());
  } catch (e) {
    console.log("⚠️ Sectie-generatie niet gequeued:", e);
  }
}

/****************************************************
 * ═══════════════════════════════════════════════════
 * FASE 3: PER-SECTIE GENERATIE (NIEUW v11)
 * ═══════════════════════════════════════════════════
 ****************************************************/

/**
 * Genereert een enkele sectie vanuit de transcriptie
 * @param {string} reportId - Report identifier (meetingFolderId)
 * @param {string} sectionId - Sectie ID uit MRA_SECTIONS (bijv. 'inleiding', 'ambitie')
 * @param {Object} overrides - Optionele overrides: {transcriptText, meetingName, extraContext}
 * @returns {Object} {sectionId, title, content, generatedAt}
 */
function generateSection(reportId, sectionId, overrides) {
  overrides = overrides || {};
  console.log(`=== GENERATE SECTION: ${sectionId} ===`);

  const section = MRA_SECTIONS.find(s => s.id === sectionId);
  if (!section) throw new Error("Onbekende sectie: " + sectionId);
  if (!section.generatable) throw new Error("Sectie " + sectionId + " is niet genereerbaar (handmatig invullen).");

  // Haal transcriptie op
  let transcriptText = overrides.transcriptText;
  if (!transcriptText) {
    const reportData = loadReportData_(reportId);
    if (!reportData || !reportData.transcriptTxtFileId) throw new Error("Geen transcriptie gevonden voor report " + reportId);
    transcriptText = DriveApp.getFileById(reportData.transcriptTxtFileId).getBlob().getDataAsString("UTF-8");
  }

  const meetingName = overrides.meetingName || loadReportData_(reportId).meetingName || "Onbekend";

  // Bouw prompt
  const systemMessage = section.systemPrompt;
  const MAX_TRANSCRIPT_CHARS = 30000;
  const truncated = transcriptText.length > MAX_TRANSCRIPT_CHARS;
  if (truncated) {
    console.log(`⚠️ Transcriptie afgekapt: ${transcriptText.length} → ${MAX_TRANSCRIPT_CHARS} tekens voor sectie ${sectionId}`);
  }
  const userMessage = [
    "VERENIGING / ORGANISATIE:",
    meetingName,
    "",
    overrides.extraContext ? ("EXTRA CONTEXT:\n" + overrides.extraContext + "\n") : "",
    "TRANSCRIPTIE:",
    transcriptText.substring(0, MAX_TRANSCRIPT_CHARS),
    truncated ? "\n[... transcriptie afgekapt, gebruik de bovenstaande informatie ...]" : ""
  ].join("\n");

  const payload = {
    model: GENERATION_MODEL,
    temperature: 0.3,
    messages: [
      { role: "system", content: systemMessage },
      { role: "user", content: userMessage }
    ]
  };

  let content = callOpenAIChat_(payload);

  // Controleer minimum woordaantal
  const minWords = extractMinWords_(systemMessage);
  const wordCount = content.trim().split(/\s+/).length;
  if (minWords > 0 && wordCount < minWords) {
    console.log(`⚠️ Sectie ${sectionId}: ${wordCount} woorden < minimum ${minWords}. Hergenereren...`);
    payload.messages.push({ role: "assistant", content: content });
    payload.messages.push({ role: "user", content: `Je tekst is ${wordCount} woorden maar het minimum is ${minWords} woorden. Schrijf de sectie opnieuw, langer en gedetailleerder. Gebruik meer informatie uit de transcriptie. Minimaal ${minWords} woorden.` });
    content = callOpenAIChat_(payload);
    const newWordCount = content.trim().split(/\s+/).length;
    console.log(`→ Hergenereerd: ${newWordCount} woorden (was ${wordCount})`);
  }

  // Sla op
  const resultKey = `report_${reportId}_section_${sectionId}`;
  const result = {
    sectionId: sectionId,
    title: section.title,
    content: content,
    generatedAt: new Date().toISOString(),
    model: GENERATION_MODEL,
    transcriptChars: transcriptText.length,
    truncated: truncated
  };

  PROPS.setProperty(resultKey, JSON.stringify(result));
  console.log(`✓ Sectie ${sectionId} gegenereerd (${content.length} chars)`);

  return result;
}

/**
 * Genereert alle genereerbare secties voor een report
 * Resumable: slaat voortgang op per sectie
 */
function generateAllSections(reportId) {
  const started = Date.now();
  const generatableSections = MRA_SECTIONS.filter(s => s.generatable);

  console.log(`=== GENERATE ALL SECTIONS (${generatableSections.length}) ===`);

  // Laad transcriptie 1x
  const reportData = loadReportData_(reportId);
  if (!reportData || !reportData.transcriptTxtFileId) {
    throw new Error("Geen transcriptie gevonden voor report " + reportId + ". Upload eerst een audio-opname.");
  }
  const transcriptText = DriveApp.getFileById(reportData.transcriptTxtFileId).getBlob().getDataAsString("UTF-8");

  const results = [];
  for (const section of generatableSections) {
    // Check of al gegenereerd
    const existingKey = `report_${reportId}_section_${section.id}`;
    const existing = PROPS.getProperty(existingKey);
    if (existing) {
      console.log(`→ ${section.id}: al gegenereerd, overslaan`);
      results.push(JSON.parse(existing));
      continue;
    }

    // Budget check
    if (Date.now() - started > (RUNTIME_BUDGET_MS - 120 * 1000)) {
      console.log("Tijdbudget bijna op → pauzeren");
      scheduleSectionGenerationResume_(reportId, 1);
      return { partial: true, completed: results.length, total: generatableSections.length, results };
    }

    try {
      const result = generateSection(reportId, section.id, { transcriptText, meetingName: reportData.meetingName });
      results.push(result);
      Utilities.sleep(1000); // rate limiting
    } catch (err) {
      console.log(`⚠️ Fout bij sectie ${section.id}:`, err);
      results.push({ sectionId: section.id, error: String(err) });
    }
  }

  console.log(`✓ Alle ${results.length} secties verwerkt`);
  return { partial: false, completed: results.length, total: generatableSections.length, results };
}

/**
 * Regenereert een sectie (overschrijft bestaande)
 */
function regenerateSection(reportId, sectionId, extraContext) {
  const key = `report_${reportId}_section_${sectionId}`;
  PROPS.deleteProperty(key); // verwijder oude
  return generateSection(reportId, sectionId, { extraContext: extraContext });
}

/****************************************************
 * SECTIE GENERATIE QUEUE & RESUME
 ****************************************************/
function queueSectionGeneration_(meetingFolderId, meetingName, transcriptTxtFileId) {
  const key = "secgen_" + meetingFolderId;
  PROPS.setProperty(key + "_state", "pending");
  PROPS.setProperty(key + "_meetingName", meetingName);
  PROPS.setProperty(key + "_meetingFolderId", meetingFolderId);
  PROPS.setProperty(key + "_transcriptTxtFileId", transcriptTxtFileId);
  console.log("✓ Sectie-generatie queued:", meetingName);
}

function scheduleSectionGenerationResume_(reportId, minutesFromNow) {
  const ms = Math.max(1, minutesFromNow || 1) * 60 * 1000;
  ScriptApp.newTrigger("resumeSectionGeneration")
    .timeBased()
    .after(ms)
    .create();
}

function resumeSectionGeneration() {
  const all = PROPS.getProperties();
  const keys = Object.keys(all)
    .filter(k => k.startsWith("secgen_") && k.endsWith("_state"))
    .map(k => k.replace(/_state$/, ""));

  for (const key of keys) {
    const state = PROPS.getProperty(key + "_state");
    if (state === "pending" || state === "generating") {
      const meetingFolderId = PROPS.getProperty(key + "_meetingFolderId");
      PROPS.setProperty(key + "_state", "generating");

      try {
        const result = generateAllSections(meetingFolderId);
        if (!result.partial) {
          PROPS.setProperty(key + "_state", "done");
          console.log("✓ Alle secties gegenereerd voor:", PROPS.getProperty(key + "_meetingName"));
        }
      } catch (err) {
        console.log("!!! Fout bij sectie-generatie:", err);
        PROPS.setProperty(key + "_state", "error");
        PROPS.setProperty(key + "_error", String(err));
      }
      return; // 1 job per run
    }
  }
}

/****************************************************
 * ═══════════════════════════════════════════════════
 * FASE 4: VERSLAG ASSEMBLY + EXPORT (NIEUW v11)
 * ═══════════════════════════════════════════════════
 ****************************************************/

/**
 * Stelt het volledige verslag samen uit gegenereerde secties
 * en maakt een Google Doc met de juiste opmaak.
 *
 * Ondersteunt:
 * - Bullet points (•) in Positie, Ambitie en Advies secties
 * - Onderstreepte kernbegrippen via _tekst_ markdown
 * - Afbeeldingen in 0-meting (twee-kolom) en Bijlage (meerdere pagina's)
 * - LiBeR/Rabobank opmaak met oranje accenten
 */
function assembleReport(reportId, headerFields) {
  console.log("=== ASSEMBLE REPORT ===");

  headerFields = headerFields || {};
  const reportData = loadReportData_(reportId);

  // Verzamel alle secties
  const sections = {};
  for (const section of MRA_SECTIONS) {
    const key = `report_${reportId}_section_${section.id}`;
    const data = PROPS.getProperty(key);
    if (data) {
      sections[section.id] = JSON.parse(data);
    }
  }

  // Maak Google Doc
  const reportTitle = `Rabobank MRA - verslag intake ${headerFields['<<naam vereniging>>'] || reportData.meetingName}`;
  const reportsFolder = DriveApp.getFolderById(REPORTS_FOLDER_ID);
  const doc = DocumentApp.create(reportTitle);
  const body = doc.getBody();

  // Stel paginamarges in
  body.setMarginTop(50);
  body.setMarginBottom(50);
  body.setMarginLeft(60);
  body.setMarginRight(60);

  // Kleurdefinities
  const ORANJE = '#E87722';
  const DONKERGRIJS = '#333333';

  // === PAGINA 1 ===
  // Header met basisgegevens
  const headerTable = body.appendTable([
    ['Opdrachtgever:', headerFields['<<opdrachtgever>>'] || 'Rabobank Kring Metropool Regio Amsterdam'],
    ['Vereniging:', headerFields['<<naam vereniging>>'] || ''],
    ['Onderwerp:', headerFields['<<onderwerp>>'] || ''],
    ['Contact met:', headerFields['<<contactpersoon>>'] || ''],
    ['Opgesteld door:', headerFields['<<opgesteld_door>>'] || 'Lutger Brenninkmeijer']
  ]);

  // Style header table
  headerTable.setBorderWidth(0);
  for (let r = 0; r < headerTable.getNumRows(); r++) {
    headerTable.getRow(r).getCell(0).getChild(0).asText().setItalic(true).setForegroundColor(DONKERGRIJS);
    headerTable.getRow(r).getCell(1).getChild(0).asText().setForegroundColor(DONKERGRIJS);
  }

  // Oranje scheidingslijn (gesimuleerd met gekleurde paragraaf)
  const divider = body.appendParagraph('____________________________________________________');
  divider.getChild(0).asText().setForegroundColor(ORANJE);

  // Inleiding
  if (sections.inleiding) {
    appendSectionHeader_(body, 'Inleiding', ORANJE);
    appendPlainText_(body, sections.inleiding.content);
  }

  // Nulmeting
  appendSectionHeader_(body, 'Nulmeting', ORANJE);

  if (sections.nulmeting_dna) {
    appendPlainText_(body, sections.nulmeting_dna.content);
  }

  if (sections.nulmeting_quickscan) {
    body.appendParagraph(''); // spacing
    appendPlainText_(body, sections.nulmeting_quickscan.content);
  }

  // === PAGINA 2 ===
  body.appendPageBreak();

  // Positie van de organisatie (losse zinnen per regel)
  if (sections.positie) {
    appendSectionHeader_(body, 'Positie van de organisatie', ORANJE);
    appendPlainText_(body, sections.positie.content);
  }

  // 0-meting en Stip op de horizon (twee-kolom tabel voor afbeeldingen)
  body.appendParagraph(''); // spacing
  const meetingImagesTable = body.appendTable([
    ['0-meting', 'Stip op de horizon']
  ]);
  meetingImagesTable.setBorderWidth(0);
  meetingImagesTable.getRow(0).getCell(0).getChild(0).asText().setBold(true).setItalic(true).setFontSize(11).setForegroundColor(ORANJE);
  meetingImagesTable.getRow(0).getCell(1).getChild(0).asText().setBold(true).setItalic(true).setFontSize(11).setForegroundColor(ORANJE);

  // Ruimte voor afbeeldingen (worden later via insertImageInReport ingevuld)
  const imageRow = meetingImagesTable.appendTableRow();
  imageRow.appendTableCell('[Afbeelding 0-meting]');
  imageRow.appendTableCell('[Afbeelding Stip op de horizon]');

  // Ambitie (losse zinnen per regel)
  body.appendParagraph(''); // spacing
  if (sections.ambitie) {
    appendSectionHeader_(body, 'Ambitie', ORANJE);
    appendPlainText_(body, sections.ambitie.content);
  }

  // === PAGINA 3 ===
  body.appendPageBreak();

  // Advies (losse alinea's per regel)
  if (sections.advies) {
    appendSectionHeader_(body, 'Advies', ORANJE);
    appendPlainText_(body, sections.advies.content);
  }

  // Ondersteuning Rabobank
  if (sections.ondersteuning) {
    body.appendParagraph(''); // spacing
    appendSectionHeader_(body, 'Ondersteuning Rabobank', ORANJE);
    appendPlainText_(body, sections.ondersteuning.content);
  }

  // Voorgestelde expert(s)
  if (sections.experts) {
    body.appendParagraph(''); // spacing
    appendSectionHeader_(body, 'Voorgestelde expert(s)', ORANJE);
    appendPlainText_(body, sections.experts.content);
  }

  // === PAGINA 4+ : BIJLAGE ===
  body.appendPageBreak();

  const bijlageHeader = body.appendParagraph('Bijlage:');
  bijlageHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  bijlageHeader.getChild(0).asText().setBold(true).setItalic(true).setFontSize(13).setForegroundColor(ORANJE);

  // Bijlage afbeeldingen worden via insertImageInReport() toegevoegd
  // Placeholder tekst
  const bijlagePlaceholder = body.appendParagraph('[Bijlage-afbeeldingen worden hier geplaatst via de editor]');
  bijlagePlaceholder.getChild(0).asText().setItalic(true).setForegroundColor('#999999').setFontSize(9);

  doc.saveAndClose();

  // Verplaats naar Reports folder
  const docFile = DriveApp.getFileById(doc.getId());
  reportsFolder.addFile(docFile);
  try { DriveApp.getRootFolder().removeFile(docFile); } catch(e) {}

  console.log("✓ Verslag samengesteld:", reportTitle);
  console.log("Doc URL:", doc.getUrl());

  // Sla report doc ID op
  PROPS.setProperty(`report_${reportId}_docId`, doc.getId());

  return {
    docId: doc.getId(),
    docUrl: doc.getUrl(),
    title: reportTitle
  };
}

/****************************************************
 * DOCUMENT OPMAAK HELPERS
 ****************************************************/

/**
 * Voegt een sectie-koptekst toe (bold, italic, oranje)
 */
function appendSectionHeader_(body, title, color) {
  const header = body.appendParagraph(title);
  header.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  header.getChild(0).asText().setBold(true).setItalic(true).setFontSize(13).setForegroundColor(color || '#E87722');
  return header;
}

/**
 * Voegt platte tekst toe (paragraaf per alinea)
 */
function appendPlainText_(body, content) {
  if (!content) return;
  const paragraphs = content.split('\n').filter(line => line.trim());
  for (const para of paragraphs) {
    const p = body.appendParagraph(para.trim());
    p.setFontSize(10);
  }
}

/**
 * Verwerkt content met bullet points (•).
 * Lijnen die beginnen met "•" worden als individuele bullet-paragrafen toegevoegd.
 * Overige tekst wordt als gewone alinea's toegevoegd.
 */
function appendBulletContent_(body, content) {
  if (!content) return;
  const lines = content.split('\n');

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    if (trimmed.startsWith('•') || trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
      // Bullet point - verwijder het bullet-teken en voeg als list item toe
      let text = trimmed.replace(/^[•\-\*]\s*/, '').trim();
      const listItem = body.appendListItem(text);
      listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
      listItem.setFontSize(10);
    } else {
      // Gewone tekst (bijv. intro-zin)
      const p = body.appendParagraph(trimmed);
      p.setFontSize(10);
    }
  }
}

/**
 * Verwerkt content met bullet points (•) EN _onderstreping_ voor kernbegrippen.
 * Markdown-style _tekst_ wordt omgezet naar onderstreepte tekst in Google Docs.
 */
function appendBulletContentWithUnderline_(body, content) {
  if (!content) return;
  const lines = content.split('\n');

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    let element;
    if (trimmed.startsWith('•') || trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
      let text = trimmed.replace(/^[•\-\*]\s*/, '').trim();
      element = body.appendListItem('');
      element.setGlyphType(DocumentApp.GlyphType.BULLET);
      element.setFontSize(10);
      applyUnderlineMarkdown_(element, text);
    } else {
      element = body.appendParagraph('');
      element.setFontSize(10);
      applyUnderlineMarkdown_(element, trimmed);
    }
  }
}

/**
 * Extraheert het minimum woordaantal uit een system prompt.
 * Zoekt naar "MINIMUM LENGTE: X woorden" in de prompt tekst.
 * Retourneert 0 als er geen minimum gevonden wordt.
 */
function extractMinWords_(systemPrompt) {
  const match = systemPrompt.match(/MINIMUM LENGTE:\s*(\d+)\s*woorden/i);
  return match ? parseInt(match[1], 10) : 0;
}

/**
 * Past _underline_ markdown toe op een paragraaf/listitem.
 * Tekst tussen _underscores_ wordt onderstreept weergegeven.
 */
function applyUnderlineMarkdown_(element, text) {
  // Splits op _markdown_ patronen
  const parts = text.split(/(_[^_]+_)/g);

  for (const part of parts) {
    if (part.startsWith('_') && part.endsWith('_') && part.length > 2) {
      // Onderstreepte tekst
      const innerText = part.slice(1, -1);
      const textEl = element.appendText(innerText);
      textEl.setUnderline(true);
    } else if (part) {
      // Gewone tekst
      element.appendText(part);
    }
  }
}

/**
 * Voegt een afbeelding toe aan een specifieke positie in het verslag
 * @param {string} reportId
 * @param {string} imageFileId - Google Drive file ID van de afbeelding
 * @param {string} position - 'nulmeting_links' | 'nulmeting_rechts' | 'bijlage'
 * @param {number} bijlageIndex - (optioneel) index voor bijlage positie
 */
function insertImageInReport(reportId, imageFileId, position, bijlageIndex) {
  const docId = PROPS.getProperty(`report_${reportId}_docId`);
  if (!docId) throw new Error("Geen verslag document gevonden voor report " + reportId);

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const image = DriveApp.getFileById(imageFileId).getBlob();

  if (position === 'nulmeting_links' || position === 'nulmeting_rechts') {
    // Zoek de 0-meting tabel
    const tables = body.getTables();
    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];
      const firstCell = table.getRow(0).getCell(0).getText();
      if (firstCell.includes('0-meting')) {
        const colIndex = position === 'nulmeting_links' ? 0 : 1;
        const row = table.getNumRows() > 1 ? table.getRow(1) : table.appendTableRow();
        const cell = row.getCell(colIndex);
        cell.clear();
        cell.appendImage(image).setWidth(280);
        break;
      }
    }
  } else if (position === 'bijlage') {
    // Verwijder de placeholder tekst als die er nog staat
    removeBijlagePlaceholder_(body);
    // Voeg afbeelding toe aan einde van document (bijlage sectie)
    // Afbeelding breedte: full-width (480px) of half-width (230px) voor side-by-side
    const width = (bijlageIndex !== undefined && bijlageIndex % 2 === 1) ? 230 : 480;
    body.appendImage(image).setWidth(width);
    body.appendParagraph(''); // spacing
  } else if (position === 'bijlage_sidebyside') {
    // Twee afbeeldingen naast elkaar in een tabel (zoals in voorbeeldverslagen)
    removeBijlagePlaceholder_(body);
    const table = body.appendTable();
    table.setBorderWidth(0);
    const row = table.appendTableRow();
    const cell1 = row.appendTableCell('');
    cell1.appendImage(image).setWidth(230);
    // Tweede afbeelding wordt via een tweede call toegevoegd
    row.appendTableCell('[volgende afbeelding]');
  }

  doc.saveAndClose();
  console.log(`✓ Afbeelding geplaatst op positie: ${position}`);
}

/**
 * Voegt meerdere bijlage-afbeeldingen toe in één keer.
 * Plaatst ze afwisselend full-width en side-by-side.
 * @param {string} reportId
 * @param {string[]} imageFileIds - Array van Google Drive file IDs
 */
function insertBijlageImages(reportId, imageFileIds) {
  if (!imageFileIds || imageFileIds.length === 0) return;

  const docId = PROPS.getProperty(`report_${reportId}_docId`);
  if (!docId) throw new Error("Geen verslag document gevonden voor report " + reportId);

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  removeBijlagePlaceholder_(body);

  let i = 0;
  while (i < imageFileIds.length) {
    if (i + 1 < imageFileIds.length) {
      // Twee afbeeldingen naast elkaar in een tabel
      const table = body.appendTable();
      table.setBorderWidth(0);
      const row = table.appendTableRow();

      const img1 = DriveApp.getFileById(imageFileIds[i]).getBlob();
      const cell1 = row.appendTableCell('');
      cell1.clear();
      cell1.appendImage(img1).setWidth(230);

      const img2 = DriveApp.getFileById(imageFileIds[i + 1]).getBlob();
      const cell2 = row.appendTableCell('');
      cell2.clear();
      cell2.appendImage(img2).setWidth(230);

      body.appendParagraph(''); // spacing
      i += 2;
    } else {
      // Laatste afbeelding full-width
      const img = DriveApp.getFileById(imageFileIds[i]).getBlob();
      body.appendImage(img).setWidth(400);
      body.appendParagraph(''); // spacing
      i++;
    }
  }

  doc.saveAndClose();
  console.log(`✓ ${imageFileIds.length} bijlage-afbeelding(en) geplaatst`);
}

/**
 * Verwijdert de bijlage placeholder tekst
 */
function removeBijlagePlaceholder_(body) {
  const numChildren = body.getNumChildren();
  for (let i = numChildren - 1; i >= 0; i--) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const text = child.asParagraph().getText();
      if (text.includes('Bijlage-afbeeldingen worden hier geplaatst')) {
        body.removeChild(child);
        break;
      }
    }
  }
}

/****************************************************
 * ═══════════════════════════════════════════════════
 * FASE 5: WEB APP API (NIEUW v11)
 * Voor de Stitch frontend applicatie
 * ═══════════════════════════════════════════════════
 ****************************************************/

/**
 * GET endpoint - haalt data op
 * Params: action, reportId, sectionId
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = params.action || 'status';

  try {
    let result;

    switch (action) {
      case 'status':
        result = { status: 'ok', version: '11.0', message: 'LiBeR Verslaggenerator API is actief' };
        break;

      case 'listReports':
        result = listAllReports_();
        break;

      case 'getReport':
        result = getReportDetails_(params.reportId);
        break;

      case 'getSection':
        result = getSectionContent_(params.reportId, params.sectionId);
        break;

      case 'getSections':
        result = getAllSections_(params.reportId);
        break;

      case 'getTemplate':
        result = { sections: MRA_SECTIONS, placeholderKeys: PLACEHOLDER_KEYS };
        break;

      case 'getTranscriptStatus':
        result = getTranscriptStatus_(params.reportId);
        break;

      default:
        result = { error: 'Onbekende actie: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * POST endpoint - mutaties
 * Body: {action, reportId, sectionId, data}
 */
function doPost(e) {
  try {
    let body = {};
    if (e && e.postData) {
      body = JSON.parse(e.postData.contents);
    }

    const action = body.action || 'webhook';
    let result;

    switch (action) {
      case 'webhook':
        // Originele Jotform webhook
        checkForNewAudioFiles();
        const pendingAnalysis = findNextPendingAnalysisJob_();
        if (pendingAnalysis) scheduleAnalysisResume_(1);
        result = { status: 'success', message: 'Check gestart' };
        break;

      case 'createReport':
        result = createNewReport_(body.data);
        break;

      case 'generateSection':
        result = generateSection(body.reportId, body.sectionId, body.overrides || {});
        break;

      case 'regenerateSection':
        result = regenerateSection(body.reportId, body.sectionId, body.extraContext || '');
        break;

      case 'generateAllSections':
        result = generateAllSections(body.reportId);
        break;

      case 'updateSection':
        result = updateSectionContent_(body.reportId, body.sectionId, body.content);
        break;

      case 'updateHeader':
        result = updateHeaderFields_(body.reportId, body.fields);
        break;

      case 'assembleReport':
        result = assembleReport(body.reportId, body.headerFields || {});
        break;

      case 'insertImage':
        insertImageInReport(body.reportId, body.imageFileId, body.position, body.index);
        result = { status: 'success' };
        break;

      case 'insertBijlageImages':
        insertBijlageImages(body.reportId, body.imageFileIds || []);
        result = { status: 'success', count: (body.imageFileIds || []).length };
        break;

      case 'deleteReport':
        result = deleteReport_(body.reportId);
        break;

      default:
        result = { error: 'Onbekende actie: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/****************************************************
 * WEB APP HELPER FUNCTIES
 ****************************************************/

function createNewReport_(data) {
  data = data || {};
  const reportId = Utilities.getUuid();

  const reportKey = "report_" + reportId;
  PROPS.setProperty(reportKey + "_meetingName", data.meetingName || "Nieuw Verslag");
  PROPS.setProperty(reportKey + "_meetingFolderId", data.meetingFolderId || reportId);
  PROPS.setProperty(reportKey + "_createdAt", new Date().toISOString());
  PROPS.setProperty(reportKey + "_template", data.template || "rabobank_mra");
  PROPS.setProperty(reportKey + "_status", "created");

  // Sla transcriptie file ID op als die is meegegeven
  if (data.transcriptTxtFileId) {
    PROPS.setProperty(reportKey + "_transcriptTxtFileId", data.transcriptTxtFileId);
  }

  // Sla header velden op
  if (data.headerFields) {
    PROPS.setProperty(reportKey + "_headerFields", JSON.stringify(data.headerFields));
  }

  console.log("✓ Nieuw report aangemaakt:", reportId);
  return { reportId, status: 'created' };
}

function deleteReport_(reportId) {
  if (!reportId) throw new Error("reportId is vereist");

  const all = PROPS.getProperties();
  const prefixes = [
    "report_" + reportId,
    "job_" + reportId,
    "secgen_" + reportId,
    "analysis_" + reportId
  ];

  let deletedCount = 0;
  for (const key of Object.keys(all)) {
    if (prefixes.some(p => key.startsWith(p + "_"))) {
      PROPS.deleteProperty(key);
      deletedCount++;
    }
  }

  // Verwijder ook section data
  for (const sec of MRA_SECTIONS) {
    const sectionKey = `report_${reportId}_section_${sec.id}`;
    if (PROPS.getProperty(sectionKey)) {
      PROPS.deleteProperty(sectionKey);
      deletedCount++;
    }
  }

  console.log(`✓ Report ${reportId} verwijderd (${deletedCount} properties)`);
  return { status: 'deleted', reportId, deletedProperties: deletedCount };
}

function listAllReports_() {
  const all = PROPS.getProperties();
  const reports = [];

  // Zoek alle report keys
  const reportIds = new Set();
  for (const key of Object.keys(all)) {
    const match = key.match(/^report_([^_]+)_status$/);
    if (match) reportIds.add(match[1]);

    // Ook transcriptie jobs als reports tonen
    const jobMatch = key.match(/^job_([^_]+)_state$/);
    if (jobMatch) reportIds.add(jobMatch[1]);
  }

  for (const id of reportIds) {
    reports.push(getReportSummary_(id));
  }

  reports.sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));
  return { reports };
}

function getReportSummary_(reportId) {
  const prefix = "report_" + reportId;
  const jobPrefix = "job_" + reportId;

  return {
    reportId,
    meetingName: PROPS.getProperty(prefix + "_meetingName") || PROPS.getProperty(jobPrefix + "_meetingName") || "Onbekend",
    status: PROPS.getProperty(prefix + "_status") || PROPS.getProperty(jobPrefix + "_state") || "unknown",
    template: PROPS.getProperty(prefix + "_template") || "rabobank_mra",
    createdAt: PROPS.getProperty(prefix + "_createdAt") || PROPS.getProperty(jobPrefix + "_createdAt") || "",
    docId: PROPS.getProperty(prefix + "_docId") || ""
  };
}

function getReportDetails_(reportId) {
  const summary = getReportSummary_(reportId);
  const sections = getAllSections_(reportId);
  const headerFields = PROPS.getProperty("report_" + reportId + "_headerFields");

  return {
    ...summary,
    headerFields: headerFields ? JSON.parse(headerFields) : {},
    sections: sections.sections,
    template: MRA_SECTIONS
  };
}

function getAllSections_(reportId) {
  const sections = {};
  for (const section of MRA_SECTIONS) {
    const key = `report_${reportId}_section_${section.id}`;
    const data = PROPS.getProperty(key);
    sections[section.id] = {
      ...section,
      generated: !!data,
      content: data ? JSON.parse(data) : null
    };
  }
  return { sections };
}

function getSectionContent_(reportId, sectionId) {
  const key = `report_${reportId}_section_${sectionId}`;
  const data = PROPS.getProperty(key);
  if (!data) return { sectionId, generated: false, content: null };
  return { sectionId, generated: true, content: JSON.parse(data) };
}

function updateSectionContent_(reportId, sectionId, content) {
  const key = `report_${reportId}_section_${sectionId}`;
  const existing = PROPS.getProperty(key);
  const data = existing ? JSON.parse(existing) : { sectionId };

  data.content = content;
  data.editedAt = new Date().toISOString();
  data.editedManually = true;

  PROPS.setProperty(key, JSON.stringify(data));
  return { status: 'updated', sectionId };
}

function updateHeaderFields_(reportId, fields) {
  PROPS.setProperty("report_" + reportId + "_headerFields", JSON.stringify(fields));
  return { status: 'updated' };
}

function getTranscriptStatus_(reportId) {
  const jobKey = "job_" + reportId;
  const state = PROPS.getProperty(jobKey + "_state");

  if (!state) return { status: 'not_found' };

  const fragmentsJson = PROPS.getProperty(jobKey + "_fragmentsJson") || "[]";
  const fragments = JSON.parse(fragmentsJson);
  const nextIndex = parseInt(PROPS.getProperty(jobKey + "_nextIndex") || "0", 10);

  return {
    status: state,
    progress: fragments.length > 0 ? Math.round((nextIndex / fragments.length) * 100) : 0,
    currentFragment: nextIndex,
    totalFragments: fragments.length,
    meetingName: PROPS.getProperty(jobKey + "_meetingName") || ""
  };
}

function loadReportData_(reportId) {
  // Probeer eerst report_ prefix, dan job_ prefix
  const reportPrefix = "report_" + reportId;
  const jobPrefix = "job_" + reportId;

  return {
    meetingName: PROPS.getProperty(reportPrefix + "_meetingName") || PROPS.getProperty(jobPrefix + "_meetingName") || "",
    meetingFolderId: PROPS.getProperty(reportPrefix + "_meetingFolderId") || PROPS.getProperty(jobPrefix + "_meetingFolderId") || reportId,
    transcriptTxtFileId: PROPS.getProperty(reportPrefix + "_transcriptTxtFileId") ||
      PROPS.getProperty("analysis_" + reportId + "_transcriptTxtFileId") ||
      PROPS.getProperty("secgen_" + reportId + "_transcriptTxtFileId") ||
      PROPS.getProperty(jobPrefix + "_transcriptTxtFileId") || "",
    docId: PROPS.getProperty(reportPrefix + "_docId") || ""
  };
}

/****************************************************
 * ═══════════════════════════════════════════════════
 * ONDERSTEUNENDE FUNCTIES (uit v10, ongewijzigd)
 * ═══════════════════════════════════════════════════
 ****************************************************/

function findAudioFileInFolder(folder) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();
    if (mimeType && mimeType.startsWith("audio/")) return file;
    const fileName = (file.getName() || "").toLowerCase();
    const audioExtensions = ['.mp3', '.wav', '.m4a', '.aac', '.ogg', '.flac', '.wma', '.aiff'];
    if (audioExtensions.some(ext => fileName.endsWith(ext))) return file;
  }
  return null;
}

function moveFileToArchive_(file, fromFolder) {
  const archiveFolder = DriveApp.getFolderById(ARCHIVE_FOLDER_ID);
  archiveFolder.addFile(file);
  try { fromFolder.removeFile(file); } catch (e) { console.log("⚠️ removeFile:", e); }
  console.log("Origineel verplaatst naar archief:", file.getName());
  return file.getId();
}

// CloudConvert functies
function getDriveMediaUrl_(fileId) { return `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`; }

function cloudConvertImportTaskForDriveFile_(file) {
  return {
    operation: "import/url",
    url: getDriveMediaUrl_(file.getId()),
    filename: file.getName(),
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  };
}

function getMetadataViaCloudConvert(file) {
  const job = cloudConvertCreateJob_({ tasks: { "import": cloudConvertImportTaskForDriveFile_(file), "meta": { operation: "metadata", input: ["import"] } } });
  const task = waitForCloudConvertTask_(job.data.id, "meta");
  return task.result && task.result.metadata;
}

function convertDirect(file) {
  const folderName = stripExtension_(file.getName());
  const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  const subFolder = targetFolder.createFolder(folderName);
  const job = cloudConvertCreateJob_({ tasks: {
    "import": cloudConvertImportTaskForDriveFile_(file),
    "convert": { operation: "convert", input: ["import"], output_format: "mp3", engine: "ffmpeg", audio_codec: "mp3", audio_bitrate: TARGET_BITRATE_KBPS },
    "export": { operation: "export/url", input: ["convert"] }
  }});
  const convertTask = waitForCloudConvertTask_(job.data.id, "convert");
  if (convertTask.status === "error") throw new Error("Conversie mislukt: " + JSON.stringify(convertTask));
  const exportTask = waitForCloudConvertTask_(job.data.id, "export");
  const url = exportTask.result.files[0].url;
  return { file: saveUrlToDrive_(url, stripExtension_(file.getName()) + ".mp3", subFolder), subFolder };
}

function splitAudioUltraRobust(file, duration, maxFragmentDuration) {
  const parts = Math.ceil(duration / maxFragmentDuration);
  const folderName = stripExtension_(file.getName());
  const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  const subFolder = targetFolder.createFolder(folderName);
  const fragments = [];
  for (let i = 0; i < parts; i++) {
    const start = i * maxFragmentDuration;
    const end = Math.min((i + 1) * maxFragmentDuration, duration);
    const fragName = `${stripExtension_(file.getName())}_part_${i + 1}.mp3`;
    let savedFile;
    try { savedFile = splitOneFragment_(file, start, end, fragName, subFolder); }
    catch (err) { savedFile = splitOneFragment_(file, toTimecode_(start), toTimecode_(end), fragName, subFolder); }
    fragments.push({ file: savedFile, startSec: start, endSec: end, index: i + 1 });
  }
  return { fragments, subFolder };
}

function splitOneFragment_(file, start, end, newName, subFolder) {
  const job = cloudConvertCreateJob_({ tasks: {
    "import": cloudConvertImportTaskForDriveFile_(file),
    "trim": { operation: "convert", input: ["import"], output_format: "mp3", engine: "ffmpeg", audio_codec: "mp3", audio_bitrate: TARGET_BITRATE_KBPS, trim_start: `${start}`, trim_end: `${end}` },
    "export": { operation: "export/url", input: ["trim"] }
  }});
  const trimTask = waitForCloudConvertTask_(job.data.id, "trim");
  if (trimTask.status === "error") throw new Error("Trim mislukt: " + JSON.stringify(trimTask));
  const exportTask = waitForCloudConvertTask_(job.data.id, "export");
  return saveUrlToDrive_(exportTask.result.files[0].url, newName, subFolder);
}

// Whisper
function transcribeWithWhisper_(file) {
  const blob = file.getBlob();
  const fileName = file.getName();
  const boundary = "----WebKitFormBoundary" + Utilities.getUuid().replace(/-/g, "");
  const requestBody = buildMultipartBody_(boundary, blob, fileName);
  const response = UrlFetchApp.fetch(WHISPER_API_URL, {
    method: "post",
    headers: { "Authorization": "Bearer " + OPENAI_API_KEY, "Content-Type": "multipart/form-data; boundary=" + boundary },
    payload: requestBody, muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) throw new Error("Whisper mislukt: " + response.getContentText());
  return JSON.parse(response.getContentText()).text || "";
}

function buildMultipartBody_(boundary, blob, fileName) {
  const fileData = blob.getBytes();
  let body = "--" + boundary + "\r\n" + 'Content-Disposition: form-data; name="file"; filename="' + fileName + '"\r\n' + "Content-Type: audio/mpeg\r\n\r\n";
  const postFile = "\r\n--" + boundary + "\r\n" + 'Content-Disposition: form-data; name="model"\r\n\r\nwhisper-1' +
    "\r\n--" + boundary + "\r\n" + 'Content-Disposition: form-data; name="language"\r\n\r\n' + TRANSCRIBE_LANGUAGE +
    "\r\n--" + boundary + "\r\n" + 'Content-Disposition: form-data; name="response_format"\r\n\r\njson' +
    "\r\n--" + boundary + "--\r\n";
  return Utilities.newBlob([...Utilities.newBlob(body).getBytes(), ...fileData, ...Utilities.newBlob(postFile).getBytes()]).getBytes();
}

// Document helpers
function createOrGetPartsFolder_(transcriptFolder, meetingName) {
  const folderName = `${meetingName} - Deeltranscripties`;
  const existing = transcriptFolder.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return transcriptFolder.createFolder(folderName);
}

function createTranscriptDoc_(title, content, folder) {
  const doc = DocumentApp.create(title);
  doc.getBody().setText(content || "");
  doc.saveAndClose();
  const docFile = DriveApp.getFileById(doc.getId());
  folder.addFile(docFile);
  try { DriveApp.getRootFolder().removeFile(docFile); } catch(e) {}
  return doc;
}

function appendToDoc_(docId, text) {
  const doc = DocumentApp.openById(docId);
  doc.getBody().appendParagraph(text);
  doc.saveAndClose();
}

function createTranscriptTxt(title, content, folder) {
  const safeTitle = String(title).replace(/[\\/:*?"<>|#%]/g, "-");
  const blob = Utilities.newBlob(content, "text/plain", safeTitle + ".txt");
  return folder.createFile(blob);
}

function cleanupMP3FragmentsByIds_(fileIds, workFolderId) {
  let count = 0;
  for (const id of fileIds || []) {
    try { DriveApp.getFileById(id).setTrashed(true); count++; } catch (e) {}
  }
  if (workFolderId) {
    try {
      const wf = DriveApp.getFolderById(workFolderId);
      if (!wf.getFiles().hasNext() && !wf.getFolders().hasNext()) wf.setTrashed(true);
    } catch (e) {}
  }
  console.log(`${count} MP3 fragment(en) verwijderd`);
}

// CloudConvert helpers
function cloudConvertCreateJob_(payload) {
  const res = UrlFetchApp.fetch(`${CLOUDCONVERT_API}/jobs`, {
    method: "post", contentType: "application/json",
    headers: { Authorization: "Bearer " + CLOUDCONVERT_TOKEN },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) throw new Error("CloudConvert fout: " + res.getContentText());
  return JSON.parse(res.getContentText());
}

function cloudConvertGetJob_(jobId) {
  const res = UrlFetchApp.fetch(`${CLOUDCONVERT_API}/jobs/${jobId}`, {
    headers: { Authorization: "Bearer " + CLOUDCONVERT_TOKEN }, muteHttpExceptions: true
  });
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) throw new Error("CloudConvert fout: " + res.getContentText());
  return JSON.parse(res.getContentText());
}

function waitForCloudConvertTask_(jobId, taskName) {
  for (let i = 0; i < 120; i++) {
    const job = cloudConvertGetJob_(jobId);
    const task = job.data.tasks.find(t => t.name === taskName);
    if (!task) throw new Error("Task niet gevonden: " + taskName);
    if (task.status === "finished") return task;
    if (task.status === "error") throw new Error(JSON.stringify(task));
    Utilities.sleep(2000);
  }
  throw new Error("CloudConvert time-out: " + taskName);
}

function saveUrlToDrive_(url, name, subFolder) {
  return subFolder.createFile(UrlFetchApp.fetch(url).getBlob().setName(name));
}

// Utilities
function calculateMaxFragmentDuration() {
  const maxSizeBits = MAX_FILE_SIZE_MB * 1024 * 1024 * 8;
  const maxDurationForSize = Math.floor(maxSizeBits / (TARGET_BITRATE_KBPS * 1000));
  return Math.min(MAX_DURATION_SEC, maxDurationForSize);
}

function toTimecode_(seconds) {
  const h = Math.floor(seconds / 3600).toString().padStart(2, "0");
  const m = Math.floor((seconds % 3600) / 60).toString().padStart(2, "0");
  const s = Math.floor(seconds % 60).toString().padStart(2, "0");
  return `${h}:${m}:${s}`;
}

function stripExtension_(name) { return String(name).replace(/\.[^/.]+$/, ""); }

function parseDuration(value) {
  if (!value) throw new Error("Geen Duration ontvangen.");
  let str = String(value).trim().replace(/\s*\(.*?\)\s*/g, '').trim();
  if (str.includes(":")) {
    const parts = str.split(":").map(Number);
    if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + Math.floor(parts[2]);
    if (parts.length === 2) return parts[0] * 60 + Math.floor(parts[1]);
  }
  const secMatch = str.match(/^([\d.]+)\s*s$/i);
  if (secMatch) return Math.floor(parseFloat(secMatch[1]));
  const minMatch = str.match(/^([\d.]+)\s*min$/i);
  if (minMatch) return Math.floor(parseFloat(minMatch[1]) * 60);
  const num = parseFloat(str);
  if (!isNaN(num)) return Math.floor(num);
  throw new Error("Kan duur niet parsen: " + str);
}

function assertConfig_() {
  if (!CLOUDCONVERT_TOKEN) throw new Error("CLOUDCONVERT_TOKEN ontbreekt.");
  if (!OPENAI_API_KEY) throw new Error("OPENAI_API_KEY ontbreekt.");
  if (!ARCHIVE_FOLDER_ID || ARCHIVE_FOLDER_ID === "VUL_HIER_JE_ARCHIVE_FOLDER_ID_IN") throw new Error("ARCHIVE_FOLDER_ID ontbreekt.");
}

// OpenAI Chat helper (NIEUW v11 - voor sectie generatie)
function callOpenAIChat_(payload) {
  const options = {
    method: "post", contentType: "application/json",
    headers: { "Authorization": "Bearer " + OPENAI_API_KEY },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };
  for (let attempt = 1; attempt <= 5; attempt++) {
    const res = UrlFetchApp.fetch(OPENAI_CHAT_URL, options);
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) {
      const json = JSON.parse(res.getContentText());
      return json.choices[0].message.content;
    }
    if (code === 429 || code >= 500) { Utilities.sleep(800 * attempt); continue; }
    throw new Error("OpenAI call failed (HTTP " + code + "): " + res.getContentText());
  }
  throw new Error("OpenAI call failed na retries.");
}

function callOpenAIJson_(payload) {
  const content = callOpenAIChat_(payload);
  return content;
}

function safeParseJson_(text) {
  try { return JSON.parse(text); }
  catch (e) { return JSON.parse(String(text || "").trim().replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/```$/i, "").trim()); }
}

// Job store/load helpers
function jobKey_(meetingFolderId) { return "job_" + meetingFolderId; }

function findNextPendingJob_() {
  const all = PROPS.getProperties();
  const jobKeys = Object.keys(all).filter(k => k.endsWith("_state") && k.startsWith("job_")).map(k => k.replace(/_state$/, ""));
  for (const key of jobKeys) {
    const state = PROPS.getProperty(key + "_state");
    if (state === "ready" || state === "transcribing" || state === "preparing") return loadJob_(key);
  }
  return null;
}

function loadJob_(key) {
  return {
    key,
    meetingName: PROPS.getProperty(key + "_meetingName") || "",
    meetingFolderId: PROPS.getProperty(key + "_meetingFolderId") || "",
    state: PROPS.getProperty(key + "_state") || "",
    nextIndex: parseInt(PROPS.getProperty(key + "_nextIndex") || "0", 10),
    fragments: JSON.parse(PROPS.getProperty(key + "_fragmentsJson") || "[]"),
    workFolderId: PROPS.getProperty(key + "_workFolderId") || "",
    docId: PROPS.getProperty(key + "_docId") || ""
  };
}

// Trigger management
function scheduleResume_(min) {
  cleanupResumeTriggers_();
  ScriptApp.newTrigger("resumePendingWork").timeBased().after(Math.max(1, min) * 60 * 1000).create();
}

function cleanupResumeTriggers_() {
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction && t.getHandlerFunction() === "resumePendingWork") ScriptApp.deleteTrigger(t);
  }
}

function scheduleAnalysisResume_(min) {
  cleanupAnalysisTriggers_();
  ScriptApp.newTrigger("resumePendingAnalysis").timeBased().after(Math.max(1, min) * 60 * 1000).create();
}

function cleanupAnalysisTriggers_() {
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction && t.getHandlerFunction() === "resumePendingAnalysis") ScriptApp.deleteTrigger(t);
  }
}

/****************************************************
 * PLACEHOLDER ANALYSE (uit v10, met uitgebreide keys)
 ****************************************************/
function analysisKey_(meetingFolderId) { return "analysis_" + meetingFolderId; }

function queuePlaceholderAnalysis_(meetingFolderId, meetingName, transcriptTxtFileId) {
  const key = analysisKey_(meetingFolderId);
  PROPS.setProperty(key + "_state", "pending");
  PROPS.setProperty(key + "_meetingName", meetingName);
  PROPS.setProperty(key + "_meetingFolderId", meetingFolderId);
  PROPS.setProperty(key + "_transcriptTxtFileId", transcriptTxtFileId);
  PROPS.setProperty(key + "_offset", "0");
  PROPS.deleteProperty(key + "_draftDocId");
  console.log("✓ Placeholder analyse queued:", meetingName);
}

function findNextPendingAnalysisJob_() {
  const all = PROPS.getProperties();
  const keys = Object.keys(all).filter(k => k.startsWith("analysis_") && k.endsWith("_state")).map(k => k.replace(/_state$/, ""));
  for (const key of keys) {
    const state = PROPS.getProperty(key + "_state");
    if (state === "pending" || state === "chunking" || state === "finalizing") return loadAnalysisJob_(key);
  }
  return null;
}

function loadAnalysisJob_(key) {
  return {
    key,
    state: PROPS.getProperty(key + "_state") || "",
    meetingName: PROPS.getProperty(key + "_meetingName") || "",
    meetingFolderId: PROPS.getProperty(key + "_meetingFolderId") || "",
    transcriptTxtFileId: PROPS.getProperty(key + "_transcriptTxtFileId") || "",
    offset: parseInt(PROPS.getProperty(key + "_offset") || "0", 10),
    draftDocId: PROPS.getProperty(key + "_draftDocId") || ""
  };
}

function resumePendingAnalysis() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  const started = Date.now();
  try {
    assertConfig_();
    const job = findNextPendingAnalysisJob_();
    if (!job) { cleanupAnalysisTriggers_(); return; }

    console.log("=== RESUME ANALYSIS ===", job.meetingName);

    const transcriptText = DriveApp.getFileById(job.transcriptTxtFileId).getBlob().getDataAsString("UTF-8");
    const transcriptFolder = DriveApp.getFolderById(TRANSCRIPT_FOLDER_ID);

    let draftDocId = job.draftDocId;
    if (!draftDocId) {
      const draftDoc = createTranscriptDoc_(job.meetingName + " - PlaceholderDraft (JSON)",
        JSON.stringify({ fields: initEmptyFields_(), qa: {} }, null, 2), transcriptFolder);
      draftDocId = draftDoc.getId();
      PROPS.setProperty(job.key + "_draftDocId", draftDocId);
    }

    PROPS.setProperty(job.key + "_state", "chunking");
    let offset = job.offset;
    let chunksDone = 0;

    while (offset < transcriptText.length && chunksDone < ANALYSIS_CHUNKS_PER_RUN) {
      if (Date.now() - started > (RUNTIME_BUDGET_MS - 90 * 1000)) {
        PROPS.setProperty(job.key + "_offset", String(offset));
        scheduleAnalysisResume_(1);
        return;
      }
      const chunk = transcriptText.substring(offset, Math.min(offset + ANALYSIS_CHUNK_CHARS, transcriptText.length));
      const partial = extractPlaceholdersFromChunk_(job.meetingName, chunk);
      mergePartialIntoDraft_(draftDocId, partial);
      offset += chunk.length;
      chunksDone++;
      PROPS.setProperty(job.key + "_offset", String(offset));
      Utilities.sleep(800);
    }

    if (offset < transcriptText.length) { scheduleAnalysisResume_(1); return; }

    PROPS.setProperty(job.key + "_state", "finalizing");
    const draftJson = readJsonFromDoc_(draftDocId);
    const finalized = finalizePlaceholderFields_(job.meetingName, (draftJson && draftJson.fields) || {});
    writePlaceholdersToSheet_(job.meetingName, finalized.fields || {}, finalized.qa || {}, draftDocId);
    PROPS.setProperty(job.key + "_state", "done");
    console.log("✓ Analysis afgerond:", job.meetingName);

    const next = findNextPendingAnalysisJob_();
    if (next) scheduleAnalysisResume_(1);
    else cleanupAnalysisTriggers_();

  } catch (err) {
    console.log("!!! FOUT analysis:", err);
    const job = findNextPendingAnalysisJob_();
    if (job) { PROPS.setProperty(job.key + "_state", "error"); PROPS.setProperty(job.key + "_error", String(err)); }
    scheduleAnalysisResume_(5);
  } finally {
    lock.releaseLock();
  }
}

function initEmptyFields_() {
  const obj = { "Verslag": "" };
  for (const k of PLACEHOLDER_KEYS) obj[k] = "";
  return obj;
}

function extractPlaceholdersFromChunk_(meetingName, chunkText) {
  const system = [
    "Je extraheert tekst uit een transcriptie en zet dit in Rabobank MRA placeholders.",
    "NIEUW: Nulmeting is opgesplitst in: kernwaarden, sterke_eigenschappen, aandachtspunten, quickscan.",
    "Regels: Verzin niets. Alleen tekst uit dit chunk. Output STRICT JSON: {\"fields\":{...}}.",
    "Per veld: korte zinnen/fragmenten. Als geen info: lege string."
  ].join("\n");

  const user = `PLACEHOLDER KEYS:\n${JSON.stringify(PLACEHOLDER_KEYS)}\n\nMEETING:\n${meetingName}\n\nTRANSCRIPT CHUNK:\n${chunkText}`;

  const payload = {
    model: ANALYSIS_MODEL, temperature: 0.1,
    response_format: { type: "json_object" },
    messages: [{ role: "system", content: system }, { role: "user", content: user }]
  };

  const content = callOpenAIJson_(payload);
  const parsed = safeParseJson_(content);
  const fields = initEmptyFields_();
  const src = (parsed && parsed.fields) || {};
  for (const k of PLACEHOLDER_KEYS) fields[k] = (src[k] || "").toString().trim();
  return { fields };
}

function finalizePlaceholderFields_(meetingName, aggregatedFields) {
  const system = [
    "Je maakt definitieve Rabobank MRA placeholders uit geaggregeerde chunk-notities.",
    "Verwijder duplicaten, maak leesbare alinea's. Output STRICT JSON: {\"fields\":{...},\"qa\":{...}}.",
    "qa.missing_fields: lege placeholders. qa.coverage_check: 1 zin over deduplicatie."
  ].join("\n");

  const user = `MEETING:\n${meetingName}\n\nAGGREGATED:\n${JSON.stringify({ fields: aggregatedFields }, null, 2)}\n\nKEYS:\n${JSON.stringify(PLACEHOLDER_KEYS)}`;

  const payload = {
    model: ANALYSIS_MODEL, temperature: 0.2,
    response_format: { type: "json_object" },
    messages: [{ role: "system", content: system }, { role: "user", content: user }]
  };

  const parsed = safeParseJson_(callOpenAIJson_(payload)) || {};
  const fields = initEmptyFields_();
  fields["Verslag"] = meetingName;
  const src = parsed.fields || {};
  for (const k of PLACEHOLDER_KEYS) fields[k] = (src[k] || "").toString().trim();
  const qa = parsed.qa || {};
  if (!qa.missing_fields) qa.missing_fields = PLACEHOLDER_KEYS.filter(k => !fields[k]);
  return { fields, qa };
}

function readJsonFromDoc_(docId) {
  return safeParseJson_(DocumentApp.openById(docId).getBody().getText());
}

function mergePartialIntoDraft_(draftDocId, partial) {
  const current = readJsonFromDoc_(draftDocId) || {};
  const baseFields = current.fields || initEmptyFields_();
  const addFields = (partial && partial.fields) || {};
  for (const k of PLACEHOLDER_KEYS) {
    const add = (addFields[k] || "").toString().trim();
    if (!add) continue;
    const cur = (baseFields[k] || "").toString().trim();
    baseFields[k] = cur ? (cur + "\n" + add) : add;
  }
  const doc = DocumentApp.openById(draftDocId);
  doc.getBody().setText(JSON.stringify({ fields: baseFields, qa: current.qa || {} }, null, 2));
  doc.saveAndClose();
}

function writePlaceholdersToSheet_(meetingName, fields, qa, draftDocId) {
  const ss = SpreadsheetApp.openById(PLACEHOLDER_SPREADSHEET_ID);
  const dataSheet = getSheetByGidOrName_(ss, PLACEHOLDER_DATA_SHEET_GID, PLACEHOLDER_SHEET_NAME, true);
  ensurePlaceholderHeaders_(dataSheet);
  const lastCol = dataSheet.getLastColumn();
  const headers = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);
  const row = new Array(headers.length).fill("");
  if (colIndex["Verslag"] !== undefined) row[colIndex["Verslag"]] = meetingName;
  for (const k of PLACEHOLDER_KEYS) { if (colIndex[k] !== undefined) row[colIndex[k]] = (fields[k] || "").toString(); }
  dataSheet.appendRow(row);

  const qaSheet = getSheetByGidOrName_(ss, null, PLACEHOLDER_QA_SHEET_NAME, true);
  ensureQaHeaders_(qaSheet);
  qaSheet.appendRow([new Date(), meetingName, ANALYSIS_MODEL, (qa.missing_fields || []).join(", "), (qa.uncertain_fields || []).join(", "), qa.coverage_check || "", qa.unassigned_notes || "", "https://docs.google.com/document/d/" + draftDocId + "/edit"]);
  console.log("✓ Placeholders → sheet:", meetingName);
}

function getSheetByGidOrName_(ss, gidOrNull, name, createIfMissing) {
  if (gidOrNull !== null && gidOrNull !== undefined && !isNaN(gidOrNull)) {
    for (const sh of ss.getSheets()) { if (sh.getSheetId && sh.getSheetId() === gidOrNull) return sh; }
  }
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  if (createIfMissing) return ss.insertSheet(name);
  return null;
}

function ensurePlaceholderHeaders_(sheet) {
  const desired = ["Verslag"].concat(PLACEHOLDER_KEYS);
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) { sheet.getRange(1, 1, 1, desired.length).setValues([desired]); sheet.setFrozenRows(1); return; }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x).trim());
  const existing = new Set(headers);
  let col = lastCol;
  for (const h of desired) { if (!existing.has(h)) { col++; sheet.getRange(1, col).setValue(h); existing.add(h); } }
  sheet.setFrozenRows(1);
}

function ensureQaHeaders_(sheet) {
  const headers = ["Timestamp", "Verslag", "Model", "MissingFields", "UncertainFields", "CoverageCheck", "UnassignedNotes", "DraftDocUrl"];
  if (sheet.getLastColumn() === 0) { sheet.getRange(1, 1, 1, headers.length).setValues([headers]); sheet.setFrozenRows(1); }
}
