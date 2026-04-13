# Graph Report - .  (2026-04-13)

## Corpus Check
- Corpus is ~23,885 words - fits in a single context window. You may not need a graph.

## Summary
- 95 nodes · 150 edges · 12 communities detected
- Extraction: 91% EXTRACTED · 9% INFERRED · 0% AMBIGUOUS · INFERRED: 13 edges (avg confidence: 0.83)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- [[_COMMUNITY_assembleAndPreview()|assembleAndPreview()]]
- [[_COMMUNITY_checkApiConfig()|checkApiConfig()]]
- [[_COMMUNITY_Advies|Advies]]
- [[_COMMUNITY_Big Data Section|Big Data Section]]
- [[_COMMUNITY_CloudConvert API|CloudConvert API]]
- [[_COMMUNITY_preview.js|preview.js]]
- [[_COMMUNITY_Rabobank|Rabobank]]
- [[_COMMUNITY_Audio Transcription Pipeline|Audio Transcription Pipeline]]
- [[_COMMUNITY_LiBeR (Advies Begeleiding Concepten)|LiBeR (Advies Begeleiding Concepten)]]
- [[_COMMUNITY_Lutger Brenninkmeijer (Author)|Lutger Brenninkmeijer (Author)]]
- [[_COMMUNITY_api.js|api.js]]
- [[_COMMUNITY_download_logos.py|download_logos.py]]

## God Nodes (most connected - your core abstractions)
1. `KNVB Intake Verslag Template` - 12 edges
2. `LiBeR Verslaggenerator` - 11 edges
3. `LiBeR` - 10 edges
4. `renderEditorDocument()` - 7 edges
5. `getEditorSections()` - 6 edges
6. `renderSectionList()` - 6 edges
7. `generateSingleSection()` - 6 edges
8. `LiBeR Logo 2008 Diapositief (White on Orange)` - 6 edges
9. `openEditor()` - 5 edges
10. `checkAndUpdateTranscriptStatus()` - 5 edges

## Surprising Connections (you probably didn't know these)
- `KNVB Intake Verslag Template` --conceptually_related_to--> `MRA Template Structure`  [INFERRED]
  KNVB - TEMPLATE verslag intake.pdf → README.md
- `LiBeR Verslaggenerator` --conceptually_related_to--> `KNVB Intake Verslag Template`  [INFERRED]
  README.md → KNVB - TEMPLATE verslag intake.pdf
- `KNVB` --conceptually_related_to--> `Rabobank`  [INFERRED]
  KNVB - TEMPLATE verslag intake.pdf → README.md
- `LiBeR (Advies Begeleiding Concepten)` --conceptually_related_to--> `LiBeR (Advies Begeleiding Concepten)`  [INFERRED]
  KNVB - TEMPLATE verslag intake.pdf → README.md
- `Lutger Brenninkmeijer (Author)` --conceptually_related_to--> `Lutger Brenninkmeijer`  [INFERRED]
  KNVB - TEMPLATE verslag intake.pdf → README.md

## Communities

### Community 0 - "assembleAndPreview()"
Cohesion: 0.2
Nodes (19): assembleAndPreview(), assembleReport(), checkAndUpdateTranscriptStatus(), generateAllSections(), generateOfflineContent(), generateSingleSection(), getEditorSections(), openEditor() (+11 more)

### Community 1 - "checkApiConfig()"
Cohesion: 0.15
Nodes (15): createReport(), deleteReport(), escapeHtml(), handleAudioFile(), loadDashboard(), loadJotformPage(), loadJotformSubmissions(), navigateTo() (+7 more)

### Community 2 - "Advies"
Cohesion: 0.25
Nodes (15): Advies, Begeleiding, Brand Color Orange (#E87722), Brand Color White (#FFFFFF), Concepten, LiBeR, Logo Design Era 2008, Nunito (Font) (+7 more)

### Community 3 - "Big Data Section"
Cohesion: 0.22
Nodes (9): Big Data Section, KNVB Intake Verslag Template, KNVB, Observatie en Vrijblijvend Advies, Onderstroom, Ontwikkelbehoefte (Development Needs), SWOT Analysis, Taxameter (+1 more)

### Community 4 - "CloudConvert API"
Cohesion: 0.22
Nodes (9): CloudConvert API, Google Apps Script, Google Docs, Google Drive, Google Sheets, GPT-4o, LiBeR Verslaggenerator, MRA Template Structure (+1 more)

### Community 5 - "preview.js"
Cohesion: 0.7
Nodes (4): buildPreviewPage(), esc(), renderPreview(), textToHtml()

### Community 6 - "Rabobank"
Cohesion: 0.67
Nodes (3): Rabobank, Rabobank Kring Metropool Regio Amsterdam, Rabobank MRA (Maatschappelijke Relevantie Analyse)

### Community 7 - "Audio Transcription Pipeline"
Cohesion: 1.0
Nodes (2): Audio Transcription Pipeline, OpenAI Whisper

### Community 8 - "LiBeR (Advies Begeleiding Concepten)"
Cohesion: 1.0
Nodes (2): LiBeR (Advies Begeleiding Concepten), LiBeR (Advies Begeleiding Concepten)

### Community 9 - "Lutger Brenninkmeijer (Author)"
Cohesion: 1.0
Nodes (2): Lutger Brenninkmeijer (Author), Lutger Brenninkmeijer

### Community 10 - "api.js"
Cohesion: 1.0
Nodes (0): 

### Community 11 - "download_logos.py"
Cohesion: 1.0
Nodes (0): 

## Knowledge Gaps
- **19 isolated node(s):** `Rabobank Kring Metropool Regio Amsterdam`, `OpenAI Whisper`, `GPT-4o`, `CloudConvert API`, `Google Apps Script` (+14 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **Thin community `Audio Transcription Pipeline`** (2 nodes): `Audio Transcription Pipeline`, `OpenAI Whisper`
  Too small to be a meaningful cluster - may be noise or needs more connections extracted.
- **Thin community `LiBeR (Advies Begeleiding Concepten)`** (2 nodes): `LiBeR (Advies Begeleiding Concepten)`, `LiBeR (Advies Begeleiding Concepten)`
  Too small to be a meaningful cluster - may be noise or needs more connections extracted.
- **Thin community `Lutger Brenninkmeijer (Author)`** (2 nodes): `Lutger Brenninkmeijer (Author)`, `Lutger Brenninkmeijer`
  Too small to be a meaningful cluster - may be noise or needs more connections extracted.
- **Thin community `api.js`** (1 nodes): `api.js`
  Too small to be a meaningful cluster - may be noise or needs more connections extracted.
- **Thin community `download_logos.py`** (1 nodes): `download_logos.py`
  Too small to be a meaningful cluster - may be noise or needs more connections extracted.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `LiBeR Verslaggenerator` connect `CloudConvert API` to `LiBeR (Advies Begeleiding Concepten)`, `Big Data Section`, `Rabobank`, `Audio Transcription Pipeline`?**
  _High betweenness centrality (0.051) - this node is a cross-community bridge._
- **Why does `KNVB Intake Verslag Template` connect `Big Data Section` to `LiBeR (Advies Begeleiding Concepten)`, `Lutger Brenninkmeijer (Author)`, `CloudConvert API`?**
  _High betweenness centrality (0.048) - this node is a cross-community bridge._
- **Are the 2 inferred relationships involving `KNVB Intake Verslag Template` (e.g. with `MRA Template Structure` and `LiBeR Verslaggenerator`) actually correct?**
  _`KNVB Intake Verslag Template` has 2 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Rabobank Kring Metropool Regio Amsterdam`, `OpenAI Whisper`, `GPT-4o` to the rest of the system?**
  _19 weakly-connected nodes found - possible documentation gaps or missing edges._