/**
 * LiBeR Verslaggenerator – Editor Module
 * Beheert de per-sectie editor, generatie en sectie-opslag.
 */

// ---- EDITOR STATE ----
const EDITOR = {
  reportId: null,
  report: null,
  sections: {},     // { sectionId: { content: '...', generated: false, editedManually: false } }
  activeSection: null,
  headerFields: {}
};

// ---- OPEN EDITOR ----
function openEditor(reportId) {
  EDITOR.reportId = reportId;
  APP.currentReportId = reportId;

  // Zoek report in lokale data
  EDITOR.report = APP.reports.find(r => r.reportId === reportId);
  if (!EDITOR.report) {
    showToast('Verslag niet gevonden.', 'error');
    navigateTo('dashboard');
    return;
  }

  EDITOR.headerFields = EDITOR.report.headerFields || {};
  EDITOR.sections = EDITOR.report.sections || {};

  // Update title
  const title = EDITOR.headerFields['<<naam vereniging>>'] || EDITOR.report.meetingName || 'Nieuw Verslag';
  document.getElementById('editor-title').textContent = `Rabobank MRA – verslag intake ${title}`;

  // Render sidebar
  renderSectionList();

  // Render document
  renderEditorDocument();

  // Probeer secties van server te laden
  if (API.isConfigured() && EDITOR.report.serverReportId) {
    API.getSections(EDITOR.report.serverReportId).then(data => {
      if (data.sections) {
        for (const [id, sec] of Object.entries(data.sections)) {
          if (sec.content && sec.content.content) {
            EDITOR.sections[id] = {
              content: sec.content.content,
              generated: true,
              generatedAt: sec.content.generatedAt
            };
          }
        }
        EDITOR.report.sections = EDITOR.sections;
        saveLocalReports();
        renderSectionList();
        renderEditorDocument();
      }
    }).catch(err => console.warn('Kon secties niet laden:', err));
  }
}

// ---- SECTION LIST (SIDEBAR) ----
function renderSectionList() {
  const list = document.getElementById('section-list');
  let generatedCount = 0;

  list.innerHTML = MRA_SECTIONS.map(sec => {
    const data = EDITOR.sections[sec.id];
    let dotClass = 'empty';
    if (data && data.content) {
      dotClass = data.editedManually ? 'editing' : 'generated';
      generatedCount++;
    }
    const activeClass = EDITOR.activeSection === sec.id ? 'active' : '';

    return `<li class="section-item ${activeClass}" onclick="selectSection('${sec.id}')">
      <span class="section-dot ${dotClass}"></span>
      <span>${sec.title}</span>
    </li>`;
  }).join('');

  document.getElementById('section-progress').textContent = `${generatedCount}/${MRA_SECTIONS.length}`;
}

// ---- SELECT SECTION ----
function selectSection(sectionId) {
  EDITOR.activeSection = sectionId;
  renderSectionList();
  renderEditorDocument();

  // Scroll to section
  const el = document.getElementById('block-' + sectionId);
  if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

// ---- RENDER EDITOR DOCUMENT ----
function renderEditorDocument() {
  const area = document.getElementById('section-content-area');
  const hf = EDITOR.headerFields;

  let html = '';

  // Header table
  html += `<div class="section-block" id="block-header">
    <table class="header-fields-table">
      <tr><td>Opdrachtgever:</td><td><input type="text" value="${escapeHtml(hf['<<opdrachtgever>>'] || 'Rabobank Kring Metropool Regio Amsterdam')}" data-field="<<opdrachtgever>>"></td></tr>
      <tr><td>Vereniging:</td><td><input type="text" value="${escapeHtml(hf['<<naam vereniging>>'] || '')}" data-field="<<naam vereniging>>"></td></tr>
      <tr><td>Onderwerp:</td><td><input type="text" value="${escapeHtml(hf['<<onderwerp>>'] || '')}" data-field="<<onderwerp>>"></td></tr>
      <tr><td>Contact met:</td><td><input type="text" value="${escapeHtml(hf['<<contactpersoon>>'] || '')}" data-field="<<contactpersoon>>"></td></tr>
      <tr><td>Opgesteld door:</td><td><input type="text" value="${escapeHtml(hf['<<opgesteld_door>>'] || 'Lutger Brenninkmeijer')}" data-field="<<opgesteld_door>>"></td></tr>
    </table>
  </div>`;

  // Sections
  for (const sec of MRA_SECTIONS) {
    if (sec.id === 'header') continue;

    const data = EDITOR.sections[sec.id] || {};
    const content = data.content || '';
    const isActive = EDITOR.activeSection === sec.id;

    html += `<div class="section-block ${isActive ? 'active' : ''}" id="block-${sec.id}">
      <div class="section-block-header">
        <span class="section-block-title">${sec.title}</span>
        <div class="section-block-actions">`;

    if (sec.generatable) {
      html += `<button class="btn btn-primary btn-sm" onclick="generateSingleSection('${sec.id}')" title="Genereer vanuit transcriptie">
            <span class="material-symbols-outlined" style="font-size:14px">auto_awesome</span> Genereer
          </button>`;
      if (content) {
        html += `<button class="btn btn-ghost btn-sm" onclick="regenerateSingleSection('${sec.id}')" title="Opnieuw genereren">
            <span class="material-symbols-outlined" style="font-size:14px">refresh</span>
          </button>`;
      }
    }
    if (sec.allowImages) {
      html += `<button class="btn btn-ghost btn-sm" onclick="uploadImageForSection('${sec.id}')">
          <span class="material-symbols-outlined" style="font-size:14px">add_photo_alternate</span> Afbeelding
        </button>`;
    }

    html += `</div></div>`;

    if (sec.allowImages && sec.id === 'nulmeting_images') {
      // Twee-kolom layout voor 0-meting / Stip
      html += `<div class="preview-two-col">
        <div>
          <h4 style="color:var(--orange);font-family:var(--font-serif);font-style:italic;font-size:13px">0-meting</h4>
          <div class="preview-image-placeholder" id="img-nulmeting-links">
            ${data.imageLeft ? `<img src="${data.imageLeft}" style="width:100%;height:100%;object-fit:contain">` : 'Klik om afbeelding toe te voegen'}
          </div>
        </div>
        <div>
          <h4 style="color:var(--orange);font-family:var(--font-serif);font-style:italic;font-size:13px">Stip op de horizon</h4>
          <div class="preview-image-placeholder" id="img-nulmeting-rechts">
            ${data.imageRight ? `<img src="${data.imageRight}" style="width:100%;height:100%;object-fit:contain">` : 'Klik om afbeelding toe te voegen'}
          </div>
        </div>
      </div>`;
    } else if (sec.allowImages && sec.id === 'bijlage') {
      // Bijlage afbeeldingen grid
      const bijlageImages = data.images || [];
      html += `<div class="bijlage-images-grid">`;
      for (const img of bijlageImages) {
        html += `<div class="preview-image-placeholder" style="background-image:url(${img});background-size:cover;background-position:center"></div>`;
      }
      html += `<div class="preview-image-placeholder" style="cursor:pointer" onclick="uploadImageForSection('bijlage')">
          <span class="material-symbols-outlined" style="font-size:32px;color:var(--gray-400)">add_photo_alternate</span>
        </div>`;
      html += `</div>`;
    } else {
      // Editable text content
      html += `<div class="section-block-content" contenteditable="true"
            data-section="${sec.id}"
            data-placeholder="Klik op 'Genereer' of typ hier de inhoud van ${sec.title}..."
            id="content-${sec.id}"
            oninput="onSectionEdit('${sec.id}')">${escapeHtml(content).replace(/\n/g, '<br>')}</div>`;
    }

    html += `</div>`;
  }

  area.innerHTML = html;

  // Bind header field changes
  area.querySelectorAll('.header-fields-table input').forEach(input => {
    input.addEventListener('change', () => {
      EDITOR.headerFields[input.dataset.field] = input.value;
      EDITOR.report.headerFields = EDITOR.headerFields;
      saveLocalReports();
    });
  });
}

// ---- SECTION EDITING ----
function onSectionEdit(sectionId) {
  const el = document.getElementById('content-' + sectionId);
  if (!el) return;
  if (!EDITOR.sections[sectionId]) EDITOR.sections[sectionId] = {};
  EDITOR.sections[sectionId].content = el.innerText;
  EDITOR.sections[sectionId].editedManually = true;
  EDITOR.report.sections = EDITOR.sections;
  // Debounced save
  clearTimeout(EDITOR._saveTimeout);
  EDITOR._saveTimeout = setTimeout(() => saveLocalReports(), 1000);
}

// ---- GENERATE SECTION ----
async function generateSingleSection(sectionId) {
  const el = document.getElementById('content-' + sectionId);
  if (el) el.classList.add('generating');

  showToast(`${sectionId} wordt gegenereerd...`);

  try {
    let content = '';

    if (API.isConfigured() && EDITOR.report.serverReportId) {
      const result = await API.generateSection(EDITOR.report.serverReportId, sectionId);
      content = result.content || '';
    } else {
      // Offline: simuleer generatie
      content = generateOfflineContent(sectionId);
    }

    EDITOR.sections[sectionId] = {
      content,
      generated: true,
      generatedAt: new Date().toISOString()
    };
    EDITOR.report.sections = EDITOR.sections;
    saveLocalReports();
    renderSectionList();
    renderEditorDocument();
    showToast(`${sectionId} gegenereerd!`, 'success');

  } catch (err) {
    showToast(`Fout bij genereren: ${err.message}`, 'error');
    if (el) el.classList.remove('generating');
  }
}

async function regenerateSingleSection(sectionId) {
  if (!confirm(`Weet je zeker dat je "${sectionId}" opnieuw wilt genereren? De huidige tekst wordt overschreven.`)) return;

  // Bewaar oude content voor herstel bij fouten
  const backup = EDITOR.sections[sectionId] ? JSON.parse(JSON.stringify(EDITOR.sections[sectionId])) : null;

  delete EDITOR.sections[sectionId];
  EDITOR.report.sections = EDITOR.sections;
  saveLocalReports();

  try {
    await generateSingleSection(sectionId);
  } catch (err) {
    // Herstel oude content als generatie mislukt
    if (backup) {
      EDITOR.sections[sectionId] = backup;
      EDITOR.report.sections = EDITOR.sections;
      saveLocalReports();
      renderEditorDocument();
      showToast('Generatie mislukt – vorige tekst hersteld', 'error');
    }
  }
}

// ---- GENERATE ALL ----
async function generateAllSections() {
  const generatable = MRA_SECTIONS.filter(s => s.generatable);
  showToast(`Alle ${generatable.length} secties worden gegenereerd...`);

  if (API.isConfigured() && EDITOR.report.serverReportId) {
    try {
      const result = await API.generateAllSections(EDITOR.report.serverReportId);
      if (result.results) {
        for (const r of result.results) {
          if (r.content) {
            EDITOR.sections[r.sectionId] = {
              content: r.content,
              generated: true,
              generatedAt: r.generatedAt
            };
          }
        }
      }
      EDITOR.report.sections = EDITOR.sections;
      saveLocalReports();
      renderSectionList();
      renderEditorDocument();
      showToast(`${result.completed || 0} secties gegenereerd!`, 'success');
    } catch (err) {
      showToast(`Fout: ${err.message}`, 'error');
    }
  } else {
    // Offline: genereer alle secties met vertraging
    for (const sec of generatable) {
      if (EDITOR.sections[sec.id] && EDITOR.sections[sec.id].content) continue;
      EDITOR.sections[sec.id] = {
        content: generateOfflineContent(sec.id),
        generated: true,
        generatedAt: new Date().toISOString()
      };
      renderSectionList();
      renderEditorDocument();
      await new Promise(r => setTimeout(r, 300));
    }
    EDITOR.report.sections = EDITOR.sections;
    saveLocalReports();
    showToast('Alle secties gegenereerd (offline modus).', 'success');
  }
}

// ---- SAVE ALL ----
function saveAllSections() {
  // Lees alle contenteditable velden
  MRA_SECTIONS.forEach(sec => {
    const el = document.getElementById('content-' + sec.id);
    if (el && el.innerText.trim()) {
      if (!EDITOR.sections[sec.id]) EDITOR.sections[sec.id] = {};
      EDITOR.sections[sec.id].content = el.innerText;
    }
  });
  EDITOR.report.sections = EDITOR.sections;
  saveLocalReports();

  // Sync naar server
  if (API.isConfigured() && EDITOR.report.serverReportId) {
    for (const [sid, data] of Object.entries(EDITOR.sections)) {
      if (data.editedManually && data.content) {
        API.updateSection(EDITOR.report.serverReportId, sid, data.content)
          .catch(err => console.warn('Sync failed for', sid, err));
      }
    }
    API.updateHeader(EDITOR.report.serverReportId, EDITOR.headerFields)
      .catch(err => console.warn('Header sync failed:', err));
  }

  showToast('Alle secties opgeslagen!', 'success');
}

// ---- ASSEMBLE (EXPORT) ----
async function assembleReport() {
  saveAllSections();

  if (API.isConfigured() && EDITOR.report.serverReportId) {
    showToast('Google Doc wordt samengesteld...');
    try {
      const result = await API.assembleReport(EDITOR.report.serverReportId, EDITOR.headerFields);
      if (result.docUrl) {
        window.open(result.docUrl, '_blank');
        showToast('Google Doc aangemaakt!', 'success');
      }
    } catch (err) {
      showToast(`Fout: ${err.message}`, 'error');
    }
  } else {
    showToast('Export vereist een actieve API verbinding.', 'error');
  }
}

function assembleAndPreview() {
  saveAllSections();
  navigateTo('preview');
}

// ---- IMAGE UPLOAD ----
function uploadImageForSection(sectionId) {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/*';
  input.multiple = sectionId === 'bijlage';
  input.onchange = () => {
    for (const file of input.files) {
      const reader = new FileReader();
      reader.onload = e => {
        if (sectionId === 'bijlage') {
          if (!EDITOR.sections.bijlage) EDITOR.sections.bijlage = { images: [] };
          if (!EDITOR.sections.bijlage.images) EDITOR.sections.bijlage.images = [];
          EDITOR.sections.bijlage.images.push(e.target.result);
        } else if (sectionId === 'nulmeting_images') {
          if (!EDITOR.sections.nulmeting_images) EDITOR.sections.nulmeting_images = {};
          // Vul links of rechts
          if (!EDITOR.sections.nulmeting_images.imageLeft) {
            EDITOR.sections.nulmeting_images.imageLeft = e.target.result;
          } else {
            EDITOR.sections.nulmeting_images.imageRight = e.target.result;
          }
        }
        EDITOR.report.sections = EDITOR.sections;
        saveLocalReports();
        renderEditorDocument();
      };
      reader.readAsDataURL(file);
    }
  };
  input.click();
}

// ---- OFFLINE CONTENT GENERATION ----
function generateOfflineContent(sectionId) {
  const name = EDITOR.headerFields['<<naam vereniging>>'] || 'de vereniging';
  const datum = EDITOR.headerFields['<<datum>>'] || 'onbekende datum';
  const contact = EDITOR.headerFields['<<contactpersoon>>'] || '';

  const templates = {
    inleiding: `Op ${datum} hebben we een bezoek gebracht aan ${name}.\nAan tafel zaten de bestuursleden en betrokken leden.${contact ? '\nContactpersoon: ' + contact + '.' : ''}`,

    nulmeting_dna: `Het DNA werd als volgt beschreven:\nKernwaarden (eigen interpretatie nav het gesprek): saamhorig, betrokken, laagdrempelig\nsterke eigenschappen: locatie, betrokkenheid leden, sociale activiteiten\naandachtspunten: ledenaantal, financiën, zichtbaarheid`,

    nulmeting_quickscan: `De door de procesbegeleider gemaakte quick scan:\nDe passie om de club overeind te houden is groot. Er is een open houding naar elkaar.\nDe aanwezigen hebben een duidelijke persoonlijke drive.\nEr wordt vooral in stellingen gesproken.\nNieuwe acties worden vooral persoonlijk en incidenteel genomen.`,

    positie: `${name} is een actieve vereniging met betrokken leden.\nHet bestuur bestaat uit een klein aantal bestuursleden.\nDe accommodatie is in redelijke staat.\nDe financiële situatie is stabiel maar kent uitdagingen.\nEr is behoefte aan meer zichtbaarheid en ledenaanwas.\nVrijwilligers zijn de drijvende kracht achter de organisatie.\nDe communicatie kan worden verbeterd.`,

    ambitie: `Er is een stip gekozen voor de komende jaren.\nDe vereniging is financieel gezond.\nEr is een helder beleidsplan met duidelijke doelen.\nDe communicatie is professioneel en bereikt alle leden.\nEr is een actieve betrokkenheid van alle leden.\nDe accommodatie is goed onderhouden.\nDe organisatie is zichtbaar in de gemeenschap.`,

    advies: `Einstein zei al: als je doet wat je deed dan krijg je wat je kreeg. ${name} mag durven nieuwe paden te bewandelen en verder te kijken dan de successen en emoties uit het verleden. Ga met de tijd mee en kijk goed om je heen. Neem de tijd om met elkaar hierover in gesprek te gaan en neem de leden hierin stap voor stap mee (zie ontwikkelmodel van Lev Vygotsky). Het ontwikkelen van een "roadmap" is hiervoor een mooi instrument.\nVoorkom dat men in algemeenheden blijven hangen en zich niet persoonlijk verantwoordelijk voelen voor de benodigde acties. Maak zaken concreet via het principe uitspreken – afspreken – aanspreken.\nBenut de zelfdeterminatietheorie van Ryan & Deci. Stimuleer autonomie bij de betrokkenen binnen de afgesproken kaders, stimuleer de onderlinge verbondenheid door een gezamenlijke focus te hebben en zet mensen in hun kracht.\nZorg dat houding (wat je voorstaat) en gedrag (wat je uiteindelijk in handelen laat zien) congruent zijn. Dan krijgen geformuleerde doelen een krachtige inhoud.`,

    ondersteuning: `De initiële wens is om aan de slag te gaan met de geformuleerde doelen.\nOp basis van het gesprek was de conclusie om breder te kijken en eerst te werken aan een herkenbaar gemeenschappelijk doel.\nDe afspraak is dat het bestuur de informatie op zich laat inwerken en beslist waar men als eerste aan wilt beginnen.`,

    experts: `Wij komen met een voorstel zodra intern duidelijk is aan welke opdracht men als eerste wilt gaan werken.`
  };

  return templates[sectionId] || `[Inhoud voor ${sectionId} wordt gegenereerd vanuit de transcriptie]`;
}
