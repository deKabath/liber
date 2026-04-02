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
  headerFields: {},
  transcriptReady: false  // Buffer: true = transcriptie volledig, mag genereren
};

// Haal de juiste template secties op voor het huidige rapport
function getEditorSections() {
  const templateId = (EDITOR.report && EDITOR.report.template) || (EDITOR.report && EDITOR.report.templateId) || 'rabobank_mra';
  return getTemplateSections(templateId);
}

// ---- GENERATION PROGRESS TRACKER ----
const GenProgress = (() => {
  let _timer = null;
  let _startTime = 0;
  let _stepStartTime = 0;
  let _steps = [];
  let _currentIdx = -1;
  let _total = 0;
  let _completed = 0;
  let _durations = []; // voor ETA berekening

  const TIPS = [
    'AI analyseert de transcriptie en schrijft de inhoud...',
    'Elke sectie wordt individueel gegenereerd vanuit de transcriptie.',
    'De AI controleert of het minimum aantal woorden is bereikt.',
    'Secties die te kort zijn worden automatisch hergegenereerd.',
    'Het verslag volgt het Rabobank MRA template.',
    'Na generatie kun je elke sectie handmatig aanpassen.',
  ];

  function show(sectionIds, sectionTitles) {
    _steps = sectionIds.map((id, i) => ({ id, title: sectionTitles[i], status: 'pending', duration: null, words: null }));
    _total = _steps.length;
    _completed = 0;
    _currentIdx = -1;
    _durations = [];
    _startTime = Date.now();

    const modal = document.getElementById('modal-generation');
    modal.hidden = false;

    document.getElementById('gen-title').textContent =
      _total === 1 ? 'Sectie genereren' : `${_total} secties genereren`;
    document.getElementById('gen-subtitle').textContent = 'Voorbereiden...';
    document.getElementById('gen-progress-bar').style.width = '0%';
    document.getElementById('gen-percent').textContent = '0%';
    document.getElementById('gen-elapsed').textContent = '0:00';
    document.getElementById('gen-eta').textContent = 'Berekenen...';
    document.getElementById('gen-counter').textContent = `0 / ${_total}`;
    document.getElementById('gen-tip').textContent = TIPS[0];

    _renderSteps();
    _startTimer();
  }

  function startStep(sectionId) {
    _currentIdx = _steps.findIndex(s => s.id === sectionId);
    if (_currentIdx === -1) return;
    _stepStartTime = Date.now();
    _steps[_currentIdx].status = 'active';
    const title = _steps[_currentIdx].title;
    document.getElementById('gen-subtitle').textContent = `Genereren: ${title}`;
    document.getElementById('gen-tip').textContent = TIPS[(_completed + 1) % TIPS.length];
    _renderSteps();
    // Scroll actieve stap in zicht
    const stepsEl = document.getElementById('gen-steps');
    const activeEl = stepsEl.querySelector('.gen-step.active');
    if (activeEl) activeEl.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }

  function completeStep(sectionId, wordCount) {
    const idx = _steps.findIndex(s => s.id === sectionId);
    if (idx === -1) return;
    const dur = (Date.now() - _stepStartTime) / 1000;
    _steps[idx].status = 'done';
    _steps[idx].duration = dur;
    _steps[idx].words = wordCount || null;
    _completed++;
    _durations.push(dur);

    const pct = Math.round((_completed / _total) * 100);
    document.getElementById('gen-progress-bar').style.width = pct + '%';
    document.getElementById('gen-percent').textContent = pct + '%';
    document.getElementById('gen-counter').textContent = `${_completed} / ${_total}`;

    // ETA
    const avgDur = _durations.reduce((a, b) => a + b, 0) / _durations.length;
    const remaining = _total - _completed;
    if (remaining > 0) {
      const etaSec = Math.round(avgDur * remaining);
      document.getElementById('gen-eta').textContent = etaSec < 60
        ? `~${etaSec}s resterend`
        : `~${Math.ceil(etaSec / 60)}min resterend`;
    } else {
      document.getElementById('gen-eta').textContent = 'Klaar!';
    }

    _renderSteps();
  }

  function failStep(sectionId) {
    const idx = _steps.findIndex(s => s.id === sectionId);
    if (idx === -1) return;
    _steps[idx].status = 'error';
    _steps[idx].duration = (Date.now() - _stepStartTime) / 1000;
    _completed++;
    _renderSteps();
    document.getElementById('gen-counter').textContent = `${_completed} / ${_total}`;
  }

  function hide() {
    clearInterval(_timer);
    _timer = null;
    document.getElementById('modal-generation').hidden = true;
  }

  function _startTimer() {
    clearInterval(_timer);
    _timer = setInterval(() => {
      const elapsed = Math.floor((Date.now() - _startTime) / 1000);
      const min = Math.floor(elapsed / 60);
      const sec = elapsed % 60;
      document.getElementById('gen-elapsed').textContent = `${min}:${sec.toString().padStart(2, '0')}`;
    }, 1000);
  }

  function _renderSteps() {
    const container = document.getElementById('gen-steps');
    container.innerHTML = _steps.map(s => {
      const icon = s.status === 'done' ? 'check_circle'
        : s.status === 'active' ? 'auto_awesome'
        : s.status === 'error' ? 'error'
        : 'radio_button_unchecked';
      const durStr = s.duration != null ? `${s.duration.toFixed(1)}s` : '';
      const wordStr = s.words != null ? `${s.words} wrd` : '';
      return `<div class="gen-step ${s.status}">
        <span class="material-symbols-outlined step-icon">${icon}</span>
        <span class="step-name">${s.title}</span>
        <span class="step-words">${wordStr}</span>
        <span class="step-time">${durStr}</span>
      </div>`;
    }).join('');
  }

  return { show, startStep, completeStep, failStep, hide };
})();

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

  // ── TRANSCRIPTIE BUFFER CHECK ──
  // Check transcriptie-status en blokkeer generatie als deze niet klaar is
  EDITOR.transcriptReady = false; // Start pessimistisch
  const serverId = EDITOR.report.serverReportId || EDITOR.report.reportId;

  if (API.isConfigured() && serverId && !serverId.startsWith('local_')) {
    // Check transcriptie completeness
    checkAndUpdateTranscriptStatus(serverId);

    // Start polling als status 'transcribing' of 'uploading' is
    if (EDITOR.report.status === 'transcribing' || EDITOR.report.status === 'uploading') {
      startTranscriptStatusPolling();
    }

    // Probeer secties van server te laden
    API.getSections(serverId).then(data => {
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

/**
 * Controleer bij de server of de transcriptie volledig is.
 * Update EDITOR.transcriptReady en de UI dienovereenkomstig.
 */
async function checkAndUpdateTranscriptStatus(serverId) {
  try {
    const status = await API.getReportStatus(serverId);

    if (status.status === 'transcribed' || status.status === 'done') {
      // Transcriptie is klaar - extra verificatie
      if (status.hasTranscript) {
        EDITOR.transcriptReady = true;
        updateGenerateButtonsState(true);
        updateTranscriptStatusUI(status);
      } else {
        EDITOR.transcriptReady = false;
        updateGenerateButtonsState(false, 'Transcriptie voltooid maar bestand ontbreekt');
      }
    } else if (status.status === 'transcribing') {
      EDITOR.transcriptReady = false;
      const pl = status.pipeline || {};
      const eta = pl.eta || {};
      let btnMsg = `Transcriptie bezig: ${status.transcriptProgress || 0}%`;
      if (eta.remainingSec > 0) btnMsg += ` (nog ~${eta.formatted})`;
      else btnMsg += ` (${status.transcriptCurrent || 0}/${status.transcriptFragments || '?'} fragmenten)`;
      updateGenerateButtonsState(false, btnMsg);
      updateTranscriptStatusUI(status);
      // Start polling als dat nog niet loopt
      if (!_statusPollInterval) startTranscriptStatusPolling();
    } else if (status.status === 'created') {
      EDITOR.transcriptReady = false;
      updateGenerateButtonsState(false, 'Upload eerst een audio-opname om secties te genereren');
      updateTranscriptStatusUI(status); // Verberg pipeline tracker
    } else if (status.status === 'error') {
      EDITOR.transcriptReady = false;
      updateGenerateButtonsState(false, 'Fout bij transcriptie: ' + (status.error || 'onbekend'));
      updateTranscriptStatusUI(status);
    } else {
      // Onbekende status - laat generatie toe als er al secties zijn
      const hasAnySections = Object.values(EDITOR.sections).some(s => s.content);
      EDITOR.transcriptReady = hasAnySections || status.hasTranscript;
      updateGenerateButtonsState(EDITOR.transcriptReady);
    }
  } catch (err) {
    console.warn('Transcriptie status check mislukt:', err);
    // Bij fout: laat generatie toe als er al secties gegenereerd zijn
    const hasAnySections = Object.values(EDITOR.sections).some(s => s.content);
    EDITOR.transcriptReady = hasAnySections;
    updateGenerateButtonsState(hasAnySections);
  }
}

/**
 * Enable/disable alle genereer-knoppen in de editor
 */
function updateGenerateButtonsState(enabled, reason) {
  // Sidebar "Alles Genereren" knop
  const genAllBtn = document.querySelector('.sidebar-footer .btn-primary');
  if (genAllBtn) {
    genAllBtn.disabled = !enabled;
    if (!enabled && reason) {
      genAllBtn.title = reason;
      genAllBtn.innerHTML = `<span class="material-symbols-outlined">hourglass_empty</span> Wacht op transcriptie`;
    } else {
      genAllBtn.title = '';
      genAllBtn.innerHTML = `<span class="material-symbols-outlined">auto_awesome</span> Alles Genereren`;
    }
  }

  // Individuele "Genereer" knoppen in sectie-headers
  document.querySelectorAll('.section-block-actions .btn-primary').forEach(btn => {
    btn.disabled = !enabled;
    if (!enabled) {
      btn.title = reason || 'Wacht op transcriptie';
    } else {
      btn.title = 'Genereer vanuit transcriptie';
    }
  });
}

// ---- SECTION LIST (SIDEBAR) ----
function renderSectionList() {
  const list = document.getElementById('section-list');
  let generatedCount = 0;

  list.innerHTML = getEditorSections().map(sec => {
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

  document.getElementById('section-progress').textContent = `${generatedCount}/${getEditorSections().length}`;
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
  for (const sec of getEditorSections()) {
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
      const isOffline = data && data.offline;
      html += `<div class="section-block-content${isOffline ? ' offline-content' : ''}" contenteditable="true"
            data-section="${sec.id}"
            data-placeholder="Klik op 'Genereer' of typ hier de inhoud van ${sec.title}..."
            id="content-${sec.id}"
            oninput="onSectionEdit('${sec.id}')">${escapeHtml(content).replace(/\n/g, '<br>')}</div>`;
      if (isOffline) {
        html += `<div class="offline-badge">⚠️ Voorbeeldtekst – niet uit transcriptie</div>`;
      }
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
async function generateSingleSection(sectionId, opts = {}) {
  // ── TRANSCRIPTIE BUFFER CHECK ──
  if (!EDITOR.transcriptReady && !opts._bulkMode) {
    showToast('Transcriptie is nog niet volledig. Wacht tot alle audio is getranscribeerd.', 'error');
    return;
  }

  const el = document.getElementById('content-' + sectionId);
  if (el) el.classList.add('generating');

  // Toon voortgangsmodal (tenzij onderdeel van bulk-generatie)
  const showProgress = !opts._bulkMode;
  const sectionDef = getEditorSections().find(s => s.id === sectionId);
  const sectionTitle = sectionDef ? sectionDef.title : sectionId;

  if (showProgress) {
    GenProgress.show([sectionId], [sectionTitle]);
    GenProgress.startStep(sectionId);
  }

  try {
    let content = '';
    let isOffline = false;

    const serverId = EDITOR.report.serverReportId || EDITOR.report.reportId;
    if (API.isConfigured() && serverId && !serverId.startsWith('local_')) {
      const result = await API.generateSection(serverId, sectionId);
      content = result.content || '';
    } else {
      // Offline: simuleer generatie met sjabloontekst (NIET uit transcriptie)
      content = generateOfflineContent(sectionId);
      isOffline = true;
    }

    const wordCount = content.trim().split(/\s+/).length;

    EDITOR.sections[sectionId] = {
      content,
      generated: true,
      offline: isOffline,
      generatedAt: new Date().toISOString()
    };
    EDITOR.report.sections = EDITOR.sections;
    saveLocalReports();
    renderSectionList();
    renderEditorDocument();

    if (showProgress) {
      GenProgress.completeStep(sectionId, wordCount);
      setTimeout(() => GenProgress.hide(), 1200);
    }

    if (isOffline) {
      showToast(`${sectionTitle} – VOORBEELDTEKST (niet uit transcriptie). Verbind de API om echte content te genereren.`, 'warning');
    } else {
      showToast(`${sectionTitle} gegenereerd (${wordCount} woorden)`, 'success');
    }

  } catch (err) {
    if (showProgress) {
      GenProgress.failStep(sectionId);
      setTimeout(() => GenProgress.hide(), 2000);
    }
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
  // ── TRANSCRIPTIE BUFFER CHECK ──
  if (!EDITOR.transcriptReady) {
    showToast('Transcriptie is nog niet volledig. Wacht tot alle audio is getranscribeerd voordat je secties genereert.', 'error');
    return;
  }

  const generatable = getEditorSections().filter(s => s.generatable);
  // Filter secties die al gegenereerd zijn
  const toGenerate = generatable.filter(s => !EDITOR.sections[s.id] || !EDITOR.sections[s.id].content);

  if (toGenerate.length === 0) {
    showToast('Alle secties zijn al gegenereerd.', 'success');
    return;
  }

  // Toon progress modal
  GenProgress.show(
    toGenerate.map(s => s.id),
    toGenerate.map(s => s.title)
  );

  const serverId = EDITOR.report.serverReportId || EDITOR.report.reportId;
  const isOnline = API.isConfigured() && serverId && !serverId.startsWith('local_');
  let completed = 0;

  for (const sec of toGenerate) {
    GenProgress.startStep(sec.id);

    try {
      let content = '';
      let isOffline = false;

      if (isOnline) {
        const result = await API.generateSection(serverId, sec.id);
        content = result.content || '';
      } else {
        // Offline: sjabloontekst
        await new Promise(r => setTimeout(r, 400)); // simuleer vertraging
        content = generateOfflineContent(sec.id);
        isOffline = true;
      }

      const wordCount = content.trim().split(/\s+/).length;

      EDITOR.sections[sec.id] = {
        content,
        generated: true,
        offline: isOffline,
        generatedAt: new Date().toISOString()
      };
      EDITOR.report.sections = EDITOR.sections;
      saveLocalReports();
      renderSectionList();
      renderEditorDocument();

      GenProgress.completeStep(sec.id, wordCount);
      completed++;

    } catch (err) {
      GenProgress.failStep(sec.id);
      console.error(`Fout bij ${sec.id}:`, err);
    }
  }

  // Alles klaar
  EDITOR.report.sections = EDITOR.sections;
  saveLocalReports();

  setTimeout(() => {
    GenProgress.hide();
    if (!isOnline) {
      showToast('⚠️ VOORBEELDTEKST gegenereerd – NIET uit audio/transcriptie!', 'warning');
    } else {
      showToast(`${completed} secties gegenereerd uit transcriptie!`, 'success');
    }
  }, 1500);
}

// ---- SAVE ALL ----
function saveAllSections() {
  // Lees alle contenteditable velden
  getEditorSections().forEach(sec => {
    const el = document.getElementById('content-' + sec.id);
    if (el && el.innerText.trim()) {
      if (!EDITOR.sections[sec.id]) EDITOR.sections[sec.id] = {};
      EDITOR.sections[sec.id].content = el.innerText;
    }
  });
  EDITOR.report.sections = EDITOR.sections;
  saveLocalReports();

  // Sync naar server
  const syncId = EDITOR.report.serverReportId || EDITOR.report.reportId;
  if (API.isConfigured() && syncId && !syncId.startsWith('local_')) {
    for (const [sid, sdata] of Object.entries(EDITOR.sections)) {
      if (sdata.editedManually && sdata.content) {
        API.updateSection(syncId, sid, sdata.content)
          .catch(err => console.warn('Sync failed for', sid, err));
      }
    }
    API.updateHeader(syncId, EDITOR.headerFields)
      .catch(err => console.warn('Header sync failed:', err));
  }

  showToast('Alle secties opgeslagen!', 'success');
}

// ---- ASSEMBLE (EXPORT) ----
async function assembleReport() {
  saveAllSections();

  const assembleId = EDITOR.report.serverReportId || EDITOR.report.reportId;
  if (API.isConfigured() && assembleId && !assembleId.startsWith('local_')) {
    showToast('Google Doc wordt samengesteld...');
    try {
      const result = await API.assembleReport(assembleId, EDITOR.headerFields);
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

// ---- DELETE CURRENT REPORT FROM EDITOR ----
function deleteCurrentReport() {
  if (!EDITOR.reportId) return;
  deleteReport(EDITOR.reportId);
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

// ---- TRANSCRIPTIE STATUS POLLING ----
let _statusPollInterval = null;

function startTranscriptStatusPolling() {
  stopTranscriptStatusPolling();

  const serverId = EDITOR.report.serverReportId || EDITOR.report.reportId;
  if (!serverId || serverId.startsWith('local_') || !API.isConfigured()) return;

  async function poll() {
    try {
      const status = await API.getReportStatus(serverId);
      updateTranscriptStatusUI(status);

      if (status.status === 'transcribed' || status.status === 'done' || status.status === 'generating') {
        // Transcriptie klaar - stop polling, activeer generatie knoppen
        stopTranscriptStatusPolling();
        EDITOR.transcriptReady = true;
        updateGenerateButtonsState(true);

        if (status.status === 'transcribed' && status.hasTranscript) {
          showToast('Transcriptie voltooid! Je kunt nu secties genereren.', 'success');
        }
      } else if (status.status === 'transcribing') {
        // Update voortgang in knoppen met pipeline info
        EDITOR.transcriptReady = false;
        const pl = status.pipeline || {};
        const eta = pl.eta || {};
        let btnMsg = `Transcriptie: ${status.transcriptProgress || 0}%`;
        if (eta.remainingSec > 0) btnMsg += ` (nog ~${eta.formatted})`;
        else btnMsg += ` (${status.transcriptCurrent || 0}/${status.transcriptFragments || '?'} fragmenten)`;
        updateGenerateButtonsState(false, btnMsg);
      } else if (status.status === 'error') {
        stopTranscriptStatusPolling();
        EDITOR.transcriptReady = false;
        updateGenerateButtonsState(false, 'Transcriptie fout: ' + (status.error || 'onbekend'));
        showToast('Fout bij transcriptie: ' + (status.error || 'onbekend'), 'error');
      }
    } catch (err) {
      console.warn('Status poll mislukt:', err);
    }
  }

  poll(); // Eerste keer direct
  _statusPollInterval = setInterval(poll, 10000); // Elke 10 seconden
}

function stopTranscriptStatusPolling() {
  if (_statusPollInterval) {
    clearInterval(_statusPollInterval);
    _statusPollInterval = null;
  }
}

function updateTranscriptStatusUI(status) {
  const tracker = document.getElementById('pipeline-tracker');
  if (!tracker) return;

  // Verberg als er geen actieve pipeline is
  const hasActiveJob = status.transcriptStatus && status.transcriptStatus !== 'none';
  const isProcessing = status.status === 'transcribing' || status.status === 'error';

  if (!hasActiveJob && !isProcessing) {
    if (status.status === 'transcribed' || status.status === 'done') {
      // Transcript beschikbaar (direct gekoppeld of al klaar) - kort "gereed" tonen
      showPipelineDone_(tracker, status);
      setTimeout(() => { tracker.hidden = true; }, 8000);
    } else {
      tracker.hidden = true;
    }
    return;
  }

  // Toon de tracker
  tracker.hidden = false;
  tracker.className = 'pipeline-tracker';

  const pipeline = status.pipeline || {};
  const steps = pipeline.steps || [];
  const eta = pipeline.eta || {};

  // === HEADER: Stap-label + ETA ===
  const labelEl = document.getElementById('pipeline-step-label');
  const etaEl = document.getElementById('pipeline-eta');

  if (status.status === 'transcribed' || status.status === 'done') {
    showPipelineDone_(tracker, status);
    setTimeout(() => { tracker.hidden = true; }, 8000);
    return;
  }

  if (status.status === 'error') {
    tracker.className = 'pipeline-tracker error';
    labelEl.textContent = 'Fout: ' + (status.error || 'Onbekende fout');
    etaEl.textContent = '';
    etaEl.className = 'pipeline-eta';
    renderPipelineSteps_(steps);
    document.getElementById('pipeline-function').textContent = pipeline.currentFunction || '';
    document.getElementById('pipeline-elapsed').textContent = '';
    return;
  }

  // Actieve status
  labelEl.textContent = pipeline.currentStepLabel || 'Pipeline actief...';

  // ETA
  if (eta.remainingSec > 0) {
    etaEl.textContent = `⏱ ${eta.formatted} resterend`;
    etaEl.className = 'pipeline-eta ' + (eta.confidence || 'medium');
  } else if (eta.remainingSec === 0) {
    etaEl.textContent = '✓ Bijna klaar';
    etaEl.className = 'pipeline-eta high';
  } else {
    etaEl.textContent = '';
    etaEl.className = 'pipeline-eta';
  }

  // === PROGRESS BAR ===
  const progressFill = document.getElementById('pipeline-progress-fill');
  const progress = status.transcriptProgress || 0;
  // Bereken gewogen voortgang: preparing=10%, transcribing=10-90%, finalizing=90-100%
  let visualProgress = progress;
  if (pipeline.currentStep === 'archiving') visualProgress = 3;
  else if (pipeline.currentStep === 'metadata') visualProgress = 6;
  else if (pipeline.currentStep === 'splitting' || pipeline.currentStep === 'converting') visualProgress = 10;
  else if (pipeline.currentStep === 'queued') visualProgress = 12;
  else if (pipeline.currentStep === 'whisper') {
    // 15% - 90% range voor whisper, gebaseerd op fragmenten
    const fragProgress = status.transcriptFragments > 0
      ? (status.transcriptCurrent / status.transcriptFragments) : 0;
    visualProgress = 15 + Math.round(fragProgress * 75);
  } else if (pipeline.currentStep === 'finalizing') visualProgress = 92;
  progressFill.style.width = visualProgress + '%';

  // === STAPPEN VISUALISATIE ===
  renderPipelineSteps_(steps);

  // === DETAIL REGEL ===
  const functionEl = document.getElementById('pipeline-function');
  const elapsedEl = document.getElementById('pipeline-elapsed');

  functionEl.textContent = pipeline.currentFunction || '';

  // Audio duur info
  let detailParts = [];
  if (pipeline.audioDurationFormatted && pipeline.audioDurationSec > 0) {
    detailParts.push(`Audio: ${pipeline.audioDurationFormatted}`);
  }
  if (pipeline.currentFragmentName && pipeline.currentStep === 'whisper') {
    detailParts.push(`Fragment: ${pipeline.currentFragmentName}`);
  }
  if (status.transcriptFragments > 0) {
    detailParts.push(`${status.transcriptCurrent}/${status.transcriptFragments} fragmenten`);
  }
  elapsedEl.textContent = detailParts.join(' · ');
}

function showPipelineDone_(tracker, status) {
  tracker.hidden = false;
  tracker.className = 'pipeline-tracker done';
  document.getElementById('pipeline-step-label').textContent = '✓ Transcriptie voltooid – je kunt nu secties genereren';
  document.getElementById('pipeline-eta').textContent = '';
  document.getElementById('pipeline-eta').className = 'pipeline-eta';
  document.getElementById('pipeline-progress-fill').style.width = '100%';
  document.getElementById('pipeline-function').textContent = '';
  document.getElementById('pipeline-elapsed').textContent = '';

  // Render alle stappen als completed
  const pipeline = status.pipeline || {};
  const steps = (pipeline.steps || []).map(s => ({ ...s, status: 'completed' }));
  renderPipelineSteps_(steps);
}

function renderPipelineSteps_(steps) {
  const container = document.getElementById('pipeline-steps');
  if (!container || !steps.length) {
    if (container) container.innerHTML = '';
    return;
  }

  container.innerHTML = steps.map(step => {
    let icon = step.icon || 'circle';
    // Override icoon op basis van status
    if (step.status === 'completed') icon = 'check_circle';
    else if (step.status === 'error') icon = 'error';
    else if (step.status === 'active') icon = step.icon || 'pending';

    let label = step.label || step.id;
    if (step.detail && step.status === 'active') {
      label += ` (${step.detail})`;
    }

    return `<div class="pipeline-step ${step.status || 'pending'}" title="${step.label}${step.elapsedFormatted ? ' - ' + step.elapsedFormatted + ' verstreken' : ''}">
      <span class="material-symbols-outlined">${icon}</span>
      <span>${label}</span>
    </div>`;
  }).join('');
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
