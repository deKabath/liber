/**
 * LiBeR Verslaggenerator – Main App Controller
 */

// ---- STATE ----
const APP = {
  currentPage: 'dashboard',
  currentReportId: null,
  reports: [],
  audioFile: null,
  uploadedImages: [],
  template: null
};

// ---- MRA Template (offline fallback) ----
const MRA_SECTIONS = [
  { id: 'header',             title: 'Basisgegevens',                generatable: false, page: 1 },
  { id: 'inleiding',          title: 'Inleiding',                   generatable: true,  page: 1 },
  { id: 'nulmeting_dna',      title: 'Nulmeting – Het DNA',         generatable: true,  page: 1 },
  { id: 'nulmeting_quickscan',title: 'Nulmeting – Quick Scan',      generatable: true,  page: 1 },
  { id: 'positie',            title: 'Positie van de organisatie',   generatable: true,  page: 2 },
  { id: 'nulmeting_images',   title: '0-meting / Stip op de horizon',generatable: false, page: 2, allowImages: true },
  { id: 'ambitie',            title: 'Ambitie',                     generatable: true,  page: 2 },
  { id: 'advies',             title: 'Advies',                      generatable: true,  page: 3 },
  { id: 'ondersteuning',      title: 'Ondersteuning Rabobank',      generatable: true,  page: 3 },
  { id: 'experts',            title: 'Voorgestelde expert(s)',       generatable: true,  page: 3 },
  { id: 'bijlage',            title: 'Bijlage',                     generatable: false, page: 4, allowImages: true }
];

// ---- NAVIGATION ----
function navigateTo(page, reportId) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));

  const pageEl = document.getElementById('page-' + page);
  if (pageEl) pageEl.classList.add('active');

  const navLink = document.querySelector(`[data-page="${page}"]`);
  if (navLink) navLink.classList.add('active');

  APP.currentPage = page;

  if (page === 'dashboard') loadDashboard();
  if (page === 'editor' && reportId) openEditor(reportId);
  if (page === 'preview') renderPreview();
}

// Nav links
document.querySelectorAll('.nav-link').forEach(link => {
  link.addEventListener('click', e => {
    e.preventDefault();
    navigateTo(link.dataset.page);
  });
});

// ---- TOAST ----
function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 4000);
}

// ---- DASHBOARD ----
async function loadDashboard() {
  if (API.isConfigured()) {
    try {
      const data = await API.listReports();
      APP.reports = (data.reports || []);
    } catch (err) {
      console.warn('API niet bereikbaar, gebruik lokale data:', err);
    }
  }
  renderDashboard();
}

function renderDashboard() {
  // Stats
  const reports = APP.reports;
  document.getElementById('stat-total').textContent = reports.length;
  document.getElementById('stat-done').textContent = reports.filter(r => r.status === 'done').length;
  document.getElementById('stat-progress').textContent = reports.filter(r => ['created','generating','chunking'].includes(r.status)).length;
  document.getElementById('stat-transcribing').textContent = reports.filter(r => ['transcribing','preparing','ready'].includes(r.status)).length;

  // Table
  const tbody = document.getElementById('reports-tbody');
  if (reports.length === 0) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="5">Geen verslagen gevonden. <a href="#" onclick="navigateTo(\'create\')">Maak een nieuw verslag aan.</a></td></tr>';
    return;
  }

  tbody.innerHTML = reports.map(r => {
    const statusMap = {
      'done':         '<span class="status-badge status-done">✓ Voltooid</span>',
      'created':      '<span class="status-badge status-progress">📝 Aangemaakt</span>',
      'generating':   '<span class="status-badge status-progress">⏳ Genereren...</span>',
      'transcribing': '<span class="status-badge status-progress">🎙️ Transcriptie...</span>',
      'preparing':    '<span class="status-badge status-progress">⏳ Voorbereiden...</span>',
      'error':        '<span class="status-badge status-error">⚠️ Fout</span>',
    };
    const status = statusMap[r.status] || `<span class="status-badge">${r.status}</span>`;
    const date = r.createdAt ? new Date(r.createdAt).toLocaleDateString('nl-NL', { day: 'numeric', month: 'short', year: 'numeric' }) : '-';

    return `<tr>
      <td><strong>${escapeHtml(r.meetingName)}</strong></td>
      <td>${status}</td>
      <td>Rabobank MRA</td>
      <td>${date}</td>
      <td>
        <button class="btn btn-ghost btn-sm" onclick="navigateTo('editor','${r.reportId}')">
          <span class="material-symbols-outlined" style="font-size:16px">edit</span> Editor
        </button>
        ${r.docId ? `<a href="https://docs.google.com/document/d/${r.docId}/edit" target="_blank" class="btn btn-ghost btn-sm">
          <span class="material-symbols-outlined" style="font-size:16px">open_in_new</span> Doc
        </a>` : ''}
        <button class="btn btn-ghost btn-sm btn-delete" onclick="deleteReport('${r.reportId}')">
          <span class="material-symbols-outlined" style="font-size:16px">delete</span>
        </button>
      </td>
    </tr>`;
  }).join('');
}

// ---- CREATE REPORT ----
async function createReport() {
  const vereniging = document.getElementById('field-vereniging').value.trim();
  const datum = document.getElementById('field-datum').value;
  const contact = document.getElementById('field-contact').value.trim();
  const onderwerp = document.getElementById('field-onderwerp').value.trim();
  const opgesteld = document.getElementById('field-opgesteld').value.trim();

  if (!vereniging) {
    showToast('Vul de naam van de vereniging in.', 'error');
    return;
  }

  if (!API.isConfigured()) {
    showToast('Configureer eerst de API URL in de instellingen.', 'error');
    return;
  }

  const headerFields = {
    '<<naam vereniging>>': vereniging,
    '<<datum>>': datum,
    '<<contactpersoon>>': contact,
    '<<onderwerp>>': onderwerp || 'verslag intake',
    '<<opgesteld_door>>': opgesteld || 'Lutger Brenninkmeijer',
    '<<opdrachtgever>>': 'Rabobank Kring Metropool Regio Amsterdam'
  };

  try {
    showToast('Verslag aanmaken op server...');

    // Maak report aan op server
    const res = await API.createReport({
      meetingName: vereniging,
      template: 'rabobank_mra',
      headerFields: headerFields
    });

    if (!res.reportId) throw new Error('Server gaf geen reportId terug');

    const serverReportId = res.reportId;
    const report = {
      reportId: serverReportId,
      serverReportId: serverReportId,
      meetingName: vereniging,
      status: 'created',
      template: 'rabobank_mra',
      createdAt: new Date().toISOString(),
      headerFields: headerFields,
      sections: {}
    };

    APP.reports.push(report);
    saveLocalReports();

    // Upload audio als er een bestand is geselecteerd
    if (APP.audioFile) {
      showToast('Audio uploaden naar server...');
      report.status = 'uploading';
      saveLocalReports();

      await uploadAudioToServer(serverReportId, APP.audioFile);
      report.status = 'transcribing';
      saveLocalReports();

      showToast(`Audio geüpload! Transcriptie is gestart.`, 'success');
    } else {
      showToast(`Verslag "${vereniging}" aangemaakt. Upload later audio om te transcriberen.`, 'success');
    }

    navigateTo('editor', serverReportId);

  } catch (err) {
    console.error('Create report failed:', err);
    showToast(`Fout bij aanmaken: ${err.message}`, 'error');
  }
}

/**
 * Upload audio bestand naar de server via base64.
 * Splitst in stukken van ~5MB voor transport naar GAS.
 * Server slaat stukken op als tijdelijke Drive-bestanden en voegt ze samen.
 * Toont voortgang in de upload-progress balk.
 */
async function uploadAudioToServer(reportId, file) {
  const MAX_CHUNK_SIZE = 5 * 1024 * 1024; // 5MB base64 per stuk
  const prog = document.getElementById('upload-progress');
  const progLabel = document.getElementById('upload-progress-label');
  const progFill = document.getElementById('upload-progress-fill');
  const progDetail = document.getElementById('upload-progress-detail');

  function showProgress(label, percent, detail) {
    prog.hidden = false;
    prog.className = 'upload-progress';
    progLabel.textContent = label;
    progFill.style.width = percent + '%';
    progDetail.textContent = detail || '';
  }

  function showDone() {
    prog.className = 'upload-progress done';
    progLabel.textContent = '✓ Upload voltooid — transcriptie gestart';
    progFill.style.width = '100%';
    progDetail.textContent = '';
    // Verberg na 5 seconden
    setTimeout(() => { prog.hidden = true; }, 5000);
  }

  showProgress('Audio inlezen...', 0, `${(file.size / 1024 / 1024).toFixed(1)} MB`);

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async function(e) {
      try {
        const base64Full = e.target.result.split(',')[1];
        const fileName = file.name;
        const mimeType = file.type || 'audio/mpeg';

        if (base64Full.length <= MAX_CHUNK_SIZE) {
          showProgress('Audio uploaden...', 50, fileName);
          const result = await API.uploadAudio(reportId, fileName, mimeType, base64Full);
          if (result.error) throw new Error(result.error);
          showDone();
          resolve(result);
        } else {
          const totalChunks = Math.ceil(base64Full.length / MAX_CHUNK_SIZE);

          for (let i = 0; i < totalChunks; i++) {
            const chunk = base64Full.substring(i * MAX_CHUNK_SIZE, (i + 1) * MAX_CHUNK_SIZE);
            const pct = Math.round(((i + 1) / (totalChunks + 1)) * 100); // +1 voor finalize stap
            showProgress(
              `Audio uploaden: deel ${i + 1} van ${totalChunks}`,
              pct,
              `${fileName} — ${pct}%`
            );
            const chunkResult = await API.uploadAudioChunk(reportId, fileName, mimeType, i, totalChunks, chunk);
            if (chunkResult.error) throw new Error(chunkResult.error);
          }

          showProgress('Audio samenvoegen en transcriptie starten...', 95, 'Even geduld...');
          const finalResult = await API.finalizeAudioUpload(reportId);
          if (finalResult.error) throw new Error(finalResult.error);
          showDone();
          resolve(finalResult);
        }
      } catch (err) {
        prog.className = 'upload-progress';
        progLabel.textContent = '⚠ Upload mislukt';
        progFill.style.width = '0%';
        progDetail.textContent = err.message;
        reject(err);
      }
    };
    reader.onerror = () => {
      prog.hidden = true;
      reject(new Error('Kon audiobestand niet lezen'));
    };
    reader.readAsDataURL(file);
  });
}

// ---- DELETE REPORT ----
async function deleteReport(reportId) {
  const report = APP.reports.find(r => r.reportId === reportId);
  const name = report ? report.meetingName : reportId;

  if (!confirm(`Weet je zeker dat je het verslag "${name}" wilt verwijderen?\n\nDit kan niet ongedaan worden gemaakt.`)) return;

  // Verwijder lokaal
  APP.reports = APP.reports.filter(r => r.reportId !== reportId);
  saveLocalReports();

  // Verwijder op server (reportId IS de serverReportId in v12)
  if (API.isConfigured()) {
    try {
      await API.deleteReport(reportId);
    } catch (err) {
      console.warn('Server delete mislukt:', err);
    }
  }

  showToast(`Verslag "${name}" verwijderd.`, 'success');
  renderDashboard();
}

// ---- AUDIO UPLOAD HANDLING ----
const audioDropZone = document.getElementById('audio-drop-zone');
const audioInput = document.getElementById('audio-input');

if (audioDropZone) {
  audioDropZone.addEventListener('click', () => audioInput.click());
  audioDropZone.addEventListener('dragover', e => { e.preventDefault(); audioDropZone.classList.add('drag-over'); });
  audioDropZone.addEventListener('dragleave', () => audioDropZone.classList.remove('drag-over'));
  audioDropZone.addEventListener('drop', e => {
    e.preventDefault();
    audioDropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleAudioFile(e.dataTransfer.files[0]);
  });
  audioInput.addEventListener('change', () => {
    if (audioInput.files.length) handleAudioFile(audioInput.files[0]);
  });
}

function handleAudioFile(file) {
  if (!file.type.startsWith('audio/') && !file.name.match(/\.(mp3|wav|m4a|aac|ogg|flac)$/i)) {
    showToast('Ongeldig audioformaat.', 'error');
    return;
  }
  APP.audioFile = file;
  document.getElementById('audio-preview').hidden = false;
  document.getElementById('audio-filename').textContent = `${file.name} (${(file.size / 1024 / 1024).toFixed(1)} MB)`;
}

function removeAudio() {
  APP.audioFile = null;
  document.getElementById('audio-preview').hidden = true;
  audioInput.value = '';
}

// ---- IMAGES UPLOAD ----
const imagesDropZone = document.getElementById('images-drop-zone');
const imagesInput = document.getElementById('images-input');

if (imagesDropZone) {
  imagesDropZone.addEventListener('click', () => imagesInput.click());
  imagesDropZone.addEventListener('dragover', e => { e.preventDefault(); imagesDropZone.classList.add('drag-over'); });
  imagesDropZone.addEventListener('dragleave', () => imagesDropZone.classList.remove('drag-over'));
  imagesDropZone.addEventListener('drop', e => {
    e.preventDefault();
    imagesDropZone.classList.remove('drag-over');
    handleImageFiles(e.dataTransfer.files);
  });
  imagesInput.addEventListener('change', () => handleImageFiles(imagesInput.files));
}

function handleImageFiles(files) {
  for (const file of files) {
    if (!file.type.startsWith('image/')) continue;
    const reader = new FileReader();
    reader.onload = e => {
      APP.uploadedImages.push({ file, dataUrl: e.target.result, position: 'bijlage' });
      renderImagesGrid();
    };
    reader.readAsDataURL(file);
  }
}

function renderImagesGrid() {
  const grid = document.getElementById('images-grid');
  grid.innerHTML = APP.uploadedImages.map((img, i) => `
    <div class="image-thumb">
      <img src="${img.dataUrl}" alt="">
      <button class="remove-btn" onclick="removeImage(${i})">×</button>
    </div>
  `).join('');
}

function removeImage(index) {
  APP.uploadedImages.splice(index, 1);
  renderImagesGrid();
}

// ---- LOCAL STORAGE ----
function saveLocalReports() {
  const toSave = APP.reports.map(r => ({
    ...r,
    // Don't store dataUrls in localStorage (too large)
  }));
  try {
    localStorage.setItem('liber_reports', JSON.stringify(toSave));
  } catch (e) { console.warn('LocalStorage full:', e); }
}

function loadLocalReports() {
  try {
    const stored = localStorage.getItem('liber_reports');
    if (stored) APP.reports = JSON.parse(stored);
  } catch (e) { console.warn('LocalStorage parse error:', e); }
}

// ---- SETTINGS CHECK ----
// v12.5: API URL is nu hardcoded als default. Deze functie is alleen nog nodig
// als iemand handmatig de localStorage wist.
function checkApiConfig() {
  if (!API.isConfigured()) {
    console.warn('[App] Geen API URL geconfigureerd - dit zou niet moeten gebeuren.');
    // Fallback: probeer opnieuw de default in te stellen
    API.setApiUrl(API.getApiUrl() || '');
  }
}

// ---- UTIL ----
function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  return d.toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' });
}

// ---- INIT ----
document.addEventListener('DOMContentLoaded', () => {
  loadLocalReports();
  checkApiConfig();
  navigateTo('dashboard');
});
