/**
 * LiBeR Verslaggenerator – API Layer
 * Communiceert met de Google Apps Script Web App backend.
 *
 * v12.6 – Vaste deployment URL. Geen localStorage meer om cache-problemen te voorkomen.
 */

const API = (() => {

  // ── Vaste deployment URL (v12.6 @13) ── Update bij elke deploy ──
  const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbz0AFVa1KwXEqTM4OQ70mYFdMBLOWLYNyZewFSWwnYpugJCTWwctJKhiQUtKwfP5gJsHg/exec';

  // Ruim eventuele oude localStorage op
  localStorage.removeItem('liber_api_url');

  function getApiUrl() { return WEBAPP_URL; }
  function isConfigured() { return true; }

  async function get(action, params = {}) {
    const url = new URL(WEBAPP_URL);
    url.searchParams.set('action', action);
    for (const [k, v] of Object.entries(params)) {
      url.searchParams.set(k, v);
    }
    const res = await fetch(url.toString());
    if (!res.ok) throw new Error(`API fout: ${res.status}`);
    return res.json();
  }

  async function post(action, data = {}) {
    const res = await fetch(WEBAPP_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ action, ...data })
    });
    if (!res.ok) throw new Error(`API fout: ${res.status}`);
    return res.json();
  }

  // ---- API Endpoints ----

  async function status()                   { return get('status'); }
  async function getTemplate()              { return get('getTemplate'); }
  async function listReports()              { return get('listReports'); }
  async function getReport(reportId)        { return get('getReport', { reportId }); }
  async function getSections(reportId)      { return get('getSections', { reportId }); }
  async function getSection(reportId, sid)  { return get('getSection', { reportId, sectionId: sid }); }
  async function getTranscriptStatus(rid)   { return get('getTranscriptStatus', { reportId: rid }); }

  async function createReport(data)                   { return post('createReport', { data }); }
  async function generateSection(reportId, sectionId) { return post('generateSection', { reportId, sectionId }); }
  async function regenerateSection(rid, sid, ctx)      { return post('regenerateSection', { reportId: rid, sectionId: sid, extraContext: ctx || '' }); }
  async function generateAllSections(reportId)        { return post('generateAllSections', { reportId }); }
  async function updateSection(rid, sid, content)      { return post('updateSection', { reportId: rid, sectionId: sid, content }); }
  async function updateHeader(reportId, fields)        { return post('updateHeader', { reportId, fields }); }
  async function assembleReport(reportId, headerFields){ return post('assembleReport', { reportId, headerFields }); }
  async function insertImage(rid, fileId, pos, idx)    { return post('insertImage', { reportId: rid, imageFileId: fileId, position: pos, index: idx }); }
  async function deleteReport(reportId)                 { return post('deleteReport', { reportId }); }

  // v12: Audio upload
  async function uploadAudio(reportId, fileName, mimeType, audioBase64) {
    return post('uploadAudio', { reportId, fileName, mimeType, audioBase64 });
  }
  async function uploadAudioChunk(reportId, fileName, mimeType, chunkIndex, totalChunks, chunkBase64) {
    return post('uploadAudioChunk', { reportId, fileName, mimeType, chunkIndex, totalChunks, chunkBase64 });
  }
  async function finalizeAudioUpload(reportId) {
    return post('finalizeAudioUpload', { reportId });
  }
  async function getReportStatus(reportId)               { return get('getReportStatus', { reportId }); }
  async function checkTranscription(reportId)            { return get('checkTranscription', { reportId }); }

  return {
    getApiUrl, isConfigured,
    status, getTemplate, listReports, getReport,
    getSections, getSection, getTranscriptStatus,
    createReport, generateSection, regenerateSection,
    generateAllSections, updateSection, updateHeader,
    assembleReport, insertImage, deleteReport,
    uploadAudio, uploadAudioChunk, finalizeAudioUpload,
    getReportStatus, checkTranscription
  };
})();
