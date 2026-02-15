/**
 * LiBeR Verslaggenerator – API Layer
 * Communiceert met de Google Apps Script Web App backend.
 *
 * CONFIGURATIE:
 * Stel de WEBAPP_URL in naar je eigen gedeployde Apps Script Web App URL.
 */

const API = (() => {

  // !! WIJZIG DIT naar je eigen Apps Script Web App URL !!
  let WEBAPP_URL = localStorage.getItem('liber_api_url') || '';

  function setApiUrl(url) {
    WEBAPP_URL = url.replace(/\/$/, '');
    localStorage.setItem('liber_api_url', WEBAPP_URL);
  }

  function getApiUrl() { return WEBAPP_URL; }

  function isConfigured() { return !!WEBAPP_URL; }

  async function get(action, params = {}) {
    if (!WEBAPP_URL) throw new Error('API URL niet geconfigureerd. Ga naar Instellingen.');
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
    if (!WEBAPP_URL) throw new Error('API URL niet geconfigureerd. Ga naar Instellingen.');
    const res = await fetch(WEBAPP_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' }, // GAS vereist text/plain
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

  return {
    setApiUrl, getApiUrl, isConfigured,
    status, getTemplate, listReports, getReport,
    getSections, getSection, getTranscriptStatus,
    createReport, generateSection, regenerateSection,
    generateAllSections, updateSection, updateHeader,
    assembleReport, insertImage, deleteReport
  };
})();
