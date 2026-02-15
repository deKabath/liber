/**
 * LiBeR Verslaggenerator – Preview Module
 * Rendert het volledige Rabobank MRA verslag als print-ready preview.
 */

function renderPreview() {
  const container = document.getElementById('preview-container');
  const hf = EDITOR.headerFields || {};
  const sec = EDITOR.sections || {};
  const name = hf['<<naam vereniging>>'] || EDITOR.report?.meetingName || '';

  container.innerHTML = '';

  // ---- PAGE 1: Header + Inleiding + Nulmeting ----
  container.innerHTML += buildPreviewPage(1, `
    <div class="page-logo">Li<span style="color:var(--orange)">Be</span>R</div>
    <div class="page-divider"></div>

    <table class="preview-header-table">
      <tr><td>Opdrachtgever:</td><td>${esc(hf['<<opdrachtgever>>'] || 'Rabobank Kring Metropool Regio Amsterdam')}</td></tr>
      <tr><td>Vereniging:</td><td>${esc(hf['<<naam vereniging>>'] || '')}</td></tr>
      <tr><td>Onderwerp:</td><td>${esc(hf['<<onderwerp>>'] || '')}</td></tr>
      <tr><td>Contact met:</td><td>${esc(hf['<<contactpersoon>>'] || '')}</td></tr>
      <tr><td>Opgesteld door:</td><td>${esc(hf['<<opgesteld_door>>'] || 'Lutger Brenninkmeijer')}</td></tr>
    </table>

    ${sec.inleiding?.content ? `
      <div class="preview-section-title">Inleiding</div>
      <div class="preview-body">${textToHtml(sec.inleiding.content)}</div>
    ` : ''}

    <div class="preview-section-title">Nulmeting</div>
    ${sec.nulmeting_dna?.content ? `
      <p style="font-style:italic;font-size:11px;margin-bottom:6px">Het DNA werd als volgt beschreven:</p>
      <div class="preview-body">${textToHtml(sec.nulmeting_dna.content)}</div>
    ` : ''}
    ${sec.nulmeting_quickscan?.content ? `
      <p style="font-style:italic;font-size:11px;margin:10px 0 6px">De door de procesbegeleider gemaakte quick scan:</p>
      <div class="preview-body">${textToHtml(sec.nulmeting_quickscan.content)}</div>
    ` : ''}

    <span class="preview-rabo-watermark">In samenwerking met Rabobank</span>
  `);

  // ---- PAGE 2: Positie + 0-meting + Ambitie ----
  container.innerHTML += buildPreviewPage(2, `
    <div class="page-logo">Li<span style="color:var(--orange)">Be</span>R</div>
    <div class="page-divider"></div>

    ${sec.positie?.content ? `
      <div class="preview-section-title">Positie van de vereniging</div>
      <div class="preview-body">${textToHtml(sec.positie.content)}</div>
    ` : ''}

    <div class="preview-two-col" style="margin:16px 0">
      <div>
        <h4 style="font-family:var(--font-serif);font-size:12px;font-weight:700;font-style:italic;color:var(--orange);margin-bottom:8px">0-meting</h4>
        ${sec.nulmeting_images?.imageLeft
          ? `<img src="${sec.nulmeting_images.imageLeft}" style="width:100%;border-radius:4px">`
          : '<div class="preview-image-placeholder">Afbeelding 0-meting</div>'}
      </div>
      <div>
        <h4 style="font-family:var(--font-serif);font-size:12px;font-weight:700;font-style:italic;color:var(--orange);margin-bottom:8px">Stip op de horizon</h4>
        ${sec.nulmeting_images?.imageRight
          ? `<img src="${sec.nulmeting_images.imageRight}" style="width:100%;border-radius:4px">`
          : '<div class="preview-image-placeholder">Afbeelding Stip op de horizon</div>'}
      </div>
    </div>

    ${sec.ambitie?.content ? `
      <div class="preview-section-title">Ambitie</div>
      <div class="preview-body">${textToHtml(sec.ambitie.content)}</div>
    ` : ''}

    <span class="preview-rabo-watermark">In samenwerking met Rabobank</span>
  `);

  // ---- PAGE 3: Advies + Ondersteuning + Experts ----
  container.innerHTML += buildPreviewPage(3, `
    <div class="page-logo">Li<span style="color:var(--orange)">Be</span>R</div>
    <div class="page-divider"></div>

    ${sec.advies?.content ? `
      <div class="preview-section-title">Advies</div>
      <div class="preview-body">${textToHtml(sec.advies.content)}</div>
    ` : ''}

    ${sec.ondersteuning?.content ? `
      <div class="preview-section-title">Ondersteuning Rabobank</div>
      <div class="preview-body">${textToHtml(sec.ondersteuning.content)}</div>
    ` : ''}

    ${sec.experts?.content ? `
      <div class="preview-section-title">Voorgestelde expert(s)</div>
      <div class="preview-body">${textToHtml(sec.experts.content)}</div>
    ` : ''}

    <span class="preview-rabo-watermark">In samenwerking met Rabobank</span>
  `);

  // ---- PAGE 4+: Bijlage ----
  const bijlageImages = sec.bijlage?.images || [];
  let bijlageHtml = `
    <div class="page-logo">Li<span style="color:var(--orange)">Be</span>R</div>
    <div class="page-divider"></div>
    <div class="preview-section-title">Bijlage:</div>
  `;

  if (bijlageImages.length > 0) {
    bijlageHtml += '<div class="bijlage-images-grid">';
    for (const img of bijlageImages) {
      bijlageHtml += `<div style="border-radius:4px;overflow:hidden;border:1px solid var(--gray-200)">
        <img src="${img}" style="width:100%;height:180px;object-fit:cover">
      </div>`;
    }
    bijlageHtml += '</div>';
  } else {
    bijlageHtml += '<p style="color:var(--gray-400);font-style:italic;font-size:11px;margin-top:12px">[Geen bijlagen toegevoegd]</p>';
  }

  bijlageHtml += '<span class="preview-rabo-watermark">In samenwerking met Rabobank</span>';
  container.innerHTML += buildPreviewPage(4, bijlageHtml);
}

// ---- HELPERS ----

function buildPreviewPage(pageNum, content) {
  return `<div class="preview-page">
    ${content}
    <span class="preview-page-number">${pageNum}</span>
  </div>`;
}

function esc(str) {
  if (!str) return '';
  const d = document.createElement('div');
  d.textContent = str;
  return d.innerHTML;
}

function textToHtml(text) {
  if (!text) return '';
  return text.split('\n')
    .filter(line => line.trim())
    .map(line => `<p>${esc(line.trim())}</p>`)
    .join('');
}
