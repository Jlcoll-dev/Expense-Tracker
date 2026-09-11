'use strict';

/* ══════════════════════════════════════════════════════
   ARCHIVE.JS — Archivar movimientos a Drive como .xlsx

   Flujo:
   1. Usuario abre pantalla Archivar
   2. Elige rango de meses (desde / hasta)
   3. Preview: cuántos movimientos abarca
   4. Confirmar → genera xlsx en el navegador con SheetJS
      (una hoja por mes + ResumenAnual, igual a tu Excel original)
   5. Sube el .xlsx a Google Drive
   6. Borra esos movimientos de la app y guarda
   ══════════════════════════════════════════════════════ */

const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

const CAT_ORDER = ['Hogar','Comida','Transporte','Entretenimiento',
                   'Salud','Moto','Salidas','Otros'];
const MET_ORDER = ['Efectivo','Débito','amex','Transferencia',
                   'MercadoPago','Otros','Visa Galicia','Master Galicia'];

/* ── Inject SheetJS from CDN (only when archive screen opens) ── */
let xlsxLoaded = false;
function loadSheetJS() {
  return new Promise((resolve, reject) => {
    if (xlsxLoaded || typeof XLSX !== 'undefined') { xlsxLoaded = true; resolve(); return; }
    const s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    s.onload  = () => { xlsxLoaded = true; resolve(); };
    s.onerror = () => reject(new Error('No se pudo cargar SheetJS'));
    document.head.appendChild(s);
  });
}

/* ══════════════════════════════════════
   RENDER PANTALLA ARCHIVO
   ══════════════════════════════════════ */
function renderArchiveScreen() {
  const txs = window._appGetTxs ? window._appGetTxs() : [];
  const mesesConData = MESES.filter(m => txs.some(t => t.mes === m));

  const screen = document.getElementById('screen-archive');
  if (!screen) return;

  screen.innerHTML = `
    <div class="safe-top"></div>
    <div class="s-header">
      <div class="s-month">Archivar</div>
    </div>
    <div class="archive-body">

      <div class="archive-info-card">
        <div class="arc-info-icon">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polyline points="21 8 21 21 3 21 3 8"/><rect x="1" y="3" width="22" height="5"/><line x1="10" y1="12" x2="14" y2="12"/></svg>
        </div>
        <div class="arc-info-text">
          Exporta los movimientos a un <strong>.xlsx</strong> en tu Drive y los borra de la app. Nada se pierde.
        </div>
      </div>

      <div class="field-group">
        <label class="field-lbl">Desde</label>
        <select id="arc-desde" class="field-input">
          ${mesesConData.map(m => `<option value="${m}">${m}</option>`).join('')}
        </select>
      </div>

      <div class="field-group">
        <label class="field-lbl">Hasta</label>
        <select id="arc-hasta" class="field-input">
          ${mesesConData.map((m, i) => `<option value="${m}" ${i === mesesConData.length - 1 ? 'selected' : ''}>${m}</option>`).join('')}
        </select>
      </div>

      <div class="arc-preview" id="arc-preview"></div>

      <div class="arc-error" id="arc-error"></div>

      <button class="cta-btn" id="arc-btn-archive" style="background:#1a73e8">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:-2px;margin-right:6px"><polyline points="21 8 21 21 3 21 3 8"/><rect x="1" y="3" width="22" height="5"/><line x1="10" y1="12" x2="14" y2="12"/></svg>
        Archivar en Drive
      </button>

      <div class="arc-confirm" id="arc-confirm" style="display:none">
        <p id="arc-confirm-text"></p>
        <div style="display:flex;gap:10px;margin-top:12px">
          <button class="arc-btn-cancel" id="arc-btn-cancel">Cancelar</button>
          <button class="arc-btn-ok" id="arc-btn-ok">Confirmar y archivar</button>
        </div>
      </div>

      <div class="arc-progress" id="arc-progress" style="display:none">
        <div class="arc-progress-bar"><div class="arc-progress-fill" id="arc-progress-fill"></div></div>
        <div class="arc-progress-text" id="arc-progress-text">Generando archivo…</div>
      </div>

    </div>
    <div class="safe-bottom"></div>
  `;

  // Bind range selects → preview
  ['arc-desde','arc-hasta'].forEach(id => {
    document.getElementById(id).addEventListener('change', updatePreview);
  });
  document.getElementById('arc-btn-archive').addEventListener('click', startArchive);
  updatePreview();
}

/* ── Preview: cuántos movimientos en el rango ── */
function updatePreview() {
  const desde = document.getElementById('arc-desde')?.value;
  const hasta = document.getElementById('arc-hasta')?.value;
  const preview = document.getElementById('arc-preview');
  const errEl   = document.getElementById('arc-error');
  if (!desde || !hasta || !preview) return;

  const desdeIdx = MESES.indexOf(desde);
  const hastaIdx = MESES.indexOf(hasta);

  if (hastaIdx < desdeIdx) {
    preview.innerHTML = '';
    errEl.textContent = '"Hasta" debe ser igual o posterior a "Desde".';
    return;
  }
  errEl.textContent = '';

  const rango = MESES.slice(desdeIdx, hastaIdx + 1);
  const txs = window._appGetTxs ? window._appGetTxs() : [];
  const sel = txs.filter(t => rango.includes(t.mes));
  const total = sel.reduce((a, t) => a + (t.tipo === 'gasto' ? t.monto : 0), 0);

  preview.innerHTML = `
    <div class="arc-preview-row">
      <span class="arc-preview-lbl">Meses</span>
      <span class="arc-preview-val">${rango.join(' → ')}</span>
    </div>
    <div class="arc-preview-row">
      <span class="arc-preview-lbl">Movimientos</span>
      <span class="arc-preview-val">${sel.length}</span>
    </div>
    <div class="arc-preview-row">
      <span class="arc-preview-lbl">Total gastos</span>
      <span class="arc-preview-val" style="color:var(--red)">$ ${Math.round(total).toLocaleString('es-AR')}</span>
    </div>
    <div class="arc-preview-row">
      <span class="arc-preview-lbl">Archivo en Drive</span>
      <span class="arc-preview-val" style="color:var(--text3);font-size:12px">${buildFileName(desde, hasta)}</span>
    </div>
  `;
}

function buildFileName(desde, hasta) {
  const now = new Date();
  const d = String(now.getDate()).padStart(2,'0');
  const m = String(now.getMonth()+1).padStart(2,'0');
  const y = now.getFullYear();
  const desdeSlug = desde.toLowerCase().slice(0,3);
  const hastaSlug = hasta.toLowerCase().slice(0,3);
  return `finanzas-${desdeSlug}-${hastaSlug}-${d}${m}${y}.xlsx`;
}

/* ═══════════════════════════════════════
   ARCHIVE FLOW
   ═══════════════════════════════════════ */
function startArchive() {
  const desde = document.getElementById('arc-desde')?.value;
  const hasta = document.getElementById('arc-hasta')?.value;
  const errEl = document.getElementById('arc-error');

  const desdeIdx = MESES.indexOf(desde);
  const hastaIdx = MESES.indexOf(hasta);

  if (hastaIdx < desdeIdx) {
    errEl.textContent = '"Hasta" debe ser igual o posterior a "Desde".';
    return;
  }

  if (!window.driveSync?.isConnected()) {
    errEl.textContent = 'Conectá Google Drive primero (pantalla Inicio).';
    return;
  }

  const rango = MESES.slice(desdeIdx, hastaIdx + 1);
  const txs = window._appGetTxs ? window._appGetTxs() : [];
  const sel = txs.filter(t => rango.includes(t.mes));

  if (sel.length === 0) {
    errEl.textContent = 'No hay movimientos en ese rango.';
    return;
  }
  errEl.textContent = '';

  // Show confirm dialog
  const confirmEl = document.getElementById('arc-confirm');
  const confirmTxt = document.getElementById('arc-confirm-text');
  confirmTxt.textContent =
    `Se van a exportar ${sel.length} movimientos (${rango.join(' a ')}) a Drive y se borrarán de la app. ¿Confirmás?`;
  confirmEl.style.display = 'block';
  document.getElementById('arc-btn-archive').style.display = 'none';

  document.getElementById('arc-btn-cancel').onclick = () => {
    confirmEl.style.display = 'none';
    document.getElementById('arc-btn-archive').style.display = 'block';
  };

  document.getElementById('arc-btn-ok').onclick = () => {
    confirmEl.style.display = 'none';
    runArchive(rango, sel);
  };
}

async function runArchive(rango, selTxs) {
  const progressEl  = document.getElementById('arc-progress');
  const progressFill = document.getElementById('arc-progress-fill');
  const progressTxt  = document.getElementById('arc-progress-text');

  progressEl.style.display = 'block';

  const setProgress = (pct, txt) => {
    progressFill.style.width = pct + '%';
    progressTxt.textContent = txt;
  };

  try {
    setProgress(10, 'Cargando generador de Excel…');
    await loadSheetJS();

    setProgress(30, 'Generando archivo .xlsx…');
    const wb = buildWorkbook(rango, selTxs);
    const xlsxBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    setProgress(60, 'Subiendo a Google Drive…');
    const fileName = buildFileName(rango[0], rango[rango.length - 1]);
    await uploadXlsxToDrive(xlsxBuffer, fileName);

    setProgress(85, 'Borrando movimientos de la app…');
    const selIds = new Set(selTxs.map(t => t.id));
    const remaining = (window._appGetTxs()).filter(t => !selIds.has(t.id));
    if (window._appSetTxs) window._appSetTxs(remaining, window._appGetNextId());

    // Also push updated JSON to Drive
    if (window.driveSync?.push) await window.driveSync.push();

    setProgress(100, `¡Listo! "${fileName}" guardado en Drive.`);
    progressFill.style.background = 'var(--green)';

    if (window._appRefresh) window._appRefresh();

    setTimeout(() => showToast(`Archivado: ${fileName}`, 'success'), 400);

  } catch (e) {
    progressFill.style.background = 'var(--red)';
    setProgress(100, 'Error: ' + (e.message || 'algo salió mal'));
    showToast('Error al archivar', 'error');
  }
}

/* ══════════════════════════════════════
   BUILD WORKBOOK (SheetJS)
   Estructura idéntica al Excel original:
   - Una hoja por mes con datos + resumen lateral
   - Hoja ResumenAnual
   ══════════════════════════════════════ */
function buildWorkbook(rango, txs) {
  const wb = XLSX.utils.book_new();

  rango.forEach(mes => {
    const mesTxs = txs.filter(t => t.mes === mes);
    const ws = buildMonthSheet(mes, mesTxs);
    XLSX.utils.book_append_sheet(wb, ws, mes);
  });

  // Resumen anual
  const wsAnual = buildAnualSheet(rango, txs);
  XLSX.utils.book_append_sheet(wb, wsAnual, 'ResumenAnual');

  return wb;
}

function buildMonthSheet(mes, txs) {
  const gastos = txs.filter(t => t.tipo === 'gasto');

  // Main data table — cols A:E
  const rows = [
    ['Fecha', 'Concepto', 'Categoría', 'Método de pago', 'Monto', '', 'Total del mes', sumBy(gastos)],
  ];

  gastos.forEach(t => rows.push([t.fecha, t.concepto, t.categoria, t.metodo, t.monto]));

  // Category summary — cols G:H (offset col 6)
  const catSummary = [['', '', 'Categoría', 'Total', '', 'Método', 'Total']];
  const maxLen = Math.max(CAT_ORDER.length, MET_ORDER.length);
  for (let i = 0; i < maxLen; i++) {
    const cat = CAT_ORDER[i];
    const met = MET_ORDER[i];
    catSummary.push([
      '', '',
      cat || '', cat ? sumBy(gastos.filter(t => t.categoria === cat)) : '',
      '',
      met || '', met ? sumBy(gastos.filter(t => t.metodo === met)) : '',
    ]);
  }

  // Merge into rows array (extend existing rows)
  catSummary.forEach((summaryRow, i) => {
    if (!rows[i]) rows[i] = new Array(8).fill('');
    summaryRow.forEach((val, j) => {
      if (j >= 2) rows[i][j] = val; // cols G+ are index 6+
    });
  });

  // Include ingresos below gastos if any
  const ingresos = txs.filter(t => t.tipo === 'ingreso');
  if (ingresos.length) {
    rows.push([]);
    rows.push(['--- INGRESOS ---']);
    rows.push(['Fecha', 'Concepto', 'Categoría', 'Método de pago', 'Monto']);
    ingresos.forEach(t => rows.push([t.fecha, t.concepto, t.categoria, t.metodo, t.monto]));
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);

  // Column widths
  ws['!cols'] = [
    { wch: 12 }, // Fecha
    { wch: 28 }, // Concepto
    { wch: 16 }, // Categoría
    { wch: 16 }, // Método
    { wch: 14 }, // Monto
    { wch: 2  }, // spacer
    { wch: 18 }, // Cat label
    { wch: 14 }, // Cat total
  ];

  return ws;
}

function buildAnualSheet(rango, txs) {
  const header = ['Categoría', ...rango, 'Total'];
  const dataRows = CAT_ORDER.map(cat => {
    const vals = rango.map(mes =>
      sumBy(txs.filter(t => t.mes === mes && t.categoria === cat && t.tipo === 'gasto'))
    );
    return [cat, ...vals, vals.reduce((a, v) => a + v, 0)];
  });

  const totalRow = ['TOTAL',
    ...rango.map(mes => sumBy(txs.filter(t => t.mes === mes && t.tipo === 'gasto'))),
    sumBy(txs.filter(t => t.tipo === 'gasto')),
  ];

  const rows = [header, ...dataRows, [], totalRow];
  const ws = XLSX.utils.aoa_to_sheet(rows);

  ws['!cols'] = [{ wch: 18 }, ...rango.map(() => ({ wch: 14 })), { wch: 14 }];
  return ws;
}

function sumBy(arr) {
  return arr.reduce((a, t) => a + (typeof t.monto === 'number' ? t.monto : 0), 0);
}

/* ══════════════════════════════════════
   UPLOAD .xlsx TO DRIVE (binary)
   ══════════════════════════════════════ */
async function uploadXlsxToDrive(arrayBuffer, fileName) {
  const token = window._driveGetToken ? window._driveGetToken() : null;
  if (!token) throw new Error('Sin token de Drive');

  const boundary = 'finanzas_xlsx_boundary';
  const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const meta = JSON.stringify({ name: fileName, mimeType });

  // Build multipart body manually (meta JSON + binary blob)
  const enc = new TextEncoder();
  const metaPart = enc.encode(
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${meta}\r\n` +
    `--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`
  );
  const closing = enc.encode(`\r\n--${boundary}--`);

  const body = new Uint8Array(metaPart.length + arrayBuffer.byteLength + closing.length);
  body.set(metaPart, 0);
  body.set(new Uint8Array(arrayBuffer), metaPart.length);
  body.set(closing, metaPart.length + arrayBuffer.byteLength);

  const res = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': `multipart/related; boundary=${boundary}`,
      },
      body: body,
    }
  );

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Drive upload error ${res.status}: ${err}`);
  }
  return await res.json();
}

/* ── Toast helper (same as drive.js but local fallback) ── */
function showToast(msg, type = 'info') {
  const el = document.getElementById('sync-toast');
  if (!el) return;
  el.textContent = msg;
  el.className = `sync-toast show ${type}`;
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.remove('show'), 4000);
}

/* ── Public init ── */
window.archiveModule = { render: renderArchiveScreen };
