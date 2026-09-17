'use strict';

/* ══════════════════════════════════════════════════════
  ARCHIVE.JS — Archivar movimientos localmente y exportar a .xlsx

   Flujo:
   1. Usuario abre pantalla Archivar
   2. Elige rango de meses (desde / hasta)
   3. Preview: cuántos movimientos abarca
  4. Confirmar → mueve los registros al archivo local del teléfono
  5. Exporta todos los registros activos y archivados a .xlsx
   ══════════════════════════════════════════════════════ */

const ARCHIVE_MONTHS = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

const CAT_ORDER = ['Hogar','Comida','Transporte','Entretenimiento',
                   'Salud','Moto','Salidas','Otros'];
const MET_ORDER = ['Efectivo','Débito','amex','Transferencia',
                   'MercadoPago','Otros','Visa Galicia','Master Galicia'];
const ARCHIVE_KEY = 'finanzas_archivo_v1';

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
  const mesesConData = ARCHIVE_MONTHS.filter(m => txs.some(t => t.mes === m));

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
          Guarda los movimientos archivados en este teléfono. Podés exportar todo a un <strong>.xlsx</strong> compatible con Google Sheets.
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

      <button class="cta-btn" id="arc-btn-archive">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:-2px;margin-right:6px"><polyline points="21 8 21 21 3 21 3 8"/><rect x="1" y="3" width="22" height="5"/><line x1="10" y1="12" x2="14" y2="12"/></svg>
        Archivar en el teléfono
      </button>

      <button class="cta-btn" id="arc-btn-export-all" style="margin-top:10px;background:var(--green)">
        Exportar todo a Excel
      </button>

      <button class="cta-btn" id="arc-btn-reset-all" style="margin-top:10px;background:var(--red)">
        Borrar todos los registros
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
  document.getElementById('arc-btn-export-all').addEventListener('click', exportAll);
  document.getElementById('arc-btn-reset-all').addEventListener('click', resetAllData);
  updatePreview();
}

function loadArchived() {
  try {
    return JSON.parse(localStorage.getItem(ARCHIVE_KEY) || '[]');
  } catch {
    return [];
  }
}

function saveArchived(records) {
  localStorage.setItem(ARCHIVE_KEY, JSON.stringify(records));
}

function getAllRecords() {
  const active = window._appGetTxs ? window._appGetTxs() : [];
  return [...active, ...loadArchived()];
}

function resetAllData() {
  const total = getAllRecords().length;
  if (!total) {
    showToast('No hay registros para borrar', 'info');
    return;
  }
  const confirmed = window.confirm(
    `Se van a borrar ${total} registros activos y archivados. Esta acción no se puede deshacer. ¿Continuar?`
  );
  if (!confirmed) return;

  if (window._appSetTxs) window._appSetTxs([], 1);
  saveArchived([]);
  if (window._appRefresh) window._appRefresh();
  renderArchiveScreen();
  showToast('Todos los registros fueron borrados', 'success');
}

function downloadWorkbook(records, fileName) {
  const meses = ARCHIVE_MONTHS.filter(m => records.some(t => t.mes === m));
  if (!meses.length) throw new Error('No hay movimientos para exportar.');
  const wb = buildWorkbook(meses, records);
  const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function exportAll() {
  const btn = document.getElementById('arc-btn-export-all');
  const errorEl = document.getElementById('arc-error');
  try {
    errorEl.textContent = '';
    btn.disabled = true;
    await loadSheetJS();
    downloadWorkbook(getAllRecords(), buildFileName('todos', 'los-registros'));
    showToast('Exportación descargada', 'success');
  } catch (e) {
    errorEl.textContent = e.message || 'No se pudo exportar.';
  } finally {
    btn.disabled = false;
  }
}

/* ── Preview: cuántos movimientos en el rango ── */
function updatePreview() {
  const desde = document.getElementById('arc-desde')?.value;
  const hasta = document.getElementById('arc-hasta')?.value;
  const preview = document.getElementById('arc-preview');
  const errEl   = document.getElementById('arc-error');
  if (!desde || !hasta || !preview) return;

  const desdeIdx = ARCHIVE_MONTHS.indexOf(desde);
  const hastaIdx = ARCHIVE_MONTHS.indexOf(hasta);

  if (hastaIdx < desdeIdx) {
    preview.innerHTML = '';
    errEl.textContent = '"Hasta" debe ser igual o posterior a "Desde".';
    return;
  }
  errEl.textContent = '';

  const rango = ARCHIVE_MONTHS.slice(desdeIdx, hastaIdx + 1);
  const txs = window._appGetTxs ? window._appGetTxs() : [];
  const sel = txs.filter(t => rango.includes(t.mes));
  const total = sel.reduce((a, t) => a + (t.tipo === 'gasto' ? t.monto : 0), 0);

  const archivedCount = loadArchived().filter(t => rango.includes(t.mes)).length;
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
      <span class="arc-preview-lbl">Ya archivados en el teléfono</span>
      <span class="arc-preview-val">${archivedCount}</span>
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

  const desdeIdx = ARCHIVE_MONTHS.indexOf(desde);
  const hastaIdx = ARCHIVE_MONTHS.indexOf(hasta);

  if (hastaIdx < desdeIdx) {
    errEl.textContent = '"Hasta" debe ser igual o posterior a "Desde".';
    return;
  }

  const rango = ARCHIVE_MONTHS.slice(desdeIdx, hastaIdx + 1);
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
    `Se van a guardar ${sel.length} movimientos (${rango.join(' a ')}) en el teléfono y dejarán de aparecer en el historial activo. ¿Confirmás?`;
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
    setProgress(20, 'Guardando en el teléfono…');
    const archived = loadArchived();
    const archivedAt = new Date().toISOString();
    saveArchived([...archived, ...selTxs.map(t => ({ ...t, archivedAt }))]);

    setProgress(70, 'Actualizando registros activos…');
    const selIds = new Set(selTxs.map(t => t.id));
    const remaining = (window._appGetTxs()).filter(t => !selIds.has(t.id));
    if (window._appSetTxs) window._appSetTxs(remaining, window._appGetNextId());

    setProgress(100, '¡Listo! Registros guardados en este teléfono.');
    progressFill.style.background = 'var(--green)';

    if (window._appRefresh) window._appRefresh();

    setTimeout(() => showToast(`${selTxs.length} registros archivados`, 'success'), 400);

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

  const cambios = txs.filter(t => t.tipo === 'cambio');
  if (cambios.length) {
    rows.push([]);
    rows.push(['--- CAMBIOS DE MONEDA ---']);
    rows.push(['Fecha', 'Concepto', 'Sale', 'Monto sale', 'Entra', 'Monto entra']);
    cambios.forEach(t => rows.push([
      t.fecha, t.concepto, t.moneda, t.monto, t.monedaDestino, t.montoDestino
    ]));
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

/* ── Toast helper ── */
function showToast(msg, type = 'info') {
  const el = document.getElementById('sync-toast');
  if (!el) return;
  el.textContent = msg;
  el.className = `sync-toast show ${type}`;
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.remove('show'), 4000);
}

/* ── Public init ── */
window.archiveModule = { render: renderArchiveScreen, getArchived: loadArchived };
