'use strict';

/* ══════════════════════════════════════════════════════
   RECURRING.JS — Gastos fijos mensuales

   Lógica:
   - El usuario configura sus gastos fijos (concepto,
     monto, categoría, método, día del mes)
   - Al abrir un mes nuevo (o manualmente), la app
     propone insertar los fijos que no están todavía
   - Los fijos se guardan en localStorage separados
     de las transacciones
   ══════════════════════════════════════════════════════ */

const RECURRING_KEY = 'finanzas_fijos_v1';

/* Fijos detectados del Excel de Juan — preconfigurados */
const FIJOS_SEED = [
  { id: 'r1', concepto: 'Alquiler',     categoria: 'Hogar', metodo: 'MercadoPago', monto: 550000, dia: 1,  activo: true  },
  { id: 'r2', concepto: 'Expensas',     categoria: 'Hogar', metodo: 'MercadoPago', monto: 208000, dia: 1,  activo: true  },
  { id: 'r3', concepto: 'Cochera',      categoria: 'Moto',  metodo: 'MercadoPago', monto: 130000, dia: 1,  activo: true  },
  { id: 'r4', concepto: 'Seguro Moto',  categoria: 'Moto',  metodo: 'amex',        monto: 145600, dia: 1,  activo: true  },
  { id: 'r5', concepto: 'Edenor',       categoria: 'Hogar', metodo: 'MercadoPago', monto: 76628,  dia: 1,  activo: true  },
  { id: 'r6', concepto: 'MetroGas',     categoria: 'Hogar', metodo: 'amex',        monto: 0,      dia: 1,  activo: false },
  { id: 'r7', concepto: 'Personal',     categoria: 'Hogar', metodo: 'amex',        monto: 48188,  dia: 1,  activo: true  },
  { id: 'r8', concepto: 'Iplan',        categoria: 'Hogar', metodo: 'amex',        monto: 33600,  dia: 1,  activo: true  },
  { id: 'r9', concepto: 'Sancor Salud', categoria: 'Salud', metodo: 'amex',        monto: 122300, dia: 1,  activo: true  },
];

const RECURRING_MONTHS = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

const CAT_OPTIONS = ['Hogar','Comida','Transporte','Entretenimiento','Salud','Moto','Salidas','Otros'];
const MET_OPTIONS = ['MercadoPago','amex','Visa Galicia','Master Galicia','Efectivo','Transferencia','Débito','Otros'];

/* ══════════════════════════════════════
   STORAGE
   ══════════════════════════════════════ */
function loadFijos() {
  try {
    const s = localStorage.getItem(RECURRING_KEY);
    return s ? JSON.parse(s) : [...FIJOS_SEED];
  } catch { return [...FIJOS_SEED]; }
}

function saveFijos(fijos) {
  localStorage.setItem(RECURRING_KEY, JSON.stringify(fijos));
}

/* ══════════════════════════════════════
   AUTO-INSERT: llama al abrir un mes
   Devuelve cuántos insertó
   ══════════════════════════════════════ */
function insertFijosForMes(mes) {
  const fijos = loadFijos().filter(f => f.activo && f.monto > 0);
  if (!fijos.length || !window._appGetTxs) return 0;

  const txs   = window._appGetTxs();
  const mesIdx = RECURRING_MONTHS.indexOf(mes);
  const year   = new Date().getFullYear();
  let inserted = 0;

  fijos.forEach(fijo => {
    // Check: ¿ya existe un tx de este fijo en este mes?
    const yaExiste = txs.some(t =>
      t.mes === mes &&
      t.tipo === 'gasto' &&
      t.concepto.toLowerCase() === fijo.concepto.toLowerCase()
    );
    if (yaExiste) return;

    // Build date string using fijo.dia
    const dia = String(fijo.dia).padStart(2, '0');
    const monStr = String(mesIdx + 1).padStart(2, '0');
    const fecha = `${year}-${monStr}-${dia}`;

    const newTx = {
      id: window._appGetNextId(),
      mes,
      fecha,
      concepto: fijo.concepto,
      categoria: fijo.categoria,
      metodo: fijo.metodo,
      monto: fijo.monto,
      moneda: fijo.moneda || 'ARS',
      tipo: 'gasto',
      esFijo: true,
    };

    const current = window._appGetTxs();
    window._appSetTxs([...current, newTx], window._appGetNextId() + 1);
    inserted++;
  });

  return inserted;
}

/* ══════════════════════════════════════
   RENDER: PANTALLA GASTOS FIJOS
   ══════════════════════════════════════ */
function renderRecurringScreen() {
  const screen = document.getElementById('screen-recurring');
  if (!screen) return;
  const fijos = loadFijos();
  let nextRId = Math.max(...fijos.map(f => parseInt(f.id.replace('r',''))), 0) + 1;

  screen.innerHTML = `
    <div class="safe-top"></div>
    <div class="s-header">
      <div>
        <div class="s-month">Gastos fijos</div>
        <div class="s-year">Se insertan automáticamente cada mes</div>
      </div>
    </div>

    <div class="rec-insert-bar">
      <div class="rec-insert-text">Insertar fijos activos en un mes</div>
      <select id="rec-mes-insert" class="rec-mes-select">
        ${RECURRING_MONTHS.map(m => `<option${m === RECURRING_MONTHS[new Date().getMonth()] ? ' selected' : ''}>${m}</option>`).join('')}
      </select>
      <button id="rec-btn-insert" class="rec-btn-insert">Insertar</button>
    </div>

    <div class="section-header">Gastos configurados</div>

    <div class="rec-list" id="rec-list">
      ${fijos.map(f => renderFijoRow(f)).join('')}
    </div>

    <div style="padding:0 16px 12px">
      <button class="rec-btn-add" id="rec-btn-add">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
        Agregar gasto fijo
      </button>
    </div>

    <!-- Form nuevo/editar (oculto) -->
    <div id="rec-form-wrap" class="rec-form-wrap" style="display:none">
      <div class="rec-form-header">
        <span id="rec-form-title">Nuevo gasto fijo</span>
        <button id="rec-form-close" class="rec-form-close" aria-label="Cerrar">✕</button>
      </div>
      <div class="rec-form-body">
        <input type="hidden" id="rec-f-id" />
        <div class="field-group">
          <label class="field-lbl">Concepto</label>
          <input type="text" id="rec-f-concepto" class="field-input" placeholder="Ej: Alquiler, Netflix…" autocomplete="off"/>
        </div>
        <div class="field-row">
          <div class="field-group">
              <label class="field-lbl">Monto</label>
            <input type="number" id="rec-f-monto" class="field-input" placeholder="0" min="0" inputmode="decimal"/>
          </div>
            <div class="field-group">
              <label class="field-lbl">Moneda</label>
              <select id="rec-f-moneda" class="field-input">
                <option value="ARS">ARS $</option>
                <option value="USD">USD u$s</option>
              </select>
            </div>
          <div class="field-group">
            <label class="field-lbl">Día del mes</label>
            <input type="number" id="rec-f-dia" class="field-input" value="1" min="1" max="28"/>
          </div>
        </div>
        <div class="field-row">
          <div class="field-group">
            <label class="field-lbl">Categoría</label>
            <select id="rec-f-cat" class="field-input">
              ${CAT_OPTIONS.map(c => `<option>${c}</option>`).join('')}
            </select>
          </div>
          <div class="field-group">
            <label class="field-lbl">Método de pago</label>
            <select id="rec-f-met" class="field-input">
              ${MET_OPTIONS.map(m => `<option>${m}</option>`).join('')}
            </select>
          </div>
        </div>
        <div class="rec-form-err" id="rec-form-err"></div>
        <button class="cta-btn" id="rec-btn-save">Guardar</button>
      </div>
    </div>

    <div class="safe-bottom"></div>
  `;

  bindRecurringEvents();
}

function renderFijoRow(f) {
  const iconMap = {
    'Hogar':'🏠','Comida':'🥗','Transporte':'🚌','Entretenimiento':'🎮',
    'Salud':'❤️','Moto':'🏍','Salidas':'🍸','Otros':'•••'
  };
  const catClass = {
    'Hogar':'cat-hogar','Comida':'cat-comida','Transporte':'cat-trans',
    'Entretenimiento':'cat-entre','Salud':'cat-salud','Moto':'cat-moto',
    'Salidas':'cat-salidas','Otros':'cat-otros'
  };
  const moneda = f.moneda || 'ARS';
  const montoStr = f.monto > 0
    ? `${moneda === 'USD' ? 'u$s' : '$'} ${Number(f.monto).toLocaleString('es-AR', { maximumFractionDigits: 2 })}`
    : '<span style="color:var(--text3)">sin monto</span>';

  return `
    <div class="rec-row ${f.activo ? '' : 'rec-row-inactive'}" data-rid="${f.id}">
      <div class="tx-icon ${catClass[f.categoria] || 'cat-otros'}" style="width:36px;height:36px;font-size:14px;flex-shrink:0">${iconMap[f.categoria] || '•'}</div>
      <div class="tx-info">
        <div class="tx-name">${f.concepto}</div>
          <div class="tx-meta">${f.categoria} · ${f.metodo} · ${moneda} · día ${f.dia}</div>
      </div>
      <div style="text-align:right;flex-shrink:0">
        <div class="tx-amount" style="color:var(--red);margin-bottom:5px">${montoStr}</div>
        <div style="display:flex;gap:8px;justify-content:flex-end;align-items:center">
          <label class="rec-toggle" title="${f.activo ? 'Desactivar' : 'Activar'}">
            <input type="checkbox" class="rec-check" data-rid="${f.id}" ${f.activo ? 'checked' : ''} />
            <span class="rec-slider"></span>
          </label>
          <button class="rec-edit-btn" data-rid="${f.id}" aria-label="Editar">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
          </button>
          <button class="rec-del-btn" data-rid="${f.id}" aria-label="Eliminar">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/></svg>
          </button>
        </div>
      </div>
    </div>
  `;
}

/* ══════════════════════════════════════
   EVENTS
   ══════════════════════════════════════ */
function bindRecurringEvents() {
  // Toggle activo/inactivo
  document.querySelectorAll('.rec-check').forEach(chk => {
    chk.addEventListener('change', () => {
      const fijos = loadFijos();
      const f = fijos.find(x => x.id === chk.dataset.rid);
      if (f) { f.activo = chk.checked; saveFijos(fijos); }
      const row = document.querySelector(`.rec-row[data-rid="${chk.dataset.rid}"]`);
      if (row) row.classList.toggle('rec-row-inactive', !chk.checked);
    });
  });

  // Editar
  document.querySelectorAll('.rec-edit-btn').forEach(btn => {
    btn.addEventListener('click', () => openForm(btn.dataset.rid));
  });

  // Eliminar
  document.querySelectorAll('.rec-del-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const fijos = loadFijos().filter(f => f.id !== btn.dataset.rid);
      saveFijos(fijos);
      renderRecurringScreen();
    });
  });

  // Agregar nuevo
  document.getElementById('rec-btn-add').addEventListener('click', () => openForm(null));

  // Cerrar form
  document.getElementById('rec-form-close').addEventListener('click', closeForm);

  // Guardar form
  document.getElementById('rec-btn-save').addEventListener('click', saveForm);

  // Insertar en mes
  document.getElementById('rec-btn-insert').addEventListener('click', () => {
    const mes = document.getElementById('rec-mes-insert').value;
    const n = insertFijosForMes(mes);
    if (n > 0) {
      showRecToast(`${n} gasto${n > 1 ? 's' : ''} fijo${n > 1 ? 's' : ''} insertado${n > 1 ? 's' : ''} en ${mes}`, 'success');
      if (window._appRefresh) window._appRefresh();
    } else {
      showRecToast(`Los fijos de ${mes} ya estaban registrados`, 'info');
    }
  });
}

function openForm(rid) {
  const wrap = document.getElementById('rec-form-wrap');
  const title = document.getElementById('rec-form-title');
  wrap.style.display = 'block';

  if (rid) {
    const f = loadFijos().find(x => x.id === rid);
    if (!f) return;
    title.textContent = 'Editar gasto fijo';
    document.getElementById('rec-f-id').value      = f.id;
    document.getElementById('rec-f-concepto').value = f.concepto;
    document.getElementById('rec-f-monto').value    = f.monto;
    document.getElementById('rec-f-moneda').value   = f.moneda || 'ARS';
    document.getElementById('rec-f-dia').value      = f.dia;
    document.getElementById('rec-f-cat').value      = f.categoria;
    document.getElementById('rec-f-met').value      = f.metodo;
  } else {
    title.textContent = 'Nuevo gasto fijo';
    document.getElementById('rec-f-id').value      = '';
    document.getElementById('rec-f-concepto').value = '';
    document.getElementById('rec-f-monto').value    = '';
    document.getElementById('rec-f-moneda').value   = 'ARS';
    document.getElementById('rec-f-dia').value      = '1';
    document.getElementById('rec-f-cat').value      = 'Hogar';
    document.getElementById('rec-f-met').value      = 'MercadoPago';
  }
  document.getElementById('rec-form-err').textContent = '';
  wrap.scrollIntoView({ behavior: 'smooth' });
}

function closeForm() {
  document.getElementById('rec-form-wrap').style.display = 'none';
}

function saveForm() {
  const concepto = document.getElementById('rec-f-concepto').value.trim();
  const monto    = parseFloat(document.getElementById('rec-f-monto').value) || 0;
  const moneda   = document.getElementById('rec-f-moneda').value;
  const dia      = parseInt(document.getElementById('rec-f-dia').value) || 1;
  const cat      = document.getElementById('rec-f-cat').value;
  const met      = document.getElementById('rec-f-met').value;
  const rid      = document.getElementById('rec-f-id').value;
  const errEl    = document.getElementById('rec-form-err');

  if (!concepto) { errEl.textContent = 'Ingresá un concepto.'; return; }
  if (dia < 1 || dia > 28) { errEl.textContent = 'El día debe ser entre 1 y 28.'; return; }
  errEl.textContent = '';

  const fijos = loadFijos();
  if (rid) {
    const f = fijos.find(x => x.id === rid);
    if (f) { f.concepto = concepto; f.monto = monto; f.moneda = moneda; f.dia = dia; f.categoria = cat; f.metodo = met; }
  } else {
    const maxId = Math.max(...fijos.map(f => parseInt(f.id.replace('r',''))), 0);
    fijos.push({ id: `r${maxId + 1}`, concepto, categoria: cat, metodo: met, monto, moneda, dia, activo: true });
  }
  saveFijos(fijos);
  closeForm();
  renderRecurringScreen();
}

function showRecToast(msg, type = 'info') {
  const el = document.getElementById('sync-toast');
  if (!el) return;
  el.textContent = msg;
  el.className = `sync-toast show ${type}`;
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.remove('show'), 3500);
}

/* ── Public API ── */
window.recurringModule = {
  render: renderRecurringScreen,
  insertForMes: insertFijosForMes,
};
