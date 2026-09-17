'use strict';

/* ══════════════════════════════════════
   DATOS INICIALES (importados del Excel)
   ══════════════════════════════════════ */
const SEED_DATA = [
  {id:1,mes:"Enero",fecha:"2026-01-01",concepto:"alquiler",categoria:"Hogar",metodo:"MercadoPago",monto:360000,moneda:"ARS",tipo:"gasto"},
  {id:2,mes:"Enero",fecha:"2026-01-01",concepto:"Expensas",categoria:"Hogar",metodo:"MercadoPago",monto:91748.22,moneda:"ARS",tipo:"gasto"},
  {id:3,mes:"Enero",fecha:"2026-01-01",concepto:"cochera",categoria:"Moto",metodo:"MercadoPago",monto:130000,moneda:"ARS",tipo:"gasto"},
  {id:4,mes:"Enero",fecha:"2026-01-01",concepto:"Seguro Moto",categoria:"Moto",metodo:"amex",monto:141000,moneda:"ARS",tipo:"gasto"},
  {id:5,mes:"Enero",fecha:"2026-01-01",concepto:"Edenor",categoria:"Hogar",metodo:"MercadoPago",monto:76700,moneda:"ARS",tipo:"gasto"},
  {id:6,mes:"Enero",fecha:"2026-01-01",concepto:"Personal",categoria:"Hogar",metodo:"amex",monto:48188.02,moneda:"ARS",tipo:"gasto"},
  {id:7,mes:"Enero",fecha:"2026-01-01",concepto:"Mercadopago6",categoria:"Entretenimiento",metodo:"amex",monto:8990,moneda:"ARS",tipo:"gasto"},
  {id:8,mes:"Enero",fecha:"2026-01-01",concepto:"Sancor Salud",categoria:"Salud",metodo:"amex",monto:122300,moneda:"ARS",tipo:"gasto"},
  {id:9,mes:"Enero",fecha:"2026-01-01",concepto:"Iplan",categoria:"Hogar",metodo:"amex",monto:33600,moneda:"ARS",tipo:"gasto"},
  {id:10,mes:"Enero",fecha:"2026-01-01",concepto:"Aspiradora",categoria:"Hogar",metodo:"amex",monto:27333,moneda:"ARS",tipo:"gasto"},
  {id:11,mes:"Enero",fecha:"2026-01-02",concepto:"Verduleria",categoria:"Comida",metodo:"MercadoPago",monto:10600,moneda:"ARS",tipo:"gasto"},
  {id:12,mes:"Enero",fecha:"2026-01-02",concepto:"Año nuevo comida",categoria:"Salidas",metodo:"MercadoPago",monto:55500,moneda:"ARS",tipo:"gasto"},
  {id:13,mes:"Enero",fecha:"2026-01-02",concepto:"Comida con Negro",categoria:"Salidas",metodo:"amex",monto:17356,moneda:"ARS",tipo:"gasto"},
  {id:14,mes:"Enero",fecha:"2026-01-02",concepto:"Año nuevo salida",categoria:"Salidas",metodo:"amex",monto:107100,moneda:"ARS",tipo:"gasto"},
  {id:15,mes:"Enero",fecha:"2026-01-02",concepto:"candado gym",categoria:"Otros",metodo:"MercadoPago",monto:4000,moneda:"ARS",tipo:"gasto"},
  {id:16,mes:"Enero",fecha:"2026-01-03",concepto:"Chino",categoria:"Comida",metodo:"MercadoPago",monto:8700,moneda:"ARS",tipo:"gasto"},
  {id:17,mes:"Enero",fecha:"2026-01-04",concepto:"Edenor",categoria:"Hogar",metodo:"MercadoPago",monto:76628,moneda:"ARS",tipo:"gasto"},
  {id:18,mes:"Enero",fecha:"2026-01-05",concepto:"bondi",categoria:"Transporte",metodo:"Visa Galicia",monto:1309,moneda:"ARS",tipo:"gasto"},
  {id:19,mes:"Enero",fecha:"2026-01-06",concepto:"Bondi",categoria:"Transporte",metodo:"Visa Galicia",monto:1309,moneda:"ARS",tipo:"gasto"},
  {id:20,mes:"Enero",fecha:"2026-01-06",concepto:"chino",categoria:"Comida",metodo:"MercadoPago",monto:3600,moneda:"ARS",tipo:"gasto"},
  {id:21,mes:"Enero",fecha:"2026-01-06",concepto:"Verduleria",categoria:"Comida",metodo:"MercadoPago",monto:8000,moneda:"ARS",tipo:"gasto"},
  {id:22,mes:"Enero",fecha:"2026-01-06",concepto:"Nafta",categoria:"Moto",metodo:"Efectivo",monto:10000,moneda:"ARS",tipo:"gasto"},
  {id:23,mes:"Enero",fecha:"2026-01-07",concepto:"verduleria",categoria:"Comida",metodo:"MercadoPago",monto:10700,moneda:"ARS",tipo:"gasto"},
  {id:24,mes:"Enero",fecha:"2026-01-07",concepto:"coto",categoria:"Comida",metodo:"amex",monto:46300,moneda:"ARS",tipo:"gasto"},
  {id:25,mes:"Enero",fecha:"2026-01-07",concepto:"farmacia",categoria:"Salud",metodo:"MercadoPago",monto:17900,moneda:"ARS",tipo:"gasto"},
  {id:26,mes:"Enero",fecha:"2026-01-07",concepto:"chino",categoria:"Comida",metodo:"Efectivo",monto:5400,moneda:"ARS",tipo:"gasto"},
  {id:27,mes:"Enero",fecha:"2026-01-08",concepto:"subte",categoria:"Transporte",metodo:"Visa Galicia",monto:2520,moneda:"ARS",tipo:"gasto"},
  {id:28,mes:"Enero",fecha:"2026-01-08",concepto:"BONDI",categoria:"Transporte",metodo:"Visa Galicia",monto:1400,moneda:"ARS",tipo:"gasto"},
  {id:29,mes:"Enero",fecha:"2026-01-08",concepto:"BALANZA",categoria:"Otros",metodo:"amex",monto:15000,moneda:"ARS",tipo:"gasto"},
  {id:30,mes:"Enero",fecha:"2026-01-08",concepto:"pollo",categoria:"Comida",metodo:"MercadoPago",monto:42900,moneda:"ARS",tipo:"gasto"},
  {id:31,mes:"Enero",fecha:"2026-01-12",concepto:"agua chino",categoria:"Comida",metodo:"MercadoPago",monto:3600,moneda:"ARS",tipo:"gasto"},
  {id:32,mes:"Enero",fecha:"2026-01-12",concepto:"espinaca",categoria:"Comida",metodo:"MercadoPago",monto:1800,moneda:"ARS",tipo:"gasto"},
  {id:33,mes:"Enero",fecha:"2026-01-12",concepto:"agua chino",categoria:"Comida",metodo:"MercadoPago",monto:2100,moneda:"ARS",tipo:"gasto"},
  {id:34,mes:"Enero",fecha:"2026-01-12",concepto:"peaje",categoria:"Moto",metodo:"Efectivo",monto:300,moneda:"ARS",tipo:"gasto"},
  {id:35,mes:"Enero",fecha:"2026-01-13",concepto:"nafta",categoria:"Moto",metodo:"Efectivo",monto:10000,moneda:"ARS",tipo:"gasto"},
  {id:36,mes:"Enero",fecha:"2026-01-13",concepto:"big pons",categoria:"Salidas",metodo:"MercadoPago",monto:23000,moneda:"ARS",tipo:"gasto"},
  {id:37,mes:"Enero",fecha:"2026-01-14",concepto:"zapatillas",categoria:"Otros",metodo:"Efectivo",monto:35000,moneda:"ARS",tipo:"gasto"},
  {id:38,mes:"Enero",fecha:"2026-01-14",concepto:"palta",categoria:"Comida",metodo:"Efectivo",monto:5500,moneda:"ARS",tipo:"gasto"},
  {id:39,mes:"Enero",fecha:"2026-01-14",concepto:"agua chino",categoria:"Comida",metodo:"MercadoPago",monto:3600,moneda:"ARS",tipo:"gasto"},
  {id:40,mes:"Enero",fecha:"2026-01-14",concepto:"gaseosa",categoria:"Comida",metodo:"MercadoPago",monto:1600,moneda:"ARS",tipo:"gasto"},
  {id:41,mes:"Enero",fecha:"2026-01-14",concepto:"bondi",categoria:"Transporte",metodo:"MercadoPago",monto:1600,moneda:"ARS",tipo:"gasto"},
  {id:42,mes:"Enero",fecha:"2026-01-15",concepto:"nafta",categoria:"Moto",metodo:"Efectivo",monto:20000,moneda:"ARS",tipo:"gasto"},
  {id:43,mes:"Enero",fecha:"2026-01-16",concepto:"bondi",categoria:"Transporte",metodo:"Visa Galicia",monto:1400,moneda:"ARS",tipo:"gasto"},
  {id:44,mes:"Enero",fecha:"2026-01-16",concepto:"vaso shaker",categoria:"Otros",metodo:"MercadoPago",monto:8000,moneda:"ARS",tipo:"gasto"},
  {id:45,mes:"Enero",fecha:"2026-01-16",concepto:"locker",categoria:"Otros",metodo:"MercadoPago",monto:24000,moneda:"ARS",tipo:"gasto"},
  {id:46,mes:"Enero",fecha:"2026-01-16",concepto:"hamburguesa",categoria:"Salidas",metodo:"MercadoPago",monto:27000,moneda:"ARS",tipo:"gasto"},
  {id:47,mes:"Enero",fecha:"2026-01-16",concepto:"verduleria",categoria:"Comida",metodo:"MercadoPago",monto:3600,moneda:"ARS",tipo:"gasto"},
  {id:48,mes:"Enero",fecha:"2026-01-17",concepto:"chino gaseosa",categoria:"Comida",metodo:"MercadoPago",monto:3000,moneda:"ARS",tipo:"gasto"},
  {id:49,mes:"Enero",fecha:"2026-01-17",concepto:"bondi",categoria:"Transporte",metodo:"Visa Galicia",monto:1400,moneda:"ARS",tipo:"gasto"},
  {id:50,mes:"Enero",fecha:"2026-01-18",concepto:"coto",categoria:"Comida",metodo:"MercadoPago",monto:38500,moneda:"ARS",tipo:"gasto"},
  {id:51,mes:"Enero",fecha:"2026-01-19",concepto:"bondi",categoria:"Transporte",metodo:"Visa Galicia",monto:1400,moneda:"ARS",tipo:"gasto"},
  {id:52,mes:"Enero",fecha:"2026-01-20",concepto:"bondi",categoria:"Transporte",metodo:"Visa Galicia",monto:1400,moneda:"ARS",tipo:"gasto"},
  {id:53,mes:"Enero",fecha:"2026-01-20",concepto:"chicle",categoria:"Otros",metodo:"MercadoPago",monto:1500,moneda:"ARS",tipo:"gasto"},
  {id:54,mes:"Enero",fecha:"2026-01-21",concepto:"Almuerzo ofi",categoria:"Comida",metodo:"MercadoPago",monto:13000,moneda:"ARS",tipo:"gasto"},
  {id:55,mes:"Enero",fecha:"2026-01-21",concepto:"Nafta moto",categoria:"Moto",metodo:"amex",monto:14000,moneda:"ARS",tipo:"gasto"},
  {id:56,mes:"Enero",fecha:"2026-01-22",concepto:"protes",categoria:"Salud",metodo:"MercadoPago",monto:64000,moneda:"ARS",tipo:"gasto"},
  {id:57,mes:"Enero",fecha:"2026-01-22",concepto:"nutricionista",categoria:"Salud",metodo:"MercadoPago",monto:175000,moneda:"ARS",tipo:"gasto"},
  {id:58,mes:"Enero",fecha:"2026-01-24",concepto:"delivery",categoria:"Salidas",metodo:"MercadoPago",monto:25000,moneda:"ARS",tipo:"gasto"},
  {id:59,mes:"Enero",fecha:"2026-01-24",concepto:"nafta",categoria:"Moto",metodo:"amex",monto:11200,moneda:"ARS",tipo:"gasto"},
  {id:60,mes:"Enero",fecha:"2026-01-25",concepto:"verduleria",categoria:"Comida",metodo:"MercadoPago",monto:3000,moneda:"ARS",tipo:"gasto"},
  {id:61,mes:"Enero",fecha:"2026-01-25",concepto:"gym",categoria:"Salidas",metodo:"MercadoPago",monto:12000,moneda:"ARS",tipo:"gasto"},
  {id:62,mes:"Enero",fecha:"2026-01-26",concepto:"bondi",categoria:"Transporte",metodo:"Visa Galicia",monto:1500,moneda:"ARS",tipo:"gasto"},
  {id:63,mes:"Enero",fecha:"2026-01-26",concepto:"verduleria",categoria:"Comida",metodo:"MercadoPago",monto:9000,moneda:"ARS",tipo:"gasto"},
  {id:64,mes:"Enero",fecha:"2026-01-27",concepto:"pelota tenis",categoria:"Otros",metodo:"MercadoPago",monto:6000,moneda:"ARS",tipo:"gasto"},
  {id:65,mes:"Enero",fecha:"2026-01-27",concepto:"mercería",categoria:"Otros",metodo:"MercadoPago",monto:2500,moneda:"ARS",tipo:"gasto"},
  {id:66,mes:"Enero",fecha:"2026-01-27",concepto:"bondi",categoria:"Transporte",metodo:"Visa Galicia",monto:1500,moneda:"ARS",tipo:"gasto"},
  {id:67,mes:"Enero",fecha:"2026-01-27",concepto:"verduleria",categoria:"Comida",metodo:"MercadoPago",monto:3000,moneda:"ARS",tipo:"gasto"},
  {id:68,mes:"Enero",fecha:"2026-01-27",concepto:"crema de maní",categoria:"Comida",metodo:"MercadoPago",monto:5900,moneda:"ARS",tipo:"gasto"},
  {id:69,mes:"Enero",fecha:"2026-01-29",concepto:"verduleria",categoria:"Comida",metodo:"MercadoPago",monto:3300,moneda:"ARS",tipo:"gasto"},
  {id:70,mes:"Enero",fecha:"2026-01-29",concepto:"super",categoria:"Comida",metodo:"MercadoPago",monto:25000,moneda:"ARS",tipo:"gasto"},
  {id:71,mes:"Enero",fecha:"2026-01-29",concepto:"super",categoria:"Otros",metodo:"MercadoPago",monto:25000,moneda:"ARS",tipo:"gasto"},
  {id:72,mes:"Febrero",fecha:"2026-02-28",concepto:"Aspiradora",categoria:"Hogar",metodo:"amex",monto:27333,moneda:"ARS",tipo:"gasto"},
  {id:73,mes:"Marzo",fecha:"2026-03-01",concepto:"alquiler",categoria:"Hogar",metodo:"MercadoPago",monto:550000,moneda:"ARS",tipo:"gasto"},
  {id:74,mes:"Marzo",fecha:"2026-03-01",concepto:"Expensas",categoria:"Hogar",metodo:"MercadoPago",monto:208000,moneda:"ARS",tipo:"gasto"},
  {id:75,mes:"Marzo",fecha:"2026-03-01",concepto:"Seguro Moto",categoria:"Moto",metodo:"amex",monto:145600,moneda:"ARS",tipo:"gasto"},
  {id:76,mes:"Marzo",fecha:"2026-03-01",concepto:"Edenor",categoria:"Hogar",metodo:"MercadoPago",monto:3100,moneda:"ARS",tipo:"gasto"},
  {id:77,mes:"Marzo",fecha:"2026-03-01",concepto:"Personal",categoria:"Hogar",metodo:"amex",monto:28400,moneda:"ARS",tipo:"gasto"},
  {id:78,mes:"Marzo",fecha:"2026-03-29",concepto:"Verduleria",categoria:"Comida",metodo:"MercadoPago",monto:9800,moneda:"ARS",tipo:"gasto"},
  {id:79,mes:"Marzo",fecha:"2026-03-29",concepto:"franks",categoria:"Salidas",metodo:"amex",monto:44600,moneda:"ARS",tipo:"gasto"},
  {id:80,mes:"Marzo",fecha:"2026-03-29",concepto:"didi",categoria:"Salidas",metodo:"amex",monto:3100,moneda:"ARS",tipo:"gasto"},
  {id:81,mes:"Marzo",fecha:"2026-03-30",concepto:"patentes moto",categoria:"Moto",metodo:"Master Galicia",monto:410000,moneda:"ARS",tipo:"gasto"},
  {id:82,mes:"Marzo",fecha:"2026-03-30",concepto:"abl",categoria:"Hogar",metodo:"Master Galicia",monto:36500,moneda:"ARS",tipo:"gasto"},
  {id:83,mes:"Abril",fecha:"2026-04-01",concepto:"alquiler",categoria:"Hogar",metodo:"MercadoPago",monto:550000,moneda:"ARS",tipo:"gasto"},
];

/* ══════════════════════════════════════
   CONFIG
   ══════════════════════════════════════ */
const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

const CAT_CLASS = {
  'Hogar':'cat-hogar','Comida':'cat-comida','Transporte':'cat-trans',
  'Entretenimiento':'cat-entre','Salud':'cat-salud','Moto':'cat-moto',
  'Salidas':'cat-salidas','Otros':'cat-otros'
};
const CAT_ICON = {
  'Hogar':'🏠','Comida':'🥗','Transporte':'🚌','Entretenimiento':'🎮',
  'Salud':'❤️','Moto':'🏍','Salidas':'🍸','Otros':'•••'
};
const CAT_COLOR = {
  'Hogar':'#0a84ff','Comida':'#30d158','Transporte':'#ff9f0a',
  'Entretenimiento':'#bf5af2','Salud':'#ff453a','Moto':'#ff9f0a',
  'Salidas':'#ff2d55','Otros':'#636366'
};
const METODO_PILL = {
  'MercadoPago':'pill-mp','amex':'pill-amx','Visa Galicia':'pill-vg',
  'Master Galicia':'pill-mg','Efectivo':'pill-ef','Transferencia':'pill-tr',
  'Débito':'pill-db','Otros':'pill-ot'
};
const METODO_COLOR = {
  'MercadoPago':'#009ee3','amex':'#4da6ff','Visa Galicia':'#8888ff',
  'Master Galicia':'#ff6680','Efectivo':'#30d158','Transferencia':'#aaaaff',
  'Débito':'#ff9f0a','Otros':'#636366'
};

/* ══════════════════════════════════════
   STATE
   ══════════════════════════════════════ */
let txs = [];
let nextId = 1;
let activeScreen = 'home';
let homeMesIdx = new Date().getMonth(); // 0-based
let repMes = MESES[new Date().getMonth()];
let selectedType = 'gasto';

/* ══════════════════════════════════════
   STORAGE
   ══════════════════════════════════════ */
function loadData() {
  try {
    const saved = localStorage.getItem('finanzas_v2');
    if (saved) {
      const parsed = JSON.parse(saved);
      txs = parsed.txs || SEED_DATA;
      nextId = parsed.nextId || (Math.max(...txs.map(t => t.id)) + 1);
    } else {
      txs = [...SEED_DATA];
      nextId = 84;
      saveData();
    }
  } catch (e) {
    txs = [...SEED_DATA];
    nextId = 84;
  }
}

function saveData() {
  try {
    localStorage.setItem('finanzas_v2', JSON.stringify({ txs, nextId }));
  } catch (e) { /* storage full */ }
}

/* ══════════════════════════════════════
   FORMATTERS
   ══════════════════════════════════════ */
function fmt(n, moneda = 'ARS') {
  const abs = Math.abs(Math.round(n));
  const str = abs.toLocaleString('es-AR');
  return moneda === 'USD' ? `u$s ${str}` : `$ ${str}`;
}
function capitalize(s) {
  if (!s) return s;
  return s.charAt(0).toUpperCase() + s.slice(1);
}

/* ══════════════════════════════════════
   RENDER: HOME
   ══════════════════════════════════════ */
function renderHome() {
  const mes = MESES[homeMesIdx];
  document.getElementById('home-month').textContent = mes;

  // Auto-insert gastos fijos si hay módulo activo
  if (window.recurringModule) {
    const n = window.recurringModule.insertForMes(mes);
    if (n > 0 && window.driveSync?.isConnected()) window.driveSync.push();
  }

  const mesData = txs.filter(t => t.mes === mes && t.moneda === 'ARS');
  const totalGasto   = mesData.filter(t => t.tipo === 'gasto').reduce((a, t) => a + t.monto, 0);
  const totalIngreso = mesData.filter(t => t.tipo === 'ingreso').reduce((a, t) => a + t.monto, 0);
  const balance = totalIngreso - totalGasto;

  document.getElementById('home-total-gasto').textContent = fmt(totalGasto);
  document.getElementById('home-total-ingreso').textContent = fmt(totalIngreso);
  const balEl = document.getElementById('home-balance');
  balEl.textContent = fmt(Math.abs(balance));
  balEl.className = 'bc-mini-val ' + (balance >= 0 ? 'green' : 'red');

  const listEl = document.getElementById('home-tx-list');
  const recent = [...mesData].reverse().slice(0, 10);

  if (!recent.length) {
    listEl.innerHTML = `<div class="empty-state">Sin movimientos en ${mes}.<br>Tocá + para agregar uno.</div>`;
    return;
  }
  listEl.innerHTML = recent.map(t => txItem(t, false)).join('');
}

/* ══════════════════════════════════════
   RENDER: HISTORIAL
   ══════════════════════════════════════ */
function renderHistory() {
  const fMes  = document.getElementById('h-mes').value;
  const fTipo = document.getElementById('h-tipo').value;
  const fCat  = document.getElementById('h-cat').value;

  let filtered = [...txs].reverse();
  if (fMes)  filtered = filtered.filter(t => t.mes === fMes);
  if (fTipo) filtered = filtered.filter(t => t.tipo === fTipo);
  if (fCat)  filtered = filtered.filter(t => t.categoria === fCat);

  const total = filtered.reduce((a, t) => a + (t.tipo === 'gasto' ? -t.monto : t.monto), 0);
  document.getElementById('h-summary').textContent =
    `${filtered.length} movimiento${filtered.length !== 1 ? 's' : ''} · ${fmt(Math.abs(total))}`;

  const listEl = document.getElementById('hist-tx-list');
  if (!filtered.length) {
    listEl.innerHTML = '<div class="empty-state">Sin resultados con ese filtro.</div>';
    return;
  }
  listEl.innerHTML = filtered.map(t => txItem(t, true)).join('');
}

/* ══════════════════════════════════════
   RENDER: REPORTES
   ══════════════════════════════════════ */
function renderReports() {
  // Month tabs
  const mesesConData = MESES.filter(m => txs.some(t => t.mes === m));
  const tabsEl = document.getElementById('rep-month-tabs');
  tabsEl.innerHTML = mesesConData.map(m =>
    `<button class="month-tab${m === repMes ? ' active' : ''}" data-mes="${m}">${m.slice(0,3)}</button>`
  ).join('');
  tabsEl.querySelectorAll('.month-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      repMes = btn.dataset.mes;
      renderReports();
    });
  });

  const mesData = txs.filter(t => t.mes === repMes && t.moneda === 'ARS');
  const totalGasto   = mesData.filter(t => t.tipo === 'gasto').reduce((a, t) => a + t.monto, 0);
  const totalIngreso = mesData.filter(t => t.tipo === 'ingreso').reduce((a, t) => a + t.monto, 0);

  document.getElementById('rep-total-gasto').textContent   = fmt(totalGasto);
  document.getElementById('rep-total-ingreso').textContent = fmt(totalIngreso);

  // Bars por categoría
  const CATS = ['Hogar','Comida','Transporte','Entretenimiento','Salud','Moto','Salidas','Otros'];
  const catTotals = CATS
    .map(c => ({ label: c, val: mesData.filter(t => t.tipo === 'gasto' && t.categoria === c).reduce((a, t) => a + t.monto, 0) }))
    .filter(x => x.val > 0)
    .sort((a, b) => b.val - a.val);
  const maxCat = catTotals[0]?.val || 1;

  document.getElementById('rep-cats').innerHTML = catTotals.length
    ? catTotals.map(({ label, val }) => `
        <div class="bar-row">
          <div class="bar-lbl">${label}</div>
          <div class="bar-out"><div class="bar-in" style="width:${Math.round(val/maxCat*100)}%;background:${CAT_COLOR[label]}"></div></div>
          <div class="bar-val">${fmt(val)}</div>
        </div>`).join('')
    : '<div style="color:var(--text3);font-size:13px;padding:4px 0">Sin gastos en este mes.</div>';

  // Bars por método
  const METODOS = ['MercadoPago','amex','Visa Galicia','Master Galicia','Efectivo','Transferencia','Débito','Otros'];
  const metTotals = METODOS
    .map(m => ({ label: m, val: mesData.filter(t => t.tipo === 'gasto' && t.metodo === m).reduce((a, t) => a + t.monto, 0) }))
    .filter(x => x.val > 0)
    .sort((a, b) => b.val - a.val);
  const maxMet = metTotals[0]?.val || 1;

  document.getElementById('rep-metodos').innerHTML = metTotals.length
    ? metTotals.map(({ label, val }) => `
        <div class="bar-row">
          <div class="bar-lbl">${label}</div>
          <div class="bar-out"><div class="bar-in" style="width:${Math.round(val/maxMet*100)}%;background:${METODO_COLOR[label]}"></div></div>
          <div class="bar-val">${fmt(val)}</div>
        </div>`).join('')
    : '<div style="color:var(--text3);font-size:13px;padding:4px 0">Sin datos.</div>';
}

/* ══════════════════════════════════════
   TX ITEM HTML
   ══════════════════════════════════════ */
function txItem(t, showDelete) {
  const isIngreso = t.tipo === 'ingreso';
  const iconClass = isIngreso ? 'cat-ingreso' : (CAT_CLASS[t.categoria] || 'cat-otros');
  const icon      = isIngreso ? '↑' : (CAT_ICON[t.categoria] || '•');
  const amtColor  = isIngreso ? 'var(--green)' : 'var(--red)';
  const amtSign   = isIngreso ? '+' : '−';
  const pillClass = METODO_PILL[t.metodo] || 'pill-ot';
  const deleteBtn = showDelete
    ? `<button class="tx-delete" data-id="${t.id}" aria-label="Eliminar">✕</button>`
    : '';

  return `
    <div class="tx-item" data-id="${t.id}">
      <div class="tx-icon ${iconClass}">${icon}</div>
      <div class="tx-info">
        <div class="tx-name">${capitalize(t.concepto)}</div>
        <div class="tx-meta">${t.categoria} · ${t.fecha}</div>
      </div>
      <div class="tx-right">
        <div class="tx-amount" style="color:${amtColor}">${amtSign} ${fmt(t.monto, t.moneda)}</div>
        <span class="tx-method ${pillClass}">${t.metodo}</span>
      </div>
      ${deleteBtn}
    </div>`;
}

/* ══════════════════════════════════════
   NAVIGATION
   ══════════════════════════════════════ */
function showScreen(name) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));

  document.getElementById(`screen-${name}`).classList.add('active');
  document.querySelector(`.tab[data-screen="${name}"]`)?.classList.add('active');
  activeScreen = name;

  if (name === 'home')      renderHome();
  if (name === 'history')   renderHistory();
  if (name === 'reports')   renderReports();
  if (name === 'add')       initForm();
  if (name === 'archive') {
    const tryArchive = () => {
      if (window.archiveModule) { window.archiveModule.render(); return; }
      document.getElementById('screen-archive').innerHTML =
        '<div class="safe-top"></div><div style="padding:2rem 1rem;color:#ff453a;font-size:14px;font-family:sans-serif">Error al cargar módulo Archivar.<br>Recargá la app.</div>';
    };
    window.archiveModule ? tryArchive() : setTimeout(tryArchive, 600);
  }
  if (name === 'recurring') {
    const tryRecurring = () => {
      if (window.recurringModule) { window.recurringModule.render(); return; }
      document.getElementById('screen-recurring').innerHTML =
        '<div class="safe-top"></div><div style="padding:2rem 1rem;color:#ff453a;font-size:14px;font-family:sans-serif">Error al cargar módulo Fijos.<br>Recargá la app.</div>';
    };
    window.recurringModule ? tryRecurring() : setTimeout(tryRecurring, 600);
  }
}

/* ══════════════════════════════════════
   ADD FORM
   ══════════════════════════════════════ */
function initForm() {
  document.getElementById('f-fecha').valueAsDate = new Date();
  // Set mes select to match current homeMes
  document.getElementById('f-mes').value = MESES[homeMesIdx];
  updateTypeButtons();
}

function updateTypeButtons() {
  document.querySelectorAll('.type-card').forEach(btn => {
    btn.classList.toggle('selected', btn.dataset.type === selectedType);
  });
}

function addTransaction() {
  const concepto  = document.getElementById('f-concepto').value.trim();
  const monto     = parseFloat(document.getElementById('f-monto').value);
  const moneda    = document.getElementById('f-moneda').value;
  const categoria = document.getElementById('f-categoria').value;
  const metodo    = document.getElementById('f-metodo').value;
  const fecha     = document.getElementById('f-fecha').value;
  const mes       = document.getElementById('f-mes').value;
  const errEl     = document.getElementById('f-error');

  errEl.textContent = '';
  if (!concepto)        { errEl.textContent = 'Ingresá un concepto.'; return; }
  if (!monto || monto <= 0) { errEl.textContent = 'Ingresá un monto válido.'; return; }
  if (!fecha)           { errEl.textContent = 'Seleccioná una fecha.'; return; }

  txs.push({ id: nextId++, mes, fecha, concepto, categoria, metodo, monto, moneda, tipo: selectedType });
  saveData();

  // Sync to Drive if connected
  if (window.driveSync?.isConnected()) window.driveSync.push();

  // Reset form
  document.getElementById('f-concepto').value = '';
  document.getElementById('f-monto').value = '';
  errEl.textContent = '';

  // Go to home showing that month
  homeMesIdx = MESES.indexOf(mes);
  showScreen('home');
}

/* ══════════════════════════════════════
   DELETE
   ══════════════════════════════════════ */
function deleteTransaction(id) {
  txs = txs.filter(t => t.id !== id);
  saveData();
  renderHistory();
  if (window.driveSync?.isConnected()) window.driveSync.push();
}

/* ══════════════════════════════════════
   EVENT LISTENERS
   ══════════════════════════════════════ */
function bindEvents() {
  // Tab navigation
  document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => showScreen(btn.dataset.screen));
  });

  // Month nav on home
  document.getElementById('month-prev').addEventListener('click', () => {
    homeMesIdx = (homeMesIdx - 1 + 12) % 12;
    renderHome();
  });
  document.getElementById('month-next').addEventListener('click', () => {
    homeMesIdx = (homeMesIdx + 1) % 12;
    renderHome();
  });

  // Type selector on add form
  document.querySelectorAll('.type-card').forEach(btn => {
    btn.addEventListener('click', () => {
      selectedType = btn.dataset.type;
      updateTypeButtons();
    });
  });

  // Register button
  document.getElementById('btn-registrar').addEventListener('click', addTransaction);

  // History filters
  ['h-mes','h-tipo','h-cat'].forEach(id => {
    document.getElementById(id).addEventListener('change', renderHistory);
  });

  // Delete buttons (delegated)
  document.getElementById('hist-tx-list').addEventListener('click', e => {
    const btn = e.target.closest('.tx-delete');
    if (btn) deleteTransaction(Number(btn.dataset.id));
  });

  // Clear form errors on input
  ['f-concepto','f-monto'].forEach(id => {
    document.getElementById(id).addEventListener('input', () => {
      document.getElementById('f-error').textContent = '';
    });
  });
}

/* ══════════════════════════════════════
   SERVICE WORKER
   ══════════════════════════════════════ */
function registerSW() {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
}

/* ══════════════════════════════════════
   BRIDGE — expone state a drive.js
   ══════════════════════════════════════ */
window._appGetTxs    = () => txs;
window._appGetNextId = () => nextId;
window._appSetTxs    = (newTxs, newNextId) => {
  txs    = newTxs;
  nextId = newNextId;
  saveData();
};
window._appRefresh   = () => {
  if (activeScreen === 'home')    renderHome();
  if (activeScreen === 'history') renderHistory();
  if (activeScreen === 'reports') renderReports();
};

/* ══════════════════════════════════════
   INIT
   ══════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
  loadData();
  bindEvents();
  window.driveSync?.init?.();
  showScreen('home');
  registerSW();
});

{
/* ══════════════════════════════════════════════════════
   DRIVE.JS — Google Drive sync para Finanzas PWA
   
   Flujo:
   1. Usuario toca "Conectar" → popup OAuth Google
   2. App obtiene access_token (scope: drive.file)
   3. Al guardar un movimiento → sube finanzas-data.json a Drive
   4. Al abrir la app → baja Drive y hace merge por ID
   
   IMPORTANTE: reemplazá GOOGLE_CLIENT_ID con el tuyo.
   Ver README.md sección "Configurar Google OAuth"
   ══════════════════════════════════════════════════════ */

const GOOGLE_CLIENT_ID = 'TU_CLIENT_ID_AQUI.apps.googleusercontent.com';
const DRIVE_SCOPE      = 'https://www.googleapis.com/auth/drive.file';
const DRIVE_FILE_NAME  = 'finanzas-data.json';
const DRIVE_FILE_KEY   = 'finanzas_drive_file_id';
const DRIVE_TOKEN_KEY  = 'finanzas_drive_token';

let driveFileId   = localStorage.getItem(DRIVE_FILE_KEY) || null;
let accessToken   = null;
let tokenClient   = null;
let syncPending   = false;

/* ── UI helpers ── */
function showToast(msg, type = 'info') {
  const el = document.getElementById('sync-toast');
  el.textContent = msg;
  el.className = `sync-toast show ${type}`;
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.remove('show'), 3000);
}

function setDriveStatus(connected, text = '', syncing = false) {
  const banner   = document.getElementById('drive-banner');
  const topRow   = document.getElementById('drive-banner-content');
  const statusRow = document.getElementById('drive-status-row');
  const dot      = document.getElementById('drive-dot');
  const statusTxt = document.getElementById('drive-status-text');

  if (!banner) return;

  if (connected) {
    topRow.style.display    = 'none';
    statusRow.style.display = 'flex';
    dot.className = 'drive-dot ' + (syncing ? 'syncing' : 'ok');
    statusTxt.textContent = text || 'Sincronizado';
  } else {
    topRow.style.display    = 'flex';
    statusRow.style.display = 'none';
  }
}

/* ── Init Google Identity Services ── */
function initDrive() {
  // Wait for GSI script to load
  if (typeof google === 'undefined') {
    setTimeout(initDrive, 500);
    return;
  }

  // Restore token from session (not localStorage — tokens expire)
  const saved = sessionStorage.getItem(DRIVE_TOKEN_KEY);
  if (saved) {
    accessToken = saved;
    setDriveStatus(true, 'Sincronizado');
    // Pull from Drive on app load to catch any changes from other devices
    pullFromDrive();
  }

  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: DRIVE_SCOPE,
    callback: onTokenReceived,
  });

  // Bind buttons
  document.getElementById('drive-connect-btn')?.addEventListener('click', () => {
    tokenClient.requestAccessToken({ prompt: 'consent' });
  });
  document.getElementById('drive-sync-btn')?.addEventListener('click', () => {
    pushToDrive();
  });
  document.getElementById('drive-disconnect-btn')?.addEventListener('click', () => {
    disconnectDrive();
  });
}

function onTokenReceived(response) {
  if (response.error) {
    showToast('No se pudo conectar con Google', 'error');
    return;
  }
  accessToken = response.access_token;
  sessionStorage.setItem(DRIVE_TOKEN_KEY, accessToken);
  setDriveStatus(true, 'Conectando…', true);
  showToast('Conectado a Google Drive', 'success');
  // First sync: pull → merge → push
  pullFromDrive().then(() => pushToDrive());
}

function disconnectDrive() {
  if (accessToken) {
    google.accounts.oauth2.revoke(accessToken, () => {});
  }
  accessToken = null;
  driveFileId = null;
  sessionStorage.removeItem(DRIVE_TOKEN_KEY);
  localStorage.removeItem(DRIVE_FILE_KEY);
  setDriveStatus(false);
  showToast('Desconectado de Google Drive');
}

/* ══════════════════════════════════════
   PUSH — local → Drive
   ══════════════════════════════════════ */
async function pushToDrive() {
  if (!accessToken) return;
  setDriveStatus(true, 'Guardando en Drive…', true);

  const payload = {
    version: 2,
    updatedAt: new Date().toISOString(),
    txs: window._appGetTxs ? window._appGetTxs() : [],
    nextId: window._appGetNextId ? window._appGetNextId() : 1,
  };
  const body = JSON.stringify(payload);

  try {
    if (driveFileId) {
      await updateDriveFile(driveFileId, body);
    } else {
      driveFileId = await createDriveFile(body);
      localStorage.setItem(DRIVE_FILE_KEY, driveFileId);
    }
    setDriveStatus(true, 'Sincronizado · ' + timeNow());
  } catch (e) {
    if (e.status === 401) {
      // Token expired → request new one silently
      tokenClient.requestAccessToken({ prompt: '' });
    } else {
      setDriveStatus(true, 'Error al sincronizar');
      showToast('Error al guardar en Drive', 'error');
    }
  }
}

async function createDriveFile(content) {
  // Multipart upload: metadata + content
  const boundary = 'finanzas_boundary_x9k';
  const meta = JSON.stringify({ name: DRIVE_FILE_NAME, mimeType: 'application/json' });
  const multipart = [
    `--${boundary}`,
    'Content-Type: application/json; charset=UTF-8',
    '',
    meta,
    `--${boundary}`,
    'Content-Type: application/json',
    '',
    content,
    `--${boundary}--`,
  ].join('\r\n');

  const res = await apiFetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
    {
      method: 'POST',
      headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
      body: multipart,
    }
  );
  const data = await res.json();
  return data.id;
}

async function updateDriveFile(fileId, content) {
  await apiFetch(
    `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`,
    {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: content,
    }
  );
}

/* ══════════════════════════════════════
   PULL — Drive → local (merge)
   ══════════════════════════════════════ */
async function pullFromDrive() {
  if (!accessToken) return;
  setDriveStatus(true, 'Bajando de Drive…', true);

  try {
    // Find the file if we don't have its ID
    if (!driveFileId) {
      const searchRes = await apiFetch(
        `https://www.googleapis.com/drive/v3/files?q=name='${DRIVE_FILE_NAME}'+and+trashed=false&fields=files(id,modifiedTime)&orderBy=modifiedTime+desc`
      );
      const searchData = await searchRes.json();
      if (searchData.files && searchData.files.length > 0) {
        driveFileId = searchData.files[0].id;
        localStorage.setItem(DRIVE_FILE_KEY, driveFileId);
      }
    }

    if (!driveFileId) {
      setDriveStatus(true, 'Sin archivo en Drive aún');
      return;
    }

    const fileRes = await apiFetch(
      `https://www.googleapis.com/drive/v3/files/${driveFileId}?alt=media`
    );
    const remote = await fileRes.json();

    if (remote && remote.txs) {
      mergeWithLocal(remote);
      setDriveStatus(true, 'Sincronizado · ' + timeNow());
    }
  } catch (e) {
    setDriveStatus(true, 'Sin conexión — datos locales');
  }
}

/* ══════════════════════════════════════
   MERGE — une local + Drive por ID
   El movimiento con ID mayor gana en
   caso de conflicto de monto/concepto.
   Los de Drive que no están en local
   se agregan, y viceversa.
   ══════════════════════════════════════ */
function mergeWithLocal(remote) {
  if (!window._appGetTxs || !window._appSetTxs) return;

  const local    = window._appGetTxs();
  const localIds = new Set(local.map(t => t.id));
  const remoteIds = new Set(remote.txs.map(t => t.id));

  // Movimientos que están en Drive pero no en local → agregar
  const toAdd = remote.txs.filter(t => !localIds.has(t.id));

  // Movimientos que están en local pero no en Drive → quedan (ya están)
  // (no borramos nada en el merge)

  if (toAdd.length > 0) {
    const merged = [...local, ...toAdd].sort((a, b) => a.id - b.id);
    const newNextId = Math.max(
      window._appGetNextId(),
      remote.nextId || 1,
      ...merged.map(t => t.id + 1)
    );
    window._appSetTxs(merged, newNextId);
    showToast(`+${toAdd.length} movimiento${toAdd.length > 1 ? 's' : ''} sincronizado${toAdd.length > 1 ? 's' : ''} de Drive`, 'success');

    // Re-render current screen
    if (window._appRefresh) window._appRefresh();
  } else {
    showToast('Todo al día', 'info');
  }
}

/* ══════════════════════════════════════
   HTTP helper with auth header
   ══════════════════════════════════════ */
async function apiFetch(url, options = {}) {
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const err = new Error(`Drive API error ${res.status}`);
    err.status = res.status;
    throw err;
  }
  return res;
}

function timeNow() {
  return new Date().toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' });
}

/* ══════════════════════════════════════
   PUBLIC API — llamado desde app.js
   ══════════════════════════════════════ */
window._driveGetToken = () => accessToken;

window.driveSync = {
  init: initDrive,
  push: pushToDrive,
  pull: pullFromDrive,
  isConnected: () => !!accessToken,
};
}

{
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
}

{
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

const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
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
  const mesIdx = MESES.indexOf(mes);
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
      moneda: 'ARS',
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
        ${MESES.map(m => `<option${m === MESES[new Date().getMonth()] ? ' selected' : ''}>${m}</option>`).join('')}
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
            <label class="field-lbl">Monto ARS</label>
            <input type="number" id="rec-f-monto" class="field-input" placeholder="0" min="0" inputmode="decimal"/>
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
  const montoStr = f.monto > 0
    ? `$ ${Math.round(f.monto).toLocaleString('es-AR')}`
    : '<span style="color:var(--text3)">sin monto</span>';

  return `
    <div class="rec-row ${f.activo ? '' : 'rec-row-inactive'}" data-rid="${f.id}">
      <div class="tx-icon ${catClass[f.categoria] || 'cat-otros'}" style="width:36px;height:36px;font-size:14px;flex-shrink:0">${iconMap[f.categoria] || '•'}</div>
      <div class="tx-info">
        <div class="tx-name">${f.concepto}</div>
        <div class="tx-meta">${f.categoria} · ${f.metodo} · día ${f.dia}</div>
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
      if (window.driveSync?.isConnected()) window.driveSync.push();
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
    document.getElementById('rec-f-dia').value      = f.dia;
    document.getElementById('rec-f-cat').value      = f.categoria;
    document.getElementById('rec-f-met').value      = f.metodo;
  } else {
    title.textContent = 'Nuevo gasto fijo';
    document.getElementById('rec-f-id').value      = '';
    document.getElementById('rec-f-concepto').value = '';
    document.getElementById('rec-f-monto').value    = '';
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
    if (f) { f.concepto = concepto; f.monto = monto; f.dia = dia; f.categoria = cat; f.metodo = met; }
  } else {
    const maxId = Math.max(...fijos.map(f => parseInt(f.id.replace('r',''))), 0);
    fijos.push({ id: `r${maxId + 1}`, concepto, categoria: cat, metodo: met, monto, dia, activo: true });
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
}
