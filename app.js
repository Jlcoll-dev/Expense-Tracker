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
  const abs = Math.abs(Number(n) || 0);
  const str = abs.toLocaleString('es-AR', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  return moneda === 'USD' ? `u$s ${str}` : `$ ${str}`;
}
function fmtBalance(n, moneda) {
  return `${n < 0 ? '− ' : ''}${fmt(n, moneda)}`;
}
function capitalize(s) {
  if (!s) return s;
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function calculateBalances(records) {
  const balances = { ARS: 0, USD: 0 };
  records.forEach(t => {
    if (t.tipo === 'cambio') {
      balances[t.moneda] -= Number(t.monto) || 0;
      balances[t.monedaDestino] += Number(t.montoDestino) || 0;
      return;
    }
    const amount = Number(t.monto) || 0;
    balances[t.moneda] += t.tipo === 'ingreso' ? amount : -amount;
  });
  return balances;
}

function monthlyCurrencySummary(records, moneda) {
  return records
    .filter(t => t.moneda === moneda && (t.tipo === 'gasto' || t.tipo === 'ingreso'))
    .reduce((summary, t) => {
      summary[t.tipo] += Number(t.monto) || 0;
      return summary;
    }, { gasto: 0, ingreso: 0 });
}

/* ══════════════════════════════════════
   RENDER: HOME
   ══════════════════════════════════════ */
function renderHome() {
  const mes = MESES[homeMesIdx];
  document.getElementById('home-month').textContent = mes;

  // Auto-insert gastos fijos si hay módulo activo
  if (window.recurringModule) {
    window.recurringModule.insertForMes(mes);
  }

  const mesData = txs.filter(t => t.mes === mes);
  const archived = window.archiveModule?.getArchived?.() || [];
  const balances = calculateBalances([...txs, ...archived]);
  const arsMonth = monthlyCurrencySummary(mesData, 'ARS');
  const usdMonth = monthlyCurrencySummary(mesData, 'USD');

  const arsBalanceEl = document.getElementById('home-ars-balance');
  const usdBalanceEl = document.getElementById('home-usd-balance');
  arsBalanceEl.textContent = fmtBalance(balances.ARS, 'ARS');
  usdBalanceEl.textContent = fmtBalance(balances.USD, 'USD');
  arsBalanceEl.className = 'bc-mini-val ' + (balances.ARS >= 0 ? 'green' : 'red');
  usdBalanceEl.className = 'bc-mini-val ' + (balances.USD >= 0 ? 'green' : 'red');
  document.getElementById('home-ars-detail').textContent = `Este mes: ${fmt(arsMonth.gasto, 'ARS')} gastos`;
  document.getElementById('home-usd-detail').textContent = `Este mes: ${fmt(usdMonth.gasto, 'USD')} gastos`;

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

  const totals = ['ARS', 'USD'].map(moneda => {
    const total = filtered.reduce((sum, t) => {
      if (t.tipo === 'cambio' || t.moneda !== moneda) return sum;
      return sum + (t.tipo === 'gasto' ? -t.monto : t.monto);
    }, 0);
    return total ? fmt(Math.abs(total), moneda) : null;
  }).filter(Boolean);
  document.getElementById('h-summary').textContent =
    `${filtered.length} movimiento${filtered.length !== 1 ? 's' : ''}${totals.length ? ` · ${totals.join(' · ')}` : ''}`;

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

  const mesData = txs.filter(t => t.mes === repMes);
  const arsSummary = monthlyCurrencySummary(mesData, 'ARS');
  const usdSummary = monthlyCurrencySummary(mesData, 'USD');

  document.getElementById('rep-ars-gasto').textContent = fmt(arsSummary.gasto, 'ARS');
  document.getElementById('rep-usd-gasto').textContent = fmt(usdSummary.gasto, 'USD');
  document.getElementById('rep-ars-ingreso').textContent = fmt(arsSummary.ingreso, 'ARS');
  document.getElementById('rep-usd-ingreso').textContent = fmt(usdSummary.ingreso, 'USD');

  // Bars por categoría
  const CATS = ['Hogar','Comida','Transporte','Entretenimiento','Salud','Moto','Salidas','Otros'];
  const catTotals = CATS
    .map(c => ({ label: c, val: mesData.filter(t => t.tipo === 'gasto' && t.moneda === 'ARS' && t.categoria === c).reduce((a, t) => a + t.monto, 0) }))
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
    .map(m => ({ label: m, val: mesData.filter(t => t.tipo === 'gasto' && t.moneda === 'ARS' && t.metodo === m).reduce((a, t) => a + t.monto, 0) }))
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
  if (t.tipo === 'cambio') {
    const deleteBtn = showDelete
      ? `<button class="tx-delete" data-id="${t.id}" aria-label="Eliminar">✕</button>`
      : '';
    return `
      <div class="tx-item" data-id="${t.id}">
        <div class="tx-icon cat-ingreso">↔</div>
        <div class="tx-info">
          <div class="tx-name">${capitalize(t.concepto)}</div>
          <div class="tx-meta">Cambio · ${t.fecha}</div>
        </div>
        <div class="tx-right">
          <div class="tx-amount" style="color:var(--blue,#0a84ff)">− ${fmt(t.monto, t.moneda)}</div>
          <span class="tx-method pill-ot">+ ${fmt(t.montoDestino, t.monedaDestino)}</span>
        </div>
        ${deleteBtn}
      </div>`;
  }
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
  document.getElementById('normal-currency-fields').style.display = selectedType === 'cambio' ? 'none' : '';
  document.getElementById('exchange-fields').style.display = selectedType === 'cambio' ? 'block' : 'none';
}

function addTransaction() {
  let concepto    = document.getElementById('f-concepto').value.trim();
  const monto     = parseFloat(document.getElementById('f-monto').value);
  const moneda    = document.getElementById('f-moneda').value;
  const categoria = document.getElementById('f-categoria').value;
  const metodo    = document.getElementById('f-metodo').value;
  const fecha     = document.getElementById('f-fecha').value;
  const mes       = document.getElementById('f-mes').value;
  const errEl     = document.getElementById('f-error');

  errEl.textContent = '';
  if (!fecha)           { errEl.textContent = 'Seleccioná una fecha.'; return; }

  if (selectedType === 'cambio') {
    const monedaOrigen = document.getElementById('f-cambio-origen').value;
    const monedaDestino = document.getElementById('f-cambio-destino').value;
    const montoOrigen = parseFloat(document.getElementById('f-cambio-monto-origen').value);
    const montoDestino = parseFloat(document.getElementById('f-cambio-monto-destino').value);
    if (monedaOrigen === monedaDestino) { errEl.textContent = 'Elegí dos monedas distintas.'; return; }
    if (!montoOrigen || montoOrigen <= 0 || !montoDestino || montoDestino <= 0) {
      errEl.textContent = 'Ingresá los dos montos del cambio.';
      return;
    }
    concepto = concepto || `Cambio ${monedaOrigen} a ${monedaDestino}`;
    txs.push({ id: nextId++, mes, fecha, concepto, categoria: 'Otros', metodo: 'Efectivo', monto: montoOrigen, moneda: monedaOrigen, tipo: 'cambio', montoDestino, monedaDestino });
  } else {
    if (!concepto) { errEl.textContent = 'Ingresá un concepto.'; return; }
    if (!monto || monto <= 0) { errEl.textContent = 'Ingresá un monto válido.'; return; }
    txs.push({ id: nextId++, mes, fecha, concepto, categoria, metodo, monto, moneda, tipo: selectedType });
  }
  saveData();

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
  document.getElementById('home-add-btn').addEventListener('click', () => showScreen('add'));

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
  BRIDGE — expone el estado a los módulos locales
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
  showScreen('home');
  registerSW();
});
