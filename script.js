/*
 * script.js — Encabezados exactos:
 * FechaEntrada, Conductor, Procedencia, CantSacos, QQs Netos, Recibidor
 * Filtros: Conductor, Recibidor, FechaEntrada
 * Descargas:
 *   - Origen (Google Sheets -> .xlsx)
 *   - Vista actual (lo que ves en la tabla, con filtros) -> .xlsx con SheetJS
 */

// -------------------- Configuración --------------------
let dataset = [];
let currentView = []; // <- mantiene lo que se está mostrando en la tabla


//https://docs.google.com/spreadsheets/d/16K-i-FwK86CKryE3JFhkvsS2JOszA-Ui/edit?usp=sharing&ouid=116572784923002226502&rtpof=true&sd=true
// Config de Google Sheets (opcional, deja vacío si no usas)
//const GOOGLE_SHEET_ID  = '1HO1dnYe55Weyswxh1qfz1Y8_5nxoGDR6'; // tu ID
const GOOGLE_SHEET_ID  = '1HO1dnYe55Weyswxh1qfz1Y8_5nxoGDR6'; // tu ID
const GOOGLE_SHEET_GID = '0';                                   // pestaña

// -------------------- Utilidades --------------------
function toISODate(val) {
  if (val == null || val === '') return '';
  if (typeof val === 'string' && /^Date\(/.test(val)) {
    const m = val.match(/^Date\((\d+),(\d+),(\d+)(?:,(\d+),(\d+),(\d+))?\)$/);
    if (m) {
      const [, y, mo, d, hh='0', mi='0', ss='0'] = m;
      const dt = new Date(Date.UTC(+y, +mo, +d, +hh, +mi, +ss));
      return dt.toISOString().slice(0, 10);
    }
  }
  if (typeof val === 'string') {
    const m = val.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m) {
      const [, dd, mm, yyyy] = m;
      return `${yyyy}-${String(mm).padStart(2,'0')}-${String(dd).padStart(2,'0')}`;
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(val)) return val;
  }
  if (typeof val === 'number') {
    const ms = Math.round((val - 25569) * 86400 * 1000);
    return new Date(ms).toISOString().slice(0, 10);
  }
  const dt = new Date(val);
  if (!isNaN(dt)) return dt.toISOString().slice(0, 10);
  return '';
}

function formatDateDisplay(val) {
  const iso = toISODate(val);
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

function debounce(fn, wait = 150) {
  let t; 
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), wait); };
}

function prepareDataset() {
  dataset = dataset.map(row => ({
    ...row,
    __fechaISO: toISODate(row['FechaEntrada'])
  }));
}

// -------------------- Carga de datos --------------------
async function loadDataset() {
  // 1) Google Sheets (gviz)
  if (GOOGLE_SHEET_ID) {
    try {
      const gvizUrl = `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/gviz/tq?gid=${GOOGLE_SHEET_GID}&tqx=out:json&t=${Date.now()}`;
      const res = await fetch(gvizUrl);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const text = await res.text();
      const jsonStr = text.substring(text.indexOf('{'), text.lastIndexOf('}') + 1);
      const gvizData = JSON.parse(jsonStr);
      const cols = gvizData.table.cols.map(c => c.label || c.id || '');
      const rows = gvizData.table.rows || [];
      dataset = rows.map(r => {
        const obj = {};
        cols.forEach((label, i) => {
          const cell = r.c?.[i];
          obj[label] = cell ? (cell.f ?? cell.v ?? '') : '';
        });
        return obj;
      });
      prepareDataset();
      initPage();
      return;
    } catch (err) {
      console.warn('Fallo Google Sheets, usando Excel local:', err);
    }
  }

  // 2) Fallback local OBD.xlsx
  try {
    const res = await fetch('OBD.xlsx');
    const ab  = await res.arrayBuffer();
    const wb  = XLSX.read(ab, { type: 'array' });
    const sh  = wb.SheetNames[0];
    dataset   = XLSX.utils.sheet_to_json(wb.Sheets[sh], { defval: '' });
    prepareDataset();
    initPage();
  } catch (err) {
    console.error('Error al cargar el Excel local:', err);
  }
}

// -------------------- Inicialización UI --------------------
function initPage() {
  populateFilters();
  renderTable(dataset); // esto también setea currentView

  const selConductor = document.getElementById('filter-conductor');
  const selRecibidor = document.getElementById('filter-recibidor');
  const inpFecha     = document.getElementById('filter-fecha');
  const btnClear     = document.getElementById('clear-filters');
  const btnRemote    = document.getElementById('download-xlsx-remote');
  const btnView      = document.getElementById('download-xlsx-view');

  const debouncedApply = debounce(applyFilters, 150);

  if (selConductor) selConductor.addEventListener('change', debouncedApply);
  if (selRecibidor) selRecibidor.addEventListener('change', debouncedApply);
  if (inpFecha)     inpFecha.addEventListener('change', debouncedApply);

  if (btnClear) {
    btnClear.addEventListener('click', () => {
      if (selConductor) selConductor.value = '';
      if (selRecibidor) selRecibidor.value = '';
      if (inpFecha)     inpFecha.value = '';
      renderTable(dataset);
    });
  }

  if (btnRemote) {
    btnRemote.addEventListener('click', () => {
      if (!GOOGLE_SHEET_ID) {
        alert('Configura GOOGLE_SHEET_ID para descargar desde Google Sheets.');
        return;
      }
     // const url = `https://drive.google.com/uc?export=download&id=1HO1dnYe55Weyswxh1qfz1Y8_5nxoGDR6`;
       const url = `https://drive.google.com/uc?export=download&id=1HO1dnYe55Weyswxh1qfz1Y8_5nxoGDR6`;
      window.open(url, '_blank');
    });
  }

  if (btnView) {
    btnView.addEventListener('click', exportCurrentViewToExcel);
  }
}

function populateFilters() {
  const selConductor = document.getElementById('filter-conductor');
  const selRecibidor = document.getElementById('filter-recibidor');
  if (!selConductor || !selRecibidor) return;

  const conductores = Array.from(new Set(dataset.map(r => r['Conductor']).filter(Boolean))).sort();
  const recibidores = Array.from(new Set(dataset.map(r => r['Recibidor']).filter(Boolean))).sort();

  conductores.forEach(v => {
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    selConductor.appendChild(opt);
  });

  recibidores.forEach(v => {
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    selRecibidor.appendChild(opt);
  });
}

// -------------------- Renderizado (con data-label para móvil) --------------------
function renderTable(data) {
  const tbody = document.getElementById('table-body');
  if (!tbody) return;
  tbody.innerHTML = '';

  currentView = data.slice(); // <- guarda la vista actual para exportar

  const LABELS = {
    'Fecha': 'FechaEntrada',
    'Nombre del Conductor': 'Conductor',
    'Cliente o Agencia': 'Procedencia',
    'Documentos': 'MTNTs',
    'Sacos': 'CantSacos',
    'QQs Netos': 'QQs Netos',
    'Recibidor': 'Recibidor'
  };

  const BATCH = 200;
  let i = 0;

  function makeTd(label, text) {
    const td = document.createElement('td');
    td.setAttribute('data-label', label);
    td.textContent = text ?? '';
    return td;
  }

  function paintChunk() {
    const frag = document.createDocumentFragment();
    const end = Math.min(i + BATCH, data.length);

    for (; i < end; i++) {
      const row = data[i];
      const tr = document.createElement('tr');

      const tdFecha = makeTd(LABELS['FechaEntrada'], formatDateDisplay(row.__fechaISO || row['FechaEntrada']));
      const tdCond  = makeTd(LABELS['Conductor'],    row['Conductor']);
      const tdProc  = makeTd(LABELS['Procedencia'],  row['Procedencia']);
      const tdDoc  = makeTd(LABELS['Documentos'],  row['MTNTs']);
      const tdSacos = makeTd(LABELS['CantSacos'],    row['CantSacos']);
      const tdQQs   = makeTd(LABELS['QQs Netos'],    row['QQs Netos']);
      const tdRec   = makeTd(LABELS['Recibidor'],    row['Recibidor']);

      tr.append(tdFecha, tdCond, tdProc, tdDoc, tdSacos, tdQQs, tdRec);
      frag.appendChild(tr);
    }

    tbody.appendChild(frag);
    if (i < data.length) requestAnimationFrame(paintChunk);
  }

  requestAnimationFrame(paintChunk);
  updateRowCount(data.length);
}

// -------------------- Filtros --------------------
function applyFilters() {
  const conductor = document.getElementById('filter-conductor')?.value || '';
  const recibidor = document.getElementById('filter-recibidor')?.value || '';
  const fechaISO  = document.getElementById('filter-fecha')?.value || '';

  let filtered = dataset;

  if (conductor) filtered = filtered.filter(r => (r['Conductor'] || '') === conductor);
  if (recibidor) filtered = filtered.filter(r => (r['Recibidor'] || '') === recibidor);
  if (fechaISO)  filtered = filtered.filter(r => r.__fechaISO === fechaISO);

  renderTable(filtered);
}

// -------------------- Contador --------------------
function updateRowCount(count) {
  const el = document.getElementById('row-count');
  if (!el) return;
  el.textContent = `Mostrando ${count} ${count === 1 ? 'registro' : 'registros'}.`;
}

// -------------------- Exportar vista actual a Excel --------------------
function exportCurrentViewToExcel() {
  if (!currentView || currentView.length === 0) {
    alert('No hay datos para exportar.');
    return;
  }

  // Orden y nombres de columnas exactamente como los encabezados
  const header = ['Fecha','Nombre del Conductor','Cliente o Agencia','Documentos','Sacos','QQs Netos','Recibidor'];

  // Construir datos (usamos ISO para fechas; Excel las reconoce o quedan como texto estándar)
  const rows = currentView.map(r => ([
    toISODate(r.__fechaISO || r['FechaEntrada']) || '',
    r['Conductor'] ?? '',
    r['Procedencia'] ?? '',
    r['MTNTs'] ?? '',
    toNumber(r['CantSacos']),
    toNumber(r['QQs Netos']),
    r['Recibidor'] ?? ''
  ]));

  const aoa = [header, ...rows];
  const ws  = XLSX.utils.aoa_to_sheet(aoa);

  // Anchos aproximados
  ws['!cols'] = [
    { wch: 12 }, // FechaEntrada
    { wch: 22 }, // Conductor
    { wch: 22 }, // Procedencia
    { wch: 40 }, // Documentos
    { wch: 10 }, // CantSacos
    { wch: 12 }, // QQs Netos
    { wch: 18 }  // Recibidor
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Vista');

  // Nombre del archivo
  const fechaSel = document.getElementById('filter-fecha')?.value || 'todos';
  const stamp = new Date().toISOString().slice(0,19).replace(/[:T]/g,'-');
  const filename = `destajo_vista_${fechaSel}_${stamp}.xlsx`;

  XLSX.writeFile(wb, filename);
}

function toNumber(val) {
  if (val == null || val === '') return '';
  const n = typeof val === 'number' ? val : parseFloat(String(val).replace(',', '.'));
  return isNaN(n) ? String(val) : n;
}

// -------------------- Arranque --------------------
document.addEventListener('DOMContentLoaded', loadDataset);



