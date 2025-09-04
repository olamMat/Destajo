/*
 * script.js — versión ajustada al archivo con encabezados:
 * FechaEntrada, Conductor, CantSacos, QQs Netos, Recibidor
 *
 * Fuente preferida: Google Sheets (gviz). Fallback: OBD.xlsx local.
 * Filtros: Conductor, Recibidor, FechaEntrada.
 * Botón: descargar Excel desde Google Sheets (export a .xlsx).
 */

// -------------------- Configuración --------------------
let dataset = [];

// Si usas Google Sheets:
const GOOGLE_SHEET_ID  = '1HO1dnYe55Weyswxh1qfz1Y8_5nxoGDR6'; // <-- el tuyo
const GOOGLE_SHEET_GID = '0';                                  // pestaña

// Si en lugar de una hoja usas un archivo Excel en Drive NO–Sheets,
// reemplaza la lógica del botón por un enlace de descarga directa de Drive.
// Ejemplo de export de Google Sheets (ya listo en el botón):
// https://docs.google.com/spreadsheets/d/ID/export?format=xlsx&id=ID&gid=GID

// -------------------- Utilidades --------------------
function toISODate(val) {
  if (val == null || val === '') return '';

  // gviz Date(YYYY,MM,DD[,hh,mm,ss])
  if (typeof val === 'string' && /^Date\(/.test(val)) {
    const m = val.match(/^Date\((\d+),(\d+),(\d+)(?:,(\d+),(\d+),(\d+))?\)$/);
    if (m) {
      const [, y, mo, d, hh='0', mi='0', ss='0'] = m;
      const dt = new Date(Date.UTC(+y, +mo, +d, +hh, +mi, +ss));
      return dt.toISOString().slice(0, 10);
    }
  }

  // "DD/MM/YYYY" o "DD-MM-YYYY"
  if (typeof val === 'string') {
    let m = val.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m) {
      const [, dd, mm, yyyy] = m;
      const d2 = String(dd).padStart(2, '0');
      const m2 = String(mm).padStart(2, '0');
      return `${yyyy}-${m2}-${d2}`;
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(val)) return val; // ISO ya
  }

  // Serial de Excel
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

// Normaliza dataset para filtros rápidos
function prepareDataset() {
  dataset = dataset.map(row => ({
    ...row,
    __fechaISO: toISODate(row['FechaEntrada'])
  }));
}

// -------------------- Carga de datos --------------------
async function loadDataset() {
  // 1) Google Sheets (gviz JSON)
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
          let v = cell ? (cell.f ?? cell.v) : '';
          obj[label] = v ?? '';
        });
        return obj;
      });

      prepareDataset();
      initPage();
      return;
    } catch (err) {
      console.warn('Fallo Google Sheets, usando Excel local:', err);
      // sigue a fallback
    }
  }

  // 2) Fallback: archivo Excel local (OBD.xlsx en la misma carpeta)
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
  renderTable(dataset);

  const selConductor = document.getElementById('filter-conductor');
  const selRecibidor = document.getElementById('filter-recibidor');
  const inpFecha     = document.getElementById('filter-fecha');
  const btnClear     = document.getElementById('clear-filters');
  const btnDownload  = document.getElementById('download-xlsx');

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

  // Botón para descargar el Excel desde Google
  if (btnDownload) {
    btnDownload.addEventListener('click', () => {
      if (!GOOGLE_SHEET_ID) {
        alert('Configura GOOGLE_SHEET_ID para descargar desde Google Sheets.');
        return;
      }
      //https://docs.google.com/spreadsheets/d/1HO1dnYe55Weyswxh1qfz1Y8_5nxoGDR6/edit?usp=drive_link&ouid=116572784923002226502&rtpof=true&sd=true
      // Export directo a .xlsx de la hoja (usa GID actual)
      const url = `https://drive.google.com/uc?export=download&id=1HO1dnYe55Weyswxh1qfz1Y8_5nxoGDR6`;
      // Abre en la misma pestaña o en una nueva:
      window.open(url, '_blank'); // o: location.href = url;
    });
  }
}

function populateFilters() {
  const selConductor = document.getElementById('filter-conductor');
  const selRecibidor = document.getElementById('filter-recibidor');
  if (!selConductor || !selRecibidor) return;

  const conductores = Array.from(
    new Set(dataset.map(r => r['Conductor']).filter(Boolean))
  ).sort();

  const recibidores = Array.from(
    new Set(dataset.map(r => r['Recibidor']).filter(Boolean))
  ).sort();

  conductores.forEach(v => {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    selConductor.appendChild(opt);
  });

  recibidores.forEach(v => {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    selRecibidor.appendChild(opt);
  });
}

// -------------------- Renderizado --------------------
function renderTable(data) {
  const tbody = document.getElementById('table-body');
  if (!tbody) return;

  tbody.innerHTML = '';

  const BATCH = 200;
  let i = 0;

  function makeTd(text) {
    const td = document.createElement('td');
    td.textContent = text ?? '';
    return td;
  }

  function paintChunk() {
    const frag = document.createDocumentFragment();
    const end = Math.min(i + BATCH, data.length);

    for (; i < end; i++) {
      const row = data[i];
      const tr = document.createElement('tr');

      const tdFecha = makeTd(formatDateDisplay(row.__fechaISO || row['FechaEntrada']));
      const tdCond  = makeTd(row['Conductor']);
      const tdSacos = makeTd(row['CantSacos']);
      const tdQQs   = makeTd(row['QQs Netos']);
      const tdRec   = makeTd(row['Recibidor']);

      tr.append(tdFecha, tdCond, tdSacos, tdQQs, tdRec);
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

  if (conductor) {
    filtered = filtered.filter(r => (r['Conductor'] || '') === conductor);
  }
  if (recibidor) {
    filtered = filtered.filter(r => (r['Recibidor'] || '') === recibidor);
  }
  if (fechaISO) {
    filtered = filtered.filter(r => r.__fechaISO === fechaISO);
  }

  renderTable(filtered);
}

function updateRowCount(count) {
  const el = document.getElementById('row-count');
  if (!el) return;
  el.textContent = `Mostrando ${count} ${count === 1 ? 'registro' : 'registros'}.`;
}

// -------------------- Arranque --------------------
document.addEventListener('DOMContentLoaded', loadDataset);
