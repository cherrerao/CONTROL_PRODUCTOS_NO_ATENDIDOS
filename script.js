// ═══════════════════════════════════════════════════════════════
// CONFIG
// ═══════════════════════════════════════════════════════════════
const SHEET_ID      = '1wXQjHUAHEnfTde4xWJujv9xMQOmbGgzaI_27rRnUOQM';
const SHEET_TAB     = 'ENTRADAS';

// ⚠️ REEMPLAZA CON EL ID DEL SHEET DE DISPO_CENTROS_ALM
const ICI_SHEET_ID  = '1Tqpua8gTpw_jstO2YOHkQ_bwtXmo2iWJ';
const ICI_SHEET_TAB = 'DISP_CENTROS_ALM_SINC';

// ═══════════════════════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════════════════════
Chart.register(ChartDataLabels);
let allData      = [];
let filteredData = [];
let iciData      = [];          // datos de disponibilidad ICI
let cruceData    = [];          // resultado del cruce
let groupedData  = [];
let currentPeriod= 'todos';
let charts       = {};
let sortCol = 'noSat', sortDir = -1;
let page = 1, iciPage = 1, groupPage = 1;
const PAGE = 50;
let iciSortCol = 'indicador', iciSortDir = 1;
let iciFilterIndicador = '', iciFilterRed = '', iciSearch = '';
let dateGroupedData = [];
let dateGroupMode = 'last'; // 'none' | 'all' | 'last' | 'today'
// ═══════════════════════════════════════════════════════════════
// STATE EXTRA — ICI por día
// ═══════════════════════════════════════════════════════════════
let iciDayFilter = '';   // fecha YYYY-MM-DD seleccionada en ICI
let iciCentroFilter = '';
let iciGroupedData = []; // cruce agrupado por fecha+centro+producto

// ═══════════════════════════════════════════════════════════════
// EXCEL EXPORT — usa SheetJS (xlsx.full.min.js)
// ═══════════════════════════════════════════════════════════════
// ─── Paleta DIRESA ───────────────────────────────────────────
const XLSX_PALETTE = {
  headerBg:   '0D2137',   // azul oscuro institucional
  headerFg:   'FFFFFF',
  subBg:      '1565C0',   // azul medio
  subFg:      'FFFFFF',
  altRow:     'EEF4FF',   // azul muy claro para filas pares
  whiteFg:    'FFFFFF',
  textDark:   '0D2137',
  red:        'C62828',
  orange:     'E65100',
  green:      '2E7D32',
  muted:      '5A7490',
  border:     'BDD4F0',
};

function xlsxCell(v, bold, bg, fg, align, border) {
  const s = { font: { bold: !!bold, color: { rgb: fg||XLSX_PALETTE.textDark }, name: 'Calibri', sz: 10 } };
  if (bg) s.fill = { patternType: 'solid', fgColor: { rgb: bg } };
  if (align) s.alignment = { horizontal: align, vertical: 'center', wrapText: false };
  if (border) {
    const b = { style: 'thin', color: { rgb: XLSX_PALETTE.border } };
    s.border = { top: b, bottom: b, left: b, right: b };
  }
  return { v, s };
}

function exportXLSX(headers, rows, filename, opts) {
  opts = opts || {};
  const title    = opts.title    || filename.replace(/_/g,' ').replace(/\.xlsx$/,'');
  const subtitle = opts.subtitle || ('DIRESA Callao · DEMID · Generado: ' + new Date().toLocaleDateString('es-PE', { day:'2-digit', month:'long', year:'numeric' }));
  const sheetName = opts.sheet  || 'Datos';

  const aoa = [];

  // Row 0: main title (merged)
  aoa.push([{ v: title, s: {
    font: { bold: true, sz: 13, color: { rgb: XLSX_PALETTE.headerFg }, name: 'Calibri' },
    fill: { patternType: 'solid', fgColor: { rgb: XLSX_PALETTE.headerBg } },
    alignment: { horizontal: 'left', vertical: 'center' }
  } }, ...Array(headers.length - 1).fill('')]);

  // Row 1: subtitle
  aoa.push([{ v: subtitle, s: {
    font: { italic: true, sz: 9, color: { rgb: XLSX_PALETTE.headerFg }, name: 'Calibri' },
    fill: { patternType: 'solid', fgColor: { rgb: XLSX_PALETTE.subBg } },
    alignment: { horizontal: 'left', vertical: 'center' }
  } }, ...Array(headers.length - 1).fill('')]);

  // Row 2: blank spacer
  aoa.push(Array(headers.length).fill(''));

  // Row 3: headers
  aoa.push(headers.map(h => xlsxCell(h, true, XLSX_PALETTE.headerBg, XLSX_PALETTE.headerFg, 'center', true)));

  // Rows 4+: data with alternating colors + conditional number formatting
  rows.forEach((row, ri) => {
    const bg = ri % 2 === 0 ? null : XLSX_PALETTE.altRow;
    aoa.push(row.map((v, ci) => {
      // Detect cobertura column (header contains %)
      const hdr = headers[ci] || '';
      let cell = xlsxCell(v, false, bg, null, null, true);
      // Numbers: right-align
      if (typeof v === 'number') {
        cell.s.alignment = { horizontal: 'right', vertical: 'center' };
        // Cobertura % coloring
        if (hdr.includes('%') || hdr.toLowerCase().includes('cobertura')) {
          const pct = parseFloat(v);
          cell.s.font.color = { rgb: pct === 0 ? XLSX_PALETTE.red : pct < 30 ? XLSX_PALETTE.orange : XLSX_PALETTE.green };
          cell.s.font.bold  = true;
        }
        // Sin atender / noSat coloring
        if (hdr.toLowerCase().includes('sin atender') || hdr.toLowerCase().includes('no atend')) {
          if (v > 0) { cell.s.font.color = { rgb: XLSX_PALETTE.red }; cell.s.font.bold = true; }
        }
      }
      return cell;
    }));
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Merge title & subtitle rows across all columns
  ws['!merges'] = [
    { s: { r:0, c:0 }, e: { r:0, c: headers.length - 1 } },
    { s: { r:1, c:0 }, e: { r:1, c: headers.length - 1 } },
    { s: { r:2, c:0 }, e: { r:2, c: headers.length - 1 } },
  ];

  // Row heights
  ws['!rows'] = [{ hpt: 22 }, { hpt: 14 }, { hpt: 6 }, { hpt: 18 },
    ...rows.map(() => ({ hpt: 16 }))];

  // Column widths
  ws['!cols'] = headers.map((h, i) => {
    const maxLen = Math.max(h.length, ...rows.map(r => String(r[i]||'').length));
    return { wch: Math.min(maxLen + 4, 52) };
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  // Add a metadata sheet
  const metaWs = XLSX.utils.aoa_to_sheet([
    [{ v: 'Reporte', s: { font: { bold: true } } }, title],
    [{ v: 'Sistema', s: { font: { bold: true } } }, 'DEMID · DIRESA Callao'],
    [{ v: 'Generado', s: { font: { bold: true } } }, new Date().toLocaleString('es-PE')],
    [{ v: 'Registros', s: { font: { bold: true } } }, rows.length],
  ]);
  metaWs['!cols'] = [{ wch: 14 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(wb, metaWs, 'Info');

  XLSX.writeFile(wb, filename);
}

function fmtFechaXLSX(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toLocaleDateString('es-PE');
  return String(v);
}

// ICI: mes clave "YYYY-MM" → nombre
function mesLabel(k) {
  if (!k || k === 'sin-fecha') return k;
  const [y, m] = k.split('-');
  const d = new Date(+y, +m - 1, 1);
  return d.toLocaleDateString('es-PE', { month: 'long', year: 'numeric' });
}

// ═══════════════════════════════════════════════════════════════
// ALERTA DE DESABASTECIMIENTO PERSISTENTE
// Productos que llevan ≥2 días consecutivos con indicador_ce DESABASTECIDO o SUBSTOCK
// ═══════════════════════════════════════════════════════════════
function detectarPersistentes(cruceArr) {
  // Agrupar por cod_pre + cod_sismed + fecha
  const byKey = {};
  cruceArr.forEach(r => {
    if (!r.tiene_match) return;
    //const fecha = r.fecha ? r.fecha.toISOString().slice(0,10) : null;
    const fecha = r.fecha
    ? (
        r.fecha.getFullYear() + '-' +
        String(r.fecha.getMonth()+1).padStart(2,'0') + '-' +
        String(r.fecha.getDate()).padStart(2,'0')
      )
    : null;
    if (!fecha) return;
    const k = `${r.cod_pre}||${r.cod_sismed}`;
    if (!byKey[k]) byKey[k] = { estab: r.estab, redes: r.redes, producto: r.producto, cod_pre: r.cod_pre, cod_sismed: r.cod_sismed, dias: {} };
    if (!byKey[k].dias[fecha]) byKey[k].dias[fecha] = { ind_ce: r.indicador_ce, noSat: 0 };
    byKey[k].dias[fecha].noSat += r.noSat;
    // keep worst indicator
    const order = ['DESABASTECIDO','SUBSTOCK','NORMOSTOCK','SOBRESTOCK','SIN_MATCH'];
    if (order.indexOf(r.indicador_ce) < order.indexOf(byKey[k].dias[fecha].ind_ce || 'SIN_MATCH')) {
      byKey[k].dias[fecha].ind_ce = r.indicador_ce;
    }
  });
  const persistentes = [];
  for (const [k, info] of Object.entries(byKey)) {
    const fechas = Object.keys(info.dias).sort();
    if (fechas.length < 2) continue;
    // check last 2+ consecutive dates with DESABASTECIDO or SUBSTOCK
    // Se cambia para que puede indicar los dias que estan corriendo desde que esta desabastecido
    /*const malas = fechas.filter(f => ['DESABASTECIDO','SUBSTOCK'].includes(info.dias[f].ind_ce));
    if (malas.length >= 2) {
      const totalNoSat = fechas.reduce((s,f) => s + info.dias[f].noSat, 0);
      const ultimaFecha = fechas[fechas.length-1];
      const primMala = malas[0];
      persistentes.push({
        estab: info.estab,
        redes: info.redes,
        producto: info.producto,
        cod_pre: info.cod_pre,
        cod_sismed: info.cod_sismed,
        diasMalos: malas.length,
        diasTotal: fechas.length,
        totalNoSat,
        ultimaFecha,
        primerFechaMala: primMala,
        indicadores: fechas.map(f => info.dias[f].ind_ce),
      });
    }*/
    let consecutivos = 0;
    let maxConsecutivos = 0;
    let primerFechaConsecutiva = null;

    for (let i = 0; i < fechas.length; i++) {

      const actual = fechas[i];

      const esMalo =
        ['DESABASTECIDO']
        .includes(info.dias[actual].ind_ce);

      if (!esMalo) {
        consecutivos = 0;
        continue;
      }

      if (i === 0) {

        consecutivos = 1;
        primerFechaConsecutiva = actual;

      } else {

        const prev = new Date(fechas[i - 1]);
        const curr = new Date(actual);

        const diff =
          (curr - prev) / (1000 * 60 * 60 * 24);

        if (diff === 1) {

          consecutivos++;

        } else {

          consecutivos = 1;
          primerFechaConsecutiva = actual;
        }
      }

      if (consecutivos > maxConsecutivos) {
        maxConsecutivos = consecutivos;
      }
    }

    if (maxConsecutivos >= 2) {

      const totalNoSat =
        fechas.reduce((s,f) =>
          s + info.dias[f].noSat, 0);

      const ultimaFecha = fechas[fechas.length - 1];

      persistentes.push({
        estab: info.estab,
        redes: info.redes,
        producto: info.producto,
        cod_pre: info.cod_pre,
        cod_sismed: info.cod_sismed,
        //diasMalos: maxConsecutivos,
        diasMalos: Math.floor(
  (
        new Date(ultimaFecha) -
        new Date(
          fechas.find(f =>
          ['DESABASTECIDO']
        .includes(info.dias[f].ind_ce)
      )
    )
  ) / (1000 * 60 * 60 * 24)
) + 1,
        diasTotal: fechas.length,
        totalNoSat,
        ultimaFecha,
        //primerFechaMala: primerFechaConsecutiva,
        primerFechaMala: fechas.find(f =>
        ['DESABASTECIDO']
        .includes(info.dias[f].ind_ce)
        ),
        indicadores: fechas.map(f => info.dias[f].ind_ce),
      });
    }

  }
  return persistentes.sort((a,b) => b.diasMalos - a.diasMalos || b.totalNoSat - a.totalNoSat);
}

// ═══════════════════════════════════════════════════════════════
// ICI AGRUPADO POR DÍA + CENTRO + PRODUCTO
// ═══════════════════════════════════════════════════════════════
function buildICIGrouped(cruceArr) {
  const map = {};
  cruceArr.forEach(r => {
    const fecha = r.fecha ? r.fecha.toISOString().slice(0,10) : 'sin-fecha';
    const k = `${fecha}||${r.cod_pre}||${r.cod_sismed}`;
    if (!map[k]) {
      map[k] = {
        fecha,
        fechaDate: r.fecha,
        cod_pre: r.cod_pre,
        cod_sismed: r.cod_sismed,
        estab: r.estab,
        redes: r.redes,
        producto: r.producto,
        servicio: r.servicio,
        noSat: 0,
        requerida: 0,
        disponible: 0,
        count: 0,
        indicador_ce: r.indicador_ce,
        indicador_alm: r.indicador_alm,
        stock_ce: r.stock_ce,
        stock_alm: r.stock_alm,
        tiene_match: r.tiene_match,
        match_nombre: r.match_nombre,
        mesKey: r.fecha ? (r.fecha.getFullYear() + '-' + String(r.fecha.getMonth()+1).padStart(2,'0')) : 'sin-fecha',
      };
    }
    const it = map[k];
    it.noSat += r.noSat;
    it.requerida += (r.requerida || 0);
    it.disponible += (r.disponible || 0);
    it.count += 1;
    // keep worst indicator
    const ord = ['DESABASTECIDO','SUBSTOCK','NORMOSTOCK','SOBRESTOCK','SIN_MATCH'];
    if (ord.indexOf(r.indicador_ce) < ord.indexOf(it.indicador_ce || 'SIN_MATCH')) it.indicador_ce = r.indicador_ce;
  });
  return Object.values(map)
    .map(it => ({ ...it, cobertura: it.requerida ? (it.disponible / it.requerida * 100) : 0 }))
    .sort((a,b) => b.fecha.localeCompare(a.fecha) || b.noSat - a.noSat); // más reciente primero
}


// ═══════════════════════════════════════════════════════════════
// FETCH CSV con fallback CORS (genérico)
// ═══════════════════════════════════════════════════════════════
async function fetchCSV(sheetId, sheetTab) {
  const directUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&sheet=${encodeURIComponent(sheetTab)}`;
  const gvizUrl   = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetTab)}`;
  try {
    const res = await fetch(directUrl, { redirect: 'follow' });
    if (res.ok) { const csv = await res.text(); if (csv && csv.length > 50) return csv; }
  } catch(e) {}
  const proxies = [
    `https://api.allorigins.win/raw?url=${encodeURIComponent(gvizUrl)}`,
    `https://corsproxy.io/?${encodeURIComponent(gvizUrl)}`,
    `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(gvizUrl)}`,
  ];
  for (const proxy of proxies) {
    try {
      const res = await fetch(proxy);
      if (res.ok) { const csv = await res.text(); if (csv && csv.length > 50 && !csv.includes('<html')) return csv; }
    } catch(e) { continue; }
  }
  throw new Error('No se pudo conectar. Verifica que el Sheet sea público.');
}

async function fetchSheetData() { return parseCSVEntradas(await fetchCSV(SHEET_ID, SHEET_TAB)); }
async function fetchICIData()   { return parseCSVICI(await fetchCSV(ICI_SHEET_ID, ICI_SHEET_TAB)); }

let RED_MAP = {};

function normalizeCodigo(value, length = 5) {
  const digits = String(value || '').replace(/[^0-9]/g, '');
  return digits ? digits.padStart(length, '0') : '';
}

async function cargarCatalogo() {
  const res  = await fetch("catalogo_establecimientos.json");
  const json = await res.json();
  json.redes.forEach(red => {
    red.establecimientos.forEach(est => { RED_MAP[est.cod_pre] = red.nombre; });
  });
}

// ═══════════════════════════════════════════════════════════════
// PARSE CSV — Hoja ENTRADAS (Productos No Atendidos)
// ═══════════════════════════════════════════════════════════════
function parseCSVEntradas(csv) {
  const lines = csv.split('\n').filter(l => l.trim());
  if (lines.length < 2) throw new Error('La hoja ENTRADAS está vacía.');
  function parseLine(line) {
    const result = []; let current = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { inQ = !inQ; continue; }
      if (ch === ',' && !inQ) { result.push(current.trim()); current = ''; continue; }
      current += ch;
    }
    result.push(current.trim()); return result;
  }
  const h   = parseLine(lines[0]).map(h => h.replace(/^"|"$/g,'').trim());
  const idx = {};
  const COL_MAP = {
    REDES:     ['Redes','redes'],
    cod_pre:   ['COD PRE','cod_pre','cod pre','COD_PRE','Codigo Pre','Código Pre'],
    cod_sismed:['Código Producto','codigo_producto','cod_sismed','CODSISMED','cod sismed','Código Producto'],
    estab:     ['Establecimiento','establecimiento'],
    producto:  ['Producto','producto'],
    servicio:  ['Tipo de Servicio','Servicio'],
    requerida: ['Cantidad Requerida','Requerida'],
    disponible:['Cantidad Disponible','Disponible'],
    noSat:     ['Demanda No Satisfecha','No Satisfecha'],
    cobertura: ['Cobertura (%)','Cobertura'],
    fecha:     ['Fecha','fecha'],
    obs:       ['Observaciones','observaciones'],
    usuario:   ['Usuario que Registró','Usuario'],
  };
  for (const [key, candidates] of Object.entries(COL_MAP)) {
    for (const c of candidates) { const i = h.findIndex(x => x === c); if (i !== -1) { idx[key] = i; break; } }
  }
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = parseLine(lines[i]);
    if (!cols.length || !cols[idx.estab || 2]) continue;
    const get    = key => idx[key] !== undefined ? (cols[idx[key]] || '').replace(/^"|"$/g,'') : '';
    const getNum = key => parseFloat(get(key).replace(/,/g,'.')) || 0;
    //Se cambia por nuevo, que permite considerar fechas desdel el 01 del mes 
    /*const fechaStr = get('fecha');
    let fecha = null, mesKey = 'Sin fecha', mesNombre = 'Sin fecha';
    if (fechaStr) {
      const d = new Date(fechaStr.replace(/(\d{2})\/(\d{2})\/(\d{4})/,'$3-$2-$1'));
      if (!isNaN(d)) { fecha = d; mesKey = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0'); mesNombre = d.toLocaleDateString('es-PE',{month:'long',year:'numeric'}); }
    }*/

    const fechaStr = get('fecha');
    let fecha = null, mesKey = 'Sin fecha', mesNombre = 'Sin fecha';

    if (fechaStr) {

      // FORMATO: YYYY-MM-DD
      const [anio, mes, dia] = fechaStr.split('-').map(Number);

      const d = new Date(anio, mes - 1, dia);

      if (!isNaN(d)) {
        fecha = d;

        mesKey =
          d.getFullYear() + '-' +
          String(d.getMonth() + 1).padStart(2,'0');

        mesNombre = d.toLocaleDateString('es-PE', {
          month:'long',
          year:'numeric'
        });
      }
    }
    const req = getNum('requerida'), disp = getNum('disponible');
    const noS = getNum('noSat') || Math.max(0, req - disp);
    const cob = getNum('cobertura');
    const codPre = normalizeCodigo(get('cod_pre'));
    const codSismed = normalizeCodigo(get('cod_sismed'));
    const red = RED_MAP[codPre] || 'Sin red asignada';
    rows.push({ redes:red, cod_pre:codPre, cod_sismed:codSismed, estab:get('estab'), producto:get('producto'), servicio:get('servicio'), requerida:req, disponible:disp, noSat:noS, cobertura:cob, fecha, mesKey, mesNombre, obs:get('obs'), usuario:get('usuario') });
  }
  return rows.filter(r => r.estab && r.producto);
}

// ═══════════════════════════════════════════════════════════════
// PARSE CSV — Hoja DISPO_CENTROS_ALM_SINC (ICI)
// ═══════════════════════════════════════════════════════════════
function parseCSVICI(csv) {
  const lines = csv.split(/\r?\n/).map(l => l.replace(/\r$/, '')).filter(l => l.trim());
  if (lines.length < 2) throw new Error('La hoja ICI está vacía.');
  function parseLine(line) {
    const result = []; let current = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { inQ = !inQ; continue; }
      if (ch === ',' && !inQ) { result.push(current.trim()); current = ''; continue; }
      current += ch;
    }
    result.push(current.trim()); return result;
  }
  const headerIndex = lines.findIndex(line => /codigo_pre|cod_pre/i.test(line) && /cod_sismed|codigo_producto/i.test(line));
  if (headerIndex === -1) throw new Error('No se encontró la fila de encabezado en la hoja ICI.');
  const h = parseLine(lines[headerIndex]).map(h => h.replace(/^"|"$/g,'').trim());
  const dataLines = lines.slice(headerIndex + 1);
  // Mapeo flexible de columnas del Sheet DISPO
  const idx = {};
  const COL_ICI = {
    cod_pre:       ['codigo_pre','cod_pre','COD_PRE','Cod_pre','Codigo_pre','Código_pre'],
    estab:         ['establecimiento','Establecimiento','ESTABLECIMIENTO'],
    red:           ['RED','Red','red'],
    cod_sismed:    ['cod_sismed','COD_SISMED','cod sismed','cod_sismed','codigo_producto','Código Producto'],
    descripcion:   ['descripcion','Descripcion','DESCRIPCION','descripción'],
    tipo:          ['TIPO','Tipo','tipo'],
    estrategico:   ['estrategico','Estrategico','ESTRATEGICO'],
    suministro:    ['SUMINISTRO','Suministro','suministro'],
    stock_ce:      ['STOCK_CENTROS','stock_centros','Stock_centros'],
    cmpa_ce:       ['CPMA_CENTROS','CMPA_CENTROS','cpma_centros'],
    disp_ce:       ['DISP_CENTROS','disp_centros','Disp_centros'],
    indicador_ce:  ['INDICADOR_CE','indicador_ce','INDICADOR_CE','INDICADOR_CENTROS','INDICADOR_CENTRO','INDICADOR_CENTROs'],
    stock_alm:     ['STOCK_ALM','stock_alm','Stock_alm'],
    cmpa_alm:      ['CMPA_ALM','cmpa_alm','CMPA ALM'],
    disp_alm:      ['DIP_ALM','DISP_ALM','dip_alm','disp_alm'],
    indicador_alm: ['INDICADOR_ALM','indicador_alm','Indicador_ALM'],
    fecha_reporte: ['FECHA_REPORTE','fecha_reporte','Fecha_Reporte'],
  };
  for (const [key, candidates] of Object.entries(COL_ICI)) {
    for (const c of candidates) { const i = h.findIndex(x => x === c); if (i !== -1) { idx[key] = i; break; } }
  }
  const rows = [];
  for (let i = 0; i < dataLines.length; i++) {
    const cols = parseLine(dataLines[i]);
    if (!cols.length) continue;
    const get    = key => idx[key] !== undefined ? (cols[idx[key]] || '').replace(/^"|"$/g,'').trim() : '';
    const getNum = key => parseFloat(get(key).replace(/,/g,'.')) || 0;
    const codPre = normalizeCodigo(get('cod_pre'));
    const codSismed = normalizeCodigo(get('cod_sismed'));
    if (!codPre || codPre === '00000') continue;
    rows.push({
      cod_pre:       codPre,
      cod_sismed:    codSismed,
      estab:         get('estab'),
      red:           get('red'),
      descripcion:   get('descripcion'),
      tipo:          get('tipo'),
      estrategico:   get('estrategico'),
      suministro:    get('suministro'),
      stock_ce:      getNum('stock_ce'),
      cmpa_ce:       getNum('cmpa_ce'),
      disp_ce:       getNum('disp_ce'),
      indicador_ce:  get('indicador_ce').toUpperCase(),
      stock_alm:     getNum('stock_alm'),
      cmpa_alm:      getNum('cmpa_alm'),
      disp_alm:      getNum('disp_alm'),
      indicador_alm: get('indicador_alm').toUpperCase(),
      fecha_reporte: get('fecha_reporte'),
    });
  }
  return rows;
}

// ═══════════════════════════════════════════════════════════════
// CRUCE: Productos No Atendidos ↔ ICI
// Criterio: COD_PRE + COD_SISMED exactos (códigos únicos)
// Si no hay match exacto → Producto Nuevo o Sin Historial en ICI
// ═══════════════════════════════════════════════════════════════

function cruzarDatos(noAtendidos, ici) {
  console.group('🔗 INICIANDO CRUCE DE DATOS');
  console.log('📦 Registros de entrada (noAtendidos):', noAtendidos.length);
  console.log('📦 Registros ICI:', ici.length);
  
  // Crear índice ICI por codigo_pre y cod_sismed (búsqueda exacta por código único)
  const iciPorClaveCompuesta = {};
  ici.forEach(row => {
    const k = `${row.cod_pre}||${row.cod_sismed}`;
    iciPorClaveCompuesta[k] = row;
  });
  console.log('🔑 Índice ICI por (COD_PRE || COD_SISMED):', Object.keys(iciPorClaveCompuesta).length, 'claves únicas');

  let contadores = { exactos: 0, sinMatch: 0 };

  const resultado = noAtendidos.map((na, idx) => {
    // Búsqueda EXACTA por codigo_pre + cod_sismed (los códigos son únicos)
    const kExacta = `${na.cod_pre}||${na.cod_sismed}`;
    const match = iciPorClaveCompuesta[kExacta] || null;

    // Si no hay match exacto = PRODUCTO NUEVO O SIN HISTORIAL EN ICI
    if (match) {
      contadores.exactos++;
    } else {
      contadores.sinMatch++;
    }

    // Log cada 100 registros (solo mostrar sin match para revisar)
    if (idx < 5 || (idx % 100 === 0)) {
      console.log(`  [${idx}] Producto: "${na.producto}" | COD_PRE: ${na.cod_pre} | COD_SISMED: ${na.cod_sismed} | Match: ${match ? '✅ EXACTO (COD_PRE + COD_SISMED)' : '❌ PRODUCTO NUEVO/SIN HISTORIAL'}`);
    }

    return {
      // Datos del no atendido
      redes:         na.redes,
      cod_pre:       na.cod_pre,
      cod_sismed:    na.cod_sismed,
      estab:         na.estab,
      producto:      na.producto,
      servicio:      na.servicio,
      noSat:         na.noSat,
      cobertura:     na.cobertura,
      fecha:         na.fecha,
      // Datos ICI cruzados (exacto o sin historial)
      match_nombre:  match ? match.descripcion : null,
      match_sim:     match ? 1.0 : 0,
      match_cod_sismed: match ? match.cod_sismed : null,
      indicador_ce:  match ? match.indicador_ce  : 'SIN_MATCH',
      indicador_alm: match ? match.indicador_alm : 'SIN_MATCH',
      stock_ce:      match ? match.stock_ce  : null,
      cmpa_ce:       match ? match.cmpa_ce   : null,
      disp_ce:       match ? match.disp_ce   : null,
      stock_alm:     match ? match.stock_alm : null,
      cmpa_alm:      match ? match.cmpa_alm  : null,
      disp_alm:      match ? match.disp_alm  : null,
      fecha_reporte: match ? match.fecha_reporte : null,
      tiene_match:   !!match,
    };
  });

  console.log('📊 RESULTADO FINAL DEL CRUCE:');
  console.log(`  ✅ Exactos (COD_PRE + COD_SISMED): ${contadores.exactos}`);
  console.log(`  ⚠️  Productos nuevos/sin historial en ICI: ${contadores.sinMatch}`);
  console.log(`  📈 Total con historial ICI: ${contadores.exactos} / ${noAtendidos.length} (${(contadores.exactos / noAtendidos.length * 100).toFixed(1)}%)`);
  console.groupEnd();
  
  return resultado;
}

// ═══════════════════════════════════════════════════════════════
// CRITERIO / DIAGNÓSTICO del cruce
// ═══════════════════════════════════════════════════════════════
function diagnosticarCruce(row) {
  if (!row.tiene_match) return { nivel: 'sin_datos', label: '⚪ PRODUCTO NUEVO', color: '#78909c', desc: 'Código no encontrado en ICI. Producto nuevo o sin historial de disponibilidad.' };

  const ie = row.indicador_ce;
  const ia = row.indicador_alm;

  // Caso crítico: desabastecido en centros Y sin stock en almacén
  if (ie === 'DESABASTECIDO' && (ia === 'DESABASTECIDO' || ia === 'SUBSTOCK')) {
    return { nivel: 'critico', label: '🔴 CRÍTICO', color: '#c62828',
      desc: `Centros: ${ie} · Almacén: ${ia}. Sin stock disponible para reponer. Requiere pedido urgente a SISMED.` };
  }
  // Alerta: desabastecido en centros pero hay algo en almacén
  if (ie === 'DESABASTECIDO' && ia === 'NORMOSTOCK') {
    return { nivel: 'alerta', label: '🟠 ALERTA — Redistribuir', color: '#e65100',
      desc: `Centros: ${ie} · Almacén: ${ia}. Hay stock en almacén, pero no llega a los centros. Revisar distribución.` };
  }
  if (ie === 'DESABASTECIDO' && ia === 'SOBRESTOCK') {
    return { nivel: 'alerta', label: '🟠 ALERTA — Redistribuir urgente', color: '#d84315',
      desc: `Centros: ${ie} · Almacén: ${ia}. Sobrestock en almacén y cero en centros. Error logístico de distribución.` };
  }
  // Substock en centros
  if (ie === 'SUBSTOCK' && (ia === 'DESABASTECIDO' || ia === 'SUBSTOCK')) {
    return { nivel: 'alto', label: '🟡 ALTO RIESGO', color: '#f9a825',
      desc: `Centros: ${ie} · Almacén: ${ia}. Stock bajo en ambos niveles. Programar reposición.` };
  }
  if (ie === 'SUBSTOCK' && (ia === 'NORMOSTOCK' || ia === 'SOBRESTOCK')) {
    return { nivel: 'moderado', label: '🟡 MODERADO — Distribuir', color: '#fbc02d',
      desc: `Centros: ${ie} · Almacén: ${ia}. Substock en centros pero almacén abastecido. Distribuir.` };
  }
  // Normostock o sobrestock en centros (producto no atendido por otra razón)
  if (ie === 'NORMOSTOCK' || ie === 'SOBRESTOCK') {
    return { nivel: 'revisar', label: '🔵 REVISAR', color: '#1565c0',
      desc: `Centros: ${ie} · Almacén: ${ia}. El ICI indica stock suficiente. Verificar si el producto no atendido corresponde a una presentación diferente o error de registro.` };
  }
  return { nivel: 'sin_datos', label: 'Indefinido', color: '#78909c', desc: `Indicadores: CE=${ie} / ALM=${ia}` };
}

// ═══════════════════════════════════════════════════════════════
// INIT / RELOAD
// ═══════════════════════════════════════════════════════════════
async function reloadData() {
  const btn = document.getElementById('refresh-btn');
  btn.classList.add('spinning');
  document.getElementById('dash-content').innerHTML = `<div class="loading-overlay"><div class="loader"></div><div class="loading-text">Cargando datos desde Google Sheets…</div></div>`;
  try {
    await cargarCatalogo();
    allData = await fetchSheetData();
    if (!allData.length) throw new Error('No se encontraron filas con datos en ENTRADAS.');

    // Intentar cargar ICI (no bloquea si falla)
    try {
      if (ICI_SHEET_ID !== 'TU_SHEET_ID_DISPO_CENTROS_ALM_AQUI') {
        document.querySelector('.loading-text') && (document.querySelector('.loading-text').textContent = 'Cruzando con datos ICI…');
        console.log('🔵 [INICIO] Cargando datos ICI desde Sheet ID:', ICI_SHEET_ID, 'Tab:', ICI_SHEET_TAB);
        iciData   = await fetchICIData();
        console.log('✅ [ICI CARGADO] Se obtuvieron', iciData.length, 'registros de ICI');
        console.log('📊 [MUESTRA ICI] Primeros 3 registros:', iciData.slice(0, 3));
        console.log('✅ [ENTRADAS CARGADAS] Se obtuvieron', allData.length, 'registros de ENTRADAS');
        console.log('📊 [MUESTRA ENTRADAS] Primeros 3 registros:', allData.slice(0, 3));
        cruceData = cruzarDatos(allData, iciData);
        console.log('🔗 [CRUCE COMPLETADO] Se cruzaron', cruceData.length, 'registros');
        const conMatch = cruceData.filter(r => r.tiene_match).length;
        const sinMatch = cruceData.filter(r => !r.tiene_match).length;
        console.log(`📈 [RESULTADO CRUCE] ${conMatch} con match · ${sinMatch} sin match`);
      } else {
        iciData   = [];
        cruceData = allData.map(r => ({ ...r, tiene_match: false, indicador_ce: 'SIN_MATCH', indicador_alm: 'SIN_MATCH' }));
        console.warn('ICI_SHEET_ID no configurado. Configura el ID del sheet de DISPO_CENTROS_ALM.');
      }
    } catch(iciErr) {
      console.warn('❌ No se pudo cargar ICI:', iciErr.message);
      console.error('Detalles del error:', iciErr);
      iciData   = [];
      cruceData = allData.map(r => ({ ...r, tiene_match: false, indicador_ce: 'SIN_MATCH', indicador_alm: 'SIN_MATCH' }));
    }

    buildPeriodTabs();
    renderAll();
  } catch(err) {
    document.getElementById('dash-content').innerHTML = `<div class="error-panel"><p>Error al cargar los datos</p><small>${err.message}</small><button class="btn-sm" onclick="reloadData()" style="margin-top:20px;background:#1565c0;border-color:#1565c0;">↺ Reintentar</button></div>`;
  } finally { btn.classList.remove('spinning'); }
}

// ═══════════════════════════════════════════════════════════════
// PERIOD TABS
// ═══════════════════════════════════════════════════════════════
function buildPeriodTabs() {
  const container = document.getElementById('period-pills');
  if (!container) return;
  container.innerHTML = '';
  const counts = {};
  allData.forEach(r => { counts[r.mesKey] = (counts[r.mesKey]||0) + 1; });
  const meses = Object.keys(counts).sort();

  const addPill = (label, key, count) => {
    const btn = document.createElement('button');
    btn.className = 'period-pill' + (currentPeriod === key ? ' active' : '');
    btn.innerHTML = `${label}<span class="ppbadge">${count}</span>`;
    btn.onclick = () => { currentPeriod = key; container.querySelectorAll('.period-pill').forEach(b=>b.classList.remove('active')); btn.classList.add('active'); page=1; renderAll(); };
    container.appendChild(btn);
  };
  addPill('Todos', 'todos', allData.length);
  const dotColors = ['#1565c0','#22c55e','#f97316','#a78bfa','#f59e0b','#06b6d4'];
  meses.forEach((m, i) => {
    const sample = allData.find(r => r.mesKey === m);
    const nombre = sample ? sample.mesNombre : m;
    addPill(nombre, m, counts[m]);
  });
}
function getFiltered() {
  return currentPeriod === 'todos' ? allData : allData.filter(r => r.mesKey === currentPeriod);
}
function getCruceFiltered() {
  const periodoData = currentPeriod === 'todos' ? allData : allData.filter(r => r.mesKey === currentPeriod);
  const codSet = new Set(periodoData.map(r => `${r.cod_pre}||${r.producto}`));
  return cruceData.filter(r => codSet.has(`${r.cod_pre}||${r.producto}`));
}

function groupRecords(data) {
  const map = {};
  data.forEach(r => {
    const key = `${r.redes}||${r.estab}||${r.cod_pre}||${r.cod_sismed}||${r.servicio}`;
    if (!map[key]) {
      map[key] = {
        redes: r.redes,
        estab: r.estab,
        cod_pre: r.cod_pre,
        cod_sismed: r.cod_sismed,
        producto: r.producto,
        servicio: r.servicio,
        requerida: 0,
        disponible: 0,
        noSat: 0,
        coberturaSum: 0,
        count: 0,
        fecha: r.fecha,
        mesKey: r.mesKey || ''
      };
    }
    const item = map[key];
    item.requerida += r.requerida;
    item.disponible += r.disponible;
    item.noSat += r.noSat;
    item.coberturaSum += r.cobertura;
    item.count += 1;
    if (!item.fecha || (r.fecha && r.fecha > item.fecha)) item.fecha = r.fecha;
  });
  return Object.values(map).map(item => ({
    ...item,
    cobertura: item.requerida ? (item.disponible / item.requerida * 100) : 0
  }));
}

function getTodayRecords(data) {
  const today = new Date();
  const todayKey = today.toISOString().slice(0,10);
  const weekAgo = new Date(today);
  weekAgo.setDate(weekAgo.getDate() - 7);
  const todayRows = data.filter(r => r.fecha && r.fecha.toISOString().slice(0,10) === todayKey);
  const recentRows = data.filter(r => r.fecha && r.fecha >= weekAgo).sort((a,b) => b.fecha - a.fecha);
  return { today: todayRows, recent: recentRows };
}

function formatDateForCSV(date) {
  return date instanceof Date ? date.toLocaleDateString('es-PE') : '';
}

function buildDateGroupedData() {
  if (!filteredData || !filteredData.length || dateGroupMode === 'none') { dateGroupedData = []; return; }
  const rows = filteredData.filter(r => r.fecha);
  if (!rows.length) { dateGroupedData = []; return; }
  // determine target date set
  const dates = [...new Set(rows.map(r => r.fecha.toISOString().slice(0,10)))].sort();
  let targetDates = [];
  if (dateGroupMode === 'all') targetDates = dates;
  if (dateGroupMode === 'last') targetDates = [dates[dates.length-1]];
  if (dateGroupMode === 'today') targetDates = [new Date().toISOString().slice(0,10)];

  const map = {};
  rows.forEach(r => {
    const dKey = r.fecha.toISOString().slice(0,10);
    if (!targetDates.includes(dKey)) return;
    const key = `${dKey}||${r.cod_pre}||${r.cod_sismed}||${r.estab}||${r.servicio}`;
    if (!map[key]) map[key] = { dateKey: dKey, fecha: r.fecha, redes: r.redes, estab: r.estab, cod_pre: r.cod_pre, cod_sismed: r.cod_sismed, producto: r.producto, servicio: r.servicio, requerida:0, disponible:0, noSat:0, count:0 };
    const it = map[key];
    it.requerida += r.requerida; it.disponible += r.disponible; it.noSat += r.noSat; it.count += 1;
    if (!it.fecha || (r.fecha && r.fecha > it.fecha)) it.fecha = r.fecha;
  });
  dateGroupedData = Object.values(map).map(it => ({ ...it, cobertura: it.requerida ? (it.disponible / it.requerida * 100) : 0 }));
}

// ═══════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════
// TAB SWITCHING
// ═══════════════════════════════════════════════════════════════
let activeTab = 'resumen';
function switchTab(tab) {
  activeTab = tab;
  document.querySelectorAll('.nav-item').forEach(el => {
    el.classList.toggle('active', el.dataset.tab === tab);
  });
  document.querySelectorAll('.tab-page').forEach(el => {
    el.classList.toggle('active', el.dataset.tab === tab);
  });
  const titles = {
    resumen: 'Resumen Ejecutivo', brecha: 'Brecha por Establecimiento',
    ici: 'Cruce ICI · Disponibilidad Diaria', detalle: 'Detalle Completo', agrupado: 'Resumen Agrupado',
    nuevos: 'Resumen por Producto · Vista General'
  };
  const titleEl = document.getElementById('tab-title');
  if (titleEl) titleEl.textContent = titles[tab] || '';
}

// ═══════════════════════════════════════════════════════════════
// RENDER ALL
// ═══════════════════════════════════════════════════════════════
function renderAll() {
  filteredData = getFiltered();
  const totalNoSat = filteredData.reduce((s,r)=>s+r.noSat,0);
  const totalReq   = filteredData.reduce((s,r)=>s+r.requerida,0);
  const totalDisp  = filteredData.reduce((s,r)=>s+r.disponible,0);
  const cobertura  = totalReq > 0 ? (totalDisp/totalReq*100) : 0;
  const periodo    = currentPeriod === 'todos' ? 'Todos los períodos' : (filteredData[0]?.mesNombre || currentPeriod);
  document.getElementById('tb-meta').textContent =
    `${periodo} · ${filteredData.length.toLocaleString()} registros · ${totalNoSat.toLocaleString()} u sin atender · Cobertura: ${cobertura.toFixed(1)}%`;
  const sbMeta = document.getElementById('sb-meta');
  if (sbMeta) sbMeta.textContent = `${filteredData.length.toLocaleString()} registros · ${cobertura.toFixed(1)}% cobertura`;
  renderBanner(cobertura, totalNoSat);
  document.getElementById('dash-content').innerHTML = buildDashHTML();
  switchTab(activeTab);
  renderKPIs(filteredData);
  renderGauge(cobertura, totalDisp, totalReq);
  renderEstabChart(filteredData);
  renderTopProds(filteredData);
  renderTodayRecords(filteredData);
  renderAlerts(filteredData);
  groupPage = 1;
  buildDateGroupedData();
  renderGroupedSummary(filteredData);
  renderServicioChart(filteredData);
  populateFilters();
  filterTable();
  renderICIKPIs();
  iciGroupedData = buildICIGrouped(getCruceFiltered());
  renderCruceTable();
  renderPersistentesAlert();
  renderRedChart(filteredData);
  renderNuevosTable();
}

// ═══════════════════════════════════════════════════════════════
// BANNER
// ═══════════════════════════════════════════════════════════════
function renderBanner(cob, noSat) {
  const el = document.getElementById('global-banner');
  if (!el) return;
  if (cob < 10) {
    el.innerHTML = `<div class="banner crit"><div class="banner-icon">🚨</div><div><div class="banner-title">Crisis crítica — Cobertura global ${cob.toFixed(1)}%</div><div class="banner-body">${noSat.toLocaleString()} unidades no atendidas. Se requiere reposición urgente de medicamentos.</div></div></div>`;
  } else if (cob < 50) {
    el.innerHTML = `<div class="banner warn"><div class="banner-icon">⚠️</div><div><div class="banner-title">Alerta de abastecimiento — Cobertura ${cob.toFixed(1)}%</div><div class="banner-body">Revisar pedidos con SISMED y coordinar redistribución entre establecimientos.</div></div></div>`;
  } else { el.innerHTML = ''; }
}

// ═══════════════════════════════════════════════════════════════
// HTML SCAFFOLD — TAB PAGES
// ═══════════════════════════════════════════════════════════════
function buildDashHTML() {
  const iciConfigured = ICI_SHEET_ID !== 'TU_SHEET_ID_DISPO_CENTROS_ALM_AQUI';
  const iciStatus = iciConfigured
    ? (iciData.length ? `<span style="color:#2e7d32">✓ ${iciData.length.toLocaleString()} reg. ICI</span>` : `<span style="color:#c62828">✗ ICI no disponible</span>`)
    : `<span style="color:#e65100">⚠ Configura ICI_SHEET_ID</span>`;
  return `
  <!-- TAB: RESUMEN -->
  <div class="tab-page active" data-tab="resumen">
    <div class="kpi-grid" id="kpi-grid" style="flex-shrink:0"></div>
    <div style="flex:1;min-height:0;display:grid;grid-template-columns:1.5fr 1fr;gap:12px">
      <div class="card" style="min-height:0;overflow:hidden">
        <div class="card-hdr" style="flex-shrink:0">
          <div><div class="card-title">Top 10 Productos con Mayor Brecha</div><div class="card-sub">Unidades sin atender por producto</div></div>
        </div>
        <div style="overflow-x:auto;flex:1;overflow-y:auto;min-height:0">
          <table class="top-table">
            <thead><tr><th style="width:24px">#</th><th>Producto</th><th>Sin atender</th><th>Estado</th></tr></thead>
            <tbody id="top-tbody"></tbody>
          </table>
        </div>
      </div>
      <div class="card" style="min-height:0;overflow:hidden">
        <div class="card-hdr" style="flex-shrink:0">
          <div><div class="card-title">🔴 Alertas Críticas</div><div class="card-sub">Productos con 0% de cobertura</div></div>
        </div>
        <div class="alert-list" id="alerts-list"></div>
      </div>
    </div>
  </div>

  <!-- TAB: BRECHA -->
  <div class="tab-page" data-tab="brecha">
    <div style="flex:1;min-height:0;display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="card" style="min-height:0">
        <div class="card-hdr" style="flex-shrink:0">
          <div><div class="card-title">Brecha por Establecimiento</div><div class="card-sub">Top 12 · atendido vs sin atender · % cobertura en barra</div></div>
          <span class="card-tag" id="tag-estab">—</span>
        </div>
        <div style="flex:1;min-height:0;position:relative"><canvas id="cEstab"></canvas></div>
      </div>
      <div style="display:flex;flex-direction:column;gap:12px;min-height:0">
        <div class="card" style="flex:1;min-height:0">
          <div class="card-hdr" style="flex-shrink:0">
            <div><div class="card-title">Brecha por Red</div><div class="card-sub">Unidades sin atender por red · % cobertura</div></div>
          </div>
          <div style="flex:1;min-height:0;position:relative"><canvas id="cRed"></canvas></div>
        </div>
        <div class="card" style="flex:0 0 auto">
          <div class="card-hdr" style="margin-bottom:8px">
            <div><div class="card-title">Cobertura Global</div></div>
          </div>
          <div class="gauge-wrap" id="gauge-wrap"></div>
        </div>
      </div>
    </div>
    <div style="flex:0 0 auto;display:grid;grid-template-columns:1.5fr 1fr;gap:12px">
      <div class="card">
        <div class="card-hdr" style="flex-shrink:0">
          <div><div class="card-title">Por Tipo de Servicio</div><div class="card-sub">Distribución de brecha por servicio</div></div>
        </div>
        <div style="height:140px;position:relative"><canvas id="cServicio"></canvas></div>
      </div>
      <div class="explainer" style="align-self:stretch">
        <div class="explainer-icon">💡</div>
        <div>
          <div class="explainer-title">¿Qué significa la cobertura?</div>
          <div class="explainer-body">La cobertura indica qué % de la demanda <strong>sí pudo atenderse</strong>. El % mostrado dentro de cada barra es la cobertura del establecimiento. Rojo = sin atender, Verde = atendido.</div>
        </div>
      </div>
    </div>
  </div>

  <!-- TAB: CRUCE ICI -->
  <div class="tab-page" data-tab="ici">
    <div class="kpi-grid" id="ici-kpi-grid" style="flex-shrink:0"></div>
    <!-- Alerta persistente -->
    <div class="card" style="flex-shrink:0">
      <div class="card-hdr" style="margin-bottom:8px">
        <div><div class="card-title">🚨 Desabastecimiento Persistente</div><div class="card-sub">Productos con ≥2 días consecutivos desabastecidos o en substock comparando con ICI</div></div>
      </div>
      <div id="persistentes-list" style="display:flex;flex-direction:column;gap:6px;max-height:160px;overflow-y:auto"></div>
    </div>
    <div class="tbl-card" style="flex:1;min-height:0">
      <div class="tbl-bar" style="flex-wrap:wrap;gap:6px">
        <span class="tbl-bar-title">Cruce ICI × No Atendidos <span style="font-size:10px;font-weight:400;margin-left:6px">${iciStatus}</span></span>
        <input class="inp" type="text" id="ici-search" placeholder="🔍 Buscar producto o centro…" style="width:160px" oninput="iciPage=1;renderCruceTable()">
        <select class="inp" id="ici-dia" onchange="iciPage=1;renderCruceTable()" style="max-width:150px">
          <option value="">Todos los días</option>
        </select>
        <select class="inp" id="ici-mes" onchange="iciPage=1;renderCruceTable()" style="max-width:140px">
          <option value="">Todos los meses</option>
        </select>
        <select class="inp" id="ici-centro" onchange="iciPage=1;renderCruceTable()" style="max-width:180px">
          <option value="">Todos los centros</option>
        </select>
        <select class="inp" id="ici-red" onchange="iciPage=1;renderCruceTable()">
          <option value="">Todas las redes</option>
        </select>
        <select class="inp" id="ici-indicador" onchange="iciPage=1;renderCruceTable()">
          <option value="">Todo CE</option>
          <option value="DESABASTECIDO">🔴 DESAB.</option>
          <option value="SUBSTOCK">🟡 SUBSTOCK</option>
          <option value="NORMOSTOCK">🔵 NORMO</option>
          <option value="SOBRESTOCK">🟢 SOBRE</option>
          <option value="SIN_MATCH">⚪ Sin match</option>
        </select>
        <select class="inp" id="ici-nivel" onchange="iciPage=1;renderCruceTable()">
          <option value="">Todo nivel</option>
          <option value="critico">🔴 Crítico</option>
          <option value="alerta">🟠 Alerta</option>
          <option value="alto">🟡 Alto riesgo</option>
          <option value="moderado">🟡 Moderado</option>
          <option value="revisar">🔵 Revisar</option>
          <option value="sin_datos">⚪ Sin datos</option>
        </select>
        <span class="tbl-count" id="ici-count"></span>
        <button class="btn-sm" onclick="exportCruceCSV()">⬇ Excel</button>
      </div>
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th onclick="sortICI('fecha')" style="white-space:nowrap">Fecha ↓</th>
            <th onclick="sortICI('redes')">Red</th>
            <th onclick="sortICI('estab')">Centro</th>
            <th onclick="sortICI('cod_sismed')">COD SISMED</th>
            <th onclick="sortICI('producto')">Producto no atendido</th>
            <th onclick="sortICI('noSat')" style="text-align:right">Sin atender</th>
            <th onclick="sortICI('indicador_ce')">ICI Centros</th>
            <th onclick="sortICI('indicador_alm')">ICI Almacén</th>
            <th style="text-align:right">Stock CE</th>
            <th style="text-align:right">Stock ALM</th>
            <th onclick="sortICI('nivel')">Diagnóstico</th>
            <th>Match ICI</th>
          </tr></thead>
          <tbody id="ici-body"></tbody>
        </table>
      </div>
      <div class="tbl-pager">
        <button class="btn-sm" onclick="prevICI()">← Anterior</button>
        <span id="ici-pager"></span>
        <button class="btn-sm" onclick="nextICI()">Siguiente →</button>
      </div>
    </div>
  </div>

  <!-- TAB: DETALLE -->
  <div class="tab-page" data-tab="detalle">
    <div class="tbl-card" style="flex:1;min-height:0">
      <div class="tbl-bar">
        <span class="tbl-bar-title">Registros</span>
        <input class="inp" type="text" id="t-search" placeholder="🔍 Buscar producto o establecimiento…" style="width:200px" oninput="filterTable()">
        <select class="inp" id="t-redes" onchange="filterTable()"><option value="">Todas las redes</option></select>
        <select class="inp" id="t-estab" onchange="filterTable()"><option value="">Todos los EESS</option></select>
        <select class="inp" id="t-servicio" onchange="filterTable()"><option value="">Todos los servicios</option></select>
        <select class="inp" id="t-cob" onchange="filterTable()">
          <option value="">Toda cobertura</option>
          <option value="0">Sin cobertura (0%)</option>
          <option value="parcial">Parcial</option>
        </select>
        <select class="inp" id="t-date-mode" onchange="dateGroupMode=this.value;buildDateGroupedData();filterTable()">
          <option value="none">Sin agrupar</option>
          <option value="all">Agrupar todo</option>
          <option value="last" selected>Último día</option>
          <option value="today">Hoy</option>
        </select>
        <span class="tbl-count" id="t-count"></span>
        <button class="btn-sm" onclick="exportCSV()">⬇ Excel</button>
      </div>
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th onclick="sortTbl('redes')">Red</th>
            <th onclick="sortTbl('cod_pre')">COD PRE</th>
            <th onclick="sortTbl('estab')">Establecimiento</th>
            <th onclick="sortTbl('cod_sismed')">COD SISMED</th>
            <th onclick="sortTbl('producto')">Producto</th>
            <th onclick="sortTbl('servicio')">Servicio</th>
            <th onclick="sortTbl('requerida')" style="text-align:right">Req.</th>
            <th onclick="sortTbl('disponible')" style="text-align:right">Disp.</th>
            <th onclick="sortTbl('noSat')">Sin atender</th>
            <th onclick="sortTbl('cobertura')">Cobertura</th>
            <th onclick="sortTbl('fecha')">Fecha</th>
          </tr></thead>
          <tbody id="t-body"></tbody>
        </table>
      </div>
      <div class="tbl-pager">
        <button class="btn-sm" onclick="prevPage()">← Anterior</button>
        <span id="t-pager"></span>
        <button class="btn-sm" onclick="nextPage()">Siguiente →</button>
      </div>
    </div>
  </div>

  <!-- TAB: AGRUPADO -->
  <div class="tab-page" data-tab="agrupado">
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;flex-shrink:0">
      <div class="card">
        <div class="card-hdr" style="margin-bottom:6px">
          <div><div class="card-title">Registros de Hoy</div><div class="card-sub">Fecha actual</div></div>
          <button class="btn-sm" onclick="exportTodayCSV()">⬇ Excel Hoy</button>
        </div>
        <div class="alert-list" id="today-list" style="max-height:180px"></div>
      </div>
      <div class="card">
        <div class="card-hdr" style="margin-bottom:6px">
          <div><div class="card-title">Criterio de Cruce ICI</div><div class="card-sub">Guía de diagnóstico</div></div>
        </div>
        <div style="font-size:11px;color:var(--muted2);line-height:1.8">
          <span style="color:#c62828">🔴 <strong>CRÍTICO</strong></span> — Desabast. en centros Y almacén → pedido urgente a SISMED<br>
          <span style="color:#e65100">🟠 <strong>ALERTA</strong></span> — Desabast. en centros, stock en ALM → falla logística<br>
          <span style="color:#f9a825">🟡 <strong>ALTO RIESGO</strong></span> — Substock en ambos niveles → programar reposición<br>
          <span style="color:#1565c0">🔵 <strong>REVISAR</strong></span> — ICI indica normostock → verificar registro<br>
          <span style="color:#78909c">⚪ <strong>NUEVO</strong></span> — Código sin historial en ICI
        </div>
      </div>
    </div>
    <div class="tbl-card" style="flex:1;min-height:0">
      <div class="tbl-bar" style="flex-wrap:wrap;gap:6px">
        <span class="tbl-bar-title">Resumen Agrupado por Producto</span>
        <input class="inp" type="text" id="g-search" placeholder="🔍 Buscar producto, centro, código…" style="width:210px" oninput="filterGrouped()">
        <select class="inp" id="g-red" onchange="filterGrouped()">
          <option value="">Todas las redes</option>
        </select>
        <select class="inp" id="g-estab" onchange="filterGrouped()" style="max-width:180px">
          <option value="">Todos los EESS</option>
        </select>
        <select class="inp" id="g-mes" onchange="filterGrouped()" style="max-width:150px">
          <option value="">Todos los meses</option>
        </select>
        <span class="tbl-count" id="group-count"></span>
        <button class="btn-sm" onclick="exportGroupedCSV()">⬇ Excel Agrupado</button>
      </div>
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th>Red</th><th>Establecimiento</th><th>COD PRE</th><th>COD SISMED</th><th>Producto</th><th>Servicio</th>
            <th style="text-align:right">Req.</th><th style="text-align:right">Disp.</th><th style="text-align:right">Sin atender</th><th>Cobertura</th>
          </tr></thead>
          <tbody id="group-body"></tbody>
        </table>
      </div>
      <div class="tbl-pager">
        <button class="btn-sm" onclick="prevGroupPage()">← Anterior</button>
        <span id="group-pager"></span>
        <button class="btn-sm" onclick="nextGroupPage()">Siguiente →</button>
      </div>
    </div>
  </div>

  <!-- TAB: PRODUCTOS NUEVOS -->
  <div class="tab-page" data-tab="nuevos">
    <!-- KPIs resumen -->
    <div class="kpi-grid" id="np-kpi-grid" style="flex-shrink:0"></div>

    <!-- Tabla principal -->
    <div class="tbl-card" style="flex:1;min-height:0">
      <div class="tbl-bar" style="flex-wrap:wrap;gap:6px">
        <span class="tbl-bar-title">Resumen por Producto</span>
        <input class="inp" type="text" id="np-search" placeholder="🔍 Buscar producto o código…" style="width:190px" oninput="npPage=1;renderNuevosTable()">
        <select class="inp" id="np-mes" onchange="npPage=1;renderNuevosTable()" style="max-width:145px">
          <option value="">Todos los meses</option>
        </select>
        <select class="inp" id="np-dia" onchange="npPage=1;renderNuevosTable()" style="max-width:145px">
          <option value="">Todos los días</option>
        </select>
        <select class="inp" id="np-origen" onchange="npPage=1;renderNuevosTable()" style="max-width:160px">
          <option value="todos">Todos los productos</option>
          <option value="nuevo">⚪ Sin historial ICI</option>
          <option value="ici">✅ Con historial ICI</option>
        </select>
        <select class="inp" id="np-vista" style="display:none">
          <option value="producto" selected>Por producto</option>
        </select>
        <span class="tbl-count" id="np-count"></span>
        <button class="btn-sm" onclick="exportNuevosXLSX()">⬇ Excel</button>
      </div>
      <div class="tbl-wrap">
        <table>
          <thead id="np-thead"><tr></tr></thead>
          <tbody id="np-body"></tbody>
        </table>
      </div>
      <div class="tbl-pager">
        <button class="btn-sm" onclick="prevNp()">← Anterior</button>
        <span id="np-pager"></span>
        <button class="btn-sm" onclick="nextNp()">Siguiente →</button>
      </div>
    </div>
  </div>
  `;
}



// KPIs ORIGINALES
// ═══════════════════════════════════════════════════════════════
function renderKPIs(data) {
  const totalNoSat  = data.reduce((s,r)=>s+r.noSat,0);
  const totalReq    = data.reduce((s,r)=>s+r.requerida,0);
  const totalDisp   = data.reduce((s,r)=>s+r.disponible,0);
  const cobertura   = totalReq > 0 ? (totalDisp/totalReq*100) : 0;
  const nEstabs     = new Set(data.map(r=>r.estab)).size;
  const nProds      = new Set(data.map(r=>r.producto)).size;
  const nRedes      = new Set(data.map(r=>r.redes)).size;
  const sin0        = data.filter(r=>r.cobertura===0).length;

  const kpis = [
    { icon:'📦', lbl:'Total sin atender', val: totalNoSat.toLocaleString('es-PE') + ' u', sub:`de ${totalReq.toLocaleString()} requeridas`, c:'c-red' },
    { icon:'📊', lbl:'Cobertura global',  val: cobertura.toFixed(1) + '%', sub:`${totalDisp.toLocaleString()} u atendidas`, c: cobertura<30?'c-red':cobertura<70?'c-orange':'c-green' },
    { icon:'🏥', lbl:'Establecimientos',  val: nEstabs, sub:`${nRedes} redes`, c:'c-blue' },
    { icon:'💊', lbl:'Productos afectados',val: nProds,  sub:'productos distintos', c:'c-blue' },
    //{ icon:'🚨', lbl:'Sin cobertura (0%)',val: sin0,    sub:'registros sin atender nada', c:'c-red' },
  ];
  document.getElementById('kpi-grid').innerHTML = kpis.map(k=>`
    <div class="kpi ${k.c}">
      <span class="kpi-icon">${k.icon}</span>
      <div class="kpi-lbl">${k.lbl}</div>
      <div class="kpi-val">${k.val}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>`).join('');
}

// ═══════════════════════════════════════════════════════════════
// KPIs ICI
// ═══════════════════════════════════════════════════════════════
function renderICIKPIs() {
  const el = document.getElementById('ici-kpi-grid');
  if (!el) return;
  const cruce = getCruceFiltered();
  if (!cruce.length || !iciData.length) {
    el.innerHTML = `<div style="color:var(--muted);font-size:12px;padding:12px;">Sin datos ICI disponibles. Configura ICI_SHEET_ID en script.js.</div>`;
    return;
  }
  const conMatch   = cruce.filter(r => r.tiene_match);
  const sinMatch   = cruce.filter(r => !r.tiene_match);
  const criticos   = cruce.filter(r => diagnosticarCruce(r).nivel === 'critico');
  const alertas    = cruce.filter(r => diagnosticarCruce(r).nivel === 'alerta');
  const desabCE    = conMatch.filter(r => r.indicador_ce  === 'DESABASTECIDO');
  const desabALM   = conMatch.filter(r => r.indicador_alm === 'DESABASTECIDO');
  const pctMatch   = cruce.length > 0 ? (conMatch.length/cruce.length*100).toFixed(0) : 0;

  const kpis = [
    { icon:'🔴', lbl:'Críticos (ambos niveles)',   val: criticos.length, sub:'pedido urgente a SISMED', c:'c-red' },
    { icon:'🟠', lbl:'Alerta redistribución',       val: alertas.length,  sub:'stock en ALM, sin llegar a centros', c:'c-orange' },
    { icon:'🏪', lbl:'Desabastecido en Centros',    val: desabCE.length,  sub:'INDICADOR_CE = DESABASTECIDO', c:'c-red' },
    { icon:'🏭', lbl:'Desabastecido en Almacén',    val: desabALM.length, sub:'INDICADOR_ALM = DESABASTECIDO', c:'c-orange' },
    { icon:'🔗', lbl:'Match ICI encontrado',        val: `${conMatch.length} / ${cruce.length}`, sub:`${pctMatch}% productos cruzados`, c:'c-blue' },
  ];
  el.innerHTML = kpis.map(k=>`
    <div class="kpi ${k.c}">
      <span class="kpi-icon">${k.icon}</span>
      <div class="kpi-lbl">${k.lbl}</div>
      <div class="kpi-val">${k.val}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>`).join('');
}

// ═══════════════════════════════════════════════════════════════
// TABLA CRUCE ICI — agrupada por día + centro + producto
// ═══════════════════════════════════════════════════════════════
function renderCruceTable() {
  const el     = document.getElementById('ici-body');
  const cnt    = document.getElementById('ici-count');
  const pager  = document.getElementById('ici-pager');
  if (!el) return;

  // Poblar selects (sólo una vez o si cambió el dataset)
  function poblarSelect(id, vals, fmtLabel) {
    const sr = document.getElementById(id);
    if (!sr) return;
    const prev = sr.value;
    sr.innerHTML = sr.options[0].outerHTML; // mantener primer "Todos"
    vals.forEach(v => { const o=document.createElement('option'); o.value=v; o.textContent=fmtLabel?fmtLabel(v):v; sr.appendChild(o); });
    sr.value = prev;
  }
  const allMeses   = [...new Set(iciGroupedData.filter(r=>r.mesKey&&r.mesKey!=='sin-fecha').map(r=>r.mesKey))].sort().reverse();
  const allDias    = [...new Set(iciGroupedData.filter(r=>r.fecha&&r.fecha!=='sin-fecha').map(r=>r.fecha))].sort().reverse();
  const allCentros = [...new Set(iciGroupedData.map(r=>r.estab))].sort();
  const allRedes2  = [...new Set(iciGroupedData.map(r=>r.redes))].sort();
  poblarSelect('ici-mes',    allMeses,   k => mesLabel(k));
  poblarSelect('ici-dia',    allDias,    d => new Date(d+'T00:00:00').toLocaleDateString('es-PE',{day:'2-digit',month:'short',year:'numeric'}));
  poblarSelect('ici-centro', allCentros, null);
  poblarSelect('ici-red',    allRedes2,  null);

  const busq    = (document.getElementById('ici-search')?.value||'').toLowerCase();
  const filtInd = document.getElementById('ici-indicador')?.value||'';
  const filtRed = document.getElementById('ici-red')?.value||'';
  const filtNiv = document.getElementById('ici-nivel')?.value||'';
  const filtDia = document.getElementById('ici-dia')?.value||'';
  const filtCen = document.getElementById('ici-centro')?.value||'';
  const filtMes = document.getElementById('ici-mes')?.value||'';

  let rows = iciGroupedData.map(r => ({ ...r, _diag: diagnosticarCruce(r) })).filter(r => {
    if (filtRed && r.redes !== filtRed) return false;
    if (filtInd && r.indicador_ce !== filtInd) return false;
    if (filtNiv && r._diag.nivel !== filtNiv) return false;
    if (filtDia && r.fecha !== filtDia) return false;
    if (filtCen && r.estab !== filtCen) return false;
    if (filtMes && r.mesKey !== filtMes) return false;
    if (busq && !r.producto.toLowerCase().includes(busq) && !r.estab.toLowerCase().includes(busq) && !r.cod_sismed?.includes(busq)) return false;
    return true;
  });

  const totalP = Math.ceil(rows.length/PAGE)||1;
  if (iciPage > totalP) iciPage = 1;
  if (cnt)   cnt.textContent  = rows.length.toLocaleString() + ' registros agrupados';
  if (pager) pager.textContent = `Pág. ${iciPage} / ${totalP}`;

  const slice = rows.slice((iciPage-1)*PAGE, iciPage*PAGE);

  const indicBadge = (ind) => {
    const map = {
      'DESABASTECIDO': 'background:#c62828;color:#ffcdd2',
      'SUBSTOCK':      'background:#e65100;color:#ffe0b2',
      'NORMOSTOCK':    'background:#1565c0;color:#bbdefb',
      'SOBRESTOCK':    'background:#2e7d32;color:#c8e6c9',
      'SIN_MATCH':     'background:#78909c;color:#eceff1',
    };
    const style = map[ind] || map['SIN_MATCH'];
    return `<span style="${style};padding:2px 8px;border-radius:4px;font-size:10px;font-family:Space Mono,monospace;white-space:nowrap">${ind}</span>`;
  };

  const fmtFecha = (d) => {
    if (!d || d === 'sin-fecha') return '—';
    return new Date(d+'T00:00:00').toLocaleDateString('es-PE',{day:'2-digit',month:'short',year:'numeric'});
  };

  el.innerHTML = slice.length ? slice.map(r => `<tr>
    <td class="mono" style="font-size:10px;white-space:nowrap;color:#1565c0;font-weight:700">${fmtFecha(r.fecha)}</td>
    <td style="font-size:11px;color:var(--muted2);max-width:90px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r.redes||'-'}</td>
    <td style="font-size:11px;max-width:140px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r.estab}</td>
    <td class="mono" style="font-size:10px">${r.cod_sismed||'-'}</td>
    <td style="font-size:11px;max-width:200px" title="${r.producto}">${r.producto.length>45?r.producto.slice(0,43)+'…':r.producto}</td>
    <td class="mono" style="text-align:right">${r.noSat.toLocaleString()}</td>
    <td>${indicBadge(r.indicador_ce)}</td>
    <td>${indicBadge(r.indicador_alm)}</td>
    <td class="mono" style="text-align:right;font-size:10px">${r.stock_ce !== null ? r.stock_ce.toLocaleString() : '—'}</td>
    <td class="mono" style="text-align:right;font-size:10px">${r.stock_alm !== null ? r.stock_alm.toLocaleString() : '—'}</td>
    <td><span style="background:${r._diag.color}22;color:${r._diag.color};padding:2px 8px;border-radius:4px;font-size:10px;white-space:nowrap" title="${r._diag.desc}">${r._diag.label}</span></td>
    <td style="font-size:10px;color:var(--muted);max-width:130px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">
      ${r.tiene_match ? `<span title="${r.match_nombre||''}">${r.match_nombre?(r.match_nombre.length>30?r.match_nombre.slice(0,28)+'…':r.match_nombre):'—'}</span>` : '<span style="color:#78909c">⚪ Nuevo</span>'}
    </td>
  </tr>`).join('') : '<tr><td colspan="12" style="text-align:center;padding:40px;color:var(--muted)">Sin resultados para los filtros seleccionados</td></tr>';
}

function sortICI(col) {
  if (iciSortCol===col) iciSortDir*=-1; else { iciSortCol=col; iciSortDir=-1; }
  renderCruceTable();
}
function prevICI(){if(iciPage>1){iciPage--;renderCruceTable();}}
function nextICI(){const tp=Math.ceil(iciGroupedData.length/PAGE);if(iciPage<tp){iciPage++;renderCruceTable();}}

// ═══════════════════════════════════════════════════════════════
// ALERTA PERSISTENTE — productos sin stock en múltiples días
// ═══════════════════════════════════════════════════════════════
function renderPersistentesAlert() {
  const el = document.getElementById('persistentes-list');
  if (!el) return;
  const cruce = getCruceFiltered();
  const pers = detectarPersistentes(cruce);
  if (!pers.length) {
    el.innerHTML = '<p style="text-align:center;color:var(--muted2);padding:20px;font-size:11px">✅ No se detectan productos con desabastecimiento persistente en múltiples días</p>';
    return;
  }
  const badge = (ind) => {
    const c = ind==='DESABASTECIDO'?'#c62828':ind==='SUBSTOCK'?'#e65100':'#78909c';
    return `<span style="background:${c}22;color:${c};padding:1px 6px;border-radius:4px;font-size:9px;font-family:Space Mono,monospace">${ind}</span>`;
  };
  el.innerHTML = pers.slice(0,20).map(p => {
    const diasBadge = p.diasMalos >= 5 ? `<span style="background:#c62828;color:#fff;padding:2px 8px;border-radius:4px;font-size:9px;font-weight:700">⚠ ${p.diasMalos} DÍAS</span>`
      : `<span style="background:#e65100;color:#fff;padding:2px 8px;border-radius:4px;font-size:9px;font-weight:700">${p.diasMalos} días</span>`;
    const indHist = [...new Set(p.indicadores)].map(badge).join(' ');
    return `<div class="a-item crit" style="flex-wrap:wrap;gap:6px">
      <div class="a-dot"></div>
      <div style="flex:1;min-width:0">
        <div class="a-name">${p.producto}</div>
        <div class="a-meta">${p.redes} · ${p.estab}</div>
        <div class="a-meta" style="margin-top:3px">Desde ${new Date(p.primerFechaMala+'T00:00:00').toLocaleDateString('es-PE',{day:'2-digit',month:'short'})} · ${p.totalNoSat.toLocaleString()} u sin atender · ${indHist}</div>
      </div>
      ${diasBadge}
    </div>`;
  }).join('');
}

// ═══════════════════════════════════════════════════════════════
// GRÁFICO POR RED (reemplaza establecimiento en brecha)
// ═══════════════════════════════════════════════════════════════
function renderRedChart(data) {
  if (charts.red) { try { charts.red.destroy(); } catch(e){} }
  const byRed = {};
  data.forEach(r => {
    if (!byRed[r.redes]) byRed[r.redes] = { noSat:0, disponible:0, requerida:0, estabs: new Set() };
    byRed[r.redes].noSat += r.noSat;
    byRed[r.redes].disponible += r.disponible;
    byRed[r.redes].requerida += r.requerida;
    byRed[r.redes].estabs.add(r.estab);
  });
  const sorted = Object.entries(byRed).sort((a,b)=>b[1].noSat-a[1].noSat);
  const canvas = document.getElementById('cRed');
  if (!canvas) return;
  const labels = sorted.map(([k,v]) => `${k.replace('RED ','')}\n(${v.estabs.size} EESS)`);
  charts.red = new Chart(canvas, {
    type: 'bar',
    data: { labels,
      datasets: [
        { label:'Atendido',    data: sorted.map(([,v])=>v.disponible), backgroundColor:'rgba(46,125,50,.75)', borderRadius:5, stack:'s' },
        { label:'Sin atender', data: sorted.map(([,v])=>v.noSat),      backgroundColor:'rgba(198,40,40,.82)', borderRadius:5, stack:'s' }
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false, indexAxis:'y',
      plugins:{
        legend:{display:true,position:'bottom',labels:{color:'#3d5a73',font:{family:'DM Sans',size:11},padding:14,boxWidth:10}},
        tooltip:{ backgroundColor:'#0d2137',titleColor:'#fff',bodyColor:'#90b4c8',padding:11,
          callbacks:{ label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.x.toLocaleString('es-PE')} u` }},
        datalabels:{ display: ctx => ctx.datasetIndex===1, color:'#fff', anchor:'center', align:'center',
          font:{size:9,family:'Space Mono'}, formatter: v => v>0?v.toLocaleString('es-PE'):'' }
      },
      scales:{
        x:{stacked:true,grid:{color:'rgba(21,101,192,.07)'},ticks:{color:'#5a7490',font:{size:9}}},
        y:{stacked:true,grid:{display:false},ticks:{color:'#0d2137',font:{size:10}}}
      }
    }
  });
}

// ═══════════════════════════════════════════════════════════════
// EXPORT CSV CRUCE
// ═══════════════════════════════════════════════════════════════
function exportCruceCSV() { exportCruceXLSX(); }
function exportCruceXLSX() {
  const cruce = iciGroupedData;
  const heads = ['Fecha','Mes','Red','Establecimiento','COD PRE','COD SISMED','Producto','Sin Atender','Requerida','Disponible','Cobertura %','ICI Centros','ICI Almacén','Stock CE','Stock ALM','Diagnóstico','Match ICI'];
  const dataRows = cruce.map(r => {
    const d = diagnosticarCruce(r);
    return [
      r.fecha||'—', mesLabel(r.mesKey), r.redes, r.estab, r.cod_pre, r.cod_sismed, r.producto,
      r.noSat, r.requerida||0, r.disponible||0,
      r.requerida ? +(r.disponible/r.requerida*100).toFixed(1) : 0,
      r.indicador_ce, r.indicador_alm,
      r.stock_ce !== null ? r.stock_ce : '—',
      r.stock_alm !== null ? r.stock_alm : '—',
      d.label.replace(/[\u{1F534}\u{1F7E0}\u{1F7E1}\u{1F535}\u26AA]/gu,'').trim(),
      r.match_nombre||'—'
    ];
  });
  exportXLSX(heads, dataRows, `CruceICI_${new Date().toISOString().slice(0,10)}.xlsx`, {
    title: 'Cruce ICI × Productos No Atendidos',
    subtitle: 'DEMID · DIRESA Callao · Disponibilidad diaria vs demanda no satisfecha · Agrupado por fecha + centro + producto'
  });
}

// ═══════════════════════════════════════════════════════════════
// GAUGE SVG
// ═══════════════════════════════════════════════════════════════
function renderGauge(cob, disp, req) {
  const el = document.getElementById('gauge-wrap');
  if (!el) return;
  const pct   = Math.min(Math.max(cob, 0), 100);
  const color = pct < 10 ? '#c62828' : pct < 50 ? '#d84315' : pct < 80 ? '#e65100' : '#2e7d32';
  const r = 70, cx = 100, cy = 95;
  const startAngle = -Math.PI, endAngle = 0;
  const angle      = startAngle + (pct / 100) * Math.PI;
  const toXY = (a, rr) => ({ x: cx + rr * Math.cos(a), y: cy + rr * Math.sin(a) });
  const arcPath = (r2, a1, a2, color) => {
    const s = toXY(a1, r2), e = toXY(a2, r2);
    const large = (a2 - a1) > Math.PI ? 1 : 0;
    return `<path d="M ${s.x} ${s.y} A ${r2} ${r2} 0 ${large} 1 ${e.x} ${e.y}" fill="none" stroke="${color}" stroke-width="14" stroke-linecap="round"/>`;
  };
  const needle = toXY(angle, r - 8);
  el.innerHTML = `
    <svg class="gauge-svg" viewBox="0 0 200 110" width="200" height="110">
      ${arcPath(r, startAngle, endAngle, '#e4eaf3')}
      ${pct > 0 ? arcPath(r, startAngle, angle, color) : ''}
      <circle cx="${needle.x}" cy="${needle.y}" r="5" fill="${color}"/>
      <text x="${cx}" y="${cy - 12}" class="gauge-label-big" fill="${color}">${pct.toFixed(1)}%</text>
      <text x="${cx}" y="${cy + 8}" class="gauge-label-sub">cobertura global</text>
      <text x="14" y="${cy + 22}" style="font-size:9px;fill:#5a7490;font-family:Space Mono,monospace">0%</text>
      <text x="172" y="${cy + 22}" style="font-size:9px;fill:#5a7490;font-family:Space Mono,monospace">100%</text>
    </svg>
    <div class="gauge-legend">
      <div class="gauge-leg-item"><div class="gauge-leg-dot" style="background:#2e7d32"></div>Atendido: ${disp.toLocaleString('es-PE')} u</div>
      <div class="gauge-leg-item"><div class="gauge-leg-dot" style="background:#c62828"></div>Sin atender: ${(req-disp).toLocaleString('es-PE')} u</div>
    </div>`;
}

// ═══════════════════════════════════════════════════════════════
// GRÁFICO ESTABLECIMIENTOS — cobertura % + brecha
// ═══════════════════════════════════════════════════════════════
function renderEstabChart(data) {
  if (charts.estab) { try { charts.estab.destroy(); } catch(e){} }
  const byEstab = {};
  data.forEach(r => {
    if (!byEstab[r.estab]) byEstab[r.estab] = { noSat:0, disponible:0, requerida:0, red:r.redes };
    byEstab[r.estab].noSat     += r.noSat;
    byEstab[r.estab].disponible+= r.disponible;
    byEstab[r.estab].requerida += r.requerida;
  });
  const sorted = Object.entries(byEstab).sort((a,b)=>b[1].noSat-a[1].noSat).slice(0,12);
  const tag = document.getElementById('tag-estab');
  if (tag) tag.textContent = `Top ${sorted.length} EESS · sin atender`;
  const canvas = document.getElementById('cEstab');
  if (!canvas) return;
  const coberturas = sorted.map(([,v]) => v.requerida ? +(v.disponible/v.requerida*100).toFixed(1) : 0);
  const labels = sorted.map(([k,v]) => `${k.length>22?k.slice(0,20)+'…':k}  ${coberturas[sorted.indexOf(Object.entries(byEstab).find(e=>e[0]===k))]?.toFixed?.(0)||'0'}%`);
  charts.estab = new Chart(canvas, {
    type: 'bar',
    data: { labels: sorted.map(([k])=>k.length>28?k.slice(0,26)+'…':k),
      datasets: [
        { label:'Atendido',     data: sorted.map(([,v]) => v.disponible), backgroundColor:'rgba(46,125,50,.70)', borderRadius:4, stack:'total' },
        { label:'Sin atender',  data: sorted.map(([,v]) => v.noSat),      backgroundColor:'rgba(198,40,40,.82)', borderRadius:4, stack:'total' }
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false, indexAxis:'y',
      plugins: {
        legend:{display:true,position:'bottom',labels:{color:'#3d5a73',font:{family:'DM Sans',size:11},padding:14,boxWidth:10}},
        tooltip:{ backgroundColor:'#0d2137',titleColor:'#fff',bodyColor:'#90b4c8',padding:11,
          callbacks:{
            title: (items) => {
              const idx = items[0].dataIndex;
              const [k,v] = sorted[idx];
              const cob = v.requerida ? (v.disponible/v.requerida*100).toFixed(1) : 0;
              return `${k} · Cob. ${cob}%`;
            },
            label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.x.toLocaleString('es-PE')} u`
          }
        },
        datalabels:{
          display: ctx => ctx.datasetIndex === 1,
          color: '#fff', anchor:'center', align:'center',
          font:{size:9,family:'Space Mono'},
          formatter: (v, ctx) => {
            const [,val] = sorted[ctx.dataIndex];
            const cob = val.requerida ? (val.disponible/val.requerida*100).toFixed(0) : 0;
            return `${cob}%`;
          }
        }
      },
      scales:{
        x:{stacked:true,grid:{color:'rgba(21,101,192,.07)'},ticks:{color:'#5a7490',font:{size:9}}},
        y:{stacked:true,grid:{display:false},ticks:{color:'#0d2137',font:{size:10}}}
      }
    }
  });
}

// ═══════════════════════════════════════════════════════════════
// TOP 10 PRODUCTOS
// ═══════════════════════════════════════════════════════════════
function renderTopProds(data) {
  // Agrupar desde el 1ero del mes: suma por producto en el período activo
  const byProd = {};
  data.forEach(r => {
    if (!byProd[r.producto]) byProd[r.producto] = { noSat: 0, meses: new Set(), estabs: new Set() };
    byProd[r.producto].noSat += r.noSat;
    if (r.mesKey) byProd[r.producto].meses.add(r.mesKey);
    byProd[r.producto].estabs.add(r.estab);
  });
  const top = Object.entries(byProd).sort((a,b)=>b[1].noSat-a[1].noSat).slice(0,10);
  const maxV = top[0]?.[1].noSat || 1;
  const tbody = document.getElementById('top-tbody');
  if (!tbody) return;
  tbody.innerHTML = top.map(([prod, info], i) => {
    const val = info.noSat;
    const pct = Math.round((val/maxV)*100);
    const rankClass = i===0?'r1':i===1?'r2':i===2?'r3':'';
    const estado = val > maxV*.5 ? 'CRÍTICO' : val > maxV*.25 ? 'ALTO' : 'MODERADO';
    const badgeCls = estado==='CRÍTICO'?'badge-red':estado==='ALTO'?'badge-orange':'badge-green';
    const subInfo = `${info.estabs.size} EESS · ${info.meses.size} mes${info.meses.size>1?'es':''}`;
    return `<tr>
      <td class="rank ${rankClass}">${i+1}</td>
      <td title="${prod}"><div class="prod-name">${prod}</div><div style="font-size:9px;color:var(--muted);margin-top:1px">${subInfo}</div></td>
      <td><div class="nosat-bar-wrap"><div class="nosat-bar-track"><div class="nosat-bar-fill" style="width:${pct}%"></div></div><span class="nosat-val">${val.toLocaleString('es-PE')} u</span></div></td>
      <td><span class="badge ${badgeCls}">${estado}</span></td>
    </tr>`;
  }).join('');
}

// ═══════════════════════════════════════════════════════════════
// ALERTAS
// ═══════════════════════════════════════════════════════════════
function renderAlerts(data) {
  const dedup = (rows) => {
    const map = {};
    rows.forEach(r => {
      const fechaKey = r.fecha ? r.fecha.toISOString().slice(0,10) : 'sin-fecha';
      const key = `${r.producto}||${r.estab}||${fechaKey}`;
      if (!map[key] || r.noSat > map[key].noSat) map[key] = r;
    });
    return Object.values(map);
  };
  const crit = dedup(data.filter(r=>r.cobertura===0)).sort((a,b)=>b.noSat-a.noSat).slice(0,15);
  const warn = dedup(data.filter(r=>r.cobertura>0&&r.cobertura<30)).sort((a,b)=>b.noSat-a.noSat).slice(0,8);
  const el = document.getElementById('alerts-list');
  if (!el) return;
  const items = [
    ...crit.map(r=>`<div class="a-item crit"><div class="a-dot"></div><div style="flex:1;min-width:0"><div class="a-name">${r.producto}</div><div class="a-meta">${r.estab}</div><div class="a-meta" style="color:rgba(198,40,40,.85);margin-top:2px">${r.noSat.toLocaleString()} u sin atender · ${r.fecha?r.fecha.toLocaleDateString('es-PE'):''}</div></div><span class="a-tag">0% atendido</span></div>`),
    ...warn.map(r=>`<div class="a-item warn"><div class="a-dot"></div><div style="flex:1;min-width:0"><div class="a-name">${r.producto}</div><div class="a-meta">${r.estab}</div><div class="a-meta" style="color:rgba(216,67,21,.85);margin-top:2px">${r.cobertura.toFixed(1)}% atendido · ${r.noSat.toLocaleString()} u sin cubrir</div></div><span class="a-tag">Parcial</span></div>`)
  ];
  el.innerHTML = items.length
    ? items.join('')
    : '<p style="text-align:center;color:var(--muted);padding:24px;font-size:12px">✅ Sin alertas críticas</p>';
}

// ═══════════════════════════════════════════════════════════════
// POR SERVICIO
// ═══════════════════════════════════════════════════════════════
function renderServicioChart(data) {
  if (charts.svc) { try { charts.svc.destroy(); } catch(e){} }
  const bySvc = {};
  data.forEach(r => { bySvc[r.servicio] = (bySvc[r.servicio]||0) + r.noSat; });
  const top = Object.entries(bySvc).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const canvas = document.getElementById('cServicio');
  if (!canvas) return;
  const colors = ['#c62828','#d84315','#e65100','#1565c0','#1976d2','#2196f3','#00695c','#2e7d32','#4527a0','#546e7a'];
  charts.svc = new Chart(canvas, {
    type:'bar',
    data:{ labels:top.map(([k])=>k.length>30?k.slice(0,28)+'…':k), datasets:[{data:top.map(([,v])=>v),backgroundColor:top.map((_,i)=>colors[i%colors.length]+'cc'),borderRadius:5}]},
    options:{
      responsive:true,maintainAspectRatio:false,indexAxis:'y',
      plugins:{legend:{display:false},tooltip:{backgroundColor:'#0d2137',borderColor:'rgba(21,101,192,.2)',borderWidth:1,titleColor:'#fff',bodyColor:'#90b4c8',padding:11,callbacks:{label:ctx=>` ${ctx.parsed.x.toLocaleString('es-PE')} unidades sin atender`}},datalabels:{display:true,color:'#0d2137',anchor:'end',align:'end',font:{size:9,family:'Space Mono'},formatter:v=>v.toLocaleString('es-PE')}},
      scales:{x:{grid:{color:'rgba(21,101,192,.07)'},ticks:{color:'#5a7490',font:{size:9}}},y:{grid:{display:false},ticks:{color:'#0d2137',font:{size:11}}}}
    }
  });
}

function renderTodayRecords(data) {
  const el = document.getElementById('today-list');
  if (!el) return;
  const { today, recent } = getTodayRecords(data);
  const todayCount = today.length;
  const recentCount = recent.length;
  const summary = todayCount
    ? `<div style="font-size:12px;color:var(--text);margin-bottom:8px"><strong>${todayCount} registros</strong> de hoy · <strong>${recentCount}</strong> registros últimos 7 días.</div>`
    : `<div style="font-size:12px;color:var(--muted);margin-bottom:8px">No hay registros con fecha de hoy.</div>`;
  const items = today.slice(0,10).map(r => `<div class="a-item ${r.cobertura===0?'crit':'warn'}"><div class="a-dot"></div><div style="flex:1;min-width:0"><div class="a-name">${r.producto}</div><div class="a-meta">${r.estab} · ${r.servicio}</div><div class="a-meta" style="margin-top:2px">${r.noSat.toLocaleString()} u sin atender · ${r.fecha.toLocaleDateString('es-PE')}</div></div><span class="a-tag">${r.cobertura===0?'0%':''}${r.cobertura>0&&r.cobertura<100?r.cobertura.toFixed(1)+'%':''}</span></div>`);
  el.innerHTML = summary + (items.length ? items.join('') : '<p style="text-align:center;color:var(--muted);padding:18px;font-size:12px">Sin registros de hoy</p>');
}

let groupedAllData = [];   // todos los grupos sin filtrar
let grpSearch = '', grpRed = '', grpEstab = '', grpMes = '', grpProd = '';

function filterGrouped() {
  const s  = (document.getElementById('g-search')?.value||'').toLowerCase();
  const fr = document.getElementById('g-red')?.value||'';
  const fe = document.getElementById('g-estab')?.value||'';
  const fm = document.getElementById('g-mes')?.value||'';
  groupedData = groupedAllData.filter(r => {
    if (fr && r.redes !== fr) return false;
    if (fe && r.estab !== fe) return false;
    if (fm && r.mesKey !== fm) return false;
    if (s && !r.producto.toLowerCase().includes(s) && !r.estab.toLowerCase().includes(s) && !(r.cod_sismed||'').includes(s)) return false;
    return true;
  });
  groupPage = 1;
  _paintGroupTable();
}

function _paintGroupTable() {
  const el = document.getElementById('group-body');
  const pager = document.getElementById('group-pager');
  const cnt = document.getElementById('group-count');
  if (!el) return;
  const totalPages = Math.ceil(groupedData.length / PAGE) || 1;
  if (groupPage > totalPages) groupPage = 1;
  const slice = groupedData.slice((groupPage-1)*PAGE, groupPage*PAGE);
  if (cnt) cnt.textContent = `${groupedData.length.toLocaleString()} grupos`;
  if (pager) pager.textContent = `Pág. ${groupPage} / ${totalPages}`;
  el.innerHTML = slice.length ? slice.map(r => `<tr>
      <td class="mono" style="font-size:11px;color:var(--muted2)">${r.redes||'-'}</td>
      <td style="font-size:11px;max-width:150px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r.estab}</td>
      <td class="mono" style="font-size:11px;color:var(--muted2)">${r.cod_pre}</td>
      <td class="mono" style="font-size:11px;color:var(--muted2)">${r.cod_sismed||''}</td>
      <td style="font-size:11px;max-width:210px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${r.producto}">${r.producto}</td>
      <td style="font-size:10px;color:var(--muted2)">${r.servicio}</td>
      <td class="mono" style="text-align:right">${r.requerida.toLocaleString()}</td>
      <td class="mono" style="text-align:right">${r.disponible.toLocaleString()}</td>
      <td class="mono" style="text-align:right;font-weight:700;color:${r.noSat>0?'#c62828':'#2e7d32'}">${r.noSat.toLocaleString()}</td>
      <td><span class="pill ${r.cobertura===0?'pill-red':r.cobertura<30?'pill-orange':'pill-green'}">${r.cobertura.toFixed(1)}%</span></td>
    </tr>`).join('') : '<tr><td colspan="10" style="text-align:center;padding:24px;color:var(--muted)">Sin resultados</td></tr>';
}

function renderGroupedSummary(data) {
  // groupRecords agrupa por cod_pre+cod_sismed+servicio sumando cantidades
  groupedAllData = groupRecords(data).sort((a,b) => b.noSat - a.noSat);
  // Poblar filtros del tab Agrupado
  const gRed   = document.getElementById('g-red');
  const gEstab = document.getElementById('g-estab');
  const gMes   = document.getElementById('g-mes');
  if (gRed && gRed.options.length <= 1) {
    [...new Set(groupedAllData.map(r=>r.redes))].sort().forEach(v=>{ const o=document.createElement('option'); o.value=o.textContent=v; gRed.appendChild(o); });
  }
  if (gEstab && gEstab.options.length <= 1) {
    [...new Set(groupedAllData.map(r=>r.estab))].sort().forEach(v=>{ const o=document.createElement('option'); o.value=o.textContent=v; gEstab.appendChild(o); });
  }
  if (gMes && gMes.options.length <= 1) {
    [...new Set(groupedAllData.filter(r=>r.mesKey).map(r=>r.mesKey))].sort().reverse().forEach(k=>{ const o=document.createElement('option'); o.value=k; o.textContent=mesLabel(k); gMes.appendChild(o); });
  }
  groupedData = groupedAllData;
  _paintGroupTable();
}

function prevGroupPage() { if (groupPage > 1) { groupPage--; _paintGroupTable(); } }
function nextGroupPage() { const total = Math.ceil(groupedData.length/PAGE); if (groupPage < total) { groupPage++; _paintGroupTable(); } }

// ═══════════════════════════════════════════════════════════════
// TAB: PRODUCTOS NUEVOS (sin historial ICI)
// Agrupados solo por producto — sin cruzar con centro ni red
// ═══════════════════════════════════════════════════════════════
let npPage = 1;
const NP_PAGE = 50;

function buildNuevosBase() {
  // Todos los registros del período filtrado (filteredData)
  // El filtro de origen ICI se aplica con el select np-origen
  const origen = document.getElementById('np-origen')?.value || 'todos';
  if (origen === 'nuevo') return cruceData.filter(r => !r.tiene_match);
  if (origen === 'ici')   return cruceData.filter(r => r.tiene_match);
  return cruceData; // todos
}

// Agrupa por PRODUCTO (solo) sumando cantidades y recolectando fechas/meses
function agruparNuevosPorProducto(rows) {
  const map = {};
  rows.forEach(r => {
    const k = r.cod_sismed || r.producto;
    if (!map[k]) map[k] = {
      cod_sismed: r.cod_sismed,
      producto:   r.producto,
      tiene_match: r.tiene_match,
      noSat: 0, requerida: 0, disponible: 0,
      fechas: new Set(),
      meses:  new Set(),
      nCentros: new Set(),
    };
    map[k].noSat      += r.noSat;
    map[k].requerida  += r.requerida || 0;
    map[k].disponible += r.disponible || 0;
    if (r.tiene_match) map[k].tiene_match = true;
    if (r.fecha) map[k].fechas.add(r.fecha instanceof Date ? r.fecha.toISOString().slice(0,10) : r.fecha);
    if (r.mesKey) map[k].meses.add(r.mesKey);
    map[k].nCentros.add(r.cod_pre);
  });
  return Object.values(map).map(it => ({
    ...it,
    fechas:   [...it.fechas].sort().reverse(),
    meses:    [...it.meses].sort().reverse(),
    nCentros: it.nCentros.size,
    cobertura: it.requerida ? (it.disponible / it.requerida * 100) : 0,
    ultimaFecha: [...it.fechas].sort().reverse()[0] || '',
    primeraFecha: [...it.fechas].sort()[0] || '',
  })).sort((a,b) => b.ultimaFecha.localeCompare(a.ultimaFecha) || b.noSat - a.noSat);
}

// Agrupa por DÍA sumando todos los productos nuevos de ese día
function agruparNuevosPorDia(rows) {
  const map = {};
  rows.forEach(r => {
    const d = r.fecha instanceof Date ? r.fecha.toISOString().slice(0,10) : (r.fecha || 'sin-fecha');
    if (!map[d]) map[d] = { fecha: d, noSat: 0, requerida: 0, disponible: 0, productos: new Set(), nCentros: new Set(), mesKey: r.mesKey||'' };
    map[d].noSat      += r.noSat;
    map[d].requerida  += r.requerida || 0;
    map[d].disponible += r.disponible || 0;
    map[d].productos.add(r.cod_sismed || r.producto);
    map[d].nCentros.add(r.cod_pre);
  });
  return Object.values(map).map(it => ({
    ...it,
    nProductos: it.productos.size,
    nCentros:   it.nCentros.size,
    cobertura:  it.requerida ? (it.disponible / it.requerida * 100) : 0,
  })).sort((a,b) => b.fecha.localeCompare(a.fecha));
}

// Agrupa por MES
function agruparNuevosPorMes(rows) {
  const map = {};
  rows.forEach(r => {
    const k = r.mesKey || 'sin-mes';
    if (!map[k]) map[k] = { mesKey: k, noSat: 0, requerida: 0, disponible: 0, productos: new Set(), nCentros: new Set(), dias: new Set() };
    map[k].noSat      += r.noSat;
    map[k].requerida  += r.requerida || 0;
    map[k].disponible += r.disponible || 0;
    map[k].productos.add(r.cod_sismed || r.producto);
    map[k].nCentros.add(r.cod_pre);
    if (r.fecha) map[k].dias.add(r.fecha instanceof Date ? r.fecha.toISOString().slice(0,10) : r.fecha);
  });
  return Object.values(map).map(it => ({
    ...it,
    nProductos: it.productos.size,
    nCentros:   it.nCentros.size,
    nDias:      it.dias.size,
    cobertura:  it.requerida ? (it.disponible / it.requerida * 100) : 0,
  })).sort((a,b) => b.mesKey.localeCompare(a.mesKey));
}

function renderNuevosKPIs(base) {
  const el = document.getElementById('np-kpi-grid');
  if (!el) return;
  const totalNoSat   = base.reduce((s,r)=>s+r.noSat,0);
  const uniqProds    = new Set(base.map(r=>r.cod_sismed||r.producto)).size;
  const uniqMeses    = new Set(base.filter(r=>r.mesKey).map(r=>r.mesKey)).size;
  const uniqDias     = new Set(base.filter(r=>r.fecha).map(r=>r.fecha instanceof Date?r.fecha.toISOString().slice(0,10):r.fecha)).size;
  const uniqCentros  = new Set(base.map(r=>r.cod_pre)).size;
  const kpis = [
    { icon:'💊', lbl:'Productos únicos',            val: uniqProds, sub:'códigos distintos en vista', c:'c-violet' },
    { icon:'📦', lbl:'Unidades sin atender',        val: totalNoSat.toLocaleString('es-PE'), sub:'total acumulado', c:'c-red' },
    { icon:'📅', lbl:'Días con registros',           val: uniqDias, sub:'fechas distintas', c:'c-blue' },
   // { icon:'🗓️', lbl:'Meses involucrados',           val: uniqMeses, sub:'períodos distintos', c:'c-cyan' },
    { icon:'🏥', lbl:'Centros que reportan',         val: uniqCentros, sub:'establecimientos', c:'c-orange' },
  ];
  el.innerHTML = kpis.map(k => `<div class="kpi ${k.c}">
    <span class="kpi-icon">${k.icon}</span>
    <div class="kpi-lbl">${k.lbl}</div>
    <div class="kpi-val">${k.val}</div>
    <div class="kpi-sub">${k.sub}</div>
  </div>`).join('');
}

function renderNuevosTable() {
  const el    = document.getElementById('np-body');
  const thead = document.getElementById('np-thead');
  const cnt   = document.getElementById('np-count');
  const pager = document.getElementById('np-pager');
  if (!el) return;

  const base   = buildNuevosBase();
  const vista  = document.getElementById('np-vista')?.value || 'producto';
  const busq   = (document.getElementById('np-search')?.value||'').toLowerCase();
  const filtMs = document.getElementById('np-mes')?.value||'';
  const filtDa = document.getElementById('np-dia')?.value||'';

  // Poblar selects de mes y día desde la base sin filtrar
  const srMes = document.getElementById('np-mes');
  if (srMes && srMes.options.length <= 1) {
    const meses = [...new Set(base.filter(r=>r.mesKey).map(r=>r.mesKey))].sort().reverse();
    meses.forEach(k => { const o=document.createElement('option'); o.value=k; o.textContent=mesLabel(k); srMes.appendChild(o); });
  }
  const srDia = document.getElementById('np-dia');
  if (srDia && srDia.options.length <= 1) {
    const dias = [...new Set(base.filter(r=>r.fecha).map(r=>r.fecha instanceof Date?r.fecha.toISOString().slice(0,10):r.fecha))].sort().reverse();
    dias.forEach(d => { const o=document.createElement('option'); o.value=d; o.textContent=new Date(d+'T00:00:00').toLocaleDateString('es-PE',{day:'2-digit',month:'short',year:'numeric'}); srDia.appendChild(o); });
  }

  // Aplicar filtros previos al agrupado
  const filtered = base.filter(r => {
    const d = r.fecha instanceof Date ? r.fecha.toISOString().slice(0,10) : (r.fecha||'');
    if (filtMs && r.mesKey !== filtMs) return false;
    if (filtDa && d !== filtDa) return false;
    if (busq && !r.producto.toLowerCase().includes(busq) && !(r.cod_sismed||'').includes(busq)) return false;
    return true;
  });

  renderNuevosKPIs(filtered);

  // Agrupar según vista
  let rows, headers, rowFn;
  if (vista === 'producto') {
    rows = agruparNuevosPorProducto(filtered);
    headers = ['COD SISMED','Producto','En ICI','Sin atender','Requerida','Disponible','Cobertura %','Días c/registro','Meses','Última fecha','1ª fecha','EESS reportan'];
    rowFn = r => `<tr>
      <td class="mono" style="font-weight:700;color:var(--accent)">${r.cod_sismed||'—'}</td>
      <td style="max-width:230px" title="${r.producto}">${r.producto.length>40?r.producto.slice(0,38)+'…':r.producto}</td>
      <td style="text-align:center">${r.tiene_match ? '<span style="color:#2e7d32;font-size:11px" title="Con historial ICI">✅</span>' : '<span style="color:#78909c;font-size:11px" title="Sin historial ICI">⚪</span>'}</td>
      <td class="mono" style="text-align:right;font-weight:700;color:#c62828">${r.noSat.toLocaleString()}</td>
      <td class="mono" style="text-align:right;color:var(--muted2)">${r.requerida.toLocaleString()}</td>
      <td class="mono" style="text-align:right;color:#2e7d32">${r.disponible.toLocaleString()}</td>
      <td><span class="pill ${r.cobertura===0?'pill-red':r.cobertura<50?'pill-orange':'pill-green'}">${r.cobertura.toFixed(1)}%</span></td>
      <td class="mono" style="text-align:center">${r.fechas.length}</td>
      <td class="mono" style="text-align:center">${r.meses.length}</td>
      <td class="mono" style="color:var(--accent)">${r.ultimaFecha ? new Date(r.ultimaFecha+'T00:00:00').toLocaleDateString('es-PE',{day:'2-digit',month:'short',year:'numeric'}) : '—'}</td>
      <td class="mono" style="color:var(--muted2)">${r.primeraFecha ? new Date(r.primeraFecha+'T00:00:00').toLocaleDateString('es-PE',{day:'2-digit',month:'short',year:'numeric'}) : '—'}</td>
      <td class="mono" style="text-align:center">${r.nCentros}</td>
    </tr>`;
  } else if (vista === 'dia') {
    rows = agruparNuevosPorDia(filtered);
    headers = ['Fecha','Productos distintos','Sin atender total','Requerida','Disponible','Cobertura %','EESS reportan'];
    rowFn = r => `<tr>
      <td class="mono" style="font-weight:700;color:var(--accent)">${r.fecha!=='sin-fecha'?new Date(r.fecha+'T00:00:00').toLocaleDateString('es-PE',{weekday:'short',day:'2-digit',month:'short',year:'numeric'}):'Sin fecha'}</td>
      <td class="mono" style="text-align:center">${r.nProductos}</td>
      <td class="mono" style="text-align:right;font-weight:700;color:#c62828">${r.noSat.toLocaleString()}</td>
      <td class="mono" style="text-align:right;color:var(--muted2)">${r.requerida.toLocaleString()}</td>
      <td class="mono" style="text-align:right;color:#2e7d32">${r.disponible.toLocaleString()}</td>
      <td><span class="pill ${r.cobertura===0?'pill-red':r.cobertura<50?'pill-orange':'pill-green'}">${r.cobertura.toFixed(1)}%</span></td>
      <td class="mono" style="text-align:center">${r.nCentros}</td>
    </tr>`;
  } else { // mes
    rows = agruparNuevosPorMes(filtered);
    headers = ['Mes','Días con registros','Productos distintos','Sin atender total','Requerida','Disponible','Cobertura %','EESS reportan'];
    rowFn = r => `<tr>
      <td style="font-weight:700;color:var(--accent)">${mesLabel(r.mesKey)}</td>
      <td class="mono" style="text-align:center">${r.nDias}</td>
      <td class="mono" style="text-align:center">${r.nProductos}</td>
      <td class="mono" style="text-align:right;font-weight:700;color:#c62828">${r.noSat.toLocaleString()}</td>
      <td class="mono" style="text-align:right;color:var(--muted2)">${r.requerida.toLocaleString()}</td>
      <td class="mono" style="text-align:right;color:#2e7d32">${r.disponible.toLocaleString()}</td>
      <td><span class="pill ${r.cobertura===0?'pill-red':r.cobertura<50?'pill-orange':'pill-green'}">${r.cobertura.toFixed(1)}%</span></td>
      <td class="mono" style="text-align:center">${r.nCentros}</td>
    </tr>`;
  }

  // Encabezados dinámicos
  if (thead) thead.innerHTML = '<tr>' + headers.map(h=>`<th>${h}</th>`).join('') + '</tr>';

  const total = Math.ceil(rows.length/NP_PAGE)||1;
  if (npPage > total) npPage = 1;
  if (cnt)   cnt.textContent  = `${rows.length.toLocaleString()} registros`;
  if (pager) pager.textContent = `Pág. ${npPage} / ${total}`;

  const slice = rows.slice((npPage-1)*NP_PAGE, npPage*NP_PAGE);
  el.innerHTML = slice.length
    ? slice.map(rowFn).join('')
    : `<tr><td colspan="${headers.length}" style="text-align:center;padding:40px;color:var(--muted)">Sin productos nuevos en el período seleccionado</td></tr>`;
}

function prevNp() { if (npPage>1){npPage--;renderNuevosTable();} }
function nextNp() {
  // recalc total without rebuilding full table — just page forward
  npPage++; renderNuevosTable();
}

function exportNuevosXLSX() {
  const base  = buildNuevosBase();
  const vista = document.getElementById('np-vista')?.value || 'producto';
  const filtMs = document.getElementById('np-mes')?.value||'';
  const filtDa = document.getElementById('np-dia')?.value||'';
  const busq   = (document.getElementById('np-search')?.value||'').toLowerCase();
  const filtered = base.filter(r => {
    const d = r.fecha instanceof Date ? r.fecha.toISOString().slice(0,10) : (r.fecha||'');
    if (filtMs && r.mesKey !== filtMs) return false;
    if (filtDa && d !== filtDa) return false;
    if (busq && !r.producto.toLowerCase().includes(busq) && !(r.cod_sismed||'').includes(busq)) return false;
    return true;
  });

  let headers, rows, filename, opts = {};
  if (vista === 'producto') {
    const agrup = agruparNuevosPorProducto(filtered);
    headers = ['COD SISMED','Producto','En ICI','Sin Atender','Requerida','Disponible','Cobertura %','Días c/Registro','Meses','Última Fecha','Primera Fecha','EESS Reportan'];
    rows = agrup.map(r => [r.cod_sismed||'—', r.producto, r.tiene_match?'Sí':'No', r.noSat, r.requerida, r.disponible, +r.cobertura.toFixed(1), r.fechas.length, r.meses.length, r.ultimaFecha||'—', r.primeraFecha||'—', r.nCentros]);
    filename = `Productos_PorProducto_${new Date().toISOString().slice(0,10)}.xlsx`; opts = { title: 'Resumen por Producto · Vista General', subtitle: 'DEMID · DIRESA Callao · Todos los productos con demanda no atendida' };
  } else if (vista === 'dia') {
    const agrup = agruparNuevosPorDia(filtered);
    headers = ['Fecha','Productos Distintos','Sin Atender','Requerida','Disponible','Cobertura %','EESS Reportan'];
    rows = agrup.map(r => [r.fecha, r.nProductos, r.noSat, r.requerida, r.disponible, +r.cobertura.toFixed(1), r.nCentros]);
    filename = `Productos_PorDia_${new Date().toISOString().slice(0,10)}.xlsx`; opts = { title: 'Resumen por Día · Vista General', subtitle: 'DEMID · DIRESA Callao · Productos agrupados por fecha de registro' };
  } else {
    const agrup = agruparNuevosPorMes(filtered);
    headers = ['Mes','Días con Registros','Productos Distintos','Sin Atender','Requerida','Disponible','Cobertura %','EESS Reportan'];
    rows = agrup.map(r => [mesLabel(r.mesKey), r.nDias, r.nProductos, r.noSat, r.requerida, r.disponible, +r.cobertura.toFixed(1), r.nCentros]);
    filename = `Productos_PorMes_${new Date().toISOString().slice(0,10)}.xlsx`; opts = { title: 'Resumen por Mes · Vista General', subtitle: 'DEMID · DIRESA Callao · Productos agrupados por mes de registro' };
  }
  exportXLSX(headers, rows, filename, opts||{});
}

function exportGroupedCSV() {
  const heads = ['Red','Establecimiento','COD PRE','COD SISMED','Producto','Servicio','Requerida','Disponible','Sin Atender','Cobertura %'];
  const dataRows = groupedData.map(r => [r.redes, r.estab, r.cod_pre, r.cod_sismed, r.producto, r.servicio, r.requerida, r.disponible, r.noSat, +(r.cobertura.toFixed(1))]);
  exportXLSX(heads, dataRows, `ProductosNA_Agrupados_${new Date().toISOString().slice(0,10)}.xlsx`, {
    title: 'Resumen Agrupado · Productos No Atendidos',
    subtitle: 'DEMID · DIRESA Callao · Suma por establecimiento y producto'
  });
}

function exportTodayCSV() {
  const { today } = getTodayRecords(filteredData);
  const heads = ['Red','Establecimiento','COD PRE','COD SISMED','Producto','Servicio','Requerida','Disponible','Sin Atender','Cobertura %','Fecha'];
  const dataRows = today.map(r => [r.redes, r.estab, r.cod_pre, r.cod_sismed, r.producto, r.servicio, r.requerida, r.disponible, r.noSat, +(r.cobertura.toFixed(1)), r.fecha instanceof Date ? r.fecha.toLocaleDateString('es-PE') : (r.fecha||'')]);
  exportXLSX(heads, dataRows, `ProductosNA_Hoy_${new Date().toISOString().slice(0,10)}.xlsx`, {
    title: 'Registros del Día · Productos No Atendidos',
    subtitle: 'DEMID · DIRESA Callao · Registros con fecha de hoy'
  });
}

// ═══════════════════════════════════════════════════════════════
// TABLA DETALLE ORIGINAL
// ═══════════════════════════════════════════════════════════════
function populateFilters() {
  const estabs    = [...new Set(allData.map(r=>r.estab))].sort();
  const servicios = [...new Set(allData.map(r=>r.servicio))].sort();
  const se = document.getElementById('t-estab');
  const ss = document.getElementById('t-servicio');
  if (!se||!ss) return;
  const pe=se.value, ps=ss.value;
  se.innerHTML = '<option value="">Todos los EESS</option>' + estabs.map(e=>`<option value="${e}">${e}</option>`).join('');
  const redes = [...new Set(allData.map(r=>r.redes))].sort();
  const sr = document.getElementById('t-redes');
  if (sr) sr.innerHTML = '<option value="">Todas las redes</option>' + redes.map(r=>`<option value="${r}">${r}</option>`).join('');
  ss.innerHTML = '<option value="">Todos los servicios</option>' + servicios.map(s=>`<option value="${s}">${s}</option>`).join('');
  se.value=pe; ss.value=ps;
}

function filterTable() {
  const s  = (document.getElementById('t-search')?.value||'').toLowerCase();
  const fe = document.getElementById('t-estab')?.value||'';
  const fs = document.getElementById('t-servicio')?.value||'';
  const fc = document.getElementById('t-cob')?.value||'';
  const fr = document.getElementById('t-redes')?.value||'';
  const baseData = (dateGroupMode && dateGroupMode !== 'none') ? dateGroupedData : filteredData;
  let rows = baseData.filter(r => {
    if (fe && r.estab!==fe) return false;
    if (fs && r.servicio!==fs) return false;
    if (fr && r.redes!==fr) return false;
    if (fc==='0' && r.cobertura!==0) return false;
    if (fc==='parcial' && !(r.cobertura>0&&r.cobertura<100)) return false;
    if (s && !r.producto.toLowerCase().includes(s) && !r.estab.toLowerCase().includes(s)) return false;
    return true;
  });
  rows.sort((a,b) => {
    const av=a[sortCol], bv=b[sortCol];
    if (typeof av==='number') return (av-bv)*sortDir;
    return String(av||'').localeCompare(String(bv||''))*sortDir;
  });
  const cnt   = document.getElementById('t-count');
  const pager = document.getElementById('t-pager');
  const tbody = document.getElementById('t-body');
  if (!cnt||!pager||!tbody) return;
  const totalP = Math.ceil(rows.length/PAGE)||1;
  if (page>totalP) page=totalP;
  const slice  = rows.slice((page-1)*PAGE, page*PAGE);
  cnt.textContent   = rows.length.toLocaleString() + ' registros';
  pager.textContent = `Pág. ${page} / ${totalP}`;
  const cobPill = v => {
    if (v===0)  return `<span class="pill pill-red">0%</span>`;
    if (v<30)   return `<span class="pill pill-orange">${v.toFixed(1)}%</span>`;
    if (v<100)  return `<span class="pill pill-yellow">${v.toFixed(1)}%</span>`;
    return `<span class="pill pill-green">${v.toFixed(1)}%</span>`;
  };
  const barNoSat = v => {
    const maxV = Math.max(...baseData.map(r=>r.noSat),1);
    const pct  = Math.min(v/maxV*80,80);
    const col  = v>maxV*.5?'#c62828':v>maxV*.15?'#d84315':'#1565c0';
    return `<div class="prog-row"><div class="prog-bar"><div class="prog-fill" style="width:${pct}px;background:${col}"></div></div><span class="prog-val">${v.toLocaleString()}</span></div>`;
  };
  tbody.innerHTML = slice.length
    ? slice.map(r=>`<tr>
        <td class="mono" style="font-size:11px;color:var(--muted2)">${r.redes||'-'}</td>
        <td class="mono" style="font-size:11px;color:var(--muted2)">${r.cod_pre}</td>
        <td style="font-size:11px;max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r.estab}</td>
        <td class="mono" style="font-size:11px;color:var(--muted2)">${r.cod_sismed||''}</td>
        <td style="font-size:11px;max-width:240px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${r.producto}">${r.producto}</td>
        <td style="font-size:10px;color:var(--muted2)">${r.servicio}</td>
        <td class="mono" style="text-align:right">${r.requerida.toLocaleString()}</td>
        <td class="mono" style="text-align:right">${r.disponible.toLocaleString()}</td>
        <td style="min-width:110px">${barNoSat(r.noSat)}</td>
        <td>${cobPill(r.cobertura)}</td>
        <td class="mono" style="font-size:10px">${r.fecha?r.fecha.toLocaleDateString('es-PE'):'-'}</td>
      </tr>`).join('')
    : '<tr><td colspan="11" style="text-align:center;padding:40px;color:var(--muted)">Sin resultados</td></tr>';
}

function sortTbl(col) {
  if (sortCol===col) sortDir*=-1; else { sortCol=col; sortDir=-1; }
  filterTable();
}
function prevPage(){if(page>1){page--;filterTable();}}
function nextPage(){const tp=Math.ceil(filteredData.length/PAGE);if(page<tp){page++;filterTable();}}

// ═══════════════════════════════════════════════════════════════
// EXPORT CSV ORIGINAL
// ═══════════════════════════════════════════════════════════════
function exportCSV() {
  const heads = ['Red','COD PRE','COD SISMED','Establecimiento','Producto','Servicio','Requerida','Disponible','Sin Atender','Cobertura %','Fecha'];
  const baseData = (dateGroupMode && dateGroupMode !== 'none') ? dateGroupedData : filteredData;
  const dataRows = baseData.map(r => [
    r.redes, r.cod_pre, r.cod_sismed, r.estab, r.producto, r.servicio,
    r.requerida, r.disponible, r.noSat,
    +(r.cobertura?.toFixed?.(1) ?? r.cobertura ?? 0),
    r.fecha instanceof Date ? r.fecha.toLocaleDateString('es-PE') : (r.fecha||'')
  ]);
  exportXLSX(heads, dataRows, `ProductosNA_${currentPeriod}_${new Date().toISOString().slice(0,10)}.xlsx`, {
    title: `Detalle Completo · Período: ${currentPeriod}`,
    subtitle: 'DEMID · DIRESA Callao · Registros individuales de demanda no satisfecha'
  });
}

window.onload = reloadData;
