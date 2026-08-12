/**
 * Parser Producción Licapa + Listado trabajadores.
 * Producción: F=Grupo · H=CI (DNI) · Q=C (jarras/cantidad)
 * Listado: A=DNI · B=Nombres y Apellidos
 */

function cleanText(value) {
  return String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function normalizeDni(value) {
  const digits = String(value ?? "").replace(/\D+/g, "");
  return digits.replace(/^0+/, "") || digits;
}

function toNumber(value) {
  if (value == null || value === "") return 0;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const raw = String(value).trim().replace(/\s/g, "").replace(",", ".");
  const n = Number(raw);
  return Number.isFinite(n) ? n : 0;
}

function excelSerialToDate(serial) {
  if (serial == null || serial === "") return "";
  if (typeof serial === "string" && /^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(serial)) {
    return serial;
  }
  const n = typeof serial === "number" ? serial : Number(serial);
  if (!Number.isFinite(n) || n < 20000) return cleanText(serial);
  const utc = Date.UTC(1899, 11, 30) + Math.round(n) * 86400000;
  const d = new Date(utc);
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const yyyy = d.getUTCFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function normHeader(value) {
  return cleanText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function looksLikeDateLabel(value) {
  const t = cleanText(value);
  if (!t) return false;
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(t)) return true;
  const n = Number(t);
  return Number.isFinite(n) && n >= 40000 && n <= 60000;
}

function detectSchema(headerRow) {
  const headers = headerRow.map((h) => normHeader(h)).filter(Boolean);
  const has = (...needles) => needles.some((n) => headers.includes(n) || headers.some((h) => h.includes(n)));

  const isProduccion =
    has("grupo") &&
    (has("ci") || has("dni")) &&
    (has("jarras") || has("cantidad") || headers.includes("c") || has("huerto") || has("etiqueta"));

  const isTareo =
    has("trabajador") ||
    has("supervisor") ||
    has("macro partida") ||
    has("horas pago") ||
    has("hora inicio") ||
    has("documento");

  return { isProduccion, isTareo, headers };
}

function pickCols(headerRow) {
  const cols = {
    grupo: -1,
    dni: -1,
    jarras: -1,
    huerto: -1,
    variedad: -1,
    apellido: -1,
    nombre: -1,
    fecha: -1
  };

  headerRow.forEach((cell, idx) => {
    const h = normHeader(cell);
    if (!h) return;
    if (h === "grupo") cols.grupo = idx;
    else if (h === "ci") cols.dni = idx;
    else if (h === "dni" && cols.dni < 0) cols.dni = idx;
    else if (h === "documento" && cols.dni < 0) cols.dni = idx;
    else if (h === "jarras" || h === "cantidad") cols.jarras = idx;
    else if (h === "c" && cols.jarras < 0) cols.jarras = idx; // col. Q encabezado "C"
    else if (h === "huerto" || h === "fundo") cols.huerto = idx;
    else if (h === "variedad") cols.variedad = idx;
    else if (h === "apellido") cols.apellido = idx;
    else if (h === "nombre") cols.nombre = idx;
    else if (h === "fecha") cols.fecha = idx;
  });

  // Fallbacks solo si el esquema es de producción (headers conocidos)
  const schema = detectSchema(headerRow);
  if (schema.isProduccion) {
    if (cols.grupo < 0) cols.grupo = 5; // F
    if (cols.dni < 0) cols.dni = 7; // H = CI
    if (cols.jarras < 0) cols.jarras = 16; // Q = C
    if (cols.huerto < 0) cols.huerto = 1;
    if (cols.variedad < 0) cols.variedad = 4;
    if (cols.apellido < 0) cols.apellido = 8;
    if (cols.nombre < 0) cols.nombre = 9;
    if (cols.fecha < 0) cols.fecha = 10;
  }

  return { cols, schema };
}

function cell(row, idx) {
  if (idx == null || idx < 0 || !row) return "";
  return row[idx];
}

function nameFromParts(apellido, nombre) {
  const a = cleanText(apellido);
  const n = cleanText(nombre);
  if ((!a || a === "S/N") && (!n || n === "S/N")) return "";
  if (!a || a === "S/N") return n === "S/N" ? "" : n;
  if (!n || n === "S/N") return a;
  return `${a} ${n}`.trim();
}

/**
 * @returns {Map<string, string>} dni -> nombre
 */
export function parseListadoBuffer(buffer, fileName = "listado.xlsx") {
  if (!window.XLSX?.read) throw new Error("SheetJS no está disponible");
  const workbook = window.XLSX.read(buffer, { type: "array", raw: false });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = window.XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false
  });

  if (!rows.length) {
    return { map: new Map(), fileName, count: 0 };
  }

  const header = rows[0].map((h) => normHeader(h));
  let dniIdx = header.findIndex((h) => h === "dni" || h.includes("documento"));
  let nameIdx = header.findIndex(
    (h) => h.includes("nombre") || h.includes("apellido") || h.includes("trabajador")
  );
  if (dniIdx < 0) dniIdx = 0;
  if (nameIdx < 0) nameIdx = 1;

  const map = new Map();
  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const dni = normalizeDni(row[dniIdx]);
    const name = cleanText(row[nameIdx]);
    if (!dni || !name) continue;
    if (!map.has(dni)) map.set(dni, name);
  }

  return { map, fileName, count: map.size, sheetName };
}

/**
 * Parsea Producción (puede ser muy grande).
 * @param {ArrayBuffer} buffer
 * @param {{ workerMap?: Map<string,string>, onProgress?: (p:number)=>void }} options
 */
export function parseProduccionBuffer(buffer, fileName = "produccion.xlsx", options = {}) {
  if (!window.XLSX?.read) throw new Error("SheetJS no está disponible");
  const workerMap = options.workerMap || new Map();

  options.onProgress?.(5);
  const workbook = window.XLSX.read(buffer, {
    type: "array",
    raw: true,
    cellDates: false,
    cellNF: false,
    cellStyles: false,
    sheetStubs: false
  });
  options.onProgress?.(45);

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const matrix = window.XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: true,
    blankrows: false
  });
  options.onProgress?.(70);

  if (!matrix.length) {
    return {
      fileName,
      sheetName,
      rows: [],
      workers: [],
      grupos: [],
      fechas: [],
      meta: { registros: 0, trabajadores: 0, jarras: 0 }
    };
  }

  const { cols, schema } = pickCols(matrix[0]);

  if (schema.isTareo && !schema.isProduccion) {
    throw new Error(
      "Este Excel parece de Tareo (Reporte_Horas), no de Producción. Sube Produccion_Licapa (columnas Grupo, CI y C/jarras)."
    );
  }
  if (!schema.isProduccion || cols.grupo < 0 || cols.dni < 0 || cols.jarras < 0) {
    throw new Error(
      "No se reconocieron columnas de producción. Se esperan encabezados Grupo (F), CI/DNI (H) y C/jarras (Q)."
    );
  }

  const rawRows = [];
  const byWorker = new Map();
  const gruposSet = new Set();
  const fechasSet = new Set();
  let totalJarras = 0;
  let dateLikeGrupos = 0;

  for (let i = 1; i < matrix.length; i += 1) {
    const row = matrix[i];
    const dni = normalizeDni(cell(row, cols.dni));
    if (!dni) continue;

    let grupo = cleanText(cell(row, cols.grupo)) || "(sin grupo)";
    // Si por error Grupo trae fechas seriales, formatea y luego rechazo abajo
    if (looksLikeDateLabel(grupo) || looksLikeDateLabel(excelSerialToDate(cell(row, cols.grupo)))) {
      dateLikeGrupos += 1;
      grupo = excelSerialToDate(cell(row, cols.grupo)) || grupo;
    }
    const jarras = toNumber(cell(row, cols.jarras));
    const huerto = cleanText(cell(row, cols.huerto));
    const variedad = cleanText(cell(row, cols.variedad));
    const fecha = excelSerialToDate(cell(row, cols.fecha));
    const excelName = nameFromParts(cell(row, cols.apellido), cell(row, cols.nombre));
    const trabajador = workerMap.get(dni) || excelName || `(DNI ${dni})`;

    gruposSet.add(grupo);
    if (fecha) fechasSet.add(fecha);
    totalJarras += jarras;

    rawRows.push({
      dni,
      trabajador,
      grupo,
      jarras,
      huerto,
      variedad,
      fecha,
      excelRow: i + 1
    });

    if (!byWorker.has(dni)) {
      byWorker.set(dni, {
        dni,
        trabajador,
        grupos: new Set(),
        jarras: 0,
        registros: 0,
        variedades: new Set(),
        huertos: new Set()
      });
    }
    const w = byWorker.get(dni);
    w.jarras += jarras;
    w.registros += 1;
    w.grupos.add(grupo);
    if (variedad) w.variedades.add(variedad);
    if (huerto) w.huertos.add(huerto);
    if (!w.trabajador || w.trabajador.startsWith("(DNI")) {
      if (trabajador && !trabajador.startsWith("(DNI")) w.trabajador = trabajador;
    }
  }

  if (rawRows.length && dateLikeGrupos / rawRows.length > 0.6) {
    throw new Error(
      "La columna Grupo viene con fechas. Este no es el Excel de Producción (usa Produccion_Licapa con Grupo LIC …)."
    );
  }

  if (rawRows.length && totalJarras <= 0) {
    throw new Error(
      "Se leyeron filas pero jarras = 0. Revisa que subas Produccion_Licapa (columna Q = C/jarras), no el tareo."
    );
  }

  options.onProgress?.(95);

  const workers = [...byWorker.values()]
    .map((w) => ({
      dni: w.dni,
      trabajador: w.trabajador,
      grupo: [...w.grupos].sort((a, b) => a.localeCompare(b, "es")).join(" · "),
      grupos: [...w.grupos],
      jarras: Math.round(w.jarras * 1000) / 1000,
      registros: w.registros,
      variedad: [...w.variedades].sort((a, b) => a.localeCompare(b, "es")).join(" · "),
      huerto: [...w.huertos].sort((a, b) => a.localeCompare(b, "es")).join(" · ")
    }))
    .sort((a, b) => b.jarras - a.jarras || a.trabajador.localeCompare(b.trabajador, "es"));

  return {
    fileName,
    sheetName,
    rows: rawRows,
    workers,
    grupos: [...gruposSet].sort((a, b) => a.localeCompare(b, "es")),
    fechas: [...fechasSet].sort((a, b) => a.localeCompare(b, "es")),
    meta: {
      registros: rawRows.length,
      trabajadores: workers.length,
      jarras: Math.round(totalJarras * 1000) / 1000
    }
  };
}

/** Re-agrega trabajadores desde filas crudas (tras filtrar). */
export function aggregateWorkers(rows) {
  const byWorker = new Map();
  rows.forEach((row) => {
    const dni = row.dni;
    if (!byWorker.has(dni)) {
      byWorker.set(dni, {
        dni,
        trabajador: row.trabajador,
        grupos: new Set(),
        jarras: 0,
        registros: 0,
        variedades: new Set(),
        huertos: new Set()
      });
    }
    const w = byWorker.get(dni);
    w.jarras += Number(row.jarras) || 0;
    w.registros += 1;
    if (row.grupo) w.grupos.add(row.grupo);
    if (row.variedad) w.variedades.add(row.variedad);
    if (row.huerto) w.huertos.add(row.huerto);
    if (row.trabajador && !String(w.trabajador || "").startsWith("(DNI")) {
      /* keep */
    } else if (row.trabajador) {
      w.trabajador = row.trabajador;
    }
  });

  return [...byWorker.values()]
    .map((w) => ({
      dni: w.dni,
      trabajador: w.trabajador,
      grupo: [...w.grupos].sort((a, b) => a.localeCompare(b, "es")).join(" · "),
      grupos: [...w.grupos],
      jarras: Math.round(w.jarras * 1000) / 1000,
      registros: w.registros,
      variedad: [...w.variedades].sort((a, b) => a.localeCompare(b, "es")).join(" · "),
      huerto: [...w.huertos].sort((a, b) => a.localeCompare(b, "es")).join(" · ")
    }))
    .sort((a, b) => b.jarras - a.jarras || a.trabajador.localeCompare(b.trabajador, "es"));
}
