/**
 * Formato de fechas Excel/SAP compartido (todos los cultivos).
 * Solo columnas cuyo encabezado parece fecha (o hint Excel validado).
 * Nunca formatea Linea Asp, Formato, Destino, etc. aunque el índice Excel esté mal.
 */

/** Hints Excel 1-based frecuentes SAP (sin LMR=51: en varios PT esa col no es fecha). */
export const DEFAULT_SAP_DATE_COLS_EXCEL = [20, 21, 41];

export function parseFlexibleDateToISO(valor) {
  if (valor == null || valor === "") return "";
  if (typeof valor === "number" && Number.isFinite(valor) && valor > 20000 && valor < 80000) {
    const fecha = new Date(Math.round((valor - 25569) * 86400 * 1000));
    if (Number.isNaN(fecha.getTime())) return "";
    const y = fecha.getUTCFullYear();
    const m = String(fecha.getUTCMonth() + 1).padStart(2, "0");
    const d = String(fecha.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  const texto = String(valor).trim();
  if (!texto) return "";

  if (/^\d{5,6}(\.\d+)?$/.test(texto)) {
    const n = Number(texto);
    if (Number.isFinite(n) && n > 20000 && n < 80000) {
      return parseFlexibleDateToISO(n);
    }
  }

  if (/^\d{8}$/.test(texto)) {
    const y = Number(texto.slice(0, 4));
    const m = Number(texto.slice(4, 6));
    const d = Number(texto.slice(6, 8));
    if (y < 1990 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return "";
    return `${texto.slice(0, 4)}-${texto.slice(4, 6)}-${texto.slice(6, 8)}`;
  }

  if (/^\d{4}-\d{2}-\d{2}/.test(texto)) {
    return texto.slice(0, 10);
  }

  if (/^\d{2}[/-]\d{2}[/-]\d{4}$/.test(texto)) {
    const [d, m, y] = texto.split(/[/-]/);
    return `${y}-${m}-${d}`;
  }

  if (/^\d{2}[/-]\d{2}[/-]\d{2}$/.test(texto)) {
    const [d, m, y] = texto.split(/[/-]/);
    const fullY = Number(y) <= 50 ? `20${y}` : `19${y}`;
    return `${fullY}-${m}-${d}`;
  }

  // No usar Date.parse genérico: convierte textos/códigos no-fecha.
  return "";
}

export function formatISOToDMY(iso) {
  if (!iso) return "";
  const [y, m, d] = String(iso).slice(0, 10).split("-");
  if (!y || !m || !d) return "";
  return `${d}/${m}/${y}`;
}

/**
 * Convierte valor de celda a DD/MM/YYYY si es fecha reconocible.
 * @returns {string|null} null si no debe modificarse
 */
export function formatDateValueToDMY(valor) {
  if (valor == null || String(valor).trim() === "") return null;
  const iso = parseFlexibleDateToISO(valor);
  if (!iso) return null;
  return formatISOToDMY(iso);
}

function normHeader(h) {
  return String(h ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{M}/gu, "")
    .trim();
}

/**
 * Encabezados que NUNCA deben formatearse como fecha
 * (aunque un hint Excel 1-based apunte mal a esa columna).
 */
export function isNonDateHeader(header) {
  const headerNorm = normHeader(header);
  if (!headerNorm) return false;
  if (/\bhora\b/.test(headerNorm) || /^hora\b/.test(headerNorm)) return true;
  if (/\blinea\b/.test(headerNorm)) return true;
  if (/\basp\b/.test(headerNorm)) return true;
  if (/\bformato\b/.test(headerNorm)) return true;
  if (/\bmercado\b/.test(headerNorm)) return true;
  if (/\blote\b/.test(headerNorm)) return true;
  if (/\busuario\b/.test(headerNorm)) return true;
  if (/\bcliente\b/.test(headerNorm)) return true;
  if (/\bdestino\b/.test(headerNorm)) return true;
  if (/\bproductor\b/.test(headerNorm)) return true;
  if (/\bguia\b/.test(headerNorm) || /\bremision\b/.test(headerNorm)) return true;
  if (/\betapa\b/.test(headerNorm)) return true;
  if (/\bcampo\b/.test(headerNorm)) return true;
  if (/\bturno\b/.test(headerNorm)) return true;
  if (/\bfundo\b/.test(headerNorm)) return true;
  if (/\bvariedad\b/.test(headerNorm)) return true;
  if (/\bcalibre\b/.test(headerNorm)) return true;
  if (/\bdefecto\b/.test(headerNorm)) return true;
  if (/\bobserv/.test(headerNorm)) return true;
  if (/\bcodig/.test(headerNorm) && !/\bfecha\b/.test(headerNorm)) return true;
  if (/\bcant\b/.test(headerNorm) || /\bmuestra\b/.test(headerNorm)) {
    if (!/\bfecha\b/.test(headerNorm)) return true;
  }
  if (/\bpeso\b/.test(headerNorm) && !/\bfecha\b/.test(headerNorm)) return true;
  if (/\blongitud\b/.test(headerNorm)) return true;
  if (/\btonalidad\b/.test(headerNorm)) return true;
  if (/\bexportable\b/.test(headerNorm)) return true;
  return false;
}

function headerLooksLikeDate(headerNorm) {
  if (!headerNorm) return false;
  if (isNonDateHeader(headerNorm)) return false;
  return (
    headerNorm.includes("fecha") ||
    headerNorm.includes("date") ||
    /\blmr\b/.test(headerNorm)
  );
}

/** ¿Esta columna JS debe tratarse como fecha según encabezado (+ hint opcional)? */
export function isDateColumnJs(jsIdx, headers = [], excelColsHint = []) {
  const js = Number(jsIdx);
  if (!Number.isFinite(js) || js < 0) return false;
  const headerNorm = normHeader(headers?.[js]);
  if (isNonDateHeader(headerNorm)) return false;
  if (headerLooksLikeDate(headerNorm)) return true;
  const excelCol = js + 1;
  const hinted = (excelColsHint || []).some((c) => Number(c) === excelCol);
  if (!hinted) return false;
  // Hint solo si no hay encabezado legible en contra.
  if (headerNorm && String(headers[js]).trim()) return false;
  return true;
}

/**
 * Índices JS (0-based) de columnas fecha: encabezados con «fecha»/LMR + hints validados.
 * Omite Linea Asp, Formato, Destino, etc. Sin fallback ciego a columnas fijas.
 */
export function resolveDateColumnJsIndexes(headers = [], excelColsHint = []) {
  const set = new Set();
  (excelColsHint || []).forEach((c) => {
    const n = Number(c);
    if (!Number.isFinite(n) || n < 1) return;
    const js = n - 1;
    if (!isDateColumnJs(js, headers, [n])) return;
    set.add(js);
  });
  (headers || []).forEach((h, i) => {
    const n = normHeader(h);
    if (!n) return;
    if (headerLooksLikeDate(n)) set.add(i);
  });
  return [...set];
}

function copyRowPreservingMeta(row) {
  const copy = [...row];
  Object.keys(row).forEach((key) => {
    if (Number.isNaN(Number(key))) copy[key] = row[key];
  });
  return copy;
}

/**
 * Aplica DD/MM/YYYY solo en columnas fecha resueltas por encabezado/hint seguro.
 * Sin fallback a DEFAULT: evita convertir Linea Asp / Destino / etc.
 *
 * @param {unknown[][]} rows
 * @param {unknown[]} [headers]
 * @param {number[]} [excelColsHint] columnas Excel 1-based
 */
export function applyDateDisplayFormatToRows(
  rows = [],
  headers = [],
  excelColsHint = DEFAULT_SAP_DATE_COLS_EXCEL
) {
  const cols = resolveDateColumnJsIndexes(headers, excelColsHint);
  if (!cols.length) {
    return (rows || []).map((row) => (Array.isArray(row) ? copyRowPreservingMeta(row) : row));
  }

  return (rows || []).map((row) => {
    if (!Array.isArray(row)) return row;
    const copy = copyRowPreservingMeta(row);
    cols.forEach((js) => {
      if (js < 0 || js >= copy.length) return;
      if (!isDateColumnJs(js, headers, excelColsHint)) return;
      const formatted = formatDateValueToDMY(copy[js]);
      if (formatted) copy[js] = formatted;
    });
    return copy;
  });
}

/** Mutación in-place de una fila (plagas legacy). */
export function formatRowDateCellsInPlace(row, headers = [], excelColsHint = DEFAULT_SAP_DATE_COLS_EXCEL) {
  if (!Array.isArray(row)) return row;
  const cols = resolveDateColumnJsIndexes(headers, excelColsHint);
  cols.forEach((js) => {
    if (js < 0 || js >= row.length) return;
    if (!isDateColumnJs(js, headers, excelColsHint)) return;
    const formatted = formatDateValueToDMY(row[js]);
    if (formatted) row[js] = formatted;
  });
  return row;
}

/**
 * Para export: formatea solo si la columna es fecha por encabezado (o hint seguro).
 * @returns {string|null} null = no es columna fecha (dejar valor original)
 */
export function formatExportDateIfApplicable(val, jsIdx, headers = [], excelColsHint = []) {
  if (!isDateColumnJs(jsIdx, headers, excelColsHint)) return null;
  if (val == null || String(val).trim() === "") return "";
  const formatted = formatDateValueToDMY(val);
  return formatted != null ? formatted : String(val).trim();
}
