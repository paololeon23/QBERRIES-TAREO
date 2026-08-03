/** Exportación Excel PT Arándano — orden, formato numérico y colores de fila. */

import { COLUMNAS_EXPORT_EXTRA } from "./arandano-pt.config.js";

const DEFAULT_TEXT_COLS = [4, 5, 10, 19, 20, 42, 49, 60];
const DEFAULT_DATE_COLS = [42, 49];

const ROW_FILL = {
  duplicate: "FFCCCC",
  green: "AFD8AF",
  orange: "FF9900"
};

function range(from, to) {
  const out = [];
  for (let c = from; c <= to; c++) out.push(c);
  return out;
}

/** PTHPAR — orden de exportación (columnas Excel 1-based). */
export function buildExportOrderPTHPA() {
  return [
    ...range(1, 67),
    69,
    70,
    71,
    72,
    73,
    74,
    78,
    75,
    76,
    77,
    80,
    79,
    68,
    ...range(81, 87),
    89,
    90,
    91,
    92,
    93,
    94,
    98,
    95,
    96,
    97,
    100,
    99,
    88
  ];
}

/** PTLPAR / PTBPAR — mismo orden que llega el archivo (1…100). */
export function buildExportOrderNatural() {
  return range(1, 100);
}

export const EXPORT_ORDER_BY_CARTILLA = {
  PTHPA: buildExportOrderPTHPA(),
  PTLPA: buildExportOrderNatural(),
  PTBPA: buildExportOrderNatural()
};

export function parseFlexibleNumber(val) {
  if (val === null || val === undefined) return NaN;
  if (typeof val === "number" && Number.isFinite(val)) return val;
  if (typeof val === "object") {
    if (val.v != null) return parseFlexibleNumber(val.v);
    if (val.w != null) return parseFlexibleNumber(val.w);
  }
  const texto = String(val).trim().replace(/\s/g, "");
  if (!texto) return NaN;
  const normal = texto.includes(",") && !texto.includes(".")
    ? texto.replace(",", ".")
    : texto.replace(/,/g, "");
  const n = Number(normal);
  return Number.isFinite(n) ? n : NaN;
}

function valueFromOriginalColumn(row, originalColJs, loadReorder) {
  if (!loadReorder?.length) return row[originalColJs];
  const internalIdx = loadReorder.indexOf(originalColJs);
  return internalIdx >= 0 ? row[internalIdx] : "";
}

function formatTextExportValue(val, excelCol, dateCols, helpers, headers = []) {
  const { formatISOToDMY, parseExcelDateISO } = helpers;
  const jsIdx = excelCol - 1;
  const header = headers[jsIdx] ?? "";
  const looksDate =
    dateCols.has(excelCol) &&
    !/linea|asp|formato|destino|lote|usuario|hora/i.test(String(header));

  if (looksDate || (dateCols.has(excelCol) && /fecha|lmr|date/i.test(String(header)))) {
    const iso = parseExcelDateISO(val);
    return iso ? formatISOToDMY(iso) : val === null || val === undefined ? "" : String(val);
  }

  if (excelCol === 4 && typeof val === "number" && Number.isFinite(val)) {
    const iso = parseExcelDateISO(val);
    return iso ? formatISOToDMY(iso) : String(val);
  }

  return val === null || val === undefined ? "" : String(val);
}

export function formatExportCellValue(val, excelCol, exportCfg, helpers, headers = []) {
  const textCols = new Set(exportCfg?.["columnas-texto"] || DEFAULT_TEXT_COLS);
  const dateCols = new Set(exportCfg?.["columnas-fecha"] || DEFAULT_DATE_COLS);

  if (textCols.has(excelCol)) {
    return formatTextExportValue(val, excelCol, dateCols, helpers, headers);
  }

  if (dateCols.has(excelCol)) {
    return formatTextExportValue(val, excelCol, dateCols, helpers, headers);
  }

  if (val === 0 || val === "0") return 0;

  const n = parseFlexibleNumber(val);
  if (!Number.isNaN(n)) return n;

  if (val === null || val === undefined || val === "") return "";

  return String(val);
}

function cellObject(val, excelCol, exportCfg) {
  const textCols = new Set(exportCfg?.["columnas-texto"] || DEFAULT_TEXT_COLS);

  if (val === "") return { v: "", t: "s" };

  if (textCols.has(excelCol)) {
    return { v: String(val), t: "s" };
  }

  if (typeof val === "number" && Number.isFinite(val)) {
    return { v: val, t: "n" };
  }

  const n = parseFlexibleNumber(val);
  if (!Number.isNaN(n)) return { v: n, t: "n" };

  return { v: String(val), t: "s" };
}

function rowFillColor(row) {
  if (row.__duplicado) return ROW_FILL.duplicate;
  if (row.__mark === "green") return ROW_FILL.green;
  if (row.__mark === "orange") return ROW_FILL.orange;
  return null;
}

function applyFill(cell, fillColor) {
  if (!fillColor) return cell;
  const base =
    cell && typeof cell === "object"
      ? { ...cell }
      : { v: cell ?? "", t: typeof cell === "number" ? "n" : "s" };
  base.s = {
    ...(base.s || {}),
    fill: { patternType: "solid", fgColor: { rgb: fillColor } }
  };
  return base;
}

export function buildPtExportSheetData(rows, cartilla, exportMeta, exportCfg, helpers) {
  const order =
    exportCfg?.["reordenar-columnas-excel"] ||
    EXPORT_ORDER_BY_CARTILLA[cartilla] ||
    buildExportOrderNatural();

  const headerOriginal = exportMeta?.headerOriginal || [];
  const loadReorder = exportMeta?.loadReorder || null;

  const wsData = [];

  const headerCells = [
    ...order.map((excelCol) => headerOriginal[excelCol - 1] ?? ""),
    ...COLUMNAS_EXPORT_EXTRA
  ];
  wsData.push(
    headerCells.map((label) => ({
      v: label || "",
      t: "s",
      s: { font: { bold: true } }
    }))
  );

  rows.forEach((row) => {
    const fillColor = rowFillColor(row);
    const line = order.map((excelCol) => {
      const raw = valueFromOriginalColumn(row, excelCol - 1, loadReorder);
      const formatted = formatExportCellValue(raw, excelCol, exportCfg, helpers, headerOriginal);
      return applyFill(cellObject(formatted, excelCol, exportCfg), fillColor);
    });

    COLUMNAS_EXPORT_EXTRA.forEach(() => {
      line.push(applyFill({ v: "", t: "s" }, fillColor));
    });

    wsData.push(line);
  });

  return wsData;
}

export function exportPtWorkbook({ filename, sheets }) {
  if (!window.XLSX?.utils || !sheets?.length) return false;

  const wb = window.XLSX.utils.book_new();
  sheets.forEach(({ name, data }) => {
    const ws = window.XLSX.utils.aoa_to_sheet(data);
    window.XLSX.utils.book_append_sheet(wb, ws, String(name).slice(0, 31));
  });
  window.XLSX.writeFile(wb, filename);
  return true;
}

export function exportPtFiltered({ rows, cartilla, fechaLabel, exportCfg, exportMeta, helpers }) {
  if (!window.XLSX?.utils) return false;

  const wsData = buildPtExportSheetData(rows, cartilla, exportMeta, exportCfg, helpers);
  const safeFecha = String(fechaLabel).replace(/\//g, "-");
  return exportPtWorkbook({
    filename: `PT_Arandano_${cartilla}_${safeFecha}.xlsx`,
    sheets: [{ name: cartilla, data: wsData }]
  });
}
