/** Exportación Excel filtrado — Uva PT (PTCUV) con colores de fila */

import { getExportOrderJs, getExportTextColsJs } from "./uva-pt.config.js";
import {
  serialExcelAFecha,
  valorCelda,
  parseFlexibleNumber,
  computeSumaTonalidades
} from "./uva-pt.validation.js";
import { isDateColumnJs, formatDateValueToDMY } from "../shared/excel-date-format.util.js";

const ROW_FILL = {
  green: "AFD8AF",
  orange: "FF9900"
};

/** JS: cosecha 51, embalaje 52, inspección 53, LMR 70 — siempre filtrado por encabezado. */
const DATE_COLS_JS = new Set([51, 52, 53, 70]);

function rowFillColor(row) {
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

function cellObject(val, jsIdx, textCols) {
  if (val === "") return { v: "", t: "s" };
  if (textCols.has(jsIdx)) return { v: String(val), t: "s" };
  if (typeof val === "number" && Number.isFinite(val)) return { v: val, t: "n" };
  const n = parseFlexibleNumber(val);
  if (Number.isFinite(n)) return { v: n, t: "n" };
  return { v: String(val), t: "s" };
}

function formatExportValue(val, jsIdx, textCols, headers) {
  const hinted = [...DATE_COLS_JS].map((js) => js + 1);
  if (DATE_COLS_JS.has(jsIdx) && isDateColumnJs(jsIdx, headers, hinted)) {
    if (typeof val === "number") return serialExcelAFecha(val);
    const asDate = formatDateValueToDMY(val);
    if (asDate) return asDate;
  }
  if (val === null || val === undefined || valorCelda(val).trim() === "") return "";
  if (textCols.has(jsIdx)) return valorCelda(val);
  const n = parseFlexibleNumber(val);
  if (Number.isFinite(n)) return n;
  return valorCelda(val);
}

export function exportUvaPtFiltered({ rows, headers, fechaLabel }) {
  if (!window.XLSX?.utils) return false;

  const exportOrder = getExportOrderJs();
  const textCols = getExportTextColsJs();
  const encabezados = [
    ...exportOrder.map((idx) => headers[idx] || `Col ${idx + 1}`),
    "Suma Tonalidades"
  ];

  const wsData = [
    encabezados.map((label) => ({
      v: label || "",
      t: "s",
      s: { font: { bold: true } }
    }))
  ];

  rows.forEach((fila) => {
    const fillColor = rowFillColor(fila);
    const rowCells = exportOrder.map((idx) => {
      const formatted = formatExportValue(fila[idx], idx, textCols, headers);
      return applyFill(cellObject(formatted, idx, textCols), fillColor);
    });
    const suma = computeSumaTonalidades(fila);
    rowCells.push(applyFill(cellObject(suma, -1, textCols), fillColor));
    wsData.push(rowCells);
  });

  const ws = window.XLSX.utils.aoa_to_sheet(wsData);
  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, "Revision_PT_Uva");
  const timestamp = Date.now().toString().slice(-4);
  const safeFecha = String(fechaLabel || "export").replace(/[\\/:*?"<>|]/g, "-");
  window.XLSX.writeFile(wb, `Reporte_PT_Uva_${safeFecha}_ID${timestamp}.xlsx`);
  return true;
}
