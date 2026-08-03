/** Exportación Excel filtrado — PT Espárrago */

import {
  getExportOrderJs,
  getExportTextColsJs,
  getExportDateColsJs,
  getColumnLabelsByIndex
} from "./esparrago-pt.config.js";
import { serialExcelAFecha, resolveLineaAspValue, LINEA_ASP_JS } from "./esparrago-pt.validation.js";
import {
  cellDisplayValue,
  parseFlexibleNumber
} from "../../../../../engine/cartilla-cell-validation.js";
import { resolvePtEsparragoColumnLabel } from "./esparrago-pt-i18n-labels.js";
import { isDateColumnJs, formatDateValueToDMY } from "../shared/excel-date-format.util.js";

function cellObject(val, jsIdx, textCols) {
  if (val === "") return { v: "", t: "s" };
  if (textCols.has(jsIdx)) return { v: String(val), t: "s" };
  if (typeof val === "number" && Number.isFinite(val)) return { v: val, t: "n" };
  const n = parseFlexibleNumber(val);
  if (Number.isFinite(n)) return { v: n, t: "n" };
  return { v: String(val), t: "s" };
}

function formatExportValue(val, jsIdx, textCols, dateCols, headers) {
  const hintedExcel = [...dateCols].map((js) => js + 1);
  if (dateCols.has(jsIdx) && isDateColumnJs(jsIdx, headers, hintedExcel)) {
    if (typeof val === "number" && Number.isFinite(val)) return serialExcelAFecha(val);
    const asDate = formatDateValueToDMY(val);
    if (asDate) return asDate;
    const txt = cellDisplayValue(val).trim();
    return txt || "";
  }
  if (val === null || val === undefined || cellDisplayValue(val).trim() === "") return "";
  if (textCols.has(jsIdx)) return cellDisplayValue(val);
  const n = parseFlexibleNumber(val);
  if (Number.isFinite(n)) return n;
  return cellDisplayValue(val);
}

export function exportEsparragoPtFiltered({ rows, headers, fechaLabel }) {
  if (!window.XLSX?.utils) return false;

  const exportOrder = getExportOrderJs();
  const textCols = getExportTextColsJs();
  const dateCols = getExportDateColsJs();
  const columnLabelsByIndex = getColumnLabelsByIndex();
  const encabezados = exportOrder.map(
    (idx) => resolvePtEsparragoColumnLabel(idx, headers, columnLabelsByIndex) || `Col ${idx + 1}`
  );

  const wsData = [
    encabezados.map((label) => ({
      v: label || "",
      t: "s",
      s: { font: { bold: true } }
    }))
  ];

  rows.forEach((fila) => {
    wsData.push(
      exportOrder.map((idx) => {
        const raw = idx === LINEA_ASP_JS ? resolveLineaAspValue(fila) : fila[idx];
        return cellObject(
          formatExportValue(raw, idx, textCols, dateCols, headers),
          idx,
          textCols
        );
      })
    );
  });

  const ws = window.XLSX.utils.aoa_to_sheet(wsData);
  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, "Revision_PT_Reordenado");
  const timestamp = Date.now().toString().slice(-4);
  const safeFecha = String(fechaLabel || "export").replace(/[\\/:*?"<>|]/g, "-");
  window.XLSX.writeFile(wb, `Reporte_PT_Esparrago_${safeFecha}_ID${timestamp}.xlsx`);
  return true;
}
