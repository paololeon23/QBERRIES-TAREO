/** Validaciones Uva PT — suma tonalidades + año/mes cosecha vs inspección */

import {
  getColCosechaJs,
  getColEmbalajeJs,
  getColInspeccionJs,
  getColLmrJs,
  getSumaTonalidadesConfig
} from "./uva-pt.config.js";
import { cellDisplayValue, parseFlexibleNumber } from "../../../../../engine/cartilla-cell-validation.js";
import {
  collectPtYearMonthMismatch,
  paintPtYearMonthMismatch,
  PT_FECHA_YM_MSG
} from "../shared/pt-fecha-ym.util.js";

const FECHA_COSECHA_SAP_JS = 19;

export function serialExcelAFecha(serial) {
  if (!serial || Number.isNaN(Number(serial))) return serial;
  const fecha = new Date(Math.round((Number(serial) - 25569) * 86400 * 1000));
  const dia = fecha.getUTCDate().toString().padStart(2, "0");
  const mes = (fecha.getUTCMonth() + 1).toString().padStart(2, "0");
  const anio = fecha.getUTCFullYear();
  return `${dia}/${mes}/${anio}`;
}

export function valorCelda(val) {
  return cellDisplayValue(val);
}

export function normalizeDateValue(val) {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number" && Number.isFinite(val)) return serialExcelAFecha(val);
  return valorCelda(val).trim();
}

export { parseFlexibleNumber };

function getSumaIndices() {
  const cfg = getSumaTonalidadesConfig();
  return {
    desdeJs: (cfg.desde_excel ?? 60) - 1,
    hastaJs: (cfg.hasta_excel ?? 69) - 1,
    igualAJs: (cfg.igual_a_excel ?? 11) - 1,
    tolerancia: cfg.tolerancia ?? 0.01
  };
}

export function computeSumaTonalidades(row) {
  const { desdeJs, hastaJs } = getSumaIndices();
  let suma = 0;
  for (let i = desdeJs; i <= hastaJs; i++) {
    const n = parseFlexibleNumber(row[i]);
    if (Number.isFinite(n)) suma += n;
  }
  return suma;
}

export function validateSumaTonalidades(row) {
  const { igualAJs, tolerancia } = getSumaIndices();
  const suma = computeSumaTonalidades(row);
  const esperado = parseFlexibleNumber(row[igualAJs]);
  if (!Number.isFinite(esperado)) {
    return { ok: true, suma, esperado: null, message: null };
  }
  const ok = Math.abs(suma - esperado) <= tolerancia;
  return {
    ok,
    suma,
    esperado,
    message: ok
      ? null
      : `Suma Tonalidades (${suma}) debe coincidir con Cant. Muestra (${esperado})`
  };
}

export function formatCellDisplay(val, colIdx) {
  const dateCols = new Set([
    getColInspeccionJs(),
    getColCosechaJs(),
    getColEmbalajeJs(),
    getColLmrJs()
  ]);
  if (dateCols.has(colIdx) && typeof val === "number") return serialExcelAFecha(val);
  if (val === null || val === undefined) return "";
  return String(val);
}

export function collectRowIncidencias(row) {
  const msgs = [];
  const result = validateSumaTonalidades(row);
  if (result.message) msgs.push(result.message);
  collectPtYearMonthMismatch(row, getColCosechaJs(), getColInspeccionJs(), PT_FECHA_YM_MSG).forEach(
    (i) => {
      if (!msgs.includes(i.message)) msgs.push(i.message);
    }
  );
  collectPtYearMonthMismatch(row, FECHA_COSECHA_SAP_JS, getColInspeccionJs(), PT_FECHA_YM_MSG).forEach(
    (i) => {
      if (!msgs.includes(i.message)) msgs.push(i.message);
    }
  );
  return msgs;
}

export function applyFechaYmCell(td, row, colJs) {
  if (
    paintPtYearMonthMismatch(td, colJs, row, getColCosechaJs(), getColInspeccionJs(), PT_FECHA_YM_MSG)
  ) {
    return true;
  }
  return paintPtYearMonthMismatch(
    td,
    colJs,
    row,
    FECHA_COSECHA_SAP_JS,
    getColInspeccionJs(),
    PT_FECHA_YM_MSG
  );
}

export function applySumaTonalidadesCell(td, row) {
  const result = validateSumaTonalidades(row);
  td.textContent = Number.isFinite(result.suma) ? String(result.suma) : "";
  td.classList.remove("agv-pt-cell-error-value");
  td.title = "";
  if (!result.ok) {
    td.classList.add("agv-pt-cell-error-value");
    td.title = result.message || "";
  }
}
