/**
 * PT: Fecha cosecha vs fecha de inspección/embalaje — solo año y mes iguales.
 * (En MP sí se exige día exacto.)
 */

import { parseFlexibleDateToISO } from "./excel-date-format.util.js";

export const PT_FECHA_YM_MSG =
  "Año y mes de cosecha deben coincidir con la fecha de inspección";

export const PT_FECHA_YM_MSG_EMBALAJE =
  "Año y mes de cosecha deben coincidir con la fecha de embalaje";

/** @returns {string} YYYY-MM o "" */
export function toYearMonth(valor) {
  const iso = parseFlexibleDateToISO(valor);
  return iso ? iso.slice(0, 7) : "";
}

export function yearMonthEqual(a, b) {
  const ya = toYearMonth(a);
  const yb = toYearMonth(b);
  if (!ya || !yb) return true; // sin ambas fechas no hay desigualdad YM
  return ya === yb;
}

/**
 * Si año/mes no coinciden, marca error en las columnas indicadas (índice JS).
 * @returns {{ colJs: number, message: string }[]}
 */
export function collectPtYearMonthMismatch(row, colCosechaJs, colRefJs, message = PT_FECHA_YM_MSG) {
  if (!row || !Number.isFinite(colCosechaJs) || !Number.isFinite(colRefJs)) return [];
  if (yearMonthEqual(row[colCosechaJs], row[colRefJs])) return [];
  return [
    { colJs: colCosechaJs, message },
    { colJs: colRefJs, message }
  ];
}

export function paintPtYearMonthMismatch(td, colJs, row, colCosechaJs, colRefJs, message = PT_FECHA_YM_MSG) {
  if (colJs !== colCosechaJs && colJs !== colRefJs) return false;
  const issues = collectPtYearMonthMismatch(row, colCosechaJs, colRefJs, message);
  if (!issues.length) return false;
  td.classList.remove("agv-pt-cell-error-empty");
  td.classList.add("agv-pt-cell-error-value");
  td.title = message;
  return true;
}
