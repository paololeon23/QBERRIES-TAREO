/** Validaciones Uva MP (MPCUV) — configuration-driven desde rules.json */

import { getValidationConfig } from "./uva-mp.config.js";
import {
  applyReglasCompuestasFila,
  cellDisplayValue,
  findDuplicates,
  getCellValidationIssues,
  indicesToValidate,
  parseFlexibleNumber
} from "../../../../../engine/cartilla-cell-validation.js";
import { parseFlexibleDateToISO } from "../shared/excel-date-format.util.js";

export const EXTRA_COL_SUMA_TONALIDADES = "__suma_tonalidades__";
export const EXTRA_COL_SUMA_CALIBRES = "__suma_calibres__";

export function valorCelda(val) {
  return cellDisplayValue(val);
}

export function celdaVacia(val) {
  return valorCelda(val).trim() === "";
}

export { parseFlexibleNumber };

export function parseExcelDateISO(v) {
  return parseFlexibleDateToISO(v);
}

export function formatISOToDMY(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}-${m}-${y}`;
}

export function formatFechaCelda(val) {
  const iso = parseExcelDateISO(val);
  return iso ? formatISOToDMY(iso) : valorCelda(val);
}

export function limpiarMarcasValidacion(rows) {
  rows.forEach((row) => {
    delete row._errors;
    delete row._errorCols;
    delete row._suma_tonalidades;
    delete row._suma_calibres;
    delete row._sumaTonalidadesError;
    delete row._sumaCalibresError;
  });
}

function compareIsoDates(a, b) {
  if (!a || !b) return null;
  return a.localeCompare(b);
}

function applyUvaMpReglasCruzadas(row, err, cfg) {
  const resumen = cfg.validaciones_resumen || {};
  // Fecha Cosecha SAP (20) no se valida en MP: no llegan datos.

  const ton = resumen.suma_tonalidades;
  if (ton) {
    const from = (ton.desde_excel ?? 56) - 1;
    const to = (ton.hasta_excel ?? 65) - 1;
    const targetJs = (ton.igual_a_excel ?? 11) - 1;
    const tol = ton.tolerancia ?? 0.01;
    let suma = 0;
    for (let i = from; i <= to; i++) {
      const n = parseFlexibleNumber(row[i]);
      suma += Number.isFinite(n) ? n : 0;
    }
    row._suma_tonalidades = suma;
    const esperado = parseFlexibleNumber(row[targetJs]);
    if (!Number.isFinite(esperado) || Math.abs(suma - esperado) > tol) {
      err(
        targetJs,
        `Suma Tonalidades (cols ${ton.desde_excel}-${ton.hasta_excel}) debe coincidir con Cant. Muestra`
      );
      row._sumaTonalidadesError = true;
    }
  }

  const cal = resumen.suma_calibres;
  if (cal) {
    const from = (cal.desde_excel ?? 134) - 1;
    const to = (cal.hasta_excel ?? 138) - 1;
    const targetVal = cal.igual_a_valor ?? 100;
    const tol = cal.tolerancia ?? 0.1;
    let suma = 0;
    for (let i = from; i <= to; i++) {
      const n = parseFlexibleNumber(row[i]);
      suma += Number.isFinite(n) ? n : 0;
    }
    row._suma_calibres = suma;
    if (Math.abs(suma - targetVal) > tol) {
      err(
        from,
        `Suma Calibres (cols ${cal.desde_excel}-${cal.hasta_excel}) debe ser exactamente ${targetVal}`
      );
      row._sumaCalibresError = true;
    }
  }
}

export function ejecutarValidacion(rows, config = null) {
  const cfg = config || getValidationConfig();
  const loteIdx = cfg.validaciones_resumen?.lote?.indice_js ?? 9;

  const lotes = rows.map((r) => valorCelda(r[loteIdx]).trim()).filter(Boolean);
  const lotesDuplicados = findDuplicates(lotes);

  rows.forEach((row) => {
    row._errors = new Set();
    row._errorCols = new Set();

    const err = (colIndex, msg) => {
      row._errors.add(`Columna ${colIndex + 1}: ${msg}`);
      row._errorCols.add(colIndex);
    };

    const ctx = {
      row,
      duplicadosLote: lotesDuplicados,
      normalizeDate: parseExcelDateISO
    };

    indicesToValidate(cfg).forEach((idx) => {
      getCellValidationIssues(idx, row[idx], ctx, cfg).forEach((issue) => {
        const colIdx = issue.colIdx ?? idx;
        err(colIdx, issue.message);
      });
    });

    applyUvaMpReglasCruzadas(row, err, cfg);

    const origen = cfg._reglasOrigen || {};
    applyReglasCompuestasFila(
      row,
      origen,
      (colJs, issue) => err(colJs, issue?.message || "Validación cruzada"),
      {
        normalizeDate: parseExcelDateISO,
        compareDates: compareIsoDates,
        parseNumber: parseFlexibleNumber
      }
    );
  });

  return { lotesDuplicados };
}

export function filaTieneError(row) {
  return (
    (row._errorCols && row._errorCols.size > 0) ||
    row._sumaTonalidadesError ||
    row._sumaCalibresError
  );
}

export function obtenerTituloColumna(c, row) {
  if (!row || !row._errors) return "";
  const prefix = `Columna ${c + 1}: `;
  for (const e of row._errors) {
    if (e.startsWith(prefix)) return e.replace(prefix, "");
  }
  if (row._errorCols?.has(c)) return "Error de validación";
  return "";
}

function extraColErrorMessage(row, extraKey) {
  if (extraKey === EXTRA_COL_SUMA_TONALIDADES && row._sumaTonalidadesError) {
    return "Suma Tonalidades no coincide con Cant. Muestra";
  }
  if (extraKey === EXTRA_COL_SUMA_CALIBRES && row._sumaCalibresError) {
    return "Suma Calibres debe ser exactamente 100";
  }
  return "";
}

export function celdaVaciaObligatoria(c, val, row) {
  if (!celdaVacia(val)) return false;
  return row._errorCols && row._errorCols.has(c);
}

export function celdaValorIncorrecto(c, val, row) {
  if (celdaVacia(val)) return false;
  return row._errorCols && row._errorCols.has(c);
}

export function getCellMeta(row, colJs) {
  if (colJs === EXTRA_COL_SUMA_TONALIDADES) {
    const val =
      row._suma_tonalidades != null && Number.isFinite(row._suma_tonalidades)
        ? row._suma_tonalidades.toFixed(2)
        : "";
    if (row._sumaTonalidadesError) {
      return {
        val,
        cellClass: "agv-mp-cell-error-value",
        title: extraColErrorMessage(row, EXTRA_COL_SUMA_TONALIDADES)
      };
    }
    return { val, cellClass: "", title: "" };
  }

  if (colJs === EXTRA_COL_SUMA_CALIBRES) {
    const val =
      row._suma_calibres != null && Number.isFinite(row._suma_calibres)
        ? row._suma_calibres.toFixed(2)
        : "";
    if (row._sumaCalibresError) {
      return {
        val,
        cellClass: "agv-mp-cell-error-value",
        title: extraColErrorMessage(row, EXTRA_COL_SUMA_CALIBRES)
      };
    }
    return { val, cellClass: "", title: "" };
  }

  const valRaw = row[colJs];
  let val;
  if ([19, 50, 69].includes(colJs)) {
    val = formatFechaCelda(valRaw);
  } else {
    val = valorCelda(valRaw);
  }

  if (celdaVaciaObligatoria(colJs, valRaw, row)) {
    return {
      val,
      cellClass: "agv-mp-cell-error-empty",
      title: obtenerTituloColumna(colJs, row)
    };
  }

  if (celdaValorIncorrecto(colJs, valRaw, row)) {
    return {
      val,
      cellClass: "agv-mp-cell-error-value",
      title: obtenerTituloColumna(colJs, row)
    };
  }

  return { val, cellClass: "", title: "" };
}

export function resolveExtraColumnKey(label) {
  const normalized = String(label || "").trim().toLowerCase();
  if (normalized === "suma tonalidades") return EXTRA_COL_SUMA_TONALIDADES;
  if (normalized === "suma calibres") return EXTRA_COL_SUMA_CALIBRES;
  return label;
}
