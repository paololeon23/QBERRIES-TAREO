/** Validaciones Palta MP — configuration-driven desde rules.json */

import {
  getColInoloroJs,
  getColInspeccionJs,
  getColLoteJs,
  getTotalColumnas,
  getValidationConfig
} from "./palta-mp.config.js";
import {
  applyReglasCompuestasFila,
  cellDisplayValue,
  findDuplicates,
  getCellValidationIssues,
  indicesToValidate,
  parseFlexibleNumber
} from "../../../../../engine/cartilla-cell-validation.js";

/** Columna solo frontend: Σ calibres 36–50 vs Cant. muestra (11). */
export const EXTRA_COL_SUMA_CALIBRES = "__suma_calibres__";

export function valorCelda(val) {
  return cellDisplayValue(val);
}

export function celdaVacia(val) {
  return valorCelda(val).trim() === "";
}

export { parseFlexibleNumber };

export function parseExcelDateISO(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    const yyyy = v.getFullYear();
    const mm = String(v.getMonth() + 1).padStart(2, "0");
    const dd = String(v.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }
  const s = valorCelda(v).trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  if (/^\d{8}$/.test(s)) return `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}`;
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
    const [d, m, y] = s.split("/");
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(s)) {
    const [d, m, y] = s.split("-");
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  // Serial Excel típico
  if (/^\d{5}(\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n) && n > 20000 && n < 80000) {
      const fecha = new Date(Math.round((n - 25569) * 86400 * 1000));
      if (!Number.isNaN(fecha.getTime())) {
        const yyyy = fecha.getUTCFullYear();
        const mm = String(fecha.getUTCMonth() + 1).padStart(2, "0");
        const dd = String(fecha.getUTCDate()).padStart(2, "0");
        return `${yyyy}-${mm}-${dd}`;
      }
    }
  }
  const d = Date.parse(s);
  return Number.isFinite(d) ? new Date(d).toISOString().slice(0, 10) : "";
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
    delete row._suma_calibres;
    delete row._suma_calibres_cant;
    delete row._sumaCalibresError;
  });
}

/** Fecha Cosecha SAP (20) no se valida en MP: no llegan datos. */
function applyPaltaMpReglasLegacyFechas(_row, _err) {
  // Intencionalmente vacío.
}

/** Solo UI: calcula Σ 36–50 y si coincide con Cant. muestra (11). */
function attachSumaCalibresFrontend(row, cfg) {
  const cal = cfg?.validaciones_resumen?.suma_calibres;
  const from = (cal?.desde_excel ?? 36) - 1;
  const to = (cal?.hasta_excel ?? 50) - 1;
  const cantJs = (cal?.igual_a_excel ?? 11) - 1;
  let suma = 0;
  for (let i = from; i <= to; i += 1) {
    const n = parseFlexibleNumber(row[i]);
    suma += Number.isFinite(n) ? n : 0;
  }
  const cant = parseFlexibleNumber(row[cantJs]);
  row._suma_calibres = suma;
  row._suma_calibres_cant = Number.isFinite(cant) ? cant : null;
  row._sumaCalibresError =
    !Number.isFinite(cant) || Math.abs(suma - cant) > 0.001;
}

function normalizeRowLength(row, totalCols) {
  if (!Array.isArray(row)) return Array.from({ length: totalCols }, () => "");
  if (row.length >= totalCols) return row;
  const padded = row.slice();
  while (padded.length < totalCols) padded.push("");
  return padded;
}

export function ejecutarValidacion(rows, config = null) {
  const cfg = config || getValidationConfig();
  if (!cfg?._reglasOrigen?.columnas?.length) {
    throw new Error("Palta MP: faltan reglas de validación (palta-mp.rules.json). Recarga el módulo.");
  }

  const totalCols = Number(cfg.total_columnas) || getTotalColumnas();
  const colInoloroJs = getColInoloroJs();
  const loteIdx = cfg.validaciones_resumen?.lote?.indice_js ?? getColLoteJs();
  const fechaInspeccionIdx = cfg.filtro_principal?.indice_js ?? getColInspeccionJs();

  const normalizedRows = (rows || []).map((row) => normalizeRowLength(row, totalCols));
  // Mutar in-place para que el service conserve las mismas refs
  rows.forEach((row, i) => {
    const norm = normalizedRows[i];
    if (row.length < totalCols) {
      for (let c = row.length; c < totalCols; c += 1) row[c] = norm[c] ?? "";
    }
  });

  const lotes = rows.map((r) => valorCelda(r[loteIdx]).trim()).filter(Boolean);
  const lotesDuplicados = findDuplicates(lotes);

  const validationHelpers = {
    parseNumber: parseFlexibleNumber,
    normalizeDate: parseExcelDateISO
  };

  const indices = indicesToValidate(cfg);

  rows.forEach((row) => {
    row._errors = new Set();
    row._errorCols = new Set();

    const err = (colIndex, msg) => {
      const js = Number(colIndex);
      if (!Number.isFinite(js) || js < 0) return;
      row._errors.add(`Columna ${js + 1}: ${msg}`);
      row._errorCols.add(js);
    };

    const ctx = {
      row,
      duplicadosLote: lotesDuplicados,
      normalizeDate: parseExcelDateISO,
      fechaInspeccionIdx
    };

    indices.forEach((idx) => {
      getCellValidationIssues(idx, row[idx], ctx, cfg).forEach((issue) => {
        const colIdx = issue.colIdx ?? idx;
        err(colIdx, issue.message);
      });
    });

    applyReglasCompuestasFila(
      row,
      cfg._reglasOrigen,
      (colIdx, issue) => err(colIdx, issue.message),
      validationHelpers
    );

    applyPaltaMpReglasLegacyFechas(row, err);

    // Inoloro (Excel 69 / JS 68): obligatorio operativo aunque la regla base no lo marque
    if (celdaVacia(row[colInoloroJs])) err(colInoloroJs, "Inoloro obligatorio");

    attachSumaCalibresFrontend(row, cfg);
  });

  return { lotesDuplicados };
}

export function filaTieneError(row) {
  return row._errorCols && row._errorCols.size > 0;
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

export function celdaVaciaObligatoria(c, val, row) {
  if (!celdaVacia(val)) return false;
  return row._errorCols && row._errorCols.has(c);
}

export function celdaValorIncorrecto(c, val, row) {
  if (celdaVacia(val)) return false;
  return row._errorCols && row._errorCols.has(c);
}

export function formatSumaCalibresDisplay(row) {
  const suma = row?._suma_calibres;
  const cant = row?._suma_calibres_cant;
  if (suma == null || !Number.isFinite(suma)) return "";
  const sumaTxt = Number.isInteger(suma) ? String(suma) : suma.toFixed(2);
  if (cant == null || !Number.isFinite(cant)) return `${sumaTxt} (sin Cant. muestra)`;
  const cantTxt = Number.isInteger(cant) ? String(cant) : cant.toFixed(2);
  if (row._sumaCalibresError) return `${sumaTxt} ≠ ${cantTxt}`;
  return `${sumaTxt} = ${cantTxt}`;
}

export function getCellMeta(row, colJs) {
  if (colJs === EXTRA_COL_SUMA_CALIBRES) {
    const val = formatSumaCalibresDisplay(row);
    if (row._sumaCalibresError) {
      return {
        val,
        cellClass: "agv-mp-cell-error-value",
        title: "Suma columnas 36–50 debe coincidir con Cant. muestra (11)"
      };
    }
    return {
      val,
      cellClass: "agv-mp-cell-suma-ok",
      title: "Suma calibres 36–50 = Cant. muestra (11)"
    };
  }

  const valRaw = row[colJs];
  let val;
  if ([19, 63, 64].includes(colJs)) {
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
