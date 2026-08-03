/** Validaciones PT Espárrago — avisos globales y por celda */

import { getPesoRango, getPesoRangosForFormato, resolveClienteKey } from "./esparrago-pt-pesos.catalog.js";
import { getValidationConfig, getReglasOrigen } from "./esparrago-pt.config.js";
import {
  getFailMessageFromReglas,
  getLongitudExactaDesdeConfig
} from "../../../../../engine/cartilla-rules.adapter.js";
import {
  cellDisplayValue,
  getCellValidationIssues,
  paintCellValidation,
  parseFlexibleNumber
} from "../../../../../engine/cartilla-cell-validation.js";
import { isPtSapDataColJs, PT_SKIP_SAP_VALIDATION } from "../shared/mp-results-perf.util.js";

const CRITICOS_VACIOS = [9, 10, 36, 37, 49, 50, 53, 58, 59, 60, 61, 62, 64];
/** Peso 01–05 → Excel 59–63 / JS 58–62 */
const MERCADOS_VALIDOS = ["USA", "EUROPA", "ASIA"];
export const LINEA_JS = 26;
export const LINEA_ASP_JS = 50;
export const FECHA_LMR_JS = 51;

/** Fecha LMR mayoritaria (ISO yyyy-mm-dd) de la inspección en revisión. */
let _fechaLmrMayoritariaISO = "";

export function setFechaLmrMayoritaria(iso) {
  _fechaLmrMayoritariaISO = iso || "";
}

export function getFechaLmrMayoritaria() {
  return _fechaLmrMayoritariaISO;
}

function looksLikeDateDisplay(value) {
  const texto = String(value ?? "").trim();
  if (!texto) return false;
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(texto)) return true;
  if (/^\d{4}-\d{2}-\d{2}/.test(texto)) return true;
  if (value instanceof Date && Number.isFinite(value.getTime())) return true;
  if (typeof value === "number" && Number.isFinite(value) && value > 20000 && value < 80000) {
    return true;
  }
  return false;
}

export function resolveLineaAspValue(row) {
  if (!row) return "";
  const lineaAspRaw = row[LINEA_ASP_JS];
  const lineaAsp = cellDisplayValue(lineaAspRaw).trim();
  const linea = cellDisplayValue(row[LINEA_JS]).trim();
  // Si Linea Asp vino como fecha (Excel mal tipado / formateo previo), usar Linea.
  if (lineaAsp && !looksLikeDateDisplay(lineaAspRaw) && !looksLikeDateDisplay(lineaAsp)) {
    return lineaAsp;
  }
  if (linea && !looksLikeDateDisplay(linea)) return linea;
  if (lineaAsp && !looksLikeDateDisplay(lineaAsp)) return lineaAsp;
  return linea || "";
}

export function serialExcelAFecha(serial) {
  if (!serial || Number.isNaN(Number(serial))) return serial;
  const fecha = new Date(Math.round((Number(serial) - 25569) * 86400 * 1000));
  const dia = fecha.getUTCDate().toString().padStart(2, "0");
  const mes = (fecha.getUTCMonth() + 1).toString().padStart(2, "0");
  const anio = fecha.getUTCFullYear();
  return `${dia}/${mes}/${anio}`;
}

/** Normaliza valor Excel/texto a ISO yyyy-mm-dd. */
export function parseFechaToISO(val) {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number" && Number.isFinite(val)) {
    const dmy = serialExcelAFecha(val);
    return dmyToISO(dmy);
  }
  const texto = String(val).trim();
  if (!texto) return "";
  if (/^\d{8}$/.test(texto)) {
    return `${texto.slice(0, 4)}-${texto.slice(4, 6)}-${texto.slice(6, 8)}`;
  }
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {
    const [d, m, y] = texto.split("/");
    return `${y}-${m}-${d}`;
  }
  if (/^\d{2}-\d{2}-\d{4}$/.test(texto)) {
    const [d, m, y] = texto.split("-");
    return `${y}-${m}-${d}`;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) return texto;
  const fecha = Date.parse(texto);
  return Number.isFinite(fecha) ? new Date(fecha).toISOString().slice(0, 10) : "";
}

function dmyToISO(dmy) {
  const partes = String(dmy || "").trim().split("/");
  if (partes.length !== 3) return "";
  return `${partes[2]}-${partes[1]}-${partes[0]}`;
}

export function formatISOToDMY(iso) {
  if (!iso) return "";
  const raw = String(iso).trim();
  // Ya viene como dd/mm/yyyy (valor del select de inspección)
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(raw)) return raw;
  if (!raw.includes("-")) return raw;
  const [y, m, d] = raw.split("-");
  if (!y || !m || !d) return raw;
  return `${d}/${m}/${y}`;
}

export function computeFechaLmrMayoritaria(rows, colLmrJs = FECHA_LMR_JS) {
  const fechas = (rows || []).map((r) => parseFechaToISO(r[colLmrJs])).filter(Boolean);
  const conteo = {};
  fechas.forEach((f) => {
    conteo[f] = (conteo[f] || 0) + 1;
  });
  const sorted = Object.entries(conteo).sort((a, b) => b[1] - a[1]);
  return sorted[0]?.[0] || "";
}

export function formatCellDisplay(val, colIdx, row = null) {
  if (colIdx === LINEA_ASP_JS && row) return resolveLineaAspValue(row);
  if ((colIdx === 3 || colIdx === 46 || colIdx === FECHA_LMR_JS) && typeof val === "number") {
    return serialExcelAFecha(val);
  }
  if (val === null || val === undefined) return "";
  return String(val);
}

function normalizeFechaInspeccion(val) {
  if (typeof val === "number" && Number.isFinite(val)) return serialExcelAFecha(val);
  return String(val ?? "").trim();
}

function getJulianoFromDMY(dmy) {
  const partes = String(dmy || "").trim().split("/");
  if (partes.length !== 3) return null;
  const fechaObj = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
  const inicioAnio = new Date(fechaObj.getFullYear(), 0, 0);
  const unDia = 1000 * 60 * 60 * 24;
  return Math.floor((fechaObj - inicioAnio) / unDia)
    .toString()
    .padStart(3, "0");
}

/** Datos SAP = cols Excel 13–27 y 29–33. Nota Condición (28) puede venir de WULUT y NO es SAP. */
function hasSapData(fila) {
  for (let i = 12; i <= 26; i++) {
    const val = fila[i];
    if (val !== undefined && val !== null && String(val).trim() !== "") return true;
  }
  for (let i = 28; i <= 32; i++) {
    const val = fila[i];
    if (val !== undefined && val !== null && String(val).trim() !== "") return true;
  }
  return false;
}

function defectoFueraRango(fila, colIdx, min, max) {
  const raw = fila[colIdx];
  if (raw === undefined || raw === null || String(raw).trim() === "") return false;
  const n = Number.parseFloat(raw);
  return !Number.isNaN(n) && (n < min || n > max);
}

function calidadIncompleta(fila) {
  for (let i = 77; i <= 132; i++) {
    const val = fila[i];
    if (val === undefined || val === null || String(val).trim() === "") return true;
  }
  return false;
}

function setCellTitle(td, title, additive = false) {
  if (!title) return;
  if (additive && td.title) {
    if (!td.title.includes(title)) td.title = `${td.title} · ${title}`;
    return;
  }
  if (!td.title) {
    td.title = title;
    return;
  }
  if (!td.title.includes(title)) td.title = `${td.title} · ${title}`;
}

export function scanGlobalWarnings(filas) {
  let errorSAP = false;
  let errorDefectos = false;
  let errorCalidad = false;

  filas.forEach((fila) => {
    if (!PT_SKIP_SAP_VALIDATION && !errorSAP && hasSapData(fila)) errorSAP = true;
    if (!errorDefectos && (defectoFueraRango(fila, 40, 0, 20) || defectoFueraRango(fila, 42, 0, 70) || defectoFueraRango(fila, 66, 0, 80))) {
      errorDefectos = true;
    }
    if (!errorCalidad && calidadIncompleta(fila)) errorCalidad = true;
  });

  return { errorSAP, errorDefectos, errorCalidad };
}

export function applyCellValidation(td, fila, colIdx, valor) {
  const valorEfectivo = colIdx === LINEA_ASP_JS ? resolveLineaAspValue(fila) : valor;
  let cfg = null;
  let reglas = null;
  try {
    cfg = getValidationConfig();
    reglas = getReglasOrigen() || cfg?._reglasOrigen || null;
  } catch {
    cfg = null;
    reglas = null;
  }

  const stickyClasses = [...td.classList].filter((cls) => cls.startsWith("agv-pt-sticky-col"));
  const copyClass = td.classList.contains("agv-pt-cell-copy");
  td.className = "";
  stickyClasses.forEach((cls) => td.classList.add(cls));
  if (copyClass) td.classList.add("agv-pt-cell-copy");
  td.title = "";

  const ctx = { row: fila };
  const configIssues =
    PT_SKIP_SAP_VALIDATION && isPtSapDataColJs(colIdx)
      ? []
      : cfg
        ? getCellValidationIssues(colIdx, valorEfectivo, ctx, cfg)
        : [];
  if (configIssues.length) {
    paintCellValidation(td, configIssues, "agv-pt");
  }

  const hasEmptyError = configIssues.some((i) => i.kind === "empty");
  const hasValueError = configIssues.some((i) => i.kind === "value");
  const valStr = cellDisplayValue(valorEfectivo);
  const valUpper = valStr.toUpperCase();
  const valNum = parseFlexibleNumber(valorEfectivo);
  const longitudLoteEsperada = colIdx === 9 && cfg ? getLongitudExactaDesdeConfig(cfg, 9) : null;
  const hasLengthError =
    colIdx === 9 &&
    longitudLoteEsperada != null &&
    valUpper &&
    valUpper.length !== longitudLoteEsperada;
  const mercado = cellDisplayValue(fila[53]).toUpperCase();
  const cliente = cellDisplayValue(fila[37]);
  const formato = cellDisplayValue(fila[49]).toUpperCase().replace(/\s/g, "");
  const embalaje = cellDisplayValue(fila[71]).toUpperCase();
  const presentacion = cellDisplayValue(fila[64]);
  const calibre = cellDisplayValue(fila[36]).toUpperCase();

  const markEmpty = (title) => {
    td.classList.add("agv-pt-cell-error-empty");
    setCellTitle(td, title);
  };

  const markValue = (title, { additive = false } = {}) => {
    td.classList.remove("agv-pt-cell-error-empty");
    td.classList.add("agv-pt-cell-error-value");
    setCellTitle(td, title, additive);
  };

  const failMsg = (col, fallback, extras = {}, tipoFallo = null) =>
    getFailMessageFromReglas(reglas, col, fallback, extras, tipoFallo);

  if (!hasEmptyError && CRITICOS_VACIOS.includes(colIdx) && !valUpper) {
    markEmpty(failMsg(colIdx, "Este campo es obligatorio y está vacío"));
    return;
  }

  if (colIdx === FECHA_LMR_JS) {
    const mayISO = getFechaLmrMayoritaria();
    const cellISO = parseFechaToISO(valorEfectivo);
    if (!hasEmptyError && !cellISO) {
      markEmpty(failMsg(colIdx, "Fecha Actualización LMR obligatoria"));
      return;
    }
    if (!hasValueError && mayISO && cellISO && cellISO !== mayISO) {
      const esperadoDMY = formatISOToDMY(mayISO);
      const detectadoDMY = formatISOToDMY(cellISO) || valStr;
      markValue(
        failMsg(
          colIdx,
          `Fecha LMR debe ser la mayoritaria de la inspección (${esperadoDMY}). Detectado: ${detectadoDMY}`,
          {
            valor: detectadoDMY,
            "valor-esperado": esperadoDMY,
            detectado: detectadoDMY
          },
          "igualdad"
        )
      );
      return;
    }
  }

  if (!hasValueError && colIdx === 40 && valStr !== "" && Number.isFinite(valNum) && (valNum < 0 || valNum > 20)) {
    markValue(failMsg(colIdx, "El valor debe estar entre 0 y 20"));
    return;
  }

  if (!hasValueError && colIdx === 42 && valStr !== "" && Number.isFinite(valNum) && (valNum < 0 || valNum > 70)) {
    markValue(failMsg(colIdx, "El valor debe estar entre 0 y 70"));
    return;
  }

  if (!hasValueError && colIdx === 66 && valStr !== "" && Number.isFinite(valNum) && (valNum < 0 || valNum > 80)) {
    markValue(failMsg(colIdx, "El valor debe estar entre 0 y 80"));
    return;
  }

  if (!hasEmptyError && colIdx >= 77 && colIdx <= 132 && !valUpper) {
    markEmpty(failMsg(colIdx, "Este campo no debe estar vacío (columnas de datos obligatorios)"));
    return;
  }

  const clienteKey = resolveClienteKey(mercado, cliente, calibre);
  const rangosFormato = getPesoRangosForFormato(mercado, cliente, formato, calibre);
  const rango = getPesoRango(mercado, cliente, formato, calibre, presentacion);

  if (colIdx === 49 && clienteKey && !rangosFormato.length) {
    td.classList.add("agv-pt-cell-warn-format");
    td.title = `Formato ${formato} no registrado para ${clienteKey}`;
  }

  if (colIdx >= 58 && colIdx <= 62 && Number.isFinite(valNum) && rango) {
    if (valNum < rango.min || valNum > rango.max) {
      markValue(
        calibre === "SMALL"
          ? `Peso incorrecto: El formato ${formato} con Calibre SMALL debe pesar entre ${rango.min}g y ${rango.max}g (Atado)`
          : `Peso incorrecto: El formato ${formato} para Calibre ${calibre || "Estándar"} debe estar entre ${rango.min}g y ${rango.max}g`
      );
      return;
    }
  }

  if (colIdx === 71 && rango && embalaje !== rango.tipo) {
    markValue(`Error de embalaje: Para ${formato} debe ser ${rango.tipo}`);
    return;
  }

  if (colIdx === 64 && rangosFormato.length) {
    const allowed = rangosFormato
      .map((item) => item.presentacion)
      .filter((pres) => pres && pres !== "no validar");
    if (allowed.length && presentacion) {
      const ok = allowed.some((pres) => presentacion.toUpperCase() === pres.toUpperCase());
      if (!ok) {
        markValue(
          `Presentación incorrecta: Para ${clienteKey} con formato ${formato} debe ser "${allowed.join('" o "')}"`
        );
        return;
      }
    }
  } else if (colIdx === 64 && rango && rango.presentacion !== "no validar") {
    if (presentacion.toUpperCase() !== rango.presentacion.toUpperCase()) {
      markValue(
        `Presentación incorrecta: Para ${clienteKey} con formato ${formato} debe ser "${rango.presentacion}"`
      );
      return;
    }
  }

  if (colIdx === 9 && valUpper && !hasEmptyError) {
    const longitudEsperada = longitudLoteEsperada ?? 13;
    let msgError = "";
    if (!hasLengthError && longitudEsperada != null && valUpper.length !== longitudEsperada) {
      msgError = failMsg(
        colIdx,
        `El lote debe tener exactamente ${longitudEsperada} caracteres (detectados: ${valUpper.length})`,
        {
          valor: valStr,
          texto: valStr,
          lote: valStr,
          detectado: valUpper.length,
          "longitud-detectada": String(valUpper.length)
        },
        "longitud"
      );
    } else if (!hasLengthError) {
      const fechaInspeccionStr = normalizeFechaInspeccion(fila[46]);
      const julianoEsperado = getJulianoFromDMY(fechaInspeccionStr);
      if (julianoEsperado) {
        const julianoEnLote = valUpper.slice(-3);
        if (julianoEnLote !== julianoEsperado) {
          msgError = `Día juliano incorrecto: para fecha de inspección ${fechaInspeccionStr} el lote debe terminar en ${julianoEsperado} (detectado: ${julianoEnLote})`;
        }
      }
    }
    if (msgError) {
      markValue(msgError, { additive: hasLengthError || configIssues.length > 0 });
      return;
    }
  }

  if (colIdx === 10 && !hasValueError && !hasEmptyError && Number.isFinite(valNum) && valNum < 100) {
    markValue(failMsg(colIdx, "La muestra mínima es de 100 unidades"));
    return;
  }

  if (colIdx === 53 && valUpper && !MERCADOS_VALIDOS.includes(valUpper)) {
    markEmpty(failMsg(colIdx, "Mercado no válido. Solo se permite USA, EUROPA o ASIA"));
  }
}
