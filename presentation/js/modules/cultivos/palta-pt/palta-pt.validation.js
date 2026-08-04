/** Validaciones Palta PT — configuration-driven desde rules.json */

import {
  getColEmbalajeJs,
  getColCosechaJs,
  getColTrazabilidadJs,
  getValidationConfig
} from "./palta-pt.config.js";
import { extraerTrazabilidad } from "../arandano-pt/arandano-pt.catalogs.js";
import {
  applyReglasCompuestasFila,
  cellDisplayValue,
  getCellValidationIssues,
  indicesToValidate,
  paintCellValidation,
  parseFlexibleNumber
} from "../../../../../engine/cartilla-cell-validation.js";
import { isPtSapDataColJs, PT_SKIP_SAP_VALIDATION } from "../shared/mp-results-perf.util.js";
import {
  paintPtYearMonthMismatch,
  PT_FECHA_YM_MSG_EMBALAJE
} from "../shared/pt-fecha-ym.util.js";
import { parseFlexibleDateToISO } from "../shared/excel-date-format.util.js";

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

export function getJulianoFromDMY(dmy) {
  const partes = String(dmy || "").trim().split("/");
  if (partes.length !== 3) return null;
  const fechaObj = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
  const inicioAnio = new Date(fechaObj.getFullYear(), 0, 0);
  const unDia = 1000 * 60 * 60 * 24;
  return Math.floor((fechaObj - inicioAnio) / unDia)
    .toString()
    .padStart(3, "0");
}

function compareDatesDMY(a, b) {
  const pa = String(a || "").split("/");
  const pb = String(b || "").split("/");
  if (pa.length !== 3 || pb.length !== 3) return null;
  const da = new Date(Number(pa[2]), Number(pa[1]) - 1, Number(pa[0]));
  const db = new Date(Number(pb[2]), Number(pb[1]) - 1, Number(pb[0]));
  if (Number.isNaN(da.getTime()) || Number.isNaN(db.getTime())) return null;
  return da.getTime() - db.getTime();
}

export function deriveEtapaCampo(trazCode) {
  const traz = extraerTrazabilidad(String(trazCode ?? ""));
  if (!traz) return "";
  return `${traz.sector.etapa}-${traz.sector.campo}`;
}

export function formatCellDisplay(val, colIdx) {
  if ((colIdx === 3 || colIdx === getColCosechaJs() || colIdx === getColEmbalajeJs()) && typeof val === "number") {
    return serialExcelAFecha(val);
  }
  if (val === null || val === undefined) return "";
  return String(val);
}

function buildValidationContext(row, fechaEmbalaje, config) {
  return {
    row,
    normalizeDate: normalizeDateValue,
    fechaEmbalaje,
    config
  };
}

function collectTrazabilidadIssues(row, fechaEmbalaje) {
  const issues = [];
  const trazIdx = getColTrazabilidadJs();
  const traz = valorCelda(row[trazIdx]);
  if (!traz) return issues;
  const trazObj = extraerTrazabilidad(traz);
  const embalaje = fechaEmbalaje || normalizeDateValue(row[getColEmbalajeJs()]);
  const julianoEmb = getJulianoFromDMY(embalaje);
  if (trazObj && julianoEmb && trazObj.juliano !== julianoEmb) {
    issues.push({
      kind: "value",
      message: `Juliano trazabilidad (${trazObj.juliano}) ≠ embalaje (${julianoEmb})`,
      colIdx: trazIdx
    });
  }
  return issues;
}

function collectAllIssues(row, fechaEmbalaje, config = null) {
  const cfg = config || getValidationConfig();
  const ctx = buildValidationContext(row, fechaEmbalaje, cfg);
  const incidencias = [];

  indicesToValidate(cfg).forEach((idx) => {
    if (PT_SKIP_SAP_VALIDATION && isPtSapDataColJs(idx)) return;
    getCellValidationIssues(idx, row[idx], ctx, cfg).forEach((issue) => {
      incidencias.push(issue.message);
    });
  });

  applyReglasCompuestasFila(
    row,
    cfg._reglasOrigen,
    (_colIdx, issue) => incidencias.push(issue.message),
    {
      parseNumber: parseFlexibleNumber,
      normalizeDate: (v) => parseFlexibleDateToISO(v) || normalizeDateValue(v),
      compareDates: compareDatesDMY
    }
  );

  collectTrazabilidadIssues(row, fechaEmbalaje).forEach((issue) => {
    incidencias.push(issue.message);
  });

  return incidencias;
}

export function collectRowIncidencias(row, fechaEmbalaje, config = null) {
  return collectAllIssues(row, fechaEmbalaje, config);
}

export function applyCellValidation(td, fila, colIdx, valor, fechaEmbalaje, config = null) {
  const cfg = config || getValidationConfig();
  const ctx = buildValidationContext(fila, fechaEmbalaje, cfg);

  const stickyClasses = [...td.classList].filter((cls) => cls.startsWith("agv-pt-sticky-col"));
  const copyClass = td.classList.contains("agv-pt-cell-copy");
  td.className = "";
  stickyClasses.forEach((cls) => td.classList.add(cls));
  if (copyClass) td.classList.add("agv-pt-cell-copy");
  td.title = "";

  const issues =
    PT_SKIP_SAP_VALIDATION && isPtSapDataColJs(colIdx)
      ? []
      : getCellValidationIssues(colIdx, valor, ctx, cfg);
  if (colIdx === getColTrazabilidadJs()) {
    issues.push(...collectTrazabilidadIssues(fila, fechaEmbalaje));
  }
  if (issues.length) {
    paintCellValidation(td, issues, "agv-pt");
    return;
  }

  if (colIdx === getColCosechaJs() || colIdx === getColEmbalajeJs() || colIdx === 19) {
    paintPtYearMonthMismatch(
      td,
      colIdx,
      fila,
      getColCosechaJs(),
      getColEmbalajeJs(),
      PT_FECHA_YM_MSG_EMBALAJE
    );
    if (colIdx === getColEmbalajeJs() || colIdx === 19) {
      paintPtYearMonthMismatch(td, colIdx, fila, 19, getColEmbalajeJs(), PT_FECHA_YM_MSG_EMBALAJE);
    }
  }
}
