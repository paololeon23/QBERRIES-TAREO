/**
 * Análisis gerencial de cartilla — compartido (MP / PT / Plagas).
 * Semáforo + conformidad + causa + lotes, sin tocar la lógica de validación.
 *
 * Política UI: bloque Excel 13–33 (datos SAP + Nota Condición) → solo pintura roja
 * en la tabla. Nunca se listan en el panel (ni «Falta datos SAP», ni «Vacío en: …»
 * de esas columnas). El resto de la cartilla sí aparece en el análisis.
 */

/** Nunca listar / agrupar el bloque SAP en el panel (todos los cultivos). */
const SHOW_SAP_MISSING_LABEL = false;

/**
 * Omitir del panel solo vacíos/faltantes del bloque SAP.
 * Errores de valor (ej. Fecha Cosecha ≠ Fecha Inspección) SÍ se listan.
 */
function shouldOmitSapFromAnalysis(column, colNum, cause, t) {
  if (SHOW_SAP_MISSING_LABEL) return false;
  if (isSapMissingCause(cause, t)) return true;
  // Desigualdad / valor incorrecto de fechas: siempre al panel.
  const low = String(cause || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{M}/gu, "");
  if (
    low.includes("debe ser igual") ||
    low.includes("igual a fecha") ||
    low.includes("año y mes") ||
    low.includes("anio y mes") ||
    low.includes("≠")
  ) {
    return false;
  }
  if (!isMissingDataCause(cause, t)) return false;
  const n = Number(colNum);
  if (Number.isFinite(n) && n >= 1) return isSapZoneColNum(n);
  const c = String(column || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{M}/gu, "")
    .trim();
  if (!c || c === "—" || c === "-") return false;
  if (/\(sap\)/.test(c) || /\bsap\b/.test(c)) return true;
  return (
    /\bproductor\b/.test(c) ||
    /\bsociedad\b/.test(c) ||
    /\balmacen\b/.test(c) ||
    /\bguia\b/.test(c) ||
    /\bremision\b/.test(c) ||
    /\btecnologia\b/.test(c) ||
    /\bembalado\b/.test(c) ||
    /\bnota condicion\b/.test(c) ||
    /\btipo de formato\b/.test(c) ||
    /\betiqueta\b/.test(c) ||
    /\bjaba\b/.test(c) ||
    /\bviaje\b/.test(c) ||
    /\bpeso bruto\b/.test(c)
  );
}

function defaultHtmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function isSapMissingCause(cause, t) {
  const c = String(cause || "").trim().toLowerCase();
  const sap = String(t?.("cartillaAnalysis.missingSap") || "Falta datos SAP")
    .trim()
    .toLowerCase();
  return (
    c === sap ||
    c.includes("falta dato sap") ||
    c.includes("falta datos sap") ||
    c.includes("faltan datos sap") ||
    c.includes("missing sap") ||
    c.includes("donnée sap") ||
    c.includes("données sap") ||
    c.includes("缺少 sap")
  );
}

function isWeakCause(cause, t) {
  const c = String(cause || "").toLowerCase();
  // «Falta datos SAP» es causa principal (no genérica).
  if (isSapMissingCause(c, t)) return false;
  return (
    c.includes("obligatorio") ||
    c.includes("desviación de validación") ||
    c.includes("desviacion de validacion")
  );
}

/** Solo «Desviación de validación» / equivalentes: no sirve como causa para listar el lote. */
function isGenericOnlyCause(cause, t) {
  const c = String(cause || "").trim().toLowerCase();
  if (!c) return true;
  if (isSapMissingCause(c, t)) return false;
  const generic = String(t?.("cartillaAnalysis.genericError") || "Desviación de validación")
    .trim()
    .toLowerCase();
  return (
    c === generic ||
    c === "desviación de validación" ||
    c === "desviacion de validacion" ||
    c === "validation deviation" ||
    c === "écart de validation" ||
    c === "校验偏差"
  );
}

/**
 * Bloque Excel 13–33 (JS 12–32): Productor…Peso Bruto + Nota Condición.
 * Vacíos → mensaje por columna (no «Falta datos SAP»). Errores de valor/formato tal cual.
 */
function isSapZoneColNum(colNum) {
  const n = Number(colNum);
  return Number.isFinite(n) && n >= 13 && n <= 33;
}

/** Columnas del bloque SAP (por nombre, por si falta colNum). */
function isSapColumnName(column) {
  const c = String(column || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{M}/gu, "")
    .trim();
  if (/\(sap\)/.test(c) || /\bsap\b/.test(c)) return true;
  // «Calibre» del bloque SAP (JS 22). No «Calibre 18/20/26…» de calidad MP.
  if (/^calibre(\s*\(.*\))?$/.test(c)) return true;
  return (
    /\bproductor\b/.test(c) ||
    /\bsociedad\b/.test(c) ||
    /\bcentro\b/.test(c) ||
    /\balmacen\b/.test(c) ||
    /\bmaterial\b/.test(c) ||
    /\bguia\b/.test(c) ||
    /\bremision\b/.test(c) ||
    /\betapa\b/.test(c) ||
    /\bcampo\b/.test(c) ||
    /\bturno\b/.test(c) ||
    /\bfundo\b/.test(c) ||
    /\bvariedad\b/.test(c) ||
    /\bcosecha\b/.test(c) ||
    /\bproduccion\b/.test(c) ||
    /\btecnologia\b/.test(c) ||
    /\bembalado\b/.test(c) ||
    /\bcategoria\b/.test(c) ||
    /\blinea\b/.test(c) ||
    /\bnota condicion\b/.test(c) ||
    /\bformato\b/.test(c) ||
    /\betiqueta\b/.test(c) ||
    /\bjaba\b/.test(c) ||
    /\bviaje\b/.test(c) ||
    /\bpeso bruto\b/.test(c)
  );
}

/** ¿El mensaje indica celda vacía / obligatorio (no un valor incorrecto)? */
function isMissingDataCause(cause, t) {
  const raw = String(cause || "").trim();
  if (!raw) return true;
  if (isSapMissingCause(raw, t)) return true;
  const low = raw
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{M}/gu, "");
  if (low.includes("obligatorio")) return true;
  if (low.includes("sin dato")) return true;
  if (low.includes("falta dato") || low.includes("faltan dato")) return true;
  if (low.includes("campo vacio") || low === "vacio" || low.includes("vacio")) return true;
  // Igualdad / rango / formato / LMR → NO es “falta SAP”.
  if (low.includes("debe ser") || low.includes("igual")) return false;
  if (low.includes("lmr") || low.includes("mayoritaria")) return false;
  if (low.includes("rango") || low.includes("entre ")) return false;
  if (low.includes("no debe") || low.includes("prohibid")) return false;
  if (low.includes("formato") || low.includes("longitud")) return false;
  return false;
}

function isSapFailure(column, colNum, cause, t) {
  if (isSapMissingCause(cause, t)) return true;
  if (!isMissingDataCause(cause, t)) return false;
  const n = Number(colNum);
  // Con índice fiable: solo el bloque SAP Excel 13–33.
  if (Number.isFinite(n) && n >= 1) {
    return isSapZoneColNum(n);
  }
  // Sin colNum: fallback por nombre canónico SAP.
  return isSapColumnName(column);
}

/** Vacío en zona Excel 13–33 → mensaje por columna (nunca «Falta datos SAP» en panel). */
function promoteCause(cause, t, column = "", colNum = null) {
  const raw = String(cause || "").trim();
  const low = raw.toLowerCase();
  // No mezclar duplicados u otros mensajes no-SAP.
  if (low.includes("duplic") || low.includes("lote duplic")) return raw;
  if (isSapFailure(column, colNum, raw, t)) {
    if (SHOW_SAP_MISSING_LABEL) return t("cartillaAnalysis.missingSap");
    const col = String(column || "").trim();
    if (col && !/^sap$/i.test(col)) return t("cartillaAnalysis.noDataIn", { column: col });
    if (raw && !isSapMissingCause(raw, t)) return raw;
    return t("cartillaAnalysis.genericError");
  }
  if (!raw) return t("cartillaAnalysis.genericError");
  return raw;
}

function cellHasDisplayData(row, colJs) {
  if (!row || !Number.isFinite(colJs) || colJs < 0) return false;
  const raw = row[colJs];
  if (raw == null) return false;
  return String(raw).trim() !== "";
}

/** T° Ambiente / T° Pulpa / temperatura → un solo mensaje corto. */
function isTempColumn(column) {
  const c = String(column || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{M}/gu, "");
  return (
    /\bambiente\b/.test(c) ||
    /\bpulpa\b/.test(c) ||
    /\btemperatura\b/.test(c) ||
    /^t\s*[°º]/.test(c) ||
    /^t\s*amb/.test(c)
  );
}

function isDuplicateCauseText(text) {
  const e = String(text || "").toLowerCase();
  return e.includes("duplic");
}

/** Fecha LMR distinta a la mayoritaria → no spamear fila por fila. */
function isLmrMajorityCauseText(text) {
  const e = String(text || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{M}/gu, "");
  return e.includes("lmr") && (e.includes("mayoritaria") || e.includes("majority") || e.includes("majoritaire"));
}

function formatPreciseErrorText(pair, row, t) {
  const col = String(pair.column || "").trim() || "Col";
  const cause = String(pair.cause || "").trim();
  if (pair.sapMissing || isSapMissingCause(cause, t)) {
    if (SHOW_SAP_MISSING_LABEL) return t("cartillaAnalysis.missingSap");
    if (col && !/^sap$/i.test(col)) return t("cartillaAnalysis.noDataIn", { column: col });
  }
  if (isDuplicateCauseText(cause)) {
    return t("plagasArandano.duplicateLots");
  }
  if (isTempColumn(col)) {
    const js = Number.isFinite(pair.colNum) ? Number(pair.colNum) - 1 : NaN;
    const empty = Number.isFinite(js) ? !cellHasDisplayData(row, js) : true;
    // Valor fuera de rango / inválido: mensaje concreto; vacío: «Error temperatura».
    if (!empty && cause && !isWeakCause(cause, t) && !isGenericOnlyCause(cause, t) && !isMissingDataCause(cause, t)) {
      return `${col}: ${cause}`;
    }
    return t("cartillaAnalysis.missingTemps");
  }
  const js = Number.isFinite(pair.colNum) ? Number(pair.colNum) - 1 : NaN;
  const empty = Number.isFinite(js) ? !cellHasDisplayData(row, js) : false;
  if (empty) {
    return t("cartillaAnalysis.noDataIn", { column: col });
  }
  // Celda con valor: nunca «vacío» → «Error en …»
  if (cause && !isWeakCause(cause, t) && !isGenericOnlyCause(cause, t) && !isMissingDataCause(cause, t)) {
    return t("cartillaAnalysis.errorInColumn", { column: col, detail: cause });
  }
  if (cause && !isMissingDataCause(cause, t)) {
    return t("cartillaAnalysis.errorInColumn", { column: col, detail: cause });
  }
  return t("cartillaAnalysis.errorFields", { columns: col });
}

/**
 * Un solo texto no-SAP por ID/Lote (no listar Ambiente + Pulpa por separado).
 */
function summarizeNonSapErrorText(pairs, row, t) {
  // Seguridad: aunque un par SAP se filtrara mal arriba, no listarlo aquí.
  const nonSap = pairs.filter(
    (p) => !p.sapMissing && !shouldOmitSapFromAnalysis(p.column, p.colNum, p.cause, t)
  );
  if (!nonSap.length) return "";

  const hasTemp = nonSap.some((p) => isTempColumn(p.column));
  const hasDup = nonSap.some((p) => isDuplicateCauseText(p.cause));
  const otherPairs = nonSap.filter(
    (p) => !isTempColumn(p.column) && !isDuplicateCauseText(p.cause)
  );

  // Duplicados: un mensaje corto (se unen por lote más abajo).
  if (hasDup && !otherPairs.length && !hasTemp) {
    return t("plagasArandano.duplicateLots");
  }

  if (hasTemp && !otherPairs.length && !hasDup) {
    return t("cartillaAnalysis.missingTemps");
  }

  const otherLabels = [];
  otherPairs.forEach((p) => {
    const label = String(p.column || "").trim();
    if (label && !otherLabels.includes(label)) otherLabels.push(label);
  });

  const pairIsEmpty = (p) => {
    const js = Number.isFinite(p.colNum) ? Number(p.colNum) - 1 : NaN;
    if (Number.isFinite(js)) return !cellHasDisplayData(row, js);
    return isMissingDataCause(p.cause, t);
  };

  const parts = [];
  if (hasDup) parts.push(t("plagasArandano.duplicateLots"));
  if (hasTemp) parts.push(t("cartillaAnalysis.missingTemps"));
  if (otherLabels.length === 1) {
    parts.push(formatPreciseErrorText(otherPairs[0], row, t));
  } else if (otherLabels.length > 1) {
    // «Vacío en: …» solo si TODAS las celdas están realmente vacías.
    // Si hay valor incorrecto (ej. año/mes distinto), mostrar la causa real.
    const allEmpty = otherPairs.every(pairIsEmpty);
    if (allEmpty) {
      parts.push(t("cartillaAnalysis.missingFields", { columns: otherLabels.join(", ") }));
    } else {
      // Valor incorrecto (no vacío): siempre «Error en: …» + causa si hay.
      const strong = [];
      otherPairs.forEach((p) => {
        const c = String(p.cause || "").trim();
        if (
          c &&
          !isWeakCause(c, t) &&
          !isGenericOnlyCause(c, t) &&
          !isMissingDataCause(c, t) &&
          !strong.includes(c)
        ) {
          strong.push(c);
        }
      });
      const errorCols = t("cartillaAnalysis.errorFields", {
        columns: otherLabels.join(", ")
      });
      parts.push(strong.length ? `${errorCols} · ${strong.join(" · ")}` : errorCols);
    }
  }
  return parts.filter(Boolean).join(" · ");
}

/**
 * Errores de fila para el panel gerencial.
 * Incluye vacíos obligatorios (SAP faltante) y celdas con dato inválido.
 * @returns {{ causes: string[], columns: string[], pairs: { cause: string, column: string, weak: boolean, colNum: number|null }[] }}
 */
export function extractRowErrorHints(row, options = {}) {
  const {
    errorMap = null,
    duplicateLotes = new Set(),
    colLoteJs = 9,
    headerByColNum = new Map(),
    t = (k) => k,
    includeEmptyObligatorio = true,
    skipSapValidation = false
  } = options;

  const isSapColNum = (colNum) => {
    const n = Number(colNum);
    return (n >= 13 && n <= 27) || (n >= 29 && n <= 33);
  };

  const pairs = [];
  const seen = new Set();

  const pushPair = (cause, column, colNum = null) => {
    if (skipSapValidation && isSapColNum(colNum)) return;
    const col = String(column || "").trim() || (colNum != null ? `Col ${colNum}` : "Col");
    if (skipSapValidation && isSapColumnName(col) && !/\bnota condicion\b/.test(
      col.toLowerCase().normalize("NFD").replace(/\p{M}/gu, "")
    )) {
      // Columna SAP por nombre sin colNum fiable: no reportar en PT.
      if (colNum == null) return;
    }
    const rawCause = cause || t("cartillaAnalysis.genericError");
    if (skipSapValidation && isSapMissingCause(rawCause, t)) return;

    // Bloque SAP (13–33): solo pintura en tabla; no entra al panel.
    if (shouldOmitSapFromAnalysis(col, colNum, rawCause, t)) return;

    const zoneEmpty = !skipSapValidation && isSapFailure(col, colNum, rawCause, t);
    const sapMissing = SHOW_SAP_MISSING_LABEL && zoneEmpty;
    let c = promoteCause(rawCause, t, col, colNum);
    if (!SHOW_SAP_MISSING_LABEL && isSapMissingCause(c, t)) return;
    if (!c) return;

    const key = sapMissing ? `SAP|${t("cartillaAnalysis.missingSap")}` : `${colNum ?? col}|${c}`;
    if (seen.has(key)) return;
    seen.add(key);
    pairs.push({
      cause: sapMissing ? t("cartillaAnalysis.missingSap") : c,
      column: sapMissing ? "SAP" : col,
      colNum: colNum == null ? null : Number(colNum),
      weak: sapMissing ? false : isWeakCause(c, t),
      sapMissing: false
    });
  };

  const resolveHeader = (colNum) => {
    const n = Number(colNum);
    if (!Number.isFinite(n) || n < 1) return "Col";
    return headerByColNum.get(n) || `Col ${n}`;
  };

  const shouldIncludeCell = (_err, _js) => true;

  const filaMap = errorMap?.get?.(row._filaNum);
  if (filaMap?.size) {
    filaMap.forEach((err, colNum) => {
      const js = Number(colNum) - 1;
      if (!shouldIncludeCell(err, js)) return;
      pushPair(
        err?.problema || err?.tipo || t("cartillaAnalysis.genericError"),
        resolveHeader(colNum),
        colNum
      );
    });
  }

  if (row._errorCols instanceof Map) {
    row._errorCols.forEach((msg, colJs) => {
      const js = Number(colJs);
      const err = typeof msg === "string" ? { problema: msg } : msg || {};
      if (!shouldIncludeCell(err, js)) return;
      const colNum = js + 1;
      const header = resolveHeader(colNum);
      let cause =
        typeof msg === "string" ? msg.trim() : String(msg?.problema || "").trim();
      if (!cause) {
        cause = cellHasDisplayData(row, js)
          ? `${header}: valor incorrecto`
          : t("cartillaAnalysis.noDataIn", { column: header });
      }
      pushPair(cause, header, colNum);
    });
  } else if (row._errorCols instanceof Set) {
    row._errorCols.forEach((colJs) => {
      const js = Number(colJs);
      if (!includeEmptyObligatorio && !cellHasDisplayData(row, js)) return;
      const colNum = js + 1;
      const header = resolveHeader(colNum);
      // Vacío obligatorio: mensaje de la columna (no asumir SAP).
      const cause = cellHasDisplayData(row, js)
        ? `${header}: valor incorrecto`
        : t("cartillaAnalysis.noDataIn", { column: header });
      pushPair(cause, header, colNum);
    });
  }

  if (Array.isArray(row._errors) && row._errors.length) {
    row._errors.forEach((err) => {
      // Plagas / módulos que guardan: "Columna 13: Campo obligatorio"
      if (typeof err === "string") {
        const m = err.match(/^Columna\s+(\d+)\s*:\s*(.*)$/i);
        if (m) {
          const colNum = Number(m[1]);
          const header = resolveHeader(colNum);
          const msg =
            String(m[2] || "").trim() || t("cartillaAnalysis.noDataIn", { column: header });
          const js = colNum - 1;
          if (!includeEmptyObligatorio && !cellHasDisplayData(row, js)) return;
          pushPair(msg, header, colNum);
          return;
        }
        if (err.trim()) pushPair(err.trim(), "—", null);
        return;
      }
      const colNum = err.colNum ?? (err.colJs != null ? err.colJs + 1 : err.col);
      if (colNum == null) return;
      const js = Number(colNum) - 1;
      if (!shouldIncludeCell(err, js)) return;
      pushPair(
        err?.message || err?.problema || err?.msg || t("cartillaAnalysis.genericError"),
        resolveHeader(colNum),
        colNum
      );
    });
  }

  const lote = String(row[colLoteJs] ?? "").trim();
  const isRealDuplicate =
    Boolean(lote) &&
    cellHasDisplayData(row, colLoteJs) &&
    (Boolean(row.__duplicado) || Boolean(duplicateLotes?.has?.(lote)));
  // Solo marcar «duplicado» si realmente lo es.
  // `_errorLote` también se usa por longitud/obligatorio → no confundir.
  if (isRealDuplicate) {
    pushPair(t("plagasArandano.duplicateLots"), resolveHeader(colLoteJs + 1), colLoteJs + 1);
  }

  pairs.sort((a, b) => Number(a.weak) - Number(b.weak));

  return {
    causes: pairs.map((p) => p.cause),
    columns: pairs.map((p) => p.column),
    pairs
  };
}

/**
 * @param {object} params
 */
export function buildCartillaAnalysis(params) {
  const {
    rows = [],
    filasConError = [],
    errorMap = null,
    duplicateLotes = new Set(),
    colLoteJs = 9,
    colIdJs = 0,
    columns = [],
    cartilla = "",
    fechaLabel = "",
    t,
    htmlEscape = defaultHtmlEscape,
    translateHeader = (h) => h,
    skipSapValidation = false,
    observations = []
  } = params;

  const total = rows.length;
  const errors = filasConError.length;
  const ok = Math.max(total - errors, 0);
  const conformity = total ? Math.round((ok / total) * 100) : 100;

  let level = "ok";
  if (errors > 0 && conformity >= 95) level = "warn";
  if (errors > 0 && conformity < 95) level = "critical";

  const headerByColNum = new Map();
  columns.forEach((col) => {
    const idx = typeof col === "object" ? col.originalIndex : col;
    const header = typeof col === "object" ? col.header : String(col);
    if (idx == null || !Number.isFinite(Number(idx))) return;
    // Solo clave Excel 1-based (igual que errorMap). No usar idx 0-based:
    // pisaría la columna anterior (ej. LMR 51 ← Observación).
    headerByColNum.set(Number(idx) + 1, translateHeader(header, idx));
  });

  const causeCount = new Map();
  const colCount = new Map();
  const loteDetails = new Map();
  /** Líneas precisas: ID · Lote · Error (una por incidencia). */
  const errorLines = [];
  const errorLineKeys = new Set();

  const pushErrorLine = (id, lote, error, sap = false) => {
    const errText = String(error || "").trim();
    if (!errText) return;
    const isDup = isDuplicateCauseText(errText);
    const isLmr = isLmrMajorityCauseText(errText);
    // Duplicados / LMR: una sola fila resumen (no spam).
    const key = isDup
      ? `DUP|${lote}`
      : isLmr
        ? `LMR|majority`
        : `${id}|${lote}|${sap ? "SAP" : "OTHER"}`;
    if (errorLineKeys.has(key)) {
      const existing = errorLines.find((line) => {
        if (isDup) return line.dup && line.lote === lote;
        if (isLmr) return line.lmrSummary;
        return line.id === (id || "—") && line.lote === lote && Boolean(line.sap) === Boolean(sap);
      });
      if (existing && isDup && id) {
        const ids = String(existing.id)
          .split(",")
          .map((s) => s.trim())
          .filter(Boolean);
        if (id && !ids.includes(id)) {
          ids.push(id);
          existing.id = ids.join(", ");
        }
      } else if (existing && isLmr) {
        existing.count = (existing.count || 1) + 1;
        existing.error = t("cartillaAnalysis.lmrMajoritySummary", {
          count: String(existing.count)
        });
        if (lote && lote !== "—") {
          const lotes = String(existing.lote || "")
            .split(",")
            .map((s) => s.trim())
            .filter(Boolean);
          if (!lotes.includes(lote) && lotes.length < 3) {
            lotes.push(lote);
            existing.lote = lotes.join(", ");
          } else if (!lotes.includes(lote) && lotes.length === 3 && !String(existing.lote).includes("…")) {
            existing.lote = `${lotes.join(", ")}…`;
          }
        }
      } else if (existing && !sap && !isDup) {
        const parts = existing.error.split(" · ").map((s) => s.trim()).filter(Boolean);
        errText.split(" · ").forEach((part) => {
          const p = part.trim();
          if (p && !parts.includes(p)) parts.push(p);
        });
        existing.error = parts.join(" · ");
      }
      return;
    }
    errorLineKeys.add(key);
    errorLines.push({
      id: isLmr ? "—" : id || "—",
      lote: isLmr ? lote || "—" : lote,
      error: isDup
        ? t("plagasArandano.duplicateLots")
        : isLmr
          ? t("cartillaAnalysis.lmrMajoritySummary", { count: "1" })
          : errText,
      sap: Boolean(sap),
      dup: isDup,
      lmrSummary: isLmr,
      count: isLmr ? 1 : undefined
    });
  };

  filasConError.forEach((row) => {
    const lote = String(row[colLoteJs] ?? "").trim() || t("cartillaAnalysis.unknownLot");
    const id = String(row[colIdJs] ?? "").trim();

    const { pairs: rawPairs } = extractRowErrorHints(row, {
      errorMap,
      duplicateLotes,
      colLoteJs,
      headerByColNum,
      t,
      includeEmptyObligatorio: true,
      skipSapValidation
    });

    // Solo lotes con causa real (no el genérico «Desviación de validación»).
    const pairs = rawPairs.filter((p) => !isGenericOnlyCause(p.cause, t));
    if (!pairs.length) return;

    const causes = pairs.map((p) => p.cause);
    const entry = loteDetails.get(lote) || {
      lote,
      ids: [],
      columns: new Set(),
      hints: [],
      count: 0,
      strong: 0
    };
    entry.count += 1;
    if (id && !entry.ids.includes(id)) entry.ids.push(id);

    pairs.forEach((p) => {
      if (p.column) {
        entry.columns.add(p.column);
        if (!p.weak) colCount.set(p.column, (colCount.get(p.column) || 0) + 1);
      }
    });

    const hasSap = SHOW_SAP_MISSING_LABEL && pairs.some((p) => p.sapMissing);
    if (hasSap) {
      pushErrorLine(id, lote, t("cartillaAnalysis.missingSap"), true);
      const short = t("cartillaAnalysis.missingSap");
      if (!entry.hints.includes(short)) entry.hints.push(short);
      entry.strong += 1;
    }

    const nonSapSummary = summarizeNonSapErrorText(pairs, row, t);
    if (nonSapSummary) {
      pushErrorLine(id, lote, nonSapSummary, false);
      if (!entry.hints.includes(nonSapSummary)) entry.hints.push(nonSapSummary);
      entry.strong += 1;
    }

    loteDetails.set(lote, entry);

    const hasStrong = pairs.some((p) => !p.weak);
    causes.forEach((cause) => {
      if (!SHOW_SAP_MISSING_LABEL && isSapMissingCause(cause, t)) return;
      if (hasSap && !isSapMissingCause(cause, t) && isWeakCause(cause, t)) return;
      if (isWeakCause(cause, t) && hasStrong && !isSapMissingCause(cause, t)) return;
      causeCount.set(cause, (causeCount.get(cause) || 0) + 1);
    });
  });

  errorLines.sort((a, b) => {
    // Primero duplicados, luego LMR resumen, luego el resto (sin bloque SAP especial).
    const rank = (line) => (line.dup ? 0 : line.lmrSummary ? 1 : line.sap ? 2 : 3);
    const ra = rank(a);
    const rb = rank(b);
    if (ra !== rb) return ra - rb;
    const loteCmp = String(a.lote).localeCompare(String(b.lote));
    if (loteCmp) return loteCmp;
    return String(a.id).localeCompare(String(b.id));
  });

  const rankedCauses = [...causeCount.entries()]
    .filter(([cause]) => SHOW_SAP_MISSING_LABEL || !isSapMissingCause(cause, t))
    .sort((a, b) => {
    const aSap = SHOW_SAP_MISSING_LABEL && isSapMissingCause(a[0], t) ? 0 : 1;
    const bSap = SHOW_SAP_MISSING_LABEL && isSapMissingCause(b[0], t) ? 0 : 1;
    if (aSap !== bSap) return aSap - bSap;
    const aWeak = isWeakCause(a[0], t) ? 1 : 0;
    const bWeak = isWeakCause(b[0], t) ? 1 : 0;
    if (aWeak !== bWeak) return aWeak - bWeak;
    return b[1] - a[1];
  });
  const topCause = rankedCauses[0] || null;
  const topColumn = [...colCount.entries()].sort((a, b) => b[1] - a[1])[0] || null;
  const allErrorLotes = [...loteDetails.values()]
    .filter((entry) => (entry.hints || []).length > 0)
    .sort(
      (a, b) =>
        (b.strong || 0) - (a.strong || 0) ||
        b.count - a.count ||
        a.lote.localeCompare(b.lote)
    );
  // Todos los lotes con causa (si a todos les falta SAP, salen todos).
  const totalErrorLotes = allErrorLotes.length;
  const topLotes = allErrorLotes.map((entry) => ({
    lote: entry.lote,
    ids: entry.ids.slice(0, 6),
    columns: [...entry.columns].slice(0, 12),
    hints: (entry.hints || []).slice(0, 4),
    count: entry.count
  }));
  const hasNonSapLines = errorLines.some((line) => !line.sap);
  const sapDominant =
    SHOW_SAP_MISSING_LABEL && topCause && isSapMissingCause(topCause[0], t) && !hasNonSapLines;
  const dupSize = duplicateLotes?.size ?? 0;

  const trafficLabel =
    level === "ok"
      ? t("cartillaAnalysis.trafficOk")
      : level === "warn"
        ? t("cartillaAnalysis.trafficWarn")
        : t("cartillaAnalysis.trafficCritical");

  const reading =
    level === "ok"
      ? t("cartillaAnalysis.readingOk")
      : sapDominant && totalErrorLotes > 0
        ? t("cartillaAnalysis.readingMissingSap", {
            lots: String(totalErrorLotes)
          })
        : topCause && totalErrorLotes > 0
          ? t("cartillaAnalysis.readingErrors", {
              cause: topCause[0],
              lots: String(totalErrorLotes)
            })
          : errors > 0
            ? t("cartillaAnalysis.readingErrorsGeneric")
            : t("cartillaAnalysis.readingOk");

  const observationNotes = (Array.isArray(observations) ? observations : [])
    .map((note) => String(note || "").trim())
    .filter(Boolean);

  return {
    cartilla,
    fechaLabel,
    total,
    errors,
    ok,
    conformity,
    level,
    trafficLabel,
    reading,
    observations: observationNotes,
    topCause,
    topColumn,
    topLotes,
    errorLines,
    totalErrorLotes,
    totalErrorLines: errorLines.length,
    sapDominant: Boolean(sapDominant),
    hasNonSapLines: Boolean(hasNonSapLines),
    duplicateCount: dupSize,
    icon: level === "ok" ? "success" : level === "warn" ? "warning" : "error",
    htmlEscape
  };
}

function renderKpis(analysis, t) {
  const esc = analysis.htmlEscape;
  return `
    <div class="agv-cartilla-analysis__kpis">
      <div class="agv-cartilla-analysis__kpi">
        <span class="agv-cartilla-analysis__kpi-label">${esc(t("cartillaAnalysis.conformity"))}</span>
        <strong class="agv-cartilla-analysis__kpi-value">${analysis.conformity}%</strong>
      </div>
      <div class="agv-cartilla-analysis__kpi">
        <span class="agv-cartilla-analysis__kpi-label">${esc(t("cartillaAnalysis.records"))}</span>
        <strong class="agv-cartilla-analysis__kpi-value">${analysis.ok}/${analysis.total}</strong>
      </div>
      <div class="agv-cartilla-analysis__kpi">
        <span class="agv-cartilla-analysis__kpi-label">${esc(t("cartillaAnalysis.errors"))}</span>
        <strong class="agv-cartilla-analysis__kpi-value">${analysis.errors}</strong>
      </div>
      <div class="agv-cartilla-analysis__kpi">
        <span class="agv-cartilla-analysis__kpi-label">${esc(t("cartillaAnalysis.traffic"))}</span>
        <strong class="agv-cartilla-analysis__kpi-value">${esc(analysis.trafficLabel)}</strong>
      </div>
    </div>`;
}

function renderObservations(analysis, t) {
  const esc = analysis.htmlEscape;
  const notes = Array.isArray(analysis.observations) ? analysis.observations : [];
  if (!notes.length) return "";
  return `<aside class="agv-cartilla-analysis__observations" role="note">
    <span class="agv-cartilla-analysis__observations-label">${esc(t("cartillaAnalysis.observations"))}</span>
    ${notes.map((note) => `<p class="agv-cartilla-analysis__observation">${esc(note)}</p>`).join("")}
  </aside>`;
}

function renderDetails(analysis, t) {
  const esc = analysis.htmlEscape;
  const sap = Boolean(analysis.sapDominant);
  const lines = Array.isArray(analysis.errorLines) ? analysis.errorLines : [];

  const causeText = sap
    ? esc(t("cartillaAnalysis.missingSap"))
    : analysis.topCause
      ? `${esc(String(analysis.topCause[0]).replace(/\s*\(SAP\)/gi, "").replace(/\bSAP\b/gi, "").trim())} (${analysis.topCause[1]})`
      : esc(t("cartillaAnalysis.noCause"));

  const lotsLabel = lines.some((l) => !l.sap)
    ? t("cartillaAnalysis.topIncidents")
    : sap
      ? t("cartillaAnalysis.topLots")
      : t("cartillaAnalysis.topLots");

  const lotesHtml = lines.length
    ? `<ul class="agv-cartilla-analysis__lots">${lines
        .map((item) => {
          if (item.lmrSummary) {
            return `<li class="agv-cartilla-analysis__lot agv-cartilla-analysis__lot--precise agv-cartilla-analysis__lot--summary">
              <span class="agv-cartilla-analysis__lot-error" title="${esc(item.error || "")}">${esc(item.error || "")}</span>
            </li>`;
          }
          const id = item.id || "—";
          const lote = item.lote || "—";
          const error = String(item.error || "")
            .replace(/\s*\(SAP\)/gi, "")
            .replace(/\bfalta datos sap\b/gi, t("cartillaAnalysis.seeTable"))
            .replace(/\bmissing sap data\b/gi, t("cartillaAnalysis.seeTable"))
            .trim();
          return `<li class="agv-cartilla-analysis__lot agv-cartilla-analysis__lot--precise">
            <span class="agv-cartilla-analysis__lot-id" title="ID">${esc(id)}</span>
            <span class="agv-cartilla-analysis__lot-sep" aria-hidden="true">·</span>
            <strong class="agv-cartilla-analysis__lot-code" title="Lote">${esc(lote)}</strong>
            <span class="agv-cartilla-analysis__lot-sep" aria-hidden="true">·</span>
            <span class="agv-cartilla-analysis__lot-error" title="${esc(error)}">${esc(error)}</span>
          </li>`;
        })
        .join("")}</ul>`
    : `<p class="agv-cartilla-analysis__empty">${esc(t("cartillaAnalysis.noLots"))}</p>`;

  const dupLine =
    analysis.duplicateCount > 0
      ? `<p class="agv-cartilla-analysis__dup">${esc(
          t("cartillaAnalysis.duplicates", { count: String(analysis.duplicateCount) })
        )}</p>`
      : "";

  if (sap && !analysis.hasNonSapLines) {
    return `
    <div class="agv-cartilla-analysis__block">
      <span class="agv-cartilla-analysis__block-label">${esc(t("cartillaAnalysis.mainCause"))}</span>
      <p class="agv-cartilla-analysis__block-value">${causeText}</p>
    </div>
    <div class="agv-cartilla-analysis__block agv-cartilla-analysis__block--lots">
      <span class="agv-cartilla-analysis__block-label">${esc(lotsLabel)}</span>
      ${lotesHtml}
    </div>
    ${dupLine}
    <p class="agv-cartilla-analysis__reading">${esc(analysis.reading)}</p>`;
  }

  const columnText = analysis.topColumn
    ? `${esc(analysis.topColumn[0])} (${analysis.topColumn[1]})`
    : esc(t("cartillaAnalysis.seeTable"));

  return `
    <div class="agv-cartilla-analysis__grid">
      <div class="agv-cartilla-analysis__block">
        <span class="agv-cartilla-analysis__block-label">${esc(t("cartillaAnalysis.mainCause"))}</span>
        <p class="agv-cartilla-analysis__block-value">${causeText}</p>
      </div>
      <div class="agv-cartilla-analysis__block">
        <span class="agv-cartilla-analysis__block-label">${esc(t("cartillaAnalysis.mainColumn"))}</span>
        <p class="agv-cartilla-analysis__block-value">${columnText}</p>
      </div>
    </div>
    <div class="agv-cartilla-analysis__block agv-cartilla-analysis__block--lots">
      <span class="agv-cartilla-analysis__block-label">${esc(lotsLabel)}</span>
      ${lotesHtml}
    </div>
    ${dupLine}
    <p class="agv-cartilla-analysis__reading">${esc(analysis.reading)}</p>`;
}

export function htmlCartillaAnalysisModal(analysis, t) {
  const esc = analysis.htmlEscape;
  // Modal / «Ver detalle»: solo errores. Observaciones no van aquí.
  return `
    <div class="agv-cartilla-analysis agv-cartilla-analysis--modal agv-cartilla-analysis--${analysis.level}">
      <p class="agv-cartilla-analysis__meta">${esc(analysis.cartilla)} · ${esc(analysis.fechaLabel)}</p>
      ${renderKpis(analysis, t)}
      ${renderDetails(analysis, t)}
    </div>`;
}

export function htmlCartillaAnalysisPanel(analysis, t) {
  const esc = analysis.htmlEscape;
  return `
    <div class="agv-cartilla-analysis agv-cartilla-analysis--panel agv-cartilla-analysis--${analysis.level}">
      <div class="agv-cartilla-analysis__head">
        <div>
          <h4 class="agv-cartilla-analysis__title">${esc(t("cartillaAnalysis.title"))}</h4>
          <p class="agv-cartilla-analysis__meta">${esc(analysis.cartilla)} · ${esc(analysis.fechaLabel)}</p>
        </div>
        <span class="agv-cartilla-analysis__badge">${esc(analysis.trafficLabel)}</span>
      </div>
      ${renderKpis(analysis, t)}
      ${renderDetails(analysis, t)}
    </div>
    ${renderObservations(analysis, t)}`;
}

/**
 * Controlador ligero para enganchar panel + modal en cualquier servicio.
 */
export function createCartillaAnalysisController(options) {
  const {
    getRoot,
    hostSelector,
    showDialog,
    t,
    htmlEscape = defaultHtmlEscape
  } = options;

  let lastAnalysis = null;

  const getHost = () => getRoot()?.querySelector(hostSelector) || null;

  const clear = () => {
    const host = getHost();
    if (host) {
      host.innerHTML = "";
      host.hidden = true;
    }
    lastAnalysis = null;
  };

  const present = (analysisParams) => {
    const analysis = buildCartillaAnalysis({ ...analysisParams, t, htmlEscape });
    lastAnalysis = analysis;
    const host = getHost();
    if (host) {
      host.innerHTML = htmlCartillaAnalysisPanel(analysis, t);
      host.hidden = false;
    }
    if (typeof showDialog === "function") {
      showDialog({
        icon: analysis.icon,
        title: t("cartillaAnalysis.title"),
        html: htmlCartillaAnalysisModal(analysis, t),
        confirmButtonText: t("cartillaAnalysis.continue"),
        wide: true
      });
    }
    return analysis;
  };

  const refreshPanel = (analysisParams) => {
    if (!analysisParams) {
      if (!lastAnalysis) return null;
      const host = getHost();
      if (host) {
        host.innerHTML = htmlCartillaAnalysisPanel(lastAnalysis, t);
        host.hidden = false;
      }
      return lastAnalysis;
    }
    const analysis = buildCartillaAnalysis({ ...analysisParams, t, htmlEscape });
    lastAnalysis = analysis;
    const host = getHost();
    if (host) {
      host.innerHTML = htmlCartillaAnalysisPanel(analysis, t);
      host.hidden = false;
    }
    return analysis;
  };

  return {
    clear,
    present,
    refreshPanel,
    getLast: () => lastAnalysis
  };
}

/**
 * Deriva filas con error desde celdas pintadas en el DOM.
 * Anota row._errorCols con los índices JS de las celdas en rojo (dataset.excelCol).
 */
export function deriveFilasConErrorFromDom(
  tbody,
  rows,
  selector = ".agv-pt-cell-error-empty, .agv-pt-cell-error-value, .agv-mp-cell-error-empty, .agv-mp-cell-error-value"
) {
  if (!tbody || !rows?.length) return [];
  const out = [];
  [...tbody.children].forEach((tr, i) => {
    const row = rows[i];
    if (!row) return;
    const errorCells = tr.querySelectorAll(selector);
    if (!errorCells.length) return;
    const cols = new Map();
    errorCells.forEach((td) => {
      const js = Number(td.dataset.excelCol);
      if (!Number.isFinite(js) || js < 0) return;
      const title = String(td.title || "").trim();
      cols.set(js, title || cols.get(js) || "");
    });
    row._errorCols = cols;
    if (!Array.isArray(row._errors) || !row._errors.length) {
      row._errors = [...cols.entries()]
        .filter(([, msg]) => msg)
        .map(([js, msg]) => `Columna ${js + 1}: ${msg}`);
    }
    out.push(row);
  });
  return out;
}

/** Convierte headers string[] a columnas para buildCartillaAnalysis. */
export function headersToAnalysisColumns(headers = []) {
  return headers.map((header, originalIndex) =>
    typeof header === "object" && header != null
      ? {
          header: header.header ?? String(header),
          originalIndex: header.originalIndex ?? originalIndex
        }
      : { header: String(header ?? ""), originalIndex }
  );
}

/** Compat: alias usado por Espárrago MP. */
export const buildEsparragoMpCartillaAnalysis = buildCartillaAnalysis;
