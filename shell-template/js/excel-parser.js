/** Parseo Excel: horas de pago por turno; el día se valida sumando turnos. */

const HOUR_BASE = 9.6;
const HOUR_HALF = 10.1;
const HOUR_MAX = 10.6;
const HOUR_DAY_CAP = 12;

const HOUR_STEPS = [HOUR_BASE, HOUR_HALF, HOUR_MAX, HOUR_DAY_CAP];
const HOUR_EXACT_EPS = 0.02;

const HOURS_LABEL = "Suma de Horas Pago";

export function matchExactHourStep(hours) {
  if (hours == null || !Number.isFinite(hours)) return null;
  let best = null;
  let bestDist = Infinity;
  for (const step of HOUR_STEPS) {
    const dist = Math.abs(hours - step);
    if (dist < bestDist) {
      bestDist = dist;
      best = step;
    }
  }
  if (bestDist <= HOUR_EXACT_EPS) return best;
  return null;
}

/**
 * Clasifica suma del día (COSTO DE COSECHA) — solo valores exactos:
 * - posible-salida: > 0 y < 9.6 (ej. 5.5 primera pasada → revisar pase)
 * - ok: exacto 9.6
 * - aviso: exacto 10.1
 * - aviso-hora: exacto 10.6 o 12
 * - rojo: cualquier otro (ej. 9.63, 11.37) o > 12
 */
function formatHoursDisplay(hours) {
  if (hours == null || !Number.isFinite(hours)) return "";
  return String(Math.round(hours * 1e3) / 1e3);
}

export function classifyDayHours(hours) {
  if (hours == null || !Number.isFinite(hours) || hours <= 0) {
    return { flag: "na", tip: "Sin horas de pago para sumar. ¿Falta dato en el Excel?" };
  }
  const rounded = Math.round(hours * 1e6) / 1e6;

  if (rounded < HOUR_BASE - HOUR_EXACT_EPS) {
    return {
      flag: "posible-salida",
      tip: `Posible pase de salida: suma ${formatHoursDisplay(rounded)} h (menor a 9.6). Verificar si hay pase registrado.`
    };
  }

  const exact = matchExactHourStep(rounded);
  if (exact === HOUR_BASE) {
    return { flag: "ok", tip: "" };
  }
  if (exact === HOUR_HALF) {
    return {
      flag: "aviso",
      tip: `Aviso: ${formatHoursDisplay(rounded)} h = media hora extra (exacto 10.1).`
    };
  }
  if (exact === HOUR_MAX) {
    return {
      flag: "aviso-hora",
      tip: `Aviso: ${formatHoursDisplay(rounded)} h = 1 h extra (exacto 10.6).`
    };
  }
  if (exact === HOUR_DAY_CAP) {
    return {
      flag: "aviso-hora",
      tip: `Aviso: ${formatHoursDisplay(rounded)} h = tope jornada (exacto 12).`
    };
  }

  if (rounded > HOUR_DAY_CAP + HOUR_EXACT_EPS) {
    return {
      flag: "rojo",
      tip: `Error: suma ${formatHoursDisplay(rounded)} h supera el tope de 12 h.`
    };
  }
  return {
    flag: "rojo",
    tip: `Error: suma ${formatHoursDisplay(rounded)} h ≠ exacto (solo 9.6 / 10.1 / 10.6 / 12).`
  };
}

export function isExactAllowedHour(hours) {
  return matchExactHourStep(hours) != null;
}

/** Fallbacks 0-based: C, D, I, M, U, V/W, Y, AB/AC/AD/AE */
const FALLBACK = {
  codigoTrabajador: 1, // B
  documento: 2,
  trabajador: 3,
  macroPartida: 8,
  actividad: 12, // M
  ceco: 20, // U
  codSupervisor: 21,
  supervisor: 22,
  fecha: 24,
  horaInicio: 27,
  horaFin: 28,
  totalHoras: 29,
  horasPago: 30
};

function normHeader(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\./g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanText(value) {
  return String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isCostoCosechaMacro(value) {
  const h = normHeader(value)
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return (
    h === "costo de cosecha" ||
    h.includes("costo de cosecha") ||
    (h.includes("costo") && h.includes("cosecha"))
  );
}

function pickSheetName(sheetNames) {
  const preferred = sheetNames.find((name) => /^validacion/i.test(String(name).trim()));
  return preferred || sheetNames[0];
}

function formatDni(value) {
  const raw = cleanText(value);
  if (!raw) return "";
  if (/^\d+(\.0+)?$/.test(raw)) return String(Math.trunc(Number(raw)));
  return raw.replace(/\.0+$/, "");
}

/**
 * Horas decimales como Excel:
 * - Horas Pago (AE) si ya viene numérico (5.5, 4.6…)
 * - si no: Total Horas (AD) × 24  ← exactamente la fórmula =AD*24
 * - si no: (Hora Fin − Hora Inicio) × 24 con seriales Excel (sin Date/timezone)
 */
export function excelDurationToHours(value) {
  if (value === null || value === undefined || value === "") return null;

  // Total Horas AD suele ser fracción de día (0.229166… = 05:30:00)
  if (typeof value === "number" && Number.isFinite(value)) {
    if (value > 0 && value < 1.5) return Math.round(value * 24 * 1e10) / 1e10;
    // Ya en horas (raro en AD, común si viniera mal mapeado)
    if (value >= 1.5 && value <= 24) return value;
    return null;
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    // Solo hora del día del Date → duración en AD tipo reloj
    const h =
      value.getUTCHours() + value.getUTCMinutes() / 60 + value.getUTCSeconds() / 3600;
    return Math.round(h * 1e10) / 1e10;
  }

  const raw = cleanText(value).replace(",", ".");
  if (!raw || /^n\.?a\.?$/i.test(raw) || raw === "-") return null;

  // "05:30:00" o "4:36:00"
  const m = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) {
    const hours = Number(m[1]) + Number(m[2]) / 60 + Number(m[3] || 0) / 3600;
    return Math.round(hours * 1e10) / 1e10;
  }

  const n = Number(raw);
  if (!Number.isFinite(n)) return null;
  return excelDurationToHours(n);
}

/** Horas Pago ya calculadas (AE): número directo, sin reinterpretar como fecha. */
export function excelPagoHours(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) {
    // AE válido: horas de un turno (p.ej. 4.1 … 10.6). No usar seriales/fecha.
    if (value >= 0.25 && value <= 24) return value;
    return null;
  }
  if (value instanceof Date) return null;
  const raw = cleanText(value).replace(",", ".");
  if (!raw || /^n\.?a\.?$/i.test(raw)) return null;
  const n = Number(raw);
  if (!Number.isFinite(n)) return null;
  return excelPagoHours(n);
}

/**
 * (Hora Fin − Hora Inicio) × 24 — solo con seriales Excel numéricos.
 * Evita Date/timezone (que producía 5.566667 / 4.733333).
 */
export function durationHours(inicio, fin) {
  const toSerial = (v) => {
    if (typeof v === "number" && Number.isFinite(v)) return v;
    if (v instanceof Date && !Number.isNaN(v.getTime())) {
      // Reconstruir serial aproximado desde UTC (solo fallback)
      const epoch = Date.UTC(1899, 11, 30);
      return (v.getTime() - epoch) / 86400000;
    }
    return null;
  };

  const a = toSerial(inicio);
  const b = toSerial(fin);
  if (a == null || b == null) return null;
  const hours = (b - a) * 24;
  if (!Number.isFinite(hours) || hours < 0 || hours > 48) return null;
  return Math.round(hours * 1e10) / 1e10;
}

/** @deprecated alias — preferir excelDurationToHours / excelPagoHours */
export function excelTimeToHours(value) {
  const pago = excelPagoHours(value);
  if (pago != null) return pago;
  return excelDurationToHours(value);
}

function formatFecha(value) {
  if (value === null || value === undefined || value === "") return "";

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const dd = String(value.getDate()).padStart(2, "0");
    const mm = String(value.getMonth() + 1).padStart(2, "0");
    const yyyy = value.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const utc = Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000;
    const d = new Date(utc);
    const dd = String(d.getUTCDate()).padStart(2, "0");
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    const yyyy = d.getUTCFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  const raw = cleanText(value);
  const m = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    let year = Number(m[3]);
    if (year < 100) year += 2000;
    const dd = String(Number(m[1])).padStart(2, "0");
    const mm = String(Number(m[2])).padStart(2, "0");
    return `${dd}/${mm}/${year}`;
  }

  const m2 = raw.match(/^(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
  if (m2) return formatFecha(m2[1]);
  return raw;
}

/**
 * De "06/08/2026 06:30" (col AB/AC) → solo la hora "06:30".
 * También acepta serial Excel o Date.
 */
export function formatClockTime(value) {
  if (value === null || value === undefined || value === "") return "";

  const pad = (hh, mm) => `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}`;

  // Texto tipo Excel: "06/08/2026 06:30" | "6/08/2026 12:00" | "06:30"
  if (typeof value === "string" || (typeof value !== "number" && !(value instanceof Date))) {
    const raw = cleanText(value);
    if (!raw) return "";
    // fecha + hora → quedarse solo con HH:MM
    const withDate = raw.match(
      /\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/
    );
    if (withDate) return pad(Number(withDate[1]), Number(withDate[2]));
    const onlyTime = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (onlyTime) return pad(Number(onlyTime[1]), Number(onlyTime[2]));
    const anyTime = raw.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (anyTime) return pad(Number(anyTime[1]), Number(anyTime[2]));
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    let frac = value % 1;
    if (frac < 0) frac += 1;
    // Fracción de día o serial con fecha+hora
    if (value >= 0 && value < 1.5) {
      const totalMinutes = Math.round(frac * 24 * 60);
      return pad(Math.floor(totalMinutes / 60) % 24, totalMinutes % 60);
    }
    if (value > 48) {
      const totalMinutes = Math.round(frac * 24 * 60);
      return pad(Math.floor(totalMinutes / 60) % 24, totalMinutes % 60);
    }
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    // Preferir hora local del reloj mostrado; si viene 0 UTC probar UTC
    let hh = value.getHours();
    let mm = value.getMinutes();
    if (hh === 0 && mm === 0 && (value.getUTCHours() !== 0 || value.getUTCMinutes() !== 0)) {
      hh = value.getUTCHours();
      mm = value.getUTCMinutes();
    }
    return pad(hh, mm);
  }

  return "";
}

export function clockToMinutes(clock) {
  const m = String(clock || "").match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return Number(m[1]) * 60 + Number(m[2]);
}

/** Lee celda: prefiere texto formateado (w) en AB/AC para “06/08/2026 06:30”. */
function readSheetCell(sheet, rowIndex, colIndex, preferDisplay = false) {
  if (colIndex == null || colIndex < 0 || rowIndex < 0) return "";
  const addr = window.XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = sheet[addr];
  if (!cell) return "";
  if (preferDisplay && cell.w !== undefined && cell.w !== null && String(cell.w).trim() !== "") {
    return cell.w;
  }
  if (cell.v !== undefined && cell.v !== null) return cell.v;
  if (cell.w !== undefined && cell.w !== null) return cell.w;
  return "";
}

function readHeaderRow(sheet, rowIndex, maxCol) {
  const headers = [];
  for (let c = 0; c <= maxCol; c += 1) {
    headers.push(cleanText(readSheetCell(sheet, rowIndex, c)));
  }
  return headers;
}

function findHeaderRow(sheet, range) {
  let best = { index: range.s.r, score: -1 };
  const scanTo = Math.min(range.s.r + 30, range.e.r);
  for (let r = range.s.r; r <= scanTo; r += 1) {
    let score = 0;
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const h = normHeader(readSheetCell(sheet, r, c));
      if (!h) continue;
      if (h === "documento" || h.startsWith("documento")) score += 3;
      if (h === "trabajador" || h.startsWith("trabajador")) score += 3;
      if (h.includes("macro partida")) score += 3;
      if (h === "hora fin" || h.startsWith("hora fin")) score += 3;
      if (h === "hora inicio" || h.startsWith("hora inicio")) score += 2;
      if (h.includes("horas pago") || h === "horas pago") score += 3;
      if (h === "fecha") score += 2;
      if (h === "supervisor") score += 2;
      if (h.includes("cod") && h.includes("supervisor")) score += 2;
      if (h === "fundo" || h.startsWith("fundo") || h === "sede") score += 2;
    }
    if (score > best.score) best = { index: r, score };
  }
  return best.score > 0 ? best.index : range.s.r;
}

/**
 * Resuelve columnas por nombre de cabecera (prioridad).
 * Letras tipo I/AB/AC solo como fallback si el Excel no trae ese título.
 */
function mapHeaders(headerRow) {
  const mapping = {
    codigoTrabajador: -1,
    documento: -1,
    trabajador: -1,
    supervisor: -1,
    codSupervisor: -1,
    fundo: -1,
    macroPartida: -1,
    actividad: -1,
    ceco: -1,
    fecha: -1,
    horaInicio: -1,
    horaFin: -1,
    totalHoras: -1,
    horasPago: -1,
    horasPagoFromHeader: false,
    horaInicioFromHeader: false,
    horaFinFromHeader: false,
    fundoFromHeader: false
  };

  const scoreMatch = (h, exacts = [], includes = [], excludes = []) => {
    if (!h) return false;
    if (excludes.some((x) => h.includes(x))) return false;
    if (exacts.some((x) => h === x)) return true;
    return includes.some((x) => h.includes(x));
  };

  const pickBest = (candidates) => {
    if (!candidates.length) return -1;
    candidates.sort((a, b) => b.score - a.score || a.index - b.index);
    return candidates[0].index;
  };

  const macroCandidates = [];
  const fundoCandidates = [];
  const fechaCandidates = [];
  const iniCandidates = [];
  const finCandidates = [];
  const pagoCandidates = [];
  const totalCandidates = [];

  headerRow.forEach((header, index) => {
    const h = normHeader(header);
    if (!h) return;

    // Código Trabajador (= Documento de referencia cuando Documento viene vacío)
    if (
      mapping.codigoTrabajador < 0 &&
      scoreMatch(
        h,
        ["codigo trabajador", "cod trabajador", "codigo del trabajador", "cod. trabajador"],
        ["codigo trabajador", "cod trabajador"],
        ["supervisor"]
      )
    ) {
      mapping.codigoTrabajador = index;
      return;
    }

    // Documento / DNI
    if (
      mapping.documento < 0 &&
      scoreMatch(h, ["documento", "dni", "nro documento", "numero documento", "n documento"], [
        "documento",
        "dni"
      ], ["tipo documento", "codigo", "cod "])
    ) {
      mapping.documento = index;
      return;
    }

    // Trabajador
    if (
      mapping.trabajador < 0 &&
      scoreMatch(
        h,
        ["trabajador", "nombre trabajador", "nombres", "apellidos y nombres", "nombre completo"],
        ["trabajador", "nombre trabajador"],
        ["codigo trabajador", "cod trabajador"]
      )
    ) {
      mapping.trabajador = index;
      return;
    }

    // Cód. Supervisor (antes que Supervisor a secas)
    if (
      mapping.codSupervisor < 0 &&
      ((h.includes("cod") && h.includes("supervisor")) ||
        h === "codigo supervisor" ||
        h === "cod supervisor")
    ) {
      mapping.codSupervisor = index;
      return;
    }

    // Supervisor (nombre)
    if (
      mapping.supervisor < 0 &&
      scoreMatch(h, ["supervisor", "nombre supervisor"], ["supervisor"], ["cod", "codigo"])
    ) {
      mapping.supervisor = index;
      return;
    }

    // Fundo / sede / finca
    if (
      scoreMatch(
        h,
        ["fundo", "fundo / sede", "fundo sede", "sede", "finca", "campo"],
        ["fundo", "sede", "finca"],
        ["supervisor"]
      )
    ) {
      let score = 40;
      if (h === "fundo") score = 100;
      else if (h.startsWith("fundo")) score = 90;
      else if (h === "sede" || h.includes("fundo")) score = 80;
      fundoCandidates.push({ index, score });
    }

    // Macro Partida (p. ej. "COSTO DE COSECHA")
    if (h.includes("macro") || h.includes("partida")) {
      let score = 0;
      if (h === "macro partida" || h === "macropartida" || h === "macro-partida") score = 100;
      else if (h.includes("macro partida")) score = 80;
      else if (h === "partida" || h === "nombre partida") score = 40;
      if (index === FALLBACK.macroPartida) score += 25;
      if (score > 0) macroCandidates.push({ index, score, header: h });
    }

    // Actividad (columna M)
    if (
      mapping.actividad < 0 &&
      scoreMatch(
        h,
        ["actividad", "actividad labor", "labor", "nombre actividad"],
        ["actividad"],
        ["macro", "partida", "cod actividad", "codigo actividad"]
      )
    ) {
      mapping.actividad = index;
    }

    // CECO / Centro de costo (columna U)
    if (
      mapping.ceco < 0 &&
      (h === "ceco" ||
        h === "ce co" ||
        h.includes("ceco") ||
        h.includes("centro de costo") ||
        h.includes("centro costo") ||
        h === "c.c." ||
        h === "cc")
    ) {
      mapping.ceco = index;
    }

    // Fecha (no confundir con fecha dentro de hora inicio/fin)
    if (scoreMatch(h, ["fecha", "fecha tareo", "fecha labor", "fecha trabajo"], ["fecha"], ["hora"])) {
      let score = 50;
      if (h === "fecha") score = 100;
      else if (h.startsWith("fecha")) score = 80;
      fechaCandidates.push({ index, score });
    }

    // Hora Inicio
    if (
      scoreMatch(
        h,
        ["hora inicio", "horainicio", "hora de inicio", "hora ini"],
        ["hora inicio", "horainicio", "hora de inicio", "hora ini"],
        ["pago", "fin", "total"]
      ) ||
      (h.includes("inicio") && h.includes("hora"))
    ) {
      let score = 50;
      if (h === "hora inicio" || h === "horainicio") score = 100;
      else if (h.includes("hora") && h.includes("inicio")) score = 90;
      iniCandidates.push({ index, score });
    }

    // Hora Fin
    if (
      scoreMatch(
        h,
        ["hora fin", "horafin", "hora de fin", "hora final"],
        ["hora fin", "horafin", "hora de fin", "hora final"],
        ["pago", "inicio", "total"]
      ) ||
      (h.includes("fin") && h.includes("hora"))
    ) {
      let score = 50;
      if (h === "hora fin" || h === "horafin") score = 100;
      else if (h.includes("hora") && h.includes("fin")) score = 90;
      finCandidates.push({ index, score });
    }

    // Total Horas (duración AD) — solo si el encabezado existe
    if (
      scoreMatch(
        h,
        ["total horas", "total hora", "horas total", "duracion", "duración"],
        ["total horas", "total hora", "duracion"]
      )
    ) {
      let score = 60;
      if (h === "total horas" || h === "total hora") score = 100;
      totalCandidates.push({ index, score });
    }

    // Horas Pago (AE = AD*24)
    if (
      scoreMatch(
        h,
        ["horas pago", "hora pago", "horas a pago", "hrs pago", "horas pagadas"],
        ["horas pago", "hora pago", "hrs pago"]
      )
    ) {
      let score = 70;
      if (h === "horas pago" || h === "hora pago") score = 100;
      pagoCandidates.push({ index, score });
    }
  });

  if (macroCandidates.length) {
    mapping.macroPartida = pickBest(macroCandidates);
  }
  if (fundoCandidates.length) {
    mapping.fundo = pickBest(fundoCandidates);
    mapping.fundoFromHeader = true;
  }
  if (fechaCandidates.length) {
    mapping.fecha = pickBest(fechaCandidates);
  }
  if (iniCandidates.length) {
    mapping.horaInicio = pickBest(iniCandidates);
    mapping.horaInicioFromHeader = true;
  }
  if (finCandidates.length) {
    mapping.horaFin = pickBest(finCandidates);
    mapping.horaFinFromHeader = true;
  }
  if (totalCandidates.length) {
    mapping.totalHoras = pickBest(totalCandidates);
  }
  if (pagoCandidates.length) {
    mapping.horasPago = pickBest(pagoCandidates);
    mapping.horasPagoFromHeader = true;
  }

  // Fallbacks solo si el título no existe
  if (mapping.codigoTrabajador < 0) mapping.codigoTrabajador = FALLBACK.codigoTrabajador;
  if (mapping.documento < 0) mapping.documento = FALLBACK.documento;
  if (mapping.trabajador < 0) mapping.trabajador = FALLBACK.trabajador;
  if (mapping.macroPartida < 0) mapping.macroPartida = FALLBACK.macroPartida;
  if (mapping.actividad < 0) mapping.actividad = FALLBACK.actividad;
  if (mapping.ceco < 0) mapping.ceco = FALLBACK.ceco;
  if (mapping.supervisor < 0) mapping.supervisor = FALLBACK.supervisor;
  if (mapping.codSupervisor < 0) mapping.codSupervisor = FALLBACK.codSupervisor;
  if (mapping.fecha < 0) mapping.fecha = FALLBACK.fecha;
  if (mapping.horaInicio < 0) mapping.horaInicio = FALLBACK.horaInicio;
  if (mapping.horaFin < 0) mapping.horaFin = FALLBACK.horaFin;
  // NO forzar AD si el Excel no trae "Total Horas" (rompe Horas Pago / reloj)
  if (mapping.horasPago < 0) {
    mapping.horasPago = FALLBACK.horasPago;
  }

  return mapping;
}

function rowHasAnyMappedData(sheet, r, mapping) {
  const cols = [
    mapping.codigoTrabajador,
    mapping.documento,
    mapping.trabajador,
    mapping.supervisor,
    mapping.macroPartida,
    mapping.actividad,
    mapping.ceco,
    mapping.fecha,
    mapping.horaInicio,
    mapping.horaFin,
    mapping.horasPago,
    mapping.fundo,
    mapping.codSupervisor
  ];
  return cols.some((c) => cleanText(readSheetCell(sheet, r, c)) !== "");
}

/**
 * Horas del turno — igual que Excel:
 *   AE = AD * 24
 * Prioriza Horas Pago si el encabezado existe (tu columna real).
 * Solo usa Total Horas (AD)×24 si ese encabezado está mapeado.
 * Si no: (Fin − Inicio) × 24 con seriales.
 */
function resolveTurnoHours(sheet, r, mapping) {
  // 1) Horas Pago (preferido cuando el Excel trae el título)
  if (mapping.horasPagoFromHeader || mapping.totalHoras < 0) {
    const fromAe = excelPagoHours(readSheetCell(sheet, r, mapping.horasPago));
    if (fromAe != null) return fromAe;
  }

  // 2) Total Horas (AD) × 24 — solo si hay columna "Total Horas"
  if (mapping.totalHoras >= 0) {
    const fromAd = excelDurationToHours(readSheetCell(sheet, r, mapping.totalHoras));
    if (fromAd != null) return fromAd;
  }

  // 3) Horas Pago (fallback numérico)
  const fromAe2 = excelPagoHours(readSheetCell(sheet, r, mapping.horasPago));
  if (fromAe2 != null) return fromAe2;

  // 4) (Hora Fin − Hora Inicio) × 24
  return durationHours(
    readSheetCell(sheet, r, mapping.horaInicio),
    readSheetCell(sheet, r, mapping.horaFin)
  );
}

export function parseExcelBuffer(buffer, fileName = "archivo.xlsx") {
  if (!window.XLSX?.read) {
    throw new Error("SheetJS no está disponible");
  }

  const workbook = window.XLSX.read(buffer, {
    type: "array",
    cellDates: false,
    raw: true,
    dense: false
  });
  const sheetName = pickSheetName(workbook.SheetNames || []);
  if (!sheetName) throw new Error("El Excel no tiene hojas");

  const sheet = workbook.Sheets[sheetName];
  if (!sheet || !sheet["!ref"]) throw new Error("La hoja está vacía");

  const range = window.XLSX.utils.decode_range(sheet["!ref"]);
  const headerIndex = findHeaderRow(sheet, range);
  const headerRow = readHeaderRow(sheet, headerIndex, range.e.c);
  const mapping = mapHeaders(headerRow);

  const rows = [];
  let costoCosechaCount = 0;
  let otherMacroCount = 0;
  let scannedRows = 0;

  for (let r = headerIndex + 1; r <= range.e.r; r += 1) {
    if (!rowHasAnyMappedData(sheet, r, mapping)) continue;
    scannedRows += 1;

    const macroPartida = cleanText(readSheetCell(sheet, r, mapping.macroPartida));
    const actividad = cleanText(readSheetCell(sheet, r, mapping.actividad));
    const ceco = cleanText(readSheetCell(sheet, r, mapping.ceco));
    const esCostoCosecha = isCostoCosechaMacro(macroPartida);
    if (esCostoCosecha) costoCosechaCount += 1;
    else otherMacroCount += 1;

    const codigoTrabajador = formatDni(readSheetCell(sheet, r, mapping.codigoTrabajador));
    const documentoCell = formatDni(readSheetCell(sheet, r, mapping.documento));
    // Código Trabajador = Documento: si Documento viene vacío, usar el código
    const documento = documentoCell || codigoTrabajador;
    const documentoVacio = Boolean(codigoTrabajador || documentoCell) && !documentoCell;
    const trabajador = cleanText(readSheetCell(sheet, r, mapping.trabajador));
    const horaInicioDisplay = readSheetCell(sheet, r, mapping.horaInicio, true);
    const horaFinDisplay = readSheetCell(sheet, r, mapping.horaFin, true);
    const horaInicioSerial = readSheetCell(sheet, r, mapping.horaInicio, false);
    const horaFinSerial = readSheetCell(sheet, r, mapping.horaFin, false);
    const horasTurno = resolveTurnoHours(sheet, r, mapping);
    const horaInicioKey =
      typeof horaInicioSerial === "number"
        ? String(horaInicioSerial)
        : cleanText(horaInicioDisplay || horaInicioSerial);

    // Reloj: texto formateado O serial; Fecha desde columna Fecha o desde el datetime del reloj
    const horaInicioTexto =
      formatClockTime(horaInicioDisplay) || formatClockTime(horaInicioSerial);
    const horaFinTexto =
      formatClockTime(horaFinDisplay) || formatClockTime(horaFinSerial);

    let fecha = formatFecha(readSheetCell(sheet, r, mapping.fecha));
    if (!fecha) {
      fecha =
        formatFecha(horaInicioDisplay) ||
        formatFecha(horaInicioSerial) ||
        formatFecha(horaFinDisplay) ||
        formatFecha(horaFinSerial);
    }

    // Clave de día numérica (serial Excel) por si la fecha texto falla
    let fechaSerial = null;
    if (typeof horaInicioSerial === "number" && Number.isFinite(horaInicioSerial)) {
      fechaSerial = Math.floor(horaInicioSerial);
    } else if (typeof horaFinSerial === "number" && Number.isFinite(horaFinSerial)) {
      fechaSerial = Math.floor(horaFinSerial);
    }

    rows.push({
      rowIndex: r + 1,
      excelRow: r + 1,
      codigoTrabajador,
      documento,
      documentoCell,
      documentoVacio,
      trabajador,
      supervisor: cleanText(readSheetCell(sheet, r, mapping.supervisor)),
      codSupervisor: formatDni(readSheetCell(sheet, r, mapping.codSupervisor)),
      fundo: cleanText(readSheetCell(sheet, r, mapping.fundo)),
      macroPartida,
      actividad,
      ceco,
      fecha,
      fechaSerial,
      horaInicioKey,
      horaInicioTexto,
      horaFinTexto,
      horaInicioMin: clockToMinutes(horaInicioTexto),
      horaFinMin: clockToMinutes(horaFinTexto),
      esCostoCosecha,
      horasTurno,
      horas: horasTurno,
      costoCosecha: horasTurno,
      totalDia: null,
      variedad: "",
      tipo: "",
      estado: "",
      sessionTipo: "",
      sessionVariedad: "",
      hoursByDay: { [HOURS_LABEL]: horasTurno }
    });
  }

  // Suma del día: turnos del mismo Documento + Fecha (+ macro)
  // Ej.: 5.5 + 4.6 = 10.1 · 5.5 + 4.1 = 9.6 — NUNCA mezclar 2 fechas.
  const dayGroups = new Map();
  rows.forEach((row) => {
    const dayPart =
      row.fecha ||
      (row.fechaSerial != null ? `serial-${row.fechaSerial}` : `fila-${row.excelRow}`);
    const key = `${row.documento}|${dayPart}|${row.macroPartida}|${row.esCostoCosecha ? "1" : "0"}`;
    if (!dayGroups.has(key)) dayGroups.set(key, []);
    dayGroups.get(key).push(row);
  });
  dayGroups.forEach((group) => {
    const ordered = group
      .slice()
      .sort((a, b) => (a.horaInicioMin ?? 0) - (b.horaInicioMin ?? 0));
    // Deduplicar el mismo turno (mismo inicio) por si el Excel repite filas
    const seenIni = new Set();
    const unique = [];
    ordered.forEach((row) => {
      const k =
        row.horaInicioTexto || row.horaInicioKey
          ? String(row.horaInicioKey || row.horaInicioTexto)
          : `row-${row.excelRow}`;
      if (seenIni.has(k)) return;
      seenIni.add(k);
      unique.push(row);
    });
    const turnRows = unique.length ? unique : ordered;
    let turnos = turnRows
      .map((row) => Number(row.horasTurno))
      .filter((n) => Number.isFinite(n));
    let total = turnos.reduce((sum, n) => sum + n, 0);
    let rounded = Math.round(total * 1e10) / 1e10;
    let detalle = turnos.map((n) => formatHoursDisplay(n)).join(" + ");

    // Si Horas Pago ya trae el total del día repetido en cada turno, no duplicar
    if (
      turnos.length >= 2 &&
      turnos.every((h) => matchExactHourStep(h) != null) &&
      turnos.every((h) => Math.abs(h - turnos[0]) <= HOUR_EXACT_EPS)
    ) {
      rounded = turnos[0];
      detalle = formatHoursDisplay(rounded);
      turnos = [rounded];
    }

    const inicios = turnRows.map((r) => r.horaInicioTexto).filter(Boolean);
    const fines = turnRows.map((r) => r.horaFinTexto).filter(Boolean);
    group.forEach((row) => {
      row.totalDia = rounded;
      row.horas = rounded;
      row.sumaHorasPago = rounded;
      row.turnosDetalle = detalle;
      row.turnosCount = turnos.length;
      row.horasInicioDetalle = inicios;
      row.horasFinDetalle = fines;
      if (row.esCostoCosecha) row.costoCosecha = rounded;
      row.hoursByDay = { [HOURS_LABEL]: rounded };
    });
  });

  const costoFromRows = rows.filter((row) => row.esCostoCosecha).length;
  if (costoFromRows !== costoCosechaCount) costoCosechaCount = costoFromRows;

  return {
    fileName,
    sheetName,
    headerIndex,
    mapping,
    rows,
    dayLabels: [HOURS_LABEL],
    meta: {
      hourBase: HOUR_BASE,
      hourHalf: HOUR_HALF,
      hourMax: HOUR_MAX,
      hourDayCap: HOUR_DAY_CAP,
      scannedRows,
      costoCosechaCount,
      otherMacroCount,
      totalRows: rows.length
    }
  };
}

export { HOUR_BASE, HOUR_HALF, HOUR_MAX, HOUR_DAY_CAP, HOUR_STEPS, HOURS_LABEL, FALLBACK };
