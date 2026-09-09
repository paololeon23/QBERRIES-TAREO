/** Vista Resumen: tabla por supervisor/fundo + gráfico + errores. */

import { collapseToDayRows } from "./validacion-table.js";
import {
  getHorariosCosecha,
  horarioMananaLabel,
  horarioTardeLabel
} from "./horarios-cosecha.js";
import {
  countScanerCosto,
  countCosechaCosto,
  countAuxiliarCalidad,
  countSupervisoresCalidad,
  normNombreKpi,
  isActividadContadaPorGrupo,
  isActividadCalidad,
  filterGruposResumen,
  filterRowsResumenExcluidos,
  isSupervisorExcluidoResumen
} from "./validacion-kpi.js";

const SUPERVISORES_LICAPA_URL = "data/supervisores-licapa.json";
const SUPERVISORES_GENERALES_URL = "data/supervisores-generales-licapa.json";
const PLANILLAS_FALTANTES_STORAGE_KEY = "qberries.supervisoresLicapa.v6";

/** No exigen cerrar planilla (lista oficial LICAPA). */
const SUPERVISORES_NO_ACTIVOS_DNI = new Set([
  "42493820", "44141396", "48590607", "61014348", "70132627",
  "72911037", "73503134", "74047419", "74068569", "74239909",
  "74317270", "74984893", "75075892", "75078541", "77914317", "78011755", "78199416"
]);

/** Excluidos de validación temporalmente (no reactivar aunque aparezcan en tareo). */
const SUPERVISORES_EXCLUIDOS_VALIDACION_DNI = new Set([
  "72911037", // VASQUEZ COTRINA EVELYN RUVIT
  "74317270", // PASTOR CUEVA SORAYDA ARACELY
  "78199416" // MARTINEZ REYES WILSON ALFREDO
]);

/** Administradores de cosecha (no cosecha de campo): al final de tablas + otro color. */
const ADMIN_COSECHA_DNI = new Set(["76769404", "48462734", "75078541", "78199416"]);
const ADMIN_COSECHA_NOMBRES = [
  "LAIZA PASTOR SARAI SIORELA",
  "PASTOR LAIZA SARAI SIORELA",
  "ARAMBULO RAZURI MIGUEL FERNANDO",
  "RAZURI ARAMBULO MIGUEL FERNANDO",
  "LEON VARGAS DEYSI TATIANA",
  "MARTINEZ REYES WILSON ALFREDO"
];

/** Personal LICAPA II — no Avísame ni validación de planillas LICAPA I. */
const SUPERVISORES_LICAPA_II_DNI = new Set([
  "74047419", // NAMOC NARRO BIVIANA DE LOS ANGELES
  "74239909" // PADILLA NUÑEZ JESUS MARIA
]);

/** Área de Calidad — tampoco exigen planilla. */
const SUPERVISORES_CALIDAD_DNI = new Set(["75501379"]);

function clockTextToMinutes(txt) {
  if (txt == null || txt === "") return null;
  const m = String(txt).match(/(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return Number(m[1]) * 60 + Number(m[2]);
}

/** Pares inicio/fin por turno del día (colapsado o fila suelta). */
function getDayTurnPairs(row) {
  const inicios = row.horasInicioDetalle || (row.horaInicioTexto ? [row.horaInicioTexto] : []);
  const fines = row.horasFinDetalle || (row.horaFinTexto ? [row.horaFinTexto] : []);
  const iniMins = inicios.map(clockTextToMinutes).filter((m) => m != null);
  const finMins = fines.map(clockTextToMinutes).filter((m) => m != null);
  const pairs = [];

  if (iniMins.length || finMins.length) {
    const n = Math.max(iniMins.length, finMins.length);
    for (let i = 0; i < n; i += 1) {
      pairs.push({
        ini: iniMins[i] ?? iniMins[0] ?? null,
        fin: finMins[i] ?? finMins[finMins.length - 1] ?? null
      });
    }
  } else if (row.horaInicioMin != null || row.horaFinMin != null) {
    pairs.push({ ini: row.horaInicioMin ?? null, fin: row.horaFinMin ?? null });
  }

  return pairs;
}

/** 1.er corte válido según fundo/actividad. */
function dayRowTieneBloqueManana(row) {
  const hz = getHorariosCosecha(row?.fundo, row?.actividad);
  return getDayTurnPairs(row).some(
    ({ ini, fin }) =>
      ini != null &&
      fin != null &&
      ini <= hz.firstStartMin + 2 &&
      fin >= hz.firstEndMin - 10 &&
      fin <= hz.firstEndMin + 10
  );
}

/** 2.º corte válido según fundo/actividad. */
function dayRowTieneBloqueTarde(row) {
  const hz = getHorariosCosecha(row?.fundo, row?.actividad);
  return getDayTurnPairs(row).some(
    ({ ini, fin }) =>
      ini != null &&
      fin != null &&
      ini >= hz.secondStartMin - 2 &&
      fin >= hz.secondEndMin - 2
  );
}

/** Persona-día con jornada completa según fundo/actividad (mañana+tarde o directo ≥ 9.6 h). */
function dayRowCerroPlanilla(row) {
  const hz = getHorariosCosecha(row?.fundo, row?.actividad);
  const pairs = getDayTurnPairs(row);
  const iniMins = pairs.map((p) => p.ini).filter((m) => m != null);
  const finMins = pairs.map((p) => p.fin).filter((m) => m != null);

  const hasMorning = iniMins.some((m) => Math.abs(m - hz.firstStartMin) <= 2);
  const hasAfternoon = iniMins.some((m) => m >= hz.secondStartMin - 2);

  if (hasMorning && hasAfternoon) return true;

  /* Reloj solo mañana → no cerró */
  if (iniMins.length && !hasAfternoon) {
    const allStartMorning = iniMins.every((m) => m < hz.secondStartMin - 2);
    const allEndNoonOrBefore =
      !finMins.length || finMins.every((f) => f <= hz.firstEndMin + 10);
    if (allStartMorning && allEndNoonOrBefore) return false;
  }

  /* Jornada directa */
  if (pairs.length === 1) {
    const { ini, fin } = pairs[0];
    if (
      ini != null &&
      fin != null &&
      ini <= hz.firstStartMin + 2 &&
      fin >= hz.directEndMin - 2
    ) {
      return true;
    }
  }

  if (row.horaInicioMin != null && row.horaFinMin != null) {
    if (
      row.horaInicioMin <= hz.firstStartMin + 2 &&
      row.horaFinMin >= hz.directEndMin - 2 &&
      row.horaFinMin > hz.firstEndMin + 10
    ) {
      return true;
    }
    if (
      row.horaInicioMin <= hz.firstStartMin + 2 &&
      row.horaFinMin <= hz.firstEndMin + 10
    ) {
      return false;
    }
  }

  /* Sin reloj claro: usar suma solo si no hay indicios de solo mañana */
  if (!iniMins.length && !finMins.length) {
    const sum = Number(row.sumaHorasPago ?? row.totalDia ?? row.horas ?? row.horasTurno ?? 0);
    if (Number.isFinite(sum) && sum >= hz.fullDayHours - 0.05) return true;
  }

  return false;
}

function motivoPlanillaFaltante(stats) {
  if (!stats) return "Sin cierre mañana ni tarde";
  if (!stats.planillas) return "Sin cierre mañana ni tarde";
  if (stats.cerradas > 0) return "Planilla cerrada";
  if (stats.bloqueManana > 0 && stats.bloqueTarde === 0) {
    return `Sin cierre tarde (${stats.soloManana} solo mañana)`;
  }
  if (stats.bloqueManana === 0 && stats.bloqueTarde === 0) {
    return "Sin cierre mañana ni tarde";
  }
  if (stats.bloqueTarde > 0 && stats.bloqueManana === 0) {
    return "Sin cierre mañana";
  }
  return "Sin planilla cerrada";
}
let chartPersonas = null;
let chartErrores = null;
let supervisoresLicapaCatalog = null;
let supervisoresGeneralesLicapa = null;
let planillasFaltantesState = {
  missing: [],
  present: [],
  inactive: [],
  reactivados: [],
  catalog: [],
  activeCatalog: [],
  personasApoyoHoy: new Map(),
  statusMap: new Map(),
  filters: {}
};
let planillasFaltantesView = "faltantes";
let planillasFaltantesResumenTab = "todos";

/** UI tabla resumen (buscador + tabs de estado). */
let resumenUiState = {
  allRows: [],
  tab: "todos",
  search: ""
};

function destroyChart(chart) {
  if (chart) {
    chart.destroy();
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function isActividadSupervisorCosecha(actividad) {
  return normNombreKpi(actividad) === "SUPERVISOR DE COSECHA";
}

/**
 * Apoyo supervisión: trabajador con actividad SUPERVISOR DE COSECHA
 * bajo OTRO supervisor (ej. ANTICONA → REBAZA).
 * Se recalcula SIEMPRE con las filas recibidas (sin cache).
 * Clave: FUNDO||SUPERVISOR para no mezclar fundos.
 */
export function buildApoyoSupervisionMap(rows) {
  const bySup = new Map();
  (rows || []).forEach((row) => {
    if (row && row.esCostoCosecha === false) return;
    if (!isActividadSupervisorCosecha(row?.actividad)) return;
    const fundo = String(row.fundo || "").trim() || "LICAPA";
    const fundoKey = fundo.toUpperCase();
    const sup = String(row.supervisor || "").trim();
    if (!sup) return;
    const supKey = normNombreKpi(sup);
    if (!supKey) return;
    const mapKey = `${fundoKey}||${supKey}`;
    if (!bySup.has(mapKey)) {
      bySup.set(mapKey, { supervisor: sup, fundo, apoyos: new Map() });
    }
    const trab = String(row.trabajador || "").trim();
    const trabKey = normNombreKpi(trab);
    if (!trabKey || trabKey === supKey) return;
    bySup.get(mapKey).apoyos.set(trabKey, {
      nombre: trab,
      dni: String(row.documento || "").replace(/\D/g, ""),
      fundo: fundoKey
    });
  });

  const out = new Map();
  bySup.forEach((v, mapKey) => {
    const list = [...v.apoyos.values()].sort((a, b) =>
      a.nombre.localeCompare(b.nombre, "es")
    );
    out.set(mapKey, {
      apoyo: list.length > 0,
      apoyoNombres: list.map((p) => p.nombre),
      apoyosDetalle: list,
      supervisor: v.supervisor,
      fundo: v.fundo
    });
  });
  return out;
}

/** Personas apoyo del alcance (Map nombreNorm → persona). Sin estado previo. */
export function collectPersonasApoyoHoy(rows, soloBajoSupervisores = null) {
  const map = buildApoyoSupervisionMap(rows);
  const out = new Map();
  map.forEach((info) => {
    const supKey = normNombreKpi(info.supervisor);
    if (soloBajoSupervisores && !soloBajoSupervisores.has(supKey)) return;
    (info.apoyosDetalle || []).forEach((p) => {
      const key = normNombreKpi(p.nombre);
      if (!key || out.has(key)) return;
      out.set(key, {
        nombre: p.nombre,
        dni: p.dni || "",
        fundo: p.fundo || "LICAPA",
        activo: true,
        nota: "",
        apoyoHoy: true
      });
    });
  });
  return out;
}

function apoyoMapLookup(apoyoMap, fundo, supervisor) {
  if (!apoyoMap) return null;
  const fundoKey = String(fundo || "").trim().toUpperCase() || "LICAPA";
  const supKey = normNombreKpi(supervisor);
  if (!supKey) return null;
  return apoyoMap.get(`${fundoKey}||${supKey}`) || null;
}

function personasApoyoHoyKeys(personasApoyoHoy) {
  if (personasApoyoHoy instanceof Map) return new Set(personasApoyoHoy.keys());
  if (personasApoyoHoy instanceof Set) return personasApoyoHoy;
  return new Set();
}

function emptyPlanillasFaltantesState() {
  return {
    missing: [],
    present: [],
    inactive: [],
    reactivados: [],
    catalog: [],
    activeCatalog: [],
    personasApoyoHoy: new Map(),
    statusMap: new Map(),
    validatedRows: [],
    filters: {}
  };
}

/** Limpia estado de planillas/apoyo (nuevo Excel o cerrar modal). */
export function resetPlanillasFaltantesState() {
  planillasFaltantesState = emptyPlanillasFaltantesState();
  planillasFaltantesResumenTab = "todos";
  planillasFaltantesView = "faltantes";
}

/** Agrupa persona-día por Fundo + Supervisor. */
export function buildResumenGroups(dayRows, apoyoMap = null) {
  const map = new Map();
  const trabajadoresGlobal = new Set();

  dayRows.forEach((row) => {
    const fundo = row.fundo || "(sin fundo)";
    const supervisor = row.supervisor || "(sin supervisor)";
    const key = `${fundo}||${supervisor}`;
    if (!map.has(key)) {
      map.set(key, {
        fundo,
        supervisor,
        planillas: 0,
        trabajadores: new Set(),
        errores: 0,
        avisos: 0,
        ok: 0
      });
    }
    const g = map.get(key);
    g.planillas += 1;
    if (row.documento) {
      g.trabajadores.add(String(row.documento));
      trabajadoresGlobal.add(String(row.documento));
    }
    const isError =
      row.status === "rojo" ||
      row.dayFlags?.horaInicio === "rojo" ||
      row.dayFlags?.horaFin === "rojo" ||
      row.dayFlags?.ceco === "rojo" ||
      row.dayFlags?.documento === "rojo" ||
      row.dayFlags?.trabajador === "rojo" ||
      (row.flags || []).includes("rojo");
    if (isError) g.errores += 1;
    else if (row.status === "aviso") g.avisos += 1;
    else g.ok += 1;
  });

  const groups = [...map.values()]
    .map((g) => {
      const apoyoInfo = apoyoMapLookup(apoyoMap, g.fundo, g.supervisor);
      return {
        fundo: g.fundo,
        supervisor: g.supervisor,
        planillas: g.planillas,
        trabajadores: g.trabajadores.size,
        errores: g.errores,
        avisos: g.avisos,
        ok: g.ok,
        apoyo: Boolean(apoyoInfo?.apoyo),
        apoyoNombres: apoyoInfo?.apoyoNombres || []
      };
    })
    .sort((a, b) => {
      const fa = a.fundo.localeCompare(b.fundo, "es");
      if (fa) return fa;
      return a.supervisor.localeCompare(b.supervisor, "es");
    });

  return {
    groups,
    kpis: {
      grupos: groups.length,
      fundos: new Set(groups.map((g) => g.fundo)).size,
      supervisores: new Set(groups.map((g) => g.supervisor)).size,
      trabajadores: trabajadoresGlobal.size,
      errores: groups.reduce((s, g) => s + g.errores, 0),
      avisos: groups.reduce((s, g) => s + g.avisos, 0),
      conApoyo: groups.filter((g) => g.apoyo).length
    }
  };
}

/** Personas únicas por Macro Partida. */
export function buildPersonasPorMacro(dayRows) {
  const map = new Map();
  dayRows.forEach((row) => {
    const macro = row.macroPartida || "(sin macro)";
    if (!map.has(macro)) map.set(macro, new Set());
    if (row.documento) map.get(macro).add(String(row.documento));
  });
  return [...map.entries()]
    .map(([label, set]) => ({ label, value: set.size }))
    .sort((a, b) => b.value - a.value);
}

/** Personas únicas por Actividad (columna M).
 *  COSECHA / SCANER / SUPERVISOR DE COSECHA → 1 por grupo (equipo), no por trabajador. */
export function buildPersonasPorActividad(dayRows) {
  const map = new Map();
  dayRows.forEach((row) => {
    const act = String(row.actividad || "").trim() || "(sin actividad)";
    if (!map.has(act)) map.set(act, new Set());
    if (row.esCostoCosecha && isActividadContadaPorGrupo(row.actividad)) {
      const sup = normNombreKpi(row.supervisor);
      if (sup) map.get(act).add(`grp:${sup}`);
    } else if (row.documento) {
      map.get(act).add(String(row.documento));
    }
  });
  return [...map.entries()]
    .map(([label, set]) => ({ label, value: set.size }))
    .sort((a, b) => b.value - a.value);
}

function isErrorDayRow(row) {
  return (
    row.status === "rojo" ||
    row.dayFlags?.horaInicio === "rojo" ||
    row.dayFlags?.horaFin === "rojo" ||
    row.dayFlags?.ceco === "rojo" ||
    row.dayFlags?.documento === "rojo" ||
    row.dayFlags?.trabajador === "rojo" ||
    (row.flags || []).includes("rojo")
  );
}

/** Supervisores con más errores (persona-día en rojo), top N. */
export function buildErroresPorSupervisor(dayRows, limit = 10) {
  const map = new Map();
  dayRows.forEach((row) => {
    if (!isErrorDayRow(row)) return;
    const name = row.supervisor || "(sin supervisor)";
    map.set(name, (map.get(name) || 0) + 1);
  });
  return [...map.entries()]
    .map(([label, value]) => ({ label, value }))
    .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label, "es"))
    .slice(0, limit);
}

const PIE_COLORS = [
  "#5dade2",
  "#58d68d",
  "#f5b041",
  "#af7ac5",
  "#5d6d7e",
  "#ec7063",
  "#48c9b0",
  "#f4d03f",
  "#85929e",
  "#a569bd"
];

const ERROR_BAR_COLORS = [
  "#e74c3c",
  "#c0392b",
  "#e67e22",
  "#d35400",
  "#cd6155",
  "#ec7063",
  "#f1948a",
  "#e59866",
  "#af7ac5",
  "#5d6d7e"
];

export async function openResumenModal(getValidated) {
  const modal = document.getElementById("modalResumen");
  if (modal) modal.hidden = false;
  resumenUiState.tab = "todos";
  resumenUiState.search = "";
  const searchEl = document.getElementById("resumenSearch");
  if (searchEl) searchEl.value = "";
  document.querySelectorAll("[data-resumen-tab]").forEach((btn) => {
    btn.classList.toggle("is-active", btn.getAttribute("data-resumen-tab") === "todos");
  });
  // Precarga catálogo para filas “No asistió” / apoyo
  await loadAndRenderResumenFaltantes(getValidated);
}

function supervisorFaltaCerrarManana(st) {
  if (!st || !st.planillas) return true;
  if (st.cerradas > 0) return false;
  return st.bloqueManana === 0;
}

function supervisorFaltaCerrarTarde(st) {
  if (!st || !st.planillas) return true;
  if (st.cerradas > 0) return false;
  return st.bloqueTarde === 0;
}

async function loadSupervisoresGeneralesLicapa() {
  try {
    const res = await fetch(SUPERVISORES_GENERALES_URL, { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    supervisoresGeneralesLicapa = Array.isArray(data?.grupos) ? data.grupos : [];
    supervisorGeneralIndex = null;
    return supervisoresGeneralesLicapa;
  } catch (err) {
    console.warn("[supervisores-generales] No se pudo cargar:", err);
    if (!Array.isArray(supervisoresGeneralesLicapa)) supervisoresGeneralesLicapa = [];
    supervisorGeneralIndex = null;
    return supervisoresGeneralesLicapa;
  }
}

/** Tokens ordenados: "PURIZAGA SAAVEDRA…" = "SAAVEDRA PURIZAGA…" */
function nombreTokensKey(name) {
  return normNombreKpi(name)
    .split(/\s+/)
    .filter((t) => t.length > 1)
    .sort()
    .join(" ");
}

/** Índice DNI / nombre → Supervisor General (lista oficial LICAPA). */
let supervisorGeneralIndex = null;

function buildSupervisorGeneralIndex() {
  const byDni = new Map();
  const byName = new Map();
  const grupos = Array.isArray(supervisoresGeneralesLicapa) ? supervisoresGeneralesLicapa : [];
  grupos.forEach((g) => {
    const info = {
      generalNombre: String(g?.general?.nombre || "").trim(),
      generalDni: String(g?.general?.dni || "").replace(/\D/g, "")
    };
    if (!info.generalNombre) return;
    (Array.isArray(g?.miembros) ? g.miembros : []).forEach((m) => {
      const dni = String(m?.dni || "").replace(/\D/g, "");
      if (dni) byDni.set(dni, info);
      const names = [
        m?.nombre,
        ...(Array.isArray(m?.alias) ? m.alias : []),
        ...(Array.isArray(m?.aliases) ? m.aliases : [])
      ].filter(Boolean);
      names.forEach((n) => {
        const key = normNombreKpi(n);
        if (key) byName.set(key, info);
        const tok = nombreTokensKey(n);
        if (tok) byName.set(`tok:${tok}`, info);
      });
    });
  });
  return { byDni, byName };
}

function getSupervisorGeneralIndex() {
  if (!supervisorGeneralIndex) supervisorGeneralIndex = buildSupervisorGeneralIndex();
  return supervisorGeneralIndex;
}

function lookupSupervisorGeneral(nombre, dni = "") {
  const idx = getSupervisorGeneralIndex();
  const dniKey = String(dni || "").replace(/\D/g, "");
  if (dniKey && idx.byDni.has(dniKey)) return idx.byDni.get(dniKey);
  const key = normNombreKpi(nombre);
  if (key && idx.byName.has(key)) return idx.byName.get(key);
  const tok = nombreTokensKey(nombre);
  if (tok && idx.byName.has(`tok:${tok}`)) return idx.byName.get(`tok:${tok}`);
  // match flexible por tokens
  let best = null;
  idx.byName.forEach((info, k) => {
    if (best || !k.startsWith("tok:")) return;
    if (nombresEquivalentes(nombre, k.slice(4).split(" ").join(" "))) best = info;
  });
  return best;
}

/**
 * LICAPA II/III → otro fundo.
 * LICAPA en lista oficial → nombre del Supervisor General.
 * LICAPA sin lista → Nuevo (sin lista).
 */
function resolveSupervisorGeneralMeta(nombre, dni, fundo) {
  const fundoNorm = String(fundo || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
  if (fundoNorm === "LICAPA II" || fundoNorm === "LICAPA III") {
    return {
      sgKey: "otro-fundo",
      sgLabel: fundoNorm,
      sgTone: "muted",
      generalNombre: fundoNorm,
      generalDni: ""
    };
  }
  const hit = lookupSupervisorGeneral(nombre, dni);
  if (hit?.generalNombre) {
    return {
      sgKey: "lista",
      sgLabel: hit.generalNombre,
      sgTone: "ok",
      generalNombre: hit.generalNombre,
      generalDni: hit.generalDni || ""
    };
  }
  const esLicapaI =
    !fundoNorm ||
    fundoNorm === "LICAPA" ||
    (fundoNorm.includes("LICAPA") && !/\bII\b|\bIII\b/.test(fundoNorm));
  if (esLicapaI) {
    return {
      sgKey: "nuevo",
      sgLabel: "Nuevo (sin lista)",
      sgTone: "warn",
      generalNombre: "",
      generalDni: ""
    };
  }
  return {
    sgKey: "otro",
    sgLabel: "—",
    sgTone: "muted",
    generalNombre: "",
    generalDni: ""
  };
}

function nombresEquivalentes(a, b) {
  const ka = nombreTokensKey(a);
  const kb = nombreTokensKey(b);
  if (ka && kb && ka === kb) return true;
  const ta = new Set(normNombreKpi(a).split(/\s+/).filter((t) => t.length > 1));
  const tb = new Set(normNombreKpi(b).split(/\s+/).filter((t) => t.length > 1));
  if (!ta.size || !tb.size) return false;
  let hit = 0;
  ta.forEach((t) => {
    if (tb.has(t)) hit += 1;
  });
  const need = Math.min(3, Math.min(ta.size, tb.size));
  return hit >= need && hit / Math.min(ta.size, tb.size) >= 0.75;
}

export function isAdminCosechaSupervisor(nombre, dni = "") {
  const dniKey = String(dni || "").replace(/\D/g, "");
  if (dniKey && ADMIN_COSECHA_DNI.has(dniKey)) return true;
  const key = normNombreKpi(nombre);
  if (!key) return false;
  if (ADMIN_COSECHA_NOMBRES.some((n) => normNombreKpi(n) === key)) return true;
  return ADMIN_COSECHA_NOMBRES.some((n) => nombresEquivalentes(nombre, n));
}

/** Busca planilla del supervisor: DNI (codSupervisor) primero, luego nombre flexible. */
function lookupStatusSupervisor(statusMap, nombre, dni = "") {
  if (!statusMap || !statusMap.size) return null;

  const dniKey = String(dni || "").replace(/\D/g, "");
  if (dniKey && statusMap.byDni instanceof Map && statusMap.byDni.has(dniKey)) {
    return statusMap.byDni.get(dniKey);
  }

  if (!nombre) return null;
  const exact = normNombreKpi(nombre);
  if (exact && statusMap.has(exact)) return statusMap.get(exact);

  const tok = nombreTokensKey(nombre);
  let best = null;
  statusMap.forEach((st, key) => {
    if (typeof key !== "string") return;
    const match =
      (tok && nombreTokensKey(key) === tok) || nombresEquivalentes(nombre, key);
    if (!match) return;
    if (!best || (st.planillas || 0) > (best.planillas || 0)) best = st;
  });
  return best;
}

/** True si la persona aparece como supervisor en el tareo (DNI o nombre). */
function supervisorApareceEnTareo(statusMap, nombre, validatedRows, dni = "", codSupSet = null) {
  const dniKey = String(dni || "").replace(/\D/g, "");

  /* 1) Código supervisor del Excel (más confiable). */
  if (dniKey && codSupSet instanceof Set && codSupSet.has(dniKey)) return true;
  if (dniKey && statusMap?.byDni instanceof Map && statusMap.byDni.has(dniKey)) {
    const st = statusMap.byDni.get(dniKey);
    if (st && st.planillas > 0) return true;
  }

  const candidatos = [nombre].filter(Boolean);
  if (dniKey && Array.isArray(supervisoresLicapaCatalog)) {
    const cat = supervisoresLicapaCatalog.find(
      (s) => String(s.dni || "").replace(/\D/g, "") === dniKey
    );
    if (cat?.nombre) candidatos.push(cat.nombre);
  }

  for (let c = 0; c < candidatos.length; c += 1) {
    const st = lookupStatusSupervisor(statusMap, candidatos[c], dniKey);
    if (st && st.planillas > 0) return true;
  }

  const rows = validatedRows || [];
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    const cod = String(row?.codSupervisor || "").replace(/\D/g, "");
    if (dniKey && cod && cod === dniKey) return true;
    const sup = String(row?.supervisor || "").trim();
    if (!sup) continue;
    for (let c = 0; c < candidatos.length; c += 1) {
      if (
        nombresEquivalentes(candidatos[c], sup) ||
        nombreTokensKey(candidatos[c]) === nombreTokensKey(sup)
      ) {
        return true;
      }
    }
  }
  return false;
}

/** Quién no tareó como supervisor, agrupado por Supervisor General. */
function buildSinTareoPorGrupo(statusMap, personasApoyoHoy, validatedRows = null) {
  const grupos = Array.isArray(supervisoresGeneralesLicapa) ? supervisoresGeneralesLicapa : [];
  const apoyoHoy = personasApoyoHoyKeys(personasApoyoHoy);
  const f = getResumenFilterValues();
  const out = [];
  let nSin = 0;
  const codSupSet = collectCodSupervisoresTareo(validatedRows);

  const esApoyoNombre = (nombre) =>
    apoyoHoy.has(normNombreKpi(nombre)) ||
    [...apoyoHoy].some((k) => nombresEquivalentes(k, nombre));

  grupos.forEach((g) => {
    const generalNombre = String(g?.general?.nombre || "").trim();
    const generalDni = String(g?.general?.dni || "").replace(/\D/g, "");
    const miembros = Array.isArray(g?.miembros) ? g.miembros : [];
    const faltan = [];

    miembros.forEach((m) => {
      const dni = String(m?.dni || "").replace(/\D/g, "");
      const nombre = String(m?.nombre || "").trim();
      if (!nombre) return;
      if (
        f.supervisor &&
        !nombresEquivalentes(f.supervisor, nombre) &&
        normNombreKpi(f.supervisor) !== normNombreKpi(nombre)
      ) {
        return;
      }
      if (isSupervisorExcluidoValidacion({ dni }) || isSupervisorLicapaII({ dni })) return;

      const aliases = [
        nombre,
        ...(Array.isArray(m.alias) ? m.alias : []),
        ...(Array.isArray(m.aliases) ? m.aliases : [])
      ].filter(Boolean);

      const aparecio = aliases.some((alias) =>
        supervisorApareceEnTareo(statusMap, alias, validatedRows, dni, codSupSet)
      );
      if (aparecio) return;

      const st =
        aliases.map((a) => lookupStatusSupervisor(statusMap, a, dni)).find((x) => x) || null;
      const esApoyo = aliases.some((a) => esApoyoNombre(a));
      faltan.push({
        dni,
        nombre,
        total: st?.trabajadores ?? 0,
        cerraron: st?.cerradas ?? 0,
        soloManana: st?.soloManana ?? 0,
        estado: esApoyo ? "Sin tareo · Apoyo" : "Sin tareo",
        cerro: false,
        faltaManana: true,
        faltaTarde: true,
        sinTareo: true,
        esApoyo,
        generalNombre,
        generalDni
      });
    });

    if (!faltan.length) return;
    faltan.sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
    nSin += faltan.length;
    out.push({
      kind: "group",
      generalNombre,
      generalDni,
      count: faltan.length,
      miembros: faltan
    });
  });

  return { grupos: out, total: nSin };
}

function buildResumenPlanillaRows() {
  const { missing, present, statusMap, activeCatalog, personasApoyoHoy } = planillasFaltantesState;
  const apoyoHoy = personasApoyoHoyKeys(personasApoyoHoy);
  const toRow = (s, cerro) => {
    const st = statusMap?.get(normNombreKpi(s.nombre));
    const total = st?.trabajadores ?? s.trabajadores ?? 0;
    const cerraron = st?.cerradas ?? s.cerradas ?? 0;
    const soloManana = st?.soloManana ?? s.soloManana ?? 0;
    const nombreKey = normNombreKpi(s.nombre);
    /* Solo quienes aparecen como apoyo en ESTE tareo/filtro (nunca catálogo ni estado viejo). */
    const esApoyo = apoyoHoy.has(nombreKey);
    const sinGrupo = !st || !st.planillas;
    const tag = s.reactivadoHoy
      ? " · Supervisor hoy (Avísame)"
      : esApoyo
        ? " · Apoyo"
        : "";
    const faltaManana = supervisorFaltaCerrarManana(st);
    const faltaTarde = supervisorFaltaCerrarTarde(st);
    let estado = cerro ? `Cerró planilla${tag}` : `No cerró${tag}`;
    if (cerro && soloManana > 0) {
      estado = `Cerró · ${cerraron} ok / ${soloManana} temprano${tag}`;
    } else if (!cerro && esApoyo && sinGrupo) {
      estado = `Apoyo · sin grupo · falta mañana y tarde`;
    } else if (!cerro && esApoyo && faltaManana && faltaTarde) {
      estado = `Apoyo · falta cerrar mañana y tarde`;
    } else if (!cerro && esApoyo && faltaManana) {
      estado = `Apoyo · falta cerrar mañana`;
    } else if (!cerro && esApoyo && faltaTarde) {
      estado = `Apoyo · falta cerrar tarde`;
    } else if (!cerro && faltaManana && faltaTarde) {
      estado = `Falta cerrar mañana y tarde`;
    } else if (!cerro && faltaManana) {
      estado = `Falta cerrar mañana`;
    } else if (!cerro && faltaTarde) {
      estado = `Falta cerrar tarde`;
    }
    return {
      dni: s.dni,
      nombre: s.nombre,
      total,
      cerraron,
      soloManana,
      estado,
      cerro,
      bloqueManana: st?.bloqueManana ?? 0,
      bloqueTarde: st?.bloqueTarde ?? 0,
      faltaManana,
      faltaTarde,
      reactivadoHoy: Boolean(s.reactivadoHoy),
      esApoyo
    };
  };

  const scope = activeCatalog || [...(missing || []), ...(present || [])];
  const allRows = scope.map((s) => {
    const st = statusMap?.get(normNombreKpi(s.nombre));
    return toRow(s, st?.cerradas > 0);
  });

  const faltaManana = allRows
    .filter((row) => row.faltaManana)
    .sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
  const faltaTarde = allRows
    .filter((row) => row.faltaTarde)
    .sort((a, b) => b.soloManana - a.soloManana || a.nombre.localeCompare(b.nombre, "es"));
  const faltaAmbos = allRows
    .filter((row) => row.faltaManana && row.faltaTarde && !row.esApoyo)
    .sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
  const apoyo = allRows
    .filter((row) => row.esApoyo)
    .sort((a, b) => Number(a.cerro) - Number(b.cerro) || a.nombre.localeCompare(b.nombre, "es"));
  const cerraron = (present || [])
    .map((s) => toRow(s, true))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
  /* Todos = solo sin tareo (no aparecen como supervisor), agrupados por Supervisor General. */
  const sinTareo = allRows
    .filter((row) => !row.bloqueManana && !row.bloqueTarde && (row.total || 0) === 0 && !row.cerro)
    .map((row) => ({ ...row, estado: "Sin tareo", sinTareo: true }));

  return { faltaManana, faltaTarde, faltaAmbos, apoyo, cerraron, todos: sinTareo, sinTareo };
}

function badgeTurno(falta, when) {
  return falta
    ? `<span class="resumen-badge resumen-badge--warn">${escapeHtml(when)} · falta</span>`
    : `<span class="resumen-badge resumen-badge--ok">${escapeHtml(when)} · ok</span>`;
}

function badgeApoyo(esApoyo) {
  return esApoyo
    ? `<span class="resumen-badge resumen-badge--apoyo">Apoyo</span>`
    : "";
}

function faltantesPersonCard(s, n) {
  const tone = s.cerro ? "is-ok" : s.esApoyo ? "is-apoyo" : "is-warn";
  const flags = [
    badgeApoyo(s.esApoyo),
    badgeTurno(s.faltaManana, "Mañ"),
    badgeTurno(s.faltaTarde, "Tar")
  ]
    .filter(Boolean)
    .join("");
  return `<article class="rf-person ${tone}">
    <div class="rf-person__index">${n}</div>
    <div class="rf-person__body">
      <div class="rf-person__name">${escapeHtml(s.nombre)}</div>
      <div class="rf-person__estado">${escapeHtml(s.estado || "")}</div>
    </div>
    <div class="rf-person__flags">${flags}</div>
  </article>`;
}

function setFaltantesTabCounts(counts) {
  Object.entries(counts).forEach(([key, value]) => {
    const el = document.querySelector(`[data-faltantes-count="${key}"]`);
    if (el) el.textContent = String(value);
  });
}

function renderResumenFaltantesPanel() {
  const { present, reactivados, activeCatalog, filters, statusMap, personasApoyoHoy, validatedRows } =
    planillasFaltantesState;
  const { faltaManana, faltaTarde, faltaAmbos, apoyo, cerraron } = buildResumenPlanillaRows();
  const sinTareoGrupos = buildSinTareoPorGrupo(statusMap, personasApoyoHoy, validatedRows);

  const tab = planillasFaltantesResumenTab || "todos";
  const list =
    tab === "cerraron"
      ? cerraron
      : tab === "apoyo"
        ? apoyo
        : tab === "falta-ambos"
          ? faltaAmbos
          : tab === "falta-manana"
            ? faltaManana
            : tab === "falta-tarde"
              ? faltaTarde
              : null;

  setFaltantesTabCounts({
    todos: sinTareoGrupos.total,
    "falta-ambos": faltaAmbos.length,
    "falta-manana": faltaManana.length,
    "falta-tarde": faltaTarde.length,
    apoyo: apoyo.length,
    cerraron: cerraron.length
  });

  document.querySelectorAll("[data-faltantes-tab]").forEach((btn) => {
    btn.classList.toggle("is-active", btn.getAttribute("data-faltantes-tab") === tab);
  });

  const hint = document.getElementById("resumenFaltantesHint");
  if (hint) {
    const f = filters || getResumenFilterValues();
    const fundoFiltro = f.fundo || "";
    if (tab === "cerraron") {
      hint.textContent = `${cerraron.length} cerraron planilla`;
    } else if (tab === "apoyo") {
      hint.textContent = `${apoyo.length} en apoyo · estado mañana / tarde`;
    } else if (tab === "todos") {
      hint.textContent = `${sinTareoGrupos.total} sin tareo · agrupados por Supervisor General`;
    } else if (tab === "falta-ambos") {
      hint.textContent = list?.length
        ? `${list.length} sin cierre mañana ni tarde`
        : "Nadie falta mañana y tarde";
    } else if (tab === "falta-manana") {
      hint.textContent = list?.length
        ? `${list.length} sin cierre mañana · ${horarioMananaLabel(fundoFiltro)}`
        : "Todos cerraron mañana";
    } else if (tab === "falta-tarde") {
      hint.textContent = list?.length
        ? `${list.length} sin cierre tarde · ${horarioTardeLabel(fundoFiltro)}`
        : "Todos cerraron tarde";
    }
  }

  const btnAvisame = document.getElementById("btnResumenAvisame");
  const countAvisame = document.getElementById("resumenAvisameCount");
  const nRe = reactivados?.length || 0;
  if (btnAvisame) {
    btnAvisame.hidden = nRe === 0;
    if (countAvisame) {
      countAvisame.hidden = nRe === 0;
      countAvisame.textContent = String(nRe);
    }
  }

  const listEl = document.getElementById("resumenFaltantesList");
  const emptyEl = document.getElementById("resumenFaltantesEmpty");
  if (!listEl) return;

  if (tab === "todos") {
    if (!sinTareoGrupos.total) {
      listEl.innerHTML = "";
      if (emptyEl) {
        emptyEl.hidden = false;
        emptyEl.textContent = "Todos los del listado por Supervisor General ya tarearon.";
      }
      return;
    }
    if (emptyEl) emptyEl.hidden = true;
    let n = 0;
    listEl.innerHTML = sinTareoGrupos.grupos
      .map((g) => {
        const people = g.miembros
          .map((s) => {
            n += 1;
            return faltantesPersonCard(
              {
                ...s,
                faltaManana: true,
                faltaTarde: true,
                esApoyo: false,
                cerro: false,
                estado: s.estado || "Sin tareo"
              },
              n
            );
          })
          .join("");
        return `<section class="rf-group">
          <header class="rf-group__head">
            <div class="rf-group__copy">
              <span class="rf-group__label">Supervisor General</span>
              <strong class="rf-group__name">${escapeHtml(g.generalNombre)}</strong>
              ${g.generalDni ? `<span class="rf-group__dni">${escapeHtml(g.generalDni)}</span>` : ""}
            </div>
            <span class="rf-group__count">${g.count}</span>
          </header>
          <div class="rf-group__body">${people}</div>
        </section>`;
      })
      .join("");
    return;
  }

  if (!list.length) {
    listEl.innerHTML = "";
    if (emptyEl) {
      emptyEl.hidden = false;
      emptyEl.textContent =
        tab === "cerraron"
          ? "Nadie cerró planilla en este filtro."
          : tab === "apoyo"
            ? "Nadie figura como apoyo en este filtro."
            : tab === "falta-ambos"
              ? "Nadie sin Apoyo falta mañana y tarde."
              : tab === "falta-manana"
                ? `Ninguno de los ${(activeCatalog || []).length} activos falta mañana.`
                : `Ninguno de los ${(activeCatalog || []).length} activos falta tarde.`;
    }
    return;
  }

  if (emptyEl) emptyEl.hidden = true;
  listEl.innerHTML = `<div class="rf-group rf-group--flat"><div class="rf-group__body">${list
    .map((s, i) => faltantesPersonCard(s, i + 1))
    .join("")}</div></div>`;
}

async function loadAndRenderResumenFaltantes(getValidated) {
  const validated = getValidated?.();
  if (!validated?.rows?.length) {
    resetPlanillasFaltantesState();
    return;
  }
  const catalog = await loadSupervisoresLicapaCatalog();
  if (!catalog.length) return;
  await loadSupervisoresGeneralesLicapa();
  planillasFaltantesState = computePlanillasFaltantes(catalog, validated, getResumenFilterValues());
  // Re-pinta tabla principal (incluye No asistió / Está de apoyo)
  renderResumenView(validated, {});
}

async function refreshPlanillasFaltantesFromFilters(getValidated) {
  const modal = document.getElementById("modalResumen");
  if (!modal || modal.hidden) return;
  await loadAndRenderResumenFaltantes(getValidated);
}

async function openPlanillasFaltantesInResumen(getValidated) {
  planillasFaltantesResumenTab = "todos";
  await loadAndRenderResumenFaltantes(getValidated);
}

export function closeResumenModal() {
  const modal = document.getElementById("modalResumen");
  if (modal) modal.hidden = true;
  resetPlanillasFaltantesState();
}

function escapeAttr(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/"/g, "&quot;");
}

/**
 * Actividades disponibles según filtros (como Excel: al elegir Macro solo salen sus actividades).
 */
function fillResumenActividadOptions(validated, { supervisor = "", fundo = "", macro = "" } = {}) {
  const actSelect = document.getElementById("resumenFltActividad");
  if (!actSelect) return;

  const prev = actSelect.value;
  const activities = [
    ...new Set(
      (validated?.rows || [])
        .filter((row) => {
          if (supervisor && row.supervisor !== supervisor) return false;
          if (fundo && row.fundo !== fundo) return false;
          if (macro && row.macroPartida !== macro) return false;
          return true;
        })
        .map((r) => String(r.actividad || "").trim())
        .filter(Boolean)
    )
  ].sort((a, b) => a.localeCompare(b, "es"));

  actSelect.innerHTML = `<option value="">Actividad</option>${activities
    .map((v) => `<option value="${escapeAttr(v)}">${escapeAttr(v)}</option>`)
    .join("")}`;
  actSelect.value = activities.includes(prev) ? prev : "";
}

function fillResumenSelect(selectId, values, allLabel = "Todos") {
  const sel = document.getElementById(selectId);
  if (!sel) return;
  const prev = sel.value;
  const list = [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, "es"));
  sel.innerHTML = `<option value="">${escapeAttr(allLabel)}</option>${list
    .map((v) => `<option value="${escapeAttr(v)}">${escapeAttr(v)}</option>`)
    .join("")}`;
  sel.value = list.includes(prev) ? prev : "";
}

/** Últimos filtros del tareo al abrir Resumen (Supervisor / Fundo / actividades). */
let resumenMainFilters = {
  supervisor: "",
  fundo: "",
  actividades: []
};

function normActividadResumen(act) {
  return String(act || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function updateResumenFiltersSyncLabel() {
  const el = document.getElementById("resumenFiltersSync");
  if (!el) return;
  const parts = [];
  if (resumenMainFilters.supervisor) parts.push(`Supervisor: ${resumenMainFilters.supervisor}`);
  if (resumenMainFilters.fundo) parts.push(`Fundo: ${resumenMainFilters.fundo}`);
  const nAct = (resumenMainFilters.actividades || []).length;
  if (nAct) parts.push(nAct === 1 ? "1 actividad" : `${nAct} actividades`);
  el.textContent = parts.length
    ? `Filtros del tareo · ${parts.join(" · ")}`
    : "Filtros del tareo · Todos los supervisores y fundos";
}

/** Llena opciones y aplica Supervisor/Fundo/actividades del tareo (sin filtrar 2 veces). */
export function syncResumenFiltersFromMain(validated, mainFilters = {}) {
  const rows = validated?.rows || [];
  const f = mainFilters || {};

  resumenMainFilters = {
    supervisor: f.supervisor || "",
    fundo: f.fundo || "",
    actividades: Array.isArray(f.actividades) ? [...f.actividades] : []
  };

  fillResumenSelect(
    "resumenFltSupervisor",
    rows.map((r) => r.supervisor),
    "Todos"
  );
  fillResumenSelect(
    "resumenFltFundo",
    rows.map((r) => r.fundo),
    "Todos"
  );

  const macroSel = document.getElementById("resumenFltMacro");
  if (macroSel) {
    const macros = [
      ...new Set(rows.map((r) => String(r.macroPartida || "").trim()).filter(Boolean))
    ].sort((a, b) => a.localeCompare(b, "es"));
    macroSel.innerHTML = `<option value="">Macro Partida</option>${macros
      .map((v) => `<option value="${escapeAttr(v)}">${escapeAttr(v)}</option>`)
      .join("")}`;
    const prefer = f.macro || "COSTO DE COSECHA";
    macroSel.value = macros.includes(prefer) ? prefer : "";
  }

  const supEl = document.getElementById("resumenFltSupervisor");
  const fundoEl = document.getElementById("resumenFltFundo");
  const actEl = document.getElementById("resumenFltActividad");
  if (supEl) supEl.value = resumenMainFilters.supervisor;
  if (fundoEl) fundoEl.value = resumenMainFilters.fundo;

  fillResumenActividadOptions(validated, {
    supervisor: resumenMainFilters.supervisor,
    fundo: resumenMainFilters.fundo,
    macro: macroSel?.value || ""
  });

  if (actEl) {
    const acts = resumenMainFilters.actividades;
    if (acts.length === 1 && [...actEl.options].some((o) => o.value === acts[0])) {
      actEl.value = acts[0];
    } else {
      actEl.value = "";
    }
  }

  updateResumenFiltersSyncLabel();
}

/** Misma data que la tabla del modal Resumen (respeta tareo + Macro/Actividad del modal). */
export function getFilteredResumenData(validated, mainFilters = {}) {
  const supEl = document.getElementById("resumenFltSupervisor");
  const fundoEl = document.getElementById("resumenFltFundo");
  const macroEl = document.getElementById("resumenFltMacro");
  const actEl = document.getElementById("resumenFltActividad");

  const mf = { ...resumenMainFilters, ...(mainFilters || {}) };
  const supervisor = (supEl ? supEl.value : "") || mf.supervisor || "";
  const fundo = (fundoEl ? fundoEl.value : "") || mf.fundo || "";
  const macro = macroEl ? macroEl.value : mf.macro || "";

  fillResumenActividadOptions(validated, { supervisor, fundo, macro });

  const actividad = actEl ? actEl.value : mf.actividad || "";
  const actSet =
    !actividad && Array.isArray(mf.actividades) && mf.actividades.length
      ? new Set(mf.actividades.map((a) => normActividadResumen(a)).filter(Boolean))
      : null;

  const baseFiltered = (validated.rows || []).filter((row) => {
    if (supervisor && row.supervisor !== supervisor) return false;
    if (fundo && row.fundo !== fundo) return false;
    if (macro && row.macroPartida !== macro) return false;
    return true;
  });

  const filtered = baseFiltered.filter((row) => {
    if (actividad && String(row.actividad || "").trim() !== actividad) return false;
    if (actSet) {
      const key = normActividadResumen(row.actividad);
      if (!key || !actSet.has(key)) return false;
    }
    return true;
  });

  const dayRows = collapseToDayRows(filtered);
  /* Apoyo siempre del alcance actual (fundo/macro/supervisor), nunca de un Excel anterior. */
  const apoyoMap = buildApoyoSupervisionMap(baseFiltered);
  const { groups, kpis } = buildResumenGroups(dayRows, apoyoMap);
  return {
    groups,
    kpis,
    filtered,
    baseFiltered,
    dayRows,
    filters: {
      supervisor,
      fundo,
      macro,
      actividad,
      actividades: actSet ? [...mf.actividades] : []
    },
    macro,
    actividad
  };
}

export function renderResumenView(validated, mainFilters = {}) {
  if (!validated) return;

  const { groups, kpis, filtered, baseFiltered, actividad } =
    getFilteredResumenData(validated, mainFilters);
  const visibleGroups = filterGruposResumen(groups);
  const equiposBase = filterRowsResumenExcluidos(baseFiltered);

  const kpiHost = document.getElementById("resumenKpis");
  const actividadEsCalidad = isActividadCalidad(actividad);

  if (kpiHost) {
    if (actividadEsCalidad) {
      const calidadPersonas = countAuxiliarCalidad(filtered);
      const calidadSupervisores = countSupervisoresCalidad(equiposBase);
      kpiHost.innerHTML = `
      <div class="resumen-kpi resumen-kpi--calidad"><span class="resumen-kpi__label">Aux. Calidad</span><span class="resumen-kpi__value">${calidadPersonas}</span></div>
      <div class="resumen-kpi resumen-kpi--calidad"><span class="resumen-kpi__label">Sup. Calidad</span><span class="resumen-kpi__value">${calidadSupervisores}</span></div>
      <div class="resumen-kpi"><span class="resumen-kpi__label">Trabajadores</span><span class="resumen-kpi__value">${kpis.trabajadores}</span></div>
      <div class="resumen-kpi resumen-kpi--danger"><span class="resumen-kpi__label">Errores</span><span class="resumen-kpi__value">${kpis.errores}</span></div>
      <div class="resumen-kpi resumen-kpi--warn"><span class="resumen-kpi__label">Extras</span><span class="resumen-kpi__value">${kpis.avisos}</span></div>
    `;
    } else {
      const supervisoresKpi = visibleGroups.length;
      const scanerKpi = countScanerCosto(equiposBase);
      const cosechaKpi = countCosechaCosto(equiposBase);
      kpiHost.innerHTML = `
      <div class="resumen-kpi"><span class="resumen-kpi__label">Supervisores</span><span class="resumen-kpi__value">${supervisoresKpi}</span></div>
      <div class="resumen-kpi"><span class="resumen-kpi__label">Scaner</span><span class="resumen-kpi__value">${scanerKpi}</span></div>
      <div class="resumen-kpi"><span class="resumen-kpi__label">Cosecha</span><span class="resumen-kpi__value">${cosechaKpi}</span></div>
      <div class="resumen-kpi"><span class="resumen-kpi__label">Trabajadores</span><span class="resumen-kpi__value">${kpis.trabajadores}</span></div>
      <div class="resumen-kpi resumen-kpi--danger"><span class="resumen-kpi__label">Errores</span><span class="resumen-kpi__value">${kpis.errores}</span></div>
      <div class="resumen-kpi resumen-kpi--warn"><span class="resumen-kpi__label">Extras</span><span class="resumen-kpi__value">${kpis.avisos}</span></div>
    `;
    }
  }

  resumenUiState.allRows = buildResumenEnrichedRows(validated, visibleGroups, baseFiltered);
  paintResumenTable();
}

function resumenTurnoBadge(falta) {
  return falta
    ? `<span class="resumen-badge resumen-badge--warn">Falta</span>`
    : `<span class="resumen-badge resumen-badge--ok">OK</span>`;
}

function resumenEstadoBadge(row) {
  if (row.isAdminCosecha) {
    return `<span class="resumen-badge resumen-badge--admin">Admin cosecha</span>`;
  }
  const map = {
    ok: ["resumen-badge--ok", "OK"],
    "falta-tarde": ["resumen-badge--warn", "Falta tarde"],
    "falta-manana": ["resumen-badge--warn", "Falta mañana"],
    "no-subio": ["resumen-badge--warn", "No subió"],
    "no-asistio": ["resumen-badge--danger", "No asistió"],
    apoyo: ["resumen-badge--apoyo", "Está de apoyo"]
  };
  const [cls, label] = map[row.estadoKey] || ["resumen-badge--muted", row.estadoLabel || "—"];
  return `<span class="resumen-badge ${cls}">${escapeHtml(label)}</span>`;
}

function resolveResumenEstado({ sinPlanilla, esApoyo, faltaManana, faltaTarde, cerro }) {
  if (esApoyo && sinPlanilla) {
    return { estadoKey: "apoyo", estadoLabel: "Está de apoyo" };
  }
  if (sinPlanilla) {
    return { estadoKey: "no-asistio", estadoLabel: "No asistió" };
  }
  if (cerro && !faltaManana && !faltaTarde) {
    return { estadoKey: "ok", estadoLabel: "OK" };
  }
  if (faltaManana && faltaTarde) {
    return { estadoKey: "no-subio", estadoLabel: "No subió" };
  }
  if (faltaTarde) {
    return { estadoKey: "falta-tarde", estadoLabel: "Falta tarde" };
  }
  if (faltaManana) {
    return { estadoKey: "falta-manana", estadoLabel: "Falta mañana" };
  }
  return { estadoKey: "ok", estadoLabel: "OK" };
}

function attachSupervisorGeneral(row, dni = "") {
  const sg = resolveSupervisorGeneralMeta(row.supervisor, dni, row.fundo);
  const isAdmin = isAdminCosechaSupervisor(row.supervisor, dni);
  return {
    ...row,
    sgKey: sg.sgKey,
    sgLabel: sg.sgLabel,
    sgTone: sg.sgTone,
    generalNombre: sg.generalNombre,
    generalDni: sg.generalDni,
    isAdminCosecha: isAdmin
  };
}

/**
 * Filas de la tabla: grupos con planilla + apoyo sin grupo + catálogo sin tareo.
 */
function buildResumenEnrichedRows(validated, visibleGroups, baseFiltered) {
  const statusMap = getSupervisoresPlanillaStatus(validated);
  const personasApoyo =
    planillasFaltantesState.personasApoyoHoy instanceof Map && planillasFaltantesState.personasApoyoHoy.size
      ? planillasFaltantesState.personasApoyoHoy
      : collectPersonasApoyoHoy(baseFiltered || validated?.rows || []);
  const apoyoKeys = personasApoyoHoyKeys(personasApoyo);
  const seen = new Set();
  const rows = [];

  visibleGroups.forEach((g) => {
    const key = `${normNombreKpi(g.fundo)}||${normNombreKpi(g.supervisor)}`;
    seen.add(normNombreKpi(g.supervisor));
    const st = lookupStatusSupervisor(statusMap, g.supervisor);
    const sinPlanilla = !st || !(st.planillas > 0);
    const esApoyo = apoyoKeys.has(normNombreKpi(g.supervisor));
    const faltaManana = supervisorFaltaCerrarManana(st);
    const faltaTarde = supervisorFaltaCerrarTarde(st);
    const cerro = Boolean(st?.cerro || (st?.cerradas || 0) > 0);
    const estado = resolveResumenEstado({ sinPlanilla, esApoyo, faltaManana, faltaTarde, cerro });
    const dniGuess = (st?.dnis && st.dnis[0]) || "";
    rows.push(
      attachSupervisorGeneral(
        {
          fundo: g.fundo,
          supervisor: g.supervisor,
          planillas: g.planillas,
          trabajadores: g.trabajadores,
          errores: g.errores,
          avisos: g.avisos,
          apoyo: Boolean(g.apoyo),
          apoyoNombres: g.apoyoNombres || [],
          faltaManana,
          faltaTarde,
          esApoyo,
          ...estado,
          rowKind: "grupo",
          sortKey: key
        },
        dniGuess
      )
    );
  });

  // Apoyo sin fila propia
  personasApoyo.forEach((p, pKey) => {
    if (!pKey || seen.has(pKey)) return;
    if (isSupervisorExcluidoResumen(p.nombre)) return;
    seen.add(pKey);
    const st = lookupStatusSupervisor(statusMap, p.nombre, p.dni);
    const sinPlanilla = !st || !(st.planillas > 0);
    const faltaManana = supervisorFaltaCerrarManana(st);
    const faltaTarde = supervisorFaltaCerrarTarde(st);
    const cerro = Boolean(st?.cerro || (st?.cerradas || 0) > 0);
    const estado = resolveResumenEstado({
      sinPlanilla,
      esApoyo: true,
      faltaManana,
      faltaTarde,
      cerro
    });
    const fundo = p.fundo || st?.fundo || "LICAPA";
    rows.push(
      attachSupervisorGeneral(
        {
          fundo,
          supervisor: p.nombre,
          planillas: st?.planillas || 0,
          trabajadores: st?.trabajadores || 0,
          errores: 0,
          avisos: 0,
          apoyo: false,
          apoyoNombres: [],
          faltaManana,
          faltaTarde,
          esApoyo: true,
          ...estado,
          rowKind: "apoyo",
          sortKey: `${normNombreKpi(fundo)}||${pKey}`
        },
        p.dni || ""
      )
    );
  });

  // Catálogo activo sin tareo → No asistió
  const activeCatalog = planillasFaltantesState.activeCatalog || [];
  activeCatalog.forEach((s) => {
    const pKey = normNombreKpi(s.nombre);
    if (!pKey || seen.has(pKey)) return;
    if (isSupervisorExcluidoResumen(s.nombre)) return;
    if (apoyoKeys.has(pKey)) return;
    const st = lookupStatusSupervisor(statusMap, s.nombre, s.dni);
    if (st && st.planillas > 0) return;
    seen.add(pKey);
    const fundo = s.fundo || "LICAPA";
    rows.push(
      attachSupervisorGeneral(
        {
          fundo,
          supervisor: s.nombre,
          planillas: 0,
          trabajadores: 0,
          errores: 0,
          avisos: 0,
          apoyo: false,
          apoyoNombres: [],
          faltaManana: true,
          faltaTarde: true,
          esApoyo: false,
          estadoKey: "no-asistio",
          estadoLabel: "No asistió",
          rowKind: "faltante",
          sortKey: `${normNombreKpi(fundo)}||${pKey}`
        },
        s.dni || ""
      )
    );
  });

  rows.sort((a, b) => {
    // Administradores de cosecha siempre al final
    const aa = a.isAdminCosecha ? 1 : 0;
    const bb = b.isAdminCosecha ? 1 : 0;
    if (aa !== bb) return aa - bb;
    const pri = { "no-asistio": 0, apoyo: 1, "no-subio": 2, "falta-tarde": 3, "falta-manana": 4, ok: 5 };
    const pa = pri[a.estadoKey] ?? 9;
    const pb = pri[b.estadoKey] ?? 9;
    if (pa !== pb) return pa - pb;
    const ff = String(a.fundo).localeCompare(String(b.fundo), "es");
    if (ff) return ff;
    return String(a.supervisor).localeCompare(String(b.supervisor), "es");
  });

  return rows;
}

export function getResumenTableRowsForExport() {
  return getResumenVisibleRows();
}

function getResumenVisibleRows() {
  const tab = resumenUiState.tab || "todos";
  const q = String(resumenUiState.search || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();

  const filtered = (resumenUiState.allRows || []).filter((row) => {
    if (tab === "falta-tarde" && !row.faltaTarde) return false;
    if (tab === "falta-manana" && !row.faltaManana) return false;
    if (tab === "no-asistio" && row.estadoKey !== "no-asistio") return false;
    if (tab === "apoyo" && !(row.esApoyo || row.estadoKey === "apoyo")) return false;
    if (tab === "nuevo" && row.sgKey !== "nuevo") return false;
    if (tab === "otro-fundo" && row.sgKey !== "otro-fundo") return false;
    if (tab === "ok" && row.estadoKey !== "ok") return false;

    if (!q) return true;
    const blob = `${row.fundo} ${row.supervisor} ${row.sgLabel || ""} ${row.estadoLabel} ${(row.apoyoNombres || []).join(" ")}`
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
    return blob.includes(q);
  });

  // Mantener admins al final también tras filtrar
  return filtered.sort((a, b) => {
    const aa = a.isAdminCosecha ? 1 : 0;
    const bb = b.isAdminCosecha ? 1 : 0;
    return aa - bb;
  });
}

function updateResumenTabCounts() {
  const all = resumenUiState.allRows || [];
  const counts = {
    todos: all.length,
    "falta-tarde": all.filter((r) => r.faltaTarde).length,
    "falta-manana": all.filter((r) => r.faltaManana).length,
    "no-asistio": all.filter((r) => r.estadoKey === "no-asistio").length,
    apoyo: all.filter((r) => r.esApoyo || r.estadoKey === "apoyo").length,
    nuevo: all.filter((r) => r.sgKey === "nuevo").length,
    "otro-fundo": all.filter((r) => r.sgKey === "otro-fundo").length,
    ok: all.filter((r) => r.estadoKey === "ok").length
  };
  Object.entries(counts).forEach(([key, n]) => {
    const el = document.querySelector(`[data-resumen-count="${key}"]`);
    if (el) el.textContent = String(n);
  });
}

function paintResumenTable() {
  updateResumenTabCounts();
  const visible = getResumenVisibleRows();
  const body = document.getElementById("resumenTableBody");
  const hint = document.getElementById("resumenTableHint");
  const range = document.getElementById("resumenPagerRange");

  document.querySelectorAll("[data-resumen-tab]").forEach((btn) => {
    btn.classList.toggle("is-active", btn.getAttribute("data-resumen-tab") === (resumenUiState.tab || "todos"));
  });

  if (hint) {
    const tab = resumenUiState.tab || "todos";
    const labels = {
      todos: "Todos los supervisores del alcance",
      "falta-tarde": "Filas con tareo de tarde pendiente",
      "falta-manana": "Filas con tareo de mañana pendiente",
      "no-asistio": "No asistieron / no subieron planilla",
      apoyo: "Están de apoyo (sin grupo propio o marcados apoyo)",
      nuevo: "LICAPA nuevos · no están en lista de Supervisor General",
      "otro-fundo": "Vienen de LICAPA II o LICAPA III",
      ok: "Tareo mañana y tarde en orden"
    };
    hint.textContent = `${labels[tab] || ""} · ${visible.length} filas`;
  }

  if (body) {
    if (!visible.length) {
      body.innerHTML = `<tr><td colspan="11" class="resumen-empty">Sin filas para este filtro o búsqueda.</td></tr>`;
    } else {
      body.innerHTML = visible
        .map((g) => {
          const errClass = g.errores > 0 ? " is-cell-danger" : "";
          const warnClass = g.avisos > 0 ? " is-cell-warn" : "";
          const apoyoTxt = g.apoyo ? `Sí · ${(g.apoyoNombres || []).join(" · ")}` : "No";
          const apoyoClass = g.apoyo ? " is-cell-apoyo" : "";
          const sgClass =
            g.sgKey === "nuevo"
              ? " is-cell-warn"
              : g.sgKey === "otro-fundo"
                ? " is-cell-muted"
                : g.sgKey === "lista"
                  ? " is-cell-apoyo"
                  : "";
          const rowTone = g.isAdminCosecha
            ? "is-row-admin-cosecha"
            : g.estadoKey === "no-asistio" || g.errores > 0 || g.sgKey === "nuevo"
              ? "is-row-danger"
              : g.estadoKey === "apoyo"
                ? "is-row-apoyo"
                : g.estadoKey === "falta-tarde" || g.estadoKey === "falta-manana" || g.estadoKey === "no-subio" || g.avisos > 0
                  ? "is-row-warn"
                  : "";
          const supLabel = g.isAdminCosecha
            ? `${escapeHtml(g.supervisor)} <span class="resumen-admin-tag">Admin</span>`
            : escapeHtml(g.supervisor);
          return `<tr class="${rowTone}">
              <td>${escapeHtml(g.fundo)}</td>
              <td>${supLabel}</td>
              <td class="${sgClass}" title="${escapeAttr(g.generalDni ? `${g.sgLabel} · ${g.generalDni}` : g.sgLabel || "")}">${escapeHtml(g.sgLabel || "—")}</td>
              <td>${g.planillas}</td>
              <td>${g.trabajadores}</td>
              <td class="${errClass}">${g.errores}</td>
              <td class="${warnClass}">${g.avisos}</td>
              <td>${resumenTurnoBadge(g.faltaManana)}</td>
              <td>${resumenTurnoBadge(g.faltaTarde)}</td>
              <td>${resumenEstadoBadge(g)}</td>
              <td class="${apoyoClass}">${escapeHtml(apoyoTxt)}</td>
            </tr>`;
        })
        .join("");
    }
  }

  if (range) {
    range.textContent = `1 – ${visible.length} of ${visible.length}`;
  }
}

function shortSupervisorName(name) {
  const parts = String(name || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  if (parts.length <= 2) return parts.join(" ");
  // Nombre + 1er apellido (más legible en barras)
  return `${parts[0]} ${parts[1]}`;
}

function shortMacroLabel(label) {
  const raw = String(label || "").trim();
  if (!raw) return "—";
  const upper = raw.toUpperCase();
  const aliases = {
    "COSTO DE COSECHA": "Cosecha",
    "REMUNERACIONES ADMINISTRATIVAS": "Remuneraciones",
    LABORES: "Labores",
    SANIDAD: "Sanidad",
    RIEGO: "Riego",
    ALMACENES: "Almacenes"
  };
  if (aliases[upper]) return aliases[upper];
  if (raw.length <= 14) return raw;
  return `${raw.slice(0, 12)}…`;
}

/** % dentro de cada tramo del donut (todos los segmentos). */
const doughnutInsidePercent = {
  id: "doughnutInsidePercent",
  afterDatasetsDraw(chart) {
    if (chart.config.type !== "doughnut") return;
    const { ctx } = chart;
    const meta = chart.getDatasetMeta(0);
    if (!meta?.data?.length) return;
    const values = (chart.data.datasets[0]?.data || []).map((v) => Number(v) || 0);
    const total = values.reduce((a, b) => a + b, 0);
    if (!total) return;

    ctx.save();
    meta.data.forEach((arc, i) => {
      if (!arc || arc.hidden) return;
      const value = values[i];
      if (!value) return;
      const pct = (value / total) * 100;

      const props =
        typeof arc.getProps === "function"
          ? arc.getProps(["startAngle", "endAngle", "innerRadius", "outerRadius", "x", "y"], true)
          : {
              startAngle: arc.startAngle,
              endAngle: arc.endAngle,
              innerRadius: arc.innerRadius,
              outerRadius: arc.outerRadius,
              x: arc.x,
              y: arc.y
            };

      const mid = (props.startAngle + props.endAngle) / 2;
      const cos = Math.cos(mid);
      const sin = Math.sin(mid);
      // Tramos chicos: etiqueta más afuera del anillo para que se lea
      const t = pct >= 8 ? 0.5 : pct >= 3 ? 0.62 : 0.78;
      const r = props.innerRadius + (props.outerRadius - props.innerRadius) * t;
      const x = props.x + cos * r;
      const y = props.y + sin * r;

      const label =
        pct >= 10 ? `${Math.round(pct)}%` : pct >= 1 ? `${pct.toFixed(1)}%` : `${pct.toFixed(1)}%`;
      const fontSize = pct >= 25 ? 13 : pct >= 8 ? 11 : pct >= 3 ? 9 : 8;

      ctx.font = `700 ${fontSize}px 'Segoe UI', system-ui, sans-serif`;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      // Borde para contraste sobre colores claros u oscuros
      ctx.lineWidth = pct >= 8 ? 3 : 2.5;
      ctx.strokeStyle = "rgba(17, 24, 39, 0.45)";
      ctx.fillStyle = "#fff";
      ctx.strokeText(label, x, y);
      ctx.fillText(label, x, y);
    });
    ctx.restore();
  }
};

function renderResumenCharts(porPie, erroresPorSup, pieTitle = "Personas por Macro Partida") {
  const canvasMacro = document.getElementById("chartResumenMacro");
  const canvasEstado = document.getElementById("chartResumenEstado");
  if (!window.Chart) return;

  destroyChart(chartPersonas);
  destroyChart(chartErrores);
  chartPersonas = null;
  chartErrores = null;

  if (canvasMacro) {
    const labels = porPie.map((x) => x.label);
    const data = porPie.map((x) => x.value);
    const total = data.reduce((a, b) => a + (Number(b) || 0), 0) || 1;
    chartPersonas = new window.Chart(canvasMacro, {
      type: "doughnut",
      data: {
        labels,
        datasets: [
          {
            data,
            backgroundColor: labels.map((_, i) => PIE_COLORS[i % PIE_COLORS.length]),
            borderWidth: 2,
            borderColor: "#fff"
          }
        ]
      },
      plugins: [doughnutInsidePercent],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        radius: "90%",
        cutout: "48%",
        layout: {
          padding: { top: 0, right: 2, bottom: 0, left: 2 }
        },
        plugins: {
          legend: {
            position: "right",
            align: "center",
            labels: {
              boxWidth: 10,
              boxHeight: 10,
              font: { size: 11 },
              padding: 10,
              color: "#4b5563",
              generateLabels(chart) {
                const ds = chart.data.datasets[0];
                return chart.data.labels.map((label, i) => {
                  const val = Number(ds.data[i]) || 0;
                  const ratio = val / total;
                  const pct = (ratio * 100).toFixed(ratio >= 0.1 ? 0 : 1);
                  return {
                    text: `${shortMacroLabel(label)}  ${val}  (${pct}%)`,
                    fillStyle: ds.backgroundColor[i],
                    strokeStyle: "#fff",
                    lineWidth: 1,
                    hidden: false,
                    index: i
                  };
                });
              }
            }
          },
          title: {
            display: true,
            text: pieTitle,
            color: "#374151",
            font: { size: 12, weight: "600" },
            padding: { top: 0, bottom: 4 },
            align: "start"
          },
          tooltip: {
            callbacks: {
              label(ctx) {
                const val = Number(ctx.raw) || 0;
                const pct = ((val / total) * 100).toFixed(1);
                return ` ${ctx.label}: ${val} (${pct}%)`;
              }
            }
          }
        }
      }
    });
  }

  if (canvasEstado) {
    const labels = erroresPorSup.map((x) => shortSupervisorName(x.label));
    const fullNames = erroresPorSup.map((x) => x.label);
    const data = erroresPorSup.map((x) => x.value);
    chartErrores = new window.Chart(canvasEstado, {
      type: "bar",
      data: {
        labels: labels.length ? labels : ["Sin errores"],
        datasets: [
          {
            label: "Errores",
            data: data.length ? data : [0],
            backgroundColor: data.map((_, i) => ERROR_BAR_COLORS[i % ERROR_BAR_COLORS.length]),
            borderRadius: 6,
            maxBarThickness: 26,
            categoryPercentage: 0.72,
            barPercentage: 0.85
          }
        ]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        layout: {
          padding: { top: 6, right: 18, bottom: 4, left: 4 }
        },
        plugins: {
          legend: { display: false },
          title: {
            display: true,
            text: "Supervisores con más errores",
            color: "#374151",
            font: { size: 13, weight: "600" },
            padding: { top: 2, bottom: 14 }
          },
          tooltip: {
            callbacks: {
              title(items) {
                const i = items[0]?.dataIndex ?? 0;
                return fullNames[i] || labels[i] || "";
              },
              label(item) {
                return ` Errores: ${item.raw}`;
              }
            }
          }
        },
        scales: {
          x: {
            beginAtZero: true,
            ticks: { precision: 0 },
            grid: { color: "#f3f4f6" },
            afterFit(scale) {
              scale.paddingRight = 8;
            }
          },
          y: {
            ticks: {
              font: { size: 10 },
              autoSkip: false,
              crossAlign: "far"
            },
            grid: { display: false },
            afterFit(scale) {
              scale.width = Math.max(scale.width, 96);
            }
          }
        }
      }
    });
  }
}

function applySupervisorActivoFlags(row) {
  const dni = String(row?.dni ?? "").replace(/\D/g, "");
  if (SUPERVISORES_CALIDAD_DNI.has(dni)) {
    return { ...row, activo: false, nota: "calidad" };
  }
  if (SUPERVISORES_LICAPA_II_DNI.has(dni)) {
    return { ...row, activo: false, nota: "licapa_ii" };
  }
  if (SUPERVISORES_NO_ACTIVOS_DNI.has(dni)) {
    return { ...row, activo: false, nota: "no_activo" };
  }
  const nota = String(row?.nota || "").toLowerCase();
  /* Apoyo supervisión: activo, pero sin grupo propio → debe salir en Falta mañana/tarde */
  if (nota === "apoyo") {
    return { ...row, activo: true, nota: "apoyo" };
  }
  return { ...row, activo: true, nota: nota === "apoyo" ? "apoyo" : "" };
}

function isSupervisorApoyo(row) {
  return String(row?.nota || "").toLowerCase() === "apoyo";
}

function isSupervisorExcluidoValidacion(row) {
  const dni = String(row?.dni ?? "").replace(/\D/g, "");
  return (
    SUPERVISORES_EXCLUIDOS_VALIDACION_DNI.has(dni) || SUPERVISORES_LICAPA_II_DNI.has(dni)
  );
}

function isSupervisorLicapaII(row) {
  const dni = String(row?.dni ?? "").replace(/\D/g, "");
  return SUPERVISORES_LICAPA_II_DNI.has(dni);
}

function isSupervisorActivo(row) {
  if (!row) return false;
  const dni = String(row.dni ?? "").replace(/\D/g, "");
  if (SUPERVISORES_CALIDAD_DNI.has(dni) || SUPERVISORES_NO_ACTIVOS_DNI.has(dni)) return false;
  if (row.activo === false || row.activo === 0 || row.activo === "false") return false;
  const nota = String(row.nota || row.estado || "").toLowerCase();
  if (nota === "no_activo" || nota === "baja" || nota === "calidad" || nota === "duda" || nota === "licapa_ii") return false;
  return true;
}

function estadoSupervisorLabel(row) {
  const dni = String(row?.dni ?? "").replace(/\D/g, "");
  if (isSupervisorLicapaII(row)) return "LICAPA II (sin validar)";
  if (isSupervisorExcluidoValidacion(row)) return "Excluido (sin validar)";
  if (SUPERVISORES_CALIDAD_DNI.has(dni) || String(row?.nota || "").toLowerCase() === "calidad") {
    return "Área de Calidad";
  }
  if (!isSupervisorActivo(row)) return "Ya no trabaja";
  if (isSupervisorApoyo(row)) return "Activo · Apoyo (sin grupo propio)";
  return "Activo";
}

function normalizeSupervisorCatalogRow(row, defaultFundo = "LICAPA") {
  const nota = String(row?.nota ?? row?.estado ?? "").trim().toLowerCase();
  const dni = String(row?.dni ?? row?.DNI ?? "").replace(/\D/g, "");
  const activoRaw = row?.activo;
  let activo =
    activoRaw === false || activoRaw === 0 || activoRaw === "false"
      ? false
      : nota === "no_activo" || nota === "baja" || nota === "calidad" || nota === "duda" || nota === "licapa_ii"
        ? false
        : true;
  const base = {
    dni,
    nombre: String(row?.nombre ?? row?.Nombre ?? row?.name ?? "").trim(),
    fundo: String(row?.fundo ?? row?.FUNDO ?? defaultFundo).trim() || defaultFundo,
    activo,
    nota: nota || (activo ? "" : "no_activo")
  };
  return applySupervisorActivoFlags(base);
}

function mergeCatalogActivoFlags(catalog, flagsByDni) {
  const merged = (catalog || []).map((row) => {
    const dni = String(row.dni ?? "").replace(/\D/g, "");
    const flags = flagsByDni.get(dni);
    if (flags) {
      return applySupervisorActivoFlags({ ...row, ...flags });
    }
    return applySupervisorActivoFlags(row);
  });
  /* Agregar activos nuevos del JSON (ej. apoyos) que no estaban en localStorage */
  const have = new Set(merged.map((r) => String(r.dni || "").replace(/\D/g, "")).filter(Boolean));
  flagsByDni.forEach((flags, dni) => {
    if (!dni || have.has(dni)) return;
    if (flags.activo === false) return;
    const nombre = String(flags.nombre || "").trim();
    if (!nombre) return;
    merged.push(
      applySupervisorActivoFlags({
        dni,
        nombre,
        fundo: flags.fundo || "LICAPA",
        activo: true,
        nota: flags.nota || ""
      })
    );
  });
  return merged;
}

async function fetchSupervisoresLicapaFlags() {
  try {
    const res = await fetch(SUPERVISORES_LICAPA_URL);
    if (!res.ok) return new Map();
    const data = await res.json();
    const rows = Array.isArray(data) ? data : data?.supervisores;
    if (!Array.isArray(rows)) return new Map();
    const map = new Map();
    rows.forEach((row) => {
      const normalized = normalizeSupervisorCatalogRow(row, data?.fundo || "LICAPA");
      if (normalized.dni) {
        map.set(normalized.dni, {
          activo: normalized.activo,
          nota: normalized.nota,
          nombre: normalized.nombre,
          fundo: normalized.fundo
        });
      }
    });
    return map;
  } catch {
    return new Map();
  }
}

function readSupervisoresLicapaFromStorage() {
  try {
    const raw = localStorage.getItem(PLANILLAS_FALTANTES_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    const rows = Array.isArray(parsed) ? parsed : parsed?.supervisores;
    if (!Array.isArray(rows) || !rows.length) return null;
    return rows.map((row) => normalizeSupervisorCatalogRow(row)).filter((row) => row.nombre);
  } catch {
    return null;
  }
}

function saveSupervisoresLicapaToStorage(rows) {
  try {
    localStorage.setItem(PLANILLAS_FALTANTES_STORAGE_KEY, JSON.stringify(rows));
  } catch {
    /* ignore quota */
  }
}

async function loadSupervisoresLicapaCatalog() {
  const flagsByDni = await fetchSupervisoresLicapaFlags();
  const stored = readSupervisoresLicapaFromStorage();
  if (stored?.length) {
    supervisoresLicapaCatalog = mergeCatalogActivoFlags(stored, flagsByDni);
    return supervisoresLicapaCatalog;
  }
  try {
    const res = await fetch(SUPERVISORES_LICAPA_URL);
    if (!res.ok) throw new Error("HTTP " + res.status);
    const data = await res.json();
    const rows = Array.isArray(data) ? data : data?.supervisores;
    if (!Array.isArray(rows)) throw new Error("Formato inválido");
    supervisoresLicapaCatalog = rows
      .map((row) => normalizeSupervisorCatalogRow(row, data?.fundo || "LICAPA"))
      .filter((row) => row.nombre);
    return supervisoresLicapaCatalog;
  } catch (err) {
    console.warn("[planillas-faltantes] No se pudo cargar catálogo:", err);
    supervisoresLicapaCatalog = [];
    return [];
  }
}

function parseSupervisoresLicapaExcel(buffer) {
  if (typeof XLSX === "undefined") throw new Error("Lector Excel no disponible");
  const wb = XLSX.read(buffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!Array.isArray(row) || row.length < 2) continue;
    const dni = String(row[0] ?? "").replace(/\D/g, "");
    const nombre = String(row[1] ?? "").trim();
    if (!nombre || /^nombre/i.test(nombre)) continue;
    if (!dni && nombre.length < 4) continue;
    const col2 = String(row[2] ?? "").trim();
    const col3 = String(row[3] ?? "").trim().toLowerCase();
    let fundo = "LICAPA";
    let activo = true;
    let nota = "";
    const estadoCol = col3 || (/activo|calidad|inactivo|baja/i.test(col2) ? col2.toLowerCase() : "");
    if (/calidad/.test(estadoCol)) {
      activo = false;
      nota = "calidad";
    } else if (/no\s*activo|inactivo|baja/.test(estadoCol)) {
      activo = false;
      nota = "no_activo";
    }
    if (col2 && !/activo|calidad|inactivo|baja/i.test(col2)) fundo = col2;
    out.push(
      applySupervisorActivoFlags({
        dni,
        nombre,
        fundo,
        activo,
        nota
      })
    );
  }
  if (!out.length) {
    throw new Error("No se encontraron supervisores en el Excel (columnas A: DNI, B: Nombre).");
  }
  return out;
}

/** Comparativa por supervisor: total trabajadores vs cuántos cerraron jornada.
 *  Algunos pueden salir temprano/enfermos (solo mañana) — eso es normal.
 *  El supervisor CERRÓ si al menos 1 trabajador tiene jornada completa (mañana+tarde).
 *  NO cerró si TODOS sus trabajadores tienen solo 06:30–12:00. */
function getSupervisoresPlanillaStatus(validated) {
  const dayRows = collapseToDayRows(validated?.rows || []);
  const byKey = new Map();
  const byDni = new Map();

  dayRows.forEach((row) => {
    if (!row.supervisor) return;
    const key = normNombreKpi(row.supervisor);
    if (!byKey.has(key)) {
      byKey.set(key, {
        nombre: row.supervisor,
        fundo: row.fundo || "—",
        planillas: 0,
        cerradas: 0,
        soloManana: 0,
        bloqueManana: 0,
        bloqueTarde: 0,
        trabajadores: new Set(),
        dnis: new Set()
      });
    }
    const g = byKey.get(key);
    g.planillas += 1;
    if (row.documento) g.trabajadores.add(String(row.documento));
    const codSup = String(row.codSupervisor || "").replace(/\D/g, "");
    if (codSup) g.dnis.add(codSup);

    const cerro = dayRowCerroPlanilla(row);
    const manana = dayRowTieneBloqueManana(row);
    const tarde = dayRowTieneBloqueTarde(row);

    if (cerro) g.cerradas += 1;
    else if (manana && !tarde) g.soloManana += 1;
    if (manana) g.bloqueManana += 1;
    if (tarde) g.bloqueTarde += 1;
  });

  byKey.forEach((g, key) => {
    const totalTrab = g.trabajadores.size;
    const cerro = g.cerradas > 0;
    const entry = {
      nombre: g.nombre,
      fundo: g.fundo,
      planillas: g.planillas,
      cerradas: g.cerradas,
      soloManana: g.soloManana,
      bloqueManana: g.bloqueManana,
      bloqueTarde: g.bloqueTarde,
      trabajadores: totalTrab,
      cerro,
      pctCerrado: totalTrab ? Math.round((g.cerradas / g.planillas) * 100) : 0,
      dnis: [...(g.dnis || [])]
    };
    byKey.set(key, entry);
    (g.dnis || []).forEach((dni) => {
      const prev = byDni.get(dni);
      if (!prev || (entry.planillas || 0) > (prev.planillas || 0)) {
        byDni.set(dni, entry);
      }
    });
  });

  byKey.byDni = byDni;
  return byKey;
}

/** DNIs que aparecen como código de supervisor en el tareo. */
function collectCodSupervisoresTareo(rows) {
  const set = new Set();
  (rows || []).forEach((row) => {
    const d = String(row.codSupervisor || "").replace(/\D/g, "");
    if (d) set.add(d);
  });
  return set;
}

function isSupervisorCalidad(row) {
  const dni = String(row?.dni ?? "").replace(/\D/g, "");
  return SUPERVISORES_CALIDAD_DNI.has(dni) || String(row?.nota || "").toLowerCase() === "calidad";
}

function getResumenFilterValues() {
  return {
    supervisor: document.getElementById("resumenFltSupervisor")?.value || "",
    fundo: document.getElementById("resumenFltFundo")?.value || "",
    macro: document.getElementById("resumenFltMacro")?.value || "",
    actividad: document.getElementById("resumenFltActividad")?.value || ""
  };
}

/** Aplica filtros del Resumen al tareo (mismo criterio que la tabla izquierda). */
function filterRowsByResumen(validated, filters) {
  const f = filters || getResumenFilterValues();
  const rows = (validated?.rows || []).filter((row) => {
    if (f.supervisor && row.supervisor !== f.supervisor) return false;
    if (f.fundo && row.fundo !== f.fundo) return false;
    if (f.macro && row.macroPartida !== f.macro) return false;
    if (f.actividad && String(row.actividad || "").trim() !== f.actividad) return false;
    return true;
  });
  return { ...(validated || {}), rows };
}

function computePlanillasFaltantes(catalog, validated, filters) {
  const f = filters || getResumenFilterValues();
  const hasScopeFilter = Boolean(f.supervisor || f.fundo || f.macro || f.actividad);

  /* Cierre mañana/tarde: siempre con todo el tareo (jornada completa del día). */
  const statusMap = getSupervisoresPlanillaStatus(validated);

  /* Misma data que la tabla izquierda del Resumen (respeta filtros). */
  const filteredScope = filterRowsByResumen(validated, f);
  const supervisorsInScope = new Set();
  (filteredScope.rows || []).forEach((row) => {
    const key = normNombreKpi(row.supervisor);
    if (key) supervisorsInScope.add(key);
  });

  /* Apoyo = actividad SUPERVISOR DE COSECHA: no filtrar por Actividad COSECHA/SCANER.
   * Solo fundo / macro / supervisor, y únicamente bajo supervisores del alcance filtrado. */
  const rowsForApoyo = (validated?.rows || []).filter((row) => {
    if (f.supervisor && row.supervisor !== f.supervisor) return false;
    if (f.fundo && row.fundo !== f.fundo) return false;
    if (f.macro && row.macroPartida !== f.macro) return false;
    return true;
  });
  const personasApoyoHoy = collectPersonasApoyoHoy(
    rowsForApoyo,
    hasScopeFilter ? supervisorsInScope : null
  );

  const catalogFiltered = (catalog || []).filter((s) => {
    const fundoCat = String(s.fundo || "LICAPA").toUpperCase();
    if (f.fundo) {
      return fundoCat === String(f.fundo).toUpperCase() || fundoCat === "—" || !s.fundo;
    }
    return fundoCat === "LICAPA" || fundoCat === "—" || !s.fundo;
  });

  const calidad = catalogFiltered.filter((s) => isSupervisorCalidad(s));
  const noActivosBase = catalogFiltered.filter(
    (s) => !isSupervisorActivo(s) && !isSupervisorCalidad(s)
  );
  let activeBase = catalogFiltered.filter((s) => isSupervisorActivo(s));
  if (f.supervisor) {
    const key = normNombreKpi(f.supervisor);
    activeBase = activeBase.filter((s) => normNombreKpi(s.nombre) === key);
  }
  /* Con filtros: solo supervisores que salen en la tabla filtrada (o son su apoyo). */
  if (hasScopeFilter) {
    const apoyoKeys = personasApoyoHoyKeys(personasApoyoHoy);
    activeBase = activeBase.filter((s) => {
      const key = normNombreKpi(s.nombre);
      return supervisorsInScope.has(key) || apoyoKeys.has(key);
    });
  }

  /* Avísame: ya no trabajan en catálogo, pero HOY subieron tareo → se tratan como supervisores.
   * Excluidos temporalmente: nunca reactivar ni validar aunque aparezcan en data. */
  let reactivados = noActivosBase
    .filter((s) => !isSupervisorExcluidoValidacion(s))
    .filter((s) => (statusMap.get(normNombreKpi(s.nombre))?.planillas || 0) > 0)
    .map((s) => ({ ...s, reactivadoHoy: true, activoHoy: true }));
  if (f.supervisor) {
    const key = normNombreKpi(f.supervisor);
    reactivados = reactivados.filter((s) => normNombreKpi(s.nombre) === key);
  }
  if (hasScopeFilter) {
    reactivados = reactivados.filter((s) => supervisorsInScope.has(normNombreKpi(s.nombre)));
  }
  const reactivadoKeys = new Set(reactivados.map((s) => normNombreKpi(s.nombre)));

  /* No activos restantes (no aparecieron hoy) + calidad */
  const inactive = [
    ...noActivosBase.filter((s) => !reactivadoKeys.has(normNombreKpi(s.nombre))),
    ...calidad
  ];

  /* Activos del alcance + apoyos dinámicos del alcance (flujo filtrado). */
  const haveKeys = new Set(
    [...activeBase, ...reactivados].map((s) => normNombreKpi(s.nombre)).filter(Boolean)
  );
  const apoyosDinamicos = [];
  personasApoyoHoy.forEach((persona, key) => {
    if (!key || haveKeys.has(key)) return;
    if (isSupervisorExcluidoValidacion(persona)) return;
    apoyosDinamicos.push({ ...persona, apoyoHoy: true });
    haveKeys.add(key);
  });
  const markApoyo = (s) =>
    personasApoyoHoy.has(normNombreKpi(s.nombre)) ? { ...s, apoyoHoy: true } : s;
  const activeCatalog = [
    ...activeBase.map(markApoyo),
    ...reactivados.map(markApoyo),
    ...apoyosDinamicos
  ];

  const present = activeCatalog.filter((s) => {
    const st = statusMap.get(normNombreKpi(s.nombre));
    return st?.cerradas > 0;
  });
  const missing = activeCatalog
    .filter((s) => {
      const st = statusMap.get(normNombreKpi(s.nombre));
      return !st || st.cerradas === 0;
    })
    .map((s) => {
      const st = statusMap.get(normNombreKpi(s.nombre));
      return {
        ...s,
        motivo: motivoPlanillaFaltante(st),
        trabajadores: st?.trabajadores ?? 0,
        cerradas: st?.cerradas ?? 0,
        soloManana: st?.soloManana ?? 0,
        planillas: st?.planillas ?? 0
      };
    })
    .sort((a, b) => (b.soloManana || 0) - (a.soloManana || 0) || a.nombre.localeCompare(b.nombre, "es"));

  return {
    missing,
    present,
    inactive,
    reactivados,
    statusMap,
    catalog: catalogFiltered,
    activeCatalog,
    personasApoyoHoy,
    validatedRows: validated?.rows || [],
    filters: f
  };
}

function renderReactivadosBanner(reactivados) {
  const btn = document.getElementById("btnPlanillasAvisame");
  const countEl = document.getElementById("planillasAvisameCount");
  if (!btn) return;
  const n = reactivados?.length || 0;
  if (!n) {
    btn.hidden = true;
    if (countEl) {
      countEl.hidden = true;
      countEl.textContent = "";
    }
    return;
  }
  btn.hidden = false;
  if (countEl) {
    countEl.hidden = false;
    countEl.textContent = String(n);
  }
}

function hideAvisameToast() {
  const el = document.getElementById("avisameToast");
  if (el) {
    el.classList.remove("is-open");
    el.hidden = true;
  }
}

function showAvisameToast(reactivados, anchorBtn) {
  let el = document.getElementById("avisameToast");
  if (!el) {
    el = document.createElement("div");
    el.id = "avisameToast";
    el.className = "avisame-toast";
    el.setAttribute("role", "status");
    document.body.appendChild(el);
  }

  const list = reactivados || [];
  if (!list.length) {
    el.innerHTML = `
      <div class="avisame-toast__head">
        <strong>Avísame</strong>
        <button type="button" class="avisame-toast__close" aria-label="Cerrar">×</button>
      </div>
      <p class="avisame-toast__empty">Ningún “ya no trabaja” apareció en el tareo. Correcto.</p>
    `;
  } else {
    el.innerHTML = `
      <div class="avisame-toast__head">
        <strong>Avísame — Ya son supervisores hoy</strong>
        <button type="button" class="avisame-toast__close" aria-label="Cerrar">×</button>
      </div>
      <p class="avisame-toast__note">Estaban como “ya no trabajan”, pero subieron tareo: se cuentan como supervisores (Faltan / Cerraron).</p>
      <ul class="avisame-toast__list">
        ${list
          .map(
            (s) => `
          <li>
            <span class="avisame-toast__name">${escapeHtml(s.nombre)}</span>
            <span class="avisame-toast__dni">(${escapeHtml(s.dni || "—")})</span>
            <span class="avisame-toast__badge">→ SUPERVISOR</span>
          </li>`
          )
          .join("")}
      </ul>
    `;
  }

  el.hidden = false;
  el.classList.add("is-open");

  // Posicionar cerca del botón si existe
  if (anchorBtn?.getBoundingClientRect) {
    const r = anchorBtn.getBoundingClientRect();
    const pad = 8;
    const width = Math.min(360, window.innerWidth - 24);
    let left = Math.min(r.right - width, window.innerWidth - width - 12);
    left = Math.max(12, left);
    let top = r.bottom + pad;
    if (top + 220 > window.innerHeight) top = Math.max(12, r.top - pad - 200);
    el.style.left = `${left}px`;
    el.style.top = `${top}px`;
    el.style.right = "auto";
  } else {
    el.style.left = "auto";
    el.style.right = "1.25rem";
    el.style.top = "1.25rem";
  }

  el.querySelector(".avisame-toast__close")?.addEventListener("click", hideAvisameToast, { once: true });

  clearTimeout(showAvisameToast._timer);
  showAvisameToast._timer = setTimeout(hideAvisameToast, 8000);
}

function alertReactivadosNoActivos(reactivados, anchorBtn) {
  showAvisameToast(reactivados, anchorBtn);
}

function renderPlanillasFaltantesModal() {
  const { missing, present, inactive, reactivados, catalog, activeCatalog } = planillasFaltantesState;
  const q = (document.getElementById("planillasFaltantesSearch")?.value || "").trim().toLowerCase();
  const matchQ = (s) => !q || s.nombre.toLowerCase().includes(q) || String(s.dni).includes(q);

  renderReactivadosBanner(reactivados || []);

  const filtered = missing.filter(matchQ);
  const inactiveFiltered = (inactive || []).filter(matchQ);
  const activeFiltered = (activeCatalog || []).filter(matchQ);

  const kpiEl = document.getElementById("planillasFaltantesKpis");
  if (kpiEl) {
    kpiEl.innerHTML = `
      <div class="resumen-kpi resumen-kpi--neutral"><span class="resumen-kpi__label">Activos</span><span class="resumen-kpi__value">${(activeCatalog || []).length}</span></div>
      <div class="resumen-kpi resumen-kpi--ok"><span class="resumen-kpi__label">Con planilla</span><span class="resumen-kpi__value">${present.length}</span></div>
      <div class="resumen-kpi resumen-kpi--warn"><span class="resumen-kpi__label">Faltan cerrar</span><span class="resumen-kpi__value">${missing.length}</span></div>
      <div class="resumen-kpi resumen-kpi--neutral"><span class="resumen-kpi__label">Ya no trabajan</span><span class="resumen-kpi__value">${(inactive || []).length}</span></div>
    `;
  }

  document.querySelectorAll("[data-planillas-tab]").forEach((btn) => {
    btn.classList.toggle("is-active", btn.getAttribute("data-planillas-tab") === planillasFaltantesView);
  });

  let rowsToShow = [];
  let countText = "";
  const showComparativa =
    planillasFaltantesView === "faltantes" || planillasFaltantesView === "activos";

  const thead = document.querySelector("#modalPlanillasFaltantes .errores-modal__table thead tr");
  if (thead) {
    thead.innerHTML = showComparativa
      ? `<th>DNI</th><th>Supervisor</th><th>Total</th><th>Cerraron</th><th>Solo mañana</th><th>Estado</th>`
      : `<th>DNI</th><th>Supervisor</th><th>Estado</th>`;
  }

  if (planillasFaltantesView === "faltantes") {
    rowsToShow = filtered
      .map((s) => ({
        dni: s.dni,
        nombre: s.nombre,
        total: s.trabajadores ?? 0,
        cerraron: s.cerradas ?? 0,
        soloManana: s.soloManana ?? 0,
        estado: s.motivo || "Sin planilla cerrada",
        kind: "missing"
      }))
      .sort((a, b) => b.soloManana - a.soloManana || a.nombre.localeCompare(b.nombre, "es"));
    countText = filtered.length
      ? `${filtered.length} activo${filtered.length === 1 ? "" : "s"} sin cerrar (comparativa por trabajadores)`
      : missing.length
        ? "Sin coincidencias en la búsqueda"
        : "Todos los activos cerraron planilla";
  } else if (planillasFaltantesView === "inactivos") {
    const reactivadoKeys = new Set(
      (reactivados || []).map((s) => normNombreKpi(s.nombre))
    );
    rowsToShow = inactiveFiltered.map((s) => {
      const enData = reactivadoKeys.has(normNombreKpi(s.nombre));
      return {
        dni: s.dni,
        nombre: s.nombre,
        estado: isSupervisorLicapaII(s)
          ? "LICAPA II (sin validar)"
          : isSupervisorExcluidoValidacion(s)
          ? "Excluido temporalmente (sin validar)"
          : isSupervisorCalidad(s)
          ? "Área de Calidad"
          : enData
            ? "Ya no trabaja · Apareció en tareo (no debería)"
            : "Ya no trabaja",
        kind: "inactive"
      };
    });
    countText = `${inactiveFiltered.length} ya no trabajan (no deben subir tareo)`;
  } else {
    const apoyoHoy = personasApoyoHoyKeys(planillasFaltantesState.personasApoyoHoy);
    rowsToShow = activeFiltered
      .map((s) => {
        const st = planillasFaltantesState.statusMap?.get(normNombreKpi(s.nombre));
        const esApoyo = apoyoHoy.has(normNombreKpi(s.nombre));
        const tagApoyo = esApoyo ? " · Apoyo" : "";
        let estado = `Activo · Sin tareo${tagApoyo}`;
        if (st?.cerro) {
          estado =
            st.soloManana > 0
              ? `Cerró · ${st.cerradas} ok / ${st.soloManana} temprano${tagApoyo}`
              : `Activo · Planilla cerrada${tagApoyo}`;
        } else if (st?.soloManana) {
          estado = `Activo · Nadie cerró tarde${tagApoyo}`;
        } else if (st) {
          estado = `Activo · Sin cierre${tagApoyo}`;
        } else if (esApoyo) {
          estado = "Apoyo · sin grupo · falta mañana y tarde";
        }
        return {
          dni: s.dni,
          nombre: s.nombre,
          total: st?.trabajadores ?? 0,
          cerraron: st?.cerradas ?? 0,
          soloManana: st?.soloManana ?? 0,
          estado,
          kind: st?.cerro ? "active" : "missing"
        };
      })
      .sort((a, b) => {
        const ac = a.kind === "missing" ? 0 : 1;
        const bc = b.kind === "missing" ? 0 : 1;
        if (ac !== bc) return ac - bc;
        return b.soloManana - a.soloManana || a.nombre.localeCompare(b.nombre, "es");
      });
    countText = `${activeFiltered.length} supervisor${activeFiltered.length === 1 ? "" : "es"} activos (sin cerrar primero)`;
  }

  const countEl = document.getElementById("planillasFaltantesCount");
  if (countEl) countEl.textContent = countText;

  const tbody = document.getElementById("planillasFaltantesBody");
  const emptyEl = document.getElementById("planillasFaltantesEmpty");
  if (!tbody) return;

  if (!rowsToShow.length) {
    tbody.innerHTML = "";
    if (emptyEl) {
      emptyEl.hidden = false;
      if (planillasFaltantesView === "faltantes") {
        emptyEl.textContent = missing.length
          ? "Sin coincidencias en la búsqueda."
          : "Todos los supervisores activos tienen planilla en el tareo cargado.";
      } else if (planillasFaltantesView === "inactivos") {
        emptyEl.textContent = "No hay supervisores marcados como ya no trabajan.";
      } else {
        emptyEl.textContent = "No hay supervisores activos en el listado.";
      }
    }
    return;
  }

  if (emptyEl) emptyEl.hidden = true;
  tbody.innerHTML = rowsToShow
    .map((s) => {
      if (showComparativa) {
        return `
    <tr class="${s.kind === "missing" ? "is-row-warn" : ""}">
      <td class="errores-modal__dni">${escapeHtml(s.dni || "—")}</td>
      <td>${escapeHtml(s.nombre)}</td>
      <td>${s.total ?? 0}</td>
      <td>${s.cerraron ?? 0}</td>
      <td>${s.soloManana ?? 0}</td>
      <td class="errores-modal__motivo">${escapeHtml(s.estado)}</td>
    </tr>`;
      }
      return `
    <tr>
      <td class="errores-modal__dni">${escapeHtml(s.dni || "—")}</td>
      <td>${escapeHtml(s.nombre)}</td>
      <td class="errores-modal__motivo">${escapeHtml(s.estado)}</td>
    </tr>`;
    })
    .join("");
}

async function openPlanillasFaltantesModal(getValidated, getMainFilters) {
  const modal = document.getElementById("modalPlanillasFaltantes");
  if (!modal) return;

  const validated = getValidated?.();
  if (!validated?.rows?.length) {
    alert("Primero cargue el Excel de tareo para comparar planillas.");
    return;
  }

  const catalog = await loadSupervisoresLicapaCatalog();
  if (!catalog.length) {
    alert(
      "No hay listado de supervisores LICAPA. Use «Actualizar listado» o coloque data/supervisores-licapa.json."
    );
    return;
  }

  planillasFaltantesState = computePlanillasFaltantes(catalog, validated, getResumenFilterValues());
  planillasFaltantesView = "faltantes";
  const search = document.getElementById("planillasFaltantesSearch");
  if (search) search.value = "";
  renderPlanillasFaltantesModal();
  modal.hidden = false;
}

function closePlanillasFaltantesModal() {
  const modal = document.getElementById("modalPlanillasFaltantes");
  if (modal) modal.hidden = true;
}

function exportPlanillasFaltantesExcel() {
  const { missing, inactive } = planillasFaltantesState;
  if (!missing.length && !(inactive || []).length) {
    alert("No hay supervisores faltantes ni no activos para exportar.");
    return;
  }
  if (typeof XLSX === "undefined") {
    alert("Exportación Excel no disponible.");
    return;
  }
  const rows = [
    ["DNI", "Supervisor", "Fundo", "Total trab.", "Cerraron", "Solo mañana", "Estado"],
    ...missing.map((s) => [
      s.dni || "",
      s.nombre,
      s.fundo || "LICAPA",
      s.trabajadores ?? 0,
      s.cerradas ?? 0,
      s.soloManana ?? 0,
      s.motivo || "Sin planilla cerrada"
    ]),
    ...(inactive || []).map((s) => [
      s.dni || "",
      s.nombre,
      s.fundo || "LICAPA",
      "",
      "",
      "",
      estadoSupervisorLabel(s)
    ])
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Faltan planilla");
  const date = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `planillas-faltantes-licapa-${date}.xlsx`);
}

async function onPlanillasFaltantesListUpload(file, getValidated, getMainFilters) {
  if (!file) return;
  try {
    const buffer = await file.arrayBuffer();
    const rows = parseSupervisoresLicapaExcel(buffer);
    saveSupervisoresLicapaToStorage(rows);
    supervisoresLicapaCatalog = rows;
    const validated = getValidated?.();
    if (validated?.rows?.length) {
      planillasFaltantesState = computePlanillasFaltantes(rows, validated);
    } else {
      const inactive = rows.filter((s) => !isSupervisorActivo(s));
      const activeCatalog = rows.filter((s) => isSupervisorActivo(s));
      planillasFaltantesState = {
        missing: activeCatalog,
        present: [],
        inactive,
        reactivados: [],
        catalog: rows,
        activeCatalog,
        statusMap: new Map()
      };
    }
    renderPlanillasFaltantesModal();
    alert(`Listado actualizado: ${rows.length} supervisores.`);
  } catch (err) {
    alert("Error al leer Excel: " + (err.message || err));
  }
}

export function bindResumenUi({ getValidated, getMainFilters, onExport }) {
  document.getElementById("btnCloseResumen")?.addEventListener("click", closeResumenModal);
  document.querySelectorAll('[data-close-modal="resumen"]').forEach((el) => {
    el.addEventListener("click", closeResumenModal);
  });

  document.getElementById("btnClosePlanillasFaltantes")?.addEventListener("click", closePlanillasFaltantesModal);
  document.getElementById("btnClosePlanillasFaltantes2")?.addEventListener("click", closePlanillasFaltantesModal);
  document.querySelectorAll('[data-close-modal="planillas-faltantes"]').forEach((el) => {
    el.addEventListener("click", closePlanillasFaltantesModal);
  });

  const refresh = () => {
    const validated = getValidated?.();
    if (!validated) return;
    renderResumenView(validated, getMainFilters?.() || {});
    refreshPlanillasFaltantesFromFilters(getValidated);
  };

  ["resumenFltSupervisor", "resumenFltFundo", "resumenFltMacro", "resumenFltActividad"].forEach((id) => {
    document.getElementById(id)?.addEventListener("change", refresh);
  });

  document.getElementById("btnResumenExport")?.addEventListener("click", () => {
    onExport?.();
  });

  document.getElementById("resumenSearch")?.addEventListener("input", (e) => {
    resumenUiState.search = e.target?.value || "";
    paintResumenTable();
  });

  document.querySelectorAll("[data-resumen-tab]").forEach((btn) => {
    btn.addEventListener("click", () => {
      resumenUiState.tab = btn.getAttribute("data-resumen-tab") || "todos";
      paintResumenTable();
    });
  });

  document.getElementById("btnResumenAvisame")?.addEventListener("click", (e) => {
    alertReactivadosNoActivos(planillasFaltantesState.reactivados || [], e.currentTarget);
  });

  document.querySelectorAll("[data-faltantes-tab]").forEach((btn) => {
    btn.addEventListener("click", () => {
      planillasFaltantesResumenTab = btn.getAttribute("data-faltantes-tab") || "todos";
      renderResumenFaltantesPanel();
    });
  });

  document.getElementById("planillasFaltantesSearch")?.addEventListener("input", renderPlanillasFaltantesModal);

  document.querySelectorAll("[data-planillas-tab]").forEach((btn) => {
    btn.addEventListener("click", () => {
      planillasFaltantesView = btn.getAttribute("data-planillas-tab") || "faltantes";
      renderPlanillasFaltantesModal();
    });
  });

  document.getElementById("btnPlanillasFaltantesExport")?.addEventListener("click", exportPlanillasFaltantesExcel);

  document.getElementById("btnPlanillasAvisame")?.addEventListener("click", (e) => {
    alertReactivadosNoActivos(planillasFaltantesState.reactivados || [], e.currentTarget);
  });

  document.getElementById("btnPlanillasFaltantesCatalog")?.addEventListener("click", () => {
    document.getElementById("planillasFaltantesCatalogInput")?.click();
  });

  document.getElementById("planillasFaltantesCatalogInput")?.addEventListener("change", (e) => {
    const file = e.target.files?.[0];
    onPlanillasFaltantesListUpload(file, getValidated, getMainFilters);
    e.target.value = "";
  });

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    const planillasModal = document.getElementById("modalPlanillasFaltantes");
    if (planillasModal && !planillasModal.hidden) {
      closePlanillasFaltantesModal();
      return;
    }
    closeResumenModal();
  });
}
