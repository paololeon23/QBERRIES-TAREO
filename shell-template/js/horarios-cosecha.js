/** Horarios COSTO DE COSECHA por actividad (nuevo vs legado). */

import { HOUR_STEPS } from "./excel-parser.js";

const HORARIOS_LEGADO = {
  firstStartMin: 6 * 60 + 30,
  firstStartLabel: "06:30",
  firstEndMin: 12 * 60,
  firstEndLabel: "12:00",
  secondStartMin: 13 * 60,
  secondStartLabel: "13:00",
  secondEndMin: 17 * 60 + 6,
  secondEndLabel: "17:06",
  directEndMin: 16 * 60 + 6,
  directEndLabel: "16:06",
  fullDayHours: 9.6,
  restGapMin: 60
};

const HORARIOS_NUEVO = {
  firstStartMin: 6 * 60 + 45,
  firstStartLabel: "06:45",
  firstEndMin: 12 * 60,
  firstEndLabel: "12:00",
  secondStartMin: 13 * 60,
  secondStartLabel: "13:00",
  secondEndMin: 17 * 60 + 21,
  secondEndLabel: "17:21",
  directEndMin: 16 * 60 + 21,
  directEndLabel: "16:21",
  fullDayHours: 9.6,
  restGapMin: 60
};

/** Actividades con horario nuevo (06:45–12:00 / 13:00–17:21). Preseleccionadas al cargar. */
export const ACTIVIDADES_HORARIO_NUEVO = [
  "CALIDAD",
  "COSECHA",
  "ESTIBA KIA",
  "EVALUADOR DE PESOS Y CALIBRES",
  "LAVADO DE JARRAS",
  "PROYECCIÓN",
  "SCANER",
  "SUPERVISOR DE ACOPIO",
  "SUPERVISOR DE COSECHA"
];

export function normActividadHorario(act) {
  return String(act || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function aliasActividadHorario(norm) {
  if (norm === "SCANNER" || norm === "ESCANER") return "SCANER";
  if (norm === "PROYECCION") return "PROYECCIÓN";
  return norm;
}

/** Clave comparable (alias SCANER / PROYECCIÓN). */
export function actividadHorarioKey(actividad) {
  return aliasActividadHorario(normActividadHorario(actividad));
}

const NUEVO_SET = new Set(
  ACTIVIDADES_HORARIO_NUEVO.map((a) => aliasActividadHorario(normActividadHorario(a)))
);

/** True si la actividad usa 06:45 / 17:21. */
export function isActividadHorarioNuevo(actividad) {
  const key = aliasActividadHorario(normActividadHorario(actividad));
  if (!key) return false;
  if (NUEVO_SET.has(key)) return true;
  /* CALIDAD: aceptar variantes del Excel */
  if (key === "CALIDAD" || (key.includes("CALIDAD") && !key.includes("SUPERVISOR"))) {
    return NUEVO_SET.has("CALIDAD");
  }
  return false;
}

/** Etiqueta canónica del set nuevo si hace match; si no, "". */
export function matchActividadHorarioNuevoLabel(actividad) {
  const key = aliasActividadHorario(normActividadHorario(actividad));
  if (!key) return "";
  for (const label of ACTIVIDADES_HORARIO_NUEVO) {
    const lk = aliasActividadHorario(normActividadHorario(label));
    if (lk === key) return label;
    if (lk === "CALIDAD" && (key === "CALIDAD" || key.includes("CALIDAD"))) return label;
  }
  return "";
}

function buildDirectEndMins(firstStartMin) {
  return new Set(HOUR_STEPS.map((h) => firstStartMin + Math.round(Number(h) * 60)));
}

export function normFundoHorario(fundo) {
  const f = String(fundo || "").trim().toUpperCase();
  if (f === "LICAPA II") return "LICAPA II";
  if (f === "LICAPA III") return "LICAPA III";
  return "LICAPA";
}

/** Compat: fundo LICAPA (sin II/III) o actividad del set nuevo. */
export function usesHorarioLicapaNuevo(fundoOrActividad) {
  const raw = String(fundoOrActividad || "").trim();
  if (!raw) return true;
  if (/^LICAPA(\s+II|\s+III)?$/i.test(raw)) {
    return normFundoHorario(raw) === "LICAPA";
  }
  return isActividadHorarioNuevo(raw);
}

/**
 * Horario nuevo (06:45 / 17:21) solo si:
 * - actividad está en el set preseleccionado (cosecha, etc.)
 * - Y el fundo es LICAPA (no II / III)
 * LICAPA II/III y el resto de actividades → 06:30 / 17:06
 */
export function getHorariosCosecha(fundoOrOpts, actividadArg) {
  let fundo = fundoOrOpts;
  let actividad = actividadArg;
  if (fundoOrOpts && typeof fundoOrOpts === "object" && !Array.isArray(fundoOrOpts)) {
    fundo = fundoOrOpts.fundo;
    actividad = fundoOrOpts.actividad ?? actividadArg;
  }

  const fundoKey = normFundoHorario(fundo);
  const esLicapaI = fundoKey === "LICAPA";
  const actEnSetNuevo =
    actividad != null && String(actividad).trim() !== ""
      ? isActividadHorarioNuevo(actividad)
      : true; /* sin actividad: decide solo el fundo */

  const useNuevo = esLicapaI && actEnSetNuevo;
  const base = useNuevo ? HORARIOS_NUEVO : HORARIOS_LEGADO;
  return {
    ...base,
    directEndMins: buildDirectEndMins(base.firstStartMin)
  };
}

/** Texto corto para hints de UI. */
export function horarioMananaLabel(fundo, actividad) {
  const hz = getHorariosCosecha(fundo, actividad);
  return `${hz.firstStartLabel}→${hz.firstEndLabel}`;
}

export function horarioTardeLabel(fundo, actividad) {
  const hz = getHorariosCosecha(fundo, actividad);
  return `${hz.secondStartLabel}→${hz.secondEndLabel}`;
}

export function horarioResumenHint(fundo, actividad) {
  if (fundo || actividad) {
    return `${horarioMananaLabel(fundo, actividad)} / ${horarioTardeLabel(fundo, actividad)}`;
  }
  return "LICAPA + actividades cosecha: 06:45→12:00 / 13:00→17:21 · LICAPA II/III u otras: 06:30→12:00 / 13:00→17:06";
}
