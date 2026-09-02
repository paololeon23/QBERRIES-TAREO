/** Horarios COSTO DE COSECHA por fundo (LICAPA vs LICAPA II / III). */

import { HOUR_STEPS } from "./excel-parser.js?v=20260821e";

const HORARIOS_LICAPA_II_III = {
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

const HORARIOS_LICAPA = {
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

function buildDirectEndMins(firstStartMin) {
  return new Set(HOUR_STEPS.map((h) => firstStartMin + Math.round(Number(h) * 60)));
}

export function normFundoHorario(fundo) {
  const f = String(fundo || "").trim().toUpperCase();
  if (f === "LICAPA II") return "LICAPA II";
  if (f === "LICAPA III") return "LICAPA III";
  return "LICAPA";
}

export function usesHorarioLicapaNuevo(fundo) {
  const key = normFundoHorario(fundo);
  return key === "LICAPA";
}

/** Horarios aplicables al fundo (LICAPA nuevo; II/III legado). */
export function getHorariosCosecha(fundo) {
  const base = usesHorarioLicapaNuevo(fundo) ? HORARIOS_LICAPA : HORARIOS_LICAPA_II_III;
  return {
    ...base,
    directEndMins: buildDirectEndMins(base.firstStartMin)
  };
}

/** Texto corto para hints de UI. */
export function horarioMananaLabel(fundo) {
  const hz = getHorariosCosecha(fundo);
  return `${hz.firstStartLabel}→${hz.firstEndLabel}`;
}

export function horarioTardeLabel(fundo) {
  const hz = getHorariosCosecha(fundo);
  return `${hz.secondStartLabel}→${hz.secondEndLabel}`;
}

export function horarioResumenHint(fundo) {
  if (!fundo) {
    return "LICAPA: 06:45→12:00 / 13:00→17:21 · LICAPA II/III: 06:30→12:00 / 13:00→17:06";
  }
  return `${horarioMananaLabel(fundo)} / ${horarioTardeLabel(fundo)}`;
}
