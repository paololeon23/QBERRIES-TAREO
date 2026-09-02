/** KPIs compartidos (pantalla principal y Resumen). */

function normActividadKpi(act) {
  return String(act || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function normNombreKpi(name) {
  return String(name || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function costoRows(rows) {
  return (rows || []).filter((r) => r.esCostoCosecha);
}

function grupoSupervisorKey(row) {
  return normNombreKpi(row.supervisor);
}

const ACTIVIDADES_POR_GRUPO = new Set([
  "cosecha",
  "scaner",
  "scanner",
  "escaner",
  "supervisor de cosecha"
]);

/** Área de Calidad — no forma parte del equipo cosecha (scaner + supervisor + cosecha). */
export function isActividadCalidad(actividad) {
  const a = normActividadKpi(actividad);
  return (
    a === "auxiliar de calidad" ||
    a === "aux calidad" ||
    a.includes("auxiliar") && a.includes("calidad") ||
    a === "area de calidad"
  );
}

export function isActividadEquipoCosecha(actividad) {
  if (isActividadCalidad(actividad)) return false;
  return ACTIVIDADES_POR_GRUPO.has(normActividadKpi(actividad));
}

function costoRowsEquipos(rows) {
  return costoRows(rows).filter((r) => !isActividadCalidad(r.actividad));
}

/** Equipos únicos en COSTO DE COSECHA (sin filas de Área de Calidad). */
function gruposCostoSet(rows) {
  return new Set(costoRowsEquipos(rows).map(grupoSupervisorKey).filter(Boolean));
}

/** Grupos con al menos una fila que cumple matchFn(actividad normalizada). */
function countGruposConActividad(rows, matchFn) {
  const grupos = new Set();
  costoRowsEquipos(rows).forEach((r) => {
    const sup = grupoSupervisorKey(r);
    if (!sup) return;
    if (matchFn(normActividadKpi(r.actividad))) grupos.add(sup);
  });
  return grupos.size;
}

/** Personas únicas (actividad Calidad u otras fuera de equipo). */
export function countPersonasActividad(rows, matchFn) {
  return new Set(
    (rows || [])
      .filter((r) => r.esCostoCosecha && matchFn(normActividadKpi(r.actividad)))
      .map((r) => String(r.documento || normNombreKpi(r.trabajador)).trim())
      .filter(Boolean)
  ).size;
}

export function countAuxiliarCalidad(rows) {
  return countPersonasActividad(rows, (a) => isActividadCalidad(a));
}

export function countSupervisoresCalidad(rows) {
  return new Set(
    costoRows(rows)
      .filter((r) => isActividadCalidad(r.actividad))
      .map(grupoSupervisorKey)
      .filter(Boolean)
  ).size;
}

/** Equipos (supervisores) en COSTO DE COSECHA — 1 por grupo. */
export function countSupervisoresCosto(rows) {
  return gruposCostoSet(rows).size;
}

/** Grupos con scaner (1 por equipo, no personas sueltas). */
export function countScanerCosto(rows) {
  return countGruposConActividad(
    rows,
    (a) => a === "scaner" || a === "scanner" || a === "escaner"
  );
}

/** Grupos con cosecha (1 integrante por equipo, no total de cosechadores). */
export function countCosechaCosto(rows) {
  return countGruposConActividad(rows, (a) => a === "cosecha");
}

/** Grupos con actividad SUPERVISOR DE COSECHA. */
export function countSupervisorCosechaActividadCosto(rows) {
  return countGruposConActividad(rows, (a) => a === "supervisor de cosecha");
}

export function isActividadContadaPorGrupo(actividad) {
  if (isActividadCalidad(actividad)) return false;
  return ACTIVIDADES_POR_GRUPO.has(normActividadKpi(actividad));
}

/** No cuentan en KPI/tabla Resumen aunque aparezcan en tareo. */
const SUPERVISORES_EXCLUIDOS_RESUMEN = new Set(
  [
    "LEON VARGAS DEYSI TATIANA",
    "VASQUEZ COTRINA EVELYN RUVIT",
    "MARTINEZ REYES WILSON ALFREDO",
    "PASTOR CUEVA SORAYDA ARACELY",
    "NAMOC NARRO BIVIANA DE LOS ANGELES",
    "PADILLA NUÑEZ JESUS MARIA"
  ].map(normNombreKpi)
);

export function isSupervisorExcluidoResumen(supervisor) {
  return SUPERVISORES_EXCLUIDOS_RESUMEN.has(normNombreKpi(supervisor));
}

export function filterGruposResumen(groups) {
  return (groups || []).filter((g) => !isSupervisorExcluidoResumen(g.supervisor));
}

export function filterRowsResumenExcluidos(rows) {
  return (rows || []).filter((r) => !isSupervisorExcluidoResumen(r.supervisor));
}

export { normActividadKpi, normNombreKpi };
