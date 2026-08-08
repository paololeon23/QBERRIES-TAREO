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

function countUniqueActividad(rows, matchFn) {
  return new Set(
    rows
      .filter((r) => matchFn(normActividadKpi(r.actividad)))
      .map((r) => normNombreKpi(r.trabajador))
      .filter(Boolean)
  ).size;
}

/** Nombres únicos de la columna Supervisor dentro de COSTO DE COSECHA (1 por persona). */
export function countSupervisoresCosto(rows) {
  const costoRows = (rows || []).filter((r) => r.esCostoCosecha);
  return new Set(
    costoRows.map((r) => normNombreKpi(r.supervisor)).filter(Boolean)
  ).size;
}

export function countScanerCosto(rows) {
  const costoRows = (rows || []).filter((r) => r.esCostoCosecha);
  return countUniqueActividad(
    costoRows,
    (a) => a === "scaner" || a === "scanner" || a === "escaner"
  );
}

export function countCosechaCosto(rows) {
  const costoRows = (rows || []).filter((r) => r.esCostoCosecha);
  return countUniqueActividad(costoRows, (a) => a === "cosecha");
}

export { normActividadKpi, normNombreKpi, countUniqueActividad };
