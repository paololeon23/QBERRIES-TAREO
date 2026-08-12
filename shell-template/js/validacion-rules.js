/** Motor de reglas: valida suma del día (turnos) solo en COSTO DE COSECHA. */

import { HOUR_BASE, HOURS_LABEL, classifyDayHours } from "./excel-parser.js";
import { countSupervisoresCosto } from "./validacion-kpi.js";

/** Horarios flexibles: no marcar rojo por hora de inicio/fin mientras la suma ≤ 12 h. */
const REST_GAP_MIN = 60; // tip informativo de descanso (aviso), no error rojo

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function rowTextBlob(row) {
  return [
    row.documento,
    row.trabajador,
    row.supervisor,
    row.fundo,
    row.macroPartida,
    row.variedad,
    row.tipo,
    row.estado,
    row.sessionTipo,
    row.sessionVariedad,
    row.horasTurno,
    row.totalDia,
    ...Object.values(row.hoursByDay || {})
  ]
    .map((v) => String(v ?? ""))
    .join(" | ");
}

function detectTipoBucket(row) {
  const blob = `${row.tipo} ${row.sessionTipo} ${row.macroPartida} ${row.fundo} ${row.variedad} ${row.sessionVariedad}`.toLowerCase();
  if (blob.includes("china")) return "china";
  if (blob.includes("convencional") || blob.includes("conv")) return "convencional";
  return row.sessionTipo || "";
}

function cleanCeco(value) {
  return String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function dayGroupKey(row) {
  if (row.fecha) return `${row.documento}|${row.fecha}|${row.macroPartida || ""}`;
  if (row.fechaSerial != null) {
    return `${row.documento}|serial-${row.fechaSerial}|${row.macroPartida || ""}`;
  }
  const clock = row.horaInicioTexto || row.horaInicioKey || `row-${row.excelRow || row.rowIndex}`;
  return `${row.documento}|sin-fecha|${clock}|${row.macroPartida || ""}`;
}

export function validateDataset(parsed) {
  const rows = (parsed.rows || []).map((row) => ({ ...row, flags: [], dayFlags: {} }));

  const findings = {
    overHours: [],
    overBase: [],
    posibleSalida: [],
    nonCosecha: [],
    duplicates: [],
    naFails: [],
    cesados: [],
    minoritaria: [],
    horario: [],
    cecoVacio: [],
    documentoVacio: [],
    trabajadorVacio: [],
    supervisorAlerts: []
  };

  // Turno duplicado: mismo DNI + fecha + misma hora de inicio (los 2 turnos del día son normales).
  const byShift = new Map();
  rows.forEach((row) => {
    if (!row.documento || !row.horaInicioKey) return;
    const shiftKey = `${row.documento}|${row.fecha || ""}|${row.horaInicioKey}`;
    if (!byShift.has(shiftKey)) byShift.set(shiftKey, []);
    byShift.get(shiftKey).push(row);
  });
  byShift.forEach((group) => {
    if (group.length <= 1) return;
    group.forEach((row, idx) => {
      if (!row.flags.includes("duplicado")) row.flags.push("duplicado");
      const excelRef = row.excelRow ? ` fila Excel ${row.excelRow}` : "";
      row.tipDuplicado = `Duplicado ${idx + 1} de ${group.length}${excelRef}: mismo DNI + fecha + hora inicio ${
        row.horaInicioTexto || ""
      }. Fin: ${row.horaFinTexto || "—"}. Revisar registro doble.`;
    });
    findings.duplicates.push({
      documento: group[0].documento,
      count: group.length,
      trabajadores: [...new Set(group.map((r) => r.trabajador).filter(Boolean))],
      fecha: group[0].fecha || "",
      horaInicio: group[0].horaInicioTexto || "",
      horasFin: group.map((r) => r.horaFinTexto).filter(Boolean),
      excelRows: group.map((r) => r.excelRow).filter(Boolean),
      macroPartida: group[0].macroPartida || "",
      supervisor: group[0].supervisor,
      rowIndex: group[0].rowIndex
    });
  });

  // Validación de horas por DÍA (suma de turnos), solo COSTO DE COSECHA.
  // < 9.6 posible pase · exacto 9.6 OK · 10.1/10.6/12 aviso · resto ≠ exacto error.
  const dayGroups = new Map();
  rows.forEach((row) => {
    if (!row.esCostoCosecha) {
      row.dayFlags[HOURS_LABEL] = row.totalDia == null && row.horasTurno == null ? "na" : "ok";
      return;
    }
    const key = dayGroupKey(row);
    if (!dayGroups.has(key)) dayGroups.set(key, []);
    dayGroups.get(key).push(row);
  });

  dayGroups.forEach((group) => {
    const total =
      group[0].totalDia != null
        ? Number(group[0].totalDia)
        : Math.round(group.reduce((s, r) => s + (Number(r.horasTurno) || 0), 0) * 1e6) / 1e6;

    const sample = group[0];
    const classified = classifyDayHours(total);

    if (classified.flag === "na") {
      group.forEach((row) => {
        row.dayFlags[HOURS_LABEL] = "na";
        if (!row.flags.includes("na")) row.flags.push("na");
        row.tipHoras = classified.tip;
      });
      findings.naFails.push({
        field: "horas:total_dia",
        documento: sample.documento,
        trabajador: sample.trabajador,
        supervisor: sample.supervisor,
        macroPartida: sample.macroPartida,
        group: sample.supervisor || sample.macroPartida || "Sin grupo",
        rowIndex: sample.rowIndex
      });
      return;
    }

    const rounded = Math.round(Number(total) * 1e6) / 1e6;

    if (classified.flag === "rojo") {
      group.forEach((row) => {
        row.dayFlags[HOURS_LABEL] = "rojo";
        if (!row.flags.includes("rojo")) row.flags.push("rojo");
        row.totalDia = rounded;
        row.sumaHorasPago = rounded;
        row.horas = rounded;
        row.hoursByDay = { [HOURS_LABEL]: rounded };
        row.tipHoras = classified.tip;
      });
      findings.overHours.push({
        documento: sample.documento,
        trabajador: sample.trabajador,
        supervisor: sample.supervisor,
        macroPartida: sample.macroPartida,
        day: HOURS_LABEL,
        hours: rounded,
        reason: rounded > 12 ? "sobre_12" : "no_exacto",
        rowIndex: sample.rowIndex
      });
      return;
    }

    group.forEach((row) => {
      row.totalDia = rounded;
      row.sumaHorasPago = rounded;
      row.horas = rounded;
      if (row.costoCosecha != null) row.costoCosecha = rounded;
      row.hoursByDay = { [HOURS_LABEL]: rounded };
      row.dayFlags[HOURS_LABEL] = classified.flag;
      row.tipHoras = classified.tip;
      if (classified.flag === "ok") return;
      if (classified.flag === "posible-salida") {
        if (!row.flags.includes("posible-salida")) row.flags.push("posible-salida");
        return;
      }
      if (!row.flags.includes("aviso")) row.flags.push("aviso");
    });

    if (classified.flag === "posible-salida") {
      findings.posibleSalida = findings.posibleSalida || [];
      findings.posibleSalida.push({
        documento: sample.documento,
        trabajador: sample.trabajador,
        supervisor: sample.supervisor,
        day: HOURS_LABEL,
        hours: rounded,
        reason: "menor_9_6",
        rowIndex: sample.rowIndex
      });
      return;
    }

    if (classified.flag !== "ok") {
      findings.overBase.push({
        documento: sample.documento,
        trabajador: sample.trabajador,
        supervisor: sample.supervisor,
        day: HOURS_LABEL,
        hours: rounded,
        extra: rounded > HOUR_BASE ? "extra-hasta-12" : "",
        rowIndex: sample.rowIndex
      });
    }
  });

  // Horarios flexibles (solo COSTO DE COSECHA):
  // Inicio/fin libres (ej. 07:00 … 19:06). Solo falta de dato o fin ≤ inicio → rojo.
  // Descanso ≠ 1 h → aviso (no asusta en rojo).
  dayGroups.forEach((group) => {
    const ordered = group
      .slice()
      .sort((a, b) => (a.horaInicioMin ?? 9999) - (b.horaInicioMin ?? 9999));
    const sample = ordered[0];
    const hardProblems = [];
    const softProblems = [];

    ordered.forEach((row) => {
      const ini = row.horaInicioMin;
      const fin = row.horaFinMin;
      row.dayFlags = row.dayFlags || {};
      row.dayFlags.horaInicio = "ok";
      row.dayFlags.horaFin = "ok";
      row.tipHoraInicio = "";
      row.tipHoraFin = "";

      if (ini == null) {
        row.dayFlags.horaInicio = "rojo";
        row.tipHoraInicio = "Error: falta hora de inicio. ¿Está vacía en el Excel?";
        hardProblems.push("Sin hora inicio");
        return;
      }

      if (fin == null) {
        row.dayFlags.horaFin = "rojo";
        row.tipHoraFin = "Error: falta hora de fin. ¿Está vacía en el Excel?";
        hardProblems.push("Sin hora fin");
        return;
      }

      if (fin <= ini) {
        row.dayFlags.horaFin = "rojo";
        row.tipHoraFin = `Error: fin ${row.horaFinTexto} debe ser después de inicio ${row.horaInicioTexto}.`;
        hardProblems.push(`Fin ${row.horaFinTexto} ≤ inicio ${row.horaInicioTexto}`);
      }
    });

    if (ordered.length >= 2) {
      for (let i = 0; i < ordered.length - 1; i += 1) {
        const a = ordered[i];
        const b = ordered[i + 1];
        if (a.horaFinMin == null || b.horaInicioMin == null) continue;
        const gap = b.horaInicioMin - a.horaFinMin;
        if (gap !== REST_GAP_MIN) {
          const msg = `Aviso: descanso entre turnos ${gap} min (habitual 60 min).`;
          if (a.dayFlags.horaFin !== "rojo") {
            a.dayFlags.horaFin = "aviso";
            a.tipHoraFin = msg;
          }
          if (b.dayFlags.horaInicio !== "rojo") {
            b.dayFlags.horaInicio = "aviso";
            b.tipHoraInicio = msg;
          }
          softProblems.push(
            `Descanso ${a.horaFinTexto || "?"} → ${b.horaInicioTexto || "?"} = ${gap} min`
          );
        }
      }
    }

    if (hardProblems.length) {
      group.forEach((row) => {
        if (!row.flags.includes("rojo")) row.flags.push("rojo");
        if (!row.flags.includes("horario")) row.flags.push("horario");
      });
      findings.horario.push({
        documento: sample.documento,
        trabajador: sample.trabajador,
        supervisor: sample.supervisor,
        fecha: sample.fecha,
        detalle: hardProblems.join("; "),
        inicios: ordered.map((r) => r.horaInicioTexto).join(" / "),
        fines: ordered.map((r) => r.horaFinTexto).join(" / "),
        rowIndex: sample.rowIndex
      });
      return;
    }

    if (softProblems.length) {
      group.forEach((row) => {
        if (!row.flags.includes("rojo") && !row.flags.includes("aviso")) row.flags.push("aviso");
        if (!row.flags.includes("horario")) row.flags.push("horario");
      });
    }
  });

  rows.forEach((row) => {
    const blob = rowTextBlob(row).toLowerCase();
    if (blob.includes("cesado") || normalizeText(row.estado).includes("cesado")) {
      row.flags.push("cesado");
      findings.cesados.push({
        documento: row.documento,
        trabajador: row.trabajador,
        supervisor: row.supervisor,
        rowIndex: row.rowIndex
      });
    }

    if (blob.includes("menoritaria")) {
      row.flags.push("menoritaria");
      findings.minoritaria.push({
        documento: row.documento,
        trabajador: row.trabajador,
        macroPartida: row.macroPartida,
        rowIndex: row.rowIndex
      });
    }

    // CECO (columna U) no debe estar vacío
    row.dayFlags = row.dayFlags || {};
    if (!cleanCeco(row.ceco)) {
      row.dayFlags.ceco = "rojo";
      row.tipCeco = "Error: CECO vacío. Debe indicar el centro de costo (columna U).";
      if (!row.flags.includes("rojo")) row.flags.push("rojo");
      if (!row.flags.includes("ceco")) row.flags.push("ceco");
      findings.cecoVacio.push({
        documento: row.documento,
        trabajador: row.trabajador,
        supervisor: row.supervisor,
        macroPartida: row.macroPartida,
        actividad: row.actividad || "",
        rowIndex: row.rowIndex
      });
    } else {
      row.dayFlags.ceco = "ok";
      row.tipCeco = "";
    }

    // Documento vacío con Código Trabajador (= Documento) → error
    // Trabajador vacío con identidad (Documento o Código) → error
    const codigoOk = String(row.codigoTrabajador || "").trim();
    const docCellOk = String(row.documentoCell || "").trim();
    const docOk = String(row.documento || "").trim(); // ya con fallback al código
    const trabOk = String(row.trabajador || "").trim();
    const docVacio = Boolean(row.documentoVacio) || (Boolean(codigoOk || docOk) && !docCellOk);

    if (row.esCostoCosecha && docVacio && (codigoOk || docOk)) {
      row.dayFlags.documento = "rojo";
      row.tipDocumento = codigoOk
        ? `Error: Documento vacío. Código Trabajador ${codigoOk} actúa como Documento.`
        : "Error: Documento vacío.";
      if (!row.flags.includes("rojo")) row.flags.push("rojo");
      if (!row.flags.includes("sin-documento")) row.flags.push("sin-documento");
      findings.documentoVacio.push({
        documento: docOk,
        codigoTrabajador: codigoOk,
        trabajador: row.trabajador || "",
        supervisor: row.supervisor,
        macroPartida: row.macroPartida,
        actividad: row.actividad || "",
        fecha: row.fecha || "",
        rowIndex: row.rowIndex
      });
    } else {
      row.dayFlags.documento = row.dayFlags.documento || "ok";
      if (!row.tipDocumento) row.tipDocumento = "";
    }

    if (row.esCostoCosecha && docOk && !trabOk) {
      row.dayFlags.trabajador = "rojo";
      row.tipTrabajador = codigoOk
        ? `Error: hay Código/Documento (${docOk}) pero Trabajador está vacío.`
        : "Error: hay Documento pero Trabajador está vacío. Completa el nombre del trabajador.";
      if (!row.flags.includes("rojo")) row.flags.push("rojo");
      if (!row.flags.includes("sin-trabajador")) row.flags.push("sin-trabajador");
      findings.trabajadorVacio.push({
        documento: row.documento,
        codigoTrabajador: codigoOk,
        trabajador: "",
        supervisor: row.supervisor,
        macroPartida: row.macroPartida,
        actividad: row.actividad || "",
        fecha: row.fecha || "",
        rowIndex: row.rowIndex
      });
    } else {
      row.dayFlags.trabajador = row.dayFlags.trabajador || "ok";
      if (!row.tipTrabajador) row.tipTrabajador = "";
    }

    row.tipoBucket = detectTipoBucket(row);
    row.status = row.flags.includes("rojo")
      ? "rojo"
      : row.flags.includes("aviso")
        ? "aviso"
        : row.flags.includes("posible-salida")
          ? "posible-salida"
          : "ok";
  });

  const bySupervisor = new Map();
  findings.overHours.forEach((item) => {
    const key = item.supervisor || "(sin supervisor)";
    if (!bySupervisor.has(key)) bySupervisor.set(key, []);
    bySupervisor.get(key).push(item);
  });
  findings.supervisorAlerts = [...bySupervisor.entries()].map(([supervisor, items]) => ({
    supervisor,
    count: items.length,
    items
  }));

  const resumenMap = new Map();
  rows.forEach((row) => {
    const key = `${row.documento}||${row.supervisor}`;
    if (!resumenMap.has(key)) {
      resumenMap.set(key, {
        documento: row.documento,
        trabajador: row.trabajador,
        supervisor: row.supervisor,
        fundo: row.fundo,
        macroPartida: row.macroPartida,
        tipoBucket: row.tipoBucket,
        variedad: row.sessionVariedad || row.variedad,
        totalHoras: 0,
        dias: 0,
        alertasRojo: 0,
        alertasAviso: 0
      });
    }
    const agg = resumenMap.get(key);
    if (row.horasTurno != null) {
      agg.totalHoras += Number(row.horasTurno) || 0;
    }
    // Contar alerta una vez por día-persona
  });

  // Alertas / días únicos por persona-día
  const seenDay = new Set();
  rows.forEach((row) => {
    const key = `${row.documento}||${row.supervisor}`;
    const agg = resumenMap.get(key);
    if (!agg) return;
    const dayKey = `${key}|${row.fecha}|${row.macroPartida}`;
    if (!seenDay.has(dayKey)) {
      seenDay.add(dayKey);
      agg.dias += 1;
      if (row.status === "rojo") agg.alertasRojo += 1;
      if (row.status === "aviso") agg.alertasAviso += 1;
    }
  });

  const naByGroup = new Map();
  findings.naFails.forEach((item) => {
    const key = item.group || "Sin grupo";
    naByGroup.set(key, (naByGroup.get(key) || 0) + 1);
  });

  // KPIs de horas por persona-día (no por fila/turno)
  const dayStatus = new Map();
  rows
    .filter((r) => r.esCostoCosecha)
    .forEach((row) => {
      const key = dayGroupKey(row);
      if (!dayStatus.has(key)) dayStatus.set(key, row.status);
    });

  return {
    rows,
    findings,
    resumen: [...resumenMap.values()],
    naByGroup: [...naByGroup.entries()].map(([group, count]) => ({ group, count })),
    kpis: {
      total: rows.length,
      costoCosecha: rows.filter((r) => r.esCostoCosecha).length,
      supervisores: countSupervisoresCosto(rows),
      cosechadores: new Set(
        rows.filter((r) => r.esCostoCosecha).map((r) => String(r.documento || "").trim()).filter(Boolean)
      ).size,
      rojo: [...dayStatus.values()].filter((s) => s === "rojo").length,
      aviso: [...dayStatus.values()].filter((s) => s === "aviso").length,
      posibleSalida: [...dayStatus.values()].filter((s) => s === "posible-salida").length,
      duplicados: new Set(findings.duplicates.map((d) => d.documento)).size,
      nonCosecha: 0,
      cesados: findings.cesados.length,
      minoritaria: findings.minoritaria.length
    }
  };
}

export function uniqueValues(rows, field) {
  return [...new Set(rows.map((r) => String(r[field] ?? "").trim()).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, "es")
  );
}
