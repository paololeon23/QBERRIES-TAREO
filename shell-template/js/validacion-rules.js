/** Motor de reglas: valida suma del día (turnos) solo en COSTO DE COSECHA. */

import { HOUR_BASE, HOURS_LABEL, classifyDayHours, isNombreTrabajadorVacio } from "./excel-parser.js?v=20260909b1";
import { countSupervisoresCosto } from "./validacion-kpi.js";
import { getHorariosCosecha, isActividadHorarioNuevo } from "./horarios-cosecha.js?v=20260909b1";

/**
 * Horarios COSTO DE COSECHA (por actividad — ver horarios-cosecha.js):
 * - Actividades cosecha (preseleccionadas): 06:45→12:00 / 13:00→17:21
 * - Otras: 06:30→12:00 / 13:00→17:06
 */
const REST_GAP_MIN = 60;

function sameTurn(a, b) {
  if (!a || !b) return false;
  if (a.horaInicioMin != null && b.horaInicioMin != null) return a.horaInicioMin === b.horaInicioMin;
  if (a.horaInicioKey && b.horaInicioKey) return a.horaInicioKey === b.horaInicioKey;
  return a.horaInicioTexto === b.horaInicioTexto;
}

/** Texto corto con lo que puso la persona: "06:30 → 13:00". */
function horarioPuesto(row) {
  const ini = row?.horaInicioTexto || "—";
  const fin = row?.horaFinTexto || "—";
  const h = row?.horasTurno;
  const horas =
    h != null && Number.isFinite(Number(h))
      ? ` (${String(Math.round(Number(h) * 1e3) / 1e3)} h)`
      : "";
  return `${ini} → ${fin}${horas}`;
}

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
    actividadNoPermitida: [],
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
  // < 9.6 posible pase · exacto 9.6/11.6 OK · 10.1/10.6/12 aviso · resto ≠ exacto error.
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

  // Horarios (solo COSTO DE COSECHA)
  dayGroups.forEach((group) => {
    const ordered = group
      .slice()
      .sort((a, b) => (a.horaInicioMin ?? 9999) - (b.horaInicioMin ?? 9999));

    const uniqueStarts = [];
    const seenIni = new Set();
    ordered.forEach((row) => {
      const key = row.horaInicioMin ?? row.horaInicioTexto ?? row.horaInicioKey;
      if (key == null || key === "") return;
      const k = String(key);
      if (seenIni.has(k)) return;
      seenIni.add(k);
      uniqueStarts.push(row);
    });

    const first = uniqueStarts[0];
    const second = uniqueStarts[1];
    const cutCount = uniqueStarts.length;
    const sample = first || ordered[0];
    const hz = getHorariosCosecha(sample?.fundo, sample?.actividad);
    const hardProblems = [];
    const softProblems = [];

    ordered.forEach((row) => {
      row.dayFlags = row.dayFlags || {};
      row.dayFlags.horaInicio = "ok";
      row.dayFlags.horaFin = "ok";
      row.tipHoraInicio = "";
      row.tipHoraFin = "";
    });

    const markIni = (row, msg) => {
      row.dayFlags.horaInicio = "rojo";
      row.tipHoraInicio = msg;
      hardProblems.push(msg.replace(/^Error:\s*/i, ""));
    };
    const markFin = (row, msg) => {
      row.dayFlags.horaFin = "rojo";
      row.tipHoraFin = msg;
      hardProblems.push(msg.replace(/^Error:\s*/i, ""));
    };
    const markIniAviso = (row, msg) => {
      if (row.dayFlags.horaInicio === "rojo") return;
      row.dayFlags.horaInicio = "aviso";
      row.tipHoraInicio = msg;
      softProblems.push(msg.replace(/^Aviso:\s*/i, ""));
    };
    const markFinAviso = (row, msg) => {
      if (row.dayFlags.horaFin === "rojo") return;
      row.dayFlags.horaFin = "aviso";
      row.tipHoraFin = msg;
      softProblems.push(msg.replace(/^Aviso:\s*/i, ""));
    };

    // Datos básicos en todos los turnos
    ordered.forEach((row) => {
      const ini = row.horaInicioMin;
      const fin = row.horaFinMin;
      if (ini == null) {
        markIni(row, "Falta hora de inicio.");
        return;
      }
      if (fin == null) {
        markFin(row, "Falta hora de fin.");
        return;
      }
      if (fin <= ini) {
        markFin(row, `Fin inválido. Puso: ${horarioPuesto(row)}.`);
      }
    });

    if (first && first.horaInicioMin != null && first.horaFinMin != null) {
      if (first.horaInicioMin !== hz.firstStartMin) {
        ordered.filter((r) => sameTurn(r, first)).forEach((r) => {
          markIni(r, `Debe iniciar ${hz.firstStartLabel}. Puso: ${horarioPuesto(r)}.`);
        });
      }

      if (cutCount === 1) {
        const fin = first.horaFinMin;
        if (fin === hz.firstEndMin) {
          // solo 1.er corte
        } else if (hz.directEndMins.has(fin)) {
          // directo OK
        } else {
          ordered.filter((r) => sameTurn(r, first)).forEach((r) => {
            markFin(
              r,
              `Horario mal. Puso: ${horarioPuesto(r)}. Correcto: ${hz.firstStartLabel}→${hz.firstEndLabel} o directo ${hz.firstStartLabel}→${hz.directEndLabel}.`
            );
          });
        }
      } else {
        if (first.horaFinMin !== hz.firstEndMin) {
          ordered.filter((r) => sameTurn(r, first)).forEach((r) => {
            markFin(
              r,
              `1.er corte debe ser ${hz.firstStartLabel}→${hz.firstEndLabel}. Puso: ${horarioPuesto(r)}.`
            );
          });
        }

        if (second && second.horaInicioMin != null) {
          if (second.horaInicioMin !== hz.secondStartMin) {
            ordered.filter((r) => sameTurn(r, second)).forEach((r) => {
              markIni(r, `2.º corte inicia ${hz.secondStartLabel}. Puso: ${horarioPuesto(r)}.`);
            });
          }

          if (second.horaFinMin != null && second.horaFinMin < hz.secondEndMin) {
            ordered.filter((r) => sameTurn(r, second)).forEach((r) => {
              markFinAviso(
                r,
                `2.º corte termina ${hz.secondEndLabel} o más. Puso: ${horarioPuesto(r)}.`
              );
            });
          }
        }

        for (let i = 0; i < uniqueStarts.length - 1; i += 1) {
          const a = uniqueStarts[i];
          const b = uniqueStarts[i + 1];
          if (a.horaFinMin == null || b.horaInicioMin == null) continue;
          const gap = b.horaInicioMin - a.horaFinMin;
          if (gap !== hz.restGapMin) {
            const msg = `Descanso debe ser 1 h (${hz.firstEndLabel}→${hz.secondStartLabel}). Puso: ${a.horaFinTexto || "—"} → ${b.horaInicioTexto || "—"} (${gap} min).`;
            ordered.filter((r) => sameTurn(r, a)).forEach((r) => markFinAviso(r, msg));
            ordered.filter((r) => sameTurn(r, b)).forEach((r) => markIniAviso(r, msg));
          }
        }
      }
    }

    if (hardProblems.length) {
      group.forEach((row) => {
        if (!row.flags.includes("rojo")) row.flags.push("rojo");
        if (!row.flags.includes("horario")) row.flags.push("horario");
      });
      findings.horario.push({
        documento: sample?.documento,
        trabajador: sample?.trabajador,
        supervisor: sample?.supervisor,
        fecha: sample?.fecha,
        detalle: [...new Set(hardProblems)].join("; "),
        inicios: ordered.map((r) => r.horaInicioTexto).join(" / "),
        fines: ordered.map((r) => r.horaFinTexto).join(" / "),
        rowIndex: sample?.rowIndex
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

    // COSTO DE COSECHA: actividad debe ser una de las 9 permitidas
    if (row.esCostoCosecha && !isActividadHorarioNuevo(row.actividad)) {
      const actTxt = String(row.actividad || "").trim() || "(vacía)";
      row.dayFlags.actividad = "rojo";
      row.tipActividad =
        `Error: actividad “${actTxt}” no permitida en COSTO DE COSECHA. ` +
        "Solo: CALIDAD, COSECHA, ESTIBA KIA, EVALUADOR DE PESOS Y CALIBRES, LAVADO DE JARRAS, PROYECCIÓN, SCANER, SUPERVISOR DE ACOPIO, SUPERVISOR DE COSECHA.";
      if (!row.flags.includes("rojo")) row.flags.push("rojo");
      if (!row.flags.includes("actividad-invalida")) row.flags.push("actividad-invalida");
      findings.actividadNoPermitida.push({
        documento: row.documento,
        trabajador: row.trabajador,
        supervisor: row.supervisor,
        fundo: row.fundo || "",
        macroPartida: row.macroPartida,
        actividad: row.actividad || "",
        rowIndex: row.rowIndex
      });
    } else {
      row.dayFlags.actividad = row.dayFlags.actividad || "ok";
      if (!row.tipActividad) row.tipActividad = "";
    }

    // Documento vacío con Código Trabajador (= Documento) → error
    // DNI/Código con nombre vacío o "NO VERIFICADO" → error
    const codigoOk = String(row.codigoTrabajador || "").trim();
    const docCellOk = String(row.documentoCell || "").trim();
    const docOk = String(row.documento || "").trim(); // ya con fallback al código
    const trabOk = !isNombreTrabajadorVacio(row.trabajador);
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
      const nombreRaw = String(row.trabajador || "").trim();
      row.tipTrabajador = nombreRaw
        ? `Error: DNI/Código ${docOk} sin nombre válido (“${nombreRaw}”). Debe figurar el nombre del trabajador.`
        : `Error: DNI/Código ${docOk} sin nombre. Completa la columna Trabajador.`;
      if (!row.flags.includes("rojo")) row.flags.push("rojo");
      if (!row.flags.includes("sin-trabajador")) row.flags.push("sin-trabajador");
      findings.trabajadorVacio.push({
        documento: row.documento,
        codigoTrabajador: codigoOk,
        trabajador: nombreRaw || "(vacío)",
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
