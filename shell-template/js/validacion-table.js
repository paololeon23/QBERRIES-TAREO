/** Tabla: una fila por persona/día con Suma de Horas Pago + reloj turnos. */

import { matchExactHourStep, classifyDayHours, HOURS_LABEL, HOUR_BASE } from "./excel-parser.js?v=20260812b";

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Hasta 3 decimales (5.5 → 5.5, 7.333333 → 7.333). */
function formatHour(value) {
  if (value == null || value === "") return "";
  const n = Number(value);
  if (!Number.isFinite(n)) return String(value);
  return String(Math.round(n * 1e3) / 1e3);
}

function fillSelect(select, values, blankLabel) {
  if (!select) return;
  const current = select.value;
  select.innerHTML = `<option value="">${blankLabel}</option>${values
    .map((v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`)
    .join("")}`;
  if (values.includes(current)) select.value = current;
}

function tipAttr(text) {
  if (!text) return "";
  return ` data-tip="${escapeHtml(text)}"`;
}

function stackTimesHtml(times, flag, tip) {
  const list = Array.isArray(times) ? times.filter(Boolean) : [];
  if (!list.length) return `<td class="cell-clock"></td>`;
  const danger = flag === "rojo";
  const warn = flag === "aviso";
  const tipClass = danger || warn ? " has-tip has-tip--cell" : "";
  const cls = danger
    ? `cell-clock is-cell-danger${tipClass}`
    : warn
      ? `cell-clock is-cell-warn${tipClass}`
      : "cell-clock";
  return `<td class="${cls}"${danger || warn ? tipAttr(tip) : ""}>${list
    .map((t) => `<span>${escapeHtml(t)}</span>`)
    .join("")}</td>`;
}

/** Clave de día: nunca mezclar dos fechas distintas. */
function dayViewKey(row) {
  if (row.fecha) return `${row.documento}|${row.fecha}|${row.macroPartida || ""}`;
  if (row.fechaSerial != null) {
    return `${row.documento}|serial-${row.fechaSerial}|${row.macroPartida || ""}`;
  }
  const clock = row.horaInicioTexto || row.horaInicioKey || `row-${row.excelRow || row.rowIndex}`;
  return `${row.documento}|sin-fecha|${clock}|${row.macroPartida || ""}`;
}

/** Agrupa turnos del mismo DNI + fecha + macro en una sola fila de vista. */
export function collapseToDayRows(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const key = dayViewKey(row);
    if (!map.has(key)) {
      map.set(key, {
        ...row,
        _turnHours: [],
        _turnStarts: [],
        _turnEnds: [],
        _iniFlag: "ok",
        _finFlag: "ok"
      });
    }
    const agg = map.get(key);
    if (row.horasTurno != null && Number.isFinite(Number(row.horasTurno))) {
      agg._turnHours.push({
        hours: Number(row.horasTurno),
        ini: row.horaInicioMin ?? 9999,
        iniTxt: row.horaInicioTexto || "",
        finTxt: row.horaFinTexto || "",
        iniKey: row.horaInicioKey || row.horaInicioTexto || `row-${row.excelRow || row.rowIndex}`,
        iniFlag: row.dayFlags?.horaInicio || "ok",
        finFlag: row.dayFlags?.horaFin || "ok"
      });
    }
    if (row.status === "rojo") agg.status = "rojo";
    else if (row.status === "aviso" && agg.status !== "rojo") agg.status = "aviso";
    else if (
      row.status === "posible-salida" &&
      agg.status !== "rojo" &&
      agg.status !== "aviso"
    ) {
      agg.status = "posible-salida";
    }
    if (row.dayFlags) {
      agg.dayFlags = { ...agg.dayFlags, ...row.dayFlags };
    }
    if (row.tipHoras) agg.tipHoras = row.tipHoras;
    if (row.dayFlags?.horaInicio === "rojo" && row.tipHoraInicio) {
      agg.tipHoraInicio = row.tipHoraInicio;
    }
    if (row.dayFlags?.horaFin === "rojo" && row.tipHoraFin) {
      agg.tipHoraFin = row.tipHoraFin;
    }
    if (row.flags?.length) {
      agg.flags = [...new Set([...(agg.flags || []), ...row.flags])];
    }
    if (row.tipDuplicado) agg.tipDuplicado = row.tipDuplicado;
    if (row.actividad) {
      if (!agg.actividad) agg.actividad = row.actividad;
      else if (agg.actividad !== row.actividad && !String(agg.actividad).includes(row.actividad)) {
        agg.actividad = `${agg.actividad} / ${row.actividad}`;
      }
    }
    if (row.ceco) {
      if (!agg.ceco) agg.ceco = row.ceco;
      else if (agg.ceco !== row.ceco && !String(agg.ceco).includes(row.ceco)) {
        agg.ceco = `${agg.ceco} / ${row.ceco}`;
      }
    }
    if (row.dayFlags?.ceco === "rojo") {
      agg.dayFlags = { ...(agg.dayFlags || {}), ceco: "rojo" };
      if (row.tipCeco) agg.tipCeco = row.tipCeco;
    }
    if (row.dayFlags?.documento === "rojo" || row.documentoVacio) {
      agg.dayFlags = { ...(agg.dayFlags || {}), documento: "rojo" };
      if (row.tipDocumento) agg.tipDocumento = row.tipDocumento;
      if (row.codigoTrabajador) agg.codigoTrabajador = row.codigoTrabajador;
      agg.documentoVacio = true;
    }
    if (row.dayFlags?.trabajador === "rojo") {
      agg.dayFlags = { ...(agg.dayFlags || {}), trabajador: "rojo" };
      if (row.tipTrabajador) agg.tipTrabajador = row.tipTrabajador;
      if (!String(agg.trabajador || "").trim()) agg.trabajador = "";
    }
    if (row.codigoTrabajador && !agg.codigoTrabajador) {
      agg.codigoTrabajador = row.codigoTrabajador;
    }
  });

  return [...map.values()].map((agg) => {
    // Todos los turnos (incluye inicio repetido si hay registro doble)
    const allTurns = (agg._turnHours || []).slice().sort((a, b) => a.ini - b.ini);
    // Para sumar: un turno por hora de inicio (mismo día)
    const seen = new Set();
    const turns = allTurns.filter((t) => {
      const k = String(t.iniKey || t.iniTxt || t.hours);
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });

    const turnos = turns.map((t) => t.hours);
    // Suma de turnos del mismo día (5.5 + 4.6 = 10.1)
    let sum = Math.round(turnos.reduce((s, n) => s + n, 0) * 1e10) / 1e10;
    let detalle = turnos.map((n) => formatHour(n)).join(" + ");

    // Si el Excel trae el TOTAL del día (9.6/10.1/10.6) repetido en cada turno, no sumar dos veces
    if (
      turnos.length >= 2 &&
      turnos.every((h) => matchExactHourStep(h) != null) &&
      turnos.every((h) => Math.abs(h - turnos[0]) <= 0.02)
    ) {
      sum = turnos[0];
      detalle = formatHour(sum);
    }

    // Mantener dayFlags de horas alineado a la suma recalculada
    let sumaFlag = agg.dayFlags?.[HOURS_LABEL] || "ok";
    if (agg.esCostoCosecha) {
      const classified = classifyDayHours(sum);
      sumaFlag = classified.flag === "na" ? "na" : classified.flag;
      if (classified.tip) agg.tipHoras = classified.tip;
      if (sumaFlag === "rojo") {
        agg.status = "rojo";
        if (!Array.isArray(agg.flags)) agg.flags = [];
        if (!agg.flags.includes("rojo")) agg.flags.push("rojo");
      } else if (
        (sumaFlag === "aviso" || sumaFlag === "aviso-hora") &&
        agg.status !== "rojo"
      ) {
        agg.status = "aviso";
        if (!Array.isArray(agg.flags)) agg.flags = [];
        if (!agg.flags.includes("aviso")) agg.flags.push("aviso");
      } else if (
        sumaFlag === "posible-salida" &&
        agg.status !== "rojo" &&
        agg.status !== "aviso"
      ) {
        agg.status = "posible-salida";
        if (!Array.isArray(agg.flags)) agg.flags = [];
        if (!agg.flags.includes("posible-salida")) agg.flags.push("posible-salida");
      }
    }
    const inicios = allTurns.map((t) => t.iniTxt).filter(Boolean);
    const fines = allTurns.map((t) => t.finTxt).filter(Boolean);
    const iniFlag = turns.some((t) => t.iniFlag === "rojo")
      ? "rojo"
      : turns.some((t) => t.iniFlag === "aviso")
        ? "aviso"
        : "ok";
    const finFlag = turns.some((t) => t.finFlag === "rojo")
      ? "rojo"
      : turns.some((t) => t.finFlag === "aviso")
        ? "aviso"
        : "ok";
    const tipIni =
      agg.tipHoraInicio ||
      (iniFlag === "rojo"
        ? "Error: falta hora de inicio."
        : iniFlag === "aviso"
          ? "Aviso: revisar horario de inicio / descanso."
          : "");
    const tipFin =
      agg.tipHoraFin ||
      (finFlag === "rojo"
        ? "Error: falta hora de fin o fin ≤ inicio."
        : finFlag === "aviso"
          ? "Aviso: revisar horario de fin / descanso."
          : "");

    delete agg._turnHours;
    delete agg._turnStarts;
    delete agg._turnEnds;
    delete agg._iniFlag;
    delete agg._finFlag;

    return {
      ...agg,
      sumaHorasPago: sum,
      totalDia: sum,
      horas: sum,
      turnosDetalle: detalle,
      turnosCount: turnos.length || 1,
      horasInicioDetalle: inicios,
      horasFinDetalle: fines,
      tipHoraInicio: tipIni,
      tipHoraFin: tipFin,
      dayFlags: {
        ...(agg.dayFlags || {}),
        [HOURS_LABEL]: sumaFlag,
        horaInicio: iniFlag,
        horaFin: finFlag
      }
    };
  });
}

export function populateFilters(state, options = {}) {
  const errorOnly = Boolean(options.errorOnly);
  const warnOnly = Boolean(options.warnOnly);
  const paseOnly = Boolean(options.paseOnly);
  const dupOnly = Boolean(options.dupOnly);
  const allRows = state?.validated?.rows || [];
  const rows = errorOnly
    ? allRows.filter((r) => r.status === "rojo")
    : warnOnly
      ? allRows.filter((r) => r.status === "aviso")
      : paseOnly
        ? allRows.filter((r) => r.status === "posible-salida")
        : dupOnly
          ? allRows.filter((r) => (r.flags || []).includes("duplicado"))
          : allRows;

  fillSelect(
    document.getElementById("fltSupervisor"),
    [...new Set(rows.map((r) => r.supervisor).filter(Boolean))].sort((a, b) => a.localeCompare(b, "es")),
    "Todos"
  );
  fillSelect(
    document.getElementById("fltFundo"),
    [...new Set(rows.map((r) => r.fundo).filter(Boolean))].sort((a, b) => a.localeCompare(b, "es")),
    "Todos"
  );
  fillSelect(
    document.getElementById("fltFecha"),
    [...new Set(rows.map((r) => r.fecha).filter(Boolean))].sort((a, b) => {
      const pa = a.split("/").reverse().join("");
      const pb = b.split("/").reverse().join("");
      return pa.localeCompare(pb);
    }),
    "Todas"
  );

  const estado = document.getElementById("fltEstado");
  if (estado && errorOnly) {
    estado.innerHTML = `<option value="rojo">Error &gt; 12</option>`;
    estado.value = "rojo";
    estado.disabled = true;
  } else if (estado && warnOnly) {
    estado.innerHTML = `<option value="aviso">Advertencia ≤ 12</option>`;
    estado.value = "aviso";
    estado.disabled = true;
  } else if (estado && paseOnly) {
    estado.innerHTML = `<option value="posible-salida">Posible pase &lt; 9.6</option>`;
    estado.value = "posible-salida";
    estado.disabled = true;
  } else if (estado) {
    const prev = estado.value;
    estado.disabled = false;
    estado.innerHTML = `
      <option value="">Todos</option>
      <option value="ok">OK = 9.6 / 11.6</option>
      <option value="posible-salida">Posible pase &lt; 9.6</option>
      <option value="aviso">Advertencia ≤ 12</option>
      <option value="rojo">Error &gt; 12</option>
    `;
    if (["", "ok", "posible-salida", "aviso", "rojo"].includes(prev)) estado.value = prev;
    else estado.value = "";
  }
}

export function readFilters() {
  return {
    supervisor: document.getElementById("fltSupervisor")?.value || "",
    fundo: document.getElementById("fltFundo")?.value || "",
    fecha: document.getElementById("fltFecha")?.value || "",
    macro: "",
    actividad: "",
    dia: "",
    estado: document.getElementById("fltEstado")?.value || "",
    tipo: document.getElementById("fltTipoLote")?.value || "",
    search: (document.getElementById("fltSearch")?.value || "").trim().toLowerCase()
  };
}

export function countTotalErrors(rows) {
  // Total persona-día en error (lo que ve la tabla al agrupar)
  return collapseToDayRows(rows || []).filter((r) => r.status === "rojo").length;
}

export function countTotalWarnings(rows) {
  return collapseToDayRows(rows || []).filter((r) => r.status === "aviso").length;
}

export function countTotalPosibleSalidas(rows) {
  return collapseToDayRows(rows || []).filter((r) => r.status === "posible-salida").length;
}

export function countTotalDuplicados(rows) {
  // Cada fila del Excel marcada como duplicada (para ver los 2+ registros)
  return (rows || []).filter((r) => (r.flags || []).includes("duplicado")).length;
}

/** Vista sin agrupar: una fila por registro Excel (para revisar duplicados). */
export function expandDuplicateRows(rows) {
  return (rows || [])
    .filter((r) => (r.flags || []).includes("duplicado"))
    .slice()
    .sort((a, b) => {
      const da = String(a.documento || "").localeCompare(String(b.documento || ""), "es");
      if (da) return da;
      const fa = String(a.fecha || "").localeCompare(String(b.fecha || ""), "es");
      if (fa) return fa;
      const ia = (a.horaInicioMin ?? 0) - (b.horaInicioMin ?? 0);
      if (ia) return ia;
      return (a.excelRow || a.rowIndex || 0) - (b.excelRow || b.rowIndex || 0);
    })
    .map((row) => {
      const horas = Number(row.horasTurno);
      const hoursOk = Number.isFinite(horas) ? horas : null;
      return {
        ...row,
        horasInicioDetalle: row.horaInicioTexto ? [row.horaInicioTexto] : [],
        horasFinDetalle: row.horaFinTexto ? [row.horaFinTexto] : [],
        turnosDetalle: hoursOk != null ? formatHour(hoursOk) : "",
        sumaHorasPago: hoursOk,
        totalDia: hoursOk,
        horas: hoursOk,
        tipDuplicado:
          row.tipDuplicado ||
          `Registro duplicado${row.excelRow ? ` (fila Excel ${row.excelRow})` : ""}: mismo DNI + fecha + hora de inicio.`
      };
    });
}

export function clearAllFilterControls() {
  ["fltSupervisor", "fltFundo", "fltFecha", "fltTipoLote"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  const estado = document.getElementById("fltEstado");
  if (estado && !estado.disabled) estado.value = "";
  const search = document.getElementById("fltSearch");
  if (search) search.value = "";
}

export function filterRows(rows, filters) {
  return rows.filter((row) => {
    if (filters.supervisor && row.supervisor !== filters.supervisor) return false;
    if (filters.fundo && row.fundo !== filters.fundo) return false;
    if (filters.fecha && row.fecha !== filters.fecha) return false;
    if (filters.macro && row.macroPartida !== filters.macro) return false;
    if (filters.estado && row.status !== filters.estado) return false;
    if (filters.soloDuplicados && !(row.flags || []).includes("duplicado")) return false;
    if (filters.tipo) {
      const bucket = (row.tipoBucket || row.sessionTipo || "").toLowerCase();
      if (bucket !== filters.tipo) return false;
    }
    if (filters.search) {
      const blob = `${row.documento} ${row.trabajador}`.toLowerCase();
      if (!blob.includes(filters.search)) return false;
    }
    return true;
  });
}

/** Estado de paginación de la tabla (vista). */
const pagerState = {
  page: 1,
  pageSize: 20,
  total: 0
};

export function getPagerState() {
  return { ...pagerState };
}

export function setPagerPageSize(size) {
  const n = Number(size);
  pagerState.pageSize = Number.isFinite(n) && n > 0 ? n : 20;
  pagerState.page = 1;
}

export function setPagerPage(page) {
  const totalPages = Math.max(1, Math.ceil(pagerState.total / pagerState.pageSize) || 1);
  const p = Number(page);
  pagerState.page = Math.min(Math.max(1, Number.isFinite(p) ? p : 1), totalPages);
}

export function resetPager() {
  pagerState.page = 1;
}

function renderPagerControls() {
  const rangeEl = document.getElementById("pagerRange");
  const first = document.getElementById("pagerFirst");
  const prev = document.getElementById("pagerPrev");
  const next = document.getElementById("pagerNext");
  const last = document.getElementById("pagerLast");
  const sizeSel = document.getElementById("pagerPageSize");

  const total = pagerState.total;
  const size = pagerState.pageSize;
  const totalPages = Math.max(1, Math.ceil(total / size) || 1);
  if (pagerState.page > totalPages) pagerState.page = totalPages;

  const page = pagerState.page;
  const from = total === 0 ? 0 : (page - 1) * size + 1;
  const to = Math.min(page * size, total);

  if (rangeEl) rangeEl.textContent = `${from} – ${to} of ${total}`;
  if (sizeSel && String(sizeSel.value) !== String(size)) sizeSel.value = String(size);

  const atStart = page <= 1 || total === 0;
  const atEnd = page >= totalPages || total === 0;
  [first, prev].forEach((btn) => {
    if (!btn) return;
    btn.disabled = atStart;
  });
  [next, last].forEach((btn) => {
    if (!btn) return;
    btn.disabled = atEnd;
  });
}

export function bindTablePager(onChange) {
  const notify = () => {
    if (typeof onChange === "function") onChange();
  };

  document.getElementById("pagerPageSize")?.addEventListener("change", (e) => {
    setPagerPageSize(e.target.value);
    notify();
  });
  document.getElementById("pagerFirst")?.addEventListener("click", () => {
    setPagerPage(1);
    notify();
  });
  document.getElementById("pagerPrev")?.addEventListener("click", () => {
    setPagerPage(pagerState.page - 1);
    notify();
  });
  document.getElementById("pagerNext")?.addEventListener("click", () => {
    setPagerPage(pagerState.page + 1);
    notify();
  });
  document.getElementById("pagerLast")?.addEventListener("click", () => {
    const totalPages = Math.max(1, Math.ceil(pagerState.total / pagerState.pageSize) || 1);
    setPagerPage(totalPages);
    notify();
  });
}

export function renderTable(state, filteredRows, options = {}) {
  const head = document.getElementById("tblValidacionHead");
  const body = document.getElementById("tblValidacionBody");
  if (!head || !body) return;

  const hourKey = state?.parsed?.dayLabels?.[0] || "Suma de Horas Pago";
  const expandDups = Boolean(options.expandDuplicates);
  const dayRows = expandDups
    ? expandDuplicateRows(filteredRows)
    : collapseToDayRows(filteredRows);

  pagerState.total = dayRows.length;
  const totalPages = Math.max(1, Math.ceil(pagerState.total / pagerState.pageSize) || 1);
  if (pagerState.page > totalPages) pagerState.page = totalPages;

  const start = (pagerState.page - 1) * pagerState.pageSize;
  const pageRows = dayRows.slice(start, start + pagerState.pageSize);

  head.innerHTML = `
    <tr>
      <th>Documento</th>
      <th>Trabajador</th>
      <th>Supervisor</th>
      <th>Fundo</th>
      <th>Macro Partida</th>
      <th>Actividad</th>
      <th>CECO</th>
      <th>Fecha</th>
      <th>H. Inicio</th>
      <th>H. Fin</th>
      <th>H. Pago</th>
      <th>Suma h.</th>
    </tr>
  `;

  body.innerHTML = pageRows
    .map((row, i) => {
      const band = (start + i) % 2 === 1 ? "is-band" : "";
      const isDup = (row.flags || []).includes("duplicado");
      const rowClass =
        row.status === "rojo"
          ? `is-row-danger ${band}`
          : row.status === "aviso"
            ? `is-row-warn ${band}`
            : row.status === "posible-salida"
              ? `is-row-pase ${band}`
              : isDup
                ? `is-row-dup ${band}`
                : band;

      const total = row.sumaHorasPago ?? row.hoursByDay?.[hourKey] ?? row.totalDia ?? row.horas;
      const flag = row.dayFlags?.[hourKey];
      const cellClass =
        flag === "rojo"
          ? "is-cell-danger"
          : flag === "aviso" || flag === "aviso-hora"
            ? "is-cell-warn"
            : flag === "posible-salida"
              ? "is-cell-pase"
              : "";

      const detalle = row.turnosDetalle
        ? escapeHtml(row.turnosDetalle)
        : escapeHtml(formatHour(row.horasTurno));

      const tipSuma =
        row.tipHoras ||
        (flag === "rojo"
          ? "Error: suma ≠ exacto (solo 9.6 / 10.1 / 10.6 / 11.6 / 12) o supera 12 h."
          : flag === "aviso" || flag === "aviso-hora"
            ? "Aviso: solo valores exactos 10.1 / 10.6 / 12."
            : flag === "posible-salida"
              ? "Posible pase de salida: suma menor a 9.6 h. Verificar si hay pase registrado."
              : "");

      const tipDup =
        row.tipDuplicado ||
        (isDup
          ? "Turno duplicado: mismo DNI + fecha + hora de inicio registrado más de una vez."
          : "");
      const docEmpty =
        Boolean(row.documentoVacio) ||
        row.dayFlags?.documento === "rojo" ||
        (Boolean(String(row.codigoTrabajador || "").trim()) &&
          !String(row.documentoCell || "").trim() &&
          row.dayFlags?.documento === "rojo");
      const tipDocumento =
        row.tipDocumento ||
        (docEmpty
          ? `Error: Documento vacío.${
              row.codigoTrabajador ? ` Código Trabajador ${row.codigoTrabajador}.` : ""
            }`
          : "");
      const docDanger = docEmpty || row.dayFlags?.documento === "rojo";
      const docClass = [
        isDup ? "is-cell-dup" : "",
        docDanger ? "is-cell-danger" : "",
        isDup || docDanger ? "has-tip has-tip--cell" : ""
      ]
        .filter(Boolean)
        .join(" ");
      const docTip = tipDocumento || tipDup;
      const docText = docDanger
        ? `(vacío)${row.codigoTrabajador ? ` · ${escapeHtml(row.codigoTrabajador)}` : ""}`
        : escapeHtml(row.documento);

      const cecoEmpty = !String(row.ceco || "").trim() || row.dayFlags?.ceco === "rojo";
      const tipCeco =
        row.tipCeco ||
        (cecoEmpty ? "Error: CECO vacío. Debe indicar el centro de costo (columna U)." : "");
      const cecoClass = cecoEmpty
        ? "is-cell-danger has-tip has-tip--cell"
        : "";
      const cecoText = cecoEmpty ? "(vacío)" : escapeHtml(row.ceco);

      const hasDocumento = Boolean(
        String(row.documento || "").trim() || String(row.codigoTrabajador || "").trim()
      );
      const trabEmpty =
        hasDocumento &&
        (!String(row.trabajador || "").trim() || row.dayFlags?.trabajador === "rojo");
      const tipTrabajador =
        row.tipTrabajador ||
        (trabEmpty
          ? "Error: hay Código/Documento pero Trabajador está vacío. Completa el nombre del trabajador."
          : "");
      const trabClass = trabEmpty ? "is-cell-danger has-tip has-tip--cell" : "";
      const trabText = trabEmpty ? "(vacío)" : escapeHtml(row.trabajador);

      const sumaClass = `cell-suma${cellClass ? ` ${cellClass} has-tip has-tip--cell` : ""}`;

      return `
        <tr class="${rowClass}" data-row-index="${row.rowIndex}">
          <td class="${docClass}"${docTip ? tipAttr(docTip) : ""}>${docText}</td>
          <td class="${trabClass}"${trabEmpty ? tipAttr(tipTrabajador) : ""}>${trabText}</td>
          <td>${escapeHtml(row.supervisor)}</td>
          <td>${escapeHtml(row.fundo)}</td>
          <td>${escapeHtml(row.macroPartida)}</td>
          <td>${escapeHtml(row.actividad || "")}</td>
          <td class="${cecoClass}"${cecoEmpty ? tipAttr(tipCeco) : ""}>${cecoText}</td>
          <td class="cell-fecha">${escapeHtml(row.fecha || "")}</td>
          ${stackTimesHtml(row.horasInicioDetalle, row.dayFlags?.horaInicio, row.tipHoraInicio)}
          ${stackTimesHtml(row.horasFinDetalle, row.dayFlags?.horaFin, row.tipHoraFin)}
          <td class="cell-turnos">${detalle}</td>
          <td class="${sumaClass}"${cellClass ? tipAttr(tipSuma) : ""}>${escapeHtml(formatHour(total))}</td>
        </tr>
      `;
    })
    .join("");

  renderPagerControls();
}
