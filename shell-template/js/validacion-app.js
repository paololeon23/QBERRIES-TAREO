/** Orquestación: upload, modal, reportes, historial. */

import { parseExcelBuffer, matchExactHourStep, classifyDayHours } from "./excel-parser.js?v=20260812b";
import { validateDataset } from "./validacion-rules.js?v=20260812b";
import {
  populateFilters,
  readFilters,
  filterRows,
  renderTable,
  collapseToDayRows,
  bindTablePager,
  resetPager,
  countTotalErrors,
  countTotalWarnings,
  countTotalPosibleSalidas,
  countTotalDuplicados,
  clearAllFilterControls
} from "./validacion-table.js?v=20260812b";
import { countSupervisoresCosto, countScanerCosto, countCosechaCosto } from "./validacion-kpi.js";
import {
  openResumenModal,
  closeResumenModal,
  syncResumenFiltersFromMain,
  renderResumenView,
  bindResumenUi,
  getFilteredResumenData
} from "./validacion-resumen.js";
import { getCurrentRoute } from "./shell.js";

const HISTORY_KEY = "qb-validacion-history";

const state = {
  parsed: null,
  validated: null,
  selectedRowIndexes: [],
  fileName: "",
  errorFocusMode: false,
  warnFocusMode: false,
  paseFocusMode: false,
  dupFocusMode: false,
  savedFiltersBeforeErrorFocus: null,
  savedFiltersBeforeWarnFocus: null,
  savedFiltersBeforePaseFocus: null,
  savedFiltersBeforeDupFocus: null
};

function $(id) {
  return document.getElementById(id);
}

function showSuccessModal(text) {
  const modal = $("modalSuccess");
  const body = $("modalSuccessText");
  if (body) body.textContent = text;
  if (modal) modal.hidden = false;
}

function hideSuccessModal() {
  const modal = $("modalSuccess");
  if (modal) modal.hidden = true;
}

const KPI_HELP = {
  supervisores: {
    title: "Supervisores",
    tone: "info",
    html: `
      <p>Personas distintas en la columna <strong>Supervisor</strong>, solo en <strong>COSTO DE COSECHA</strong>.</p>
      <ul>
        <li>Cada nombre se cuenta una sola vez (aunque trabaje en varios fundos).</li>
      </ul>`
  },
  scaner: {
    title: "Scaner",
    tone: "info",
    html: `
      <p>Personas con Actividad <strong>SCANER</strong> (columna M), solo en <strong>COSTO DE COSECHA</strong>.</p>
      <ul>
        <li>También reconoce SCANNER / ESCÁNER.</li>
        <li>Cada nombre se cuenta una sola vez.</li>
      </ul>`
  },
  cosecha: {
    title: "Cosecha",
    tone: "info",
    html: `
      <p>Personas con Actividad exactamente <strong>COSECHA</strong> (columna M), solo en <strong>COSTO DE COSECHA</strong>.</p>
      <ul>
        <li>No incluye SUPERVISOR DE COSECHA ni SCANER.</li>
        <li>Cada nombre se cuenta una sola vez.</li>
      </ul>`
  },
  error: {
    title: "Error ≠ exacto",
    tone: "danger",
    html: `
      <p>Personas-día en <strong>error</strong> dentro de <strong>COSTO DE COSECHA</strong>: suma distinta de <strong>9.6 / 10.1 / 10.6 / 11.6 / 12</strong> (ej. 9.63, 11.37), suma &gt; 12 h, horario incompleto/inválido, CECO vacío, <strong>Documento vacío</strong> o <strong>Trabajador vacío</strong>.</p>
      <ul>
        <li>Pasa el mouse sobre celdas rojas para ver el detalle.</li>
      </ul>`
  },
  extra: {
    title: "Extra / advertencia ≤ 12",
    tone: "warn",
    html: `
      <p>Solo valores <strong>exactos</strong> sobre la jornada: <strong>10.1</strong> (media hora extra), <strong>10.6</strong> (1 h extra) o <strong>12</strong> (tope).</p>
      <ul>
        <li>Cualquier otro número (ej. 11.37) no es aviso: va en <strong>rojo</strong> como ≠ exacto.</li>
      </ul>`
  },
  pase: {
    title: "Posible pase de salida",
    tone: "pase",
    html: `
      <p>Suma de horas <strong>menor a 9.6</strong> en <strong>COSTO DE COSECHA</strong>.</p>
      <ul>
        <li>Puede indicar salida anticipada: conviene cruzar con <strong>Pases de salida</strong>.</li>
        <li>Usa la barra «Posibles salidas» o el filtro Estado horas.</li>
      </ul>`
  }
};

function openKpiHelp(key) {
  const info = KPI_HELP[key];
  if (!info) return;
  const modal = $("modalKpiHelp");
  const title = $("modalKpiHelpTitle");
  const body = $("modalKpiHelpBody");
  const icon = $("modalKpiHelpIcon");
  if (title) title.textContent = info.title;
  if (body) body.innerHTML = info.html;
  if (icon) {
    icon.textContent = "i";
    icon.className = `modal__help-icon modal__help-icon--${info.tone || "info"}`;
  }
  if (modal) modal.hidden = false;
}

function hideKpiHelp() {
  const modal = $("modalKpiHelp");
  if (modal) modal.hidden = true;
}

function kpiCard(key, label, value, tone = "") {
  const toneClass = tone ? ` kpi--${tone}` : "";
  return `
    <div class="kpi${toneClass}">
      <div class="kpi__head">
        <span class="kpi__label">${label}</span>
        <button type="button" class="kpi__info" data-kpi-help="${key}" aria-label="Qué significa: ${label}" title="Qué significa">
          <span aria-hidden="true">i</span>
        </button>
      </div>
      <span class="kpi__value">${value}</span>
    </div>
  `;
}

function renderKpis(rows) {
  const host = $("kpiRow");
  if (!host || !state.validated) return;
  const k = buildKpisFromRows(rows);

  host.innerHTML = [
    kpiCard("supervisores", "Supervisores", k.supervisores),
    kpiCard("scaner", "Scaner", k.scaner),
    kpiCard("cosecha", "Cosecha", k.cosecha),
    kpiCard("error", "Error ≠ exacto", k.rojo, "danger"),
    kpiCard("extra", "Advertencia ≤12", k.aviso, "warn"),
    kpiCard("pase", "Posible pase <9.6", k.posibleSalida || 0, "pase")
  ].join("");
}

function readHistory() {
  try {
    return JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]");
  } catch {
    return [];
  }
}

function writeHistory(entry) {
  const list = readHistory();
  list.unshift(entry);
  localStorage.setItem(HISTORY_KEY, JSON.stringify(list.slice(0, 30)));
}

const historyPager = {
  page: 1,
  pageSize: 10,
  bound: false
};

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderHistoryPager(total) {
  const rangeEl = $("historyPagerRange");
  const first = $("historyPagerFirst");
  const prev = $("historyPagerPrev");
  const next = $("historyPagerNext");
  const last = $("historyPagerLast");
  const sizeSel = $("historyPagerPageSize");
  const pager = $("historyPager");
  if (pager) pager.hidden = false;

  const size = historyPager.pageSize;
  const totalPages = Math.max(1, Math.ceil(total / size) || 1);
  if (historyPager.page > totalPages) historyPager.page = totalPages;
  if (historyPager.page < 1) historyPager.page = 1;

  const page = historyPager.page;
  const from = total === 0 ? 0 : (page - 1) * size + 1;
  const to = Math.min(page * size, total);
  if (rangeEl) rangeEl.textContent = `${from} – ${to} of ${total}`;
  if (sizeSel && String(sizeSel.value) !== String(size)) sizeSel.value = String(size);

  const atStart = page <= 1 || total === 0;
  const atEnd = page >= totalPages || total === 0;
  [first, prev].forEach((btn) => {
    if (btn) btn.disabled = atStart;
  });
  [next, last].forEach((btn) => {
    if (btn) btn.disabled = atEnd;
  });
}

function bindHistoryPager() {
  if (historyPager.bound) return;
  historyPager.bound = true;

  $("historyPagerPageSize")?.addEventListener("change", (e) => {
    const n = Number(e.target.value);
    historyPager.pageSize = Number.isFinite(n) && n > 0 ? n : 10;
    historyPager.page = 1;
    renderHistory();
  });
  $("historyPagerFirst")?.addEventListener("click", () => {
    historyPager.page = 1;
    renderHistory();
  });
  $("historyPagerPrev")?.addEventListener("click", () => {
    historyPager.page = Math.max(1, historyPager.page - 1);
    renderHistory();
  });
  $("historyPagerNext")?.addEventListener("click", () => {
    const total = readHistory().length;
    const totalPages = Math.max(1, Math.ceil(total / historyPager.pageSize) || 1);
    historyPager.page = Math.min(totalPages, historyPager.page + 1);
    renderHistory();
  });
  $("historyPagerLast")?.addEventListener("click", () => {
    const total = readHistory().length;
    historyPager.page = Math.max(1, Math.ceil(total / historyPager.pageSize) || 1);
    renderHistory();
  });
}

function renderHistory() {
  const host = $("historyList");
  if (!host) return;
  bindHistoryPager();

  const list = readHistory();
  const meta = $("historyMeta");
  if (meta) meta.textContent = `${list.length} registro${list.length === 1 ? "" : "s"}`;

  renderHistoryPager(list.length);

  if (!list.length) {
    host.innerHTML = `<p class="history-empty">Aún no hay validaciones en este dispositivo.</p>`;
    return;
  }

  const start = (historyPager.page - 1) * historyPager.pageSize;
  const pageItems = list.slice(start, start + historyPager.pageSize);

  host.innerHTML = pageItems
    .map((item) => {
      const rojo = Number(item.rojo) || 0;
      const aviso = Number(item.aviso) || 0;
      const badges = [
        rojo > 0
          ? `<span class="history-badge history-badge--rojo">Rojo ${rojo}</span>`
          : `<span class="history-badge history-badge--ok">Sin rojo</span>`,
        aviso > 0
          ? `<span class="history-badge history-badge--aviso">Aviso ${aviso}</span>`
          : ""
      ]
        .filter(Boolean)
        .join("");

      return `
      <article class="history-item">
        <div>
          <p class="history-item__name">${escapeHtml(item.fileName || "Archivo")}</p>
          <div class="history-item__meta">
            <span>${escapeHtml(item.sheetName || "Hoja")}</span>
            <span>·</span>
            <span>${escapeHtml(item.rows ?? 0)} filas</span>
          </div>
          <div class="history-item__badges">${badges}</div>
        </div>
        <time>${escapeHtml(item.at || "")}</time>
      </article>`;
    })
    .join("");
}

function uniqueSortedDates(rows) {
  return [...new Set(rows.map((r) => String(r.fecha || "").trim()).filter(Boolean))].sort((a, b) => {
    const pa = a.split("/").reverse().join("");
    const pb = b.split("/").reverse().join("");
    return pa.localeCompare(pb);
  });
}

function buildKpisFromRows(rows) {
  // Solo se trabaja con COSTO DE COSECHA para los KPI de actividad
  const costoRows = rows.filter((r) => r.esCostoCosecha);

  const supervisores = countSupervisoresCosto(rows);
  const scaner = countScanerCosto(rows);
  const cosecha = countCosechaCosto(rows);

  // Error / Extra por persona-día (suma de turnos), solo COSTO DE COSECHA
  const dayStatus = new Map();
  costoRows.forEach((row) => {
    const key = `${row.documento}|${row.fecha || ""}|${row.macroPartida || ""}`;
    if (!dayStatus.has(key)) dayStatus.set(key, row.status);
  });

  return {
    total: rows.length,
    costoCosecha: costoRows.length,
    supervisores,
    scaner,
    cosecha,
    rojo: [...dayStatus.values()].filter((s) => s === "rojo").length,
    aviso: [...dayStatus.values()].filter((s) => s === "aviso").length,
    posibleSalida: [...dayStatus.values()].filter((s) => s === "posible-salida").length,
    duplicados: new Set(
      costoRows
        .filter((r) => (r.flags || []).includes("duplicado"))
        .map((r) => r.documento)
        .filter(Boolean)
    ).size,
    cesados: rows.filter((r) => (r.flags || []).includes("cesado")).length,
    fechas: uniqueSortedDates(rows)
  };
}

function renderFechaMeta(rows) {
  const host = $("tareoFechaMeta");
  if (!host) return;
  const fechas = uniqueSortedDates(rows);
  if (!fechas.length) {
    host.hidden = true;
    host.textContent = "";
    return;
  }
  host.hidden = false;
  host.textContent =
    fechas.length === 1 ? `Fecha del tareo: ${fechas[0]}` : `Fechas del tareo: ${fechas.join(" · ")}`;
}

function syncFilterTips(filteredCount) {
  const total = state.validated?.rows?.length || 0;
  const f = readFilters();
  const living = `${filteredCount} de ${total} filas visibles`;

  const tips = [
    [
      "fltSupervisor",
      f.supervisor
        ? `Activo: ${f.supervisor}. ${living}. Elige “Todos” para limpiar.`
        : `Filtra por supervisor. Ahora ves ${living}.`
    ],
    [
      "fltFundo",
      f.fundo
        ? `Activo: fundo ${f.fundo}. ${living}.`
        : `Filtra por fundo / sede. Ahora ves ${living}.`
    ],
    [
      "fltFecha",
      f.fecha
        ? `Activo: fecha ${f.fecha}. ${living}.`
        : `Filtra por fecha del tareo. Ahora ves ${living}.`
    ],
    [
      "fltEstado",
      f.estado === "ok"
        ? `Mostrando solo OK (suma = 9.6 o 11.6). ${living}.`
        : f.estado === "posible-salida"
          ? `Mostrando posibles pases (suma < 9.6). ${living}.`
          : f.estado === "aviso"
            ? `Mostrando advertencias (exacto 10.1 / 10.6 / 12). ${living}.`
            : f.estado === "rojo"
              ? `Mostrando errores (suma > 12 u horario incompleto). ${living}.`
              : `Filtra OK, Posible pase, Advertencia o Error. Ahora ves ${living}.`
    ],
    [
      "fltSearch",
      f.search
        ? `Buscando “${f.search}”. ${living}.`
        : `Busca por DNI o nombre. Ahora ves ${living}.`
    ]
  ];

  tips.forEach(([id, text]) => {
    const el = $(id);
    const host = el?.closest(".has-tip");
    if (host) host.setAttribute("data-tip", text);
  });
}

function applyFiltersObject(filters) {
  if (!filters) return;
  const setVal = (id, value) => {
    const el = $(id);
    if (el && !el.disabled) el.value = value || "";
  };
  setVal("fltSupervisor", filters.supervisor);
  setVal("fltFundo", filters.fundo);
  setVal("fltFecha", filters.fecha);
  setVal("fltEstado", filters.estado);
  setVal("fltTipoLote", filters.tipo);
  const search = $("fltSearch");
  if (search) search.value = filters.search || "";
}

function syncErrorFocusUi(totalErrors) {
  const countEl = $("errorFocusCount");
  const bar = $("errorFocusBar");
  const btnVer = $("btnVerErrores");
  const btnBack = $("btnRegresarErrores");

  if (countEl) {
    countEl.textContent =
      totalErrors === 1 ? "1 error" : `${totalErrors} errores`;
  }
  bar?.classList.toggle("is-active", state.errorFocusMode);
  bar?.classList.toggle("has-errors", totalErrors > 0);

  if (btnVer) {
    btnVer.classList.toggle("is-hidden", state.errorFocusMode);
    btnVer.disabled = totalErrors === 0;
  }
  btnBack?.classList.toggle("is-hidden", !state.errorFocusMode);
}

function syncWarnFocusUi(totalWarnings) {
  const countEl = $("warnFocusCount");
  const bar = $("warnFocusBar");
  const btnVer = $("btnVerAvisos");
  const btnBack = $("btnRegresarAvisos");

  if (countEl) {
    countEl.textContent =
      totalWarnings === 1 ? "1 advertencia" : `${totalWarnings} advertencias`;
  }
  bar?.classList.toggle("is-active", state.warnFocusMode);
  bar?.classList.toggle("has-warnings", totalWarnings > 0);

  if (btnVer) {
    btnVer.classList.toggle("is-hidden", state.warnFocusMode);
    btnVer.disabled = totalWarnings === 0;
  }
  btnBack?.classList.toggle("is-hidden", !state.warnFocusMode);
}

function syncPaseFocusUi(totalPases) {
  const countEl = $("paseFocusCount");
  const bar = $("paseFocusBar");
  const btnVer = $("btnVerPases");
  const btnBack = $("btnRegresarPases");

  if (countEl) {
    countEl.textContent =
      totalPases === 1 ? "1 posible salida" : `${totalPases} posibles salidas`;
  }
  bar?.classList.toggle("is-active", state.paseFocusMode);
  bar?.classList.toggle("has-pases", totalPases > 0);

  if (btnVer) {
    btnVer.classList.toggle("is-hidden", state.paseFocusMode);
    btnVer.disabled = totalPases === 0;
  }
  btnBack?.classList.toggle("is-hidden", !state.paseFocusMode);
}

function syncDupFocusUi(totalDups) {
  const countEl = $("dupFocusCount");
  const bar = $("dupFocusBar");
  const btnVer = $("btnVerDups");
  const btnBack = $("btnRegresarDups");

  if (countEl) {
    countEl.textContent = totalDups === 1 ? "1 duplicado" : `${totalDups} duplicados`;
  }
  bar?.classList.toggle("is-active", state.dupFocusMode);
  bar?.classList.toggle("has-dups", totalDups > 0);

  if (btnVer) {
    btnVer.classList.toggle("is-hidden", state.dupFocusMode);
    btnVer.disabled = totalDups === 0;
  }
  btnBack?.classList.toggle("is-hidden", !state.dupFocusMode);
}

function enterErrorFocusMode() {
  if (!state.validated) return;
  const total = countTotalErrors(state.validated.rows);
  if (!total) {
    window.alert("No hay registros con error.");
    return;
  }
  if (state.warnFocusMode) exitWarnFocusMode({ restoreSaved: false });
  if (state.paseFocusMode) exitPaseFocusMode({ restoreSaved: false });
  if (state.dupFocusMode) exitDupFocusMode({ restoreSaved: false });
  state.savedFiltersBeforeErrorFocus = readFilters();
  state.errorFocusMode = true;
  clearAllFilterControls();
  populateFilters(state, { errorOnly: true });
  const estado = $("fltEstado");
  if (estado) {
    estado.disabled = false;
    estado.innerHTML = `<option value="rojo">Error &gt; 12</option>`;
    estado.value = "rojo";
    estado.disabled = true;
  }
  refreshView();
}

function exitErrorFocusMode({ restoreSaved = true } = {}) {
  const saved = state.savedFiltersBeforeErrorFocus;
  state.errorFocusMode = false;
  state.savedFiltersBeforeErrorFocus = null;
  populateFilters(state, { errorOnly: false });
  if (restoreSaved && saved) applyFiltersObject(saved);
  else clearAllFilterControls();
  refreshView();
}

function enterWarnFocusMode() {
  if (!state.validated) return;
  const total = countTotalWarnings(state.validated.rows);
  if (!total) {
    window.alert("No hay registros con advertencia.");
    return;
  }
  if (state.errorFocusMode) exitErrorFocusMode({ restoreSaved: false });
  if (state.paseFocusMode) exitPaseFocusMode({ restoreSaved: false });
  if (state.dupFocusMode) exitDupFocusMode({ restoreSaved: false });
  state.savedFiltersBeforeWarnFocus = readFilters();
  state.warnFocusMode = true;
  clearAllFilterControls();
  populateFilters(state, { warnOnly: true });
  const estado = $("fltEstado");
  if (estado) {
    estado.disabled = false;
    estado.innerHTML = `<option value="aviso">Advertencia ≤ 12</option>`;
    estado.value = "aviso";
    estado.disabled = true;
  }
  refreshView();
}

function exitWarnFocusMode({ restoreSaved = true } = {}) {
  const saved = state.savedFiltersBeforeWarnFocus;
  state.warnFocusMode = false;
  state.savedFiltersBeforeWarnFocus = null;
  populateFilters(state, { errorOnly: false, warnOnly: false, paseOnly: false, dupOnly: false });
  if (restoreSaved && saved) applyFiltersObject(saved);
  else clearAllFilterControls();
  refreshView();
}

function enterPaseFocusMode() {
  if (!state.validated) return;
  const total = countTotalPosibleSalidas(state.validated.rows);
  if (!total) {
    window.alert("No hay registros con posible pase de salida (< 9.6 h).");
    return;
  }
  if (state.errorFocusMode) exitErrorFocusMode({ restoreSaved: false });
  if (state.warnFocusMode) exitWarnFocusMode({ restoreSaved: false });
  if (state.dupFocusMode) exitDupFocusMode({ restoreSaved: false });
  state.savedFiltersBeforePaseFocus = readFilters();
  state.paseFocusMode = true;
  clearAllFilterControls();
  populateFilters(state, { paseOnly: true });
  const estado = $("fltEstado");
  if (estado) {
    estado.disabled = false;
    estado.innerHTML = `<option value="posible-salida">Posible pase &lt; 9.6</option>`;
    estado.value = "posible-salida";
    estado.disabled = true;
  }
  refreshView();
}

function exitPaseFocusMode({ restoreSaved = true } = {}) {
  const saved = state.savedFiltersBeforePaseFocus;
  state.paseFocusMode = false;
  state.savedFiltersBeforePaseFocus = null;
  populateFilters(state, { errorOnly: false, warnOnly: false, paseOnly: false, dupOnly: false });
  if (restoreSaved && saved) applyFiltersObject(saved);
  else clearAllFilterControls();
  refreshView();
}

function enterDupFocusMode() {
  if (!state.validated) return;
  const total = countTotalDuplicados(state.validated.rows);
  if (!total) {
    window.alert("No hay turnos duplicados (mismo DNI + fecha + hora de inicio).");
    return;
  }
  if (state.errorFocusMode) exitErrorFocusMode({ restoreSaved: false });
  if (state.warnFocusMode) exitWarnFocusMode({ restoreSaved: false });
  if (state.paseFocusMode) exitPaseFocusMode({ restoreSaved: false });
  state.savedFiltersBeforeDupFocus = readFilters();
  state.dupFocusMode = true;
  clearAllFilterControls();
  populateFilters(state, { dupOnly: true });
  refreshView();
}

function exitDupFocusMode({ restoreSaved = true } = {}) {
  const saved = state.savedFiltersBeforeDupFocus;
  state.dupFocusMode = false;
  state.savedFiltersBeforeDupFocus = null;
  populateFilters(state, { errorOnly: false, warnOnly: false, paseOnly: false, dupOnly: false });
  if (restoreSaved && saved) applyFiltersObject(saved);
  else clearAllFilterControls();
  refreshView();
}

function restoreAllFilters() {
  state.errorFocusMode = false;
  state.warnFocusMode = false;
  state.paseFocusMode = false;
  state.dupFocusMode = false;
  state.savedFiltersBeforeErrorFocus = null;
  state.savedFiltersBeforeWarnFocus = null;
  state.savedFiltersBeforePaseFocus = null;
  state.savedFiltersBeforeDupFocus = null;
  populateFilters(state, { errorOnly: false, warnOnly: false, paseOnly: false, dupOnly: false });
  clearAllFilterControls();
  refreshView();
}

function refreshView({ keepPage = false } = {}) {
  if (!state.validated) return;
  if (!keepPage) resetPager();

  if (state.errorFocusMode) {
    const estado = $("fltEstado");
    if (estado) {
      if (estado.value !== "rojo") {
        estado.disabled = false;
        estado.innerHTML = `<option value="rojo">Error &gt; 12</option>`;
        estado.value = "rojo";
        estado.disabled = true;
      }
    }
  }

  if (state.warnFocusMode) {
    const estado = $("fltEstado");
    if (estado) {
      if (estado.value !== "aviso") {
        estado.disabled = false;
        estado.innerHTML = `<option value="aviso">Advertencia ≤ 12</option>`;
        estado.value = "aviso";
        estado.disabled = true;
      }
    }
  }

  if (state.paseFocusMode) {
    const estado = $("fltEstado");
    if (estado) {
      if (estado.value !== "posible-salida") {
        estado.disabled = false;
        estado.innerHTML = `<option value="posible-salida">Posible pase &lt; 9.6</option>`;
        estado.value = "posible-salida";
        estado.disabled = true;
      }
    }
  }

  const filters = readFilters();
  if (state.dupFocusMode) filters.soloDuplicados = true;
  const filtered = filterRows(state.validated.rows, filters);
  renderFechaMeta(state.validated.rows);
  renderKpis(filtered);
  renderTable(state, filtered, { expandDuplicates: state.dupFocusMode });
  syncFilterTips(filtered.length);
  syncErrorFocusUi(countTotalErrors(state.validated.rows));
  syncWarnFocusUi(countTotalWarnings(state.validated.rows));
  syncPaseFocusUi(countTotalPosibleSalidas(state.validated.rows));
  syncDupFocusUi(countTotalDuplicados(state.validated.rows));
}

function revalidate() {
  if (!state.parsed) return;
  state.validated = validateDataset(state.parsed);
  populateFilters(state, {
    errorOnly: state.errorFocusMode,
    warnOnly: state.warnFocusMode,
    paseOnly: state.paseFocusMode,
    dupOnly: state.dupFocusMode
  });
  refreshView();
}

function bindRowSelection() {
  document.querySelectorAll("[data-row-select]").forEach((cb) => {
    cb.addEventListener("change", () => {
      const idx = Number(cb.value);
      if (cb.checked) {
        if (!state.selectedRowIndexes.includes(idx)) state.selectedRowIndexes.push(idx);
      } else {
        state.selectedRowIndexes = state.selectedRowIndexes.filter((x) => x !== idx);
      }
    });
  });
}

function applySessionTag(tipo) {
  if (!state.validated) return;
  const selected = new Set(state.selectedRowIndexes);
  if (!selected.size) {
    window.alert("Selecciona filas en la tabla (checkbox) para etiquetar.");
    return;
  }
  state.parsed.rows.forEach((row) => {
    if (selected.has(row.rowIndex)) row.sessionTipo = tipo;
  });
  revalidate();
}

function applyVariedad() {
  if (!state.validated) return;
  const selected = new Set(state.selectedRowIndexes);
  if (!selected.size) {
    window.alert("Selecciona filas en la tabla para agregar variedad.");
    return;
  }
  const value = window.prompt("Nombre de variedad:");
  if (!value) return;
  state.parsed.rows.forEach((row) => {
    if (selected.has(row.rowIndex)) row.sessionVariedad = value.trim();
  });
  revalidate();
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function formatHourExport(value) {
  if (value == null || value === "") return "";
  const n = Number(value);
  if (!Number.isFinite(n)) return String(value);
  return String(Math.round(n * 1e3) / 1e3);
}

function stackExportHtml(times) {
  const list = Array.isArray(times) ? times.filter(Boolean) : [];
  return list.map((t) => escapeXml(t)).join("<br/>");
}

function estadoLabel(status) {
  if (status === "rojo") return "error";
  if (status === "aviso") return "aviso";
  if (status === "posible-salida") return "posible pase";
  return "ok";
}

function downloadBlobFile(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

/**
 * Export HTML que Excel abre bien, con colores de celda.
 * cells: string | { html, tone?: 'danger'|'warn'|'ok'|'', tip?: string }
 */
function downloadHtmlExcel({ filename, sheetName, headers, rows }) {
  const headHtml = headers
    .map(
      (h) =>
        `<th style="background:#145a34;color:#ffffff;font-weight:700;border:1px solid #0f4a2b;padding:7px 9px;text-align:left;">${escapeXml(h)}</th>`
    )
    .join("");

  const toneCss = {
    danger: "background:#fecaca !important;color:#991b1b;font-weight:700;",
    warn: "background:#fde68a !important;color:#92400e;font-weight:600;",
    pase: "background:#bae6fd !important;color:#075985;font-weight:600;",
    ok: "background:#dcfce7 !important;color:#166534;font-weight:600;",
    softDanger: "background:#fff5f5 !important;",
    softWarn: "background:#fffaf0 !important;",
    softPase: "background:#f0f9ff !important;"
  };

  const bodyHtml = rows
    .map((row, idx) => {
      const stripe = idx % 2 === 1 ? "#fafbfc" : "#ffffff";
      const rowTone = row.rowTone || "";
      const rowBg = toneCss[rowTone] || `background:${stripe};`;
      const titleAttr = row.rowTip ? ` title="${escapeXml(row.rowTip)}"` : "";

      const tds = (row.cells || [])
        .map((cell) => {
          const isObj = cell && typeof cell === "object" && !Array.isArray(cell);
          const html = isObj ? cell.html ?? escapeXml(cell.value ?? "") : escapeXml(cell);
          const tone = isObj ? cell.tone || "" : "";
          const tip = isObj && cell.tip ? ` title="${escapeXml(cell.tip)}"` : "";
          const numAsText = isObj && cell.text === true ? "mso-number-format:'\\@';" : "";
          const style = `border:1px solid #e5e7eb;padding:6px 8px;vertical-align:top;white-space:pre-line;${rowBg}${
            toneCss[tone] || ""
          }${numAsText}`;
          return `<td style="${style}"${tip}>${html}</td>`;
        })
        .join("");

      return `<tr${titleAttr}>${tds}</tr>`;
    })
    .join("");

  const safeSheet = String(sheetName || "Hoja1").slice(0, 31);
  const html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="UTF-8" />
<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
<x:Name>${escapeXml(safeSheet)}</x:Name>
<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
<style>
  table { border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
</style>
</head>
<body>
<table>
  <thead><tr>${headHtml}</tr></thead>
  <tbody>${bodyHtml || `<tr><td colspan="${headers.length}" style="padding:8px;color:#6b7280;">Sin datos para exportar.</td></tr>`}</tbody>
</table>
</body>
</html>`;

  // BOM ayuda a Excel a leer tildes/ñ
  const blob = new Blob(["\uFEFF", html], { type: "application/vnd.ms-excel;charset=utf-8;" });
  downloadBlobFile(filename, blob);
}

function resolveSumaFlag(row, total) {
  const hourKey = state.parsed?.dayLabels?.[0] || "Suma de Horas Pago";
  const fromFlags = row.dayFlags?.[hourKey];
  if (fromFlags && fromFlags !== "ok") return fromFlags;

  if (!row.esCostoCosecha) return fromFlags || "ok";

  return classifyDayHours(total).flag;
}

/** Exporta Detalle igual al frontend, con celdas de error en rojo. */
function downloadDetalleLikeFrontend() {
  const filters = readFilters();
  const filtered = filterRows(state.validated.rows, filters);
  const dayRows = collapseToDayRows(filtered);
  const stamp = new Date().toISOString().slice(0, 10);
  const hourKey = state.parsed?.dayLabels?.[0] || "Suma de Horas Pago";

  const rows = dayRows.map((row) => {
    const total = row.sumaHorasPago ?? row.hoursByDay?.[hourKey] ?? row.totalDia ?? row.horas;
    const sumaFlag = resolveSumaFlag(row, total);
    const iniFlag = row.dayFlags?.horaInicio;
    const finFlag = row.dayFlags?.horaFin;
    const rowTone =
      row.status === "rojo"
        ? "softDanger"
        : row.status === "aviso"
          ? "softWarn"
          : row.status === "posible-salida"
            ? "softPase"
            : "";

    return {
      rowTone,
      rowTip: row.tipHoras || row.tipHoraInicio || row.tipHoraFin || "",
      cells: [
        {
          value:
            row.documentoVacio || row.dayFlags?.documento === "rojo"
              ? `(vacío)${row.codigoTrabajador ? ` · ${row.codigoTrabajador}` : ""}`
              : row.documento,
          text: true,
          tone: row.documentoVacio || row.dayFlags?.documento === "rojo" ? "danger" : "",
          tip: row.tipDocumento || ""
        },
        {
          value: String(row.trabajador || "").trim()
            ? row.trabajador
            : String(row.documento || "").trim() || String(row.codigoTrabajador || "").trim()
              ? "(vacío)"
              : "",
          tone:
            (String(row.documento || "").trim() || String(row.codigoTrabajador || "").trim()) &&
            (!String(row.trabajador || "").trim() || row.dayFlags?.trabajador === "rojo")
              ? "danger"
              : "",
          tip: row.tipTrabajador || ""
        },
        row.supervisor,
        row.fundo,
        row.macroPartida,
        row.actividad || "",
        {
          value: String(row.ceco || "").trim() ? row.ceco : "(vacío)",
          tone: !String(row.ceco || "").trim() || row.dayFlags?.ceco === "rojo" ? "danger" : "",
          tip: row.tipCeco || ""
        },
        row.fecha || "",
        {
          html: stackExportHtml(row.horasInicioDetalle),
          tone: iniFlag === "rojo" ? "danger" : "",
          tip: row.tipHoraInicio || ""
        },
        {
          html: stackExportHtml(row.horasFinDetalle),
          tone: finFlag === "rojo" ? "danger" : "",
          tip: row.tipHoraFin || ""
        },
        row.turnosDetalle || formatHourExport(row.horasTurno),
        {
          value: formatHourExport(total),
          tone:
            sumaFlag === "rojo"
              ? "danger"
              : sumaFlag === "aviso" || sumaFlag === "aviso-hora"
                ? "warn"
                : sumaFlag === "posible-salida"
                  ? "pase"
                  : "",
          tip: row.tipHoras || ""
        },
        {
          value: estadoLabel(row.status),
          tone:
            row.status === "rojo"
              ? "danger"
              : row.status === "aviso"
                ? "warn"
                : row.status === "posible-salida"
                  ? "pase"
                  : "ok"
        }
      ]
    };
  });

  downloadHtmlExcel({
    filename: `QBerries_Detalle_${stamp}.xls`,
    sheetName: "DetalleDiario",
    headers: [
      "Documento",
      "Trabajador",
      "Supervisor",
      "Fundo",
      "Macro Partida",
      "Actividad",
      "CECO",
      "Fecha",
      "Hora Inicio",
      "Hora Fin",
      "Horas Pago (turnos)",
      "Suma de Horas Pago",
      "Estado"
    ],
    rows
  });
}

function downloadResumenExcel() {
  const stamp = new Date().toISOString().slice(0, 10);
  const mainFilters = readFilters();
  const { groups } = getFilteredResumenData(state.validated, mainFilters);

  if (!groups.length) {
    window.alert("No hay filas en el resumen con los filtros actuales.");
    return;
  }

  const rows = groups.map((g) => ({
    rowTone: g.errores > 0 ? "softDanger" : g.avisos > 0 ? "softWarn" : "",
    cells: [
      g.fundo,
      g.supervisor,
      g.planillas,
      g.trabajadores,
      {
        value: g.errores,
        tone: g.errores > 0 ? "danger" : ""
      },
      {
        value: g.avisos,
        tone: g.avisos > 0 ? "warn" : ""
      },
      g.ok
    ]
  }));

  downloadHtmlExcel({
    filename: `QBerries_Resumen_Supervisores_${stamp}.xls`,
    sheetName: "Resumen",
    headers: ["Fundo", "Supervisor", "Planillas", "Trabajadores", "Errores", "Extras", "OK"],
    rows
  });
}

function downloadHallazgosExcel() {
  const stamp = new Date().toISOString().slice(0, 10);
  const f = state.validated.findings || {};
  const filters = readFilters();

  const inFilter = (item) => {
    const empty =
      !filters.macro &&
      !filters.supervisor &&
      !filters.fundo &&
      !filters.fecha &&
      !filters.estado &&
      !filters.search;
    if (empty) return true;
    if (filters.supervisor && item.supervisor && item.supervisor !== filters.supervisor) return false;
    if (filters.fecha && item.fecha && item.fecha !== filters.fecha) return false;
    if (filters.search) {
      const blob = `${item.documento || ""} ${item.trabajador || ""} ${(item.trabajadores || []).join(" ")}`.toLowerCase();
      if (!blob.includes(String(filters.search).toLowerCase())) return false;
    }
    return true;
  };

  const pushRows = [];

  const add = (tipo, item, detalle, valor, tone) => {
    if (!inFilter(item)) return;
    pushRows.push({
      rowTone: tone === "danger" ? "softDanger" : tone === "warn" ? "softWarn" : "",
      cells: [
        { value: tipo, tone },
        { value: item.documento, text: true, tone },
        item.trabajador || (item.trabajadores || []).join(", "),
        item.supervisor || "",
        detalle ?? "",
        valor ?? "",
        item.rowIndex ?? ""
      ]
    });
  };

  (f.overHours || []).forEach((i) =>
    add("Hora no exacta", i, i.day || "Suma de Horas Pago", i.hours, "danger")
  );
  (f.horario || []).forEach((i) =>
    add("Horario", i, i.detalle, `${i.inicios || ""} → ${i.fines || ""}`, "danger")
  );
  (f.duplicates || []).forEach((i) =>
    add("Duplicado DNI", i, i.fecha || "", i.count, "danger")
  );
  (f.naFails || []).forEach((i) => add("N.A.", i, i.field, i.group, "danger"));
  (f.cecoVacio || []).forEach((i) =>
    add("CECO vacío", i, i.macroPartida || i.actividad || "Columna U", "", "danger")
  );
  (f.cesados || []).forEach((i) => add("Cesado", i, "", "", "warn"));
  (f.minoritaria || []).forEach((i) => add("Menoritaria", i, i.macroPartida, "", "warn"));
  (f.overBase || []).forEach((i) =>
    add(`Extra ${i.extra || ""}`.trim(), i, i.day || "Suma de Horas Pago", i.hours, "warn")
  );

  downloadHtmlExcel({
    filename: `QBerries_Hallazgos_${stamp}.xls`,
    sheetName: "Hallazgos",
    headers: ["Tipo", "Documento", "Trabajador", "Supervisor", "Detalle", "Valor", "Fila"],
    rows: pushRows
  });
}

function downloadReport(kind) {
  try {
    if (!state.validated) {
      window.alert("Primero sube un Excel de tareo.");
      return;
    }

    if (kind === "detalle") {
      downloadDetalleLikeFrontend();
      return;
    }
    if (kind === "resumen") {
      downloadResumenExcel();
      return;
    }
    if (kind === "hallazgos") {
      downloadHallazgosExcel();
      return;
    }

    window.alert("Tipo de reporte no reconocido.");
  } catch (err) {
    console.error(err);
    window.alert(`No se pudo exportar: ${err?.message || err}`);
  }
}

async function handleFile(file) {
  if (!file) return;
  const buffer = await file.arrayBuffer();
  const parsed = parseExcelBuffer(buffer, file.name);
  state.parsed = parsed;
  state.fileName = file.name;
  state.selectedRowIndexes = [];
  state.errorFocusMode = false;
  state.warnFocusMode = false;
  state.paseFocusMode = false;
  state.dupFocusMode = false;
  state.savedFiltersBeforeErrorFocus = null;
  state.savedFiltersBeforeWarnFocus = null;
  state.savedFiltersBeforePaseFocus = null;
  state.savedFiltersBeforeDupFocus = null;
  state.validated = validateDataset(parsed);

  $("uploadZone")?.classList.add("is-hidden");
  $("validacionWorkspace")?.classList.remove("is-hidden");

  populateFilters(state, { errorOnly: false });
  refreshView();

  writeHistory({
    fileName: file.name,
    sheetName: parsed.sheetName,
    rows: state.validated.kpis.total,
    rojo: state.validated.kpis.rojo,
    aviso: state.validated.kpis.aviso,
    at: new Date().toLocaleString("es-PE")
  });
  renderHistory();

  const costo = state.validated.kpis.costoCosecha;
  const total = state.validated.kpis.total;
  const metaCosto = state.parsed?.meta?.costoCosechaCount;
  showSuccessModal(
    `Listo. Costo de cosecha: ${costo}${metaCosto != null && metaCosto !== costo ? ` (lectura ${metaCosto})` : ""} · Total filas: ${total}. Los contadores cambian al filtrar.`
  );
}

function bindUpload() {
  const input = $("inputExcel");
  const pick = $("btnPickExcel");
  const card = $("uploadDropCard");

  pick?.addEventListener("click", () => input?.click());
  input?.addEventListener("change", () => {
    const file = input.files?.[0];
    handleFile(file).catch((err) => window.alert(err.message || String(err)));
    input.value = "";
  });

  ["dragenter", "dragover"].forEach((evt) => {
    card?.addEventListener(evt, (e) => {
      e.preventDefault();
      card.classList.add("is-dragover");
    });
  });
  ["dragleave", "drop"].forEach((evt) => {
    card?.addEventListener(evt, (e) => {
      e.preventDefault();
      card.classList.remove("is-dragover");
    });
  });
  card?.addEventListener("drop", (e) => {
    const file = e.dataTransfer?.files?.[0];
    handleFile(file).catch((err) => window.alert(err.message || String(err)));
  });
}

function bindUi() {
  [
    "fltSupervisor",
    "fltFundo",
    "fltFecha",
    "fltEstado",
    "fltTipoLote",
    "fltSearch"
  ].forEach((id) => {
    $(id)?.addEventListener("input", refreshView);
    $(id)?.addEventListener("change", refreshView);
  });

  $("btnMarkChina")?.addEventListener("click", () => applySessionTag("china"));
  $("btnMarkConv")?.addEventListener("click", () => applySessionTag("convencional"));
  $("btnSetVariedad")?.addEventListener("click", applyVariedad);
  $("btnReportResumen")?.addEventListener("click", () => {
    if (!state.validated) {
      window.alert("Primero sube un Excel de tareo.");
      return;
    }
    syncResumenFiltersFromMain(state.validated);
    renderResumenView(state.validated, readFilters());
    openResumenModal();
  });
  $("btnReportHallazgos")?.addEventListener("click", () => downloadReport("hallazgos"));
  $("btnReportDetalle")?.addEventListener("click", () => downloadReport("detalle"));
  $("btnExportar")?.addEventListener("click", () => downloadReport("detalle"));
  $("btnNewUpload")?.addEventListener("click", () => {
    state.parsed = null;
    state.validated = null;
    state.selectedRowIndexes = [];
    state.errorFocusMode = false;
    state.warnFocusMode = false;
    state.paseFocusMode = false;
    state.dupFocusMode = false;
    state.savedFiltersBeforeErrorFocus = null;
    state.savedFiltersBeforeWarnFocus = null;
    state.savedFiltersBeforePaseFocus = null;
    state.savedFiltersBeforeDupFocus = null;
    closeResumenModal();
    $("validacionWorkspace")?.classList.add("is-hidden");
    $("uploadZone")?.classList.remove("is-hidden");
  });

  $("btnVerErrores")?.addEventListener("click", () => enterErrorFocusMode());
  $("btnRegresarErrores")?.addEventListener("click", () => exitErrorFocusMode({ restoreSaved: true }));
  $("btnVerAvisos")?.addEventListener("click", () => enterWarnFocusMode());
  $("btnRegresarAvisos")?.addEventListener("click", () => exitWarnFocusMode({ restoreSaved: true }));
  $("btnVerPases")?.addEventListener("click", () => enterPaseFocusMode());
  $("btnRegresarPases")?.addEventListener("click", () => exitPaseFocusMode({ restoreSaved: true }));
  $("btnVerDups")?.addEventListener("click", () => enterDupFocusMode());
  $("btnRegresarDups")?.addEventListener("click", () => exitDupFocusMode({ restoreSaved: true }));
  $("btnRestaurarTodo")?.addEventListener("click", () => restoreAllFilters());

  $("btnCloseSuccessModal")?.addEventListener("click", hideSuccessModal);
  $("btnCloseKpiHelp")?.addEventListener("click", hideKpiHelp);
  $("btnKpiHelpOk")?.addEventListener("click", hideKpiHelp);

  document.querySelectorAll("[data-close-modal]").forEach((el) => {
    el.addEventListener("click", () => {
      const which = el.getAttribute("data-close-modal");
      if (which === "kpi") hideKpiHelp();
      else if (which === "resumen") closeResumenModal();
      else hideSuccessModal();
    });
  });

  // Info de KPIs (delegación: se re-renderizan)
  $("kpiRow")?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-kpi-help]");
    if (!btn) return;
    openKpiHelp(btn.getAttribute("data-kpi-help"));
  });

  bindTablePager(() => refreshView({ keepPage: true }));

  bindResumenUi({
    getValidated: () => state.validated,
    getMainFilters: () => readFilters(),
    onExport: () => downloadReport("resumen")
  });

  window.addEventListener("qb:route-changed", (evt) => {
    if (evt.detail?.route === "historial") renderHistory();
  });
}

bindUpload();
bindUi();
renderHistory();
if (getCurrentRoute() === "historial") renderHistory();
