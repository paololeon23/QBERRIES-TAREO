/** Orquestación: upload, modal, reportes, historial. */

import { parseExcelBuffer, matchExactHourStep, classifyDayHours, isNombreTrabajadorVacio } from "./excel-parser.js?v=20260909b1";
import { validateDataset } from "./validacion-rules.js?v=20260909b1";
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
  clearAllFilterControls,
  getActividadFilterState,
  setSelectedActividades,
  resetActividadFilterToDefaults,
  resetActividadFilterState,
  defaultSelectedActividades,
  syncActividadFilterButton
} from "./validacion-table.js?v=20260916p1";
import { countSupervisoresCosto, countScanerCosto, countCosechaCosto } from "./validacion-kpi.js";
import {
  openResumenModal,
  closeResumenModal,
  syncResumenFiltersFromMain,
  renderResumenView,
  bindResumenUi,
  getFilteredResumenData,
  getResumenTableRowsForExport,
  resetPlanillasFaltantesState,
  isAdminCosechaSupervisor,
  personaUnicaKey,
  keysConSupervisorCosecha,
  esExcluidoDeCosechadores,
  collectCosechadoresUnicosKeys
} from "./validacion-resumen.js?v=20260916c804b";
import { isActividadHorarioNuevo } from "./horarios-cosecha.js?v=20260909b1";
import { getCurrentRoute } from "./shell.js";

const HISTORY_KEY = "qb-validacion-history";

const state = {
  parsed: null,
  validated: null,
  selectedRowIndexes: [],
  fileName: "",
  uploading: false,
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
      <p>Equipos en <strong>COSTO DE COSECHA</strong>: cada supervisor = 1 grupo (1 scaner + cosecha + supervisor de cosecha).</p>
      <ul>
        <li>Cuenta <strong>grupos</strong>, no filas repetidas.</li>
      </ul>`
  },
  scaner: {
    title: "Scaner",
    tone: "info",
    html: `
      <p>Grupos con <strong>SCANER</strong> en COSTO DE COSECHA.</p>
      <ul>
        <li>1 scaner por equipo → debe coincidir con Supervisores y Cosecha si el grupo está completo.</li>
      </ul>`
  },
  cosecha: {
    title: "Cosecha",
    tone: "info",
    html: `
      <p>Grupos con actividad <strong>COSECHA</strong> en COSTO DE COSECHA.</p>
      <ul>
        <li>Cuenta <strong>equipos</strong> con cosechadores (1 por grupo), no el total de personas en cosecha.</li>
        <li>Debe coincidir con Supervisores y Scaner si cada grupo está completo.</li>
      </ul>`
  },
  error: {
    title: "Error ≠ exacto",
    tone: "danger",
    html: `
      <p>Personas-día en <strong>error</strong> dentro de <strong>COSTO DE COSECHA</strong>: suma distinta de <strong>9.6 / 10.1 / 10.6 / 11.6 / 12</strong> (ej. 9.63, 11.37), suma &gt; 12 h, horario incompleto/inválido, CECO vacío, <strong>Documento vacío</strong>, <strong>DNI/Código sin nombre</strong> (vacío o “NO VERIFICADO”) o <strong>actividad no permitida</strong> (fuera de las 9 autorizadas).</p>
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
      "btnFltActividad",
      (f.actividades || []).length
        ? `Activo: ${(f.actividades || []).length} actividad(es). Cosecha → 06:45/17:21 · otras → 06:30/17:06. ${living}.`
        : `Sin actividades seleccionadas. ${living}.`
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

const tareoActModal = {
  draft: [],
  search: ""
};

function escapeActHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderTareoActFilterList() {
  const list = $("tareoActFilterList");
  const count = $("tareoActFilterCount");
  if (!list) return;
  const { available } = getActividadFilterState();
  const q = tareoActModal.search.trim().toLowerCase();
  const filtered = q
    ? available.filter((a) => a.toLowerCase().includes(q))
    : available;
  const draftSet = new Set(tareoActModal.draft);

  if (!filtered.length) {
    list.innerHTML = `<p class="rt-filter-modal__empty">Sin coincidencias</p>`;
  } else {
    list.innerHTML = filtered
      .map((a) => {
        const safe = escapeActHtml(a);
        const isOn = draftSet.has(a);
        const horario = isActividadHorarioNuevo(a)
          ? "LICAPA 06:45 · II/III 06:30"
          : "06:30 / 17:06";
        return `<label class="rt-filter-modal__option${isOn ? " is-checked" : ""}">
        <input type="checkbox" class="rt-filter-modal__check" value="${safe}" ${isOn ? "checked" : ""} />
        <span class="rt-filter-modal__box" aria-hidden="true">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
        </span>
        <span class="rt-filter-modal__label">${safe}<small class="tareo-act-horario">${horario}</small></span>
      </label>`;
      })
      .join("");
  }

  if (count) {
    const n = tareoActModal.draft.length;
    count.textContent = n === 1 ? "1 seleccionada" : `${n} seleccionadas`;
  }
}

function openTareoActFilterModal() {
  const modal = $("modalTareoActFilter");
  if (!modal) return;
  const { selected } = getActividadFilterState();
  tareoActModal.draft = [...selected];
  tareoActModal.search = "";
  const search = $("tareoActFilterSearch");
  if (search) search.value = "";
  renderTareoActFilterList();
  modal.hidden = false;
}

function closeTareoActFilterModal() {
  const modal = $("modalTareoActFilter");
  if (modal) modal.hidden = true;
}

function bindTareoActFilterUi() {
  $("btnFltActividad")?.addEventListener("click", () => openTareoActFilterModal());
  $("btnCloseTareoActFilter")?.addEventListener("click", closeTareoActFilterModal);
  $("btnTareoActCancel")?.addEventListener("click", closeTareoActFilterModal);
  document.querySelectorAll('[data-close-modal="tareo-act"]').forEach((el) => {
    el.addEventListener("click", closeTareoActFilterModal);
  });

  $("tareoActFilterSearch")?.addEventListener("input", (e) => {
    tareoActModal.search = e.target.value || "";
    renderTareoActFilterList();
  });

  $("btnTareoActDefault")?.addEventListener("click", () => {
    const { available } = getActividadFilterState();
    tareoActModal.draft = defaultSelectedActividades(available);
    renderTareoActFilterList();
  });
  $("btnTareoActAll")?.addEventListener("click", () => {
    tareoActModal.draft = [...getActividadFilterState().available];
    renderTareoActFilterList();
  });
  $("btnTareoActNone")?.addEventListener("click", () => {
    tareoActModal.draft = [];
    renderTareoActFilterList();
  });

  $("tareoActFilterList")?.addEventListener("change", (e) => {
    const input = e.target;
    if (!(input instanceof HTMLInputElement) || !input.classList.contains("rt-filter-modal__check")) return;
    const value = input.value;
    const set = new Set(tareoActModal.draft);
    if (input.checked) set.add(value);
    else set.delete(value);
    tareoActModal.draft = [...set];
    input.closest(".rt-filter-modal__option")?.classList.toggle("is-checked", input.checked);
    const count = $("tareoActFilterCount");
    if (count) {
      const n = tareoActModal.draft.length;
      count.textContent = n === 1 ? "1 seleccionada" : `${n} seleccionadas`;
    }
  });

  $("btnTareoActApply")?.addEventListener("click", () => {
    setSelectedActividades(tareoActModal.draft);
    closeTareoActFilterModal();
    refreshView();
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
  setVal("fltEstado", filters.estado);
  setVal("fltTipoLote", filters.tipo);
  const search = $("fltSearch");
  if (search) search.value = filters.search || "";
  if (Array.isArray(filters.actividades)) {
    setSelectedActividades(filters.actividades);
  }
}

function syncErrorFocusUi(totalErrors) {
  const countEl = $("errorFocusCount");
  const bar = $("errorFocusBar");
  const btnVer = $("btnVerErrores");
  const btnPersonas = $("btnPersonasErrores");
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
  if (btnPersonas) {
    btnPersonas.disabled = totalErrors === 0;
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

function getScopeFilters(base = null, { ignoreActividades = false } = {}) {
  const f = base || readFilters();
  return {
    supervisor: f.supervisor || "",
    fundo: f.fundo || "",
    fecha: "",
    macro: f.macro || "",
    actividad: f.actividad || "",
    actividades: ignoreActividades
      ? undefined
      : Array.isArray(f.actividades)
        ? [...f.actividades]
        : [],
    dia: "",
    estado: "",
    tipo: f.tipo || "",
    search: f.search || "",
    soloDuplicados: false
  };
}

/**
 * Alcance para contadores (errores/avisos): fundo/supervisor/búsqueda.
 * Ignora filtro de actividades para no ocultar actividades no permitidas.
 */
function getScopedRows(baseFilters = null, opts = { ignoreActividades: true }) {
  if (!state.validated) return [];
  return filterRows(state.validated.rows, getScopeFilters(baseFilters, opts));
}

function lockEstadoFilter(value, labelHtml) {
  const estado = $("fltEstado");
  if (!estado) return;
  estado.disabled = false;
  estado.innerHTML = `<option value="${value}">${labelHtml}</option>`;
  estado.value = value;
  estado.disabled = true;
}

function enterErrorFocusMode() {
  if (!state.validated) return;
  const total = countTotalErrors(getScopedRows());
  if (!total) {
    window.alert("No hay registros con error en el filtro actual.");
    return;
  }
  if (state.warnFocusMode) exitWarnFocusMode({ restoreSaved: false });
  if (state.paseFocusMode) exitPaseFocusMode({ restoreSaved: false });
  if (state.dupFocusMode) exitDupFocusMode({ restoreSaved: false });
  state.savedFiltersBeforeErrorFocus = readFilters();
  state.errorFocusMode = true;
  populateFilters(state, { errorOnly: true });
  /* Errores: no ocultar por filtro de las 9 actividades (p. ej. sin actividad / no permitida) */
  const { available } = getActividadFilterState();
  setSelectedActividades(available);
  applyFiltersObject({
    ...state.savedFiltersBeforeErrorFocus,
    actividades: available,
    estado: "rojo"
  });
  lockEstadoFilter("rojo", "Error &gt; 12");
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
  const total = countTotalWarnings(getScopedRows());
  if (!total) {
    window.alert("No hay registros con advertencia en el filtro actual.");
    return;
  }
  if (state.errorFocusMode) exitErrorFocusMode({ restoreSaved: false });
  if (state.paseFocusMode) exitPaseFocusMode({ restoreSaved: false });
  if (state.dupFocusMode) exitDupFocusMode({ restoreSaved: false });
  state.savedFiltersBeforeWarnFocus = readFilters();
  state.warnFocusMode = true;
  populateFilters(state, { warnOnly: true });
  applyFiltersObject({
    ...state.savedFiltersBeforeWarnFocus,
    estado: "aviso"
  });
  lockEstadoFilter("aviso", "Advertencia ≤ 12");
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
  const total = countTotalPosibleSalidas(getScopedRows());
  if (!total) {
    window.alert("No hay posibles pases de salida (< 9.6 h) en el filtro actual.");
    return;
  }
  if (state.errorFocusMode) exitErrorFocusMode({ restoreSaved: false });
  if (state.warnFocusMode) exitWarnFocusMode({ restoreSaved: false });
  if (state.dupFocusMode) exitDupFocusMode({ restoreSaved: false });
  state.savedFiltersBeforePaseFocus = readFilters();
  state.paseFocusMode = true;
  populateFilters(state, { paseOnly: true });
  applyFiltersObject({
    ...state.savedFiltersBeforePaseFocus,
    estado: "posible-salida"
  });
  lockEstadoFilter("posible-salida", "Posible pase &lt; 9.6");
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
  const total = countTotalDuplicados(getScopedRows());
  if (!total) {
    window.alert("No hay turnos duplicados en el filtro actual.");
    return;
  }
  if (state.errorFocusMode) exitErrorFocusMode({ restoreSaved: false });
  if (state.warnFocusMode) exitWarnFocusMode({ restoreSaved: false });
  if (state.paseFocusMode) exitPaseFocusMode({ restoreSaved: false });
  state.savedFiltersBeforeDupFocus = readFilters();
  state.dupFocusMode = true;
  populateFilters(state, { dupOnly: true });
  applyFiltersObject(state.savedFiltersBeforeDupFocus);
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
  /* En foco de errores, no restringir por actividad (coincide con el contador). */
  if (state.errorFocusMode) filters.actividades = undefined;

  let filtered = filterRows(state.validated.rows, filters);

  /* Si el filtro de las 9 deja 0 filas pero hay data, ampliar a todas las actividades del Excel. */
  if (
    !filtered.length &&
    state.validated.rows.length &&
    !state.errorFocusMode &&
    !state.warnFocusMode &&
    !state.paseFocusMode &&
    !state.dupFocusMode &&
    Array.isArray(filters.actividades) &&
    filters.actividades.length > 0
  ) {
    const scoped = getScopedRows(filters);
    if (scoped.length) {
      const { available } = getActividadFilterState();
      if (available.length) {
        setSelectedActividades(available);
        filters.actividades = [...available];
        filtered = filterRows(state.validated.rows, filters);
      }
    }
  }

  const scoped = getScopedRows(filters);
  renderFechaMeta(state.validated.rows);
  renderKpis(filtered);
  renderTable(state, filtered, { expandDuplicates: state.dupFocusMode });
  syncFilterTips(filtered.length);
  syncErrorFocusUi(countTotalErrors(scoped));
  syncWarnFocusUi(countTotalWarnings(scoped));
  syncPaseFocusUi(countTotalPosibleSalidas(scoped));
  syncDupFocusUi(countTotalDuplicados(scoped));
  renderModulosCosechaIfOpen();
  renderLotesHoyIfOpen();
}

function normExactToken(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function isActividadCosechaExact(actividad) {
  return normExactToken(actividad) === "COSECHA";
}

function workerUniqueKey(row) {
  // Misma clave que Resumen (DNI → código → nombre)
  return personaUnicaKey(row);
}

/** Filas del alcance fundo (todas las actividades) para alinear con KPI Resumen. */
function rowsAlcanceFundo(fundoFiltro = "") {
  const fundoNorm = normExactToken(fundoFiltro);
  return (state.validated?.rows || []).filter((row) => {
    if (fundoNorm && normExactToken(row.fundo) !== fundoNorm) return false;
    return true;
  });
}

/** Personas únicas cosecha al estilo Resumen (misma regla: sin supervisor dual). */
function collectResumenCosechaKeys(fundoFiltro = "") {
  return collectCosechadoresUnicosKeys(rowsAlcanceFundo(fundoFiltro));
}

/** Solo módulos reales (MODULO 1, 2, 5…). No QBERRIES 01 ni otros textos. */
function isModuloCosechaValido(modulo) {
  const raw = String(modulo || "").trim();
  if (!raw) return false;
  const key = normExactToken(raw);
  if (!key || key.includes("QBERRIES")) return false;
  if (key === "(SIN MODULO)" || key === "SIN MODULO") return false;
  return /MODULO\s*\d+/.test(key) || /^M\s*\d+$/.test(key);
}

function moduloCosechaLabel(row) {
  const modulo = String(row?.modulo || "").trim();
  return isModuloCosechaValido(modulo) ? modulo : "";
}

/** Vista del modal módulos: resumen (COSECHA) | detalle (actividades). */
const modulosUiState = { view: "resumen" };

/** Actividades permitidas en vista detalle (pivote). */
const ACTIVIDADES_DETALLE_MODULO = new Set([
  "COSECHA",
  "ESTIBA KIA",
  "SCANER",
  "SUPERVISOR DE COSECHA",
  "CALIDAD"
]);

/** Etiqueta canónica de actividad (evita SCANER/SCANNER duplicados). */
function actividadDetalleLabel(actividad) {
  const a = normExactToken(actividad);
  if (!a) return "";
  if (a === "SCANNER" || a === "ESCANER" || a === "ESCANEAR") return "SCANER";
  if (a.includes("SUPERVISOR") && a.includes("COSECHA")) return "SUPERVISOR DE COSECHA";
  if (a.includes("CALIDAD")) return "CALIDAD";
  if (a.includes("ESTIBA") && a.includes("KIA")) return "ESTIBA KIA";
  if (a === "COSECHA") return "COSECHA";
  return a;
}

/**
 * Tabla detallada tipo pivote:
 * Módulo × Fundo × Actividad → personas únicas.
 * Si se movió, cuenta en destino. En origen COSECHA (solo si hubo salida):
 * "inicio → neto (quedaron después HH:MM)" con horas/cantidades del Excel del día.
 */
function buildModulosDetalleReport() {
  const rows = state.validated?.rows || [];
  const fundoFiltro = String(readFilters().fundo || "").trim();
  const fundoNorm = normExactToken(fundoFiltro);
  const excluirCosecha = keysConSupervisorCosecha(rowsAlcanceFundo(fundoFiltro));

  /** wKey → { mods, actsByMod } */
  const byWorker = new Map();
  let filasFuente = 0;

  rows.forEach((row) => {
    if (fundoNorm && normExactToken(row.fundo) !== fundoNorm) return;
    const actividad = actividadDetalleLabel(row.actividad);
    if (!actividad || !ACTIVIDADES_DETALLE_MODULO.has(actividad)) return;
    const wKey = workerUniqueKey(row);
    if (!wKey) return;
    const modulo = moduloCosechaLabel(row);
    if (!modulo) return;

    filasFuente += 1;
    const fundo = String(row.fundo || "").trim() || "(sin fundo)";
    const groupKey = `${normExactToken(fundo)}||${normExactToken(modulo)}`;
    const supervisor = String(row.supervisor || "").trim() || "(sin supervisor)";
    const startMin =
      row.horaInicioMin != null && Number.isFinite(row.horaInicioMin)
        ? row.horaInicioMin
        : Number.POSITIVE_INFINITY;

    if (!byWorker.has(wKey)) {
      byWorker.set(wKey, { mods: new Map(), actsByMod: new Map() });
    }
    const w = byWorker.get(wKey);
    const prev = w.mods.get(groupKey);
    if (!prev || startMin < prev.startMin) {
      w.mods.set(groupKey, { fundo, modulo, startMin, supervisor });
    }
    if (!w.actsByMod.has(groupKey)) w.actsByMod.set(groupKey, new Set());
    // Supervisor con fila COSECHA extra: no inflar conteo de cosechadores
    if (actividad === "COSECHA" && esExcluidoDeCosechadores(row, excluirCosecha)) return;
    w.actsByMod.get(groupKey).add(actividad);
  });

  const buckets = new Map();
  const visitMap = new Map(); // bKey → Set(wKey) actividad en ese módulo (inicio)
  const leaveHoraByMod = new Map(); // groupKey → min hora destino (salieron)
  const moveMap = new Map();

  byWorker.forEach((w, wKey) => {
    const modList = [...w.mods.entries()];
    if (!modList.length) return;

    let first = modList[0];
    let last = modList[0];
    modList.forEach((entry) => {
      const t = entry[1].startMin;
      if (t < first[1].startMin) first = entry;
      if (t > last[1].startMin) last = entry;
    });

    const [firstKey, firstMeta] = first;
    const [assignKey, assignMeta] = last;

    // Visitas (inicio) por módulo + actividad
    w.actsByMod.forEach((acts, gKey) => {
      const meta = w.mods.get(gKey);
      if (!meta) return;
      acts.forEach((actividad) => {
        const bKey = `${gKey}||${actividad}`;
        if (!visitMap.has(bKey)) visitMap.set(bKey, new Set());
        visitMap.get(bKey).add(wKey);
      });
    });

    if (firstKey !== assignKey) {
      const supervisor = firstMeta.supervisor || assignMeta.supervisor || "(sin supervisor)";
      if (!isAdminCosechaSupervisor(supervisor)) {
        const moveKey = [
          normExactToken(supervisor),
          firstKey,
          assignKey
        ].join("||");
        if (!moveMap.has(moveKey)) {
          moveMap.set(moveKey, {
            supervisor,
            moduloDesde: firstMeta.modulo,
            moduloHacia: assignMeta.modulo,
            horaHaciaMin: assignMeta.startMin,
            cosechadores: 0
          });
        }
        moveMap.get(moveKey).cosechadores += 1;
      }

      const prevH = leaveHoraByMod.get(firstKey);
      if (
        prevH == null ||
        (Number.isFinite(assignMeta.startMin) && assignMeta.startMin < prevH)
      ) {
        leaveHoraByMod.set(firstKey, assignMeta.startMin);
      }
    }

    const acts = w.actsByMod.get(assignKey) || new Set();
    acts.forEach((actividad) => {
      const bKey = `${assignKey}||${actividad}`;
      if (!buckets.has(bKey)) {
        buckets.set(bKey, {
          fundo: assignMeta.fundo,
          modulo: assignMeta.modulo,
          groupKey: assignKey,
          actividad,
          workers: new Set()
        });
      }
      buckets.get(bKey).workers.add(wKey);
    });
  });

  // Asegurar filas de origen COSECHA aunque neto < inicio
  visitMap.forEach((workers, bKey) => {
    if (buckets.has(bKey)) return;
    const parts = bKey.split("||");
    const actividad = parts[parts.length - 1];
    if (actividad !== "COSECHA") return;
    const groupKey = parts.slice(0, -1).join("||");
    // buscar meta de algún worker
    let meta = null;
    for (const wKey of workers) {
      const w = byWorker.get(wKey);
      meta = w?.mods?.get(groupKey);
      if (meta) break;
    }
    if (!meta) return;
    buckets.set(bKey, {
      fundo: meta.fundo,
      modulo: meta.modulo,
      groupKey,
      actividad,
      workers: new Set() // neto 0 posible
    });
  });

  const items = [...buckets.entries()]
    .map(([bKey, g]) => {
      const inicio = visitMap.get(bKey)?.size || g.workers.size;
      const cantidad = g.workers.size;
      const horaMin = leaveHoraByMod.get(g.groupKey);
      const salieron = inicio > cantidad;
      return {
        fundo: g.fundo,
        modulo: g.modulo,
        actividad: g.actividad,
        cantidad,
        inicio,
        salieron,
        horaDespues: salieron ? minutesToClock(horaMin) : ""
      };
    })
    .filter((it) => it.cantidad > 0 || (it.actividad === "COSECHA" && it.inicio > 0))
    .sort((a, b) => {
      const mm = a.modulo.localeCompare(b.modulo, "es", { numeric: true });
      if (mm) return mm;
      const ff = a.fundo.localeCompare(b.fundo, "es");
      if (ff) return ff;
      const order = [
        "CALIDAD",
        "COSECHA",
        "ESTIBA KIA",
        "SCANER",
        "SUPERVISOR DE COSECHA"
      ];
      const ia = order.indexOf(a.actividad);
      const ib = order.indexOf(b.actividad);
      if (ia !== -1 || ib !== -1) return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
      return a.actividad.localeCompare(b.actividad, "es");
    });

  const movers = [...moveMap.values()]
    .sort((a, b) => {
      const ss = a.supervisor.localeCompare(b.supervisor, "es");
      if (ss) return ss;
      return a.moduloDesde.localeCompare(b.moduloDesde, "es", { numeric: true });
    });

  const sumaCuentas = items.reduce((s, it) => s + it.cantidad, 0);
  const sumaCosecha = items
    .filter((it) => it.actividad === "COSECHA")
    .reduce((s, it) => s + it.cantidad, 0);
  const modulos = new Set(items.map((it) => `${normExactToken(it.fundo)}||${normExactToken(it.modulo)}`));

  return {
    fundoFiltro: fundoFiltro || "",
    items,
    movers,
    totalUnicos: byWorker.size,
    sumaCuentas,
    sumaCosecha,
    modulosCount: modulos.size,
    filasFuente
  };
}

/**
 * COSECHA (exacta) + Fundo del filtro (exacto) o todos.
 * Cada persona cuenta en UN solo módulo: el destino (última hora).
 * Si empezó en M3 y luego M5 → cuenta en M5; M3 muestra el neto en rojo.
 * Avisos: por supervisor (de → a, horas, N cosechadores).
 */
function buildModulosCosechaReport() {
  const rows = state.validated?.rows || [];
  const fundoFiltro = String(readFilters().fundo || "").trim();
  const fundoNorm = normExactToken(fundoFiltro);
  const alcance = rowsAlcanceFundo(fundoFiltro);
  const excluir = keysConSupervisorCosecha(alcance);

  const scoped = rows.filter((row) => {
    if (!isActividadCosechaExact(row.actividad)) return false;
    if (!moduloCosechaLabel(row)) return false;
    if (fundoNorm) return normExactToken(row.fundo) === fundoNorm;
    return true;
  });

  /** worker → módulo destino (última hora de inicio entre módulos) */
  const assigned = new Map();
  /** worker → primer módulo (más temprano) — para avisos de movimiento */
  const firstStep = new Map();
  /** groupKey → visitas */
  const visits = new Map();
  /** worker → Map(groupKey → { modulo, fundo, startMin, supervisor }) */
  const workerMods = new Map();

  scoped.forEach((row) => {
    const wKey = workerUniqueKey(row);
    if (!wKey || esExcluidoDeCosechadores(row, excluir)) return;

    const modulo = moduloCosechaLabel(row);
    if (!modulo) return;
    const documento = String(row.documento || "").trim();
    const trabajador = String(row.trabajador || "").trim();
    const fundo = String(row.fundo || "").trim() || "(sin fundo)";
    const groupKey = fundoFiltro ? modulo : `${fundo}||${modulo}`;
    const supervisor = String(row.supervisor || "").trim() || "(sin supervisor)";
    const startMin =
      row.horaInicioMin != null && Number.isFinite(row.horaInicioMin)
        ? row.horaInicioMin
        : Number.POSITIVE_INFINITY;
    const excelRow = Number(row.excelRow ?? row.rowIndex ?? 0) || 0;

    if (!visits.has(groupKey)) {
      visits.set(groupKey, { fundo, modulo, workers: new Map() });
    }
    const visitBucket = visits.get(groupKey);
    if (!visitBucket.workers.has(wKey)) {
      visitBucket.workers.set(wKey, { documento, trabajador, supervisor });
    }

    if (!workerMods.has(wKey)) workerMods.set(wKey, new Map());
    const wMods = workerMods.get(wKey);
    const prevMod = wMods.get(groupKey);
    if (!prevMod || startMin < prevMod.startMin) {
      wMods.set(groupKey, { fundo, modulo, startMin, supervisor, excelRow, documento, trabajador });
    }
  });

  workerMods.forEach((mods, wKey) => {
    let first = null;
    let last = null;
    mods.forEach((meta, groupKey) => {
      const slot = { ...meta, groupKey };
      if (!first || meta.startMin < first.startMin) first = slot;
      if (!last || meta.startMin > last.startMin) last = slot;
    });
    if (first) firstStep.set(wKey, first);
    if (last) assigned.set(wKey, last);
  });

  const items = [...visits.entries()]
    .map(([groupKey, g]) => {
      let asignados = 0;
      let reubicadosCount = 0; // pisaron aquí pero cuentan en otro (salieron)
      g.workers.forEach((_info, wKey) => {
        const slot = assigned.get(wKey);
        if (!slot) return;
        if (slot.groupKey === groupKey) asignados += 1;
        else reubicadosCount += 1;
      });
      return {
        fundo: g.fundo,
        modulo: g.modulo,
        cantidad: asignados,
        pisaronAqui: g.workers.size,
        reubicadosCount,
        /** Neto reducido porque gente salió a otro módulo */
        salieron: reubicadosCount > 0
      };
    })
    .sort((a, b) => {
      if (!fundoFiltro) {
        const ff = a.fundo.localeCompare(b.fundo, "es");
        if (ff) return ff;
      }
      return a.modulo.localeCompare(b.modulo, "es", { numeric: true });
    });

  /** supervisor + de (primero) → a (destino) */
  const moveMap = new Map();
  assigned.forEach((dest, wKey) => {
    const origin = firstStep.get(wKey);
    if (!origin || origin.groupKey === dest.groupKey) return;
    const supervisor = dest.supervisor || origin.supervisor || "(sin supervisor)";
    const moveKey = [
      normExactToken(supervisor),
      origin.groupKey,
      dest.groupKey
    ].join("||");
    if (!moveMap.has(moveKey)) {
      moveMap.set(moveKey, {
        supervisor,
        moduloDesde: origin.modulo,
        fundoDesde: origin.fundo,
        moduloHacia: dest.modulo,
        fundoHacia: dest.fundo,
        horaDesdeMin: origin.startMin,
        horaHaciaMin: dest.startMin,
        workers: new Set()
      });
    }
    const m = moveMap.get(moveKey);
    m.workers.add(wKey);
    if (Number.isFinite(origin.startMin) && origin.startMin < m.horaDesdeMin) {
      m.horaDesdeMin = origin.startMin;
    }
    if (Number.isFinite(dest.startMin) && dest.startMin < m.horaHaciaMin) {
      m.horaHaciaMin = dest.startMin;
    }
  });

  const movers = [...moveMap.values()]
    .map((m) => ({
      supervisor: m.supervisor,
      moduloDesde: m.moduloDesde,
      fundoDesde: m.fundoDesde,
      moduloHacia: m.moduloHacia,
      fundoHacia: m.fundoHacia,
      horaDesde: minutesToClock(m.horaDesdeMin),
      horaHacia: minutesToClock(m.horaHaciaMin),
      cosechadores: m.workers.size
    }))
    .sort((a, b) => {
      const ss = a.supervisor.localeCompare(b.supervisor, "es");
      if (ss) return ss;
      const md = a.moduloDesde.localeCompare(b.moduloDesde, "es", { numeric: true });
      if (md) return md;
      return a.moduloHacia.localeCompare(b.moduloHacia, "es", { numeric: true });
    });

  const totalUnicos = assigned.size;
  const sumaModulos = items.reduce((s, it) => s + it.cantidad, 0);
  const totalReubicados = movers.reduce((s, m) => s + m.cosechadores, 0);

  const sinDniMap = new Map();
  assigned.forEach((slot, wKey) => {
    if (wKey.startsWith("d:")) return;
    const k = wKey;
    if (sinDniMap.has(k)) return;
    sinDniMap.set(k, {
      trabajador: slot.trabajador || "(sin nombre)",
      documento: "",
      codigo: wKey.startsWith("c:") ? wKey.slice(2) : "",
      modulo: slot.modulo || "",
      fundo: slot.fundo || "",
      supervisor: slot.supervisor || ""
    });
  });
  // También filas COSECHA del alcance sin DNI (por si no entraron a assigned)
  scoped.forEach((row) => {
    const doc = String(row.documento || "").trim();
    if (doc) return;
    const wKey = workerUniqueKey(row);
    if (!wKey || wKey.startsWith("d:") || sinDniMap.has(wKey)) return;
    sinDniMap.set(wKey, {
      trabajador: String(row.trabajador || "").trim() || "(sin nombre)",
      documento: "",
      codigo: String(row.codigoTrabajador || "").replace(/\D/g, ""),
      modulo: moduloCosechaLabel(row) || "",
      fundo: String(row.fundo || "").trim(),
      supervisor: String(row.supervisor || "").trim() || "(sin supervisor)"
    });
  });
  const sinDni = [...sinDniMap.values()].sort((a, b) =>
    a.trabajador.localeCompare(b.trabajador, "es")
  );

  const resumenKeys = collectResumenCosechaKeys(fundoFiltro);
  const modulosKeys = new Set(assigned.keys());
  const extraVsResumen = [];
  modulosKeys.forEach((k) => {
    if (resumenKeys.has(k)) return;
    let info = null;
    assigned.forEach((slot, wKey) => {
      if (wKey === k) info = slot;
    });
    extraVsResumen.push({
      documento: info?.documento || "",
      trabajador: info?.trabajador || k,
      modulo: info?.modulo || "",
      motivo: "En módulos, no en KPI Resumen"
    });
  });
  const faltanEnModulos = [];
  resumenKeys.forEach((k) => {
    if (!modulosKeys.has(k)) {
      faltanEnModulos.push({ documento: k, motivo: "En Resumen, sin módulo válido COSECHA" });
    }
  });

  return {
    fundoFiltro: fundoFiltro || "",
    showFundo: !fundoFiltro,
    items,
    movers,
    totalReubicados,
    totalUnicos,
    sumaModulos,
    filasFuente: scoped.length,
    resumenCosechadores: resumenKeys.size,
    sinDni,
    extraVsResumen,
    faltanEnModulos
  };
}

function minutesToClock(min) {
  if (min == null || !Number.isFinite(min) || min === Number.POSITIVE_INFINITY) return "—";
  const h = Math.floor(min / 60);
  const m = Math.round(min % 60);
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

function renderModulosCosechaTable() {
  const isDetalle = modulosUiState.view === "detalle";
  const head = $("tblModulosCosechaHead");
  const body = $("tblModulosCosechaBody");
  const meta = $("modulosCosechaMeta");
  const count = $("modulosCosechaCount");
  const hint = $("modulosCosechaHint");
  const avisos = $("modulosCosechaAvisos");
  const title = $("modulosCosechaTitle");
  const btnToggle = $("btnToggleModulosDetalle");

  if (btnToggle) {
    btnToggle.textContent = isDetalle ? "Resumen" : "Detalle";
  }
  if (title) {
    title.textContent = isDetalle ? "Detalle por actividad" : "Por módulo";
  }

  if (isDetalle) {
    const report = buildModulosDetalleReport();

    if (hint) hint.innerHTML = `5 actividades · sin duplicar · neto en destino`;
    if (meta) {
      meta.textContent = `${report.fundoFiltro || "todos"} · ${report.modulosCount} mód. · COSECHA ${report.sumaCosecha}`;
    }

    if (avisos) {
      // Lista dinámica: solo aparece si hoy hay movimientos reales
      if (!report.movers?.length) {
        avisos.hidden = true;
        avisos.innerHTML = "";
      } else {
        avisos.hidden = false;
        avisos.innerHTML = `
          <p class="modulos-aviso__title">Supervisores con cambio de módulo</p>
          <ul class="modulos-aviso__list">${report.movers
            .map(
              (m) =>
                `<li><strong>${escapeHtml(m.supervisor)}</strong>: ${escapeHtml(m.moduloDesde)} - ${escapeHtml(m.moduloHacia)}</li>`
            )
            .join("")}</ul>`;
      }
    }

    if (head) {
      head.innerHTML = `<tr><th>Módulo</th><th>Fundo</th><th>Actividad</th><th>Cant.</th></tr>`;
    }

    if (body) {
      if (!report.items.length) {
        body.innerHTML = `<tr><td colspan="4" class="errores-modal__empty">Sin datos.</td></tr>`;
      } else {
        let prevMod = null;
        let prevFundo = null;
        const rowsHtml = report.items
          .map((it) => {
            const showMod = it.modulo !== prevMod;
            const showFundo = showMod || it.fundo !== prevFundo;
            prevMod = it.modulo;
            prevFundo = it.fundo;
            const rowClass = showMod ? " class=\"modulos-detalle__group\"" : "";
            let cantHtml = String(it.cantidad);
            // Solo si hoy hubo salida real de este módulo (números/hora del Excel)
            if (it.actividad === "COSECHA" && it.salieron && it.inicio > it.cantidad) {
              const hora =
                it.horaDespues && it.horaDespues !== "—" ? it.horaDespues : "";
              const nota = hora
                ? ` <small>(quedaron después ${escapeHtml(hora)})</small>`
                : "";
              cantHtml = `<span class="modulos-num--rojo">${it.inicio} → ${it.cantidad}${nota}</span>`;
            }
            return `<tr${rowClass}>
              <td class="modulos-detalle__mod">${showMod ? escapeHtml(it.modulo) : ""}</td>
              <td>${showFundo ? escapeHtml(it.fundo) : ""}</td>
              <td>${escapeHtml(it.actividad)}</td>
              <td>${cantHtml}</td>
            </tr>`;
          })
          .join("");
        body.innerHTML =
          rowsHtml +
          `<tr class="modulos-table__total"><td colspan="3">TOTAL COSECHA</td><td>${report.sumaCosecha}</td></tr>`;
      }
    }

    if (count) count.textContent = `COSECHA ${report.sumaCosecha}`;
    return report;
  }

  const report = buildModulosCosechaReport();

  if (hint) {
    hint.innerHTML = `COSECHA · únicos (DNI/nombre) · sin duplicados · <b style="color:#b91c1c">rojo</b> = neto tras salida`;
  }
  if (meta) {
    const f = report.fundoFiltro || "todos";
    const r = report.resumenCosechadores;
    meta.textContent =
      r != null
        ? `${f} · ${report.items.length} mód. · ${report.totalUnicos} únicos (Resumen ${r})`
        : `${f} · ${report.items.length} mód. · ${report.totalUnicos} únicos`;
  }

  if (head) {
    head.innerHTML = report.showFundo
      ? `<tr><th>Fundo</th><th>Módulo</th><th>Neto</th><th>Inicio</th><th>Salieron</th></tr>`
      : `<tr><th>Módulo</th><th>Neto</th><th>Inicio</th><th>Salieron</th></tr>`;
  }

  if (body) {
    if (!report.items.length) {
      const cols = report.showFundo ? 5 : 4;
      body.innerHTML = `<tr><td colspan="${cols}" class="errores-modal__empty">Sin datos.</td></tr>`;
    } else {
      const rowsHtml = report.items
        .map((it) => {
          const warnClass = it.salieron ? " class=\"modulos-table__warn\"" : "";
          const cantHtml = it.salieron
            ? `<span class="modulos-num--rojo" title="Inicio ${it.pisaronAqui}">${it.cantidad}</span>`
            : String(it.cantidad);
          const reuCell =
            it.reubicadosCount > 0
              ? `<span class="modulos-reu-badge">${it.reubicadosCount}</span>`
              : "0";
          return report.showFundo
            ? `<tr${warnClass}><td>${escapeHtml(it.fundo)}</td><td>${escapeHtml(it.modulo)}</td><td>${cantHtml}</td><td>${it.pisaronAqui}</td><td>${reuCell}</td></tr>`
            : `<tr${warnClass}><td>${escapeHtml(it.modulo)}</td><td>${cantHtml}</td><td>${it.pisaronAqui}</td><td>${reuCell}</td></tr>`;
        })
        .join("");

      const totalRow = report.showFundo
        ? `<tr class="modulos-table__total"><td colspan="2">TOTAL</td><td>${report.totalUnicos}</td><td>—</td><td>${report.totalReubicados || 0}</td></tr>`
        : `<tr class="modulos-table__total"><td>TOTAL</td><td>${report.totalUnicos}</td><td>—</td><td>${report.totalReubicados || 0}</td></tr>`;

      body.innerHTML = rowsHtml + totalRow;
    }
  }

  if (avisos) {
    const blocks = [];
    if (report.sinDni?.length) {
      blocks.push(`
        <p class="modulos-aviso__title">Sin DNI (${report.sinDni.length})</p>
        <ul class="modulos-aviso__list">${report.sinDni
          .map((p) => {
            const extra = [
              p.codigo ? `cód. ${p.codigo}` : "",
              p.modulo || "",
              p.supervisor || ""
            ]
              .filter(Boolean)
              .join(" · ");
            return `<li><strong>${escapeHtml(p.trabajador)}</strong>${extra ? ` — ${escapeHtml(extra)}` : ""}</li>`;
          })
          .join("")}</ul>`);
    }
    if (report.extraVsResumen?.length) {
      blocks.push(`
        <p class="modulos-aviso__title">Solo en módulos (no en Resumen)</p>
        <ul class="modulos-aviso__list">${report.extraVsResumen
          .map(
            (p) =>
              `<li><strong>${escapeHtml(p.trabajador || p.documento)}</strong> (${escapeHtml(p.documento)})</li>`
          )
          .join("")}</ul>`);
    }
    if (report.movers?.length) {
      blocks.push(`
        <p class="modulos-aviso__title">Movimientos</p>
        <ul class="modulos-aviso__list">${report.movers
          .map(
            (m) =>
              `<li><strong>${escapeHtml(m.supervisor)}</strong>: ${escapeHtml(m.moduloDesde)} ${escapeHtml(m.horaDesde)} → ${escapeHtml(m.moduloHacia)} ${escapeHtml(m.horaHacia)} · ${m.cosechadores}</li>`
          )
          .join("")}</ul>`);
    }
    if (!blocks.length) {
      avisos.hidden = true;
      avisos.innerHTML = "";
    } else {
      avisos.hidden = false;
      avisos.innerHTML = blocks.join("");
    }
  }

  if (count) count.textContent = `${report.totalUnicos} únicos`;
  return report;
}

function toggleModulosDetalleView() {
  modulosUiState.view = modulosUiState.view === "detalle" ? "resumen" : "detalle";
  renderModulosCosechaTable();
}

function openModulosCosechaModal() {
  if (!state.validated) {
    window.alert("Primero sube un Excel de tareo.");
    return;
  }
  const modal = $("modalModulosCosecha");
  if (!modal) return;
  modulosUiState.view = "resumen";
  renderModulosCosechaTable();
  modal.hidden = false;
}

function closeModulosCosechaModal() {
  const modal = $("modalModulosCosecha");
  if (modal) modal.hidden = true;
}

function renderModulosCosechaIfOpen() {
  const modal = $("modalModulosCosecha");
  if (!modal || modal.hidden) return;
  renderModulosCosechaTable();
}

function classifyActividadLote(actividad) {
  const a = normExactToken(actividad);
  if (a === "COSECHA") return "cosecha";
  if (a === "SCANER" || a === "SCANNER" || a === "ESCANER") return "scaner";
  if (a.includes("CALIDAD")) return "calidad";
  if (a === "SUPERVISOR DE COSECHA" || a === "SUPERVISOR DE ACOPIO") return "supervisor";
  return "otro";
}

const LOTES_LICAPA_URL = "data/lotes-licapa.json";
let lotesLicapaCatalog = null; // { byLote: Map, byCod: Map, list: [] }

function loteNumKey(value) {
  const digits = String(value ?? "").replace(/\D/g, "");
  if (!digits) return "";
  // evita claves vacías / leading zeros → "003" = "3"
  return String(Number(digits));
}

async function loadLotesLicapaCatalog() {
  if (lotesLicapaCatalog) return lotesLicapaCatalog;
  try {
    const res = await fetch(LOTES_LICAPA_URL, { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    const list = Array.isArray(data) ? data : Array.isArray(data?.lotes) ? data.lotes : [];
    const byLote = new Map();
    const byCod = new Map();
    list.forEach((item) => {
      const lote = String(item?.lote ?? "").trim();
      const modulo = String(item?.modulo ?? "").trim();
      const turno = String(item?.turno ?? "").trim();
      const codLote = String(item?.codLote ?? item?.cod ?? "").trim();
      const meta = {
        lote,
        modulo,
        turno,
        codLote,
        etapa: String(item?.etapa ?? "").trim(),
        variedad: String(item?.variedad ?? "").trim()
      };
      const num = loteNumKey(lote);
      if (num) byLote.set(num, meta);
      if (codLote) byCod.set(normExactToken(codLote), meta);
    });
    lotesLicapaCatalog = { list, byLote, byCod };
  } catch (err) {
    console.warn("[lotes-licapa] No se pudo cargar:", err);
    lotesLicapaCatalog = { list: [], byLote: new Map(), byCod: new Map() };
  }
  return lotesLicapaCatalog;
}

/** Extrae nº de lote desde col P / texto (189, LOTE 189, LT189, M5T2LT189…). */
function extractLoteNumberFromText(text) {
  const raw = String(text ?? "").trim();
  if (!raw) return "";
  const upper = raw.toUpperCase();
  const patterns = [
    /\bLT\s*0*(\d+)\b/i,
    /\bLOTE\s*0*(\d+)\b/i,
    /\bL\s*0*(\d+)\b/i,
    /^0*(\d+)$/
  ];
  for (const re of patterns) {
    const m = upper.match(re);
    if (m?.[1]) return String(Number(m[1]));
  }
  const onlyDigits = raw.replace(/\D/g, "");
  if (onlyDigits && onlyDigits.length <= 4) return String(Number(onlyDigits));
  return "";
}

function formatLoteLabel(loteNum) {
  const n = String(loteNum || "").trim();
  if (!n) return "";
  return `LOTE ${n}`;
}

/** Catálogo M5/T2 → "MODULO 5 - TURNO 2" */
function formatModuloTurnoDetalle(modulo, turno) {
  const modRaw = String(modulo || "").trim();
  const turRaw = String(turno || "").trim();
  if (!modRaw && !turRaw) return "—";
  const modNum = (modRaw.match(/(\d+)/) || [])[1] || modRaw.replace(/^m/i, "");
  const turNum = (turRaw.match(/(\d+)/) || [])[1] || turRaw.replace(/^turno\s*/i, "");
  if (modNum && turNum) return `MODULO ${modNum} - TURNO ${turNum}`;
  if (modNum) return `MODULO ${modNum}`;
  if (turNum) return `TURNO ${turNum}`;
  return "—";
}

/**
 * Lote = columna P del Excel. Módulo/Turno = lotes-licapa.json.
 * CECO (QBE01-…) NO se usa como lote.
 */
function resolveLoteLicapaMeta(row, catalog) {
  const loteExcel = String(row?.lote || "").trim();
  const moduloExcel = String(row?.modulo || "").trim();
  const ceco = String(row?.ceco || "").trim();

  const cat = catalog || lotesLicapaCatalog;
  const loteNum =
    extractLoteNumberFromText(loteExcel) ||
    extractLoteNumberFromText(moduloExcel) ||
    "";

  if (!loteNum) {
    // Sin lote en col P: no inventar con CECO
    return {
      lote: loteExcel || "(sin lote)",
      loteNum: "",
      modulo: "",
      turno: "",
      moduloTurno: "—",
      fromCatalog: false
    };
  }

  let meta = cat?.byLote?.get(loteNum) || null;
  if (!meta && loteExcel) {
    meta = cat?.byCod?.get(normExactToken(loteExcel)) || null;
  }
  if (!meta && ceco) {
    meta = cat?.byCod?.get(normExactToken(ceco)) || null;
  }

  if (!meta) {
    return {
      lote: formatLoteLabel(loteNum),
      loteNum,
      modulo: "",
      turno: "",
      moduloTurno: "—",
      fromCatalog: false
    };
  }

  const hasModulo = Boolean(meta.modulo);
  const hasTurno = Boolean(meta.turno);
  return {
    lote: formatLoteLabel(meta.lote || loteNum),
    loteNum: loteNumKey(meta.lote) || loteNum,
    modulo: hasModulo ? meta.modulo : "",
    turno: hasTurno ? meta.turno : "",
    moduloTurno: formatModuloTurnoDetalle(meta.modulo, meta.turno),
    codLote: meta.codLote || "",
    fromCatalog: true,
    soloLote: !hasModulo && !hasTurno
  };
}

let lotesHoyUiState = { search: "", report: null };

/**
 * Una fila = supervisor + fundo + lote (col P).
 * Módulo/Turno = lotes-licapa → "MODULO 5 - TURNO 2".
 */
function buildLotesHoyReport() {
  const rows = state.validated?.rows || [];
  const fundoFiltro = String(readFilters().fundo || "").trim();
  const fundoNorm = normExactToken(fundoFiltro);
  const catalog = lotesLicapaCatalog;

  const scoped = rows.filter((row) => {
    if (!row.esCostoCosecha) return false;
    if (fundoNorm) return normExactToken(row.fundo) === fundoNorm;
    return true;
  });

  const byKey = new Map();
  const fechas = new Set();

  scoped.forEach((row) => {
    const kind = classifyActividadLote(row.actividad);
    if (kind === "otro") return;

    const supervisor = String(row.supervisor || "").trim() || "(sin supervisor)";
    const fundo = String(row.fundo || "").trim() || "(sin fundo)";
    const loteMeta = resolveLoteLicapaMeta(row, catalog);
    const lote = loteMeta.lote;
    const wKey = workerUniqueKey(row);
    const groupKey = `${normExactToken(fundo)}||${normExactToken(supervisor)}||${normExactToken(loteMeta.loteNum || lote)}`;

    if (row.fecha) fechas.add(row.fecha);

    if (!byKey.has(groupKey)) {
      byKey.set(groupKey, {
        fundo,
        supervisor,
        lote,
        loteNum: loteMeta.loteNum || "",
        modulo: loteMeta.modulo || "",
        turno: loteMeta.turno || "",
        moduloTurno: loteMeta.moduloTurno || "—",
        fromCatalog: Boolean(loteMeta.fromCatalog),
        cosecha: new Set(),
        scaner: new Set(),
        calidad: new Set(),
        supervisorAct: new Set(),
        total: new Set()
      });
    }
    const g = byKey.get(groupKey);
    if (loteMeta.fromCatalog) {
      g.fromCatalog = true;
      g.lote = loteMeta.lote;
      if (loteMeta.loteNum) g.loteNum = loteMeta.loteNum;
      if (loteMeta.modulo) g.modulo = loteMeta.modulo;
      if (loteMeta.turno) g.turno = loteMeta.turno;
      g.moduloTurno = loteMeta.moduloTurno || formatModuloTurnoDetalle(g.modulo, g.turno);
    }
    if (!wKey) return;
    g.total.add(wKey);
    if (kind === "cosecha") g.cosecha.add(wKey);
    else if (kind === "scaner") g.scaner.add(wKey);
    else if (kind === "calidad") g.calidad.add(wKey);
    else if (kind === "supervisor") g.supervisorAct.add(wKey);
  });

  const items = [...byKey.values()]
    .map((g) => ({
      fundo: g.fundo,
      supervisor: g.supervisor,
      lote: g.lote,
      loteNum: g.loteNum,
      moduloTurno: g.moduloTurno || formatModuloTurnoDetalle(g.modulo, g.turno),
      fromCatalog: g.fromCatalog,
      isAdminCosecha: isAdminCosechaSupervisor(g.supervisor),
      cosecha: g.cosecha.size,
      scaner: g.scaner.size,
      calidad: g.calidad.size,
      supervisorAct: g.supervisorAct.size,
      total: g.total.size
    }))
    .filter(
      (it) =>
        it.cosecha > 0 ||
        it.scaner > 0 ||
        it.calidad > 0 ||
        it.supervisorAct > 0 ||
        it.total > 0
    )
    .sort((a, b) => {
      const aa = a.isAdminCosecha ? 1 : 0;
      const bb = b.isAdminCosecha ? 1 : 0;
      if (aa !== bb) return aa - bb;
      const ss = a.supervisor.localeCompare(b.supervisor, "es");
      if (ss) return ss;
      if (!fundoFiltro) {
        const ff = a.fundo.localeCompare(b.fundo, "es");
        if (ff) return ff;
      }
      const na = Number(a.loteNum) || 0;
      const nb = Number(b.loteNum) || 0;
      if (na && nb && na !== nb) return na - nb;
      return a.lote.localeCompare(b.lote, "es", { numeric: true });
    });

  const totCosecha = items.reduce((s, it) => s + it.cosecha, 0);
  const totScaner = items.reduce((s, it) => s + it.scaner, 0);
  const totCalidad = items.reduce((s, it) => s + it.calidad, 0);
  const totSupervisor = items.reduce((s, it) => s + it.supervisorAct, 0);
  const supervisoresUnicos = new Set(items.map((it) => it.supervisor)).size;
  const lotesUnicos = new Set(items.map((it) => `${it.fundo}||${it.lote}`)).size;
  const conCatalogo = items.filter((it) => it.fromCatalog).length;

  return {
    fundoFiltro: fundoFiltro || "",
    showFundo: !fundoFiltro,
    items,
    fechas: [...fechas].sort(),
    totCosecha,
    totScaner,
    totCalidad,
    totSupervisor,
    supervisoresUnicos,
    lotesUnicos,
    conCatalogo,
    filasFuente: scoped.length
  };
}

function getLotesHoyVisibleItems(report) {
  const q = String(lotesHoyUiState.search || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
  const items = report?.items || [];
  if (!q) return items;
  return items.filter((it) => {
    const blob = `${it.supervisor} ${it.fundo} ${it.lote} ${it.moduloTurno || ""}`
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
    return blob.includes(q);
  });
}

function renderLotesHoyCards() {
  const report = buildLotesHoyReport();
  lotesHoyUiState.report = report;
  const body = $("lotesHoyBody");
  const empty = $("lotesHoyEmpty");
  const meta = $("lotesHoyMeta");
  const count = $("lotesHoyCount");
  const footerMeta = $("lotesHoyFooterMeta");
  const hint = $("lotesHoyHint");
  const wrap = $("lotesHoyGrid");
  const kpiSup = $("lotesKpiSup");
  const kpiLotes = $("lotesKpiLotes");
  const kpiCosecha = $("lotesKpiCosecha");
  const kpiScaner = $("lotesKpiScaner");
  const kpiCalidad = $("lotesKpiCalidad");

  if (hint) {
    hint.innerHTML = report.fundoFiltro
      ? `Filtro: fundo <b>${escapeHtml(report.fundoFiltro)}</b>. Lote = col. P del Excel · detalle = <b>MODULO X - TURNO Y</b> (lotes-licapa).`
      : `Lote = <b>columna P</b> del Excel. Detalle = <b>MODULO X - TURNO Y</b> desde lotes-licapa. Si no hay lote, no se inventa con CECO.`;
  }

  if (meta) {
    const fechaTxt =
      report.fechas.length === 1
        ? `Fecha ${report.fechas[0]}`
        : report.fechas.length
          ? `${report.fechas.length} fechas`
          : "Sin fecha";
    meta.textContent = fechaTxt;
  }

  if (kpiSup) kpiSup.textContent = String(report.supervisoresUnicos);
  if (kpiLotes) kpiLotes.textContent = String(report.lotesUnicos);
  if (kpiCosecha) kpiCosecha.textContent = String(report.totCosecha);
  if (kpiScaner) kpiScaner.textContent = String(report.totScaner);
  if (kpiCalidad) kpiCalidad.textContent = String(report.totSupervisor || 0);

  const visible = getLotesHoyVisibleItems(report);

  if (!report.items.length) {
    if (body) body.innerHTML = "";
    if (wrap) wrap.hidden = true;
    if (empty) {
      empty.hidden = false;
      empty.textContent = "Sin lotes con personal en este alcance.";
    }
  } else if (!visible.length) {
    if (body) body.innerHTML = "";
    if (wrap) wrap.hidden = true;
    if (empty) {
      empty.hidden = false;
      empty.textContent = "Sin coincidencias en la búsqueda.";
    }
  } else {
    if (empty) empty.hidden = true;
    if (wrap) wrap.hidden = false;
    if (body) {
      body.innerHTML = visible
        .map((it) => {
          const mtClass =
            !it.moduloTurno || it.moduloTurno === "—"
              ? "lotes-modturno is-muted"
              : "lotes-modturno";
          const loteTitle = it.fromCatalog
            ? `${it.lote} · ${it.moduloTurno}`
            : `${it.lote} (sin catálogo)`;
          const rowClass = it.isAdminCosecha ? "is-row-admin-cosecha" : "";
          const supHtml = it.isAdminCosecha
            ? `${escapeHtml(it.supervisor)} <span class="resumen-admin-tag">Admin</span>`
            : escapeHtml(it.supervisor);
          return `<tr class="${rowClass}">
            <td class="lotes-table__sup">${supHtml}</td>
            <td class="lotes-table__fundo">${escapeHtml(it.fundo)}</td>
            <td class="lotes-table__lote" title="${escapeHtml(loteTitle)}">${escapeHtml(it.lote)}</td>
            <td><span class="${mtClass}">${escapeHtml(it.moduloTurno || "—")}</span></td>
            <td class="lotes-num lotes-num--cosecha">${it.cosecha}</td>
            <td class="lotes-num lotes-num--scaner">${it.scaner}</td>
            <td class="lotes-num lotes-num--sup">${it.supervisorAct}</td>
            <td class="lotes-num lotes-num--total">${it.total}</td>
          </tr>`;
        })
        .join("");
    }
  }

  if (count) {
    count.textContent =
      visible.length === 1 ? "1 fila visible" : `${visible.length} filas visibles`;
  }
  if (footerMeta) {
    footerMeta.textContent = `Totales: Cosecha ${report.totCosecha} · Escanear ${report.totScaner} · Supervisor ${report.totSupervisor || 0} · ${report.supervisoresUnicos} supervisores · ${report.lotesUnicos} lotes`;
  }

  return report;
}

async function openLotesHoyModal() {
  if (!state.validated) {
    window.alert("Primero sube un Excel de tareo.");
    return;
  }
  const modal = $("modalLotesHoy");
  if (!modal) return;
  await loadLotesLicapaCatalog();
  lotesHoyUiState.search = "";
  const search = $("lotesHoySearch");
  if (search) search.value = "";
  renderLotesHoyCards();
  modal.hidden = false;
}

function closeLotesHoyModal() {
  const modal = $("modalLotesHoy");
  if (modal) modal.hidden = true;
}

function renderLotesHoyIfOpen() {
  const modal = $("modalLotesHoy");
  if (!modal || modal.hidden) return;
  loadLotesLicapaCatalog().then(() => renderLotesHoyCards());
}

function exportLotesHoyExcel() {
  if (!state.validated) {
    window.alert("Primero sube un Excel de tareo.");
    return;
  }
  const report = lotesHoyUiState.report || buildLotesHoyReport();
  const items = getLotesHoyVisibleItems(report);
  const stamp = new Date().toISOString().slice(0, 10);
  const fundoTag = report.fundoFiltro
    ? report.fundoFiltro.replace(/\s+/g, "-")
    : "TODOS";

  const headers = [
    "Supervisor",
    "Fundo",
    "Lote",
    "Módulo - Turno",
    "Cosecha",
    "Escanear",
    "Supervisor act.",
    "Total personas"
  ];

  const rows = items.map((it) => ({
    cells: [
      { value: it.supervisor },
      { value: it.fundo },
      { value: it.lote },
      { value: it.moduloTurno },
      { value: it.cosecha },
      { value: it.scaner },
      { value: it.supervisorAct },
      { value: it.total }
    ]
  }));

  rows.push({
    cells: [
      { value: "TOTALES" },
      { value: `${report.supervisoresUnicos} supervisores` },
      { value: `${report.lotesUnicos} lotes` },
      { value: "" },
      { value: report.totCosecha },
      { value: report.totScaner },
      { value: report.totSupervisor || 0 },
      { value: "" }
    ]
  });

  try {
    downloadXlsxExcel({
      filename: `lotes-hoy-${fundoTag}-${stamp}.xlsx`,
      sheetName: "Lotes hoy",
      headers,
      rows
    });
  } catch (err) {
    window.alert(err?.message || "No se pudo exportar Excel.");
  }
}

function exportModulosCosechaExcel() {
  if (!state.validated) {
    window.alert("Primero sube un Excel de tareo.");
    return;
  }

  const stamp = new Date().toISOString().slice(0, 10);

  if (modulosUiState.view === "detalle") {
    const report = buildModulosDetalleReport();
    const fundoTag = report.fundoFiltro
      ? report.fundoFiltro.replace(/\s+/g, "-")
      : "TODOS";
    const headers = ["Módulo", "Fundo", "Actividad", "Cant."];
    const rows = report.items.map((it) => ({
      cells: [
        { value: it.modulo },
        { value: it.fundo },
        { value: it.actividad },
        { value: it.cantidad }
      ]
    }));
    rows.push({
      cells: [
        { value: "TOTAL COSECHA" },
        { value: "" },
        { value: "" },
        { value: report.sumaCosecha }
      ]
    });

    try {
      downloadXlsxExcel({
        filename: `cosecha-modulos-detalle-${fundoTag}-${stamp}.xlsx`,
        sheetName: "Detalle módulos",
        headers,
        rows
      });
    } catch (err) {
      window.alert(err?.message || "No se pudo exportar Excel.");
    }
    return;
  }

  const report = buildModulosCosechaReport();
  const fundoTag = report.fundoFiltro
    ? report.fundoFiltro.replace(/\s+/g, "-")
    : "TODOS";

  const headers = report.showFundo
    ? ["Fundo", "Módulo", "Neto", "Inicio", "Salieron"]
    : ["Módulo", "Neto", "Inicio", "Salieron"];

  const rows = report.items.map((it) => ({
    cells: report.showFundo
      ? [
          { value: it.fundo },
          { value: it.modulo },
          { value: it.cantidad },
          { value: it.pisaronAqui },
          { value: it.reubicadosCount }
        ]
      : [
          { value: it.modulo },
          { value: it.cantidad },
          { value: it.pisaronAqui },
          { value: it.reubicadosCount }
        ]
  }));

  rows.push({
    cells: report.showFundo
      ? [
          { value: "TOTAL" },
          { value: "" },
          { value: report.totalUnicos },
          { value: "" },
          { value: report.totalReubicados || 0 }
        ]
      : [
          { value: "TOTAL" },
          { value: report.totalUnicos },
          { value: "" },
          { value: report.totalReubicados || 0 }
        ]
  });

  if (report.movers.length) {
    rows.push({ cells: [{ value: "" }] });
    rows.push({
      cells: [{ value: "AVISO: movimientos de supervisor" }]
    });
    rows.push({
      cells: [
        { value: "Supervisor" },
        { value: "De módulo" },
        { value: "Hora desde" },
        { value: "A módulo" },
        { value: "Hora hacia" },
        { value: "Cosechadores" }
      ]
    });
    report.movers.forEach((m) => {
      const desde = report.showFundo
        ? `${m.fundoDesde} · ${m.moduloDesde}`
        : m.moduloDesde;
      const hacia = report.showFundo
        ? `${m.fundoHacia} · ${m.moduloHacia}`
        : m.moduloHacia;
      rows.push({
        cells: [
          { value: m.supervisor },
          { value: desde },
          { value: m.horaDesde },
          { value: hacia },
          { value: m.horaHacia },
          { value: m.cosechadores }
        ]
      });
    });
  }

  try {
    downloadXlsxExcel({
      filename: `cosecha-modulos-${fundoTag}-${stamp}.xlsx`,
      sheetName: "COSECHA módulos",
      headers,
      rows
    });
  } catch (err) {
    window.alert(err?.message || "No se pudo exportar Excel.");
  }
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

function motivoErrorPersona(row) {
  const parts = [];

  // Identidad primero: DNI sin nombre / documento vacío
  if (row.tipTrabajador) parts.push(row.tipTrabajador);
  if (row.tipDocumento) parts.push(row.tipDocumento);
  if (row.tipCeco) parts.push(row.tipCeco);
  if (row.tipActividad) parts.push(row.tipActividad);

  // Preferir el tip de horario (ya trae "Puso: …")
  if (row.tipHoraFin) parts.push(row.tipHoraFin);
  else if (row.tipHoraInicio) parts.push(row.tipHoraInicio);

  const inis = row.horasInicioDetalle?.length
    ? row.horasInicioDetalle
    : row.horaInicioTexto
      ? [row.horaInicioTexto]
      : [];
  const fins = row.horasFinDetalle?.length
    ? row.horasFinDetalle
    : row.horaFinTexto
      ? [row.horaFinTexto]
      : [];
  const bloques = [];
  const n = Math.max(inis.length, fins.length);
  for (let i = 0; i < n; i += 1) {
    bloques.push(`${inis[i] || "—"}→${fins[i] || "—"}`);
  }
  const suma =
    row.sumaHorasPago != null || row.totalDia != null || row.horas != null
      ? Math.round(Number(row.sumaHorasPago ?? row.totalDia ?? row.horas) * 1e3) / 1e3
      : null;

  if (row.tipHoras) {
    if (bloques.length) parts.push(`Puso: ${bloques.join(" · ")}. ${row.tipHoras}`);
    else parts.push(row.tipHoras);
  } else if (row.tipDuplicado) {
    parts.push(row.tipDuplicado);
  } else if (!parts.length && bloques.length) {
    parts.push(`Puso: ${bloques.join(" · ")}${suma != null ? ` (${suma} h)` : ""}`);
  }

  if (parts.length) return parts.join(" · ");
  return "Error de horas / horario";
}

function collectPersonasConError() {
  if (!state.validated?.rows?.length) return [];
  const dayRows = collapseToDayRows(state.validated.rows).filter((r) => r.status === "rojo");
  return dayRows
    .map((r) => ({
      documento: r.documento || "",
      trabajador: r.trabajador || "",
      supervisor: r.supervisor || "",
      fecha: r.fecha || "",
      fundo: r.fundo || "",
      motivo: motivoErrorPersona(r),
      inicios: (r.horasInicioDetalle || []).join(" / "),
      fines: (r.horasFinDetalle || []).join(" / ")
    }))
    .sort((a, b) => {
      const sa = String(a.supervisor).localeCompare(String(b.supervisor), "es");
      if (sa) return sa;
      return String(a.trabajador).localeCompare(String(b.trabajador), "es");
    });
}

function renderErroresPersonasList(query = "") {
  const body = $("erroresPersonasBody");
  const empty = $("erroresPersonasEmpty");
  const countEl = $("erroresPersonasCount");
  if (!body) return;
  const q = String(query || "")
    .trim()
    .toLowerCase();
  const all = collectPersonasConError();
  const rows = q
    ? all.filter((p) =>
        `${p.documento} ${p.trabajador} ${p.supervisor} ${p.fecha} ${p.motivo}`
          .toLowerCase()
          .includes(q)
      )
    : all;
  if (countEl) {
    countEl.textContent = rows.length === 1 ? "1 persona" : `${rows.length} personas`;
  }
  if (!rows.length) {
    body.innerHTML = "";
    if (empty) empty.hidden = false;
    return;
  }
  if (empty) empty.hidden = true;
  body.innerHTML = rows
    .map(
      (p) => `<tr data-err-dni="${escapeHtml(p.documento)}" data-err-fecha="${escapeHtml(p.fecha)}">
      <td class="errores-modal__dni">${escapeHtml(p.documento || "—")}</td>
      <td>${escapeHtml(isNombreTrabajadorVacio(p.trabajador) ? "(sin nombre)" : p.trabajador || "—")}</td>
      <td>${escapeHtml(p.supervisor || "—")}</td>
      <td>${escapeHtml(p.fecha || "—")}</td>
      <td class="errores-modal__motivo" title="${escapeHtml(p.motivo)}">${escapeHtml(p.motivo)}</td>
    </tr>`
    )
    .join("");
}

function collectSupervisoresConError() {
  const personas = collectPersonasConError();
  const map = new Map();
  personas.forEach((p) => {
    const key = String(p.supervisor || "").trim() || "(sin supervisor)";
    if (!map.has(key)) {
      map.set(key, {
        supervisor: key,
        errores: 0,
        personas: new Set(),
        motivos: []
      });
    }
    const g = map.get(key);
    g.errores += 1;
    if (p.documento) g.personas.add(String(p.documento));
    else if (p.trabajador) g.personas.add(String(p.trabajador));
    if (p.motivo && g.motivos.length < 3 && !g.motivos.includes(p.motivo)) {
      g.motivos.push(p.motivo);
    }
  });
  return [...map.values()]
    .map((g) => ({
      supervisor: g.supervisor,
      errores: g.errores,
      personas: g.personas.size,
      motivos: g.motivos.join(" · ")
    }))
    .sort((a, b) => b.errores - a.errores || a.supervisor.localeCompare(b.supervisor, "es"));
}

function renderSupervisoresErroresList(query = "") {
  const body = $("supErroresBody");
  const empty = $("supErroresEmpty");
  const countEl = $("supErroresCount");
  if (!body) return;
  const q = String(query || "")
    .trim()
    .toLowerCase();
  const all = collectSupervisoresConError();
  const rows = q
    ? all.filter((s) => `${s.supervisor} ${s.motivos}`.toLowerCase().includes(q))
    : all;
  if (countEl) {
    countEl.textContent =
      rows.length === 1 ? "1 supervisor" : `${rows.length} supervisores`;
  }
  if (!rows.length) {
    body.innerHTML = "";
    if (empty) empty.hidden = false;
    return;
  }
  if (empty) empty.hidden = true;
  body.innerHTML = rows
    .map(
      (s) => `<tr data-sup-err="${escapeHtml(s.supervisor)}">
      <td class="errores-modal__dni">${escapeHtml(s.supervisor)}</td>
      <td><strong>${s.errores}</strong></td>
      <td>${s.personas}</td>
      <td class="errores-modal__motivo" title="${escapeHtml(s.motivos)}">${escapeHtml(s.motivos || "—")}</td>
    </tr>`
    )
    .join("");
}

function openSupervisoresErroresModal() {
  const modal = $("modalSupervisoresErrores");
  if (!modal) return;
  const search = $("supErroresSearch");
  if (search) search.value = "";
  renderSupervisoresErroresList("");
  modal.hidden = false;
}

function closeSupervisoresErroresModal() {
  const modal = $("modalSupervisoresErrores");
  if (modal) modal.hidden = true;
}

function focusSupervisorFromError(supervisor) {
  closeSupervisoresErroresModal();
  enterErrorFocusMode();
  const sel = $("fltSupervisor");
  if (sel && supervisor && supervisor !== "(sin supervisor)") {
    const has = [...sel.options].some((o) => o.value === supervisor);
    if (has) {
      sel.disabled = false;
      sel.value = supervisor;
    }
  }
  refreshView();
}

function openErroresPersonasModal() {
  const modal = $("modalErroresPersonas");
  if (!modal) return;
  const search = $("erroresPersonasSearch");
  if (search) search.value = "";
  renderErroresPersonasList("");
  modal.hidden = false;
}

function closeErroresPersonasModal() {
  const modal = $("modalErroresPersonas");
  if (modal) modal.hidden = true;
}

function focusPersonaFromError(dni, fecha) {
  closeErroresPersonasModal();
  enterErrorFocusMode();
  const search = $("fltSearch");
  if (search && dni) {
    search.value = dni;
  }
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

function stackExportText(times) {
  const list = Array.isArray(times) ? times.filter(Boolean) : [];
  return list.join(" | ");
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

function cellExportValue(cell) {
  if (cell == null) return "";
  if (typeof cell !== "object" || Array.isArray(cell)) return cell;
  if (cell.html != null) {
    return String(cell.html)
      .replace(/<br\s*\/?>/gi, " | ")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .trim();
  }
  return cell.value ?? "";
}

/** Exporta .xlsx real (SheetJS) — Excel lo abre sin aviso de extensión. */
function downloadXlsxExcel({ filename, sheetName, headers, rows }) {
  if (typeof XLSX === "undefined" || !XLSX?.utils) {
    throw new Error("No se cargó el exportador Excel (XLSX).");
  }

  const safeName = String(filename || "export.xlsx").replace(/\.xls$/i, ".xlsx");
  const safeSheet = String(sheetName || "Hoja1").slice(0, 31);
  const aoa = [
    headers,
    ...(rows || []).map((row) => (row.cells || []).map((cell) => cellExportValue(cell)))
  ];
  if (aoa.length === 1) {
    aoa.push(headers.map(() => ""));
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = headers.map((h, i) => {
    let max = String(h || "").length;
    aoa.forEach((r) => {
      const len = String(r[i] ?? "").length;
      if (len > max) max = len;
    });
    return { wch: Math.min(Math.max(max + 2, 10), 40) };
  });

  /* Documentos como texto para no perder ceros a la izquierda */
  headers.forEach((h, col) => {
    if (!/documento|dni/i.test(String(h))) return;
    for (let r = 1; r < aoa.length; r += 1) {
      const addr = XLSX.utils.encode_cell({ r, c: col });
      const cell = ws[addr];
      if (!cell) continue;
      cell.t = "s";
      cell.v = String(cell.v ?? "");
      cell.z = "@";
    }
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, safeSheet);
  XLSX.writeFile(wb, safeName);
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
          value: !isNombreTrabajadorVacio(row.trabajador)
            ? row.trabajador
            : String(row.documento || "").trim() || String(row.codigoTrabajador || "").trim()
              ? "(sin nombre)"
              : "",
          tone:
            (String(row.documento || "").trim() || String(row.codigoTrabajador || "").trim()) &&
            (isNombreTrabajadorVacio(row.trabajador) || row.dayFlags?.trabajador === "rojo")
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
          value: stackExportText(row.horasInicioDetalle),
          tone: iniFlag === "rojo" ? "danger" : "",
          tip: row.tipHoraInicio || ""
        },
        {
          value: stackExportText(row.horasFinDetalle),
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

  downloadXlsxExcel({
    filename: `QBerries_Detalle_${stamp}.xlsx`,
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
  const rowsData = getResumenTableRowsForExport();

  if (!rowsData.length) {
    window.alert("No hay filas en el resumen con los filtros actuales.");
    return;
  }

  const rows = rowsData.map((g) => ({
    rowTone:
      g.estadoKey === "no-asistio" || g.errores > 0
        ? "softDanger"
        : g.estadoKey === "falta-tarde" || g.estadoKey === "falta-manana" || g.estadoKey === "no-subio" || g.avisos > 0
          ? "softWarn"
          : g.estadoKey === "apoyo"
            ? "softOk"
            : "",
    cells: [
      g.fundo,
      g.supervisor,
      g.sgLabel || "—",
      g.planillas,
      g.cosechadores ?? g.trabajadores ?? 0,
      { value: g.errores, tone: g.errores > 0 ? "danger" : "" },
      { value: g.avisos, tone: g.avisos > 0 ? "warn" : "" },
      { value: g.faltaManana ? "Falta" : "OK", tone: g.faltaManana ? "warn" : "ok" },
      { value: g.faltaTarde ? "Falta" : "OK", tone: g.faltaTarde ? "warn" : "ok" },
      {
        value: g.estadoLabel,
        tone:
          g.estadoKey === "no-asistio"
            ? "danger"
            : g.estadoKey === "ok"
              ? "ok"
              : g.estadoKey === "apoyo"
                ? "ok"
                : "warn"
      },
      {
        value: g.apoyo ? `Sí · ${(g.apoyoNombres || []).join(" · ")}` : "No",
        tone: g.apoyo ? "ok" : ""
      }
    ]
  }));

  downloadXlsxExcel({
    filename: `QBerries_Resumen_Supervisores_${stamp}.xlsx`,
    sheetName: "Resumen",
    headers: [
      "Fundo",
      "Supervisor",
      "Sup. General",
      "Planillas",
      "Cosechadores",
      "Errores",
      "Extras",
      "Mañana",
      "Tarde",
      "Estado",
      "Apoyo equipo"
    ],
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
      !filters.estado &&
      !filters.search;
    if (empty) return true;
    if (filters.supervisor && item.supervisor && item.supervisor !== filters.supervisor) return false;
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
  (f.documentoVacio || []).forEach((i) =>
    add(
      "Documento vacío",
      i,
      i.codigoTrabajador ? `Código ${i.codigoTrabajador}` : i.macroPartida || "",
      i.fecha || "",
      "danger"
    )
  );
  (f.trabajadorVacio || []).forEach((i) =>
    add(
      "DNI sin nombre",
      i,
      i.codigoTrabajador ? `Código ${i.codigoTrabajador}` : i.macroPartida || "",
      i.fecha || "",
      "danger"
    )
  );
  (f.actividadNoPermitida || []).forEach((i) =>
    add(
      "Actividad no permitida",
      i,
      i.macroPartida || "COSTO DE COSECHA",
      i.actividad || "(vacía)",
      "danger"
    )
  );
  (f.cesados || []).forEach((i) => add("Cesado", i, "", "", "warn"));
  (f.minoritaria || []).forEach((i) => add("Menoritaria", i, i.macroPartida, "", "warn"));
  (f.overBase || []).forEach((i) =>
    add(`Extra ${i.extra || ""}`.trim(), i, i.day || "Suma de Horas Pago", i.hours, "warn")
  );

  downloadXlsxExcel({
    filename: `QBerries_Hallazgos_${stamp}.xlsx`,
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
  if (state.uploading) return;
  state.uploading = true;
  const pick = $("btnPickExcel");
  const card = $("uploadDropCard");
  if (pick) {
    pick.disabled = true;
    pick.setAttribute("aria-busy", "true");
  }
  card?.classList.add("is-busy");
  try {
    resetPlanillasFaltantesState();
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

    resetActividadFilterState();
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
  } catch (err) {
    console.error("[tareo] upload failed", err);
    window.alert(err?.message || "No se pudo leer el Excel. Revisa el archivo e inténtalo de nuevo.");
  } finally {
    state.uploading = false;
    if (pick) {
      pick.disabled = false;
      pick.removeAttribute("aria-busy");
    }
    card?.classList.remove("is-busy");
  }
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
    "fltEstado",
    "fltTipoLote"
  ].forEach((id) => {
    $(id)?.addEventListener("input", refreshView);
    $(id)?.addEventListener("change", refreshView);
  });

  // Búsqueda: debounce para no re-renderizar en cada tecla
  let searchTimer = 0;
  $("fltSearch")?.addEventListener("input", () => {
    window.clearTimeout(searchTimer);
    searchTimer = window.setTimeout(() => refreshView({ keepPage: true }), 180);
  });
  $("fltSearch")?.addEventListener("change", () => refreshView({ keepPage: true }));

  $("btnMarkChina")?.addEventListener("click", () => applySessionTag("china"));
  $("btnMarkConv")?.addEventListener("click", () => applySessionTag("convencional"));
  $("btnSetVariedad")?.addEventListener("click", applyVariedad);
  $("btnReportResumen")?.addEventListener("click", async () => {
    if (!state.validated) {
      window.alert("Primero sube un Excel de tareo.");
      return;
    }
    resetPlanillasFaltantesState();
    syncResumenFiltersFromMain(state.validated, readFilters());
    renderResumenView(state.validated, readFilters());
    await openResumenModal(() => state.validated);
  });
  $("btnReportHallazgos")?.addEventListener("click", () => downloadReport("hallazgos"));
  $("btnSupervisoresErrores")?.addEventListener("click", () => openSupervisoresErroresModal());
  $("btnCloseSupErrores")?.addEventListener("click", closeSupervisoresErroresModal);
  $("btnCloseSupErrores2")?.addEventListener("click", closeSupervisoresErroresModal);
  $("btnSupErroresFiltrar")?.addEventListener("click", () => {
    closeSupervisoresErroresModal();
    enterErrorFocusMode();
  });
  $("supErroresSearch")?.addEventListener("input", () => {
    renderSupervisoresErroresList($("supErroresSearch")?.value || "");
  });
  $("supErroresBody")?.addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-sup-err]");
    if (!tr) return;
    focusSupervisorFromError(tr.getAttribute("data-sup-err") || "");
  });
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
    resetPlanillasFaltantesState();
    resetActividadFilterState();
    closeResumenModal();
    $("validacionWorkspace")?.classList.add("is-hidden");
    $("uploadZone")?.classList.remove("is-hidden");
  });

  $("btnVerErrores")?.addEventListener("click", () => enterErrorFocusMode());
  $("btnPersonasErrores")?.addEventListener("click", () => openErroresPersonasModal());
  $("btnCloseErroresPersonas")?.addEventListener("click", closeErroresPersonasModal);
  $("btnCloseErroresPersonas2")?.addEventListener("click", closeErroresPersonasModal);
  $("btnErroresFiltrarTabla")?.addEventListener("click", () => {
    closeErroresPersonasModal();
    enterErrorFocusMode();
  });
  $("erroresPersonasSearch")?.addEventListener("input", () => {
    renderErroresPersonasList($("erroresPersonasSearch")?.value || "");
  });
  $("erroresPersonasBody")?.addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-err-dni]");
    if (!tr) return;
    focusPersonaFromError(tr.getAttribute("data-err-dni") || "", tr.getAttribute("data-err-fecha") || "");
  });
  $("btnRegresarErrores")?.addEventListener("click", () => exitErrorFocusMode({ restoreSaved: true }));
  $("btnVerAvisos")?.addEventListener("click", () => enterWarnFocusMode());
  $("btnRegresarAvisos")?.addEventListener("click", () => exitWarnFocusMode({ restoreSaved: true }));
  $("btnVerPases")?.addEventListener("click", () => enterPaseFocusMode());
  $("btnRegresarPases")?.addEventListener("click", () => exitPaseFocusMode({ restoreSaved: true }));
  $("btnVerDups")?.addEventListener("click", () => enterDupFocusMode());
  $("btnRegresarDups")?.addEventListener("click", () => exitDupFocusMode({ restoreSaved: true }));
  $("btnRestaurarTodo")?.addEventListener("click", () => restoreAllFilters());

  $("btnModulosCosecha")?.addEventListener("click", () => openModulosCosechaModal());
  $("btnCloseModulosCosecha")?.addEventListener("click", closeModulosCosechaModal);
  $("btnCloseModulosCosecha2")?.addEventListener("click", closeModulosCosechaModal);
  $("btnToggleModulosDetalle")?.addEventListener("click", toggleModulosDetalleView);
  $("btnExportModulosCosecha")?.addEventListener("click", exportModulosCosechaExcel);
  document.querySelectorAll('[data-close-modal="modulos-cosecha"]').forEach((el) => {
    el.addEventListener("click", closeModulosCosechaModal);
  });

  $("btnLotesHoy")?.addEventListener("click", () => openLotesHoyModal());
  $("btnCloseLotesHoy")?.addEventListener("click", closeLotesHoyModal);
  $("btnCloseLotesHoy2")?.addEventListener("click", closeLotesHoyModal);
  $("btnExportLotesHoy")?.addEventListener("click", exportLotesHoyExcel);
  $("lotesHoySearch")?.addEventListener("input", (e) => {
    lotesHoyUiState.search = e.target?.value || "";
    renderLotesHoyCards();
  });
  document.querySelectorAll('[data-close-modal="lotes-hoy"]').forEach((el) => {
    el.addEventListener("click", closeLotesHoyModal);
  });

  $("btnCloseSuccessModal")?.addEventListener("click", hideSuccessModal);
  $("btnCloseKpiHelp")?.addEventListener("click", hideKpiHelp);
  $("btnKpiHelpOk")?.addEventListener("click", hideKpiHelp);

  document.querySelectorAll("[data-close-modal]").forEach((el) => {
    el.addEventListener("click", () => {
      const which = el.getAttribute("data-close-modal");
      if (which === "kpi") hideKpiHelp();
      else if (which === "resumen") closeResumenModal();
      else if (which === "errores-personas") closeErroresPersonasModal();
      else if (which === "supervisores-errores") closeSupervisoresErroresModal();
      else if (which === "modulos-cosecha") closeModulosCosechaModal();
      else if (which === "lotes-hoy") closeLotesHoyModal();
      else if (which === "planillas-faltantes") {
        document.getElementById("modalPlanillasFaltantes").hidden = true;
      } else hideSuccessModal();
    });
  });

  // Info de KPIs (delegación: se re-renderizan)
  $("kpiRow")?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-kpi-help]");
    if (!btn) return;
    openKpiHelp(btn.getAttribute("data-kpi-help"));
  });

  bindTablePager(() => refreshView({ keepPage: true }));

  bindTareoActFilterUi();
  syncActividadFilterButton();

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
