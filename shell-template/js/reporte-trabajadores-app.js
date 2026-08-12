/**
 * Reporte de Trabajadores:
 * 1) Fundo (múltiple) → 2) Supervisor (col. W) → 3) Actividad (col. M) → Trabajador (col. D)
 */

import { parseExcelBuffer } from "./excel-parser.js";

const state = {
  fileName: "",
  rows: [],
  workers: [],
  filtered: [],
  allFundos: [],
  availableSupervisores: [],
  availableActividades: [],
  selectedFundos: [],
  selectedActividades: [],
  selectedSupervisores: []
};

const pager = {
  page: 1,
  pageSize: 20,
  bound: false
};

/** Modal draft for multi-select filters */
const filterModal = {
  kind: null, // "fundo" | "supervisor" | "actividad"
  options: [],
  draft: new Set(),
  bound: false
};

function $(id) {
  return document.getElementById(id);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function setStatus(kind, text) {
  const el = $("rtStatus");
  if (!el) return;
  if (!text) {
    el.hidden = true;
    el.textContent = "";
    el.removeAttribute("data-kind");
    return;
  }
  el.hidden = false;
  el.dataset.kind = kind || "";
  el.textContent = text;
}

function uniqueSorted(values) {
  return [...new Set(values.map((v) => String(v || "").trim()).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, "es")
  );
}

function formatMultiLabel(selected, placeholder, singular, plural) {
  if (!selected.length) return placeholder;
  if (selected.length === 1) return selected[0];
  return `${selected.length} ${plural}`;
}

function rowsForSelectedFundos() {
  const fundos = state.selectedFundos;
  if (!fundos.length) return [];
  const fundoSet = new Set(fundos);
  return state.rows.filter((r) => fundoSet.has(String(r.fundo || "").trim()));
}

function rowsForSelectedFundoSupervisor() {
  const fundos = state.selectedFundos;
  const supervisores = state.selectedSupervisores;
  if (!fundos.length || !supervisores.length) return [];
  const fundoSet = new Set(fundos);
  const supSet = new Set(supervisores);
  return state.rows.filter(
    (r) =>
      fundoSet.has(String(r.fundo || "").trim()) &&
      supSet.has(String(r.supervisor || "").trim())
  );
}

function syncFilterButtons() {
  const btnFundo = $("rtBtnFundo");
  const txtFundo = $("rtBtnFundoText");
  const btnSup = $("rtBtnSupervisor");
  const txtSup = $("rtBtnSupervisorText");
  const btnAct = $("rtBtnActividad");
  const txtAct = $("rtBtnActividadText");
  const search = $("rtFltSearch");

  const hasRows = state.rows.length > 0;
  const hasFundo = state.selectedFundos.length > 0;
  const hasSup = state.selectedSupervisores.length > 0;
  const hasAct = state.selectedActividades.length > 0;

  if (btnFundo) btnFundo.disabled = !hasRows;
  if (txtFundo) {
    txtFundo.textContent = formatMultiLabel(
      state.selectedFundos,
      btnFundo?.dataset.placeholder || "Selecciona fundo(s)…",
      "fundo",
      "fundos"
    );
  }
  btnFundo?.classList.toggle("has-value", hasFundo);

  if (btnSup) {
    btnSup.disabled = !hasFundo;
    btnSup.dataset.placeholder = hasFundo
      ? "Selecciona supervisor(es)…"
      : "Primero elige fundo…";
  }
  if (txtSup) {
    txtSup.textContent = formatMultiLabel(
      state.selectedSupervisores,
      btnSup?.dataset.placeholder || "Primero elige fundo…",
      "supervisor",
      "supervisores"
    );
  }
  btnSup?.classList.toggle("has-value", hasSup);

  if (btnAct) {
    btnAct.disabled = !hasSup;
    btnAct.dataset.placeholder = hasSup
      ? "Selecciona actividad(es)…"
      : "Primero elige supervisor…";
  }
  if (txtAct) {
    txtAct.textContent = formatMultiLabel(
      state.selectedActividades,
      btnAct?.dataset.placeholder || "Primero elige supervisor…",
      "actividad",
      "actividades"
    );
  }
  btnAct?.classList.toggle("has-value", hasAct);

  if (search) {
    search.disabled = !(hasFundo && hasSup && hasAct);
    if (search.disabled) search.value = "";
  }
}

function refreshFundoOptions() {
  state.allFundos = uniqueSorted(state.rows.map((r) => r.fundo));
  state.availableSupervisores = [];
  state.availableActividades = [];
  state.selectedFundos = [];
  state.selectedSupervisores = [];
  state.selectedActividades = [];
  syncFilterButtons();
}

function refreshSupervisorOptions({ selectAll = true } = {}) {
  const scoped = rowsForSelectedFundos();
  if (!scoped.length) {
    state.availableSupervisores = [];
    state.availableActividades = [];
    state.selectedSupervisores = [];
    state.selectedActividades = [];
    syncFilterButtons();
    return;
  }

  state.availableSupervisores = uniqueSorted(scoped.map((r) => r.supervisor));
  if (selectAll) {
    state.selectedSupervisores = [...state.availableSupervisores];
  } else {
    const allowed = new Set(state.availableSupervisores);
    state.selectedSupervisores = state.selectedSupervisores.filter((s) => allowed.has(s));
  }
  refreshActividadOptions({ selectAll: true });
}

function refreshActividadOptions({ selectAll = true } = {}) {
  const scoped = rowsForSelectedFundoSupervisor();
  if (!scoped.length) {
    state.availableActividades = [];
    state.selectedActividades = [];
    syncFilterButtons();
    return;
  }

  state.availableActividades = uniqueSorted(scoped.map((r) => r.actividad));

  if (selectAll) {
    state.selectedActividades = [...state.availableActividades];
  } else {
    const allowed = new Set(state.availableActividades);
    state.selectedActividades = state.selectedActividades.filter((a) => allowed.has(a));
  }
  syncFilterButtons();
}

function updateFilterCount() {
  const el = $("rtFilterCount");
  if (!el) return;
  const n = filterModal.draft.size;
  let noun = "items";
  let ending = n === 1 ? "o" : "os";
  if (filterModal.kind === "fundo") {
    noun = n === 1 ? "fundo" : "fundos";
  } else if (filterModal.kind === "actividad") {
    noun = n === 1 ? "actividad" : "actividades";
    ending = n === 1 ? "a" : "as";
  } else {
    noun = n === 1 ? "supervisor" : "supervisores";
  }
  el.textContent = `${n} ${noun} seleccionad${ending}`;
}

function renderFilterList(query = "") {
  const list = $("rtFilterList");
  if (!list) return;

  const q = String(query || "")
    .trim()
    .toLowerCase();
  const options = filterModal.options.filter((v) => !q || v.toLowerCase().includes(q));

  if (!options.length) {
    list.innerHTML = `<p class="rt-filter-modal__empty">${
      filterModal.options.length ? "Sin coincidencias" : "Sin opciones"
    }</p>`;
    return;
  }

  list.innerHTML = options
    .map((v) => {
      const isOn = filterModal.draft.has(v);
      const safe = escapeHtml(v);
      return `<label class="rt-filter-modal__option${isOn ? " is-checked" : ""}">
        <input type="checkbox" class="rt-filter-modal__check" value="${safe}" ${isOn ? "checked" : ""} />
        <span class="rt-filter-modal__box" aria-hidden="true">
          <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="3.5 8.5 6.5 11.5 12.5 4.5"></polyline>
          </svg>
        </span>
        <span class="rt-filter-modal__label">${safe}</span>
      </label>`;
    })
    .join("");
}

function openFilterModal(kind) {
  const modal = $("modalRtFilter");
  if (!modal) return;

  filterModal.kind = kind;
  if (kind === "fundo") {
    filterModal.options = [...state.allFundos];
    filterModal.draft = new Set(state.selectedFundos);
    $("rtFilterEyebrow").textContent = "Paso 1";
    $("rtFilterTitle").textContent = "Fundo";
    $("rtFilterHint").textContent = "Marca uno o varios fundos.";
    $("btnRtFilterAll").textContent = "Todos";
    $("btnRtFilterNone").textContent = "Ninguno";
  } else if (kind === "supervisor") {
    filterModal.options = [...state.availableSupervisores];
    filterModal.draft = new Set(state.selectedSupervisores);
    $("rtFilterEyebrow").textContent = "Paso 2";
    $("rtFilterTitle").textContent = "Supervisor";
    $("rtFilterHint").textContent = "Marca uno o varios supervisores del fundo elegido (columna W).";
    $("btnRtFilterAll").textContent = "Todos";
    $("btnRtFilterNone").textContent = "Ninguno";
  } else {
    filterModal.options = [...state.availableActividades];
    filterModal.draft = new Set(state.selectedActividades);
    $("rtFilterEyebrow").textContent = "Paso 3";
    $("rtFilterTitle").textContent = "Actividad";
    $("rtFilterHint").textContent = "Marca una o varias actividades (columna M).";
    $("btnRtFilterAll").textContent = "Todas";
    $("btnRtFilterNone").textContent = "Ninguna";
  }

  const search = $("rtFilterSearch");
  if (search) search.value = "";
  renderFilterList("");
  updateFilterCount();
  modal.hidden = false;
  window.setTimeout(() => search?.focus(), 40);
}

function closeFilterModal() {
  const modal = $("modalRtFilter");
  if (modal) modal.hidden = true;
  filterModal.kind = null;
  filterModal.options = [];
  filterModal.draft = new Set();
}

function applyFilterModal() {
  const kind = filterModal.kind;
  const selected = [...filterModal.draft].sort((a, b) => a.localeCompare(b, "es"));

  if (kind === "fundo") {
    const prev = state.selectedFundos.join("\0");
    state.selectedFundos = selected;
    const changed = prev !== selected.join("\0");
    refreshSupervisorOptions({ selectAll: changed || !state.selectedSupervisores.length });
  } else if (kind === "supervisor") {
    const prev = state.selectedSupervisores.join("\0");
    state.selectedSupervisores = selected;
    const changed = prev !== selected.join("\0");
    refreshActividadOptions({ selectAll: changed || !state.selectedActividades.length });
  } else if (kind === "actividad") {
    state.selectedActividades = selected;
    syncFilterButtons();
  }

  closeFilterModal();
  applyReportFilters({ resetPage: true });
}

/** Personas únicas por DNI dentro de los filtros (soporta múltiples). */
function buildWorkers(rows, { fundos = [], actividades = [], supervisores = [] } = {}) {
  const fundoSet = new Set(fundos);
  const actSet = new Set(actividades);
  const supSet = new Set(supervisores);
  const map = new Map();

  rows.forEach((row) => {
    const fundo = String(row.fundo || "").trim();
    const act = String(row.actividad || "").trim();
    const sup = String(row.supervisor || "").trim();
    if (fundoSet.size && !fundoSet.has(fundo)) return;
    if (actSet.size && !actSet.has(act)) return;
    if (supSet.size && !supSet.has(sup)) return;

    const doc = String(row.documento || "").trim();
    const name = String(row.trabajador || "").trim();
    if (!doc && !name) return;

    const key = doc || `nombre:${name.toLowerCase()}`;
    if (!map.has(key)) {
      map.set(key, {
        documento: doc,
        trabajador: name,
        actividades: new Set(),
        supervisores: new Set(),
        fundos: new Set(),
        registros: 0,
        fechas: new Set()
      });
    }
    const w = map.get(key);
    w.registros += 1;
    if (row.fecha) w.fechas.add(String(row.fecha));
    if (!w.trabajador && name) w.trabajador = name;
    if (act) w.actividades.add(act);
    if (sup) w.supervisores.add(sup);
    if (fundo) w.fundos.add(fundo);
  });

  return [...map.values()]
    .map((w) => ({
      documento: w.documento,
      trabajador: w.trabajador,
      actividad: [...w.actividades].sort((a, b) => a.localeCompare(b, "es")).join(" · "),
      supervisor: [...w.supervisores].sort((a, b) => a.localeCompare(b, "es")).join(" · "),
      fundo: [...w.fundos].sort((a, b) => a.localeCompare(b, "es")).join(" · "),
      registros: w.registros,
      dias: w.fechas.size
    }))
    .sort((a, b) => {
      const byName = String(a.trabajador || "").localeCompare(String(b.trabajador || ""), "es");
      if (byName) return byName;
      return String(a.documento || "").localeCompare(String(b.documento || ""), "es");
    });
}

function applyReportFilters({ resetPage = true } = {}) {
  const fundos = state.selectedFundos;
  const actividades = state.selectedActividades;
  const supervisores = state.selectedSupervisores;
  const search = ($("rtFltSearch")?.value || "").trim().toLowerCase();

  if (fundos.length && actividades.length && supervisores.length) {
    state.workers = buildWorkers(state.rows, { fundos, actividades, supervisores });
  } else {
    state.workers = [];
  }

  state.filtered = state.workers.filter((w) => {
    if (!search) return true;
    const blob = `${w.documento} ${w.trabajador}`.toLowerCase();
    return blob.includes(search);
  });

  if (resetPage) pager.page = 1;
  syncFilterButtons();
  renderKpis();
  renderMeta();
  renderTable();
  syncExportButton();
}

function renderKpis() {
  const fileEl = $("rtKpiFile");
  const actEl = $("rtKpiActividades");
  const supEl = $("rtKpiSupervisores");
  const trabEl = $("rtKpiTrabajadores");

  if (fileEl) fileEl.textContent = state.fileName || "—";
  if (supEl) {
    const count = state.selectedFundos.length
      ? state.availableSupervisores.length
      : uniqueSorted(state.rows.map((r) => r.supervisor)).length;
    supEl.textContent = String(count);
  }
  if (actEl) {
    const count = state.selectedSupervisores.length
      ? state.availableActividades.length
      : uniqueSorted(state.rows.map((r) => r.actividad)).length;
    actEl.textContent = String(count);
  }
  if (trabEl) trabEl.textContent = String(state.filtered.length);
}

function summarizeList(list, emptyLabel) {
  if (!list.length) return emptyLabel;
  if (list.length <= 2) return list.join(", ");
  return `${list.slice(0, 2).join(", ")} +${list.length - 2}`;
}

function renderMeta() {
  const meta = $("rtMeta");
  if (!meta) return;
  const fundos = state.selectedFundos;
  const actividades = state.selectedActividades;
  const supervisores = state.selectedSupervisores;

  if (!state.rows.length) {
    meta.textContent = "Sube un Excel para comenzar.";
    return;
  }
  if (!fundos.length) {
    meta.textContent = `${state.rows.length} filas leídas · elige uno o varios fundos.`;
    return;
  }
  if (!supervisores.length) {
    meta.textContent = `Fundo: ${summarizeList(fundos, "—")} · elige uno o varios supervisores.`;
    return;
  }
  if (!actividades.length) {
    meta.textContent = `Supervisor: ${summarizeList(supervisores, "—")} · elige una o varias actividades.`;
    return;
  }
  meta.textContent = `${state.filtered.length} trabajador${
    state.filtered.length === 1 ? "" : "es"
  } · ${summarizeList(fundos, "—")} · ${summarizeList(supervisores, "—")} · ${summarizeList(
    actividades,
    "—"
  )}`;
}

function renderPagerControls() {
  const rangeEl = $("rtPagerRange");
  const first = $("rtPagerFirst");
  const prev = $("rtPagerPrev");
  const next = $("rtPagerNext");
  const last = $("rtPagerLast");
  const sizeSel = $("rtPagerPageSize");

  const total = state.filtered.length;
  const size = pager.pageSize;
  const totalPages = Math.max(1, Math.ceil(total / size) || 1);
  if (pager.page > totalPages) pager.page = totalPages;
  if (pager.page < 1) pager.page = 1;

  const page = pager.page;
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

function renderTable() {
  const body = $("rtTableBody");
  const empty = $("rtEmpty");
  if (!body) return;

  renderPagerControls();
  const total = state.filtered.length;
  const start = (pager.page - 1) * pager.pageSize;
  const pageRows = state.filtered.slice(start, start + pager.pageSize);

  const fundos = state.selectedFundos;
  const actividades = state.selectedActividades;
  const supervisores = state.selectedSupervisores;

  if (!pageRows.length) {
    body.innerHTML = "";
    if (empty) {
      empty.hidden = false;
      const title = empty.querySelector(".pases-empty__title, .rt-empty__title");
      const text = empty.querySelector(".pases-empty__text, .rt-empty__text");
      if (!state.rows.length) {
        if (title) title.textContent = "Aún no hay data clasificada";
        if (text) text.textContent = "Sube un Excel para comenzar el reporte.";
      } else if (!fundos.length) {
        if (title) title.textContent = "Aún no hay data clasificada";
        if (text)
          text.textContent =
            "Los datos ya se leyeron. Paso 1: toca Fundo y marca uno o varios.";
      } else if (!supervisores.length) {
        if (title) title.textContent = "Aún no hay data clasificada";
        if (text)
          text.textContent =
            "Paso 2: toca Supervisor y marca uno o varios del fundo elegido.";
      } else if (!actividades.length) {
        if (title) title.textContent = "Aún no hay data clasificada";
        if (text)
          text.textContent =
            "Paso 3: toca Actividad y marca las del supervisor para ver los trabajadores.";
      } else {
        if (title) title.textContent = "Sin trabajadores";
        if (text) text.textContent = "No hay trabajadores para esa combinación de filtros.";
      }
    }
    return;
  }

  if (empty) empty.hidden = true;
  body.innerHTML = pageRows
    .map((row, i) => {
      const n = start + i + 1;
      return `<tr>
        <td class="rt-table__num">${n}</td>
        <td class="rt-table__doc">${escapeHtml(row.documento || "—")}</td>
        <td>
          <div class="rt-table__primary">${escapeHtml(row.trabajador || "—")}</div>
        </td>
        <td><span class="rt-pill" title="${escapeHtml(row.actividad || "")}">${escapeHtml(
          row.actividad || "—"
        )}</span></td>
        <td title="${escapeHtml(row.supervisor || "")}">${escapeHtml(row.supervisor || "—")}</td>
        <td>${escapeHtml(row.fundo || "—")}</td>
        <td class="rt-table__num">${row.registros}</td>
      </tr>`;
    })
    .join("");
}

function syncExportButton() {
  const btn = $("btnRtExport");
  if (btn) btn.disabled = !state.filtered.length;
}

function showWorkspace(show) {
  // Estilo Pases: UI siempre visible; Subir Excel en el header.
  const newBtn = $("btnRtNewFile");
  if (newBtn) newBtn.hidden = !show;
}

async function handleFile(file) {
  if (!file) return;
  try {
    setStatus("loading", "Leyendo Excel…");
    const buffer = await file.arrayBuffer();
    const parsed = parseExcelBuffer(buffer, file.name);
    state.fileName = file.name;
    state.rows = Array.isArray(parsed?.rows) ? parsed.rows : [];
    state.workers = [];
    state.filtered = [];
    state.selectedFundos = [];
    state.selectedActividades = [];
    state.selectedSupervisores = [];
    pager.page = 1;

    showWorkspace(true);
    refreshFundoOptions();
    applyReportFilters({ resetPage: true });
    if (!state.rows.length) setStatus("empty", "No se encontraron filas en el Excel.");
    else setStatus("", "");
  } catch (err) {
    console.error("[reporte-trabajadores]", err);
    setStatus("error", err?.message || "No se pudo leer el Excel.");
  }
}

function exportExcel() {
  if (!state.filtered.length) {
    setStatus("empty", "No hay trabajadores para exportar. Elige fundo, supervisor y actividad.");
    return;
  }
  const XLSX = window.XLSX;
  if (!XLSX?.utils || typeof XLSX.write !== "function") {
    setStatus("error", "No se pudo cargar el exportador Excel (XLSX).");
    return;
  }

  try {
    const rows = state.filtered.map((w, i) => ({
      "#": i + 1,
      Documento: w.documento || "",
      Trabajador: w.trabajador || "",
      Actividad: w.actividad || "",
      Supervisor: w.supervisor || "",
      Fundo: w.fundo || "",
      Registros: w.registros
    }));

    const sheet = XLSX.utils.json_to_sheet(rows);
    sheet["!cols"] = [
      { wch: 5 },
      { wch: 14 },
      { wch: 36 },
      { wch: 28 },
      { wch: 34 },
      { wch: 16 },
      { wch: 10 }
    ];
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, "Trabajadores");

    const safeAct =
      state.selectedActividades.length === 1
        ? state.selectedActividades[0].replace(/[^\wÀ-ÿ\-]+/gi, "_").slice(0, 24)
        : `${state.selectedActividades.length}actividades`;
    const safeSup =
      state.selectedSupervisores.length === 1
        ? state.selectedSupervisores[0].replace(/[^\wÀ-ÿ\-]+/gi, "_").slice(0, 24)
        : `${state.selectedSupervisores.length}supervisores`;
    const stamp = new Date().toISOString().slice(0, 10);
    const fileName = `reporte-trabajadores-${safeAct}-${safeSup}-${stamp}.xlsx`;

    const buffer = XLSX.write(book, { bookType: "xlsx", type: "array" });
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1500);
    setStatus("", "");
  } catch (err) {
    console.error("[reporte-trabajadores] export", err);
    setStatus("error", err?.message || "Error al exportar Excel.");
  }
}

function bindPager() {
  if (pager.bound) return;
  pager.bound = true;

  $("rtPagerPageSize")?.addEventListener("change", (e) => {
    const n = Number(e.target.value);
    pager.pageSize = Number.isFinite(n) && n > 0 ? n : 20;
    pager.page = 1;
    renderTable();
  });
  $("rtPagerFirst")?.addEventListener("click", () => {
    pager.page = 1;
    renderTable();
  });
  $("rtPagerPrev")?.addEventListener("click", () => {
    pager.page = Math.max(1, pager.page - 1);
    renderTable();
  });
  $("rtPagerNext")?.addEventListener("click", () => {
    const totalPages = Math.max(1, Math.ceil(state.filtered.length / pager.pageSize) || 1);
    pager.page = Math.min(totalPages, pager.page + 1);
    renderTable();
  });
  $("rtPagerLast")?.addEventListener("click", () => {
    pager.page = Math.max(1, Math.ceil(state.filtered.length / pager.pageSize) || 1);
    renderTable();
  });
}

function bindUpload() {
  const input = $("rtInputExcel");
  const pick = $("btnRtPickExcel");
  const card = $("rtDropCard");

  pick?.addEventListener("click", () => input?.click());
  input?.addEventListener("change", () => {
    const file = input.files?.[0];
    handleFile(file);
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
    handleFile(file);
  });
}

function bindFilterModal() {
  if (filterModal.bound) return;
  filterModal.bound = true;

  $("rtBtnFundo")?.addEventListener("click", () => {
    if (!state.rows.length) return;
    openFilterModal("fundo");
  });
  $("rtBtnSupervisor")?.addEventListener("click", () => {
    if (!state.selectedFundos.length) return;
    openFilterModal("supervisor");
  });
  $("rtBtnActividad")?.addEventListener("click", () => {
    if (!state.selectedSupervisores.length) return;
    openFilterModal("actividad");
  });

  $("btnCloseRtFilter")?.addEventListener("click", closeFilterModal);
  $("btnRtFilterCancel")?.addEventListener("click", closeFilterModal);
  $("btnRtFilterApply")?.addEventListener("click", applyFilterModal);

  document.querySelectorAll('[data-close-modal="rt-filter"]').forEach((el) => {
    el.addEventListener("click", closeFilterModal);
  });

  $("rtFilterSearch")?.addEventListener("input", (e) => {
    renderFilterList(e.target.value);
  });

  $("btnRtFilterAll")?.addEventListener("click", () => {
    filterModal.options.forEach((v) => filterModal.draft.add(v));
    renderFilterList($("rtFilterSearch")?.value || "");
    updateFilterCount();
  });
  $("btnRtFilterNone")?.addEventListener("click", () => {
    filterModal.draft.clear();
    renderFilterList($("rtFilterSearch")?.value || "");
    updateFilterCount();
  });

  $("rtFilterList")?.addEventListener("change", (e) => {
    const input = e.target;
    if (!(input instanceof HTMLInputElement) || !input.classList.contains("rt-filter-modal__check")) return;
    if (input.checked) filterModal.draft.add(input.value);
    else filterModal.draft.delete(input.value);
    const option = input.closest(".rt-filter-modal__option");
    option?.classList.toggle("is-checked", input.checked);
    updateFilterCount();
  });

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    const modal = $("modalRtFilter");
    if (modal && !modal.hidden) closeFilterModal();
  });
}

function resetModule() {
  state.fileName = "";
  state.rows = [];
  state.workers = [];
  state.filtered = [];
  state.allFundos = [];
  state.availableSupervisores = [];
  state.availableActividades = [];
  state.selectedFundos = [];
  state.selectedActividades = [];
  state.selectedSupervisores = [];
  pager.page = 1;
  closeFilterModal();
  showWorkspace(false);
  setStatus("", "");
  syncFilterButtons();
  syncExportButton();
  const meta = $("rtMeta");
  if (meta) meta.textContent = "Sube un Excel para comenzar.";
}

function bindUi() {
  bindUpload();
  bindPager();
  bindFilterModal();

  $("rtFltSearch")?.addEventListener("input", () => applyReportFilters({ resetPage: true }));
  $("btnRtExport")?.addEventListener("click", () => exportExcel());
  $("btnRtNewFile")?.addEventListener("click", () => resetModule());
}

let started = false;

function init() {
  if (started) return;
  started = true;
  bindUi();
  renderTable();
}

window.addEventListener("qb:route-changed", (evt) => {
  if (evt.detail?.route === "reporte-trabajadores") init();
});

if (window.location.hash.replace(/^#\/?/, "").split("/")[0] === "reporte-trabajadores") {
  init();
}
