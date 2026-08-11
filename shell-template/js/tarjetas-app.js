/**
 * Tarjeta Pallet — consulta solo lectura (proxy /api/tarjetas).
 */

import {
  reporteTarjetas,
  peekCachedReporteTarjetas,
  limaTodayYmd,
  normalizeText,
  emptyKpis
} from "./tarjetas-api.js";

const EXPORT_COLUMNS = [
  ["fecha", "Fecha"],
  ["nGuia", "N° Guía"],
  ["correlativo", "Correlativo"],
  ["placa", "Placa"],
  ["lugar", "Lugar"],
  ["variedad", "Variedad"],
  ["modulo", "Módulo"],
  ["turno", "Turno"],
  ["lote", "Lote"],
  ["jarras", "Jarras"],
  ["jabas", "Jabas"],
  ["pesoBruto", "Peso bruto"],
  ["pesoNeto", "Peso neto"],
  ["dni", "DNI"],
  ["horaRegistrada", "Hora registrada"]
];

const state = {
  rows: [],
  filtered: [],
  loading: false,
  error: "",
  selectedIndex: -1,
  loadedAt: null,
  kpis: emptyKpis(),
  apiRango: null,
  fromCache: false,
  dataSig: "",
  /** 'hoy' | 'todas' | 'custom' */
  fechaMode: "hoy"
};

const pagerState = {
  page: 1,
  pageSize: 20
};

let started = false;
let activeRoute = false;
let loadSeq = 0;
let lastQuerySig = "";

function querySig(opts) {
  return opts?.todas ? "todas" : `fecha:${opts?.fecha || limaTodayYmd()}`;
}

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

function display(value) {
  const text = String(value ?? "").trim();
  return text || "—";
}

function formatNum(value) {
  if (value === "" || value == null) return "—";
  const n = Number(value);
  if (!Number.isFinite(n)) return display(value);
  return n.toLocaleString("es-PE", { maximumFractionDigits: 2 });
}

function getFilters() {
  return {
    fecha: ($("fltTarjetaFecha")?.value || "").trim(),
    placa: ($("fltTarjetaPlaca")?.value || "").trim(),
    q: ($("fltTarjetaSearch")?.value || "").trim()
  };
}

function uniqueSorted(values) {
  return [...new Set(values.map((v) => String(v || "").trim()).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, "es")
  );
}

function syncSelectOptions() {
  const placaSel = $("fltTarjetaPlaca");
  if (!placaSel) return;

  const currentPlaca = placaSel.value;
  const placas = uniqueSorted(state.rows.map((r) => r.placa));

  placaSel.innerHTML =
    `<option value="">Todas</option>` +
    placas.map((p) => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join("");

  if (placas.includes(currentPlaca)) placaSel.value = currentPlaca;
}

function applyFilters({ resetPage = true } = {}) {
  syncSelectOptions();

  const f = getFilters();
  const q = normalizeText(f.q);
  const placaFilter = normalizeText(f.placa);

  state.filtered = state.rows.filter((row) => {
    if (placaFilter && normalizeText(row.placa) !== placaFilter) return false;

    if (q) {
      const qDigits = q.replace(/\D/g, "");
      const blob = normalizeText(
        `${row.nGuia} ${row.correlativo} ${row.placa} ${row.lugar} ${row.variedad} ${row.lote} ${row.modulo} ${row.dni}`
      );
      const dniOk = qDigits && String(row.dni || "").includes(qDigits);
      if (!blob.includes(q) && !dniOk) return false;
    }
    return true;
  });

  if (resetPage) pagerState.page = 1;
  renderKpis();
  renderTable();
  renderMeta();
}

function renderKpis() {
  const k = state.kpis || emptyKpis();
  const set = (id, val) => {
    const el = $(id);
    if (el) el.textContent = String(val);
  };
  const num = (v) => formatNum(v).replace("—", "0");
  set("kpiTarjetasGuias", k.totalGuias);
  set("kpiTarjetasJarras", num(k.totalJarras));
  set("kpiTarjetasJabas", num(k.totalJabas));
  set("kpiTarjetasPesoBruto", num(k.pesoBrutoTotal));
  set("kpiTarjetasPesoNeto", num(k.pesoNetoTotal));
}

function setStatus(kind, message) {
  const el = $("tarjetasStatus");
  if (!el) return;
  el.hidden = !message;
  el.dataset.kind = kind || "";
  el.textContent = message || "";
}

function renderMeta() {
  const meta = $("tarjetasMeta");
  if (!meta) return;
  if (!state.loadedAt && !state.rows.length) {
    meta.textContent = "Sin datos cargados";
    return;
  }
  const mode =
    state.fechaMode === "todas"
      ? "todas las fechas"
      : state.fechaMode === "hoy"
        ? "hoy (Lima)"
        : ($("fltTarjetaFecha")?.value || "día");
  const stamp = state.loadedAt
    ? state.loadedAt.toLocaleString("es-PE", { timeZone: "America/Lima" })
    : "—";
  const cacheNote = state.fromCache ? " · respaldo local" : "";
  meta.textContent = `${state.filtered.length} fila${
    state.filtered.length === 1 ? "" : "s"
  } · ${mode} · actualizado ${stamp}${cacheNote}`;
}

function renderPagerControls() {
  const rangeEl = $("tarjetasPagerRange");
  const first = $("tarjetasPagerFirst");
  const prev = $("tarjetasPagerPrev");
  const next = $("tarjetasPagerNext");
  const last = $("tarjetasPagerLast");
  const sizeSel = $("tarjetasPagerPageSize");

  const total = state.filtered.length;
  const size = pagerState.pageSize;
  const totalPages = Math.max(1, Math.ceil(total / size) || 1);
  if (pagerState.page > totalPages) pagerState.page = totalPages;
  if (pagerState.page < 1) pagerState.page = 1;

  const page = pagerState.page;
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

function getPageRows() {
  const start = (pagerState.page - 1) * pagerState.pageSize;
  return state.filtered.slice(start, start + pagerState.pageSize);
}

function renderTable() {
  const tbody = $("tarjetasTableBody");
  const empty = $("tarjetasEmpty");
  if (!tbody) return;

  renderPagerControls();
  const pageRows = getPageRows();

  if (!state.filtered.length) {
    tbody.innerHTML = "";
    if (empty) {
      empty.hidden = state.loading || Boolean(state.error);
      const titleEl = empty.querySelector(".tarjetas-empty__title");
      const textEl = empty.querySelector(".tarjetas-empty__text");
      if (titleEl) titleEl.textContent = state.rows.length ? "Sin coincidencias" : "Sin tarjetas";
      if (textEl) {
        textEl.textContent = state.rows.length
          ? "Ninguna fila coincide con los filtros."
          : "No hay tarjetas registradas para este rango.";
      }
    }
    return;
  }

  if (empty) empty.hidden = true;
  tbody.innerHTML = pageRows
    .map((row) => {
      const absIndex = state.filtered.indexOf(row);
      return `<tr class="tarjetas-row" data-index="${absIndex}" tabindex="0" role="button">
        <td>
          <div class="tarjetas-table__primary">${escapeHtml(display(row.nGuia))}</div>
          <div class="tarjetas-table__sub">${escapeHtml(display(row.correlativo))}</div>
        </td>
        <td>
          <div>${escapeHtml(display(row.fecha))}</div>
          <div class="tarjetas-table__sub">${escapeHtml(display(row.horaRegistrada))}</div>
        </td>
        <td><span class="tarjeta-badge tarjeta-badge--placa">${escapeHtml(display(row.placa))}</span></td>
        <td>
          <div>${escapeHtml(display(row.lugar))}</div>
          <div class="tarjetas-table__sub">${escapeHtml(display(row.variedad))}</div>
        </td>
        <td>
          <div>${escapeHtml(display(row.modulo))} · ${escapeHtml(display(row.lote))}</div>
          <div class="tarjetas-table__sub">${escapeHtml(display(row.turno))}</div>
        </td>
        <td class="tarjetas-table__num">${escapeHtml(formatNum(row.jarras))}</td>
        <td class="tarjetas-table__num">${escapeHtml(formatNum(row.jabas))}</td>
        <td class="tarjetas-table__num">${escapeHtml(formatNum(row.pesoNeto))}</td>
        <td class="tarjetas-table__doc">${escapeHtml(display(row.dni))}</td>
      </tr>`;
    })
    .join("");
}

function syncFechaModeButtons() {
  const hoy = $("btnTarjetasHoy");
  const todas = $("btnTarjetasClearFecha");
  hoy?.classList.toggle("is-active", state.fechaMode === "hoy");
  todas?.classList.toggle("is-active", state.fechaMode === "todas");
  hoy?.classList.toggle("btn--secondary", state.fechaMode === "hoy");
  hoy?.classList.toggle("btn--ghost", state.fechaMode !== "hoy");
  todas?.classList.toggle("btn--secondary", state.fechaMode === "todas");
  todas?.classList.toggle("btn--ghost", state.fechaMode !== "todas");
}

function setFechaMode(mode, fechaValue) {
  state.fechaMode = mode;
  const el = $("fltTarjetaFecha");
  if (el) el.value = fechaValue || "";
  syncFechaModeButtons();
}

function rowsSignature(rows) {
  return `${rows.length}|${rows
    .slice(0, 3)
    .map((r) => `${r.nGuia}:${r.correlativo}:${r.horaRegistrada}`)
    .join(";")}|${rows[rows.length - 1]?.nGuia || ""}`;
}

function applyApiPayload(resp, { resetPage = true, touchMeta = true } = {}) {
  state.rows = Array.isArray(resp.data) ? resp.data : [];
  state.kpis = resp.kpis || emptyKpis();
  state.apiRango = resp.filtros
    ? {
        modo: resp.filtros.todas ? "todas" : "dia",
        fecha: resp.filtros.fecha || null,
        todas: Boolean(resp.filtros.todas)
      }
    : null;
  state.fromCache = Boolean(resp.fromCache);
  state.dataSig = rowsSignature(state.rows);
  state.error = "";
  if (touchMeta) state.loadedAt = new Date();
  applyFilters({ resetPage });
}

function listarOptsFromUi(keepFilters) {
  const fechaInput = state.fechaMode === "todas"
    ? ""
    : state.fechaMode === "hoy"
      ? limaTodayYmd()
      : ($("fltTarjetaFecha")?.value || "").trim() || limaTodayYmd();

  if (!keepFilters) {
    const fechaEl = $("fltTarjetaFecha");
    if (fechaEl) fechaEl.value = fechaInput;
  }
  return {
    fechaInput,
    opts: fechaInput ? { fecha: fechaInput, limit: 2000 } : { todas: true, limit: 2000 }
  };
}

/**
 * Carga tarjetas.
 * - force=false: memoria/caché local (sin GET si hay respaldo de esa fecha)
 * - force=true: nuevo GET (Actualizar o cambio de fecha)
 */
async function loadTarjetas({ keepFilters = true, force = false } = {}) {
  const seq = ++loadSeq;
  const btn = $("btnTarjetasRefresh");
  const { fechaInput, opts } = listarOptsFromUi(keepFilters);
  const sig = querySig(opts);

  syncFechaModeButtons();

  const peeked = peekCachedReporteTarjetas(opts);
  const memoryHit = !force && state.rows.length && lastQuerySig === sig;

  if (memoryHit) {
    applyFilters({ resetPage: false });
    renderMeta();
    setStatus("", "");
    return;
  }

  if (!force && peeked) {
    applyApiPayload(peeked, { resetPage: true, touchMeta: true });
    lastQuerySig = sig;
    setStatus("", "");
    return;
  }

  if (force && peeked) {
    applyApiPayload(peeked, { resetPage: true, touchMeta: true });
  }

  setStatus(
    "loading",
    force
      ? fechaInput
        ? "Actualizando día seleccionado…"
        : "Actualizando todas las fechas…"
      : fechaInput
        ? "Cargando tarjetas del día…"
        : "Cargando todas las fechas…"
  );
  if (btn) btn.disabled = true;
  state.loading = true;

  try {
    const resp = await reporteTarjetas({ ...opts, force: Boolean(force) });
    if (seq !== loadSeq) return;

    applyApiPayload(resp, { resetPage: true, touchMeta: true });
    lastQuerySig = sig;

    if (state.fromCache) {
      setStatus(
        "empty",
        "Mostrando respaldo local (caché). Pulsa Actualizar para pedir datos nuevos."
      );
    } else {
      setStatus(
        state.filtered.length ? "" : "empty",
        state.filtered.length
          ? ""
          : state.rows.length
            ? "Ninguna fila coincide con los filtros."
            : "No hay tarjetas registradas."
      );
    }
  } catch (err) {
    if (seq !== loadSeq) return;
    const msg = err?.message || "Error al cargar tarjetas";
    if (peeked || state.rows.length) {
      if (peeked && !state.rows.length) applyApiPayload(peeked, { resetPage: true, touchMeta: true });
      lastQuerySig = sig;
      setStatus("error", `${msg} · mostrando respaldo local`);
    } else {
      state.error = msg;
      state.rows = [];
      state.filtered = [];
      state.kpis = emptyKpis();
      state.apiRango = null;
      state.dataSig = "";
      lastQuerySig = "";
      renderKpis();
      renderTable();
      renderMeta();
      syncFechaModeButtons();
      setStatus("error", state.error);
    }
  } finally {
    if (seq === loadSeq) {
      state.loading = false;
      if (btn) btn.disabled = false;
      syncFechaModeButtons();
    }
  }
}

function fieldCell(label, value, extraClass = "") {
  return `<div class="tarjeta-sheet__field${extraClass ? ` ${extraClass}` : ""}">
    <span class="tarjeta-sheet__field-label">${escapeHtml(label)}</span>
    <span class="tarjeta-sheet__field-value">${escapeHtml(value)}</span>
  </div>`;
}

function openDetail(index) {
  const row = state.filtered[index];
  const modal = $("modalTarjetaDetail");
  const body = $("tarjetaDetailBody");
  const title = $("tarjetaDetailTitle");
  const sub = $("tarjetaDetailSub");
  const badges = $("tarjetaDetailBadges");
  const metrics = $("tarjetaDetailMetrics");
  if (!row || !modal || !body) return;

  state.selectedIndex = index;
  if (title) title.textContent = row.nGuia ? `Guía ${row.nGuia}` : "Detalle de tarjeta";
  if (sub) {
    sub.textContent = [display(row.lugar), display(row.variedad)].filter((v) => v !== "—").join(" · ") || "Sin lugar / variedad";
  }
  if (badges) {
    badges.innerHTML = `
      <span class="tarjeta-sheet__badge">${escapeHtml(display(row.placa))}</span>
      <span class="tarjeta-sheet__badge tarjeta-sheet__badge--soft">${escapeHtml(display(row.fecha))}</span>
    `;
  }
  if (metrics) {
    metrics.innerHTML = `
      <div class="tarjeta-sheet__metric">
        <span class="tarjeta-sheet__metric-label">Jarras</span>
        <strong class="tarjeta-sheet__metric-value">${escapeHtml(formatNum(row.jarras))}</strong>
      </div>
      <div class="tarjeta-sheet__metric">
        <span class="tarjeta-sheet__metric-label">Jabas</span>
        <strong class="tarjeta-sheet__metric-value">${escapeHtml(formatNum(row.jabas))}</strong>
      </div>
      <div class="tarjeta-sheet__metric">
        <span class="tarjeta-sheet__metric-label">Peso bruto</span>
        <strong class="tarjeta-sheet__metric-value">${escapeHtml(formatNum(row.pesoBruto))}</strong>
      </div>
      <div class="tarjeta-sheet__metric">
        <span class="tarjeta-sheet__metric-label">Peso neto</span>
        <strong class="tarjeta-sheet__metric-value">${escapeHtml(formatNum(row.pesoNeto))}</strong>
      </div>
    `;
  }

  body.innerHTML = `
    <section class="tarjeta-sheet__section">
      <h3 class="tarjeta-sheet__section-title">Identificación</h3>
      <div class="tarjeta-sheet__grid">
        ${fieldCell("Correlativo / Guía campo", display(row.correlativo), "tarjeta-sheet__field--span2")}
        ${fieldCell("N° Guía", display(row.nGuia))}
        ${fieldCell("DNI", display(row.dni))}
        ${fieldCell("Hora registrada", display(row.horaRegistrada), "tarjeta-sheet__field--span2")}
      </div>
    </section>
    <section class="tarjeta-sheet__section">
      <h3 class="tarjeta-sheet__section-title">Operación</h3>
      <div class="tarjeta-sheet__grid">
        ${fieldCell("Lugar / Fundo", display(row.lugar))}
        ${fieldCell("Variedad", display(row.variedad))}
        ${fieldCell("Módulo", display(row.modulo))}
        ${fieldCell("Turno", display(row.turno))}
        ${fieldCell("Lote", display(row.lote))}
        ${fieldCell("Placa", display(row.placa))}
      </div>
    </section>
  `;
  modal.hidden = false;
}

function closeDetail() {
  const modal = $("modalTarjetaDetail");
  if (modal) modal.hidden = true;
  state.selectedIndex = -1;
}

function exportExcel() {
  const rows = state.filtered;
  if (!rows.length) {
    setStatus("empty", "No hay filas para exportar.");
    return;
  }
  const XLSX = window.XLSX;
  if (!XLSX?.utils || typeof XLSX.write !== "function") {
    setStatus("error", "No se pudo cargar el exportador Excel (XLSX).");
    return;
  }

  try {
    const data = rows.map((row, i) => {
      const out = { "#": i + 1 };
      EXPORT_COLUMNS.forEach(([key, label]) => {
        out[label] = row[key] ?? "";
      });
      return out;
    });
    const sheet = XLSX.utils.json_to_sheet(data);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, "Tarjetas");
    const stamp = limaTodayYmd();
    const buffer = XLSX.write(book, { bookType: "xlsx", type: "array" });
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `tarjeta-pallet-${stamp}.xlsx`;
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1500);
    setStatus("", "");
  } catch (err) {
    console.error("[tarjetas] export", err);
    setStatus("error", err?.message || "Error al exportar Excel.");
  }
}

function bindPager() {
  $("tarjetasPagerPageSize")?.addEventListener("change", (e) => {
    const n = Number(e.target.value);
    pagerState.pageSize = Number.isFinite(n) && n > 0 ? n : 20;
    pagerState.page = 1;
    renderTable();
  });
  $("tarjetasPagerFirst")?.addEventListener("click", () => {
    pagerState.page = 1;
    renderTable();
  });
  $("tarjetasPagerPrev")?.addEventListener("click", () => {
    pagerState.page = Math.max(1, pagerState.page - 1);
    renderTable();
  });
  $("tarjetasPagerNext")?.addEventListener("click", () => {
    const totalPages = Math.max(1, Math.ceil(state.filtered.length / pagerState.pageSize) || 1);
    pagerState.page = Math.min(totalPages, pagerState.page + 1);
    renderTable();
  });
  $("tarjetasPagerLast")?.addEventListener("click", () => {
    pagerState.page = Math.max(1, Math.ceil(state.filtered.length / pagerState.pageSize) || 1);
    renderTable();
  });
}

const CHART_PALETTE = ["#145a34", "#27ae60", "#4fb04a", "#c4d62e", "#f6921e", "#1a7a4a", "#6bbf8a", "#94a3b8"];
/** @type {Record<string, any>} */
const resumenCharts = {};

function destroyResumenCharts() {
  Object.keys(resumenCharts).forEach((key) => {
    try {
      resumenCharts[key]?.destroy();
    } catch (_) {}
    delete resumenCharts[key];
  });
}

function aggregateBy(rows, keyFn) {
  const map = new Map();
  rows.forEach((row) => {
    const key = String(keyFn(row) || "").trim() || "Sin dato";
    if (!map.has(key)) map.set(key, { label: key, jarras: 0, jabas: 0, filas: 0, guias: new Set() });
    const item = map.get(key);
    item.jarras += Number(row.jarras) || 0;
    item.jabas += Number(row.jabas) || 0;
    item.filas += 1;
    if (row.nGuia) item.guias.add(`${row.nGuia}|${row.fecha || ""}`);
  });
  return [...map.values()].sort((a, b) => b.jarras - a.jarras || a.label.localeCompare(b.label, "es"));
}

function makeChart(key, canvas, config) {
  if (!window.Chart || !canvas) return;
  if (resumenCharts[key]) {
    try {
      resumenCharts[key].destroy();
    } catch (_) {}
  }
  resumenCharts[key] = new Chart(canvas.getContext("2d"), config);
}

function aggregateGuias(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const guia = String(row.nGuia || "").trim() || "Sin guía";
    const fecha = String(row.fecha || "").trim();
    const key = `${guia}|${fecha}`;
    if (!map.has(key)) {
      map.set(key, {
        label: fecha ? `${guia} (${fecha.slice(5)})` : guia,
        nGuia: guia,
        jarras: 0,
        jabas: 0,
        pesoNeto: 0,
        filas: 0
      });
    }
    const item = map.get(key);
    item.jarras += Number(row.jarras) || 0;
    item.jabas += Number(row.jabas) || 0;
    item.pesoNeto += Number(row.pesoNeto) || 0;
    item.filas += 1;
  });
  return [...map.values()].sort((a, b) => b.jarras - a.jarras || a.label.localeCompare(b.label, "es"));
}

function renderResumenDashboard() {
  const rows = state.filtered.length ? state.filtered : state.rows;
  const kpisHost = $("tarjetasResumenKpis");
  const sub = $("tarjetasResumenSub");

  const guías = new Set(rows.map((r) => `${r.nGuia}|${r.fecha || ""}`).filter(Boolean)).size;
  let jarras = 0;
  let jabas = 0;
  let pesoNeto = 0;
  rows.forEach((r) => {
    jarras += Number(r.jarras) || 0;
    jabas += Number(r.jabas) || 0;
    pesoNeto += Number(r.pesoNeto) || 0;
  });

  const topGuia = aggregateGuias(rows)[0];
  const topMod = aggregateBy(rows, (r) => r.modulo)[0];

  if (sub) {
    const mode =
      state.fechaMode === "todas"
        ? "todas las fechas"
        : state.fechaMode === "hoy"
          ? "hoy (Lima)"
          : ($("fltTarjetaFecha")?.value || "día seleccionado");
    const highlight = [
      topGuia ? `Guía líder: ${topGuia.nGuia} (${formatNum(topGuia.jarras)} jarras)` : null,
      topMod ? `Módulo líder: ${topMod.label}` : null
    ]
      .filter(Boolean)
      .join(" · ");
    sub.textContent = `${rows.length} filas · ${mode}${highlight ? ` · ${highlight}` : ""}`;
  }

  if (kpisHost) {
    kpisHost.innerHTML = [
      ["Guías", guías],
      ["Jarras", Math.round(jarras * 100) / 100],
      ["Jabas", Math.round(jabas * 100) / 100],
      ["Peso neto", Math.round(pesoNeto * 100) / 100]
    ]
      .map(
        ([label, val]) => `<div class="tj-resumen__kpi">
        <span class="tj-resumen__kpi-label">${escapeHtml(label)}</span>
        <strong class="tj-resumen__kpi-value">${escapeHtml(formatNum(val).replace("—", "0"))}</strong>
      </div>`
      )
      .join("");
  }

  if (!window.Chart) {
    setStatus("error", "Chart.js no está disponible para el resumen.");
    return;
  }

  Chart.defaults.font.family = '"Source Sans 3", "Segoe UI", sans-serif';
  Chart.defaults.color = "#4a6354";

  const topGuias = aggregateGuias(rows).slice(0, 8);
  makeChart("guias", $("chartTarjetasGuias"), {
    type: "bar",
    data: {
      labels: topGuias.map((x) => x.label),
      datasets: [
        {
          label: "Jarras",
          data: topGuias.map((x) => x.jarras),
          backgroundColor: "#145a34",
          borderRadius: 6,
          maxBarThickness: 36
        }
      ]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            afterBody(items) {
              const i = items?.[0]?.dataIndex ?? -1;
              const row = topGuias[i];
              if (!row) return [];
              return [`Jabas: ${row.jabas}`, `Peso neto: ${row.pesoNeto || 0}`];
            }
          }
        }
      },
      scales: {
        x: { beginAtZero: true, title: { display: true, text: "Jarras" }, grid: { color: "rgba(20,90,52,0.08)" } },
        y: { grid: { display: false } }
      }
    }
  });

  const byMod = aggregateBy(rows, (r) => r.modulo).slice(0, 8);
  makeChart("modulo", $("chartTarjetasModulo"), {
    type: "bar",
    data: {
      labels: byMod.map((x) => x.label),
      datasets: [
        {
          label: "Jarras",
          data: byMod.map((x) => x.jarras),
          backgroundColor: "#27ae60",
          borderRadius: 6,
          maxBarThickness: 28
        }
      ]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { beginAtZero: true, title: { display: true, text: "Jarras" }, grid: { color: "rgba(20,90,52,0.08)" } },
        y: { grid: { display: false } }
      }
    }
  });

  const byVar = aggregateBy(rows, (r) => r.variedad).slice(0, 6);
  makeChart("variedad", $("chartTarjetasVariedad"), {
    type: "doughnut",
    data: {
      labels: byVar.map((x) => x.label),
      datasets: [
        {
          data: byVar.map((x) => x.jarras || x.filas),
          backgroundColor: CHART_PALETTE,
          borderWidth: 2,
          borderColor: "#fff"
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "58%",
      plugins: {
        legend: { position: "bottom", labels: { boxWidth: 12, boxHeight: 12 } },
        tooltip: {
          callbacks: {
            label(ctx) {
              const n = Number(ctx.raw) || 0;
              return ` ${ctx.label}: ${n} jarras`;
            }
          }
        }
      }
    }
  });

  const byDiaMap = new Map();
  rows.forEach((row) => {
    const day = String(row.fecha || "").trim() || "Sin fecha";
    if (!byDiaMap.has(day)) byDiaMap.set(day, { label: day, jarras: 0, jabas: 0, pesoNeto: 0 });
    const item = byDiaMap.get(day);
    item.jarras += Number(row.jarras) || 0;
    item.jabas += Number(row.jabas) || 0;
    item.pesoNeto += Number(row.pesoNeto) || 0;
  });
  const byDia = [...byDiaMap.values()].sort((a, b) => String(a.label).localeCompare(String(b.label)));

  makeChart("dia", $("chartTarjetasDia"), {
    type: "bar",
    data: {
      labels: byDia.map((x) => x.label),
      datasets: [
        {
          type: "bar",
          label: "Jarras",
          data: byDia.map((x) => x.jarras),
          backgroundColor: "#145a34",
          borderRadius: 6,
          yAxisID: "y",
          order: 2
        },
        {
          type: "bar",
          label: "Jabas",
          data: byDia.map((x) => x.jabas),
          backgroundColor: "#27ae60",
          borderRadius: 6,
          yAxisID: "y",
          order: 3
        },
        {
          type: "line",
          label: "Peso neto",
          data: byDia.map((x) => x.pesoNeto),
          borderColor: "#f6921e",
          backgroundColor: "rgba(246, 146, 30, 0.15)",
          tension: 0.3,
          pointRadius: 4,
          pointBackgroundColor: "#f6921e",
          yAxisID: "y1",
          order: 1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "top" } },
      scales: {
        x: { grid: { display: false } },
        y: {
          beginAtZero: true,
          position: "left",
          title: { display: true, text: "Jarras / Jabas" },
          grid: { color: "rgba(20,90,52,0.08)" }
        },
        y1: {
          beginAtZero: true,
          position: "right",
          title: { display: true, text: "Peso neto" },
          grid: { drawOnChartArea: false }
        }
      }
    }
  });
}

function openTarjetasResumen() {
  const modal = $("modalTarjetasResumen");
  if (!modal) return;
  if (!state.rows.length) {
    setStatus("empty", "No hay datos para el resumen. Carga o actualiza el listado.");
    return;
  }
  modal.hidden = false;
  window.setTimeout(() => renderResumenDashboard(), 30);
}

function closeTarjetasResumen() {
  const modal = $("modalTarjetasResumen");
  if (modal) modal.hidden = true;
  destroyResumenCharts();
}

function bindUi() {
  ["fltTarjetaPlaca", "fltTarjetaSearch"].forEach((id) => {
    $(id)?.addEventListener("input", () => applyFilters({ resetPage: true }));
    $(id)?.addEventListener("change", () => applyFilters({ resetPage: true }));
  });

  $("btnTarjetasResumen")?.addEventListener("click", () => openTarjetasResumen());
  $("btnCloseTarjetasResumen")?.addEventListener("click", closeTarjetasResumen);
  document.querySelectorAll('[data-close-modal="tarjeta-resumen"]').forEach((el) => {
    el.addEventListener("click", closeTarjetasResumen);
  });

  $("fltTarjetaFecha")?.addEventListener("change", () => {
    const value = ($("fltTarjetaFecha")?.value || "").trim();
    const today = limaTodayYmd();
    if (!value) state.fechaMode = "todas";
    else if (value === today) state.fechaMode = "hoy";
    else state.fechaMode = "custom";
    syncFechaModeButtons();
    loadTarjetas({ keepFilters: true, force: true });
  });

  $("btnTarjetasRefresh")?.addEventListener("click", () =>
    loadTarjetas({ keepFilters: true, force: true })
  );
  $("btnTarjetasExport")?.addEventListener("click", () => exportExcel());
  $("btnTarjetasHoy")?.addEventListener("click", () => {
    setFechaMode("hoy", limaTodayYmd());
    loadTarjetas({ keepFilters: true, force: true });
  });
  $("btnTarjetasClearFecha")?.addEventListener("click", () => {
    setFechaMode("todas", "");
    loadTarjetas({ keepFilters: true, force: true });
  });

  bindPager();

  $("tarjetasTableBody")?.addEventListener("click", (evt) => {
    const row = evt.target.closest("tr[data-index]");
    if (!row) return;
    openDetail(Number(row.dataset.index));
  });

  $("tarjetasTableBody")?.addEventListener("keydown", (evt) => {
    if (evt.key !== "Enter" && evt.key !== " ") return;
    const row = evt.target.closest("tr[data-index]");
    if (!row) return;
    evt.preventDefault();
    openDetail(Number(row.dataset.index));
  });

  $("btnCloseTarjetaDetail")?.addEventListener("click", closeDetail);
  document.querySelectorAll("[data-close-modal='tarjeta']").forEach((el) => {
    el.addEventListener("click", closeDetail);
  });

  document.addEventListener("keydown", (evt) => {
    if (evt.key !== "Escape") return;
    closeDetail();
    closeTarjetasResumen();
  });
}

function initTarjetasModule() {
  if (started) return;
  started = true;
  const fecha = $("fltTarjetaFecha");
  if (fecha && !fecha.value) {
    fecha.value = limaTodayYmd();
    state.fechaMode = "hoy";
  } else if (fecha?.value) {
    state.fechaMode = fecha.value === limaTodayYmd() ? "hoy" : "custom";
  }
  bindUi();
  syncFechaModeButtons();
}

function onRoute(route) {
  activeRoute = route === "tarjeta-pallet";
  if (!activeRoute) return;
  initTarjetasModule();
  // Sin poll: usa respaldo local; GET solo si no hay caché de esa fecha
  if (!state.loading) loadTarjetas({ keepFilters: true, force: false });
}

window.addEventListener("qb:route-changed", (evt) => {
  onRoute(evt.detail?.route);
});

if (window.location.hash.replace(/^#\/?/, "").split("/")[0] === "tarjeta-pallet") {
  activeRoute = true;
  initTarjetasModule();
  loadTarjetas({ keepFilters: true, force: false });
}
