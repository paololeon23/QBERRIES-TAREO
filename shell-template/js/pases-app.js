import {
  listarPermisos,
  limaTodayYmd,
  normalizeText
} from "./pases-api.js";

const DETAIL_FIELDS = [
  ["nombres", "Nombres"],
  ["dni", "DNI"],
  ["cargo", "Cargo"],
  ["motivo", "Motivo"],
  ["observacion", "Observación"],
  ["carnetVerificado", "Carnet verificado"],
  ["carnetVerificadoAt", "Hora carnet verificado"]
];

/** Solo columnas visibles / usadas en frontend (no device, sync, etc.) */
const EXPORT_COLUMNS = [
  ["nombres", "Nombres"],
  ["dni", "DNI"],
  ["cargo", "Cargo"],
  ["fechaRegistro", "Fecha registro"],
  ["horaRegistro", "H. registro"],
  ["motivo", "Motivo"],
  ["fechaSalida", "Fecha salida"],
  ["horaSalida", "H. salida"],
  ["responsable", "Responsable"],
  ["dniResponsable", "DNI responsable"],
  ["puestoResponsable", "Cargo responsable"],
  ["carnetVerificado", "Carnet"],
  ["observacion", "Observación"],
  ["carnetVerificadoAt", "H. carnet verificado"]
];

const state = {
  rows: [],
  filtered: [],
  loading: false,
  error: "",
  selectedIndex: -1,
  loadedAt: null,
  apiTotal: 0,
  apiMotivos: [],
  apiTopMotivo: { motivo: "", count: 0 }
};

const pagerState = {
  page: 1,
  pageSize: 10
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

function display(value) {
  const text = String(value ?? "").trim();
  return text || "—";
}

function dniCargoLine(dni, cargo) {
  const dniText = String(dni ?? "").trim() || "—";
  const cargoText = String(cargo ?? "").trim() || "—";
  return `DNI ${dniText} - (${cargoText})`;
}

function carnetIsSi(value) {
  return normalizeText(value) === "si";
}

function getFilters() {
  return {
    fecha: ($("fltPaseFecha")?.value || "").trim(),
    motivo: ($("fltPaseMotivo")?.value || "").trim(),
    responsable: ($("fltPaseResponsable")?.value || "").trim(),
    q: ($("fltPaseSearch")?.value || "").trim()
  };
}

function applyFilters({ resetPage = true } = {}) {
  syncSelectOptions();

  const f = getFilters();
  const q = normalizeText(f.q);
  const motivoFilter = normalizeText(f.motivo);
  const respFilter = normalizeText(f.responsable);

  // Fecha se filtra en el GET (params fecha / todas=1). Aquí solo motivo / responsable / búsqueda.
  state.filtered = state.rows.filter((row) => {
    if (motivoFilter) {
      if (normalizeText(row.motivo) !== motivoFilter) return false;
    }

    if (respFilter) {
      if (normalizeText(row.responsable) !== respFilter) return false;
    }

    if (q) {
      const qDigits = q.replace(/\D/g, "");
      const nombreOk =
        normalizeText(row.nombres).includes(q) || normalizeText(row.responsable).includes(q);
      const dniTrab = String(row.dni || "").replace(/\D/g, "");
      const dniResp = String(row.dniResponsable || "").replace(/\D/g, "");
      const dniOk =
        (qDigits && (dniTrab.includes(qDigits) || dniResp.includes(qDigits))) ||
        normalizeText(row.dni).includes(q) ||
        normalizeText(row.dniResponsable).includes(q);
      if (!nombreOk && !dniOk) return false;
    }

    return true;
  });

  if (resetPage) pagerState.page = 1;

  renderKpis();
  renderTable();
}

function uniqueSorted(values) {
  return [...new Set(values.map((v) => String(v || "").trim()).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, "es")
  );
}

function syncSelectOptions() {
  const motivoSel = $("fltPaseMotivo");
  const respSel = $("fltPaseResponsable");
  if (!motivoSel || !respSel) return;

  const currentMotivo = motivoSel.value;
  const currentResp = respSel.value;

  const motivos = uniqueSorted(state.rows.map((r) => r.motivo));
  const responsables = uniqueSorted(state.rows.map((r) => r.responsable));

  motivoSel.innerHTML =
    `<option value="">Todos</option>` +
    motivos.map((m) => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join("");
  respSel.innerHTML =
    `<option value="">Todos</option>` +
    responsables.map((r) => `<option value="${escapeHtml(r)}">${escapeHtml(r)}</option>`).join("");

  if (motivos.includes(currentMotivo)) motivoSel.value = currentMotivo;
  if (responsables.includes(currentResp)) respSel.value = currentResp;
}

function countToday(rows) {
  const today = limaTodayYmd();
  return rows.filter((r) => String(r.fechaRegistro || "").trim() === today).length;
}

function topMotivoFromRows(rows) {
  const counts = new Map();
  rows.forEach((r) => {
    const key = String(r.motivo || "").trim();
    if (!key) return;
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  let best = "";
  let bestN = 0;
  counts.forEach((n, key) => {
    if (n > bestN || (n === bestN && key.localeCompare(best, "es") < 0)) {
      best = key;
      bestN = n;
    }
  });
  return { motivo: best, count: bestN };
}

function renderKpis() {
  const source = state.rows;
  const todayCount = countToday(source);
  const conCarnet = source.filter((r) => carnetIsSi(r.carnetVerificado)).length;

  // Total y motivo: preferir agregados del GET (Code.gs). Fallback local si aún no vienen.
  const total = Number.isFinite(Number(state.apiTotal)) && state.apiTotal >= 0
    ? Number(state.apiTotal)
    : source.length;

  const top =
    state.apiTopMotivo?.motivo
      ? state.apiTopMotivo
      : state.apiMotivos?.length
        ? [...state.apiMotivos].sort(
            (a, b) => b.count - a.count || String(a.motivo).localeCompare(String(b.motivo), "es")
          )[0]
        : topMotivoFromRows(source);

  const elHoy = $("kpiPasesHoy");
  const elSi = $("kpiPasesCarnetSi");
  const elTotal = $("kpiPasesTotal");
  const elTop = $("kpiPasesTopMotivo");

  if (elHoy) elHoy.textContent = String(todayCount);
  if (elSi) elSi.textContent = String(conCarnet);
  if (elTotal) elTotal.textContent = String(total);
  if (elTop) {
    const n = Number(top?.count) || 0;
    const nombre = top?.motivo || "";
    elTop.textContent = nombre
      ? n
        ? `${nombre} - ${n} pase${n === 1 ? "" : "s"}`
        : nombre
      : "—";
  }
}

function setStatus(kind, message) {
  const el = $("pasesStatus");
  if (!el) return;
  el.hidden = !message;
  el.dataset.kind = kind || "";
  el.textContent = message || "";
}

function motivoBadge(motivo) {
  const text = display(motivo);
  return `<span class="pase-badge pase-badge--motivo" title="${escapeHtml(text)}">${escapeHtml(text)}</span>`;
}

function carnetBadge(value) {
  const si = carnetIsSi(value);
  const label = si ? "SI" : "NO";
  const cls = si ? "pase-badge--ok" : "pase-badge--no";
  return `<span class="pase-badge ${cls}">Carnet ${label}</span>`;
}

function renderPagerControls() {
  const rangeEl = $("pasesPagerRange");
  const first = $("pasesPagerFirst");
  const prev = $("pasesPagerPrev");
  const next = $("pasesPagerNext");
  const last = $("pasesPagerLast");
  const sizeSel = $("pasesPagerPageSize");
  const pager = $("pasesTablePager");

  if (pager) pager.hidden = false;

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
  const total = state.filtered.length;
  const totalPages = Math.max(1, Math.ceil(total / pagerState.pageSize) || 1);
  if (pagerState.page > totalPages) pagerState.page = totalPages;
  const start = (pagerState.page - 1) * pagerState.pageSize;
  return state.filtered.slice(start, start + pagerState.pageSize);
}

function renderTable() {
  const tbody = $("pasesTableBody");
  const empty = $("pasesEmpty");
  if (!tbody) return;

  renderPagerControls();
  const pageRows = getPageRows();

  if (!state.filtered.length) {
    tbody.innerHTML = "";
    if (empty) {
      empty.hidden = state.loading || Boolean(state.error);
      empty.textContent = state.loading
        ? ""
        : state.error
          ? ""
          : state.rows.length
            ? "Ningún pase coincide con los filtros."
            : "No hay pases para mostrar.";
    }
    return;
  }

  if (empty) empty.hidden = true;

  tbody.innerHTML = pageRows
    .map((row, index) => {
      const absoluteIndex = (pagerState.page - 1) * pagerState.pageSize + index;
      return `<tr class="pases-table__row" data-index="${absoluteIndex}" tabindex="0">
        <td>
          <div class="pases-table__primary">${escapeHtml(display(row.nombres))}</div>
          <div class="pases-table__meta">${escapeHtml(dniCargoLine(row.dni, row.cargo))}</div>
        </td>
        <td>
          <div>${escapeHtml(display(row.fechaRegistro))}</div>
          <div class="pases-table__meta">${escapeHtml(display(row.horaRegistro))}</div>
        </td>
        <td>${motivoBadge(row.motivo)}</td>
        <td>
          <div>${escapeHtml(display(row.fechaSalida))}</div>
          <div class="pases-table__meta">${escapeHtml(display(row.horaSalida))}</div>
        </td>
        <td>
          <div class="pases-table__primary">${escapeHtml(display(row.responsable))}</div>
          <div class="pases-table__meta">${escapeHtml(dniCargoLine(row.dniResponsable, row.puestoResponsable))}</div>
        </td>
        <td>${carnetBadge(row.carnetVerificado)}</td>
      </tr>`;
    })
    .join("");
}

function openDetail(index) {
  const row = state.filtered[index];
  const modal = $("modalPaseDetail");
  const body = $("paseDetailBody");
  const title = $("paseDetailTitle");
  if (!row || !modal || !body) return;

  state.selectedIndex = index;
  if (title) title.textContent = row.nombres ? String(row.nombres) : `DNI ${display(row.dni)}`;

  body.innerHTML = DETAIL_FIELDS.map(([key, label]) => {
    const raw = row[key];
    let valueHtml = escapeHtml(display(raw));
    if (key === "carnetVerificado") valueHtml = carnetBadge(raw);
    if (key === "motivo") valueHtml = motivoBadge(raw);
    return `<div class="pase-detail__item">
      <dt>${escapeHtml(label)}</dt>
      <dd>${valueHtml}</dd>
    </div>`;
  }).join("");

  modal.hidden = false;
}

function closeDetail() {
  const modal = $("modalPaseDetail");
  if (modal) modal.hidden = true;
  state.selectedIndex = -1;
}

function exportExcel() {
  const rowsSrc = state.filtered.length ? state.filtered : state.rows;
  if (!rowsSrc.length) {
    setStatus("empty", "No hay datos para exportar.");
    return;
  }

  const XLSX = window.XLSX;
  if (!XLSX?.utils || typeof XLSX.write !== "function") {
    setStatus("error", "No se pudo cargar el exportador Excel (XLSX).");
    return;
  }

  try {
    const rows = rowsSrc.map((row) => {
      const out = {};
      EXPORT_COLUMNS.forEach(([key, label]) => {
        const val = row[key];
        out[label] = val == null || val === "" ? "" : String(val);
      });
      return out;
    });

    const sheet = XLSX.utils.json_to_sheet(rows);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, "Pases");

    const stamp = limaTodayYmd();
    const fileName = `pases-salida-qberries-${stamp}.xlsx`;

    // Descarga manual: más estable que writeFile en algunos navegadores / Live Server
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
    console.error("[pases] export Excel", err);
    setStatus("error", err?.message || "Error al exportar Excel.");
  }
}

const POLL_MS = 3000;

let started = false;
let pollTimer = null;
let activeRoute = false;

function stopPolling() {
  if (pollTimer) {
    window.clearInterval(pollTimer);
    pollTimer = null;
  }
}

function startPolling() {
  stopPolling();
  pollTimer = window.setInterval(() => {
    if (!activeRoute || state.loading) return;
    const modal = $("modalPaseDetail");
    if (modal && !modal.hidden) return;
    loadPases({ keepFilters: true, silent: true });
  }, POLL_MS);
}

async function loadPases({ keepFilters = true, silent = false } = {}) {
  if (state.loading) return;

  const btn = $("btnPasesRefresh");
  state.loading = true;
  if (!silent) {
    state.error = "";
    setStatus("loading", "Cargando pases desde BD-PERMISOS…");
    if (btn) btn.disabled = true;
  }

  try {
    const fechaInput = keepFilters
      ? ($("fltPaseFecha")?.value || "").trim()
      : limaTodayYmd();
    if (!keepFilters) {
      const fechaEl = $("fltPaseFecha");
      if (fechaEl) fechaEl.value = fechaInput;
    }

    // Code.gs: día concreto con fecha=… ; vacío → todas=1
    const resp = await listarPermisos(
      fechaInput ? { fecha: fechaInput } : { todas: true }
    );
    state.rows = resp.data || [];
    state.apiTotal = Number(resp.kpis?.total ?? resp.count ?? state.rows.length) || 0;
    state.apiMotivos = Array.isArray(resp.kpis?.motivos) ? resp.kpis.motivos : [];
    state.apiTopMotivo = resp.kpis?.topMotivo || { motivo: "", count: 0 };
    state.loadedAt = new Date();
    state.loading = false;
    state.error = "";

    // En auto-refresh conservar página; en carga manual/filtro resetear
    applyFilters({ resetPage: !silent });
    const meta = $("pasesMeta");
    if (meta) {
      const ts = state.loadedAt.toLocaleString("es-PE", { timeZone: "America/Lima" });
      const n = Number(resp.count ?? state.rows.length) || state.rows.length;
      const rango =
        resp.rango?.modo === "todas" || !fechaInput
          ? "todas las fechas"
          : `día ${fechaInput}`;
      meta.textContent = `${n} pases · ${rango} · actualizado ${ts} (hora Lima)`;
    }
    if (!silent) {
      setStatus(
        state.filtered.length ? "" : "empty",
        state.filtered.length
          ? ""
          : state.rows.length
            ? "Ningún pase coincide con los filtros."
            : "No hay pases registrados."
      );
    } else if (!state.filtered.length && !state.rows.length) {
      setStatus("empty", "No hay pases registrados.");
    } else if (state.filtered.length) {
      setStatus("", "");
    }
  } catch (err) {
    state.loading = false;
    const msg = err?.message || "Error al cargar pases";
    if (silent && state.rows.length) {
      // Mantener datos anteriores en auto-refresh
    } else {
      state.error = msg;
      state.rows = [];
      state.filtered = [];
      state.apiTotal = 0;
      state.apiMotivos = [];
      state.apiTopMotivo = { motivo: "", count: 0 };
      renderKpis();
      renderTable();
      setStatus("error", state.error);
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}

function bindPager() {
  $("pasesPagerPageSize")?.addEventListener("change", (e) => {
    const n = Number(e.target.value);
    pagerState.pageSize = Number.isFinite(n) && n > 0 ? n : 10;
    pagerState.page = 1;
    renderTable();
  });
  $("pasesPagerFirst")?.addEventListener("click", () => {
    pagerState.page = 1;
    renderTable();
  });
  $("pasesPagerPrev")?.addEventListener("click", () => {
    pagerState.page = Math.max(1, pagerState.page - 1);
    renderTable();
  });
  $("pasesPagerNext")?.addEventListener("click", () => {
    const totalPages = Math.max(1, Math.ceil(state.filtered.length / pagerState.pageSize) || 1);
    pagerState.page = Math.min(totalPages, pagerState.page + 1);
    renderTable();
  });
  $("pasesPagerLast")?.addEventListener("click", () => {
    const totalPages = Math.max(1, Math.ceil(state.filtered.length / pagerState.pageSize) || 1);
    pagerState.page = totalPages;
    renderTable();
  });
}

function bindUi() {
  ["fltPaseMotivo", "fltPaseResponsable", "fltPaseSearch"].forEach((id) => {
    $(id)?.addEventListener("input", () => applyFilters({ resetPage: true }));
    $(id)?.addEventListener("change", () => applyFilters({ resetPage: true }));
  });

  // Cambio de fecha → nuevo GET (hoy / día / todas)
  $("fltPaseFecha")?.addEventListener("change", () => {
    loadPases({ keepFilters: true, silent: false });
  });

  $("btnPasesRefresh")?.addEventListener("click", () => loadPases({ keepFilters: true, silent: false }));
  $("btnPasesExport")?.addEventListener("click", () => exportExcel());
  $("btnPasesHoy")?.addEventListener("click", () => {
    const fecha = $("fltPaseFecha");
    if (fecha) fecha.value = limaTodayYmd();
    loadPases({ keepFilters: true, silent: false });
  });
  $("btnPasesClearFecha")?.addEventListener("click", () => {
    const fecha = $("fltPaseFecha");
    if (fecha) fecha.value = "";
    loadPases({ keepFilters: true, silent: false });
  });

  bindPager();

  $("pasesTableBody")?.addEventListener("click", (evt) => {
    const row = evt.target.closest("tr[data-index]");
    if (!row) return;
    openDetail(Number(row.dataset.index));
  });

  $("pasesTableBody")?.addEventListener("keydown", (evt) => {
    if (evt.key !== "Enter" && evt.key !== " ") return;
    const row = evt.target.closest("tr[data-index]");
    if (!row) return;
    evt.preventDefault();
    openDetail(Number(row.dataset.index));
  });

  $("btnClosePaseDetail")?.addEventListener("click", closeDetail);
  document.querySelectorAll("[data-close-modal='pase']").forEach((el) => {
    el.addEventListener("click", closeDetail);
  });

  document.addEventListener("keydown", (evt) => {
    if (evt.key === "Escape") closeDetail();
  });
}

function initPasesModule() {
  if (started) return;
  started = true;
  const fecha = $("fltPaseFecha");
  if (fecha && !fecha.value) fecha.value = limaTodayYmd();
  bindUi();
}

function onRoute(route) {
  activeRoute = route === "pases-de-salida";
  if (!activeRoute) {
    stopPolling();
    return;
  }
  initPasesModule();
  startPolling();
  if (!state.loading) {
    loadPases({ keepFilters: true, silent: Boolean(state.rows.length) });
  }
}

window.addEventListener("qb:route-changed", (evt) => {
  onRoute(evt.detail?.route);
});

document.addEventListener("visibilitychange", () => {
  if (document.hidden) {
    stopPolling();
    return;
  }
  if (activeRoute) {
    startPolling();
    if (!state.loading) loadPases({ keepFilters: true, silent: true });
  }
});

if (window.location.hash.replace(/^#\/?/, "").split("/")[0] === "pases-de-salida") {
  activeRoute = true;
  initPasesModule();
  startPolling();
  loadPases({ keepFilters: true, silent: false });
}
