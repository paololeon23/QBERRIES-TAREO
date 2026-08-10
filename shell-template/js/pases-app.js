import {
  listarPermisos,
  peekCachedListarPermisos,
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
  apiTopMotivo: { motivo: "", count: 0 },
  apiPorDia: [],
  apiRango: null,
  fromCache: false,
  dataSig: "",
  /** 'hoy' | 'todas' | 'custom' — controla el pintado de botones */
  fechaMode: "hoy"
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

  // Fecha/todas van en el GET. Aquí: motivo, responsable, búsqueda/DNI.
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

function renderKpis() {
  const source = state.rows;
  const today = limaTodayYmd();
  const todayCount =
    state.apiRango?.modo === "dia" && state.apiRango?.fecha === today
      ? Number(state.apiTotal) || source.length
      : countToday(source);
  const conCarnet = source.filter((r) => carnetIsSi(r.carnetVerificado)).length;

  const total = Number.isFinite(Number(state.apiTotal)) && state.apiTotal >= 0
    ? Number(state.apiTotal)
    : source.length;

  const motivos =
    Array.isArray(state.apiMotivos) && state.apiMotivos.length
      ? [...state.apiMotivos].sort(
          (a, b) => b.count - a.count || String(a.motivo).localeCompare(String(b.motivo), "es")
        )
      : (() => {
          const map = new Map();
          source.forEach((r) => {
            const key = String(r.motivo || "").trim();
            if (!key) return;
            map.set(key, (map.get(key) || 0) + 1);
          });
          return [...map.entries()]
            .map(([motivo, count]) => ({ motivo, count }))
            .sort((a, b) => b.count - a.count || a.motivo.localeCompare(b.motivo, "es"));
        })();

  const top =
    state.apiTopMotivo?.motivo
      ? state.apiTopMotivo
      : motivos[0] || null;

  const elHoy = $("kpiPasesHoy");
  const elSi = $("kpiPasesCarnetSi");
  const elTotal = $("kpiPasesTotal");
  const elMotivos = $("kpiPasesMotivos");

  if (elHoy) elHoy.textContent = String(todayCount);
  if (elSi) elSi.textContent = String(conCarnet);
  if (elTotal) elTotal.textContent = String(total);
  if (elMotivos) {
    if (!top?.motivo) {
      elMotivos.textContent = "—";
      elMotivos.classList.add("is-empty");
    } else {
      const n = Number(top.count) || 0;
      elMotivos.classList.remove("is-empty");
      elMotivos.innerHTML = `<p class="pases-kpi__motif-name">${escapeHtml(top.motivo)} - ${n}</p>`;
    }
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
      const textEl = empty.querySelector(".pases-empty__text");
      const titleEl = empty.querySelector(".pases-empty__title");
      const msg = state.rows.length
        ? "Ningún pase coincide con los filtros."
        : "No hay pases para mostrar.";
      const title = state.rows.length ? "Sin coincidencias" : "Sin pases";
      if (textEl) textEl.textContent = msg;
      else empty.textContent = msg;
      if (titleEl) titleEl.textContent = title;
      empty.dataset.kind = state.rows.length ? "filtered" : "empty";
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

const POLL_MS = 15000;

let started = false;
let pollTimer = null;
let activeRoute = false;
/** Descarta respuestas viejas si el usuario cambia Hoy/Todas a mitad del GET */
let loadSeq = 0;

function rowsSignature(rows) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return "0";
  // Huella liviana: si no cambia, no re-pintamos en el poll
  return list
    .map(
      (r) =>
        `${r.dni}|${r.fechaRegistro}|${r.horaRegistro}|${r.motivo}|${r.fechaSalida}|${r.horaSalida}|${r.carnetVerificado}`
    )
    .join(";");
}

function syncFechaModeButtons() {
  const mode = state.fechaMode || "hoy";
  const isHoy = mode === "hoy";
  const isTodas = mode === "todas";
  const btnHoy = $("btnPasesHoy");
  const btnTodas = $("btnPasesClearFecha");

  if (btnHoy) {
    btnHoy.classList.toggle("is-active", isHoy);
    btnHoy.classList.toggle("btn--secondary", isHoy);
    btnHoy.classList.toggle("btn--ghost", !isHoy);
    btnHoy.setAttribute("aria-pressed", isHoy ? "true" : "false");
  }
  if (btnTodas) {
    btnTodas.classList.toggle("is-active", isTodas);
    btnTodas.classList.toggle("btn--secondary", isTodas);
    btnTodas.classList.toggle("btn--ghost", !isTodas);
    btnTodas.setAttribute("aria-pressed", isTodas ? "true" : "false");
  }
}

function setFechaMode(mode, fechaValue) {
  state.fechaMode = mode;
  const fecha = $("fltPaseFecha");
  if (fecha) fecha.value = fechaValue == null ? "" : String(fechaValue);
  syncFechaModeButtons();
}

function applyApiPayload(resp, { resetPage = true, touchMeta = true } = {}) {
  state.rows = resp.data || [];
  state.apiTotal = Number(resp.kpis?.total ?? resp.count ?? state.rows.length) || 0;
  state.apiMotivos = Array.isArray(resp.kpis?.motivos) ? resp.kpis.motivos : [];
  state.apiTopMotivo = resp.kpis?.topMotivo || { motivo: "", count: 0 };
  state.apiPorDia = Array.isArray(resp.kpis?.porDia) ? resp.kpis.porDia : [];
  state.apiRango = resp.rango || resp.kpis?.rango || null;
  state.fromCache = Boolean(resp.fromCache);
  state.loadedAt = resp.cachedAt ? new Date(resp.cachedAt) : new Date();
  state.dataSig = rowsSignature(state.rows);
  state.error = "";
  applyFilters({ resetPage });
  syncFechaModeButtons();

  if (!touchMeta) return;
  const meta = $("pasesMeta");
  if (!meta) return;
  const fechaInput = ($("fltPaseFecha")?.value || "").trim();
  const ts = state.loadedAt.toLocaleString("es-PE", { timeZone: "America/Lima" });
  const n = Number(resp.count ?? state.rows.length) || state.rows.length;
  // El texto sigue al modo del UI (botones), no al rango que mande la API
  const rango =
    state.fechaMode === "todas" || !fechaInput
      ? "todas las fechas"
      : `día ${fechaInput || limaTodayYmd()}`;
  const cacheNote = state.fromCache ? " · caché local" : "";
  meta.textContent = `${n} pases · ${rango} · actualizado ${ts} (hora Lima)${cacheNote}`;
}

function listarOptsFromUi(keepFilters) {
  const fechaInput = keepFilters
    ? ($("fltPaseFecha")?.value || "").trim()
    : limaTodayYmd();
  if (!keepFilters) {
    const fechaEl = $("fltPaseFecha");
    if (fechaEl) fechaEl.value = fechaInput;
  }
  return {
    fechaInput,
    opts: fechaInput ? { fecha: fechaInput } : { todas: true }
  };
}

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
  // Poll no pisa un GET del usuario; clics del usuario SÍ pueden cancelar el anterior
  if (silent && state.loading) return;

  const seq = ++loadSeq;
  const btn = $("btnPasesRefresh");
  const { fechaInput, opts } = listarOptsFromUi(keepFilters);

  if (!silent) {
    syncFechaModeButtons();
    // Pintado inmediato desde caché de ese modo (Hoy / Todas) — no se cuelga esperando
    const peeked = peekCachedListarPermisos(opts);
    if (peeked) {
      applyApiPayload(peeked, { resetPage: true, touchMeta: true });
      setStatus(
        "loading",
        fechaInput ? "Actualizando día seleccionado…" : "Cargando todas las fechas…"
      );
    } else {
      state.error = "";
      setStatus(
        "loading",
        fechaInput ? "Cargando pases del día…" : "Cargando todas las fechas…"
      );
    }
    if (btn) btn.disabled = true;
  }

  state.loading = true;

  try {
    const resp = await listarPermisos(silent ? opts : { ...opts, force: true });
    if (seq !== loadSeq) return;

    const nextSig = rowsSignature(resp.data || []);
    if (silent && nextSig === state.dataSig && state.rows.length) {
      state.fromCache = Boolean(resp.fromCache);
      return;
    }

    applyApiPayload(resp, { resetPage: !silent, touchMeta: true });

    if (state.fromCache) {
      if (!silent) {
        setStatus(
          "empty",
          "Mostrando último listado guardado (caché). Se actualizará al recuperar la conexión."
        );
      }
    } else if (!silent) {
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
    if (seq !== loadSeq) return;
    const msg = err?.message || "Error al cargar pases";
    if (state.rows.length) {
      if (!silent) setStatus("error", msg);
    } else {
      state.error = msg;
      state.rows = [];
      state.filtered = [];
      state.apiTotal = 0;
      state.apiMotivos = [];
      state.apiTopMotivo = { motivo: "", count: 0 };
      state.apiPorDia = [];
      state.apiRango = null;
      state.dataSig = "";
      renderKpis();
      renderTable();
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
    const value = ($("fltPaseFecha")?.value || "").trim();
    const today = limaTodayYmd();
    if (!value) state.fechaMode = "todas";
    else if (value === today) state.fechaMode = "hoy";
    else state.fechaMode = "custom";
    syncFechaModeButtons();
    loadPases({ keepFilters: true, silent: false });
  });

  $("btnPasesRefresh")?.addEventListener("click", () => loadPases({ keepFilters: true, silent: false }));
  $("btnPasesExport")?.addEventListener("click", () => exportExcel());
  $("btnPasesHoy")?.addEventListener("click", () => {
    setFechaMode("hoy", limaTodayYmd());
    loadPases({ keepFilters: true, silent: false });
  });
  $("btnPasesClearFecha")?.addEventListener("click", () => {
    setFechaMode("todas", "");
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
