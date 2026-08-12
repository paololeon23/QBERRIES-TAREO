/**
 * Módulo Producción: Excel Grupo (F) · DNI/CI (H) · Jarras/C (Q)
 * + Listado trabajadores · filtros · export · resumen mejores/peores
 */

import {
  parseListadoBuffer,
  parseProduccionBuffer,
  aggregateWorkers
} from "./produccion-parser.js";

const LISTADO_KEY = "qb-produccion-listado-v1";
const LISTADO_JSON_URL = "./data/listado-trabajadores.json";

const state = {
  prodFile: "",
  listadoFile: "",
  listadoSource: "", // "json" | "upload" | "cache"
  workerMap: new Map(),
  rows: [],
  workers: [],
  filtered: [],
  grupos: [],
  fechas: [],
  selectedGrupos: [],
  rankingMode: "all", // all | top10 | bottom10
  meta: { registros: 0, trabajadores: 0, jarras: 0 }
};

const pager = { page: 1, pageSize: 20, bound: false };
const charts = {};
const filterModal = { options: [], draft: new Set(), bound: false };

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

function fmt(n) {
  return Number(n || 0).toLocaleString("es-PE");
}

function setStatus(kind, text) {
  const el = $("prodStatus");
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

function syncListadoUi() {
  const label = $("prodListadoMeta");
  const inline = $("prodListadoMetaInline");
  const text = !state.workerMap.size
    ? "cargando listado…"
    : `${fmt(state.workerMap.size)} en JSON`;
  if (label) label.textContent = text;
  if (inline) inline.textContent = text;
  const badge = $("prodKpiListado");
  if (badge) badge.textContent = state.workerMap.size ? fmt(state.workerMap.size) : "—";
}

function applyWorkerMap(map, fileName, source) {
  state.workerMap = map instanceof Map ? map : new Map(map);
  state.listadoFile = fileName || "Listado trabajadores";
  state.listadoSource = source || "";
  syncListadoUi();
}

function mapFromJsonTrabajadores(list) {
  const map = new Map();
  (list || []).forEach((row) => {
    const dni = String(row.dni || "").replace(/\D+/g, "").replace(/^0+/, "") || String(row.dni || "").replace(/\D+/g, "");
    const nombre = String(row.nombre || "").trim();
    if (!dni || !nombre) return;
    if (!map.has(dni)) map.set(dni, nombre);
  });
  return map;
}

async function loadPackedListado() {
  try {
    const res = await fetch(LISTADO_JSON_URL, { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    const map = mapFromJsonTrabajadores(data.trabajadores || data.people || []);
    if (!map.size) throw new Error("JSON de listado vacío");
    applyWorkerMap(map, data.source || "listado-trabajadores.json", "json");
    return true;
  } catch (err) {
    console.warn("[produccion] listado JSON", err);
    return false;
  }
}

function loadSavedListado() {
  try {
    const raw = localStorage.getItem(LISTADO_KEY);
    if (!raw) return false;
    const data = JSON.parse(raw);
    if (!data?.entries?.length) return false;
    applyWorkerMap(new Map(data.entries), data.fileName || "Listado guardado", "cache");
    return true;
  } catch {
    return false;
  }
}

function showWorkspace(_show) {
  // Layout estilo Pases: siempre visible; solo cambia empty / botones.
  const newBtn = $("btnProdNewFile");
  if (newBtn) newBtn.hidden = !state.rows.length;
}

function formatMultiLabel(selected, placeholder) {
  if (!selected.length) return placeholder;
  if (selected.length === 1) return selected[0];
  return `${selected.length} grupos`;
}

function syncGrupoButton() {
  const btn = $("prodBtnGrupo");
  const txt = $("prodBtnGrupoText");
  const has = state.rows.length > 0;
  if (btn) btn.disabled = !has;
  if (txt) {
    txt.textContent = formatMultiLabel(
      state.selectedGrupos,
      btn?.dataset.placeholder || "Todos los grupos…"
    );
  }
  btn?.classList.toggle("has-value", state.selectedGrupos.length > 0);
}

function grupoKey(value) {
  return String(value || "").trim() || "(sin grupo)";
}

/** Ranking de grupos por total de jarras (desc). */
function rankGruposByJarras(rows) {
  const map = new Map();
  (rows || []).forEach((row) => {
    const g = grupoKey(row.grupo);
    map.set(g, (map.get(g) || 0) + (Number(row.jarras) || 0));
  });
  return [...map.entries()]
    .map(([grupo, jarras]) => ({ grupo, jarras: Math.round(jarras * 1000) / 1000 }))
    .sort((a, b) => b.jarras - a.jarras || a.grupo.localeCompare(b.grupo, "es"));
}

function rankingGrupoSet(rows, mode = state.rankingMode) {
  if (mode !== "top10" && mode !== "bottom10") return null;
  const ranked = rankGruposByJarras(rows);
  if (!ranked.length) return new Set();
  const slice =
    mode === "top10" ? ranked.slice(0, 10) : ranked.slice(Math.max(0, ranked.length - 10));
  return new Set(slice.map((r) => r.grupo));
}

function syncRankingButtons() {
  ["btnProdRankAll", "btnProdRankTop", "btnProdRankBottom"].forEach((id) => {
    const btn = $(id);
    if (!btn) return;
    const active = btn.dataset.rank === state.rankingMode;
    btn.classList.toggle("is-active", active);
    btn.classList.toggle("btn--secondary", active);
    btn.classList.toggle("btn--ghost", !active);
  });
}

function setRankingMode(mode) {
  state.rankingMode = mode === "top10" || mode === "bottom10" ? mode : "all";
  syncRankingButtons();
  applyFilters({ resetPage: true });
}

function applyFilters({ resetPage = true } = {}) {
  const search = ($("prodFltSearch")?.value || "").trim().toLowerCase();
  const fecha = $("prodFltFecha")?.value || "";
  const grupoSet = new Set(state.selectedGrupos);

  const baseForRank = state.rows.filter((row) => {
    if (grupoSet.size && !grupoSet.has(row.grupo)) return false;
    if (fecha && row.fecha !== fecha) return false;
    return true;
  });
  const rankSet = rankingGrupoSet(baseForRank);

  const scoped = state.rows.filter((row) => {
    if (grupoSet.size && !grupoSet.has(row.grupo)) return false;
    if (fecha && row.fecha !== fecha) return false;
    if (rankSet && !rankSet.has(grupoKey(row.grupo))) return false;
    if (search) {
      const blob = `${row.dni} ${row.trabajador} ${row.grupo}`.toLowerCase();
      if (!blob.includes(search)) return false;
    }
    return true;
  });

  state.filtered = aggregateWorkers(scoped);
  if (resetPage) pager.page = 1;
  renderKpis(scoped);
  renderMeta(scoped);
  renderTable();
  syncExport();
  syncGrupoButton();
}

function renderKpis(scopedRows) {
  const jarras = scopedRows.reduce((s, r) => s + (Number(r.jarras) || 0), 0);
  if ($("prodKpiFile")) $("prodKpiFile").textContent = state.prodFile || "—";
  if ($("prodKpiGrupos")) {
    $("prodKpiGrupos").textContent = String(
      state.selectedGrupos.length || state.grupos.length || 0
    );
  }
  if ($("prodKpiTrabajadores")) $("prodKpiTrabajadores").textContent = fmt(state.filtered.length);
  if ($("prodKpiJarras")) $("prodKpiJarras").textContent = fmt(Math.round(jarras * 1000) / 1000);
  syncListadoUi();
}

function renderMeta(scopedRows) {
  const meta = $("prodMeta");
  if (!meta) return;
  const jarras = scopedRows.reduce((s, r) => s + (Number(r.jarras) || 0), 0);
  const gLabel = state.selectedGrupos.length
    ? state.selectedGrupos.length <= 2
      ? state.selectedGrupos.join(", ")
      : `${state.selectedGrupos.length} grupos`
    : "Todos los grupos";
  const rankLabel =
    state.rankingMode === "top10"
      ? "· Mejores 10 grupos"
      : state.rankingMode === "bottom10"
        ? "· Peores 10 grupos"
        : "";
  meta.textContent = `${fmt(state.filtered.length)} trabajadores · ${fmt(
    Math.round(jarras * 1000) / 1000
  )} jarras · ${gLabel}${rankLabel} · ${fmt(scopedRows.length)} registros`;
}

function renderPager() {
  const total = state.filtered.length;
  const totalPages = Math.max(1, Math.ceil(total / pager.pageSize) || 1);
  if (pager.page > totalPages) pager.page = totalPages;
  const start = total ? (pager.page - 1) * pager.pageSize + 1 : 0;
  const end = Math.min(total, pager.page * pager.pageSize);
  const range = $("prodPagerRange");
  if (range) range.textContent = `${start} – ${end} of ${total}`;
  const atStart = pager.page <= 1;
  const atEnd = pager.page >= totalPages;
  ["prodPagerFirst", "prodPagerPrev"].forEach((id) => {
    const btn = $(id);
    if (btn) btn.disabled = atStart;
  });
  ["prodPagerNext", "prodPagerLast"].forEach((id) => {
    const btn = $(id);
    if (btn) btn.disabled = atEnd;
  });
}

function renderTable() {
  const body = $("prodTableBody");
  const empty = $("prodEmpty");
  if (!body) return;
  renderPager();
  const start = (pager.page - 1) * pager.pageSize;
  const pageRows = state.filtered.slice(start, start + pager.pageSize);

  if (!pageRows.length) {
    body.innerHTML = "";
    if (empty) {
      empty.hidden = false;
      const title = empty.querySelector(".pases-empty__title, .prod-empty__title");
      const text = empty.querySelector(".pases-empty__text, .prod-empty__text");
      if (!state.rows.length) {
        if (title) title.textContent = "Sin producción cargada";
        if (text) text.textContent = "Sube el Excel de producción para analizar jarras por trabajador.";
      } else {
        if (title) title.textContent = "Sin resultados";
        if (text) text.textContent = "Prueba otro grupo, fecha o búsqueda.";
      }
    }
    return;
  }

  if (empty) empty.hidden = true;
  body.innerHTML = pageRows
    .map((row, i) => {
      const n = start + i + 1;
      return `<tr>
        <td class="prod-table__num">${n}</td>
        <td class="prod-table__doc">${escapeHtml(row.dni)}</td>
        <td><div class="prod-table__primary">${escapeHtml(row.trabajador)}</div></td>
        <td title="${escapeHtml(row.grupo)}">${escapeHtml(row.grupo || "—")}</td>
        <td class="prod-table__num prod-table__jarras">${fmt(row.jarras)}</td>
        <td class="prod-table__num">${fmt(row.registros)}</td>
      </tr>`;
    })
    .join("");
}

function syncExport() {
  const btn = $("btnProdExport");
  if (btn) btn.disabled = !state.filtered.length;
  const resumen = $("btnProdResumen");
  if (resumen) resumen.disabled = !state.filtered.length;
}

function fillFechaOptions() {
  const sel = $("prodFltFecha");
  if (!sel) return;
  const prev = sel.value;
  const fechas = sortedFechas();
  sel.innerHTML = `<option value="">Todas</option>${fechas
    .map((f) => `<option value="${escapeHtml(f)}">${escapeHtml(f)}</option>`)
    .join("")}`;
  sel.value = fechas.includes(prev) ? prev : "";
}

function saveListado(fileName, map) {
  try {
    localStorage.setItem(
      LISTADO_KEY,
      JSON.stringify({
        fileName,
        savedAt: new Date().toISOString(),
        entries: [...map.entries()]
      })
    );
  } catch {
    /* quota */
  }
}

async function handleListado(file) {
  if (!file) return;
  try {
    setStatus("loading", "Leyendo listado de trabajadores…");
    const buffer = await file.arrayBuffer();
    const parsed = parseListadoBuffer(buffer, file.name);
    applyWorkerMap(parsed.map, file.name, "upload");
    saveListado(file.name, parsed.map);
    setStatus("", "");
    if (state.rows.length) {
      state.rows.forEach((row) => {
        const name = state.workerMap.get(row.dni);
        if (name) row.trabajador = name;
      });
      applyFilters({ resetPage: false });
    }
  } catch (err) {
    console.error("[produccion] listado", err);
    setStatus("error", err?.message || "No se pudo leer el listado.");
  }
}

async function handleProduccion(file) {
  if (!file) return;
  const nameLower = String(file.name || "").toLowerCase();
  if (/reporte[_\s-]?horas|tareo|validacion/.test(nameLower)) {
    setStatus(
      "error",
      "Ese archivo parece Tareo/Reporte_Horas. Aquí solo va Produccion_Licapa."
    );
    return;
  }
  if (!state.workerMap.size) {
    const ok = await loadPackedListado();
    if (!ok) {
      setStatus("error", "No se pudo cargar el JSON de trabajadores. Recarga la página.");
      return;
    }
  }
  try {
    setStatus("loading", "Leyendo Produccion_Licapa (puede tardar si el archivo es grande)…");
    showWorkspace(true);
    const buffer = await file.arrayBuffer();
    const parsed = parseProduccionBuffer(buffer, file.name, {
      workerMap: state.workerMap,
      onProgress: (p) =>
        setStatus("loading", `Procesando producción… ${Math.round(p)}%`)
    });

    state.prodFile = file.name;
    state.rows = parsed.rows;
    state.grupos = parsed.grupos;
    state.fechas = parsed.fechas;
    state.meta = parsed.meta;
    state.selectedGrupos = [];
    fillFechaOptions();
    applyFilters({ resetPage: true });
    setStatus("", "");
    if (!parsed.rows.length) {
      setStatus("empty", "No se encontraron filas con DNI (col. H) en el Excel.");
    }
  } catch (err) {
    console.error("[produccion]", err);
    setStatus("error", err?.message || "No se pudo leer el Excel de producción.");
  }
}

function exportExcel() {
  if (!state.filtered.length) {
    setStatus("empty", "No hay trabajadores para exportar con el filtro actual.");
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
      DNI: w.dni,
      Trabajador: w.trabajador,
      Grupo: w.grupo,
      Jarras: w.jarras,
      Registros: w.registros,
      Variedad: w.variedad || "",
      Huerto: w.huerto || ""
    }));
    const sheet = XLSX.utils.json_to_sheet(rows);
    sheet["!cols"] = [
      { wch: 5 },
      { wch: 12 },
      { wch: 36 },
      { wch: 22 },
      { wch: 10 },
      { wch: 10 },
      { wch: 18 },
      { wch: 14 }
    ];
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, "Produccion");
    const stamp = new Date().toISOString().slice(0, 10);
    const gTag = state.selectedGrupos.length
      ? `${state.selectedGrupos.length}grupos`
      : "todos";
    const fileName = `produccion-trabajadores-${gTag}-${stamp}.xlsx`;
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
    console.error("[produccion] export", err);
    setStatus("error", err?.message || "Error al exportar.");
  }
}

function destroyChart(key) {
  if (charts[key]) {
    charts[key].destroy();
    delete charts[key];
  }
}

function makeChart(key, canvas, config) {
  destroyChart(key);
  if (!canvas || !window.Chart) return;
  charts[key] = new window.Chart(canvas, config);
}

function parseFechaSortKey(fecha) {
  const m = String(fecha || "").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return 0;
  return Number(m[3]) * 10000 + Number(m[2]) * 100 + Number(m[1]);
}

function sortedFechas() {
  return [...state.fechas].sort((a, b) => parseFechaSortKey(a) - parseFechaSortKey(b));
}

/** Jarras por grupo para una fecha (respeta filtro de grupos + búsqueda). */
function jarrasPorGrupoEnFecha(fecha) {
  const map = new Map();
  if (!fecha) return map;
  const search = ($("prodFltSearch")?.value || "").trim().toLowerCase();
  const grupoSet = new Set(state.selectedGrupos);
  state.rows.forEach((row) => {
    if (row.fecha !== fecha) return;
    if (grupoSet.size && !grupoSet.has(row.grupo)) return;
    if (search) {
      const blob = `${row.dni} ${row.trabajador} ${row.grupo}`.toLowerCase();
      if (!blob.includes(search)) return;
    }
    const g = row.grupo || "(sin grupo)";
    map.set(g, (map.get(g) || 0) + (Number(row.jarras) || 0));
  });
  return map;
}

function fillResumenFechaSelects() {
  const fechas = sortedFechas();
  const actualSel = $("prodResumenFechaActual");
  const antSel = $("prodResumenFechaAnterior");
  if (!actualSel || !antSel) return;

  const opts = fechas
    .map((f) => `<option value="${escapeHtml(f)}">${escapeHtml(f)}</option>`)
    .join("");
  actualSel.innerHTML = opts || `<option value="">Sin fechas</option>`;
  antSel.innerHTML = opts || `<option value="">Sin fechas</option>`;

  const mainFecha = $("prodFltFecha")?.value || "";
  let actual = mainFecha && fechas.includes(mainFecha) ? mainFecha : fechas[fechas.length - 1] || "";
  const idx = fechas.indexOf(actual);
  let anterior = idx > 0 ? fechas[idx - 1] : fechas.length > 1 ? fechas[fechas.length - 2] : "";

  if (actual) actualSel.value = actual;
  if (anterior) antSel.value = anterior;
  else if (fechas[0]) antSel.value = fechas[0];
}

function renderDiffGruposChart() {
  const actual = $("prodResumenFechaActual")?.value || "";
  const anterior = $("prodResumenFechaAnterior")?.value || "";
  const hint = $("prodResumenDiffHint");
  const meta = $("prodResumenDiffMeta");

  if (hint) {
    hint.textContent =
      actual && anterior
        ? `${actual} − ${anterior} (mismo Excel)`
        : "Elige fecha actual y fecha anterior";
  }

  if (!actual || !anterior) {
    if (meta) meta.textContent = "Selecciona dos fechas para ver la diferencia por grupo.";
    destroyChart("grupos");
    return;
  }

  if (actual === anterior) {
    if (meta) meta.textContent = "Elige dos fechas distintas.";
    destroyChart("grupos");
    return;
  }

  const mapActual = jarrasPorGrupoEnFecha(actual);
  const mapAnterior = jarrasPorGrupoEnFecha(anterior);
  const grupos = new Set([...mapActual.keys(), ...mapAnterior.keys()]);
  const rows = [...grupos]
    .map((g) => {
      const a = mapActual.get(g) || 0;
      const b = mapAnterior.get(g) || 0;
      const diff = Math.round((a - b) * 1000) / 1000;
      return { grupo: g, actual: a, anterior: b, diff };
    })
    .sort((x, y) => Math.abs(y.diff) - Math.abs(x.diff) || y.actual - x.actual)
    .slice(0, 15);

  const sumDiff = rows.reduce((s, r) => s + r.diff, 0);
  const up = rows.filter((r) => r.diff > 0).length;
  const down = rows.filter((r) => r.diff < 0).length;
  if (meta) {
    meta.textContent = `Δ total visible: ${sumDiff >= 0 ? "+" : ""}${Math.round(sumDiff * 1000) / 1000} jarras · ${up} grupos ↑ · ${down} grupos ↓`;
  }

  makeChart("grupos", $("chartProdGrupos"), {
    type: "bar",
    data: {
      labels: rows.map((r) => r.grupo),
      datasets: [
        {
          label: `Δ jarras (${actual} − ${anterior})`,
          data: rows.map((r) => r.diff),
          backgroundColor: rows.map((r) =>
            r.diff > 0 ? "#27ae60" : r.diff < 0 ? "#e31e24" : "#94a3b8"
          ),
          borderRadius: 6
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
              const i = items?.[0]?.dataIndex;
              if (i == null || !rows[i]) return [];
              const r = rows[i];
              return [
                `Actual (${actual}): ${r.actual}`,
                `Anterior (${anterior}): ${r.anterior}`
              ];
            }
          }
        }
      },
      scales: {
        x: {
          grid: { color: "rgba(20,90,52,0.08)" },
          title: { display: true, text: "Diferencia de jarras" }
        },
        y: { grid: { display: false } }
      }
    }
  });
}

function openResumen() {
  const modal = $("modalProdResumen");
  if (!modal || !state.filtered.length) return;

  const search = ($("prodFltSearch")?.value || "").trim().toLowerCase();
  const fecha = $("prodFltFecha")?.value || "";
  const grupoSet = new Set(state.selectedGrupos);
  const baseForRank = state.rows.filter((row) => {
    if (grupoSet.size && !grupoSet.has(row.grupo)) return false;
    if (fecha && row.fecha !== fecha) return false;
    return true;
  });
  const ranked = rankGruposByJarras(baseForRank);
  const top = ranked.slice(0, 10);
  const bottom = ranked.slice(Math.max(0, ranked.length - 10));

  const jarrasTotal = state.filtered.reduce((s, w) => s + w.jarras, 0);
  const kpis = $("prodResumenKpis");
  if (kpis) {
    kpis.innerHTML = [
      ["Grupos", ranked.length],
      ["Trabajadores", state.filtered.length],
      ["Jarras", Math.round(jarrasTotal * 1000) / 1000],
      ["Mejor grupo", top[0] ? `${top[0].grupo} · ${fmt(top[0].jarras)}` : "—"],
      ["Peor grupo", bottom[0] ? `${bottom[0].grupo} · ${fmt(bottom[0].jarras)}` : "—"]
    ]
      .map(
        ([label, val]) => `<div class="tj-resumen__kpi">
        <span class="tj-resumen__kpi-label">${escapeHtml(label)}</span>
        <strong class="tj-resumen__kpi-value">${escapeHtml(String(val))}</strong>
      </div>`
      )
      .join("");
  }

  const sub = $("prodResumenSub");
  if (sub) {
    const parts = [];
    if (state.selectedGrupos.length) parts.push(`Grupos: ${state.selectedGrupos.join(", ")}`);
    if (fecha) parts.push(`Fecha: ${fecha}`);
    if (search) parts.push(`Buscar: ${search}`);
    if (state.rankingMode === "top10") parts.push("Vista: mejores 10");
    if (state.rankingMode === "bottom10") parts.push("Vista: peores 10");
    sub.textContent = parts.length ? parts.join(" · ") : "Todos los grupos del archivo cargado";
  }

  makeChart("mejores", $("chartProdMejores"), {
    type: "bar",
    data: {
      labels: top.map((g) => g.grupo),
      datasets: [
        {
          label: "Jarras",
          data: top.map((g) => g.jarras),
          backgroundColor: "#27ae60",
          borderRadius: 6
        }
      ]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { beginAtZero: true, grid: { color: "rgba(20,90,52,0.08)" } },
        y: { grid: { display: false } }
      }
    }
  });

  makeChart("peores", $("chartProdPeores"), {
    type: "bar",
    data: {
      labels: bottom.map((g) => g.grupo),
      datasets: [
        {
          label: "Jarras",
          data: bottom.map((g) => g.jarras),
          backgroundColor: "#e31e24",
          borderRadius: 6
        }
      ]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { beginAtZero: true, grid: { color: "rgba(227,30,36,0.08)" } },
        y: { grid: { display: false } }
      }
    }
  });

  fillResumenFechaSelects();
  renderDiffGruposChart();
  modal.hidden = false;
}

function closeResumen() {
  const modal = $("modalProdResumen");
  if (modal) modal.hidden = true;
  Object.keys(charts).forEach(destroyChart);
}

function openGrupoModal() {
  const modal = $("modalProdGrupo");
  if (!modal || !state.grupos.length) return;
  filterModal.options = [...state.grupos];
  filterModal.draft = new Set(state.selectedGrupos);
  const search = $("prodGrupoSearch");
  if (search) search.value = "";
  renderGrupoList("");
  updateGrupoCount();
  modal.hidden = false;
  window.setTimeout(() => search?.focus(), 40);
}

function closeGrupoModal() {
  const modal = $("modalProdGrupo");
  if (modal) modal.hidden = true;
}

function updateGrupoCount() {
  const el = $("prodGrupoCount");
  if (!el) return;
  const n = filterModal.draft.size;
  el.textContent = n ? `${n} grupo${n === 1 ? "" : "s"} seleccionado${n === 1 ? "" : "s"}` : "Todos los grupos (sin filtro)";
}

function renderGrupoList(query = "") {
  const list = $("prodGrupoList");
  if (!list) return;
  const q = String(query || "").trim().toLowerCase();
  const options = filterModal.options.filter((v) => !q || v.toLowerCase().includes(q));
  if (!options.length) {
    list.innerHTML = `<p class="rt-filter-modal__empty">Sin coincidencias</p>`;
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

function resetModule() {
  state.prodFile = "";
  state.rows = [];
  state.workers = [];
  state.filtered = [];
  state.grupos = [];
  state.fechas = [];
  state.selectedGrupos = [];
  state.rankingMode = "all";
  state.meta = { registros: 0, trabajadores: 0, jarras: 0 };
  pager.page = 1;
  closeResumen();
  closeGrupoModal();
  showWorkspace(false);
  setStatus("", "");
  syncExport();
  syncGrupoButton();
  syncRankingButtons();
  if ($("prodMeta")) $("prodMeta").textContent = "Sube el Excel de producción para comenzar.";
  fillFechaOptions();
}

function bindPager() {
  if (pager.bound) return;
  pager.bound = true;
  $("prodPagerPageSize")?.addEventListener("change", (e) => {
    pager.pageSize = Number(e.target.value) || 20;
    pager.page = 1;
    renderTable();
  });
  $("prodPagerFirst")?.addEventListener("click", () => {
    pager.page = 1;
    renderTable();
  });
  $("prodPagerPrev")?.addEventListener("click", () => {
    pager.page = Math.max(1, pager.page - 1);
    renderTable();
  });
  $("prodPagerNext")?.addEventListener("click", () => {
    const totalPages = Math.max(1, Math.ceil(state.filtered.length / pager.pageSize) || 1);
    pager.page = Math.min(totalPages, pager.page + 1);
    renderTable();
  });
  $("prodPagerLast")?.addEventListener("click", () => {
    pager.page = Math.max(1, Math.ceil(state.filtered.length / pager.pageSize) || 1);
    renderTable();
  });
}

function bindUploads() {
  const bindFile = (cardId, inputId, pickId, handler) => {
    const input = $(inputId);
    const card = $(cardId);
    $(pickId)?.addEventListener("click", (e) => {
      e.stopPropagation();
      input?.click();
    });
    input?.addEventListener("change", () => {
      const file = input.files?.[0];
      handler(file);
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
      handler(file);
    });
  };

  // Solo producción visible; listado queda oculto (JSON empaquetado)
  bindFile("prodDropCard", "prodInputExcel", "btnProdPickExcel", handleProduccion);
}

function bindGrupoModal() {
  if (filterModal.bound) return;
  filterModal.bound = true;

  $("prodBtnGrupo")?.addEventListener("click", openGrupoModal);
  $("btnCloseProdGrupo")?.addEventListener("click", closeGrupoModal);
  $("btnProdGrupoCancel")?.addEventListener("click", closeGrupoModal);
  $("btnProdGrupoApply")?.addEventListener("click", () => {
    state.selectedGrupos = [...filterModal.draft].sort((a, b) => a.localeCompare(b, "es"));
    closeGrupoModal();
    applyFilters({ resetPage: true });
  });
  document.querySelectorAll('[data-close-modal="prod-grupo"]').forEach((el) => {
    el.addEventListener("click", closeGrupoModal);
  });
  $("prodGrupoSearch")?.addEventListener("input", (e) => renderGrupoList(e.target.value));
  $("btnProdGrupoAll")?.addEventListener("click", () => {
    filterModal.options.forEach((v) => filterModal.draft.add(v));
    renderGrupoList($("prodGrupoSearch")?.value || "");
    updateGrupoCount();
  });
  $("btnProdGrupoNone")?.addEventListener("click", () => {
    filterModal.draft.clear();
    renderGrupoList($("prodGrupoSearch")?.value || "");
    updateGrupoCount();
  });
  $("prodGrupoList")?.addEventListener("change", (e) => {
    const input = e.target;
    if (!(input instanceof HTMLInputElement) || !input.classList.contains("rt-filter-modal__check"))
      return;
    if (input.checked) filterModal.draft.add(input.value);
    else filterModal.draft.delete(input.value);
    input.closest(".rt-filter-modal__option")?.classList.toggle("is-checked", input.checked);
    updateGrupoCount();
  });
}

async function bindUi() {
  bindUploads();
  bindPager();
  bindGrupoModal();

  // Prioridad: JSON empaquetado → caché local (si el JSON falló)
  const packed = await loadPackedListado();
  if (!packed) loadSavedListado();
  syncListadoUi();
  syncExport();

  $("prodFltSearch")?.addEventListener("input", () => applyFilters({ resetPage: true }));
  $("prodFltFecha")?.addEventListener("change", () => applyFilters({ resetPage: true }));
  ["btnProdRankAll", "btnProdRankTop", "btnProdRankBottom"].forEach((id) => {
    $(id)?.addEventListener("click", () => setRankingMode($(id).dataset.rank || "all"));
  });
  syncRankingButtons();
  $("btnProdExport")?.addEventListener("click", exportExcel);
  $("btnProdNewFile")?.addEventListener("click", resetModule);
  $("btnProdResumen")?.addEventListener("click", openResumen);
  $("btnCloseProdResumen")?.addEventListener("click", closeResumen);
  $("prodResumenFechaActual")?.addEventListener("change", renderDiffGruposChart);
  $("prodResumenFechaAnterior")?.addEventListener("change", renderDiffGruposChart);
  document.querySelectorAll('[data-close-modal="prod-resumen"]').forEach((el) => {
    el.addEventListener("click", closeResumen);
  });

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    if ($("modalProdResumen") && !$("modalProdResumen").hidden) closeResumen();
    if ($("modalProdGrupo") && !$("modalProdGrupo").hidden) closeGrupoModal();
  });
}

let started = false;
async function init() {
  if (started) return;
  started = true;
  await bindUi();
  renderTable();
}

window.addEventListener("qb:route-changed", (evt) => {
  if (evt.detail?.route === "produccion") init();
});

if (window.location.hash.replace(/^#\/?/, "").split("/")[0] === "produccion") {
  init();
}
