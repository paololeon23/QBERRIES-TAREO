/** Dashboard Inicio — trabajadores + fundos/variedades */

const DATA_URL = "./data/trabajadores-resumen.json";

const FUNDO_CATALOG = [
  { id: "licapa", name: "LICAPA", variedades: ["Sekoya Pop"] },
  { id: "licapa-ii", name: "LICAPA II", variedades: ["Magica"] },
  { id: "licapa-iii", name: "LICAPA III", variedades: ["Sekoya Pop", "Magica"] }
];

const PALETTE = [
  "#27ae60",
  "#4fb04a",
  "#c4d62e",
  "#f6921e",
  "#e31e24",
  "#1a5c38",
  "#6bbf8a",
  "#8bc34a",
  "#ffb74d",
  "#ef5350",
  "#66bb6a",
  "#9ccc65",
  "#94a3b8"
];

const state = {
  raw: null,
  people: [],
  filtered: [],
  fundo: "",
  year: "",
  cargo: "",
  q: "",
  charts: {}
};

function $(id) {
  return document.getElementById(id);
}

function fmt(n) {
  return Number(n || 0).toLocaleString("es-PE");
}

function chartDefaults() {
  if (!window.Chart) return;
  Chart.defaults.font.family = '"Source Sans 3", "Segoe UI", sans-serif';
  Chart.defaults.color = "#4a6354";
  Chart.defaults.plugins.legend.labels.boxWidth = 12;
  Chart.defaults.plugins.legend.labels.boxHeight = 12;
}

function destroyChart(key) {
  if (state.charts[key]) {
    state.charts[key].destroy();
    delete state.charts[key];
  }
}

function getFilters() {
  return {
    year: ($("fltInicioYear")?.value || "").trim(),
    cargo: ($("fltInicioCargo")?.value || "").trim(),
    q: ($("fltInicioSearch")?.value || "").trim().toLowerCase(),
    fundo: state.fundo
  };
}

function applyFilters() {
  const f = getFilters();
  state.filtered = state.people.filter((p) => {
    if (f.year && p.y !== f.year) return false;
    if (f.cargo && p.c !== f.cargo) return false;
    if (f.q && !String(p.c || "").toLowerCase().includes(f.q)) return false;
    return true;
  });
  renderAll();
}

function aggregate(people) {
  const byCargo = new Map();
  const byYear = new Map();
  const byMonth = new Map();
  let obreros = 0;
  people.forEach((p) => {
    const c = p.c || "(Sin cargo)";
    byCargo.set(c, (byCargo.get(c) || 0) + 1);
    if (c === "OBRERO") obreros++;
    if (p.y) byYear.set(p.y, (byYear.get(p.y) || 0) + 1);
    if (p.m) byMonth.set(p.m, (byMonth.get(p.m) || 0) + 1);
  });
  const cargoArr = [...byCargo.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name, "es"));
  const top = cargoArr.slice(0, 10);
  const otros = cargoArr.slice(10).reduce((s, x) => s + x.count, 0);
  if (otros > 0) top.push({ name: "Otros", count: otros });
  return {
    total: people.length,
    obreros,
    staff: people.length - obreros,
    cargosUnicos: byCargo.size,
    byCargo: cargoArr,
    byCargoTop: top,
    byYear: [...byYear.entries()]
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => a.name.localeCompare(b.name)),
    byMonth: [...byMonth.entries()]
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => a.name.localeCompare(b.name))
  };
}

function renderKpis(agg) {
  const map = {
    kpiInicioTotal: agg.total,
    kpiInicioObreros: agg.obreros,
    kpiInicioStaff: agg.staff,
    kpiInicioFundos: FUNDO_CATALOG.length
  };
  Object.entries(map).forEach(([id, val]) => {
    const el = $(id);
    if (el) el.textContent = fmt(val);
  });
}

function renderFundos() {
  const wrap = $("inicioFundos");
  if (!wrap) return;
  wrap.innerHTML = FUNDO_CATALOG.map((f) => {
    const active = state.fundo === f.id ? " is-active" : "";
    const chips = f.variedades
      .map((v) => `<span class="inicio-chip">${v}</span>`)
      .join("");
    return `<article class="inicio-fundo${active}" data-fundo="${f.id}" tabindex="0">
      <p class="inicio-fundo__name">${f.name}</p>
      <div class="inicio-fundo__vars">${chips}</div>
    </article>`;
  }).join("");

  wrap.querySelectorAll(".inicio-fundo").forEach((el) => {
    el.addEventListener("click", () => {
      const id = el.dataset.fundo;
      state.fundo = state.fundo === id ? "" : id;
      renderFundos();
      renderFundoChart();
    });
  });
}

function makeOrUpdateChart(key, canvasId, config) {
  const canvas = $(canvasId);
  if (!canvas || !window.Chart) return;
  destroyChart(key);
  state.charts[key] = new Chart(canvas.getContext("2d"), config);
}

function renderCargoDonut(agg) {
  const labels = ["OBRERO", "Otros cargos"];
  const data = [agg.obreros, agg.staff];
  makeOrUpdateChart("donut", "chartInicioDonut", {
    type: "doughnut",
    data: {
      labels,
      datasets: [
        {
          data,
          backgroundColor: ["#27ae60", "#c4d62e"],
          borderWidth: 0,
          hoverOffset: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "62%",
      plugins: {
        legend: { position: "bottom" },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const v = ctx.parsed || 0;
              const pct = agg.total ? ((v / agg.total) * 100).toFixed(1) : 0;
              return ` ${ctx.label}: ${fmt(v)} (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

function renderCargoBars(agg) {
  const rows = agg.byCargoTop.slice(0, 12);
  makeOrUpdateChart("bars", "chartInicioBars", {
    type: "bar",
    data: {
      labels: rows.map((r) => r.name),
      datasets: [
        {
          label: "Trabajadores",
          data: rows.map((r) => r.count),
          backgroundColor: rows.map((_, i) => PALETTE[i % PALETTE.length]),
          borderRadius: 8,
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
        x: { grid: { color: "rgba(0,0,0,0.05)" }, ticks: { precision: 0 } },
        y: { grid: { display: false } }
      }
    }
  });
}

function renderYearBars(agg) {
  makeOrUpdateChart("years", "chartInicioYears", {
    type: "bar",
    data: {
      labels: agg.byYear.map((r) => r.name),
      datasets: [
        {
          label: "Ingresos",
          data: agg.byYear.map((r) => r.count),
          backgroundColor: "#27ae60",
          borderRadius: 10,
          maxBarThickness: 48
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, grid: { color: "rgba(0,0,0,0.05)" }, ticks: { precision: 0 } },
        x: { grid: { display: false } }
      }
    }
  });
}

function renderMonthLine(agg) {
  makeOrUpdateChart("months", "chartInicioMonths", {
    type: "line",
    data: {
      labels: agg.byMonth.map((r) => r.name),
      datasets: [
        {
          label: "Ingresos por mes",
          data: agg.byMonth.map((r) => r.count),
          borderColor: "#27ae60",
          backgroundColor: "rgba(39, 174, 96, 0.18)",
          fill: true,
          tension: 0.35,
          pointRadius: 3,
          pointHoverRadius: 5
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, grid: { color: "rgba(0,0,0,0.05)" }, ticks: { precision: 0 } },
        x: { grid: { display: false }, ticks: { maxRotation: 45, minRotation: 0 } }
      }
    }
  });
}

function renderFundoChart() {
  const active = FUNDO_CATALOG.find((f) => f.id === state.fundo);
  const labels = active
    ? active.variedades
    : ["Sekoya Pop (LICAPA + III)", "Magica (LICAPA II + III)"];
  // Cobertura de variedades por fundo (catálogo gerencial, no viene en el Excel de personas)
  const data = active
    ? active.variedades.map(() => 1)
    : [2, 2]; // aparece en 2 fundos cada una según reglas
  const title = $("chartInicioFundoTitle");
  if (title) {
    title.textContent = active
      ? `Variedades · ${active.name}`
      : "Variedades por cobertura de fundo";
  }
  makeOrUpdateChart("fundo", "chartInicioFundo", {
    type: "pie",
    data: {
      labels,
      datasets: [
        {
          data,
          backgroundColor: active ? ["#27ae60", "#c4d62e"] : ["#27ae60", "#f6921e"],
          borderWidth: 0
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "bottom" },
        tooltip: {
          callbacks: {
            label: (ctx) =>
              active
                ? ` ${ctx.label}`
                : ` ${ctx.label}`
          }
        }
      }
    }
  });
}

function populateCargoSelect(people) {
  const sel = $("fltInicioCargo");
  if (!sel) return;
  const current = sel.value;
  const cargos = [...new Set(people.map((p) => p.c).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, "es")
  );
  sel.innerHTML =
    `<option value="">Todos</option>` +
    cargos.map((c) => `<option value="${c.replace(/"/g, "&quot;")}">${c}</option>`).join("");
  if (cargos.includes(current)) sel.value = current;
}

function populateYearSelect(people) {
  const sel = $("fltInicioYear");
  if (!sel) return;
  const current = sel.value;
  const years = [...new Set(people.map((p) => p.y).filter(Boolean))].sort();
  sel.innerHTML =
    `<option value="">Todos</option>` +
    years.map((y) => `<option value="${y}">${y}</option>`).join("");
  if (years.includes(current)) sel.value = current;
}

function renderAll() {
  const agg = aggregate(state.filtered);
  renderKpis(agg);
  renderCargoDonut(agg);
  renderCargoBars(agg);
  renderYearBars(agg);
  renderMonthLine(agg);
  renderFundoChart();
  const meta = $("inicioMeta");
  if (meta) {
    meta.textContent = `Mostrando ${fmt(agg.total)} de ${fmt(state.people.length)} trabajadores · fuente: listado Excel`;
  }
}

function bindDragCards() {
  const grid = $("inicioCharts");
  if (!grid) return;
  let dragEl = null;

  grid.querySelectorAll(".inicio-chart-card").forEach((card) => {
    card.setAttribute("draggable", "true");
    card.addEventListener("dragstart", () => {
      dragEl = card;
      card.classList.add("is-dragging");
    });
    card.addEventListener("dragend", () => {
      card.classList.remove("is-dragging");
      grid.querySelectorAll(".inicio-chart-card").forEach((c) => c.classList.remove("is-drop-target"));
      dragEl = null;
    });
    card.addEventListener("dragover", (evt) => {
      evt.preventDefault();
      if (!dragEl || dragEl === card) return;
      card.classList.add("is-drop-target");
    });
    card.addEventListener("dragleave", () => card.classList.remove("is-drop-target"));
    card.addEventListener("drop", (evt) => {
      evt.preventDefault();
      card.classList.remove("is-drop-target");
      if (!dragEl || dragEl === card) return;
      const cards = [...grid.querySelectorAll(".inicio-chart-card")];
      const from = cards.indexOf(dragEl);
      const to = cards.indexOf(card);
      if (from < 0 || to < 0) return;
      if (from < to) grid.insertBefore(dragEl, card.nextSibling);
      else grid.insertBefore(dragEl, card);
    });
  });
}

function bindFilters() {
  ["fltInicioYear", "fltInicioCargo", "fltInicioSearch"].forEach((id) => {
    $(id)?.addEventListener("input", () => applyFilters());
    $(id)?.addEventListener("change", () => applyFilters());
  });
  $("btnInicioReset")?.addEventListener("click", () => {
    state.fundo = "";
    if ($("fltInicioYear")) $("fltInicioYear").value = "";
    if ($("fltInicioCargo")) $("fltInicioCargo").value = "";
    if ($("fltInicioSearch")) $("fltInicioSearch").value = "";
    renderFundos();
    applyFilters();
  });
}

async function loadData() {
  const status = $("inicioStatus");
  if (status) {
    status.hidden = false;
    status.textContent = "Cargando listado de trabajadores…";
  }
  try {
    const res = await fetch(`${DATA_URL}?v=${Date.now()}`, { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    state.raw = json;
    state.people = Array.isArray(json.people) ? json.people : [];
    populateYearSelect(state.people);
    populateCargoSelect(state.people);
    state.filtered = state.people.slice();
    if (status) status.hidden = true;
    renderFundos();
    renderAll();
  } catch (err) {
    if (status) {
      status.hidden = false;
      status.textContent = err?.message || "No se pudo cargar trabajadores-resumen.json";
    }
  }
}

let started = false;

function initInicio() {
  if (started) return;
  started = true;
  chartDefaults();
  bindFilters();
  bindDragCards();
  loadData();
}

window.addEventListener("qb:route-changed", (evt) => {
  if (evt.detail?.route === "inicio") initInicio();
});

if (window.location.hash.replace(/^#\/?/, "").split("/")[0] === "inicio" || !window.location.hash || window.location.hash === "#/") {
  initInicio();
}
