/** Vista Resumen: tabla por supervisor/fundo + gráfico + errores. */

import { collapseToDayRows } from "./validacion-table.js";
import { countSupervisoresCosto } from "./validacion-kpi.js";

let chartPersonas = null;
let chartErrores = null;

function destroyChart(chart) {
  if (chart) {
    chart.destroy();
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Agrupa persona-día por Fundo + Supervisor. */
export function buildResumenGroups(dayRows) {
  const map = new Map();
  const trabajadoresGlobal = new Set();

  dayRows.forEach((row) => {
    const fundo = row.fundo || "(sin fundo)";
    const supervisor = row.supervisor || "(sin supervisor)";
    const key = `${fundo}||${supervisor}`;
    if (!map.has(key)) {
      map.set(key, {
        fundo,
        supervisor,
        planillas: 0,
        trabajadores: new Set(),
        errores: 0,
        avisos: 0,
        ok: 0
      });
    }
    const g = map.get(key);
    g.planillas += 1;
    if (row.documento) {
      g.trabajadores.add(String(row.documento));
      trabajadoresGlobal.add(String(row.documento));
    }
    if (row.status === "rojo") g.errores += 1;
    else if (row.status === "aviso") g.avisos += 1;
    else g.ok += 1;
  });

  const groups = [...map.values()]
    .map((g) => ({
      fundo: g.fundo,
      supervisor: g.supervisor,
      planillas: g.planillas,
      trabajadores: g.trabajadores.size,
      errores: g.errores,
      avisos: g.avisos,
      ok: g.ok
    }))
    .sort((a, b) => {
      const fa = a.fundo.localeCompare(b.fundo, "es");
      if (fa) return fa;
      return a.supervisor.localeCompare(b.supervisor, "es");
    });

  return {
    groups,
    kpis: {
      grupos: groups.length,
      fundos: new Set(groups.map((g) => g.fundo)).size,
      supervisores: new Set(groups.map((g) => g.supervisor)).size,
      trabajadores: trabajadoresGlobal.size,
      errores: groups.reduce((s, g) => s + g.errores, 0),
      avisos: groups.reduce((s, g) => s + g.avisos, 0)
    }
  };
}

/** Personas únicas por Macro Partida. */
export function buildPersonasPorMacro(dayRows) {
  const map = new Map();
  dayRows.forEach((row) => {
    const macro = row.macroPartida || "(sin macro)";
    if (!map.has(macro)) map.set(macro, new Set());
    if (row.documento) map.get(macro).add(String(row.documento));
  });
  return [...map.entries()]
    .map(([label, set]) => ({ label, value: set.size }))
    .sort((a, b) => b.value - a.value);
}

/** Personas únicas por Actividad (columna M). */
export function buildPersonasPorActividad(dayRows) {
  const map = new Map();
  dayRows.forEach((row) => {
    const act = String(row.actividad || "").trim() || "(sin actividad)";
    if (!map.has(act)) map.set(act, new Set());
    if (row.documento) map.get(act).add(String(row.documento));
  });
  return [...map.entries()]
    .map(([label, set]) => ({ label, value: set.size }))
    .sort((a, b) => b.value - a.value);
}

/** Supervisores con más errores (persona-día en rojo), top N. */
export function buildErroresPorSupervisor(dayRows, limit = 10) {
  const map = new Map();
  dayRows.forEach((row) => {
    if (row.status !== "rojo") return;
    const name = row.supervisor || "(sin supervisor)";
    map.set(name, (map.get(name) || 0) + 1);
  });
  return [...map.entries()]
    .map(([label, value]) => ({ label, value }))
    .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label, "es"))
    .slice(0, limit);
}

const PIE_COLORS = [
  "#5dade2",
  "#58d68d",
  "#f5b041",
  "#af7ac5",
  "#5d6d7e",
  "#ec7063",
  "#48c9b0",
  "#f4d03f",
  "#85929e",
  "#a569bd"
];

const ERROR_BAR_COLORS = [
  "#e74c3c",
  "#c0392b",
  "#e67e22",
  "#d35400",
  "#cd6155",
  "#ec7063",
  "#f1948a",
  "#e59866",
  "#af7ac5",
  "#5d6d7e"
];

export function openResumenModal() {
  const modal = document.getElementById("modalResumen");
  if (modal) modal.hidden = false;
}

export function closeResumenModal() {
  const modal = document.getElementById("modalResumen");
  if (modal) modal.hidden = true;
}

function escapeAttr(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/"/g, "&quot;");
}

/**
 * Actividades disponibles según filtros (como Excel: al elegir Macro solo salen sus actividades).
 */
function fillResumenActividadOptions(validated, { supervisor = "", fundo = "", macro = "" } = {}) {
  const actSelect = document.getElementById("resumenFltActividad");
  if (!actSelect) return;

  const prev = actSelect.value;
  const activities = [
    ...new Set(
      (validated?.rows || [])
        .filter((row) => {
          if (supervisor && row.supervisor !== supervisor) return false;
          if (fundo && row.fundo !== fundo) return false;
          if (macro && row.macroPartida !== macro) return false;
          return true;
        })
        .map((r) => String(r.actividad || "").trim())
        .filter(Boolean)
    )
  ].sort((a, b) => a.localeCompare(b, "es"));

  actSelect.innerHTML = `<option value="">Todas</option>${activities
    .map((v) => `<option value="${escapeAttr(v)}">${escapeAttr(v)}</option>`)
    .join("")}`;
  actSelect.value = activities.includes(prev) ? prev : "";
}

export function syncResumenFiltersFromMain(validated) {
  const copy = (fromId, toId) => {
    const from = document.getElementById(fromId);
    const to = document.getElementById(toId);
    if (!from || !to) return;
    to.innerHTML = from.innerHTML;
    to.value = from.value;
  };
  copy("fltSupervisor", "resumenFltSupervisor");
  copy("fltFundo", "resumenFltFundo");

  const macroSel = document.getElementById("resumenFltMacro");
  if (macroSel) {
    const prev = macroSel.value;
    const macros = [
      ...new Set(
        (validated?.rows || [])
          .map((r) => String(r.macroPartida || "").trim())
          .filter(Boolean)
      )
    ].sort((a, b) => a.localeCompare(b, "es"));
    macroSel.innerHTML = `<option value="">Todas</option>${macros
      .map((v) => `<option value="${escapeAttr(v)}">${escapeAttr(v)}</option>`)
      .join("")}`;
    macroSel.value = macros.includes(prev) ? prev : "";
  }

  const supervisor = document.getElementById("resumenFltSupervisor")?.value || "";
  const fundo = document.getElementById("resumenFltFundo")?.value || "";
  const macro = document.getElementById("resumenFltMacro")?.value || "";
  fillResumenActividadOptions(validated, { supervisor, fundo, macro });
}

/** Misma data que la tabla del modal Resumen (respeta filtros del modal). */
export function getFilteredResumenData(validated, mainFilters = {}) {
  const supervisor =
    document.getElementById("resumenFltSupervisor")?.value || mainFilters.supervisor || "";
  const fundo = document.getElementById("resumenFltFundo")?.value || mainFilters.fundo || "";
  const macro = document.getElementById("resumenFltMacro")?.value || mainFilters.macro || "";

  fillResumenActividadOptions(validated, { supervisor, fundo, macro });

  const actividad =
    document.getElementById("resumenFltActividad")?.value || mainFilters.actividad || "";

  const filtered = (validated.rows || []).filter((row) => {
    if (supervisor && row.supervisor !== supervisor) return false;
    if (fundo && row.fundo !== fundo) return false;
    if (macro && row.macroPartida !== macro) return false;
    if (actividad && String(row.actividad || "").trim() !== actividad) return false;
    return true;
  });

  const dayRows = collapseToDayRows(filtered);
  const { groups, kpis } = buildResumenGroups(dayRows);
  return {
    groups,
    kpis,
    filtered,
    dayRows,
    filters: { supervisor, fundo, macro, actividad },
    macro,
    actividad
  };
}

export function renderResumenView(validated, mainFilters = {}) {
  if (!validated) return;

  const { groups, kpis, filtered, dayRows, macro } = getFilteredResumenData(validated, mainFilters);

  // Misma lógica que la card Supervisores de la pantalla principal
  const supervisoresKpi = countSupervisoresCosto(filtered);

  // Con Macro elegida: dona por Actividad; si no, por Macro Partida
  const porPie = macro ? buildPersonasPorActividad(dayRows) : buildPersonasPorMacro(dayRows);
  const pieTitle = macro ? "Personas por Actividad" : "Personas por Macro Partida";
  const erroresPorSup = buildErroresPorSupervisor(dayRows, 10);

  const kpiHost = document.getElementById("resumenKpis");
  if (kpiHost) {
    kpiHost.innerHTML = `
      <div class="resumen-kpi"><span class="resumen-kpi__label">Supervisores</span><span class="resumen-kpi__value">${supervisoresKpi}</span></div>
      <div class="resumen-kpi"><span class="resumen-kpi__label">Trabajadores</span><span class="resumen-kpi__value">${kpis.trabajadores}</span></div>
      <div class="resumen-kpi resumen-kpi--danger"><span class="resumen-kpi__label">Errores</span><span class="resumen-kpi__value">${kpis.errores}</span></div>
      <div class="resumen-kpi resumen-kpi--warn"><span class="resumen-kpi__label">Extras</span><span class="resumen-kpi__value">${kpis.avisos}</span></div>
    `;
  }

  const body = document.getElementById("resumenTableBody");
  if (body) {
    body.innerHTML = groups.length
      ? groups
          .map((g) => {
            const errClass = g.errores > 0 ? " is-cell-danger" : "";
            const warnClass = g.avisos > 0 ? " is-cell-warn" : "";
            return `<tr class="${g.errores > 0 ? "is-row-danger" : g.avisos > 0 ? "is-row-warn" : ""}">
              <td>${escapeHtml(g.fundo)}</td>
              <td>${escapeHtml(g.supervisor)}</td>
              <td>${g.planillas}</td>
              <td>${g.trabajadores}</td>
              <td class="${errClass}">${g.errores}</td>
              <td class="${warnClass}">${g.avisos}</td>
              <td>${g.ok}</td>
            </tr>`;
          })
          .join("")
      : `<tr><td colspan="7" class="resumen-empty">Sin datos para estos filtros.</td></tr>`;
  }

  const range = document.getElementById("resumenPagerRange");
  if (range) range.textContent = `1 – ${groups.length} of ${groups.length}`;

  const titleCompany = document.getElementById("resumenChartTitle");
  if (titleCompany) {
    titleCompany.textContent = "QBerries";
  }
  const subtitle = document.getElementById("resumenChartSubtitle");
  if (subtitle) {
    subtitle.textContent = `Total trabajadores: ${kpis.trabajadores}`;
  }

  renderResumenCharts(porPie, erroresPorSup, pieTitle);
}

function shortSupervisorName(name) {
  const parts = String(name || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  if (parts.length <= 2) return parts.join(" ");
  // Nombre + 1er apellido (más legible en barras)
  return `${parts[0]} ${parts[1]}`;
}

function shortMacroLabel(label) {
  const raw = String(label || "").trim();
  if (!raw) return "—";
  const upper = raw.toUpperCase();
  const aliases = {
    "COSTO DE COSECHA": "Cosecha",
    "REMUNERACIONES ADMINISTRATIVAS": "Remuneraciones",
    LABORES: "Labores",
    SANIDAD: "Sanidad",
    RIEGO: "Riego",
    ALMACENES: "Almacenes"
  };
  if (aliases[upper]) return aliases[upper];
  if (raw.length <= 14) return raw;
  return `${raw.slice(0, 12)}…`;
}

/** % dentro de cada tramo del donut (todos los segmentos). */
const doughnutInsidePercent = {
  id: "doughnutInsidePercent",
  afterDatasetsDraw(chart) {
    if (chart.config.type !== "doughnut") return;
    const { ctx } = chart;
    const meta = chart.getDatasetMeta(0);
    if (!meta?.data?.length) return;
    const values = (chart.data.datasets[0]?.data || []).map((v) => Number(v) || 0);
    const total = values.reduce((a, b) => a + b, 0);
    if (!total) return;

    ctx.save();
    meta.data.forEach((arc, i) => {
      if (!arc || arc.hidden) return;
      const value = values[i];
      if (!value) return;
      const pct = (value / total) * 100;

      const props =
        typeof arc.getProps === "function"
          ? arc.getProps(["startAngle", "endAngle", "innerRadius", "outerRadius", "x", "y"], true)
          : {
              startAngle: arc.startAngle,
              endAngle: arc.endAngle,
              innerRadius: arc.innerRadius,
              outerRadius: arc.outerRadius,
              x: arc.x,
              y: arc.y
            };

      const mid = (props.startAngle + props.endAngle) / 2;
      const cos = Math.cos(mid);
      const sin = Math.sin(mid);
      // Tramos chicos: etiqueta más afuera del anillo para que se lea
      const t = pct >= 8 ? 0.5 : pct >= 3 ? 0.62 : 0.78;
      const r = props.innerRadius + (props.outerRadius - props.innerRadius) * t;
      const x = props.x + cos * r;
      const y = props.y + sin * r;

      const label =
        pct >= 10 ? `${Math.round(pct)}%` : pct >= 1 ? `${pct.toFixed(1)}%` : `${pct.toFixed(1)}%`;
      const fontSize = pct >= 25 ? 13 : pct >= 8 ? 11 : pct >= 3 ? 9 : 8;

      ctx.font = `700 ${fontSize}px 'Segoe UI', system-ui, sans-serif`;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      // Borde para contraste sobre colores claros u oscuros
      ctx.lineWidth = pct >= 8 ? 3 : 2.5;
      ctx.strokeStyle = "rgba(17, 24, 39, 0.45)";
      ctx.fillStyle = "#fff";
      ctx.strokeText(label, x, y);
      ctx.fillText(label, x, y);
    });
    ctx.restore();
  }
};

function renderResumenCharts(porPie, erroresPorSup, pieTitle = "Personas por Macro Partida") {
  const canvasMacro = document.getElementById("chartResumenMacro");
  const canvasEstado = document.getElementById("chartResumenEstado");
  if (!window.Chart) return;

  destroyChart(chartPersonas);
  destroyChart(chartErrores);
  chartPersonas = null;
  chartErrores = null;

  if (canvasMacro) {
    const labels = porPie.map((x) => x.label);
    const data = porPie.map((x) => x.value);
    const total = data.reduce((a, b) => a + (Number(b) || 0), 0) || 1;
    chartPersonas = new window.Chart(canvasMacro, {
      type: "doughnut",
      data: {
        labels,
        datasets: [
          {
            data,
            backgroundColor: labels.map((_, i) => PIE_COLORS[i % PIE_COLORS.length]),
            borderWidth: 2,
            borderColor: "#fff"
          }
        ]
      },
      plugins: [doughnutInsidePercent],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        radius: "90%",
        cutout: "48%",
        layout: {
          padding: { top: 0, right: 2, bottom: 0, left: 2 }
        },
        plugins: {
          legend: {
            position: "right",
            align: "center",
            labels: {
              boxWidth: 10,
              boxHeight: 10,
              font: { size: 11 },
              padding: 10,
              color: "#4b5563",
              generateLabels(chart) {
                const ds = chart.data.datasets[0];
                return chart.data.labels.map((label, i) => {
                  const val = Number(ds.data[i]) || 0;
                  const ratio = val / total;
                  const pct = (ratio * 100).toFixed(ratio >= 0.1 ? 0 : 1);
                  return {
                    text: `${shortMacroLabel(label)}  ${val}  (${pct}%)`,
                    fillStyle: ds.backgroundColor[i],
                    strokeStyle: "#fff",
                    lineWidth: 1,
                    hidden: false,
                    index: i
                  };
                });
              }
            }
          },
          title: {
            display: true,
            text: pieTitle,
            color: "#374151",
            font: { size: 12, weight: "600" },
            padding: { top: 0, bottom: 4 },
            align: "start"
          },
          tooltip: {
            callbacks: {
              label(ctx) {
                const val = Number(ctx.raw) || 0;
                const pct = ((val / total) * 100).toFixed(1);
                return ` ${ctx.label}: ${val} (${pct}%)`;
              }
            }
          }
        }
      }
    });
  }

  if (canvasEstado) {
    const labels = erroresPorSup.map((x) => shortSupervisorName(x.label));
    const fullNames = erroresPorSup.map((x) => x.label);
    const data = erroresPorSup.map((x) => x.value);
    chartErrores = new window.Chart(canvasEstado, {
      type: "bar",
      data: {
        labels: labels.length ? labels : ["Sin errores"],
        datasets: [
          {
            label: "Errores",
            data: data.length ? data : [0],
            backgroundColor: data.map((_, i) => ERROR_BAR_COLORS[i % ERROR_BAR_COLORS.length]),
            borderRadius: 6,
            maxBarThickness: 26,
            categoryPercentage: 0.72,
            barPercentage: 0.85
          }
        ]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        layout: {
          padding: { top: 6, right: 18, bottom: 4, left: 4 }
        },
        plugins: {
          legend: { display: false },
          title: {
            display: true,
            text: "Supervisores con más errores",
            color: "#374151",
            font: { size: 13, weight: "600" },
            padding: { top: 2, bottom: 14 }
          },
          tooltip: {
            callbacks: {
              title(items) {
                const i = items[0]?.dataIndex ?? 0;
                return fullNames[i] || labels[i] || "";
              },
              label(item) {
                return ` Errores: ${item.raw}`;
              }
            }
          }
        },
        scales: {
          x: {
            beginAtZero: true,
            ticks: { precision: 0 },
            grid: { color: "#f3f4f6" },
            afterFit(scale) {
              scale.paddingRight = 8;
            }
          },
          y: {
            ticks: {
              font: { size: 10 },
              autoSkip: false,
              crossAlign: "far"
            },
            grid: { display: false },
            afterFit(scale) {
              scale.width = Math.max(scale.width, 96);
            }
          }
        }
      }
    });
  }
}

export function bindResumenUi({ getValidated, getMainFilters, onExport }) {
  document.getElementById("btnCloseResumen")?.addEventListener("click", closeResumenModal);
  document.querySelectorAll('[data-close-modal="resumen"]').forEach((el) => {
    el.addEventListener("click", closeResumenModal);
  });

  const refresh = () => {
    const validated = getValidated?.();
    if (!validated) return;
    renderResumenView(validated, getMainFilters?.() || {});
  };

  ["resumenFltSupervisor", "resumenFltFundo", "resumenFltMacro", "resumenFltActividad"].forEach((id) => {
    document.getElementById(id)?.addEventListener("change", refresh);
  });

  document.getElementById("btnResumenExport")?.addEventListener("click", () => {
    onExport?.();
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeResumenModal();
  });
}
