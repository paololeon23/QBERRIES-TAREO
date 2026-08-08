/** Gráficos de barras Chart.js. */

let chartSupervisor;
let chartMacro;
let chartTipo;

function destroyChart(chart) {
  if (chart) chart.destroy();
}

function baseOptions(onClickBar) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false }
    },
    onClick: (_evt, elements, chart) => {
      if (!elements.length) return;
      const index = elements[0].index;
      const label = chart.data.labels[index];
      onClickBar?.(label, chart.canvas.id);
    },
    scales: {
      x: {
        ticks: { maxRotation: 45, minRotation: 0, autoSkip: true, maxTicksLimit: 8 }
      },
      y: { beginAtZero: true }
    }
  };
}

export function renderCharts(filteredRows, onBarClick) {
  const bySupervisor = new Map();
  const byMacro = new Map();
  let china = 0;
  let convencional = 0;
  let sinTipo = 0;

  filteredRows.forEach((row) => {
    const sup = row.supervisor || "(sin supervisor)";
    const macro = row.macroPartida || "(sin macro)";
    const hours = Object.values(row.hoursByDay || {}).reduce((acc, h) => acc + (h || 0), 0);
    bySupervisor.set(sup, (bySupervisor.get(sup) || 0) + hours);

    const alertCount = row.status === "rojo" || row.status === "aviso" ? 1 : 0;
    byMacro.set(macro, (byMacro.get(macro) || 0) + alertCount);

    const tipo = (row.tipoBucket || row.sessionTipo || "").toLowerCase();
    if (tipo === "china") china += 1;
    else if (tipo === "convencional") convencional += 1;
    else sinTipo += 1;
  });

  const supLabels = [...bySupervisor.keys()].slice(0, 20);
  const macroLabels = [...byMacro.keys()].slice(0, 20);

  destroyChart(chartSupervisor);
  destroyChart(chartMacro);
  destroyChart(chartTipo);

  const canvasSup = document.getElementById("chartSupervisor");
  const canvasMacro = document.getElementById("chartMacro");
  const canvasTipo = document.getElementById("chartTipo");
  if (!window.Chart || !canvasSup || !canvasMacro || !canvasTipo) return;

  chartSupervisor = new window.Chart(canvasSup, {
    type: "bar",
    data: {
      labels: supLabels,
      datasets: [
        {
          data: supLabels.map((l) => Number((bySupervisor.get(l) || 0).toFixed(2))),
          backgroundColor: "#27ae60"
        }
      ]
    },
    options: {
      ...baseOptions((label) => onBarClick?.({ type: "supervisor", value: label })),
      maintainAspectRatio: false
    }
  });

  chartMacro = new window.Chart(canvasMacro, {
    type: "bar",
    data: {
      labels: macroLabels,
      datasets: [
        {
          data: macroLabels.map((l) => byMacro.get(l) || 0),
          backgroundColor: "#e31e24"
        }
      ]
    },
    options: {
      ...baseOptions((label) => onBarClick?.({ type: "macro", value: label })),
      maintainAspectRatio: false
    }
  });

  chartTipo = new window.Chart(canvasTipo, {
    type: "bar",
    data: {
      labels: ["China", "Convencional", "Sin etiqueta"],
      datasets: [
        {
          data: [china, convencional, sinTipo],
          backgroundColor: ["#e31e24", "#27ae60", "#b9d8c5"]
        }
      ]
    },
    options: {
      ...baseOptions((label) => {
        if (label === "China") onBarClick?.({ type: "tipo", value: "china" });
        if (label === "Convencional") onBarClick?.({ type: "tipo", value: "convencional" });
        if (label === "Sin etiqueta") onBarClick?.({ type: "tipo", value: "" });
      }),
      maintainAspectRatio: false
    }
  });

  // altura visual
  [canvasSup, canvasMacro, canvasTipo].forEach((c) => {
    c.parentElement.style.height = "220px";
  });
}
