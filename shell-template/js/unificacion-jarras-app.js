(function () {
"use strict";

/**
 * Unificación de jarras — Suma / Descuento / Descarte
 * Sube varios Excel del mismo tipo y los une en una sola tabla descargable.
 */

const TYPES = {
  suma: {
    key: "suma",
    label: "Suma de jarras",
    tipoMatch: ["SUMA"],
    sheetName: "Suma"
  },
  descuento: {
    key: "descuento",
    label: "Descuento jarras",
    tipoMatch: ["DESCUENTO"],
    sheetName: "Descuento"
  },
  descarte: {
    key: "descarte",
    label: "Descarte jarras",
    tipoMatch: ["DESCARTE"],
    sheetName: "Descarte"
  }
};

const HEADERS = [
  "NOMBRE DEL SUPERVISOR",
  "TIPO",
  "TOTAL TURNO DÍA",
  "TOTAL TURNO TARDE",
  "TOTAL DE TODO",
  "TOTAL DE TRABAJADORES",
  "LOTE O LOTES",
  "DÍA",
  "HORA REGISTRO"
];

const state = {
  active: "suma",
  buckets: {
    suma: { files: [], rows: [] },
    descuento: { files: [], rows: [] },
    descarte: { files: [], rows: [] }
  }
};

function $(id) {
  return document.getElementById(id);
}

function escapeHtml(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fmt(n) {
  return Number(n || 0).toLocaleString("es-PE");
}

function setStatus(kind, text) {
  const el = $("ujStatus");
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

function normHeader(v) {
  return String(v ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function isHeaderRow(row) {
  const a = normHeader(row[0]);
  return a.includes("SUPERVISOR") || a.includes("NOMBRE");
}

function rowToObject(row) {
  return {
    supervisor: String(row[0] ?? "").trim(),
    tipo: String(row[1] ?? "").trim().toUpperCase(),
    turnoDia: row[2] ?? "",
    turnoTarde: row[3] ?? "",
    totalTodo: row[4] ?? "",
    totalTrabajadores: row[5] ?? "",
    lotes: String(row[6] ?? "").trim(),
    dia: String(row[7] ?? "").trim(),
    hora: String(row[8] ?? "").trim(),
    _source: ""
  };
}

function objectToRow(o) {
  return [
    o.supervisor,
    o.tipo,
    o.turnoDia,
    o.turnoTarde,
    o.totalTodo,
    o.totalTrabajadores,
    o.lotes,
    o.dia,
    o.hora
  ];
}

function matchesType(tipo, typeKey) {
  const t = String(tipo || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
  const cfg = TYPES[typeKey];
  if (!cfg) return false;
  return cfg.tipoMatch.some((m) => t === m);
}

function parseMatrix(matrix, fileName, typeKey) {
  if (!matrix?.length) return [];
  let start = 0;
  if (isHeaderRow(matrix[0])) start = 1;
  const out = [];
  for (let r = start; r < matrix.length; r++) {
    const row = matrix[r] || [];
    const sup = String(row[0] ?? "").trim();
    if (!sup) continue;
    const obj = rowToObject(row);
    obj._source = fileName;
    // Si trae TIPO y no coincide, saltar; si no trae tipo, aceptar en el bucket activo
    if (obj.tipo && !matchesType(obj.tipo, typeKey)) continue;
    if (!obj.tipo) obj.tipo = TYPES[typeKey].tipoMatch[0];
    out.push(obj);
  }
  return out;
}

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    if (!file) return reject(new Error("Sin archivo"));
    if (!window.XLSX) return reject(new Error("No se cargó XLSX"));
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = window.XLSX.read(reader.result, { type: "array" });
        const sheet = wb.Sheets[wb.Sheets[0]?.Name || wb.SheetNames[0]];
        const matrix = window.XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "",
          raw: false
        });
        resolve(matrix);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("No se pudo leer el archivo"));
    reader.readAsArrayBuffer(file);
  });
}

function activeBucket() {
  return state.buckets[state.active];
}

function rebuildRows(typeKey) {
  const bucket = state.buckets[typeKey];
  bucket.rows = bucket.files.flatMap((f) => f.rows);
}

async function addFiles(fileList, typeKey) {
  const bucket = state.buckets[typeKey];
  const files = [...(fileList || [])];
  if (!files.length) return;
  let added = 0;
  for (const file of files) {
    const matrix = await readExcelFile(file);
    const rows = parseMatrix(matrix, file.name, typeKey);
    bucket.files.push({ name: file.name, rows });
    added += rows.length;
  }
  rebuildRows(typeKey);
  renderAll();
  setStatus("ok", `${fmt(added)} filas agregadas en «${TYPES[typeKey].label}».`);
}

function clearBucket(typeKey) {
  state.buckets[typeKey] = { files: [], rows: [] };
  renderAll();
  setStatus("", "");
}

function exportUnified(typeKey) {
  const bucket = state.buckets[typeKey];
  if (!bucket.rows.length) {
    setStatus("error", "No hay filas para exportar.");
    return;
  }
  if (!window.XLSX) {
    setStatus("error", "No se cargó el exportador Excel.");
    return;
  }
  const data = [HEADERS, ...bucket.rows.map(objectToRow)];
  const ws = window.XLSX.utils.aoa_to_sheet(data);
  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, TYPES[typeKey].sheetName);
  const stamp = new Date().toISOString().slice(0, 10);
  window.XLSX.writeFile(wb, `unificacion-${typeKey}-${stamp}.xlsx`);
  setStatus("ok", `Excel unificado descargado (${fmt(bucket.rows.length)} filas).`);
}

function renderTabs() {
  document.querySelectorAll("[data-uj-tab]").forEach((btn) => {
    const on = btn.dataset.ujTab === state.active;
    btn.classList.toggle("is-active", on);
    btn.classList.toggle("btn--secondary", on);
    btn.classList.toggle("btn--ghost", !on);
  });
}

function renderFileList() {
  const host = $("ujFileList");
  if (!host) return;
  const bucket = activeBucket();
  if (!bucket.files.length) {
    host.innerHTML = `<p class="uj-empty">Aún no hay archivos en este apartado.</p>`;
    return;
  }
  host.innerHTML = bucket.files
    .map(
      (f) => `<div class="uj-file">
        <span class="uj-file__name">${escapeHtml(f.name)}</span>
        <span class="uj-file__count">${fmt(f.rows.length)} filas</span>
      </div>`
    )
    .join("");
}

function renderTable() {
  const host = $("ujTableBody");
  const empty = $("ujTableEmpty");
  const count = $("ujRowCount");
  if (!host) return;
  const bucket = activeBucket();
  if (count) {
    count.textContent = bucket.rows.length
      ? `${fmt(bucket.rows.length)} filas unificadas · ${fmt(bucket.files.length)} archivo(s)`
      : "";
  }
  if (!bucket.rows.length) {
    host.innerHTML = "";
    if (empty) empty.hidden = false;
    return;
  }
  if (empty) empty.hidden = true;
  host.innerHTML = bucket.rows
    .map(
      (r, i) => `<tr>
        <td>${i + 1}</td>
        <td>${escapeHtml(r.supervisor)}</td>
        <td><span class="uj-tipo">${escapeHtml(r.tipo)}</span></td>
        <td class="uj-num">${escapeHtml(r.turnoDia)}</td>
        <td class="uj-num">${escapeHtml(r.turnoTarde)}</td>
        <td class="uj-num">${escapeHtml(r.totalTodo)}</td>
        <td class="uj-num">${escapeHtml(r.totalTrabajadores)}</td>
        <td>${escapeHtml(r.lotes)}</td>
        <td>${escapeHtml(r.dia)}</td>
        <td>${escapeHtml(r.hora)}</td>
        <td class="uj-source">${escapeHtml(r._source)}</td>
      </tr>`
    )
    .join("");
}

function renderKpis() {
  if ($("ujKpiSuma")) $("ujKpiSuma").textContent = fmt(state.buckets.suma.rows.length);
  if ($("ujKpiDescuento")) $("ujKpiDescuento").textContent = fmt(state.buckets.descuento.rows.length);
  if ($("ujKpiDescarte")) $("ujKpiDescarte").textContent = fmt(state.buckets.descarte.rows.length);
  const meta = $("ujMeta");
  if (meta) {
    meta.textContent = `Apartado activo: ${TYPES[state.active].label}. Sube uno o varios Excel y descarga el unificado.`;
  }
  const btnExport = $("btnUjExport");
  if (btnExport) btnExport.disabled = !activeBucket().rows.length;
  const btnClear = $("btnUjClear");
  if (btnClear) btnClear.disabled = !activeBucket().files.length;
}

function renderAll() {
  renderTabs();
  renderKpis();
  renderFileList();
  renderTable();
}

function switchTab(key) {
  if (!TYPES[key]) return;
  state.active = key;
  renderAll();
}

function bindUi() {
  document.querySelectorAll("[data-uj-tab]").forEach((btn) => {
    btn.addEventListener("click", () => switchTab(btn.dataset.ujTab));
  });

  const pick = () => $("ujInputExcel")?.click();

  $("btnUjPick")?.addEventListener("click", pick);
  $("btnUjDropPick")?.addEventListener("click", (e) => {
    e.stopPropagation();
    pick();
  });

  $("ujInputExcel")?.addEventListener("change", async () => {
    const input = $("ujInputExcel");
    try {
      await addFiles(input?.files, state.active);
    } catch (err) {
      setStatus("error", err.message || "Error al leer Excel.");
    }
    if (input) input.value = "";
  });

  $("btnUjExport")?.addEventListener("click", () => exportUnified(state.active));
  $("btnUjClear")?.addEventListener("click", () => {
    if (!window.confirm(`¿Vaciar archivos de «${TYPES[state.active].label}»?`)) return;
    clearBucket(state.active);
  });

  const zone = $("ujDropZone");
  if (zone) {
    zone.addEventListener("click", (e) => {
      if (e.target.closest("button")) return;
      pick();
    });
    ;["dragenter", "dragover"].forEach((ev) => {
      zone.addEventListener(ev, (e) => {
        e.preventDefault();
        zone.classList.add("is-dragover");
      });
    });
    ;["dragleave", "drop"].forEach((ev) => {
      zone.addEventListener(ev, (e) => {
        e.preventDefault();
        zone.classList.remove("is-dragover");
      });
    });
    zone.addEventListener("drop", (e) => {
      const files = e.dataTransfer?.files;
      if (files?.length) addFiles(files, state.active).catch((err) => setStatus("error", err.message));
    });
  }
}

let started = false;

function init() {
  if (started) return;
  started = true;
  bindUi();
  renderAll();
}

window.addEventListener("qb:route-changed", (e) => {
  if (e.detail?.route === "unificacion-jarras") init();
});

if (document.querySelector('.view[data-view="unificacion-jarras"]:not([hidden])')) {
  init();
}

})();
