/**
 * Trabajadores Jarras — Grupo → Trabajadores → Lote → Guías
 * (Sin capa “supervisor”: es el mismo trabajador)
 * Lote auto-llena Módulo + Turno. Persistencia localStorage.
 */

import {
  loadPackedLotes,
  loadPackedGrupos,
  parseDataLicapaMatrix,
  parseGruposFromProduccionMatrix,
  findLoteMeta,
  loteDisplay,
  loteSearchBlob
} from "./jarras-parser.js";

const STORE_KEY = "qb-trabajadores-jarras-v2";

const state = {
  lotes: [],
  grupos: [],
  groups: [],
  openPicker: null, // { type: 'lote'|'grupo', gid, wid? }
  filters: { grupo: "", lote: "", turno: "", search: "" }
};

function $(id) {
  return document.getElementById(id);
}

function uid(prefix = "id") {
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`;
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

function num(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function blankGuia() {
  return { id: uid("g"), jarras: "", jabas: "" };
}

function blankWorker() {
  return {
    id: uid("w"),
    nombre: "",
    lote: "",
    modulo: "",
    turno: "",
    guias: [blankGuia()]
  };
}

function blankGroup(grupo = "") {
  return { id: uid("grp"), grupo: String(grupo || "").trim(), workers: [blankWorker()] };
}

function addGroupFromInput() {
  const input = $("tjAddGrupoInput");
  const name = String(input?.value || "").trim();
  if (!name) {
    setStatus("error", "Escribe el nombre del grupo para agregarlo.");
    input?.focus();
    return;
  }
  const exists = state.groups.some((g) => String(g.grupo || "").toLowerCase() === name.toLowerCase());
  if (exists) {
    setStatus("error", `Ya existe el grupo «${name}».`);
    input?.focus();
    return;
  }
  if (!state.grupos.some((g) => String(g).toLowerCase() === name.toLowerCase())) {
    state.grupos = [...state.grupos, name].sort((a, b) => a.localeCompare(b, "es"));
  }
  clearFilters();
  state.groups.push(blankGroup(name));
  saveStore();
  if (input) input.value = "";
  setStatus("ok", `Grupo «${name}» agregado.`);
  renderAll();
  const host = $("tjSupervisors");
  if (host) host.lastElementChild?.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function migrateOldStore(data) {
  if (Array.isArray(data.groups)) return data.groups;
  if (!Array.isArray(data.supervisors)) return [];
  const groups = [];
  data.supervisors.forEach((s) => {
    (s.groups || []).forEach((g) => {
      const workers = (g.workers || []).map((w) => ({
        ...w,
        nombre: w.nombre || s.nombre || ""
      }));
      if (!workers.length && s.nombre) {
        workers.push({ ...blankWorker(), nombre: s.nombre });
      }
      groups.push({
        id: g.id || uid("grp"),
        grupo: g.grupo || "",
        workers: workers.length ? workers : [blankWorker()]
      });
    });
  });
  return groups;
}

function loadStore() {
  try {
    let raw = localStorage.getItem(STORE_KEY);
    if (!raw) raw = localStorage.getItem("qb-trabajadores-jarras-v1");
    if (!raw) return;
    const data = JSON.parse(raw);
    state.groups = migrateOldStore(data);
  } catch (err) {
    console.warn("[jarras] store", err);
  }
}

function saveStore() {
  try {
    localStorage.setItem(
      STORE_KEY,
      JSON.stringify({ savedAt: new Date().toISOString(), groups: state.groups })
    );
  } catch (err) {
    console.warn("[jarras] save", err);
  }
}

function findPath(gid, wid) {
  const g = state.groups.find((x) => x.id === gid);
  const w = g?.workers?.find((x) => x.id === wid);
  return { g, w };
}

function totalsOf(scope) {
  let jarras = 0;
  let jabas = 0;
  let trabajadores = 0;
  const list = Array.isArray(scope) ? scope : scope ? [scope] : [];
  list.forEach((g) => {
    (g.workers || []).forEach((w) => {
      if (String(w.nombre || "").trim()) trabajadores += 1;
      (w.guias || []).forEach((guia) => {
        jarras += num(guia.jarras);
        jabas += num(guia.jabas);
      });
    });
  });
  return { jarras, jabas, trabajadores };
}

function workerTotals(w) {
  return (w.guias || []).reduce(
    (acc, g) => {
      acc.jarras += num(g.jarras);
      acc.jabas += num(g.jabas);
      return acc;
    },
    { jarras: 0, jabas: 0 }
  );
}

function setStatus(kind, text) {
  const el = $("tjStatus");
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

let confirmResolver = null;

function closeTjConfirm(result) {
  const modal = $("modalTjConfirm");
  if (modal) modal.hidden = true;
  const resolve = confirmResolver;
  confirmResolver = null;
  if (resolve) resolve(Boolean(result));
}

function confirmTj({
  title = "¿Seguro que deseas eliminar?",
  text = "Esta acción no se puede deshacer.",
  confirmLabel = "Sí, eliminar"
} = {}) {
  const modal = $("modalTjConfirm");
  if (!modal) return Promise.resolve(window.confirm(`${title}\n${text}`));
  if (confirmResolver) closeTjConfirm(false);
  const titleEl = $("tjConfirmTitle");
  const textEl = $("tjConfirmText");
  const okBtn = $("btnTjConfirmOk");
  if (titleEl) titleEl.textContent = title;
  if (textEl) textEl.textContent = text;
  if (okBtn) okBtn.textContent = confirmLabel;
  modal.hidden = false;
  queueMicrotask(() => okBtn?.focus());
  return new Promise((resolve) => {
    confirmResolver = resolve;
  });
}

function renderKpis() {
  const t = totalsOf(state.groups);
  if ($("tjKpiSup")) $("tjKpiSup").textContent = fmt(state.groups.length);
  if ($("tjKpiTrab")) $("tjKpiTrab").textContent = fmt(t.trabajadores);
  if ($("tjKpiJarras")) $("tjKpiJarras").textContent = fmt(t.jarras);
  if ($("tjKpiJabas")) $("tjKpiJabas").textContent = fmt(t.jabas);
  const meta = $("tjMeta");
  if (meta) {
    const visible = filteredGroups().length;
    const total = state.groups.length;
    const filterNote =
      hasActiveFilters() && total ? ` · mostrando ${fmt(visible)} de ${fmt(total)}` : "";
    meta.textContent = `${fmt(state.lotes.length)} lotes · ${fmt(state.grupos.length)} grupos · guardado aquí${filterNote}`;
  }
}

function hasActiveFilters() {
  const f = state.filters;
  return Boolean(f.grupo || f.lote || f.turno || String(f.search || "").trim());
}

function groupMatches(group) {
  const f = state.filters;
  const workers = Array.isArray(group.workers) ? group.workers : [];
  if (f.grupo && group.grupo !== f.grupo) return false;
  if (f.lote && !workers.some((w) => w.lote === f.lote)) return false;
  if (f.turno && !workers.some((w) => String(w.turno || "") === f.turno)) return false;
  const q = String(f.search || "").trim().toLowerCase();
  if (q) {
    const inGrupo = String(group.grupo || "").toLowerCase().includes(q);
    const inWorker = workers.some((w) => String(w.nombre || "").toLowerCase().includes(q));
    if (!inGrupo && !inWorker) return false;
  }
  return true;
}

function filteredGroups() {
  if (!hasActiveFilters()) return state.groups;
  return state.groups.filter(groupMatches);
}

function fillSelect(el, values, current, emptyLabel = "Todos") {
  if (!el) return;
  const opts = [`<option value="">${escapeHtml(emptyLabel)}</option>`].concat(
    values.map((v) => {
      const val = String(v);
      const sel = val === current ? " selected" : "";
      return `<option value="${escapeHtml(val)}"${sel}>${escapeHtml(val)}</option>`;
    })
  );
  el.innerHTML = opts.join("");
}

function syncFilterControls() {
  const f = state.filters;
  const grupoSet = new Set();
  const loteSet = new Set();
  const turnoSet = new Set();
  state.groups.forEach((g) => {
    if (g.grupo) grupoSet.add(g.grupo);
    (g.workers || []).forEach((w) => {
      if (w.lote) loteSet.add(w.lote);
      if (w.turno) turnoSet.add(String(w.turno));
    });
  });
  state.grupos.forEach((g) => {
    if (g) grupoSet.add(g);
  });
  const lotesSorted = [...loteSet].sort((a, b) => String(a).localeCompare(String(b), "es"));
  const lotesLabels = lotesSorted.map((id) => {
    const meta = findLoteMeta(state.lotes, id);
    return { id, label: meta ? loteDisplay(meta) : id };
  });
  const selGrupo = $("tjFltGrupo");
  const selLote = $("tjFltLote");
  const selTurno = $("tjFltTurno");
  fillSelect(selGrupo, [...grupoSet].sort((a, b) => a.localeCompare(b, "es")), f.grupo);
  if (selLote) {
    const opts = [`<option value="">Todos</option>`].concat(
      lotesLabels.map(({ id, label }) => {
        const sel = id === f.lote ? " selected" : "";
        return `<option value="${escapeHtml(id)}"${sel}>${escapeHtml(label)}</option>`;
      })
    );
    selLote.innerHTML = opts.join("");
  }
  fillSelect(selTurno, [...turnoSet].sort((a, b) => a.localeCompare(b, "es")), f.turno);
  const search = $("tjFltSearch");
  if (search && search.value !== f.search) search.value = f.search;
}

function readFiltersFromUi() {
  state.filters = {
    grupo: $("tjFltGrupo")?.value || "",
    lote: $("tjFltLote")?.value || "",
    turno: $("tjFltTurno")?.value || "",
    search: $("tjFltSearch")?.value || ""
  };
}

function clearFilters() {
  state.filters = { grupo: "", lote: "", turno: "", search: "" };
  const search = $("tjFltSearch");
  if (search) search.value = "";
  syncFilterControls();
}

function renderGuias(gid, wid, worker) {
  const guias = worker.guias || [];
  const rows = guias
    .map((guia, idx) => {
      const isLast = idx === guias.length - 1;
      return `<div class="tj-guia">
        <div class="tj-pair" title="Jarras y jabas (van juntas)">
          <input type="number" min="0" step="1" inputmode="numeric" placeholder="Jarras"
            data-act="guia-jarras" data-gid="${escapeHtml(gid)}" data-wid="${escapeHtml(wid)}" data-guia="${escapeHtml(guia.id)}"
            value="${escapeHtml(guia.jarras)}" aria-label="Jarras ${idx + 1}" />
          <input type="number" min="0" step="1" inputmode="numeric" placeholder="Jabas"
            data-act="guia-jabas" data-gid="${escapeHtml(gid)}" data-wid="${escapeHtml(wid)}" data-guia="${escapeHtml(guia.id)}"
            value="${escapeHtml(guia.jabas)}" aria-label="Jabas ${idx + 1}" />
        </div>
        <div class="tj-mini">
          <button type="button" class="tj-icon-btn ${isLast ? "" : "is-ghost"}" title="Otra guía" data-act="add-guia"
            data-gid="${escapeHtml(gid)}" data-wid="${escapeHtml(wid)}" ${isLast ? "" : "tabindex='-1' aria-hidden='true'"}>+</button>
          <button type="button" class="tj-icon-btn tj-icon-btn--danger" title="Quitar guía" data-act="rm-guia"
            data-gid="${escapeHtml(gid)}" data-wid="${escapeHtml(wid)}" data-guia="${escapeHtml(guia.id)}">−</button>
        </div>
      </div>`;
    })
    .join("");
  const tot = workerTotals(worker);
  return `<div class="tj-guias">
    ${rows}
    <span class="tj-guia-total">${fmt(tot.jarras)} j · ${fmt(tot.jabas)} b</span>
  </div>`;
}

function renderWorker(gid, worker, { showRemove = true } = {}) {
  const meta = findLoteMeta(state.lotes, worker.lote);
  const loteTxt = worker.lote ? loteDisplay(meta || { lote: worker.lote }) : "Buscar lote…";
  return `<div class="tj-worker" data-worker="${escapeHtml(worker.id)}">
    <div class="tj-combo" data-combo="lote">
      <button type="button" class="tj-combo__btn ${worker.lote ? "has-value" : ""}" data-act="open-lote"
        data-gid="${escapeHtml(gid)}" data-wid="${escapeHtml(worker.id)}">
        <span>${escapeHtml(loteTxt)}</span><span aria-hidden="true">▾</span>
      </button>
      <div class="tj-combo__panel" hidden data-panel="lote:${escapeHtml(gid)}:${escapeHtml(worker.id)}">
        <input type="search" class="tj-combo__search" placeholder="Buscar lote…" autocomplete="off"
          data-act="lote-search" data-gid="${escapeHtml(gid)}" data-wid="${escapeHtml(worker.id)}" />
        <div class="tj-combo__list" data-list="lote:${escapeHtml(gid)}:${escapeHtml(worker.id)}"></div>
      </div>
    </div>
    <input class="tj-readonly" type="text" disabled value="${escapeHtml(worker.modulo || "—")}" title="Módulo" aria-label="Módulo" />
    <input class="tj-readonly" type="text" disabled value="${escapeHtml(worker.turno || "—")}" title="Turno" aria-label="Turno" />
    ${renderGuias(gid, worker.id, worker)}
    ${
      showRemove
        ? `<button type="button" class="tj-icon-btn tj-icon-btn--danger" title="Quitar trabajador" data-act="rm-worker"
      data-gid="${escapeHtml(gid)}" data-wid="${escapeHtml(worker.id)}">−</button>`
        : `<span></span>`
    }
  </div>`;
}

function renderGroup(group, isLast) {
  if (!Array.isArray(group.workers) || !group.workers.length) {
    group.workers = [blankWorker()];
  }
  const primary = group.workers[0];
  const rest = group.workers.slice(1);
  const tot = totalsOf(group);
  const grupoTxt = group.grupo || "Buscar grupo…";
  const restHtml = rest
    .map((w) => {
      return `<div class="tj-worker-block">
      <div class="tj-worker-block__name">
        <input type="text" placeholder="Trabajador" value="${escapeHtml(w.nombre)}"
          data-act="worker-nombre" data-gid="${escapeHtml(group.id)}" data-wid="${escapeHtml(w.id)}" />
        <div class="tj-mini">
          <button type="button" class="tj-icon-btn" title="Agregar trabajador" data-act="add-worker" data-gid="${escapeHtml(group.id)}">+</button>
          <button type="button" class="tj-icon-btn tj-icon-btn--danger" title="Quitar" data-act="rm-worker"
            data-gid="${escapeHtml(group.id)}" data-wid="${escapeHtml(w.id)}">−</button>
        </div>
      </div>
      ${renderWorker(group.id, w, { showRemove: false })}
    </div>`;
    })
    .join("");

  return `<section class="tj-wrap" data-group="${escapeHtml(group.id)}">
    <div class="tj-wrap__main">
      <header class="tj-wrap__head">
        <div class="tj-head-row">
          <div class="tj-combo tj-combo--grupo" data-combo="grupo">
            <button type="button" class="tj-combo__btn ${group.grupo ? "has-value" : ""}" data-act="open-grupo"
              data-gid="${escapeHtml(group.id)}">
              <span>${escapeHtml(grupoTxt)}</span><span aria-hidden="true">▾</span>
            </button>
            <div class="tj-combo__panel" hidden data-panel="grupo:${escapeHtml(group.id)}">
              <input type="search" class="tj-combo__search" placeholder="Buscar grupo…" autocomplete="off"
                data-act="grupo-search" data-gid="${escapeHtml(group.id)}" />
              <div class="tj-combo__list" data-list="grupo:${escapeHtml(group.id)}"></div>
            </div>
          </div>
          <input type="text" class="tj-head-worker" placeholder="Trabajador"
            value="${escapeHtml(primary.nombre)}"
            data-act="worker-nombre" data-gid="${escapeHtml(group.id)}" data-wid="${escapeHtml(primary.id)}" />
          <span class="tj-wrap__sum">${fmt(tot.jarras)} j · ${fmt(tot.jabas)} b</span>
        </div>
      </header>
      <div class="tj-wrap__cols" aria-hidden="true">
        <span>Lote</span><span>Módulo</span><span>Turno</span><span>Jarras · Jabas</span><span></span>
      </div>
      <div class="tj-wrap__body">
        ${renderWorker(group.id, primary, { showRemove: false })}
        ${restHtml}
      </div>
    </div>
    <aside class="tj-wrap__actions" aria-label="Acciones del grupo">
      <button type="button" class="btn btn--ghost btn--sm tj-side-btn" data-act="export-excel-group" data-gid="${escapeHtml(group.id)}">Excel</button>
      <button type="button" class="btn btn--ghost btn--sm tj-side-btn" data-act="export-pdf-group" data-gid="${escapeHtml(group.id)}">PDF</button>
      <button type="button" class="btn btn--ghost btn--sm tj-side-btn tj-side-btn--danger" data-act="rm-group" data-gid="${escapeHtml(group.id)}" title="Eliminar este grupo">− Grupo</button>
      ${
        isLast
          ? `<button type="button" class="btn btn--primary btn--sm tj-side-btn" data-act="add-group">+ Grupo</button>`
          : ""
      }
    </aside>
  </section>`;
}

function renderAll() {
  const host = $("tjSupervisors");
  if (!host) return;
  syncFilterControls();
  if (!state.groups.length) {
    host.innerHTML = `<div class="tj-empty-box">
      <p class="tj-empty-hint">Aún no hay grupos.</p>
      <button type="button" class="btn btn--primary btn--sm" data-act="add-group">+ Grupo</button>
    </div>`;
    renderKpis();
    return;
  }
  const visible = filteredGroups();
  if (!visible.length) {
    host.innerHTML = `<div class="tj-empty-box">
      <p class="tj-empty-hint">Ningún grupo coincide con el filtro.</p>
      <button type="button" class="btn btn--ghost btn--sm" id="btnTjFltEmptyClear">Limpiar filtros</button>
    </div>`;
    $("btnTjFltEmptyClear")?.addEventListener("click", () => {
      clearFilters();
      renderAll();
    });
    renderKpis();
    return;
  }
  const lastId = visible[visible.length - 1]?.id;
  host.innerHTML = visible.map((g) => renderGroup(g, g.id === lastId)).join("");
  renderKpis();
}

function closePickers() {
  document.querySelectorAll(".tj-combo__panel").forEach((p) => {
    p.hidden = true;
  });
  state.openPicker = null;
}

function fillLoteList(gid, wid, query = "") {
  const key = `lote:${gid}:${wid}`;
  const list = document.querySelector(`[data-list="${key}"]`);
  if (!list) return;
  const q = String(query || "").trim().toLowerCase();
  const rows = state.lotes.filter((l) => !q || loteSearchBlob(l).includes(q));
  if (!rows.length) {
    list.innerHTML = `<p class="tj-empty-hint">Sin lotes</p>`;
    return;
  }
  list.innerHTML = rows
    .slice(0, 100)
    .map((l) => {
      const title = loteDisplay(l);
      const sub = `${l.modulo || ""} · T${l.turno || ""}`;
      return `<button type="button" class="tj-combo__opt" data-act="pick-lote"
        data-gid="${escapeHtml(gid)}" data-wid="${escapeHtml(wid)}" data-lote="${escapeHtml(l.lote)}">
        <strong>${escapeHtml(title)}</strong>
        <small>${escapeHtml(sub)}</small>
      </button>`;
    })
    .join("");
}

function fillGrupoList(gid, query = "") {
  const key = `grupo:${gid}`;
  const list = document.querySelector(`[data-list="${key}"]`);
  if (!list) return;
  const q = String(query || "").trim().toLowerCase();
  const rows = state.grupos.filter((g) => !q || g.toLowerCase().includes(q));
  if (!rows.length) {
    list.innerHTML = `<p class="tj-empty-hint">Sin grupos</p>`;
    return;
  }
  list.innerHTML = rows
    .map(
      (g) => `<button type="button" class="tj-combo__opt" data-act="pick-grupo"
        data-gid="${escapeHtml(gid)}" data-grupo="${escapeHtml(g)}">
        <strong>${escapeHtml(g)}</strong>
      </button>`
    )
    .join("");
}

function openPicker(type, gid, wid, btn) {
  closePickers();
  const key = type === "lote" ? `lote:${gid}:${wid}` : `grupo:${gid}`;
  const panel = document.querySelector(`[data-panel="${key}"]`);
  if (!panel) return;
  panel.hidden = false;
  state.openPicker = { type, gid, wid };
  if (type === "lote") fillLoteList(gid, wid, "");
  else fillGrupoList(gid, "");
  window.setTimeout(() => panel.querySelector(".tj-combo__search")?.focus(), 20);
}

function applyLote(gid, wid, loteId) {
  const { w } = findPath(gid, wid);
  if (!w) return;
  const meta = findLoteMeta(state.lotes, loteId);
  w.lote = meta?.lote || loteId;
  w.modulo = meta?.modulo || "";
  w.turno = meta?.turno || "";
  saveStore();
  closePickers();
  renderAll();
}

function applyGrupo(gid, value) {
  const { g } = findPath(gid);
  if (!g) return;
  g.grupo = value;
  saveStore();
  closePickers();
  renderAll();
}

function flatRows(onlyGid = null) {
  const rows = [];
  const groups = onlyGid
    ? state.groups.filter((g) => g.id === onlyGid)
    : state.groups;
  groups.forEach((g) => {
    (g.workers || []).forEach((w) => {
      const tot = workerTotals(w);
      const meta = findLoteMeta(state.lotes, w.lote);
      rows.push({
        Grupo: g.grupo || "",
        Trabajador: w.nombre || "",
        Lote: loteDisplay(meta || { lote: w.lote }),
        Modulo: w.modulo || "",
        Turno: w.turno || "",
        Jarras: tot.jarras,
        Jabas: tot.jabas
      });
    });
  });
  return rows;
}

function exportExcel(onlyGid = null) {
  const rows = flatRows(onlyGid);
  if (!rows.length) {
    setStatus("empty", "No hay datos para exportar.");
    return;
  }
  if (!window.XLSX) {
    setStatus("error", "No se cargó el exportador Excel.");
    return;
  }
  const g = onlyGid ? state.groups.find((x) => x.id === onlyGid) : null;
  const tag = g?.grupo ? g.grupo.replace(/\s+/g, "-") : "todos";
  const wb = window.XLSX.utils.book_new();
  const ws = window.XLSX.utils.json_to_sheet(rows);
  window.XLSX.utils.book_append_sheet(wb, ws, "Jarras");
  window.XLSX.writeFile(wb, `trabajadores-jarras-${tag}-${new Date().toISOString().slice(0, 10)}.xlsx`);
  setStatus("", "");
}

function exportPdf(onlyGid = null) {
  const rows = flatRows(onlyGid);
  if (!rows.length) {
    setStatus("empty", "No hay datos para exportar.");
    return;
  }
  const JsPDF = window.jspdf?.jsPDF;
  if (!JsPDF) {
    setStatus("error", "No se cargó el exportador PDF.");
    return;
  }
  const g = onlyGid ? state.groups.find((x) => x.id === onlyGid) : null;
  const tag = g?.grupo ? g.grupo.replace(/\s+/g, "-") : "todos";
  const doc = new JsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  doc.setFontSize(14);
  doc.text(g?.grupo ? `QBerries · ${g.grupo}` : "QBerries · Trabajadores Jarras", 40, 36);
  doc.setFontSize(9);
  doc.text(`Generado ${new Date().toLocaleString("es-PE")}`, 40, 52);
  const body = rows.map((r) => [r.Grupo, r.Trabajador, r.Lote, r.Modulo, r.Turno, String(r.Jarras), String(r.Jabas)]);
  if (typeof doc.autoTable === "function") {
    doc.autoTable({
      startY: 64,
      head: [["Grupo", "Trabajador", "Lote", "Módulo", "Turno", "Jarras", "Jabas"]],
      body,
      styles: { fontSize: 8, cellPadding: 3 },
      headStyles: { fillColor: [20, 90, 52] }
    });
  }
  doc.save(`trabajadores-jarras-${tag}-${new Date().toISOString().slice(0, 10)}.pdf`);
  setStatus("", "");
}

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("No se pudo leer el archivo"));
    reader.onload = () => {
      try {
        const wb = window.XLSX.read(reader.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        resolve(window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }));
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

async function onUploadLotes(file) {
  if (!file) return;
  setStatus("loading", "Leyendo lotes…");
  try {
    state.lotes = parseDataLicapaMatrix(await readExcelFile(file));
    if (!state.lotes.length) throw new Error("No se encontraron lotes.");
    setStatus("", "");
    renderAll();
  } catch (err) {
    setStatus("error", err?.message || "Error al leer lotes.");
  }
}

async function onUploadGrupos(file) {
  if (!file) return;
  setStatus("loading", "Leyendo grupos…");
  try {
    state.grupos = parseGruposFromProduccionMatrix(await readExcelFile(file));
    if (!state.grupos.length) throw new Error("No se encontraron grupos.");
    setStatus("", "");
    renderAll();
  } catch (err) {
    setStatus("error", err?.message || "Error al leer grupos.");
  }
}

function onClick(e) {
  const btn = e.target.closest("[data-act]");
  if (!btn) return;
  const act = btn.dataset.act;
  const gid = btn.dataset.gid;
  const wid = btn.dataset.wid;
  const guia = btn.dataset.guia;

  if (act === "add-group") {
    clearFilters();
    state.groups.push(blankGroup());
    saveStore();
    renderAll();
    const host = $("tjSupervisors");
    if (host) host.lastElementChild?.scrollIntoView({ behavior: "smooth", block: "nearest" });
    return;
  }
  if (act === "export-excel-group") {
    exportExcel(gid);
    return;
  }
  if (act === "export-pdf-group") {
    exportPdf(gid);
    return;
  }
  if (act === "rm-group") {
    const { g } = findPath(gid);
    const label = g?.grupo ? `«${g.grupo}»` : "este grupo";
    confirmTj({
      title: "¿Eliminar grupo?",
      text: `Se eliminará ${label} con sus trabajadores y guías. No se puede deshacer.`,
      confirmLabel: "Sí, eliminar grupo"
    }).then((ok) => {
      if (!ok) return;
      state.groups = state.groups.filter((x) => x.id !== gid);
      saveStore();
      renderAll();
    });
    return;
  }
  if (act === "add-worker") {
    const { g } = findPath(gid);
    g?.workers?.push(blankWorker());
    saveStore();
    renderAll();
    return;
  }
  if (act === "rm-worker") {
    const { g, w } = findPath(gid, wid);
    if (!g || !w) return;
    const name = String(w.nombre || "").trim();
    confirmTj({
      title: "¿Eliminar trabajador?",
      text: name
        ? `Se quitará a «${name}» de este grupo, con sus lotes y guías.`
        : "Se quitará este trabajador del grupo, con sus lotes y guías.",
      confirmLabel: "Sí, eliminar"
    }).then((ok) => {
      if (!ok) return;
      g.workers = (g.workers || []).filter((x) => x.id !== wid);
      if (!g.workers.length) g.workers = [blankWorker()];
      saveStore();
      renderAll();
    });
    return;
  }
  if (act === "add-guia") {
    if (btn.classList.contains("is-ghost")) return;
    addGuiaRow(gid, wid, { focus: true });
    return;
  }
  if (act === "rm-guia") {
    const { w } = findPath(gid, wid);
    if (!w) return;
    const gRow = (w.guias || []).find((x) => x.id === guia);
    const hasData = gRow && (num(gRow.jarras) > 0 || num(gRow.jabas) > 0 || String(gRow.jarras || gRow.jabas || "").trim());
    const runRemove = () => {
      w.guias = (w.guias || []).filter((x) => x.id !== guia);
      if (!w.guias.length) w.guias = [blankGuia()];
      saveStore();
      renderAll();
    };
    if (!hasData && (w.guias || []).length <= 1) {
      runRemove();
      return;
    }
    confirmTj({
      title: "¿Eliminar guía?",
      text: hasData
        ? "Esta guía tiene jarras o jabas registradas. ¿Seguro que deseas eliminarla?"
        : "Se eliminará esta fila de guía.",
      confirmLabel: "Sí, eliminar guía"
    }).then((ok) => {
      if (ok) runRemove();
    });
    return;
  }
  if (act === "open-lote") {
    e.stopPropagation();
    openPicker("lote", gid, wid, btn);
    return;
  }
  if (act === "open-grupo") {
    e.stopPropagation();
    openPicker("grupo", gid, null, btn);
    return;
  }
  if (act === "pick-lote") {
    applyLote(gid, wid, btn.dataset.lote);
    return;
  }
  if (act === "pick-grupo") {
    applyGrupo(gid, btn.dataset.grupo);
  }
}

function onKeydown(e) {
  if (e.key !== "Enter") return;
  const el = e.target;
  if (!(el instanceof HTMLElement)) return;
  const act = el.dataset.act;
  // Enter en Jarras → nueva guía abajo y foco listo para seguir tecleando
  if (act === "guia-jarras") {
    e.preventDefault();
    const gid = el.dataset.gid;
    const wid = el.dataset.wid;
    const guiaId = el.dataset.guia;
    const { w } = findPath(gid, wid);
    const guia = w?.guias?.find((x) => x.id === guiaId);
    if (guia) guia.jarras = el.value;
    saveStore();
    addGuiaRow(gid, wid, { focus: true });
  }
}

function addGuiaRow(gid, wid, { focus = false } = {}) {
  const { w } = findPath(gid, wid);
  if (!w) return;
  if (!Array.isArray(w.guias)) w.guias = [];
  w.guias.push(blankGuia());
  saveStore();
  renderAll();
  if (!focus) return;
  queueMicrotask(() => {
    const wrap = document.querySelector(`.tj-wrap[data-group="${CSS.escape(gid)}"]`);
    const worker = wrap?.querySelector(`[data-worker="${CSS.escape(wid)}"]`);
    const inputs = worker?.querySelectorAll('[data-act="guia-jarras"]');
    const last = inputs?.[inputs.length - 1];
    last?.focus();
    last?.select?.();
  });
}

function onInput(e) {
  const el = e.target;
  if (!(el instanceof HTMLElement)) return;
  const act = el.dataset.act;
  if (!act) return;
  const gid = el.dataset.gid;
  const wid = el.dataset.wid;
  const guiaId = el.dataset.guia;

  if (act === "worker-nombre") {
    const { w } = findPath(gid, wid);
    if (w) w.nombre = el.value;
    saveStore();
    renderKpis();
    return;
  }
  if (act === "guia-jarras" || act === "guia-jabas") {
    const { w } = findPath(gid, wid);
    const guia = w?.guias?.find((x) => x.id === guiaId);
    if (!guia) return;
    if (act === "guia-jarras") guia.jarras = el.value;
    else guia.jabas = el.value;
    saveStore();
    const card = el.closest(".tj-worker");
    if (card && w) {
      const tot = workerTotals(w);
      const totEl = card.querySelector(".tj-guia-total");
      if (totEl) totEl.textContent = `${fmt(tot.jarras)} j · ${fmt(tot.jabas)} b`;
    }
    const wrap = el.closest(".tj-wrap");
    if (wrap) {
      const { g } = findPath(gid);
      const sum = wrap.querySelector(".tj-wrap__sum");
      if (sum && g) {
        const t = totalsOf(g);
        sum.textContent = `${fmt(t.jarras)} jarras · ${fmt(t.jabas)} jabas`;
      }
    }
    renderKpis();
    return;
  }
  if (act === "lote-search") {
    fillLoteList(gid, wid, el.value);
    return;
  }
  if (act === "grupo-search") {
    fillGrupoList(gid, el.value);
  }
}

function bindUi() {
  $("btnTjExportExcel")?.addEventListener("click", () => exportExcel());
  $("btnTjExportPdf")?.addEventListener("click", () => exportPdf());
  $("btnTjClear")?.addEventListener("click", () => {
    confirmTj({
      title: "¿Vaciar todo el registro?",
      text: "Se borrarán todos los grupos guardados en este dispositivo. No se puede deshacer.",
      confirmLabel: "Sí, vaciar todo"
    }).then((ok) => {
      if (!ok) return;
      state.groups = [];
      clearFilters();
      saveStore();
      renderAll();
    });
  });
  $("btnTjUploadLotes")?.addEventListener("click", () => $("tjInputLotes")?.click());
  $("btnTjUploadGrupos")?.addEventListener("click", () => $("tjInputGrupos")?.click());
  $("tjInputLotes")?.addEventListener("change", () => {
    onUploadLotes($("tjInputLotes").files?.[0]);
    $("tjInputLotes").value = "";
  });
  $("tjInputGrupos")?.addEventListener("change", () => {
    onUploadGrupos($("tjInputGrupos").files?.[0]);
    $("tjInputGrupos").value = "";
  });

  const onFilterChange = () => {
    readFiltersFromUi();
    renderAll();
  };
  $("tjFltGrupo")?.addEventListener("change", onFilterChange);
  $("tjFltLote")?.addEventListener("change", onFilterChange);
  $("tjFltTurno")?.addEventListener("change", onFilterChange);
  $("tjFltSearch")?.addEventListener("input", onFilterChange);
  $("btnTjFltClear")?.addEventListener("click", () => {
    clearFilters();
    renderAll();
  });
  $("btnTjAddGrupo")?.addEventListener("click", () => addGroupFromInput());
  $("tjAddGrupoInput")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      addGroupFromInput();
    }
  });

  $("btnTjConfirmOk")?.addEventListener("click", () => closeTjConfirm(true));
  $("btnTjConfirmCancel")?.addEventListener("click", () => closeTjConfirm(false));
  document.querySelectorAll('[data-close-modal="tj-confirm"]').forEach((el) => {
    el.addEventListener("click", () => closeTjConfirm(false));
  });
  document.addEventListener("keydown", (e) => {
    const modal = $("modalTjConfirm");
    if (!modal || modal.hidden) return;
    if (e.key === "Escape") {
      e.preventDefault();
      closeTjConfirm(false);
    } else if (e.key === "Enter") {
      e.preventDefault();
      closeTjConfirm(true);
    }
  });

  $("tjSupervisors")?.addEventListener("click", onClick);
  $("tjSupervisors")?.addEventListener("input", onInput);
  $("tjSupervisors")?.addEventListener("keydown", onKeydown);
  document.addEventListener("click", (e) => {
    if (!e.target.closest(".tj-combo")) closePickers();
  });
}

let started = false;

async function init() {
  if (started) return;
  started = true;
  loadStore();
  bindUi();
  try {
    const [lotes, grupos] = await Promise.all([loadPackedLotes(), loadPackedGrupos()]);
    state.lotes = lotes;
    state.grupos = grupos;
  } catch (err) {
    console.warn("[jarras] catalogs", err);
    setStatus("error", "No se cargaron catálogos. Usa Actualizar lotes / grupos.");
  }
  renderAll();
}

window.addEventListener("qb:route-changed", (evt) => {
  if (evt.detail?.route === "trabajadores-jarras") init();
});

if (window.location.hash.replace(/^#\/?/, "").split("/")[0] === "trabajadores-jarras") {
  init();
}
