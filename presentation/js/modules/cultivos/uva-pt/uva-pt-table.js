/** Tabla Uva PT — suma tonalidades, sticky, menú columnas */

import {
  getColumnasFront,
  getStickyColsPt,
  getProtectedColIndicesPt
} from "./uva-pt.config.js";
import {
  formatCellDisplay,
  collectRowIncidencias,
  applySumaTonalidadesCell,
  applyFechaYmCell
} from "./uva-pt.validation.js";
import { hydrateLucideIcons } from "../../../utils/lucide-icon.util.js";
import { translateExcelHeader } from "../../../utils/excel-header-i18n.util.js";
import {
  resolveSapZoneHeader,
  SAP_ZONE_HEADER_LABELS_BY_JS
} from "../shared/mp-results-perf.util.js";
import {
  applyPtColumnVisibility,
  bindColumnContextMenu,
  bindCopyableCells,
  syncTableColgroup
} from "../arandano-pt/arandano-pt-table.js";

const WHATSAPP_SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" fill="white" viewBox="0 0 16 16"><path d="M13.601 2.326A7.854 7.854 0 0 0 7.994 0C3.627 0 .068 3.558.064 7.926c0 1.399.366 2.76 1.06 3.973L0 16l4.104-1.076a7.863 7.863 0 0 0 3.89.593c4.365 0 7.923-3.559 7.923-7.928a7.858 7.858 0 0 0-2.316-5.563zM7.994 14.52a6.573 6.573 0 0 1-3.356-.92l-.24-.144-2.494.654.666-2.433-.156-.251a6.56 6.56 0 0 1-1.007-3.505c0-3.626 2.957-6.584 6.591-6.584a6.56 6.56 0 0 1 4.66 1.931 6.557 6.557 0 0 1 1.928 4.66c-.004 3.639-2.961 6.592-6.592 6.592zm3.615-4.934c-.197-.099-1.17-.578-1.353-.646-.182-.065-.315-.099-.445.099-.133.197-.513.646-.627.775-.114.133-.232.148-.43.05-.197-.1-.836-.308-1.592-.985-.59-.525-.985-1.175-1.103-1.372-.114-.198-.011-.304.088-.403.087-.088.197-.232.296-.346.1-.114.133-.198.198-.33.065-.134.034-.248-.015-.347-.05-.099-.445-1.076-.612-1.47-.16-.389-.323-.335-.445-.34-.114-.007-.247-.007-.38-.007a.729.729 0 0 0-.529.247c-.182.198-.691.677-.691 1.654 0 .977.71 1.916.81 2.049.098.133 1.394 2.132 3.383 2.992.47.205.84.326 1.129.418.475.152.904.129 1.246.08.38-.058 1.171-.48 1.338-.943.164-.464.164-.86.114-.943-.049-.084-.182-.133-.38-.232z"/></svg>`;

function applySticky(el, excelCol) {
  if (!getStickyColsPt().includes(excelCol)) return;
  el.classList.add("agv-pt-sticky-col", `agv-pt-sticky-col-${excelCol}`);
}

function createColorBtn(className, title, onClick) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = `agv-pt-row-color-btn ${className}`;
  btn.title = title;
  btn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    onClick();
  });
  return btn;
}

function setRowMark(tr, row, mark) {
  row.__mark = mark;
  tr.classList.remove("agv-pt-row--mark-green", "agv-pt-row--mark-orange");
  if (mark === "green") tr.classList.add("agv-pt-row--mark-green");
  else if (mark === "orange") tr.classList.add("agv-pt-row--mark-orange");
}

function makeCopyableCell(td, rawValue) {
  const text = String(rawValue ?? "").trim();
  if (!text) return;
  td.classList.add("agv-pt-cell-copy");
  td.dataset.copyValue = text;
  td.title = `Clic para copiar: ${text}`;
}

function buildActionsCell(row, tr, onRowMark, onCopyReport) {
  const td = document.createElement("td");
  td.className = "agv-pt-action-cell";

  const wrap = document.createElement("div");
  wrap.className = "agv-pt-action-cell__content";

  const drag = document.createElement("span");
  drag.className = "agv-pt-drag-handle";
  drag.textContent = "☰";
  drag.title = "Arrastrar para reordenar fila";
  drag.draggable = true;
  wrap.appendChild(drag);

  wrap.appendChild(
    createColorBtn("agv-pt-row-color-btn--green", "Marcar fila correcta (verde)", () => {
      setRowMark(tr, row, "green");
      onRowMark?.();
    })
  );
  wrap.appendChild(
    createColorBtn("agv-pt-row-color-btn--orange", "Marcar fila con advertencia (naranja)", () => {
      setRowMark(tr, row, "orange");
      onRowMark?.();
    })
  );

  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.className = "agv-pt-whatsapp-btn";
  copyBtn.title = "Copiar reporte de errores para WhatsApp";
  copyBtn.innerHTML = WHATSAPP_SVG;
  copyBtn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    onCopyReport(row);
  });
  wrap.appendChild(copyBtn);

  td.appendChild(wrap);
  return td;
}

function columnLabel(col, headers) {
  if (col === "Suma Tonalidades") return translateExcelHeader(col, -1);
  if (typeof col === "number" && SAP_ZONE_HEADER_LABELS_BY_JS[col] != null) {
    return resolveSapZoneHeader(col, headers?.[col] || "").label;
  }
  const raw = headers[col] || `Col ${Number(col) + 1}`;
  return translateExcelHeader(raw, typeof col === "number" ? col : -1);
}

function buildPlainHeader(th, excelCol, displayIdx, headers) {
  th.className = "agv-pt-data-header";
  th.dataset.colIndex = String(displayIdx);
  if (typeof excelCol === "number") th.dataset.excelCol = String(excelCol);
  th.dataset.excelHeader =
    excelCol === "Suma Tonalidades" ? excelCol : headers[excelCol] || `Col ${Number(excelCol) + 1}`;
  const label = columnLabel(excelCol, headers);
  th.textContent = label;
  if (SAP_ZONE_HEADER_LABELS_BY_JS[excelCol] != null) {
    th.classList.add("agv-pt-table__col-header--sap");
    th.title = resolveSapZoneHeader(excelCol, headers?.[excelCol] || "").title;
  } else {
    th.title = `${label} — clic para ocultar/mostrar columnas`;
  }
  if (typeof excelCol === "number") applySticky(th, excelCol);
}

export function refreshUvaPtHeaderLabels(headerRow, headers) {
  if (!headerRow) return;
  headerRow.querySelectorAll("th[data-excel-col], th[data-col-index]").forEach((th) => {
    const excelColRaw = th.dataset.excelCol;
    const excelCol =
      excelColRaw != null && excelColRaw !== ""
        ? Number.isFinite(Number(excelColRaw))
          ? Number(excelColRaw)
          : excelColRaw
        : null;
    if (excelCol == null && th.dataset.excelHeader === "Suma Tonalidades") {
      th.textContent = columnLabel("Suma Tonalidades", headers);
      return;
    }
    if (excelCol == null) return;
    const label = columnLabel(excelCol, headers);
    th.dataset.excelHeader =
      typeof excelCol === "number"
        ? headers[excelCol] || `Col ${excelCol + 1}`
        : String(excelCol);
    th.textContent = label;
  });
}

function ensureProtectedColsVisible(tableEl) {
  if (!tableEl?._ptHiddenCols) return;
  getProtectedColIndicesPt().forEach((idx) => tableEl._ptHiddenCols.delete(idx));
}

function renderBodyRows(tbody, rows, headers, tableEl, onRowMark, onCopyReport) {
  tbody.innerHTML = "";
  rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    tr.dataset.rowIndex = String(rowIndex);

    if (row.__mark === "green") tr.classList.add("agv-pt-row--mark-green");
    else if (row.__mark === "orange") tr.classList.add("agv-pt-row--mark-orange");

    tr.appendChild(buildActionsCell(row, tr, onRowMark, onCopyReport));

    getColumnasFront().forEach((col, displayIdx) => {
      const td = document.createElement("td");
      td.dataset.colIndex = String(displayIdx);
      if (typeof col === "number") td.dataset.excelCol = String(col);

      if (col === "Suma Tonalidades") {
        applySumaTonalidadesCell(td, row);
      } else {
        td.textContent = formatCellDisplay(row[col], col);
        if (col === 0 || col === 9) makeCopyableCell(td, row[col]);
        if (typeof col === "number") applyFechaYmCell(td, row, col);
      }
      if (typeof col === "number") applySticky(td, col);
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  if (tableEl?._ptHiddenCols) {
    ensureProtectedColsVisible(tableEl);
    applyPtColumnVisibility(tableEl);
  }
}

function bindRowDragDrop(tbody, rows, onReorder) {
  if (!tbody || tbody.dataset.dragBound === "1") return;
  tbody.dataset.dragBound = "1";

  let dragFrom = null;

  tbody.addEventListener("dragstart", (e) => {
    if (!e.target.closest(".agv-pt-drag-handle")) {
      e.preventDefault();
      return;
    }
    const dragRow = e.target.closest("tr");
    if (!dragRow || !tbody.contains(dragRow)) return;
    dragFrom = [...tbody.children].indexOf(dragRow);
    dragRow.classList.add("agv-pt-row--dragging");
    e.dataTransfer.effectAllowed = "move";
  });

  tbody.addEventListener("dragover", (e) => {
    if (dragFrom == null) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
  });

  tbody.addEventListener("drop", (e) => {
    e.preventDefault();
    if (dragFrom == null) return;
    const tr = e.target.closest("tr");
    if (!tr || !tbody.contains(tr)) return;
    const to = [...tbody.children].indexOf(tr);
    if (to !== dragFrom) onReorder?.(dragFrom, to, rows);
  });

  tbody.addEventListener("dragend", (e) => {
    e.target.closest("tr")?.classList.remove("agv-pt-row--dragging");
    dragFrom = null;
  });
}

export function bindUvaPtColumnMenu(tableEl, menuEl) {
  bindColumnContextMenu(tableEl, menuEl, { protectedColIndices: getProtectedColIndicesPt() });
  if (!tableEl || tableEl.dataset.colClickBound === "1") return;
  tableEl.dataset.colClickBound = "1";
  tableEl.querySelector("thead")?.addEventListener("click", (e) => {
    if (e.target.closest("select, option, .agv-pt-filter-select")) return;
    const th = e.target.closest("th.agv-pt-data-header[data-col-index]");
    if (!th || !tableEl.contains(th)) return;
    e.preventDefault();
    e.stopPropagation();

    const colIndex = Number(th.dataset.colIndex);
    if (!Number.isFinite(colIndex)) return;
    ensureProtectedColsVisible(tableEl);

    const hideBtn = menuEl.querySelector('[data-action="hide"]');
    const showAllBtn = menuEl.querySelector('[data-action="show-all"]');
    const visibleCount = tableEl.querySelectorAll("thead th[data-col-index]:not(.agv-pt-col-hidden)").length;
    const isProtected = getProtectedColIndicesPt().has(colIndex);
    if (hideBtn) {
      hideBtn.disabled = isProtected || tableEl._ptHiddenCols?.has(colIndex) || visibleCount <= 1;
    }
    if (showAllBtn) showAllBtn.disabled = !tableEl._ptHiddenCols?.size;

    if (!menuEl.hidden && menuEl.parentElement === th) {
      menuEl.hidden = true;
      document.getElementById("agv-pt-table-wrap")?.appendChild(menuEl);
      return;
    }
    th.appendChild(menuEl);
    menuEl.hidden = false;
    hydrateLucideIcons(menuEl);
  });
}

export function renderUvaPtTable({
  headerRow,
  bodyRows,
  headers,
  tableEl,
  filteredRows,
  onFilterChange,
  onReorder,
  onCopyReport
}) {
  headerRow.innerHTML = "";

  const thAction = document.createElement("th");
  thAction.textContent = "···";
  thAction.className = "agv-pt-action-header";
  thAction.title = "Acciones";
  headerRow.appendChild(thAction);

  getColumnasFront().forEach((col, displayIdx) => {
    const th = document.createElement("th");
    buildPlainHeader(th, col, displayIdx, headers);
    headerRow.appendChild(th);
  });

  if (tableEl) {
    tableEl.classList.add("agv-pt-table--uva");
    syncTableColgroup(tableEl, getColumnasFront().length);
    if (!tableEl._ptHiddenCols) tableEl._ptHiddenCols = new Set();
    ensureProtectedColsVisible(tableEl);
  }

  const paint = (rows) => {
    renderBodyRows(bodyRows, rows, headers, tableEl, null, onCopyReport);
    bindCopyableCells(bodyRows);
    bindRowDragDrop(bodyRows, rows, onReorder);
  };

  paint(filteredRows);
  onFilterChange?.(filteredRows.length, filteredRows);

  return filteredRows;
}

export function bindTableSearch(inputEl, tbody) {
  if (!inputEl || !tbody) return;
  inputEl.addEventListener("input", () => {
    const term = inputEl.value.trim().toUpperCase();
    tbody.querySelectorAll("tr").forEach((tr) => {
      if (!term) {
        tr.classList.remove("agv-pt-row--search-hidden");
        return;
      }
      let idText = "";
      let loteText = "";
      tr.querySelectorAll("td[data-excel-col]").forEach((td) => {
        if (td.dataset.excelCol === "0") idText = td.textContent.trim().toUpperCase();
        if (td.dataset.excelCol === "9") loteText = td.textContent.trim().toUpperCase();
      });
      tr.classList.toggle(
        "agv-pt-row--search-hidden",
        !(idText.includes(term) || loteText.includes(term))
      );
    });
  });
}

export function buildWhatsappReport(row) {
  const incidencias = collectRowIncidencias(row);
  if (!incidencias.length) return null;
  const lista = incidencias.map((i) => `      • ${i}`).join("\n");
  return `*Usuario:* ${row[6] ?? ""}
*ID:* ${row[0] ?? ""}
*Lote:* ${row[9] ?? ""}
*Incidencias:*
${lista}
*Acción: Corregir inspección por favor.*`;
}
