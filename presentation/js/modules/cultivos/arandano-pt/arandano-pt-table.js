/** Renderizado tabla PT — acciones, validación visual, drag & drop. */

import { i18nService } from "../../../services/i18n.service.js";
import { hydrateLucideIcons } from "../../../utils/lucide-icon.util.js";
import { translateExcelHeader } from "../../../utils/excel-header-i18n.util.js";
import { refreshTranslatedHeaderRow } from "../../../utils/table-header-i18n.util.js";
import {
  CALIBRES_MAP,
  CATEGORIAS_FRUTIST,
  VAR_MAP,
  usaCalibreEspecial,
  esSubgrupoNA,
  extraerTrazabilidad,
  getJulianoFromDate
} from "./arandano-pt.catalogs.js";
import {
  buildCalibreCalculado,
  formatCellValue,
  formatSubgrupoCellValue,
  applyDestinoValidation
} from "./arandano-pt.validation.js";
import {
  paintPtYearMonthMismatch,
  PT_FECHA_YM_MSG
} from "../shared/pt-fecha-ym.util.js";

const PALABRAS_PROHIBIDAS_REGULAR = [
  "JUMBO", "MIXED", "REGULAR", "SUPER JUMBO", "MEDIUM", "EXTRA JUMBO", "MIXTO", "NO COMBINADO", "SIN CALIBRAR"
];

const WHATSAPP_SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" fill="white" viewBox="0 0 16 16"><path d="M13.601 2.326A7.854 7.854 0 0 0 7.994 0C3.627 0 .068 3.558.064 7.926c0 1.399.366 2.76 1.06 3.973L0 16l4.104-1.076a7.863 7.863 0 0 0 3.89.593c4.365 0 7.923-3.559 7.923-7.928a7.858 7.858 0 0 0-2.316-5.563zM7.994 14.52a6.573 6.573 0 0 1-3.356-.92l-.24-.144-2.494.654.666-2.433-.156-.251a6.56 6.56 0 0 1-1.007-3.505c0-3.626 2.957-6.584 6.591-6.584a6.56 6.56 0 0 1 4.66 1.931 6.557 6.557 0 0 1 1.928 4.66c-.004 3.639-2.961 6.592-6.592 6.592zm3.615-4.934c-.197-.099-1.17-.578-1.353-.646-.182-.065-.315-.099-.445.099-.133.197-.513.646-.627.775-.114.133-.232.148-.43.05-.197-.1-.836-.308-1.592-.985-.59-.525-.985-1.175-1.103-1.372-.114-.198-.011-.304.088-.403.087-.088.197-.232.296-.346.1-.114.133-.198.198-.33.065-.134.034-.248-.015-.347-.05-.099-.445-1.076-.612-1.47-.16-.389-.323-.335-.445-.34-.114-.007-.247-.007-.38-.007a.729.729 0 0 0-.529.247c-.182.198-.691.677-.691 1.654 0 .977.71 1.916.81 2.049.098.133 1.394 2.132 3.383 2.992.47.205.84.326 1.129.418.475.152.904.129 1.246.08.38-.058 1.171-.48 1.338-.943.164-.464.164-.86.114-.943-.049-.084-.182-.133-.38-.232z"/></svg>`;

function paintRedText(td, text) {
  td.textContent = "";
  for (const ch of String(text || "ERROR")) {
    const span = document.createElement("span");
    span.textContent = ch;
    span.className = "agv-pt-cell-error-char";
    td.appendChild(span);
  }
}

function paintTrazabilidadCell(td, rawValue, cliente) {
  const text = String(rawValue ?? "");
  td.textContent = "";
  if (!text) return;

  // Desactivado por ahora: resaltado letra E (Sekoya Pop Orgánica)
  // const traz = extraerTrazabilidad(text);
  // const varNombre = traz ? VAR_MAP[traz.variedad]?.[0] || "" : "";
  // const sekoyaBad = varNombre === "Sekoya Pop Orgánica" && text.length > 4 && text[4] !== "E";
  td.textContent = text;
}

function createColorBtn(className, title, onClick) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = `agv-pt-row-color-btn ${className}`;
  btn.title = title;
  btn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    onClick(e);
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
  const prev = td.title;
  td.title = prev ? `${prev} · Clic para copiar` : `Clic para copiar: ${text}`;
}

export function bindCopyableCells(tbody) {
  if (!tbody || tbody.dataset.copyBound === "1") return;
  tbody.dataset.copyBound = "1";
  tbody.addEventListener("click", async (e) => {
    const td = e.target.closest("td.agv-pt-cell-copy");
    if (!td) return;
    const text = td.dataset.copyValue || td.textContent.trim();
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      td.classList.add("agv-pt-cell-copy--flash");
      window.setTimeout(() => td.classList.remove("agv-pt-cell-copy--flash"), 450);
    } catch {
      /* portapapeles no disponible */
    }
  });
}

function buildActionsCell(row, tr, onCopyReport, onRowMark) {
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
      onRowMark?.(row);
    })
  );
  wrap.appendChild(
    createColorBtn("agv-pt-row-color-btn--orange", "Marcar fila con advertencia (naranja)", () => {
      setRowMark(tr, row, "orange");
      onRowMark?.(row);
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

function applyCellValidations(td, colKey, row, profile, fechaInspeccion) {
  const c = profile.cols;
  const cliente = String(row[c.cliente] ?? "").trim();
  const calibreAR = row[c.calibreAr];
  const subGrupo = row[c.subgrupo];
  const destino = row[c.destino];
  const especial = usaCalibreEspecial(cliente, destino);
  const val = td.textContent;

  if (colKey === c.destino) applyDestinoValidation(td, cliente, destino);

  if (especial) {
    if (colKey === c.calibreAr) {
      if (!calibreAR) td.classList.add("agv-pt-cell-error-empty");
      else if (!["JUMBO", "MIXED", "REGULAR", "SUPER JUMBO"].includes(String(calibreAR).toUpperCase().trim())) {
        td.classList.add("agv-pt-cell-error-value");
        td.title = "Calibre debe ser REGULAR, MIXED, JUMBO o SUPER JUMBO";
      }
    }
    if (colKey === c.subgrupo) {
      if (esSubgrupoNA(subGrupo)) {
        td.classList.add("agv-pt-cell-error-value");
        td.title = "Subgrupo no puede ser N/A para cliente especial";
      } else if (!String(subGrupo ?? "").trim()) {
        td.classList.add("agv-pt-cell-error-empty");
        td.title = "Subgrupo obligatorio (R1, M1, J1, SJ1, etc.)";
      } else if (calibreAR) {
        const categoriaReal = String(calibreAR).toUpperCase().trim();
        const permitidos = CATEGORIAS_FRUTIST[categoriaReal];
        if (permitidos && !permitidos.includes(String(subGrupo).trim())) {
          td.classList.add("agv-pt-cell-error-value");
          td.title = `Subgrupo ${subGrupo} no pertenece a ${categoriaReal}`;
        }
      }
    }
  } else {
    if (cliente && !calibreAR && colKey === c.calibreAr) td.classList.add("agv-pt-cell-error-empty");
    if (cliente && calibreAR && colKey === c.calibreAr) {
      const trad = CALIBRES_MAP[calibreAR] || calibreAR;
      if (PALABRAS_PROHIBIDAS_REGULAR.includes(trad)) {
        td.classList.add("agv-pt-cell-error-value");
        td.title = "Calibre debe ser letra (A, B, C, K, etc.)";
      }
    }
    if (colKey === c.subgrupo) {
      if (!esSubgrupoNA(subGrupo)) {
        if (!String(subGrupo ?? "").trim()) {
          td.classList.add("agv-pt-cell-error-empty");
          td.title = "Subgrupo debe ser N/A (no puede estar vacío)";
        } else {
          td.classList.add("agv-pt-cell-error-value");
          td.title = "Subgrupo debe ser N/A";
        }
      }
    }
  }

  if (colKey === c.trazabilidad) {
    const trazCheck = extraerTrazabilidad(val);
    const julianoFecha = getJulianoFromDate(fechaInspeccion);
    if (!val) td.classList.add("agv-pt-cell-error-empty");
    else {
      const esDriscoll = cliente.toUpperCase().startsWith("DRISCOLL");
      if (trazCheck && julianoFecha) {
        const julianoActual = parseInt(julianoFecha, 10);
        const julianoTraz = parseInt(trazCheck.juliano, 10);
        const bad = esDriscoll
          ? julianoTraz !== julianoActual && julianoTraz !== julianoActual + 1
          : julianoTraz !== julianoActual;
        if (bad) td.classList.add("agv-pt-cell-error-value");
      }
      if (!VAR_MAP[trazCheck?.variedad]) td.classList.add("agv-pt-cell-error-value");
      if (val.length > 13) td.classList.add("agv-pt-cell-error-value");
    }
  }

  if (colKey === c.tPulpa) {
    const num = Number(val);
    if (!val) td.classList.add("agv-pt-cell-error-empty");
    else if (Number.isNaN(num) || num <= 0 || num > 5) td.classList.add("agv-pt-cell-error-value");
  }

  if (colKey === c.cantMuestra) {
    const n = Number(val);
    if (!val || n <= 100 || n > 530) td.classList.add("agv-pt-cell-error-value");
  }
  if (colKey === c.medMuestra && !val) {
    td.classList.add("agv-pt-cell-error-empty");
  }
  if (colKey === c.lote && !val) td.classList.add("agv-pt-cell-error-empty");
  if (colKey === c.lote && val && val.length !== 12) td.classList.add("agv-pt-cell-error-value");
  if (colKey === c.notaCondicion && !val) td.classList.add("agv-pt-cell-error-empty");
  if (colKey === c.cliente && !val) td.classList.add("agv-pt-cell-error-empty");
  if (colKey === c.cliente && val === "TF INTERNATIONAL") td.classList.add("agv-pt-cell-error-value");

  // PT: solo año-mes (Fecha Cosecha SAP 20 vs Fecha inspección).
  const cosechaJs = 19;
  const inspJs = c.fechaInspeccion;
  if (Number.isFinite(inspJs) && (colKey === cosechaJs || colKey === inspJs)) {
    paintPtYearMonthMismatch(td, colKey, row, cosechaJs, inspJs, PT_FECHA_YM_MSG);
  }
}

const COLUMN_LABELS = {
  CALIBRE: "CALIBRE",
  E_C: "E_C",
  VARIEDAD: "VARIEDAD"
};

function columnLabel(col, profile, headers) {
  const raw = COLUMN_LABELS[col] || profile.headerLabels?.[col] || headers[col] || String(col);
  return translateExcelHeader(raw, typeof col === "number" ? col : -1);
}

function applyErrorMap(td, row, colKey, errorMap) {
  if (typeof colKey !== "number" || !errorMap) return;
  const err = errorMap.get(row._filaNum)?.get(colKey + 1);
  if (!err) return;
  if (err.tipo === "obligatorio") td.classList.add("agv-pt-cell-error-empty");
  else if (err.tipo === "duplicado") td.classList.add("agv-pt-cell-error-duplicate");
  else td.classList.add("agv-pt-cell-error-value");
  td.title = err.problema;
}

export function syncTableColgroup(tableEl, dataColCount) {
  if (!tableEl) return;
  let cg = tableEl.querySelector("colgroup");
  if (!cg) {
    cg = document.createElement("colgroup");
    tableEl.insertBefore(cg, tableEl.firstChild);
  }
  cg.replaceChildren();
  const colAction = document.createElement("col");
  colAction.className = "agv-pt-col-action";
  cg.appendChild(colAction);
  for (let i = 0; i < dataColCount; i += 1) {
    const col = document.createElement("col");
    col.className = "agv-pt-col-data";
    col.dataset.colIndex = String(i);
    cg.appendChild(col);
  }
}

function initDefaultHiddenCols(tableEl) {
  if (!tableEl._ptHiddenCols) tableEl._ptHiddenCols = new Set();
}

export function applyPtColumnVisibility(tableEl) {
  if (!tableEl?._ptHiddenCols) return;
  tableEl.querySelectorAll("[data-col-index]").forEach((el) => {
    const idx = Number(el.dataset.colIndex);
    el.classList.toggle("agv-pt-col-hidden", tableEl._ptHiddenCols.has(idx));
  });
}

function normalizeLoteKey(val) {
  const raw = String(val ?? "").trim();
  if (!raw) return "";
  const n = Number(raw);
  if (!Number.isNaN(n) && Number.isFinite(n) && n > 0) return String(Math.trunc(n));
  return raw;
}

export function markDuplicateLoteRows(rows, colLoteJs) {
  const groups = new Map();
  rows.forEach((row) => {
    const lote = normalizeLoteKey(row[colLoteJs]);
    if (!lote) return;
    if (!groups.has(lote)) groups.set(lote, []);
    groups.get(lote).push(row);
  });
  groups.forEach((group) => {
    if (group.length > 1) group.forEach((row) => {
      row.__duplicado = true;
    });
  });
}

export function renderPtTable({
  refs,
  rows,
  headers,
  profile,
  fechaInspeccion,
  errorMap,
  onReorder,
  onCopyReport,
  onRowMark
}) {
  const { resultsHeader, resultsBody } = refs;
  resultsHeader.innerHTML = "";
  resultsBody.innerHTML = "";

  const thAction = document.createElement("th");
  thAction.textContent = "···";
  thAction.className = "agv-pt-action-header";
  thAction.title = i18nService.translate("arandanoPt.actionsHeader");
  resultsHeader.appendChild(thAction);

  profile.columnasFront.forEach((col, dataColIdx) => {
    const th = document.createElement("th");
    const raw = COLUMN_LABELS[col] || profile.headerLabels?.[col] || headers[col] || String(col);
    th.dataset.excelHeader = raw;
    th.textContent = columnLabel(col, profile, headers);
    th.className = "agv-pt-data-header";
    th.dataset.colIndex = String(dataColIdx);
    resultsHeader.appendChild(th);
  });

  rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    tr.dataset.rowIndex = String(rowIndex);
    const cliente = String(row[profile.cols.cliente] ?? "").trim();

    if (row.__mark === "green") tr.classList.add("agv-pt-row--mark-green");
    else if (row.__mark === "orange") tr.classList.add("agv-pt-row--mark-orange");
    if (row.__duplicado) {
      tr.classList.add("agv-pt-row--duplicate");
      tr.title = i18nService.translate("arandanoPt.duplicateLot");
    }

    tr.appendChild(buildActionsCell(row, tr, onCopyReport, onRowMark));

    const traz = extraerTrazabilidad(row[profile.cols.trazabilidad]);
    const ext = {
      CALIBRE: buildCalibreCalculado(row, profile),
      E_C: traz ? `${traz.sector.etapa}-${traz.sector.campo}` : "",
      VARIEDAD: VAR_MAP[traz?.variedad]?.[0] || ""
    };

    profile.columnasFront.forEach((col, dataColIdx) => {
      const td = document.createElement("td");
      td.dataset.colIndex = String(dataColIdx);
      if (col === "CALIBRE") {
        const { text, error } = ext.CALIBRE;
        if (error) {
          paintRedText(td, text);
          td.title = "Calibre no coincide con Subgrupo o es inválido";
        } else td.textContent = text;
      } else if (col === "E_C" || col === "VARIEDAD") {
        td.textContent = ext[col] ?? "";
      } else if (col === profile.cols.trazabilidad) {
        paintTrazabilidadCell(td, row[col], cliente);
        applyCellValidations(td, col, row, profile, fechaInspeccion);
        if (errorMap) applyErrorMap(td, row, col, errorMap);
      } else if (col === profile.cols.subgrupo) {
        td.textContent = formatSubgrupoCellValue(row, profile, cliente);
        applyCellValidations(td, col, row, profile, fechaInspeccion);
        if (errorMap) applyErrorMap(td, row, col, errorMap);
      } else {
        td.textContent = formatCellValue(row[col], col);
        applyCellValidations(td, col, row, profile, fechaInspeccion);
        if (errorMap) applyErrorMap(td, row, col, errorMap);
        if (col === profile.cols.id || col === profile.cols.lote) {
          makeCopyableCell(td, row[col]);
        }
      }
      tr.appendChild(td);
    });

    resultsBody.appendChild(tr);
  });

  bindRowDragDrop(resultsBody, onReorder);
  bindCopyableCells(resultsBody);

  const tableEl = refs.resultsTable;
  if (tableEl) {
    syncTableColgroup(tableEl, profile.columnasFront.length);
    initDefaultHiddenCols(tableEl);
    applyPtColumnVisibility(tableEl);
  }
}

export function refreshArandanoPtHeaderLabels(headerRow, headers, profile) {
  refreshTranslatedHeaderRow(headerRow, (idx) => headers[idx] || profile.headerLabels?.[idx] || "");
}

function closePtColumnMenu(menuEl) {
  if (!menuEl) return;
  menuEl.hidden = true;
  const wrap = document.getElementById("agv-pt-table-wrap");
  if (wrap && menuEl.parentElement !== wrap) wrap.appendChild(menuEl);
}

export function bindColumnContextMenu(tableEl, menuEl, options = {}) {
  if (!tableEl || !menuEl || tableEl.dataset.colMenuBound === "1") return;
  tableEl.dataset.colMenuBound = "1";
  initDefaultHiddenCols(tableEl);

  const protectedColIndices =
    options.protectedColIndices instanceof Set ? options.protectedColIndices : new Set();
  const isProtectedCol = (colIndex) => protectedColIndices.has(colIndex);

  if (!menuEl.querySelector("[data-action]")) {
    menuEl.innerHTML = `
      <button type="button" class="agv-pt-col-menu__item" data-action="hide" role="menuitem">
        <i data-lucide="eye-off" aria-hidden="true"></i>
        <span>${i18nService.translate("plagasArandano.hideColumn")}</span>
      </button>
      <button type="button" class="agv-pt-col-menu__item" data-action="show-all" role="menuitem">
        <i data-lucide="columns-3" aria-hidden="true"></i>
        <span>${i18nService.translate("plagasArandano.showAllColumnsShort")}</span>
      </button>`;
    hydrateLucideIcons(menuEl);
  }

  tableEl.addEventListener("contextmenu", (e) => {
    const th = e.target.closest("th[data-col-index]");
    if (!th || !tableEl.contains(th)) return;
    e.preventDefault();

    const colIndex = Number(th.dataset.colIndex);
    if (!Number.isFinite(colIndex)) return;

    const hideBtn = menuEl.querySelector('[data-action="hide"]');
    const showAllBtn = menuEl.querySelector('[data-action="show-all"]');
    const visibleCount = tableEl.querySelectorAll("thead th[data-col-index]:not(.agv-pt-col-hidden)").length;
    if (hideBtn) {
      hideBtn.disabled =
        isProtectedCol(colIndex) ||
        tableEl._ptHiddenCols.has(colIndex) ||
        visibleCount <= 1;
    }
    if (showAllBtn) showAllBtn.disabled = tableEl._ptHiddenCols.size === 0;

    th.appendChild(menuEl);
    menuEl.hidden = false;
    hydrateLucideIcons(menuEl);
  });

  menuEl.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-action]");
    if (!btn || btn.disabled) return;
    e.preventDefault();
    e.stopPropagation();

    const anchor = menuEl.parentElement;
    const colIndex = Number(anchor?.dataset?.colIndex);

    if (btn.dataset.action === "hide" && Number.isFinite(colIndex)) {
      if (isProtectedCol(colIndex)) return;
      const visible = tableEl.querySelectorAll("thead th[data-col-index]:not(.agv-pt-col-hidden)").length;
      if (visible <= 1) return;
      tableEl._ptHiddenCols.add(colIndex);
      applyPtColumnVisibility(tableEl);
      options.onVisibilityChange?.();
    } else if (btn.dataset.action === "show-all") {
      tableEl._ptHiddenCols.clear();
      applyPtColumnVisibility(tableEl);
      options.onVisibilityChange?.();
    }
    closePtColumnMenu(menuEl);
  });

  document.addEventListener("click", (e) => {
    if (menuEl.hidden) return;
    if (menuEl.contains(e.target)) return;
    closePtColumnMenu(menuEl);
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closePtColumnMenu(menuEl);
  });
}

function bindRowDragDrop(tbody, onReorder) {
  if (!tbody || tbody.dataset.dragBound === "1") return;
  tbody.dataset.dragBound = "1";

  let dragRow = null;
  let dragFrom = null;

  tbody.addEventListener("dragstart", (e) => {
    if (!e.target.closest(".agv-pt-drag-handle")) {
      e.preventDefault();
      return;
    }
    dragRow = e.target.closest("tr");
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
    if (to !== dragFrom) onReorder(dragFrom, to);
  });

  tbody.addEventListener("dragend", () => {
    dragRow?.classList.remove("agv-pt-row--dragging");
    dragRow = null;
    dragFrom = null;
  });
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
      tr.querySelectorAll("td[data-col-index]").forEach((td) => {
        if (td.dataset.colIndex === "0") idText = td.textContent.trim().toUpperCase();
        if (td.dataset.colIndex === "3") loteText = td.textContent.trim().toUpperCase();
      });
      tr.classList.toggle(
        "agv-pt-row--search-hidden",
        !(idText.includes(term) || loteText.includes(term))
      );
    });
  });
}

export function buildWhatsappReport(row, profile, incidencias) {
  if (!incidencias?.length) return null;
  const v = {
    usuario: row[profile.cols.usuario],
    id: row[profile.cols.id],
    lote: row[profile.cols.lote]
  };
  const lista = incidencias.map((i) => `      • ${i}`).join("\n");
  return `*Usuario:* ${v.usuario}
*ID:* ${v.id}
*Lote:* ${v.lote}
*Incidencias:*
${lista}
*Acción: Corregir inspección por favor.*`;
}
