/** Tabla Uva MP — columnas visibles + sumas calculadas, filas con error */

import { getColumnsToShow, getStickyCols } from "./uva-mp.config.js";
import {
  filaTieneError,
  getCellMeta,
  formatFechaCelda,
  valorCelda,
  resolveExtraColumnKey,
  EXTRA_COL_SUMA_TONALIDADES,
  EXTRA_COL_SUMA_CALIBRES
} from "./uva-mp.validation.js";
import { hydrateLucideIcons } from "../../../utils/lucide-icon.util.js";
import { translateExcelHeader } from "../../../utils/excel-header-i18n.util.js";
import { refreshTranslatedHeaderRow } from "../../../utils/table-header-i18n.util.js";
import { isDateColumnJs, formatDateValueToDMY } from "../shared/excel-date-format.util.js";

/** JS: cosecha ~19, inspección 50, LMR 69 — siempre filtrado por encabezado. */
const DATE_COLS = new Set([19, 50, 69]);

function isPinnedColumn(index) {
  return getStickyCols().includes(index);
}

export function applyStickyColumnClasses(el, index) {
  if (!isPinnedColumn(index)) return;
  el.classList.add("agv-mp-sticky-col", `agv-mp-sticky-col-${index}`);
}

function formatUvaMpCell(raw, colJs, headers) {
  const hinted = [...DATE_COLS].map((js) => js + 1);
  if (DATE_COLS.has(colJs) && isDateColumnJs(colJs, headers, hinted)) {
    const asDate = formatDateValueToDMY(raw);
    if (asDate) return asDate;
    return formatFechaCelda(raw);
  }
  return valorCelda(raw);
}

function buildDisplayColumns(headers) {
  const { indices, extra } = getColumnsToShow();
  return [
    ...indices.map((colJs) => ({
      key: colJs,
      rawLabel: headers[colJs] || `Col ${colJs + 1}`,
      label: translateExcelHeader(headers[colJs] || `Col ${colJs + 1}`, colJs),
      sticky: colJs
    })),
    ...extra.map((label) => ({
      key: resolveExtraColumnKey(label),
      rawLabel: label,
      label: translateExcelHeader(label),
      sticky: null
    }))
  ];
}

function formatCellDisplay(row, colJs, headers = []) {
  return formatUvaMpCell(row[colJs], colJs, headers);
}

function formatExtraDisplay(row, key) {
  if (key === EXTRA_COL_SUMA_TONALIDADES) {
    return row._suma_tonalidades != null && Number.isFinite(row._suma_tonalidades)
      ? row._suma_tonalidades.toFixed(2)
      : "";
  }
  if (key === EXTRA_COL_SUMA_CALIBRES) {
    return row._suma_calibres != null && Number.isFinite(row._suma_calibres)
      ? row._suma_calibres.toFixed(2)
      : "";
  }
  return "";
}

export function renderUvaMpResultsTable({
  refs,
  headers,
  allRows,
  filasConError,
  fechaISO,
  formatISOToDMY,
  t
}) {
  const { resultsHeader, resultsBody, resultsTable, resultsSection, resultsTitleEl, resultsSubtitleEl, resultsIconEl, totalFilasDiv } =
    refs;

  if (resultsHeader) resultsHeader.innerHTML = "";
  if (resultsBody) resultsBody.innerHTML = "";

  const hasErrors = filasConError.length > 0;
  const shellCls = (part) => `agv-mp-${part}`;
  const displayCols = buildDisplayColumns(headers);

  if (resultsSection) {
    resultsSection.classList.remove(`${shellCls("results")}--ok`, `${shellCls("results")}--errors`);
    resultsSection.classList.add(hasErrors ? `${shellCls("results")}--errors` : `${shellCls("results")}--ok`);
    resultsSection.classList.add("is-visible");
  }

  if (resultsTitleEl) {
    resultsTitleEl.textContent = hasErrors
      ? t("uvaMp.errorRowsTitle")
      : t("plagasArandano.allCorrect");
  }

  if (resultsSubtitleEl) {
    resultsSubtitleEl.textContent = t("plagasArandano.resultsInspectionDate", {
      date: formatISOToDMY(fechaISO)
    });
  }

  if (resultsIconEl) {
    resultsIconEl.innerHTML = hasErrors
      ? '<i data-lucide="triangle-alert"></i>'
      : '<i data-lucide="circle-check"></i>';
  }

  if (totalFilasDiv) {
    totalFilasDiv.textContent = t("uvaMp.totalInspectionRows", { count: allRows.length });
  }

  if (resultsTable) {
    resultsTable.classList.add("agv-mp-table--uva");
    resultsTable.hidden = false;
  }

  if (!hasErrors) {
    const tr = document.createElement("tr");
    tr.className = "agv-mp-row-ok";
    const td = document.createElement("td");
    td.colSpan = Math.max(displayCols.length, 1);
    td.textContent = t("uvaMp.noInspectionErrors");
    tr.appendChild(td);
    resultsBody?.appendChild(tr);
    return;
  }

  displayCols.forEach((col) => {
    const th = document.createElement("th");
    th.className = "agv-mp-table__col-header";
    th.dataset.colIndex = String(col.key);
    th.dataset.excelHeader = col.rawLabel;
    th.textContent = col.label;
    if (col.sticky != null) applyStickyColumnClasses(th, col.sticky);
    resultsHeader?.appendChild(th);
  });

  filasConError.forEach((row) => {
    const tr = document.createElement("tr");
    displayCols.forEach((col) => {
      const { val, cellClass, title } = getCellMeta(row, col.key);
      const td = document.createElement("td");
      td.dataset.colIndex = String(col.key);
      if (cellClass) td.className = cellClass;
      if (title) td.title = title;
      td.textContent =
        val ??
        (typeof col.key === "number"
          ? formatCellDisplay(row, col.key, headers)
          : formatExtraDisplay(row, col.key));
      if (col.sticky != null) applyStickyColumnClasses(td, col.sticky);
      tr.appendChild(td);
    });
    resultsBody?.appendChild(tr);
  });
}

export function refreshUvaMpHeaderLabels(headerRow, headers) {
  refreshTranslatedHeaderRow(headerRow, (idx) => {
    if (typeof idx === "number" && Number.isFinite(idx)) return headers[idx] || "";
    return "";
  });
}

export function htmlTablaFilasConError(headers, filas, { htmlEscape, t, titled = true }) {
  if (!filas?.length) return "";

  const displayCols = buildDisplayColumns(headers);
  const thead = displayCols
    .map((col) => {
      const sticky =
        col.sticky != null ? ` agv-mp-sticky-col agv-mp-sticky-col-${col.sticky}` : "";
      return `<th class="agv-mp-table__col-header${sticky}">${htmlEscape(col.label)}</th>`;
    })
    .join("");

  const tbody = filas
    .map((row) => {
      const tds = displayCols
        .map((col) => {
          const { val, cellClass, title } = getCellMeta(row, col.key);
          const sticky =
            col.sticky != null ? `agv-mp-sticky-col agv-mp-sticky-col-${col.sticky}` : "";
          const classes = [cellClass, sticky].filter(Boolean).join(" ");
          const classAttr = classes ? ` class="${htmlEscape(classes)}"` : "";
          const titleAttr = title ? ` title="${htmlEscape(title)}"` : "";
          const displayVal =
            val ??
            (typeof col.key === "number"
          ? formatCellDisplay(row, col.key, headers)
          : formatExtraDisplay(row, col.key));
          return `<td${classAttr}${titleAttr}>${htmlEscape(displayVal)}</td>`;
        })
        .join("");
      return `<tr>${tds}</tr>`;
    })
    .join("");

  const titleBlock = titled
    ? `<p class="agv-mp-nested-table-title">${htmlEscape(t("uvaMp.errorRowsTitle"))}</p>`
    : "";

  return `
    <div class="agv-mp-nested-table-wrap">
      ${titleBlock}
      <div class="agv-mp-table-scroll">
        <table class="agv-mp-table agv-mp-table--uva">
          <thead><tr>${thead}</tr></thead>
          <tbody>${tbody}</tbody>
        </table>
      </div>
    </div>`;
}

export { filaTieneError, hydrateLucideIcons };
