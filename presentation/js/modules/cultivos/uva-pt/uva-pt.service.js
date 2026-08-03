import { appConfig } from "../../../config/app.config.js";
import { AGV_PT_SHELL_IDS } from "../shared/cartilla-shell.ids.js";
import { CartillaShellUi } from "../shared/cartilla-shell.ui.js";
import { hydrateLucideIcons } from "../../../utils/lucide-icon.util.js";
import { i18nService } from "../../../services/i18n.service.js";
import { showPtDialog } from "../arandano-pt/arandano-pt-dialog.js";
import {
  CARTILLA_CODE,
  MIN_FILAS,
  HEADER_ROW_INDEX,
  DATA_START_INDEX,
  loadUvaPtValidaciones,
  getTotalColumnas,
  getColInspeccionJs,
  getColCosechaJs,
  getColLmrJs,
  getExcelCabecera,
  getValidacionArchivo
} from "./uva-pt.config.js";
import { normalizeDateValue } from "./uva-pt.validation.js";
import {
  renderUvaPtTable,
  bindTableSearch,
  bindUvaPtColumnMenu,
  buildWhatsappReport,
  refreshUvaPtHeaderLabels
} from "./uva-pt-table.js";
import { exportUvaPtFiltered } from "./uva-pt-export.js";
import {
  createCartillaAnalysisController,
  deriveFilasConErrorFromDom,
  headersToAnalysisColumns
} from "../shared/cartilla-analysis.js";
import { applyDateDisplayFormatToRows } from "../shared/excel-date-format.util.js";
import { PT_SKIP_SAP_VALIDATION } from "../shared/mp-results-perf.util.js";

function t(key, vars = {}) {
  let text = i18nService.translate(key);
  Object.entries(vars).forEach(([name, value]) => {
    text = text.replace(`{{${name}}}`, String(value));
  });
  return text;
}

function ensureXlsx() {
  if (window.XLSX?.read && window.XLSX?.utils) return true;
  showPtDialog({
    icon: "error",
    title: t("plagasArandano.error"),
    text: t("plagasArandano.errorXlsxLibrary")
  });
  return false;
}

function htmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function uniqueInspectionDates(rows) {
  const col = getColInspeccionJs();
  return [...new Set(rows.map((r) => normalizeDateValue(r[col])).filter(Boolean))].sort();
}

function rowsForInspection(rows, fechaInspeccion) {
  const col = getColInspeccionJs();
  return rows.filter((r) => normalizeDateValue(r[col]) === fechaInspeccion);
}

function majorityDate(rows, colJs) {
  const dates = rows.map((r) => normalizeDateValue(r[colJs])).filter(Boolean);
  const conteo = {};
  dates.forEach((d) => {
    conteo[d] = (conteo[d] || 0) + 1;
  });
  const sorted = Object.entries(conteo).sort((a, b) => b[1] - a[1]);
  return sorted[0]?.[0] || "";
}

function setReadonlySelect(selectEl, value, placeholder, shell, warning = false) {
  if (!selectEl) return;
  selectEl.innerHTML = "";
  if (value) {
    selectEl.add(new Option(value, value, true, true));
  } else {
    selectEl.add(new Option(placeholder, "", true, true));
  }
  selectEl.disabled = true;
  selectEl.classList.add(`${shell.cls("input")}--readonly`);
  selectEl.classList.toggle(`${shell.cls("input")}--warning`, warning);
}

export class UvaPtService {
  constructor() {
    this.shell = null;
    this.headers = [];
    this.dataRows = [];
    this.excelCabecera = null;
    this.filteredTableRows = [];
    this.colMenuBound = false;
    this.searchBound = false;
    this.abortController = null;
  }

  async init(appRoot) {
    await loadUvaPtValidaciones(appConfig.cacheBustingVersion);

    this.shell = new CartillaShellUi({
      root: appRoot,
      ids: AGV_PT_SHELL_IDS,
      cssPrefix: "agv-pt",
      i18nPrefix: "ptUva"
    });
    this.shell.cacheDom();
    this.cartillaAnalysis = createCartillaAnalysisController({
      getRoot: () => this.shell?.root,
      hostSelector: "#agv-pt-cartilla-analysis",
      showDialog: (opts) => showPtDialog(opts),
      t,
      htmlEscape
    });
    this.bindEvents();
    this.shell.resetDashboard();
    this.shell.renderExcelInsightEmpty();
    this.shell.setLiveStatus(false);
    hydrateLucideIcons(appRoot);
  }

  bindEvents() {
    this.abortController?.abort();
    this.abortController = new AbortController();
    const { signal } = this.abortController;
    const refs = this.shell.refs;

    refs.fileInput?.addEventListener("change", (e) => this.onFileSelected(e), { signal });
    refs.inspectionSelect?.addEventListener("change", () => this.onInspectionDateChange(), { signal });
    refs.runReviewBtn?.addEventListener("click", () => this.onRunReview(), { signal });
    refs.exportBtn?.addEventListener("click", () => this.onExportFiltered(), { signal });
    refs.clearBtn?.addEventListener("click", () => this.onClear(), { signal });
  }

  syncButtons() {
    const hasFile = this.dataRows.length > 0;
    const hasInspection = Boolean(this.shell.refs.inspectionSelect?.value);
    const disabled = !hasFile || !hasInspection;
    if (this.shell.refs.runReviewBtn) this.shell.refs.runReviewBtn.disabled = disabled;
    if (this.shell.refs.exportBtn) this.shell.refs.exportBtn.disabled = disabled;
  }

  async onFileSelected(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!ensureXlsx()) return;

    try {
      const buffer = await file.arrayBuffer();
      const wb = window.XLSX.read(buffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

      if (data.length < MIN_FILAS) {
        showPtDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("ptUva.invalidMinRows", { min: MIN_FILAS, found: data.length })
        });
        event.target.value = "";
        return;
      }

      const valArch = getValidacionArchivo();
      const grupoCell = data[valArch.fila_grupo_js ?? 3]?.[valArch.col_grupo_js ?? 8];
      const grupoEsperado = (valArch.grupo_esperado || `G${CARTILLA_CODE}`).toUpperCase();
      if (String(grupoCell ?? "").trim().toUpperCase() !== grupoEsperado) {
        showPtDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("ptUva.invalidCartilla")
        });
        event.target.value = "";
        return;
      }

      const estadoCell = data[valArch.fila_estado_js ?? 3]?.[valArch.col_estado_js ?? 13];
      if (valArch.estado_esperado && String(estadoCell ?? "").trim().toUpperCase() !== valArch.estado_esperado) {
        showPtDialog({
          icon: "warning",
          title: t("ptUva.warnTitle"),
          text: t("ptUva.warnEstado", { estado: valArch.estado_esperado })
        });
      }

      this.headers = data[HEADER_ROW_INDEX] || [];
      if (this.headers.length < getTotalColumnas()) {
        showPtDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("ptUva.invalidColumns", {
            expected: getTotalColumnas(),
            found: this.headers.length
          })
        });
        event.target.value = "";
        return;
      }

      this.excelCabecera = this.parseExcelCabecera(data);
      // Excel 1-based: cosecha 52, embalaje 53, inspección 54, LMR 71 (nunca 51 genérico).
      this.dataRows = applyDateDisplayFormatToRows(
        data
          .slice(DATA_START_INDEX)
          .filter((row) => row.some((c) => String(c ?? "").trim())),
        this.headers,
        [20, 21, 52, 53, 54, 71]
      );

      const tipoFila = String(this.dataRows[0]?.[1] ?? "").trim().toUpperCase();
      if (!tipoFila || tipoFila !== CARTILLA_CODE) {
        showPtDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("ptUva.invalidCartilla")
        });
        event.target.value = "";
        return;
      }

      this.resetResultsUi();

      const fechas = uniqueInspectionDates(this.dataRows);
      const sel = this.shell.refs.inspectionSelect;
      if (sel) {
        sel.innerHTML = `<option value="" disabled selected>${t("ptUva.selectDate")}</option>`;
        fechas.forEach((f) => sel.add(new Option(f, f)));
        sel.disabled = fechas.length === 0;
      }

      this.resetAutoDateFields();

      this.shell.setLiveStatus(true);
      this.syncButtons();
      this.renderExcelInsight();

      showPtDialog({
        icon: "success",
        title: "Excel cargado",
        html: `Cartilla <b>${htmlEscape(CARTILLA_CODE)}</b> · <b>${this.dataRows.length}</b> registros · <b>${getTotalColumnas()}</b> columnas`,
        timer: 1800,
        showConfirmButton: false
      });
    } catch (err) {
      console.error(err);
      showPtDialog({
        icon: "error",
        title: t("plagasArandano.error"),
        text: t("plagasArandano.errorArchivoVacio")
      });
    }
  }

  readSheetCell(sheet, fila, columna) {
    const value = sheet[(fila ?? 1) - 1]?.[(columna ?? 1) - 1];
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  parseExcelCabecera(sheet) {
    const cab = getExcelCabecera();
    if (!cab) return null;
    const meta = { titulo: this.readSheetCell(sheet, cab.titulo.fila, cab.titulo.columna) };
    cab.campos.forEach((field) => {
      meta[field.clave] = this.readSheetCell(sheet, field.fila, field.columna);
    });
    return meta;
  }

  renderExcelInsight() {
    const { excelInsightEl } = this.shell.refs;
    if (!excelInsightEl) return;
    const p = (part) => this.shell.cls(part);
    const meta = this.excelCabecera;
    if (!meta) {
      this.shell.renderExcelInsightEmpty();
      return;
    }
    const ringRadius = 42;
    const ringCircumference = 2 * Math.PI * ringRadius;
    const grupo = meta.grupo || CARTILLA_CODE;
    const estado = meta.estado || "—";
    const reportTitle = meta.titulo || "";
    excelInsightEl.className = `${p("excel-insight")} ${p("excel-insight")}--loaded`;
    excelInsightEl.innerHTML = `
      ${reportTitle ? `<p class="${p("excel-insight__report")}">${htmlEscape(reportTitle)}</p>` : ""}
      <div class="${p("excel-insight__body")}">
        <div class="${p("excel-insight__top")}">
          <div class="${p("excel-insight__primary")}">
            <div class="${p("excel-insight__stat")} ${p("excel-insight__stat")}--primary">
              <span class="${p("excel-insight__stat-label")}">${htmlEscape(t("plagasArandano.metaCompany"))}</span>
              <strong>${htmlEscape(meta.empresa || "—")}</strong>
            </div>
            <div class="${p("excel-insight__stat")} ${p("excel-insight__stat")}--primary">
              <span class="${p("excel-insight__stat-label")}">${htmlEscape(t("plagasArandano.metaClient"))}</span>
              <strong>${htmlEscape(meta.mandante || "—")}</strong>
            </div>
          </div>
          <div class="${p("excel-insight__ring")}" aria-hidden="true">
            <svg viewBox="0 0 100 100" shape-rendering="geometricPrecision">
              <defs>
                <linearGradient id="agvPtInsightRingGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                  <stop offset="0%" stop-color="#5eb8d9"></stop>
                  <stop offset="100%" stop-color="#22c55e"></stop>
                </linearGradient>
              </defs>
              <circle class="${p("excel-insight__ring-glow")}" cx="50" cy="50" r="${ringRadius}"></circle>
              <circle class="${p("excel-insight__ring-track")}" cx="50" cy="50" r="${ringRadius}"></circle>
              <circle
                class="${p("excel-insight__ring-value")}"
                cx="50"
                cy="50"
                r="${ringRadius}"
                stroke-dasharray="${ringCircumference}"
                stroke-dashoffset="0"
              ></circle>
            </svg>
            <div class="${p("excel-insight__ring-label")}">
              <strong>${htmlEscape(grupo)}</strong>
              <span>${htmlEscape(t("plagasArandano.metaGroup"))}</span>
            </div>
          </div>
        </div>
        <div class="${p("excel-insight__stats")}">
          <div class="${p("excel-insight__stat")}">
            <span class="${p("excel-insight__stat-label")}">${htmlEscape(t("plagasArandano.metaCrop"))}</span>
            <strong>${htmlEscape(meta.cultivo || "—")}</strong>
          </div>
          <div class="${p("excel-insight__stat")} ${p("excel-insight__stat")}--status">
            <span class="${p("excel-insight__stat-label")}">${htmlEscape(t("plagasArandano.metaStatus"))}</span>
            <strong>${htmlEscape(estado)}</strong>
          </div>
        </div>
      </div>`;
    hydrateLucideIcons(excelInsightEl);
  }

  resetAutoDateFields() {
    const { cosechaSelect, lmrSelect } = this.shell.refs;
    setReadonlySelect(cosechaSelect, "", t("ptUva.autoDate"), this.shell);
    if (lmrSelect) {
      lmrSelect.innerHTML = `<option value="" selected>${t("ptUva.lmrAutoDate")}</option>`;
      lmrSelect.disabled = true;
      lmrSelect.classList.add(`${this.shell.cls("input")}--readonly`);
      lmrSelect.classList.remove(`${this.shell.cls("input")}--warning`);
    }
  }

  onInspectionDateChange() {
    const fechaIns = this.shell.refs.inspectionSelect?.value;
    this.resetResultsUi();

    if (!fechaIns) {
      this.resetAutoDateFields();
      this.syncButtons();
      return;
    }

    const rows = rowsForInspection(this.dataRows, fechaIns);
    const cosecha = majorityDate(rows, getColCosechaJs());
    const lmrRows = rows.map((r) => normalizeDateValue(r[getColLmrJs()])).filter(Boolean);
    const uniqueLmr = [...new Set(lmrRows)];
    const lmrMayoritaria = majorityDate(rows, getColLmrJs());

    setReadonlySelect(this.shell.refs.cosechaSelect, cosecha, t("ptUva.autoDate"), this.shell);

    const lmrSel = this.shell.refs.lmrSelect;
    if (lmrSel) {
      if (lmrMayoritaria) {
        lmrSel.innerHTML = `<option value="${htmlEscape(lmrMayoritaria)}" selected>${htmlEscape(lmrMayoritaria)}</option>`;
      } else {
        lmrSel.innerHTML = `<option value="" selected>${t("ptUva.lmrAutoDate")}</option>`;
      }
      lmrSel.disabled = true;
      lmrSel.classList.add(`${this.shell.cls("input")}--readonly`);

      if (uniqueLmr.length > 1) {
        lmrSel.classList.add(`${this.shell.cls("input")}--warning`);
        const conteo = {};
        lmrRows.forEach((f) => {
          conteo[f] = (conteo[f] || 0) + 1;
        });
        const detalles = Object.entries(conteo)
          .map(
            ([fecha, count]) =>
              `${htmlEscape(fecha)}: <b>${count} registros</b>${fecha === lmrMayoritaria ? " (MAYORITARIA)" : ""}`
          )
          .join("<br>");

        showPtDialog({
          icon: "warning",
          title: t("ptUva.multipleLmrTitle"),
          html: `<div class="agv-pt-dialog__html--body">
            Se detectaron <b>${uniqueLmr.length}</b> fechas LMR diferentes:<br><br>
            ${detalles}<br><br>
            Se usará la fecha mayoritaria.
          </div>`,
          wide: true
        });
      } else {
        lmrSel.classList.remove(`${this.shell.cls("input")}--warning`);
      }
    }

    this.syncButtons();
  }

  getRowsForSelectedInspection() {
    const fechaIns = this.shell.refs.inspectionSelect?.value;
    if (!fechaIns) return [];
    return rowsForInspection(this.dataRows, fechaIns);
  }

  async onRunReview() {
    const fechaIns = this.shell.refs.inspectionSelect?.value;
    if (!fechaIns) {
      showPtDialog({
        icon: "warning",
        title: t("ptUva.dateRequiredTitle"),
        text: t("ptUva.dateRequiredText")
      });
      return;
    }

    const filtradas = this.getRowsForSelectedInspection();
    this.filteredTableRows = filtradas;

    const refs = this.shell.refs;
    if (refs.resultsSection) {
      refs.resultsSection.hidden = false;
      refs.resultsSection.classList.add("is-visible");
    }
    if (refs.tableSearchWrap) refs.tableSearchWrap.hidden = false;
    if (refs.resultsTable) refs.resultsTable.hidden = false;

    const onCopyReport = async (row) => {
      const msg = buildWhatsappReport(row);
      if (!msg) {
        await showPtDialog({
          icon: "info",
          title: t("ptUva.rowOkTitle"),
          text: t("ptUva.rowOkText"),
          timer: 1000,
          showConfirmButton: false
        });
        return;
      }
      try {
        await navigator.clipboard.writeText(msg);
        await showPtDialog({
          icon: "success",
          title: t("ptUva.copySuccessTitle"),
          text: t("ptUva.copySuccessText"),
          timer: 1200,
          showConfirmButton: false
        });
      } catch {
        await showPtDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("ptUva.copyErrorText")
        });
      }
    };

    renderUvaPtTable({
      headerRow: refs.resultsHeader,
      bodyRows: refs.resultsBody,
      headers: this.headers,
      tableEl: refs.resultsTable,
      filteredRows: filtradas,
      onFilterChange: (count, rows) => {
        if (rows) this.filteredTableRows = rows;
        if (refs.totalFilasDiv) {
          refs.totalFilasDiv.textContent = t("ptUva.totalRowsShown", { count });
        }
      },
      onReorder: (from, to) => {
        const moved = this.filteredTableRows.splice(from, 1)[0];
        this.filteredTableRows.splice(to, 0, moved);
        const tbody = refs.resultsBody;
        if (!tbody) return;
        const tr = tbody.children[from];
        if (!tr) return;
        const refNode = tbody.children[to > from ? to + 1 : to] ?? null;
        if (refNode) tbody.insertBefore(tr, refNode);
        else tbody.appendChild(tr);
      },
      onCopyReport
    });

    if (refs.totalFilasDiv) {
      refs.totalFilasDiv.textContent = t("ptUva.totalRowsShown", { count: filtradas.length });
    }

    if (!this.colMenuBound && refs.resultsTable && refs.colMenuEl) {
      bindUvaPtColumnMenu(refs.resultsTable, refs.colMenuEl);
      this.colMenuBound = true;
    }

    if (!this.searchBound && refs.tableSearch && refs.resultsBody) {
      bindTableSearch(refs.tableSearch, refs.resultsBody);
      this.searchBound = true;
    }

    this.syncButtons();
    hydrateLucideIcons(this.shell.root);
    refs.resultsSection?.scrollIntoView({ behavior: "smooth", block: "nearest" });

    const filasConError = deriveFilasConErrorFromDom(refs.resultsBody, filtradas);
    this.cartillaAnalysis?.present({
      rows: filtradas,
      filasConError,
      errorMap: null,
      duplicateLotes: new Set(),
      colLoteJs: 9,
      columns: headersToAnalysisColumns(this.headers),
      cartilla: CARTILLA_CODE,
      fechaLabel: fechaIns || "—",
      skipSapValidation: PT_SKIP_SAP_VALIDATION
    });
  }

  onExportFiltered() {
    const fechaIns = this.shell.refs.inspectionSelect?.value;
    const rows = this.filteredTableRows.length
      ? this.filteredTableRows
      : this.getRowsForSelectedInspection();
    if (!rows.length) {
      showPtDialog({
        icon: "warning",
        title: t("ptUva.dateRequiredTitle"),
        text: t("ptUva.noExportData")
      });
      return;
    }
    if (!exportUvaPtFiltered({ rows, headers: this.headers, fechaLabel: fechaIns })) {
      showPtDialog({
        icon: "error",
        title: t("plagasArandano.error"),
        text: t("plagasArandano.errorXlsxLibrary")
      });
      return;
    }
    showPtDialog({
      icon: "success",
      title: t("ptUva.exportSuccessTitle"),
      text: t("ptUva.exportSuccessText"),
      timer: 2000,
      showConfirmButton: false
    });
  }

  resetResultsUi() {
    const refs = this.shell.refs;
    if (refs.resultsHeader) refs.resultsHeader.innerHTML = "";
    if (refs.resultsBody) refs.resultsBody.innerHTML = "";
    if (refs.resultsTable) refs.resultsTable.hidden = true;
    if (refs.resultsSection) {
      refs.resultsSection.hidden = true;
      refs.resultsSection.classList.remove("is-visible");
    }
    if (refs.tableSearchWrap) refs.tableSearchWrap.hidden = true;
    if (refs.tableSearch) refs.tableSearch.value = "";
    if (refs.totalFilasDiv) refs.totalFilasDiv.textContent = "";
    this.filteredTableRows = [];
    this.cartillaAnalysis?.clear();
  }

  onClear() {
    this.dataRows = [];
    this.headers = [];
    this.excelCabecera = null;
    this.resetResultsUi();
    if (this.shell.refs.fileInput) this.shell.refs.fileInput.value = "";
    if (this.shell.refs.inspectionSelect) {
      this.shell.refs.inspectionSelect.innerHTML = `<option value="" disabled selected>${t("ptUva.selectDate")}</option>`;
      this.shell.refs.inspectionSelect.disabled = true;
    }
    this.resetAutoDateFields();
    this.shell.setLiveStatus(false);
    this.shell.renderExcelInsightEmpty();
    this.syncButtons();
    showPtDialog({
      icon: "success",
      title: t("ptUva.clearSuccessTitle"),
      text: t("ptUva.clearSuccessText"),
      timer: 1000,
      showConfirmButton: false
    });
  }

  onLanguageChange() {
    refreshUvaPtHeaderLabels(this.shell?.refs?.resultsHeader, this.headers);
  }

  destroy() {
    this.abortController?.abort();
    this.abortController = null;
    this.shell = null;
  }
}
