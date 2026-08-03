import { appConfig } from "../../../config/app.config.js";
import { AGV_PT_SHELL_IDS } from "../shared/cartilla-shell.ids.js";
import { CartillaShellUi } from "../shared/cartilla-shell.ui.js";
import { hydrateLucideIcons } from "../../../utils/lucide-icon.util.js";
import { i18nService } from "../../../services/i18n.service.js";
import { showPtDialog } from "../arandano-pt/arandano-pt-dialog.js";
import {
  CARTILLA_CODE,
  CARTILLA_HEADER_ROW,
  CARTILLA_HEADER_COL,
  HEADER_ROW_INDEX,
  DATA_START_INDEX,
  loadEsparragoPtValidaciones,
  getTotalColumnas,
  getColInspeccionJs,
  getColLmrJs,
  getExcelCabecera,
  getColumnLabelsByIndex,
  getExportDateColsJs
} from "./esparrago-pt.config.js";
import { scanGlobalWarnings, serialExcelAFecha, computeFechaLmrMayoritaria, formatISOToDMY, parseFechaToISO } from "./esparrago-pt.validation.js";
import {
  renderEsparragoPtTable,
  bindTableSearch,
  bindEsparragoPtColumnMenu,
  refreshEsparragoPtHeaderLabels
} from "./esparrago-pt-table.js";
import { exportEsparragoPtFiltered } from "./esparrago-pt-export.js";
import { expandMissingSapLayout } from "../shared/mp-sap-layout.util.js?v=2026072219";
import { applyDateDisplayFormatToRows } from "../shared/excel-date-format.util.js";
import { loadSapColumnasCatalog, getSapPerfil } from "../../../config/sap-columnas.registry.js";
import { PT_SKIP_SAP_VALIDATION } from "../shared/mp-results-perf.util.js";
import {
  createCartillaAnalysisController,
  deriveFilasConErrorFromDom,
  headersToAnalysisColumns
} from "../shared/cartilla-analysis.js";
import {
  setEsparragoPtAiSnapshot,
  clearEsparragoPtAiSnapshot
} from "./esparrago-pt-ai-store.js";
import { buildEsparragoPtAiDataBrief } from "./esparrago-pt-ai-data.js";
import { resolvePtEsparragoColumnLabel } from "./esparrago-pt-i18n-labels.js";

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
  return [
    ...new Set(rows.map((r) => normalizeInspectionDateValue(r[col])).filter(Boolean))
  ].sort();
}

function normalizeInspectionDateValue(val) {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number" && Number.isFinite(val)) return serialExcelAFecha(val);
  return String(val).trim();
}

function rowMatchesInspectionDate(row, fechaSel) {
  return normalizeInspectionDateValue(row[getColInspeccionJs()]) === fechaSel;
}

export class EsparragoPtService {
  constructor() {
    this.shell = null;
    this.headers = [];
    this.dataRows = [];
    this.excelCabecera = null;
    this.filteredTableRows = [];
    this.colMenuBound = false;
    this.searchBound = false;
    this.abortController = null;
    this.aiAssistant = null;
    this.lastAiSnapshot = null;
  }

  async init(appRoot) {
    await Promise.all([
      loadEsparragoPtValidaciones(appConfig.cacheBustingVersion),
      loadSapColumnasCatalog(appConfig.cacheBustingVersion)
    ]);

    this.shell = new CartillaShellUi({
      root: appRoot,
      ids: AGV_PT_SHELL_IDS,
      cssPrefix: "agv-pt",
      i18nPrefix: "ptEsparrago"
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
    const hasDate = Boolean(this.shell.refs.inspectionSelect?.value);
    if (this.shell.refs.runReviewBtn) {
      this.shell.refs.runReviewBtn.disabled = !hasFile || !hasDate;
    }
    if (this.shell.refs.exportBtn) {
      this.shell.refs.exportBtn.disabled = !hasFile || !hasDate;
    }
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

      const cartillaCell = data[CARTILLA_HEADER_ROW - 1]?.[CARTILLA_HEADER_COL - 1];
      if (String(cartillaCell ?? "").trim().toUpperCase() !== CARTILLA_CODE) {
        showPtDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("ptEsparrago.invalidCartilla")
        });
        event.target.value = "";
        return;
      }

      const rawHeaders = data[HEADER_ROW_INDEX] || [];
      const rawDataRows = data
        .slice(DATA_START_INDEX)
        .filter((row) => row.some((c) => String(c ?? "").trim()));

      // Mismo criterio MP: Nota Condición → col 28; 15+5 huecos SAP si faltan.
      const {
        headers,
        rows: layoutRows,
        expanded: sapLayoutExpanded,
        insertedSap15,
        insertedSap5
      } = expandMissingSapLayout(rawHeaders, rawDataRows, getSapPerfil("pt"));

      if (headers.length < getTotalColumnas()) {
        showPtDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("ptEsparrago.invalidColumns", {
            expected: getTotalColumnas(),
            found: headers.length
          })
        });
        event.target.value = "";
        return;
      }

      this.headers = headers;
      this.excelCabecera = this.parseExcelCabecera(data);
      // Fechas reales (JS→Excel 1-based). Nunca usar 51 aquí: en PTES la col 51 es Linea Asp.
      const dateExcelHints = [...getExportDateColsJs()].map((js) => js + 1);
      this.dataRows = applyDateDisplayFormatToRows(layoutRows, headers, dateExcelHints);
      this._sapLayoutNotice = sapLayoutExpanded
        ? { insertedSap15, insertedSap5 }
        : null;
      this.resetResultsUi();

      const fechas = uniqueInspectionDates(this.dataRows);
      const sel = this.shell.refs.inspectionSelect;
      if (sel) {
        sel.innerHTML = `<option value="" disabled selected>${t("ptEsparrago.selectDate")}</option>`;
        fechas.forEach((f) => sel.add(new Option(f, f)));
        sel.disabled = fechas.length === 0;
      }

      if (this.shell.refs.lmrSelect) {
        this.shell.refs.lmrSelect.innerHTML = `<option value="" selected>${t("ptEsparrago.lmrAutoDate")}</option>`;
      }

      this.shell.setLiveStatus(true);
      this.syncButtons();
      this.renderExcelInsight();

      const sapNote =
        !PT_SKIP_SAP_VALIDATION && this._sapLayoutNotice
          ? `<br><small>Se completaron huecos SAP vacíos (+${this._sapLayoutNotice.insertedSap15} + ${this._sapLayoutNotice.insertedSap5}) para alinear Nota Condición (col 28) y el bloque siguiente (col 34).</small>`
          : "";
      showPtDialog({
        icon: "success",
        title: "Excel cargado",
        html: `Cartilla <b>${htmlEscape(CARTILLA_CODE)}</b> · <b>${this.dataRows.length}</b> registros · <b>${getTotalColumnas()}</b> columnas${sapNote}`,
        timer: sapNote ? 3200 : 1800,
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

  onInspectionDateChange() {
    const fecha = this.shell.refs.inspectionSelect?.value;
    const lmrSel = this.shell.refs.lmrSelect;
    if (!fecha || !lmrSel) {
      this.syncButtons();
      return;
    }
    const rows = this.dataRows.filter((r) => rowMatchesInspectionDate(r, fecha));
    const colLmr = getColLmrJs();
    const mayoritariaISO = computeFechaLmrMayoritaria(rows, colLmr);
    const lmrDates = [...new Set(rows.map((r) => parseFechaToISO(r[colLmr])).filter(Boolean))];
    const lmrDisplay = formatISOToDMY(mayoritariaISO) || "PENDIENTE";

    lmrSel.innerHTML = `<option value="${htmlEscape(mayoritariaISO || lmrDisplay)}" selected>${htmlEscape(lmrDisplay)}</option>`;
    lmrSel.disabled = true;

    if (lmrDates.length > 1) {
      lmrSel.classList.add(`${this.shell.cls("input")}--warning`);
      const conteo = {};
      rows.forEach((r) => {
        const iso = parseFechaToISO(r[colLmr]);
        if (!iso) return;
        conteo[iso] = (conteo[iso] || 0) + 1;
      });
      const detalles = Object.entries(conteo)
        .sort((a, b) => b[1] - a[1])
        .map(([iso, n]) => `• ${htmlEscape(formatISOToDMY(iso))}: <b>${n}</b> fila(s)`)
        .join("<br>");
      showPtDialog({
        icon: "warning",
        title: "Múltiples fechas LMR detectadas",
        html: `<div class="agv-pt-dialog__html--body">
          Se detectaron <b>${lmrDates.length}</b> fechas LMR diferentes:<br><br>
          ${detalles}<br><br>
          Se usará la fecha mayoritaria <b>${htmlEscape(lmrDisplay)}</b>. Las filas con fechas minoritarias se marcarán como error.
        </div>`,
        wide: true
      });
    } else {
      lmrSel.classList.remove(`${this.shell.cls("input")}--warning`);
    }

    this.resetResultsUi();
    this.syncButtons();
  }

  getRowsForSelectedDate() {
    const fecha = this.shell.refs.inspectionSelect?.value;
    if (!fecha) return [];
    return this.dataRows.filter((r) => rowMatchesInspectionDate(r, fecha));
  }

  async onRunReview() {
    const fecha = this.shell.refs.inspectionSelect?.value;
    if (!fecha) {
      showPtDialog({
        icon: "warning",
        title: t("ptEsparrago.dateRequiredTitle"),
        text: t("ptEsparrago.dateRequiredText")
      });
      return;
    }

    const filtradas = this.getRowsForSelectedDate();
    this.filteredTableRows = filtradas;

    const { errorSAP, errorDefectos, errorCalidad } = scanGlobalWarnings(filtradas);
    if ((!PT_SKIP_SAP_VALIDATION && errorSAP) || errorDefectos || errorCalidad) {
      const mensajes = [];
      if (!PT_SKIP_SAP_VALIDATION && errorSAP) mensajes.push(t("ptEsparrago.warnSap"));
      if (errorDefectos) mensajes.push(t("ptEsparrago.warnDefectos"));
      if (errorCalidad) mensajes.push(t("ptEsparrago.warnCalidad"));
      await showPtDialog({
        icon: "warning",
        title: t("ptEsparrago.warnTitle"),
        html: `<div class="agv-pt-dialog__html--warnings">${mensajes.join("<br>")}</div>`,
        confirmButtonText: t("ptEsparrago.warnConfirm")
      });
    }

    const refs = this.shell.refs;
    if (refs.resultsSection) {
      refs.resultsSection.hidden = false;
      refs.resultsSection.classList.add("is-visible");
    }
    if (refs.tableSearchWrap) refs.tableSearchWrap.hidden = false;
    if (refs.resultsTable) refs.resultsTable.hidden = false;

    renderEsparragoPtTable({
      headerRow: refs.resultsHeader,
      bodyRows: refs.resultsBody,
      headers: this.headers,
      tableEl: refs.resultsTable,
      allRowsForDate: filtradas,
      filteredRows: filtradas,
      onFilterChange: (count, rows) => {
        if (rows) this.filteredTableRows = rows;
        if (refs.totalFilasDiv) {
          refs.totalFilasDiv.textContent = t("ptEsparrago.totalRowsShown", { count });
        }
      }
    });

    if (refs.totalFilasDiv) {
      refs.totalFilasDiv.textContent = t("ptEsparrago.totalRowsShown", { count: filtradas.length });
    }

    if (!this.colMenuBound && refs.resultsTable && refs.colMenuEl) {
      bindEsparragoPtColumnMenu(refs.resultsTable, refs.colMenuEl);
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
    const columnLabelsByIndex = getColumnLabelsByIndex();
    const analysis = this.cartillaAnalysis?.present({
      rows: filtradas,
      filasConError,
      errorMap: null,
      duplicateLotes: new Set(),
      colLoteJs: 9,
      columns: headersToAnalysisColumns(this.headers),
      cartilla: CARTILLA_CODE,
      fechaLabel: formatISOToDMY(fecha) || fecha || "—",
      skipSapValidation: PT_SKIP_SAP_VALIDATION,
      translateHeader: (_header, idx) =>
        resolvePtEsparragoColumnLabel(idx, this.headers, columnLabelsByIndex) ||
        String(_header ?? `Col ${idx + 1}`)
    });
    const base = analysis || this.cartillaAnalysis?.getLast?.() || {};
    this.lastAiSnapshot = buildEsparragoPtAiDataBrief(filtradas, base, this.headers);
    setEsparragoPtAiSnapshot(this.lastAiSnapshot);
  }

  /** Tras cambiar idioma: conservar datos y solo refrescar textos/cabeceras. */
  onLanguageChange() {
    if (!this.shell?.refs) return;
    this.renderExcelInsight();
    this.shell.setLiveStatus(this.dataRows.length > 0);

    const refs = this.shell.refs;
    if (refs.resultsHeader?.children?.length) {
      refreshEsparragoPtHeaderLabels(refs.resultsHeader, this.headers);
    }
    if (refs.totalFilasDiv && this.filteredTableRows.length) {
      refs.totalFilasDiv.textContent = t("ptEsparrago.totalRowsShown", {
        count: this.filteredTableRows.length
      });
    }

    const insp = refs.inspectionSelect;
    if (insp?.disabled && insp.options.length === 1 && !insp.value) {
      insp.options[0].textContent = t("ptEsparrago.selectDate");
    }
    const lmr = refs.lmrSelect;
    if (lmr?.disabled && lmr.options.length === 1 && !this.dataRows.length) {
      lmr.options[0].textContent = t("ptEsparrago.lmrAutoDate");
    }
  }

  onExportFiltered() {
    const fecha = this.shell.refs.inspectionSelect?.value;
    const rows = this.filteredTableRows.length
      ? this.filteredTableRows
      : this.getRowsForSelectedDate();
    if (!rows.length) {
      showPtDialog({
        icon: "warning",
        title: t("ptEsparrago.dateRequiredTitle"),
        text: t("ptEsparrago.noExportData")
      });
      return;
    }
    if (!exportEsparragoPtFiltered({ rows, headers: this.headers, fechaLabel: fecha })) {
      showPtDialog({
        icon: "error",
        title: t("plagasArandano.error"),
        text: t("plagasArandano.errorXlsxLibrary")
      });
      return;
    }
    showPtDialog({
      icon: "success",
      title: t("ptEsparrago.exportSuccessTitle"),
      text: t("ptEsparrago.exportSuccessText"),
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
    this.lastAiSnapshot = null;
    clearEsparragoPtAiSnapshot();
  }

  onClear() {
    this.dataRows = [];
    this.headers = [];
    this.excelCabecera = null;
    this.resetResultsUi();
    if (this.shell.refs.fileInput) this.shell.refs.fileInput.value = "";
    if (this.shell.refs.inspectionSelect) {
      this.shell.refs.inspectionSelect.innerHTML = `<option value="" disabled selected>${t("ptEsparrago.selectDate")}</option>`;
      this.shell.refs.inspectionSelect.disabled = true;
    }
    if (this.shell.refs.lmrSelect) {
      this.shell.refs.lmrSelect.innerHTML = `<option value="" selected>${t("ptEsparrago.lmrAutoDate")}</option>`;
    }
    this.shell.setLiveStatus(false);
    this.shell.renderExcelInsightEmpty();
    this.syncButtons();
    showPtDialog({
      icon: "success",
      title: t("ptEsparrago.clearSuccessTitle"),
      text: t("ptEsparrago.clearSuccessText"),
      timer: 1000,
      showConfirmButton: false
    });
  }

  destroy() {
    this.abortController?.abort();
    this.abortController = null;
    this.aiAssistant = null;
    this.lastAiSnapshot = null;
    clearEsparragoPtAiSnapshot();
    this.shell = null;
  }
}
