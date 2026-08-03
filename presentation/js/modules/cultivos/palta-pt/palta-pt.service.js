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
  loadPaltaPtValidaciones,
  getTotalColumnas,
  getColEmbalajeJs,
  getColCosechaJs,
  getExcelCabecera,
  getValidacionArchivo
} from "./palta-pt.config.js";
import { normalizeDateValue } from "./palta-pt.validation.js";
import {
  renderPaltaPtTable,
  bindTableSearch,
  bindPaltaPtColumnMenu,
  buildWhatsappReport,
  refreshPaltaPtHeaderLabels
} from "./palta-pt-table.js";
import { exportPaltaPtFiltered } from "./palta-pt-export.js";
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

function uniquePackagingDates(rows) {
  const col = getColEmbalajeJs();
  return [...new Set(rows.map((r) => normalizeDateValue(r[col])).filter(Boolean))].sort();
}

function uniqueHarvestDates(rows, fechaEmbalaje) {
  const colEmb = getColEmbalajeJs();
  const colCos = getColCosechaJs();
  return [
    ...new Set(
      rows
        .filter((r) => normalizeDateValue(r[colEmb]) === fechaEmbalaje)
        .map((r) => normalizeDateValue(r[colCos]))
        .filter(Boolean)
    )
  ].sort();
}

function rowMatchesFilters(row, fechaEmbalaje, fechaCosecha) {
  const colEmb = getColEmbalajeJs();
  if (normalizeDateValue(row[colEmb]) !== fechaEmbalaje) return false;
  if (!fechaCosecha) return true;
  return normalizeDateValue(row[getColCosechaJs()]) === fechaCosecha;
}

export class PaltaPtService {
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
    await loadPaltaPtValidaciones(appConfig.cacheBustingVersion);

    this.shell = new CartillaShellUi({
      root: appRoot,
      ids: AGV_PT_SHELL_IDS,
      cssPrefix: "agv-pt",
      i18nPrefix: "ptPalta"
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
    refs.inspectionSelect?.addEventListener("change", () => this.onPackagingDateChange(), { signal });
    refs.cosechaSelect?.addEventListener("change", () => this.syncButtons(), { signal });
    refs.runReviewBtn?.addEventListener("click", () => this.onRunReview(), { signal });
    refs.exportBtn?.addEventListener("click", () => this.onExportFiltered(), { signal });
    refs.clearBtn?.addEventListener("click", () => this.onClear(), { signal });
  }

  syncButtons() {
    const hasFile = this.dataRows.length > 0;
    const hasEmbalaje = Boolean(this.shell.refs.inspectionSelect?.value);
    const disabled = !hasFile || !hasEmbalaje;
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

      const valArch = getValidacionArchivo();
      const cartillaCell = data[valArch.fila_grupo_js ?? 3]?.[valArch.col_grupo_js ?? 8];
      if (String(cartillaCell ?? "").trim().toUpperCase() !== CARTILLA_CODE) {
        showPtDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("ptPalta.invalidCartilla")
        });
        event.target.value = "";
        return;
      }

      const estadoCell = data[valArch.fila_estado_js ?? 3]?.[valArch.col_estado_js ?? 13];
      if (valArch.estado_esperado && String(estadoCell ?? "").trim().toUpperCase() !== valArch.estado_esperado) {
        showPtDialog({
          icon: "warning",
          title: t("ptPalta.warnTitle"),
          text: t("ptPalta.warnEstado", { estado: valArch.estado_esperado })
        });
      }

      this.headers = data[HEADER_ROW_INDEX] || [];
      if (this.headers.length < getTotalColumnas()) {
        showPtDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("ptPalta.invalidColumns", {
            expected: getTotalColumnas(),
            found: this.headers.length
          })
        });
        event.target.value = "";
        return;
      }

      this.excelCabecera = this.parseExcelCabecera(data);
      // Excel 1-based: embalaje 55, LMR 72. No usar 51 (en PTCP Destino/otras cols).
      this.dataRows = applyDateDisplayFormatToRows(
        data.slice(DATA_START_INDEX).filter((row) => row.some((c) => String(c ?? "").trim())),
        this.headers,
        [20, 21, 55, 72]
      );
      this.resetResultsUi();

      const fechas = uniquePackagingDates(this.dataRows);
      const sel = this.shell.refs.inspectionSelect;
      if (sel) {
        sel.innerHTML = `<option value="" disabled selected>${t("ptPalta.selectPackagingDate")}</option>`;
        fechas.forEach((f) => sel.add(new Option(f, f)));
        sel.disabled = fechas.length === 0;
      }

      this.resetHarvestSelect();

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

  resetHarvestSelect() {
    const cosechaSel = this.shell.refs.cosechaSelect;
    if (!cosechaSel) return;
    cosechaSel.innerHTML = `<option value="" selected>${t("ptPalta.selectHarvestDate")}</option>`;
    cosechaSel.disabled = true;
    cosechaSel.classList.add(`${this.shell.cls("input")}--readonly`);
  }

  onPackagingDateChange() {
    const fechaEmb = this.shell.refs.inspectionSelect?.value;
    const cosechaSel = this.shell.refs.cosechaSelect;
    this.resetResultsUi();

    if (!fechaEmb || !cosechaSel) {
      this.syncButtons();
      return;
    }

    const cosechas = uniqueHarvestDates(this.dataRows, fechaEmb);
    cosechaSel.innerHTML = "";
    if (cosechas.length === 0) {
      cosechaSel.add(new Option(t("ptPalta.noHarvestDate"), "", true, true));
    } else if (cosechas.length === 1) {
      cosechaSel.add(new Option(cosechas[0], cosechas[0], true, true));
    } else {
      cosechaSel.add(new Option(t("ptPalta.allHarvestDates"), "", true, true));
      cosechas.forEach((c) => cosechaSel.add(new Option(c, c)));
    }
    cosechaSel.disabled = true;
    cosechaSel.classList.add(`${this.shell.cls("input")}--readonly`);

    this.syncButtons();
  }

  getRowsForSelectedFilters() {
    const fechaEmb = this.shell.refs.inspectionSelect?.value;
    if (!fechaEmb) return [];
    const fechaCos = this.shell.refs.cosechaSelect?.value || "";
    return this.dataRows.filter((r) => rowMatchesFilters(r, fechaEmb, fechaCos));
  }

  async onRunReview() {
    const fechaEmb = this.shell.refs.inspectionSelect?.value;
    if (!fechaEmb) {
      showPtDialog({
        icon: "warning",
        title: t("ptPalta.dateRequiredTitle"),
        text: t("ptPalta.dateRequiredText")
      });
      return;
    }

    const filtradas = this.getRowsForSelectedFilters();
    this.filteredTableRows = filtradas;

    const refs = this.shell.refs;
    if (refs.resultsSection) {
      refs.resultsSection.hidden = false;
      refs.resultsSection.classList.add("is-visible");
    }
    if (refs.tableSearchWrap) refs.tableSearchWrap.hidden = false;
    if (refs.resultsTable) refs.resultsTable.hidden = false;

    const onCopyReport = async (row, fechaEmbalaje) => {
      const msg = buildWhatsappReport(row, fechaEmbalaje);
      if (!msg) {
        await showPtDialog({
          icon: "info",
          title: t("ptPalta.rowOkTitle"),
          text: t("ptPalta.rowOkText"),
          timer: 1000,
          showConfirmButton: false
        });
        return;
      }
      try {
        await navigator.clipboard.writeText(msg);
        await showPtDialog({
          icon: "success",
          title: t("ptPalta.copySuccessTitle"),
          text: t("ptPalta.copySuccessText"),
          timer: 1200,
          showConfirmButton: false
        });
      } catch {
        await showPtDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("ptPalta.copyErrorText")
        });
      }
    };

    renderPaltaPtTable({
      headerRow: refs.resultsHeader,
      bodyRows: refs.resultsBody,
      headers: this.headers,
      tableEl: refs.resultsTable,
      allRowsForDate: filtradas,
      filteredRows: filtradas,
      fechaEmbalaje: fechaEmb,
      onFilterChange: (count, rows) => {
        if (rows) this.filteredTableRows = rows;
        if (refs.totalFilasDiv) {
          refs.totalFilasDiv.textContent = t("ptPalta.totalRowsShown", { count });
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
      refs.totalFilasDiv.textContent = t("ptPalta.totalRowsShown", { count: filtradas.length });
    }

    if (!this.colMenuBound && refs.resultsTable && refs.colMenuEl) {
      bindPaltaPtColumnMenu(refs.resultsTable, refs.colMenuEl);
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
      fechaLabel: fechaEmb || "—",
      skipSapValidation: PT_SKIP_SAP_VALIDATION
    });
  }

  onExportFiltered() {
    const fechaEmb = this.shell.refs.inspectionSelect?.value;
    const rows = this.filteredTableRows.length
      ? this.filteredTableRows
      : this.getRowsForSelectedFilters();
    if (!rows.length) {
      showPtDialog({
        icon: "warning",
        title: t("ptPalta.dateRequiredTitle"),
        text: t("ptPalta.noExportData")
      });
      return;
    }
    if (!exportPaltaPtFiltered({ rows, headers: this.headers, fechaLabel: fechaEmb })) {
      showPtDialog({
        icon: "error",
        title: t("plagasArandano.error"),
        text: t("plagasArandano.errorXlsxLibrary")
      });
      return;
    }
    showPtDialog({
      icon: "success",
      title: t("ptPalta.exportSuccessTitle"),
      text: t("ptPalta.exportSuccessText"),
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
      this.shell.refs.inspectionSelect.innerHTML = `<option value="" disabled selected>${t("ptPalta.selectPackagingDate")}</option>`;
      this.shell.refs.inspectionSelect.disabled = true;
    }
    this.resetHarvestSelect();
    this.shell.setLiveStatus(false);
    this.shell.renderExcelInsightEmpty();
    this.syncButtons();
    showPtDialog({
      icon: "success",
      title: t("ptPalta.clearSuccessTitle"),
      text: t("ptPalta.clearSuccessText"),
      timer: 1000,
      showConfirmButton: false
    });
  }

  onLanguageChange() {
    refreshPaltaPtHeaderLabels(this.shell?.refs?.resultsHeader, this.headers);
  }

  destroy() {
    this.abortController?.abort();
    this.abortController = null;
    this.shell = null;
  }
}
