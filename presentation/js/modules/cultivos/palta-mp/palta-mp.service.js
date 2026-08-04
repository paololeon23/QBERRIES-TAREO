import { appConfig } from "../../../config/app.config.js";
import { AGV_MP_SHELL_IDS } from "../shared/cartilla-shell.ids.js";
import { CartillaShellUi } from "../shared/cartilla-shell.ui.js";
import { hydrateLucideIcons } from "../../../utils/lucide-icon.util.js";
import { i18nService } from "../../../services/i18n.service.js";
import { showMpDialog, showMpConfirmDialog } from "../arandano-mp/arandano-mp-dialog.js";
import {
  CARTILLA_CODE,
  FILAS_SKIP,
  loadPaltaMpValidaciones,
  getTotalColumnas,
  getColInspeccionJs,
  getColLmrJs,
  getColLoteJs,
  getExcelCabecera,
  getValidacionArchivo,
  getPaltaMpValidaciones,
  getStickyCols
} from "./palta-mp.config.js";
import {
  ejecutarValidacion,
  limpiarMarcasValidacion,
  filaTieneError,
  parseExcelDateISO,
  formatISOToDMY,
  valorCelda,
  getCellMeta
} from "./palta-mp.validation.js";
import {
  renderPaltaMpResultsTable,
  htmlTablaFilasConError,
  refreshPaltaMpHeaderLabels
} from "./palta-mp-table.js";
import {
  buildFilteredSheetData,
  buildFullSheetDataWithErrors,
  writePaltaWorkbook
} from "./palta-mp-export.js";
import {
  buildLazyDateDetailPlaceholders,
  bindLazyDateDetailTables
} from "../shared/mp-results-perf.util.js";
import { expandMissingSapLayout } from "../shared/mp-sap-layout.util.js";
import { applyDateDisplayFormatToRows } from "../shared/excel-date-format.util.js";
import { loadSapColumnasCatalog, getSapPerfil } from "../../../config/sap-columnas.registry.js";
import {
  createCartillaAnalysisController,
  filterFilasConErrorExcludingSapOnly,
  headersToAnalysisColumns
} from "../shared/cartilla-analysis.js";
import {
  applyMpColumnVisibility,
  bindMpColumnContextMenu,
  bindMpTableSearch
} from "../shared/mp-column-menu.util.js";
import { syncMpStickyOffsets } from "../shared/mp-sticky-offsets.util.js";

function t(key, vars = {}) {
  let text = i18nService.translate(key);
  Object.entries(vars).forEach(([name, value]) => {
    text = text.replace(`{{${name}}}`, String(value));
  });
  return text;
}

function parseTempNumber(val) {
  const raw = String(valorCelda(val) ?? "")
    .trim()
    .replace(",", ".");
  if (!raw) return null;
  const n = Number(raw);
  return Number.isFinite(n) ? n : null;
}

/** Una sola observación si hay T° Transporte > T° Pulpa (no es error de celda). */
function buildPaltaMpTempObservations(rows) {
  const temps = getPaltaMpValidaciones()?.validaciones_resumen?.temperaturas;
  const pulpaJs = (temps?.pulpa_excel ?? 79) - 1;
  const transporteJs = (temps?.ambiente_excel ?? 81) - 1;
  const hasMayor = rows.some((row) => {
    const pulpa = parseTempNumber(row[pulpaJs]);
    const transporte = parseTempNumber(row[transporteJs]);
    return pulpa != null && transporte != null && transporte > pulpa;
  });
  return hasMayor ? [t("cartillaAnalysis.obsTransporteMayorPulpa")] : [];
}

function htmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function ensureXlsxLibrary() {
  if (window.XLSX?.read && window.XLSX?.utils) return true;
  showMpDialog({
    icon: "error",
    title: t("plagasArandano.error"),
    text: t("plagasArandano.errorXlsxLibrary")
  });
  return false;
}

function computeFechaLmrMayoritaria(rows, colLmrJs) {
  const lmrDates = rows.map((r) => parseExcelDateISO(r[colLmrJs])).filter(Boolean);
  const conteo = {};
  lmrDates.forEach((f) => {
    conteo[f] = (conteo[f] || 0) + 1;
  });
  const sorted = Object.entries(conteo).sort((a, b) => b[1] - a[1]);
  return sorted[0]?.[0] || "";
}

export class PaltaMpService {
  constructor() {
    this.shell = null;
    this.rawRows = [];
    this.headers = [];
    this.excelCabecera = null;
    this.notificationErrors = [];
    this.processedRows = [];
    this.lotesDuplicados = [];
    this.excelLoaded = false;
    this.lastReviewKey = "";
    this.lastReviewAllKey = "";
    this.abortController = null;
    this.root = null;
    this._sapInsertedJs = [];
    this._sapLayoutNotice = null;
    this.searchBound = false;
  }

  async init(appRoot) {
    this.root = appRoot;
    await Promise.all([
      loadPaltaMpValidaciones(appConfig.cacheBustingVersion),
      loadSapColumnasCatalog(appConfig.cacheBustingVersion).catch(() => null)
    ]);

    this.shell = new CartillaShellUi({
      root: appRoot,
      ids: AGV_MP_SHELL_IDS,
      cssPrefix: "agv-mp",
      i18nPrefix: "paltaMp"
    });
    this.shell.cacheDom();
    this.cartillaAnalysis = createCartillaAnalysisController({
      getRoot: () => this.root,
      hostSelector: "#agv-mp-cartilla-analysis",
      showDialog: (opts) => showMpDialog(opts),
      t,
      htmlEscape
    });
    this.bindEvents();
    this.shell.resetDashboard();
    hydrateLucideIcons(appRoot);
  }

  readSheetCell(sheet, fila, columna) {
    const value = sheet[(fila ?? 1) - 1]?.[(columna ?? 1) - 1];
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  parseExcelCabecera(data) {
    const cfg = getExcelCabecera();
    if (!cfg) return null;

    const meta = {
      titulo: this.readSheetCell(data, cfg.titulo.fila, cfg.titulo.columna)
    };

    cfg.campos.forEach((field) => {
      meta[field.clave] = this.readSheetCell(data, field.fila, field.columna);
    });

    return meta;
  }

  syncActionButtons() {
    const refs = this.shell?.refs;
    if (!refs) return;

    const fecha = refs.inspectionSelect?.value || "";
    const reviewKey = fecha ? `MPCP|${fecha}` : "";
    const hasCurrentReview =
      this.excelLoaded && reviewKey && this.lastReviewKey === reviewKey && this.processedRows.length > 0;
    const canUseActions = this.excelLoaded && this.rawRows.length > 0;

    if (refs.runReviewBtn) refs.runReviewBtn.disabled = !canUseActions || !fecha;
    if (refs.exportBtn) refs.exportBtn.disabled = !hasCurrentReview;
    if (refs.reviewAllBtn) refs.reviewAllBtn.disabled = !canUseActions;
    if (refs.exportExcelErroresBtn) {
      refs.exportExcelErroresBtn.disabled = !(canUseActions && this.lastReviewAllKey === CARTILLA_CODE);
    }
  }

  bindEvents() {
    this.abortController?.abort();
    this.abortController = new AbortController();
    const { signal } = this.abortController;
    const refs = this.shell.refs;

    refs.clearBtn?.addEventListener("click", () => this.onClear(), { signal });
    refs.fileInput?.addEventListener("change", (event) => this.onFileSelected(event), { signal });
    refs.inspectionSelect?.addEventListener("change", () => this.onInspectionDateChange(), { signal });
    refs.runReviewBtn?.addEventListener("click", () => this.onRunReview(), { signal });
    refs.reviewAllBtn?.addEventListener("click", () => this.onReviewAll(), { signal });
    refs.exportExcelErroresBtn?.addEventListener("click", () => this.onExportErrors(), { signal });
    refs.exportBtn?.addEventListener("click", () => this.onExportFiltered(), { signal });
    refs.notificationIcon?.addEventListener("click", () => this.onNotificationClick(), { signal });

    this.bindResultsColumnMenu();
    if (!this.searchBound) {
      bindMpTableSearch(refs.tableSearch, refs.resultsBody, {
        idColJs: 0,
        loteColJs: getColLoteJs()
      });
      this.searchBound = true;
    }
  }

  bindResultsColumnMenu() {
    const refs = this.shell?.refs;
    if (!refs?.resultsTable) return;
    let menuEl = refs.colMenuEl;
    if (!menuEl) {
      menuEl = document.getElementById("agv-mp-col-menu");
      if (!menuEl) {
        menuEl = document.createElement("div");
        menuEl.id = "agv-mp-col-menu";
        menuEl.className = "agv-mp-col-menu";
        menuEl.hidden = true;
        menuEl.setAttribute("role", "menu");
        (refs.resultsTable.closest(".agv-mp-table-box") || this.root)?.appendChild(menuEl);
      }
      if (this.shell?.refs) this.shell.refs.colMenuEl = menuEl;
    }
    bindMpColumnContextMenu(refs.resultsTable, menuEl, {
      protectedColIndices: new Set(getStickyCols()),
      onVisibilityChange: () => syncMpStickyOffsets(refs.resultsTable, getStickyCols())
    });
  }

  onClear() {
    this.resetData();
    this.shell.resetDashboard();
    showMpDialog({
      icon: "success",
      title: t("plagasArandano.cleared"),
      text: t("plagasArandano.clearedText"),
      timer: 1200,
      showConfirmButton: false
    });
  }

  resetData() {
    this.rawRows = [];
    this.headers = [];
    this.excelCabecera = null;
    this.notificationErrors = [];
    this.processedRows = [];
    this.lotesDuplicados = [];
    this.excelLoaded = false;
    this.lastReviewKey = "";
    this.lastReviewAllKey = "";
    this._sapInsertedJs = [];
    this._sapLayoutNotice = null;
    this.hideResumenTodasFechas();
    this.hideSingleDateResults();
  }

  async onFileSelected(event) {
    const file = event.target.files?.[0];
    const refs = this.shell.refs;

    if (!file) return;
    if (!ensureXlsxLibrary()) {
      if (refs.fileInput) refs.fileInput.value = "";
      return;
    }

    this.resetData();
    this.shell.resetDashboard({ preserveFileInput: true });

    try {
      const buffer = await file.arrayBuffer();
      const wb = window.XLSX.read(new Uint8Array(buffer), { type: "array" });
      const data = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
        header: 1,
        raw: false
      });

      const valArch = getValidacionArchivo();
      const fila4 = data[valArch.fila_grupo_js ?? 3] || [];
      const cartilla = valorCelda(fila4[valArch.col_grupo_js ?? 8]).toUpperCase().trim();
      const estado = valorCelda(fila4[valArch.col_estado_js ?? 13]).toUpperCase().trim();

      if (cartilla !== CARTILLA_CODE) {
        showMpDialog({
          icon: "error",
          title: "Cartilla no válida",
          html: `Se esperaba <b>${htmlEscape(CARTILLA_CODE)}</b>.<br>Valor encontrado: <b>${htmlEscape(cartilla || "vacío")}</b>`
        });
        if (refs.fileInput) refs.fileInput.value = "";
        return;
      }

      if (estado !== (valArch.estado_esperado ?? "ENVIADA")) {
        showMpDialog({
          icon: "error",
          title: "Estado incorrecto",
          html: `La cartilla debe estar en estado <b>ENVIADA</b>.<br>Valor: <b>${htmlEscape(estado || "vacío")}</b>`
        });
        if (refs.fileInput) refs.fileInput.value = "";
        return;
      }

      this.excelCabecera = this.parseExcelCabecera(data);
      const sheet = data.slice(FILAS_SKIP);

      if (sheet.length < 2) {
        showMpDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("plagasArandano.errorArchivoVacio")
        });
        if (refs.fileInput) refs.fileInput.value = "";
        return;
      }

      const rawHeaders = sheet[0] || [];
      const rawDataRows = sheet
        .slice(1)
        .filter((row) => row.some((cell) => cell !== "" && cell != null));

      // Si Nota Condición no está en col 28 → faltan SAP: insertar 15 + 5 vacías (Hora Insp → 34).
      const {
        headers,
        rows: layoutRows,
        expanded: sapLayoutExpanded,
        insertedSap15,
        insertedSap5,
        insertedJsIndexes
      } = expandMissingSapLayout(rawHeaders, rawDataRows, getSapPerfil("mp"));

      if (headers.length !== getTotalColumnas()) {
        showMpDialog({
          icon: "error",
          title: "Estructura incorrecta",
          html: `El archivo tiene <b>${headers.length}</b> columnas${
            sapLayoutExpanded
              ? ` (tras completar huecos SAP: +${insertedSap15 + insertedSap5})`
              : ""
          }.<br>Se requieren <b>${getTotalColumnas()}</b>.`
        });
        if (refs.fileInput) refs.fileInput.value = "";
        return;
      }

      this.headers = headers;
      // Excel 1-based: inspección 65, LMR 72 (headers mandan; sin 51 genérico).
      this.rawRows = applyDateDisplayFormatToRows(layoutRows, headers, [20, 21, 64, 65, 72]).map(
        (row) => {
          const copy = Array.isArray(row) ? [...row] : [];
          while (copy.length < getTotalColumnas()) copy.push("");
          if (sapLayoutExpanded) copy._sapLayoutExpanded = true;
          return copy;
        }
      );
      this._sapInsertedJs = insertedJsIndexes || [];
      this._sapLayoutNotice = sapLayoutExpanded
        ? { insertedSap15, insertedSap5 }
        : null;

      if (!this.rawRows.length) {
        showMpDialog({
          icon: "error",
          title: t("plagasArandano.error"),
          text: t("plagasArandano.errorArchivoVacio")
        });
        if (refs.fileInput) refs.fileInput.value = "";
        return;
      }

      if (refs.fileFieldEl) refs.fileFieldEl.classList.add("is-loaded");
      if (refs.fileInput) refs.fileInput.title = file.name;

      this.excelLoaded = true;
      this.shell.setLiveStatus(true);
      this.fillInspectionDates();
      this.setNotification(this.detectMissingInspectionDates());
      this.renderExcelInsight();
      this.syncActionButtons();

      const sapNote = this._sapLayoutNotice
        ? `<br><small>Se alineó el bloque SAP (+${this._sapLayoutNotice.insertedSap15} + ${this._sapLayoutNotice.insertedSap5}). Las columnas SAP vacías sí se validan como obligatorias.</small>`
        : "";
      showMpDialog({
        icon: "success",
        title: "Excel cargado",
        html: `Cartilla <b>${htmlEscape(CARTILLA_CODE)}</b> · <b>${this.rawRows.length}</b> registros · <b>${getTotalColumnas()}</b> columnas${sapNote}`,
        timer: sapNote ? 3200 : 1800,
        showConfirmButton: false
      });
    } catch (error) {
      showMpDialog({
        icon: "error",
        title: t("plagasArandano.error"),
        text: error.message || t("plagasArandano.errorArchivoInvalido")
      });
      this.resetData();
      this.shell.resetDashboard();
    }
  }

  fillInspectionDates() {
    const select = this.shell.refs.inspectionSelect;
    if (!select) return;

    const colJs = getColInspeccionJs();
    const fechas = [
      ...new Set(this.rawRows.map((r) => parseExcelDateISO(r[colJs])).filter(Boolean))
    ].sort();

    select.innerHTML = `<option value="" disabled selected>${htmlEscape(t("paltaMp.selectDate"))}</option>`;
    fechas.forEach((iso) => {
      const opt = document.createElement("option");
      opt.value = iso;
      opt.textContent = formatISOToDMY(iso);
      select.appendChild(opt);
    });
    select.disabled = !fechas.length;

    const lmrSelect = this.shell.refs.lmrSelect;
    if (lmrSelect) {
      lmrSelect.innerHTML = `<option value="" selected>${htmlEscape(t("paltaMp.lmrAutoDate"))}</option>`;
      lmrSelect.disabled = true;
      lmrSelect.classList.remove(`${this.shell.cls("input")}--warning`);
    }
  }

  detectMissingInspectionDates() {
    const colJs = getColInspeccionJs();
    return this.rawRows
      .filter((r) => !parseExcelDateISO(r[colJs]))
      .map((r) => ({
        id: valorCelda(r[0]),
        lote: valorCelda(r[getColLoteJs()])
      }));
  }

  setNotification(errors) {
    this.notificationErrors = errors;
    const { notificationIcon, notificationCount } = this.shell.refs;
    if (!notificationIcon || !notificationCount) return;

    if (errors.length > 0) {
      notificationIcon.classList.remove("ok");
      notificationIcon.classList.add("error");
      notificationCount.textContent = String(errors.length);
    } else {
      notificationIcon.classList.remove("error");
      notificationIcon.classList.add("ok");
      notificationCount.textContent = "0";
    }
  }

  onNotificationClick() {
    if (!this.notificationErrors.length) return;
    showMpDialog({
      icon: "warning",
      title: t("paltaMp.missingInspectionTitle"),
      html: `<div class="agv-mp-dialog__html--scroll">
        ${this.notificationErrors
          .map((e) => `• <b>ID:</b> ${htmlEscape(e.id)} &nbsp; <b>Lote:</b> ${htmlEscape(e.lote)}`)
          .join("<br>")}
      </div>`,
      wide: true
    });
  }

  renderExcelInsight() {
    const { excelInsightEl } = this.shell.refs;
    if (!excelInsightEl) return;

    const meta = this.excelCabecera;
    const p = (part) => this.shell.cls(part);

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
                <linearGradient id="pmparInsightRingGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                  <stop offset="0%" stop-color="#5eb8d9"></stop>
                  <stop offset="100%" stop-color="#22c55e"></stop>
                </linearGradient>
              </defs>
              <circle class="${p("excel-insight__ring-glow")}" cx="50" cy="50" r="${ringRadius}"></circle>
              <circle class="${p("excel-insight__ring-track")}" cx="50" cy="50" r="${ringRadius}"></circle>
              <circle class="${p("excel-insight__ring-value")}" cx="50" cy="50" r="${ringRadius}"
                stroke-dasharray="${ringCircumference}" stroke-dashoffset="0"></circle>
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
    const fechaISO = this.shell.refs.inspectionSelect?.value;
    const lmrSelect = this.shell.refs.lmrSelect;
    if (!fechaISO || !lmrSelect) {
      this.syncActionButtons();
      return;
    }

    const colInspeccionJs = getColInspeccionJs();
    const colLmrJs = getColLmrJs();
    const rows = this.rawRows.filter((r) => parseExcelDateISO(r[colInspeccionJs]) === fechaISO);

    const lmrDates = rows.map((r) => parseExcelDateISO(r[colLmrJs])).filter(Boolean);
    const unique = [...new Set(lmrDates)];
    const fechaMayoritaria = computeFechaLmrMayoritaria(rows, colLmrJs);

    lmrSelect.innerHTML = "";
    if (fechaMayoritaria) {
      const opt = document.createElement("option");
      opt.value = fechaMayoritaria;
      opt.textContent = formatISOToDMY(fechaMayoritaria);
      lmrSelect.appendChild(opt);
      lmrSelect.value = fechaMayoritaria;
    } else {
      lmrSelect.innerHTML = `<option value="" selected>${htmlEscape(t("paltaMp.lmrAutoDate"))}</option>`;
    }
    lmrSelect.disabled = true;

    if (unique.length > 1) {
      lmrSelect.classList.add(`${this.shell.cls("input")}--warning`);
      const conteo = {};
      lmrDates.forEach((f) => {
        conteo[f] = (conteo[f] || 0) + 1;
      });
      const detalles = Object.entries(conteo)
        .map(
          ([fecha, count]) =>
            `${formatISOToDMY(fecha)}: <b>${count} registros</b>${fecha === fechaMayoritaria ? " (MAYORITARIA)" : ""}`
        )
        .join("<br>");

      showMpDialog({
        icon: "warning",
        title: t("paltaMp.multipleLmrTitle"),
        html: `<div class="agv-mp-dialog__html--body">
          Se detectaron <b>${unique.length}</b> fechas LMR diferentes:<br><br>
          ${detalles}<br><br>
          Se usará la fecha mayoritaria.
        </div>`,
        wide: true
      });
    } else {
      lmrSelect.classList.remove(`${this.shell.cls("input")}--warning`);
    }

    this.syncActionButtons();
  }

  getRowsForDate(fechaISO) {
    const colJs = getColInspeccionJs();
    return this.rawRows
      .filter((r) => parseExcelDateISO(r[colJs]) === fechaISO)
      .map((row, idx) => {
        const copy = [...row];
        copy._filaNum = idx + 1;
        return copy;
      });
  }

  hideResumenTodasFechas() {
    const el = this.shell.refs.resumenTodasFechasEl;
    if (!el) return;
    el.innerHTML = "";
    el.hidden = true;
  }

  hideSingleDateResults() {
    const refs = this.shell.refs;
    if (refs.resultsHeader) refs.resultsHeader.innerHTML = "";
    if (refs.resultsBody) refs.resultsBody.innerHTML = "";
    if (refs.resultsTable) refs.resultsTable.hidden = true;
    if (refs.tableSearchWrap) refs.tableSearchWrap.hidden = true;
    if (refs.tableSearch) refs.tableSearch.value = "";
    if (refs.resultsSection) {
      refs.resultsSection.classList.remove(
        "is-visible",
        `${this.shell.cls("results")}--ok`,
        `${this.shell.cls("results")}--errors`
      );
    }
    if (refs.resultsSubtitleEl) refs.resultsSubtitleEl.textContent = "";
    if (refs.totalFilasDiv) refs.totalFilasDiv.textContent = "";
    this.cartillaAnalysis?.clear();
    this.processedRows = [];
    this.lastReviewKey = "";
    this.lastReviewAllKey = "";
    this.syncActionButtons();
  }

  async onRunReview() {
    const fechaISO = this.shell.refs.inspectionSelect?.value;

    if (!this.excelLoaded) {
      showMpDialog({ icon: "warning", title: t("plagasArandano.attention"), text: t("paltaMp.noFile") });
      return;
    }

    if (!fechaISO) {
      showMpDialog({
        icon: "warning",
        title: t("paltaMp.missingInspectionTitle"),
        html: "Debes seleccionar una <b>fecha de inspección</b> antes de ejecutar."
      });
      return;
    }

    const rows = this.getRowsForDate(fechaISO);

    const confirm = await showMpConfirmDialog({
      icon: "info",
      title: t("paltaMp.reviewDialogTitle"),
      html: `<div class="agv-mp-dialog__html--compact">
        Se va a revisar la inspección del<br><br>
        <b>${htmlEscape(formatISOToDMY(fechaISO))}</b><br><br>
        <b>${rows.length}</b> registro(s)
      </div>`,
      confirmButtonText: "Continuar",
      cancelButtonText: "Cancelar",
      wide: true
    });

    if (!confirm?.isConfirmed) return;

    this.hideResumenTodasFechas();
    this.runValidationAndRender(rows, fechaISO);
  }

  /** Misma validación (reglas + SAP + fechas 64=20) para una o todas las fechas. */
  validateRowsSameRules(rows) {
    limpiarMarcasValidacion(rows);
    return ejecutarValidacion(rows);
  }

  runValidationAndRender(rows, fechaISO) {
    const { lotesDuplicados } = this.validateRowsSameRules(rows);
    this.lotesDuplicados = lotesDuplicados;
    this.processedRows = rows;

    const filasConErrorAll = rows.filter((row) => filaTieneError(row));
    const filasConError = filterFilasConErrorExcludingSapOnly(filasConErrorAll, {
      errorMap: null,
      duplicateLotes: new Set(lotesDuplicados || []),
      colLoteJs: getColLoteJs(),
      t: (k, v) => t(k, v),
      skipSapValidation: false
    });

    renderPaltaMpResultsTable({
      refs: this.shell.refs,
      headers: this.headers,
      allRows: rows,
      filasConError,
      fechaISO,
      formatISOToDMY,
      t
    });

    this.lastReviewKey = `MPCP|${fechaISO}`;
    this.syncActionButtons();

    const refs = this.shell.refs;
    if (refs.tableSearchWrap) refs.tableSearchWrap.hidden = false;
    if (refs.tableSearch?.value) {
      refs.tableSearch.dispatchEvent(new Event("input", { bubbles: true }));
    }

    this.bindResultsColumnMenu();
    applyMpColumnVisibility(refs.resultsTable);
    requestAnimationFrame(() => syncMpStickyOffsets(refs.resultsTable, getStickyCols()));

    this.cartillaAnalysis?.present({
      rows,
      filasConError,
      errorMap: null,
      duplicateLotes: new Set(lotesDuplicados || []),
      colLoteJs: getColLoteJs(),
      columns: headersToAnalysisColumns(this.headers),
      cartilla: CARTILLA_CODE,
      fechaLabel: formatISOToDMY(fechaISO),
      observations: buildPaltaMpTempObservations(rows)
    });
  }

  buildReviewAllItems() {
    const colJs = getColInspeccionJs();
    const fechas = [
      ...new Set(this.rawRows.map((r) => parseExcelDateISO(r[colJs])).filter(Boolean))
    ].sort();

    // Una pasada por cada fecha de inspección con las mismas reglas (incl. 64 = 20)
    return fechas.map((fechaISO) => {
      const rows = this.getRowsForDate(fechaISO);
      const { lotesDuplicados } = this.validateRowsSameRules(rows);
      const filasDetalle = rows.filter((row) => filaTieneError(row));

      return {
        fecha: formatISOToDMY(fechaISO),
        fechaISO,
        totalFilas: rows.length,
        filasConError: filasDetalle.length,
        filasDetalle,
        lotesDuplicados,
        tieneErrores: filasDetalle.length > 0
      };
    });
  }

  renderResumenTodasFechas(items) {
    const el = this.shell.refs.resumenTodasFechasEl;
    if (!el) return;

    const ok = items.filter((item) => !item.tieneErrores).length;
    const bad = items.length - ok;
    const totalErrors = items.reduce((sum, item) => sum + item.filasConError, 0);
    const totalRows = items.reduce((sum, item) => sum + item.totalFilas, 0);
    const avgRate = totalRows ? Math.round((totalErrors / totalRows) * 100) : 0;

    const tiles = items
      .map((item) => {
        const tileClass = item.tieneErrores ? "agv-mp-tile--error" : "agv-mp-tile--ok";
        const badgeClass = item.tieneErrores ? "agv-mp-tile__badge--error" : "agv-mp-tile__badge--ok";
        const estado = item.tieneErrores ? t("plagasArandano.statusWithIssues") : t("plagasArandano.statusOk");
        const rate = item.totalFilas ? Math.round((item.filasConError / item.totalFilas) * 100) : 0;
        const dupTxt = item.lotesDuplicados?.length
          ? `<p class="agv-mp-tile__dup">${htmlEscape(t("plagasArandano.duplicateLots"))}: ${htmlEscape(item.lotesDuplicados.join(", "))}</p>`
          : "";

        return `<article class="agv-mp-tile ${tileClass}">
          <div class="agv-mp-tile__head">
            <span class="agv-mp-tile__date">${htmlEscape(item.fecha)}</span>
            <span class="agv-mp-tile__badge ${badgeClass}">${htmlEscape(estado)}</span>
          </div>
          <div class="agv-mp-tile__stats">
            <div class="agv-mp-tile__stat">
              <span class="agv-mp-tile__stat-val">${item.totalFilas}</span>
              <span class="agv-mp-tile__stat-lbl">${htmlEscape(t("plagasArandano.tileRecords"))}</span>
            </div>
            <div class="agv-mp-tile__stat">
              <span class="agv-mp-tile__stat-val">${item.filasConError}</span>
              <span class="agv-mp-tile__stat-lbl">${htmlEscape(t("plagasArandano.tileErrors"))}</span>
            </div>
            <div class="agv-mp-tile__stat">
              <span class="agv-mp-tile__stat-val">${rate}%</span>
              <span class="agv-mp-tile__stat-lbl">${htmlEscape(t("plagasArandano.tileRate"))}</span>
            </div>
          </div>
          ${dupTxt}
        </article>`;
      })
      .join("");

    const details = buildLazyDateDetailPlaceholders(items, htmlEscape, (item) =>
      t("plagasArandano.errorRowsCount", { errors: item.filasConError, total: item.totalFilas })
    );

    const detailsBlock = details
      ? `<div class="agv-mp-dashboard__details">
          <h3 class="agv-mp-dashboard__details-title">${htmlEscape(t("plagasArandano.errorsDetailHeading"))}</h3>
          ${details}
        </div>`
      : "";

    el.innerHTML = `
      <div class="agv-mp-dashboard">
        <div>
          <h3 class="agv-mp-dashboard__title">${htmlEscape(t("paltaMp.reviewAllDialogTitle", { cartilla: CARTILLA_CODE }))}</h3>
          <p class="agv-mp-dashboard__subtitle">${htmlEscape(t("plagasArandano.analysisByDate"))}</p>
        </div>
        <div class="agv-mp-kpi-grid">
          <div class="agv-mp-kpi">
            <span class="agv-mp-kpi__icon agv-mp-kpi__icon--dates" aria-hidden="true">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></svg>
            </span>
            <div class="agv-mp-kpi__body">
              <span class="agv-mp-kpi__value">${items.length}</span>
              <span class="agv-mp-kpi__label">${htmlEscape(t("plagasArandano.kpiDates"))}</span>
            </div>
          </div>
          <div class="agv-mp-kpi">
            <span class="agv-mp-kpi__icon agv-mp-kpi__icon--ok" aria-hidden="true">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
            </span>
            <div class="agv-mp-kpi__body">
              <span class="agv-mp-kpi__value">${ok}</span>
              <span class="agv-mp-kpi__label">${htmlEscape(t("plagasArandano.kpiOk"))}</span>
            </div>
          </div>
          <div class="agv-mp-kpi">
            <span class="agv-mp-kpi__icon agv-mp-kpi__icon--error" aria-hidden="true">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/></svg>
            </span>
            <div class="agv-mp-kpi__body">
              <span class="agv-mp-kpi__value">${bad}</span>
              <span class="agv-mp-kpi__label">${htmlEscape(t("plagasArandano.kpiIssues"))}</span>
            </div>
          </div>
          <div class="agv-mp-kpi">
            <span class="agv-mp-kpi__icon agv-mp-kpi__icon--rows" aria-hidden="true">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 6h16M4 12h16M4 18h10"/></svg>
            </span>
            <div class="agv-mp-kpi__body">
              <span class="agv-mp-kpi__value">${totalErrors}</span>
              <span class="agv-mp-kpi__label">${htmlEscape(t("plagasArandano.kpiErrorRows"))}</span>
            </div>
          </div>
        </div>
        <p class="agv-mp-dashboard__avg">${htmlEscape(t("plagasArandano.avgErrorRate", { pct: avgRate }))}</p>
        <div class="agv-mp-tile-grid">${tiles}</div>
        ${detailsBlock}
      </div>`;
    el.hidden = false;
    bindLazyDateDetailTables(el, items, (item) =>
      htmlTablaFilasConError(this.headers, item.filasDetalle, { htmlEscape, t, titled: false })
    );
  }

  onReviewAll() {
    if (!this.excelLoaded) {
      showMpDialog({ icon: "warning", title: t("plagasArandano.attention"), text: t("paltaMp.noFile") });
      return;
    }

    const items = this.buildReviewAllItems();
    if (!items.length) {
      showMpDialog({
        icon: "info",
        title: t("plagasArandano.noDates"),
        text: t("plagasArandano.noDatesText")
      });
      return;
    }

    this.hideSingleDateResults();
    this.renderResumenTodasFechas(items);
    this.lastReviewAllKey = CARTILLA_CODE;
    this.syncActionButtons();

    showMpDialog({
      icon: items.some((item) => item.tieneErrores) ? "warning" : "success",
      title: t("plagasArandano.analysisComplete"),
      text: t("plagasArandano.analysisCompleteText", { count: items.length }),
      timer: 2200,
      showConfirmButton: false
    });
  }

  onExportErrors() {
    if (!this.excelLoaded || !this.rawRows.length) {
      showMpDialog({
        icon: "warning",
        title: t("plagasArandano.attention"),
        text: t("paltaMp.noFile")
      });
      return;
    }
    if (this.lastReviewAllKey !== CARTILLA_CODE) {
      showMpDialog({
        icon: "warning",
        title: "Revisión requerida",
        html: "Primero pulsa <b>Todo</b> para revisar todas las fechas; luego podrás descargar el Excel completo."
      });
      return;
    }

    if (!ensureXlsxLibrary()) return;

    // Botón de «Todo el Excel»: todas las fechas del archivo.
    const rows = this.rawRows.map((row, idx) => {
      const copy = [...row];
      copy._filaNum = idx + 1;
      return copy;
    });
    if (!rows.length) {
      showMpDialog({ icon: "info", title: t("plagasArandano.attention"), text: t("plagasArandano.errorArchivoVacio") });
      return;
    }

    this.validateRowsSameRules(rows);

    const wsData = buildFullSheetDataWithErrors(rows, this.headers, getTotalColumnas(), getCellMeta);
    writePaltaWorkbook(`PALTA_MPCP_Errores_TodasFechas.xlsx`, "MPCP_Errores", wsData);

    showMpDialog({
      icon: "success",
      title: t("paltaMp.exportGenerated"),
      text: t("plagasArandano.exportGeneratedHighlight"),
      timer: 2200,
      showConfirmButton: false
    });
  }

  onExportFiltered() {
    if (this.shell.refs.exportBtn?.disabled) {
      showMpDialog({
        icon: "warning",
        title: t("paltaMp.reviewRequiredTitle"),
        html: "Selecciona <b>fecha de inspección</b>, luego ejecuta <b>Revisar Excel</b> antes de exportar."
      });
      return;
    }

    const fechaISO = this.shell.refs.inspectionSelect?.value;
    if (!fechaISO) {
      showMpDialog({
        icon: "warning",
        title: t("paltaMp.missingInspectionTitle"),
        html: "Selecciona una <b>fecha de inspección</b>."
      });
      return;
    }

    if (!ensureXlsxLibrary()) return;

    const rows = this.getRowsForDate(fechaISO);
    if (!rows.length) {
      showMpDialog({ icon: "info", title: t("plagasArandano.attention"), text: t("plagasArandano.errorArchivoVacio") });
      return;
    }

    const wsData = buildFilteredSheetData(rows, this.headers);
    const fechaLabel = formatISOToDMY(fechaISO).replaceAll("-", "");
    writePaltaWorkbook(`PALTA_MPCP_Filtrado_${fechaLabel}.xlsx`, "MPCP", wsData);

    showMpDialog({
      icon: "success",
      title: t("plagasArandano.exportGenerated"),
      text: `${CARTILLA_CODE}: ${rows.length} inspecciones exportadas.`,
      timer: 2200,
      showConfirmButton: false
    });
  }

  onLanguageChange() {
    const refs = this.shell?.refs;
    if (!refs) return;
    refreshPaltaMpHeaderLabels(refs.resultsHeader, this.headers);
    const count = this.processedRows?.length;
    if (refs.totalFilasDiv && count != null) {
      refs.totalFilasDiv.textContent = t("paltaMp.totalInspectionRows", { count });
    }
    refs.resultsBody?.querySelectorAll("tr.agv-mp-row-ok td").forEach((td) => {
      td.textContent = t("paltaMp.noInspectionErrors");
    });
  }

  destroy() {
    this.abortController?.abort();
    this.abortController = null;
    this.root = null;
    this.shell = null;
  }
}
