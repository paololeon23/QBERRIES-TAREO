import { AGV_PT_SHELL_IDS } from "../shared/cartilla-shell.ids.js";
import { CartillaShellUi } from "../shared/cartilla-shell.ui.js";
import { hydrateLucideIcons } from "../../../utils/lucide-icon.util.js";
import { i18nService } from "../../../services/i18n.service.js";
import { appConfig } from "../../../config/app.config.js";
import { showPtDialog, showPtConfirmDialog, showPtExportChoiceDialog } from "./arandano-pt-dialog.js";
import {
  CARTILLA_RAW_MAP,
  CARTILLA_ORDER,
  FILAS_SKIP,
  CABECERA_CARTILLA,
  CABECERA_ESTADO,
  reorderRow
} from "./arandano-pt.config.js";
import {
  loadReglasPt,
  REGLAS_POR_CARTILLA,
  resolvePlantillaId,
  buildProfileFromReglas,
  fixProfileColsFromReglas,
  computeFechaLmrMayoritaria,
  analyzePtRows,
} from "./arandano-pt-rules.helper.js";
import {
  PT_SKIP_SAP_VALIDATION,
  stripPtSapValidationErrors
} from "../shared/mp-results-perf.util.js";
import {
  renderPtTable,
  bindTableSearch,
  bindColumnContextMenu,
  buildWhatsappReport,
  markDuplicateLoteRows,
  refreshArandanoPtHeaderLabels
} from "./arandano-pt-table.js";
import {
  applyPtTableState,
  savePtTableState,
  clearAllPtTableState
} from "./arandano-pt-state.js";
import { buildPtExportSheetData, exportPtFiltered, exportPtWorkbook } from "./arandano-pt-export.js";
import { collectWhatsappIncidents } from "./arandano-pt.validation.js";
import {
  createCartillaAnalysisController,
  headersToAnalysisColumns
} from "../shared/cartilla-analysis.js";
import { expandMissingSapLayout } from "../shared/mp-sap-layout.util.js";
import { applyDateDisplayFormatToRows } from "../shared/excel-date-format.util.js";
import { loadSapColumnasCatalog, getSapPerfil } from "../../../config/sap-columnas.registry.js";

function t(key, vars = {}) {
  let text = i18nService.translate(key);
  Object.entries(vars).forEach(([name, value]) => {
    text = text.replace(`{{${name}}}`, String(value));
  });
  return text;
}

function htmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function ensureXlsx() {
  if (window.XLSX?.read && window.XLSX?.utils) return true;
  showPtDialog({ icon: "error", title: t("plagasArandano.error"), text: t("plagasArandano.errorXlsxLibrary") });
  return false;
}

function parseExcelDateISO(valor) {
  const texto = String(valor ?? "").trim();
  if (!texto) return "";
  if (typeof valor === "number" && Number.isFinite(valor)) {
    const d = new Date(Math.round((valor - 25569) * 86400 * 1000));
    return d.toISOString().slice(0, 10);
  }
  if (/^\d{8}$/.test(texto)) {
    return `${texto.slice(0, 4)}-${texto.slice(4, 6)}-${texto.slice(6, 8)}`;
  }
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {
    const [d, m, y] = texto.split("/");
    return `${y}-${m}-${d}`;
  }
  const fecha = Date.parse(texto);
  return Number.isFinite(fecha) ? new Date(fecha).toISOString().slice(0, 10) : "";
}

function formatISOToDMY(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}
export class ArandanoPtService {
  constructor() {
    this.shell = null;
    this.headers = [];
    this.dataRows = [];
    this.profile = null;
    this.cartillaStatus = { PTHPA: false, PTLPA: false, PTBPA: false };
    this.currentFilteredRows = [];
    this.errorMap = new Map();
    this.reglasByCartilla = {};
    this.reglasActivas = null;
    this.catalogos = null;
    this.abortController = null;
    this.searchBound = false;
    this.colMenuBound = false;
    this.excelCabeceraByCartilla = {};
    this.exportMetaByCartilla = {};
  }

  async init(appRoot) {
    this.shell = new CartillaShellUi({
      root: appRoot,
      ids: AGV_PT_SHELL_IDS,
      cssPrefix: "agv-pt",
      i18nPrefix: "plagasArandano"
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
    const [loaded] = await Promise.all([
      loadReglasPt(appConfig.cacheBustingVersion),
      loadSapColumnasCatalog(appConfig.cacheBustingVersion)
    ]);
    this.reglasByCartilla = loaded.reglasByCartilla;
    this.catalogos = loaded.catalogos;
    hydrateLucideIcons(appRoot);
  }

  getReglas(cartilla) {
    return this.reglasByCartilla[cartilla] || null;
  }

  readSheetCell(sheet, fila, columna) {
    const value = sheet[(fila ?? 1) - 1]?.[(columna ?? 1) - 1];
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  parseExcelCabecera(sheet, cartilla) {
    const cfg = this.getReglas(cartilla)?.["configuracion-cabecera-excel"];
    if (!cfg) return null;

    const meta = {
      titulo: this.readSheetCell(sheet, cfg.titulo?.fila, cfg.titulo?.columna)
    };

    (cfg.campos || []).forEach((field) => {
      meta[field.clave] = this.readSheetCell(sheet, field.fila, field.columna);
    });

    return meta;
  }

  getInsightCartilla() {
    const selected = this.shell.refs.inspectionTypeSelect?.value;
    if (selected && this.excelCabeceraByCartilla[selected]) return selected;
    return CARTILLA_ORDER.find((cartilla) => this.excelCabeceraByCartilla[cartilla]);
  }

  renderExcelInsight() {
    const { excelInsightEl } = this.shell.refs;
    if (!excelInsightEl) return;

    const cartilla = this.getInsightCartilla();
    const meta = cartilla ? this.excelCabeceraByCartilla[cartilla] : null;
    const p = (part) => this.shell.cls(part);

    if (!meta) {
      this.shell.renderExcelInsightEmpty();
      return;
    }

    const ringRadius = 42;
    const ringCircumference = 2 * Math.PI * ringRadius;
    const grupo = meta.grupo || cartilla || "—";
    const estado = meta.estado || "—";
    const reportTitle = meta.titulo || "";

    excelInsightEl.className = `${p("excel-insight")} ${p("excel-insight")}--loaded`;
    excelInsightEl.innerHTML = `
      ${
        reportTitle
          ? `<p class="${p("excel-insight__report")}">${htmlEscape(reportTitle)}</p>`
          : ""
      }
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

  applyPlantilla(headerRow, cartilla) {
    const reglas = this.getReglas(cartilla);
    if (!reglas || !resolvePlantillaId(headerRow.length, reglas)) return false;

    this.reglasActivas = reglas;
    this.profile = fixProfileColsFromReglas(buildProfileFromReglas(reglas), reglas);

    const reorder = reglas["reorden-ptpha-ptlpa"];
    this._reorderOrder = reorder ? reorder.map((n) => n - 1) : null;
    return true;
  }

  bindEvents() {
    this.abortController?.abort();
    this.abortController = new AbortController();
    const { signal } = this.abortController;
    const refs = this.shell.refs;

    refs.clearBtn?.addEventListener("click", () => this.onClear(), { signal });
    refs.fileInput?.addEventListener("change", (e) => this.onFileSelected(e), { signal });
    refs.inspectionTypeSelect?.addEventListener("change", () => this.onCartillaChange(), { signal });
    refs.inspectionSelect?.addEventListener("change", () => this.onInspectionDateChange(), { signal });
    refs.runReviewBtn?.addEventListener("click", () => this.onRunReview(), { signal });
    refs.exportBtn?.addEventListener("click", () => this.onExportFiltered(), { signal });
  }

  onLanguageChange() {
    const refs = this.shell?.refs;
    if (refs?.resultsHeader?.children?.length) {
      refreshArandanoPtHeaderLabels(refs.resultsHeader, this.headers, this.profile || {});
    }
  }

  destroy() {
    this.abortController?.abort();
    this.abortController = null;
  }

  syncButtons() {
    const refs = this.shell.refs;
    const cartilla = refs.inspectionTypeSelect?.value || "";
    const fecha = refs.inspectionSelect?.value || "";
    const loaded = Boolean(cartilla && this.cartillaStatus[cartilla]);
    if (refs.runReviewBtn) refs.runReviewBtn.disabled = !loaded || !fecha;
    const canExport =
      Boolean(fecha && this.profile) &&
      this.getLoadedCartillas().some((c) => this.getFilteredRows(c, fecha).length > 0);
    if (refs.exportBtn) refs.exportBtn.disabled = !canExport;
  }

  resetResultsUi() {
    const refs = this.shell.refs;
    if (refs.resultsSection) {
      refs.resultsSection.hidden = true;
      refs.resultsSection.classList.remove("is-visible");
    }
    if (refs.tableSearchWrap) refs.tableSearchWrap.hidden = true;
    if (refs.tableSearch) refs.tableSearch.value = "";
    if (refs.resultsHeader) refs.resultsHeader.innerHTML = "";
    if (refs.resultsBody) refs.resultsBody.innerHTML = "";
    if (refs.totalFilasDiv) refs.totalFilasDiv.textContent = "";
    this.currentFilteredRows = [];
    this.cartillaAnalysis?.clear();
    this.syncButtons();
  }

  async onFileSelected(event) {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;
    if (!ensureXlsx()) return;

    this.dataRows = [];
    this.headers = [];
    this.profile = null;
    this.excelCabeceraByCartilla = {};
    this.exportMetaByCartilla = {};
    const cartillasCargadas = new Set();

    for (const file of files) {
      try {
        const buffer = await file.arrayBuffer();
        const wb = window.XLSX.read(new Uint8Array(buffer), { type: "array" });
        const sheetData = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
          header: 1,
          defval: ""
        });

        const fila4 = sheetData[CABECERA_CARTILLA.fila - 1] || [];
        const cartillaRaw = String(fila4[CABECERA_CARTILLA.columna - 1] ?? "").toUpperCase().trim();
        const estado = String(fila4[CABECERA_ESTADO.columna - 1] ?? "").toUpperCase().trim();

        if (!CARTILLA_RAW_MAP[cartillaRaw]) {
          showPtDialog({
            icon: "error",
            title: "Cartilla no permitida",
            html: `Solo se aceptan <b>PTHPAR</b>, <b>PTLPAR</b> y <b>PTBPAR</b>.<br>Archivo: <b>${htmlEscape(file.name)}</b> (${htmlEscape(cartillaRaw || "DESCONOCIDA")}).`
          });
          event.target.value = "";
          return;
        }

        const cartilla = CARTILLA_RAW_MAP[cartillaRaw];
        if (cartillasCargadas.has(cartilla)) {
          showPtDialog({
            icon: "error",
            title: "Cartilla duplicada",
            html: `Ya cargaste la cartilla <b>${htmlEscape(cartilla)}</b>.`
          });
          event.target.value = "";
          return;
        }

        if (estado !== "ENVIADA") {
          showPtDialog({
            icon: "error",
            title: "Estado ≠ ENVIADA",
            html: `El archivo <b>${htmlEscape(file.name)}</b> debe estar en estado <b>ENVIADA</b>.`
          });
          event.target.value = "";
          return;
        }

        this.excelCabeceraByCartilla[cartilla] = this.parseExcelCabecera(sheetData, cartilla);

        sheetData.splice(0, FILAS_SKIP);
        if (!sheetData.length) continue;

        let headerRow = sheetData[0];
        let bodyRows = sheetData.slice(1).filter((r) => r.some((c) => c !== ""));

        // SAP PT: 15 cols + Nota Condición (28) + 5 cols → alinea plantilla 100.
        const {
          headers: sapHeaders,
          rows: sapRows,
          expanded: sapExpanded,
          insertedSap15,
          insertedSap5,
          insertedJsIndexes
        } = expandMissingSapLayout(headerRow, bodyRows, getSapPerfil("pt"));
        headerRow = sapHeaders;
        bodyRows = sapRows;
        const headerOriginal = [...headerRow];

        if (!this.applyPlantilla(headerRow, cartilla)) {
          showPtDialog({
            icon: "error",
            title: "Estructura incorrecta",
            html: `El archivo <b>${htmlEscape(file.name)}</b> tiene <b>${headerRow.length}</b> columnas.<br>Se requiere plantilla de <b>100</b> columnas.`
          });
          event.target.value = "";
          return;
        }

        if (this._reorderOrder) {
          headerRow = this._reorderOrder.map((i) => headerRow[i]);
          bodyRows = bodyRows.map((row) => reorderRow(row, this._reorderOrder));
        }

        const loadReorderRegla = this.getReglas(cartilla)?.["reorden-ptpha-ptlpa"];
        this.exportMetaByCartilla[cartilla] = {
          headerOriginal,
          loadReorder: loadReorderRegla ? loadReorderRegla.map((n) => n - 1) : null,
          sapExpanded,
          insertedSap15,
          insertedSap5,
          insertedJsIndexes: insertedJsIndexes || []
        };

        if (!this.headers.length) {
          this.headers = headerRow;
        }

        if (headerRow.length !== this.profile.totalColumnas) {
          showPtDialog({
            icon: "error",
            title: "Columnas inconsistentes",
            html: `Todos los archivos deben usar la misma plantilla (${this.profile.totalColumnas} columnas).`
          });
          event.target.value = "";
          return;
        }

        // Excel 1-based desde rules: inspección 42, LMR 49 (+ cosecha/prod. SAP).
        const filas = applyDateDisplayFormatToRows(bodyRows, headerRow, [20, 21, 42, 49])
          .filter((r) => r.some((c) => c !== ""))
          .map((r) => {
            const copy = [...r];
            copy.__cartilla = cartilla;
            return copy;
          })
          .sort((a, b) =>
            String(a[this.profile.cols.usuario] ?? "")
              .toUpperCase()
              .localeCompare(String(b[this.profile.cols.usuario] ?? "").toUpperCase())
          );

        this.dataRows.push(...filas);
        cartillasCargadas.add(cartilla);
        this.cartillaStatus[cartilla] = true;
      } catch (err) {
        console.error(err);
        showPtDialog({
          icon: "error",
          title: "Error al leer archivo",
          text: `No se pudo procesar ${file.name}.`
        });
        event.target.value = "";
        return;
      }
    }

    this.populateCartillaSelect(Array.from(cartillasCargadas));
    this.resetResultsUi();
    this.shell.setLiveStatus(true);

    const refs = this.shell.refs;
    if (refs.fileFieldEl) refs.fileFieldEl.classList.add("is-loaded");
    if (refs.fileInput && files.length) {
      refs.fileInput.title = files.map((f) => f.name).join(", ");
    }

    this.renderExcelInsight();

    if (cartillasCargadas.size) {
      const first = Array.from(cartillasCargadas)[0];
      this.shell.refs.inspectionTypeSelect.value = first;
      this.updateDatesForCartilla(first);
    }

    if (this.dataRows.length) {
      const cartillasLabel = Array.from(cartillasCargadas).join(", ");
      const sapNotes = PT_SKIP_SAP_VALIDATION
        ? []
        : CARTILLA_ORDER
            .map((c) => this.exportMetaByCartilla[c])
            .filter((m) => m?.sapExpanded)
            .map((m) => `+${m.insertedSap15 || 0}/+${m.insertedSap5 || 0}`);
      const sapNote = sapNotes.length
        ? `<br><small>Huecos SAP completados (${sapNotes.join(" · ")}) — Nota Condición col 28 / bloque col 34.</small>`
        : "";
      showPtDialog({
        icon: "success",
        title: "Excel cargado",
        html: `Cartilla(s) <b>${htmlEscape(cartillasLabel)}</b> · <b>${this.dataRows.length}</b> registros · <b>${this.profile.totalColumnas}</b> columnas${sapNote}`,
        timer: sapNote ? 3200 : 1800,
        showConfirmButton: false
      });
    }
  }

  populateCartillaSelect(cartillas) {
    const select = this.shell.refs.inspectionTypeSelect;
    if (!select) return;
    select.innerHTML = `<option value="" disabled selected>${htmlEscape(t("arandanoPt.selectCartilla"))}</option>`;
    cartillas.forEach((c) => {
      const opt = document.createElement("option");
      opt.value = c;
      opt.textContent = c;
      select.appendChild(opt);
    });
    select.disabled = false;
  }

  updateDatesForCartilla(cartilla) {
    const refs = this.shell.refs;
    const colFecha = this.profile.cols.fechaInspeccion;
    const fechas = [
      ...new Set(
        this.dataRows
          .filter((r) => r.__cartilla === cartilla)
          .map((r) => parseExcelDateISO(r[colFecha]))
          .filter(Boolean)
      )
    ].sort();

    refs.inspectionSelect.innerHTML = `<option value="" disabled selected>${htmlEscape(t("plagasArandano.selectDate"))}</option>`;
    fechas.forEach((iso) => {
      const opt = document.createElement("option");
      opt.value = iso;
      opt.textContent = formatISOToDMY(iso);
      refs.inspectionSelect.appendChild(opt);
    });
    refs.inspectionSelect.disabled = !fechas.length;

    if (refs.lmrSelect) {
      refs.lmrSelect.innerHTML = `<option value="" selected>${htmlEscape(t("arandanoPt.lmrAutoDate"))}</option>`;
      refs.lmrSelect.disabled = true;
    }

    refs.runReviewBtn.disabled = true;
    this.syncButtons();
  }

  onCartillaChange() {
    this.resetResultsUi();
    const cartilla = this.shell.refs.inspectionTypeSelect?.value;
    if (cartilla) this.updateDatesForCartilla(cartilla);
    this.renderExcelInsight();
  }

  onInspectionDateChange() {
    const refs = this.shell.refs;
    const fechaISO = refs.inspectionSelect?.value;
    const cartilla = refs.inspectionTypeSelect?.value;
    const lmrSelect = refs.lmrSelect;
    if (!fechaISO || !cartilla || !this.profile || !lmrSelect) return;

    const colFecha = this.profile.cols.fechaInspeccion;
    const colLmr = this.profile.cols.fechaLmr;
    const rows = this.dataRows.filter(
      (r) => r.__cartilla === cartilla && parseExcelDateISO(r[colFecha]) === fechaISO
    );

    const lmrDates = rows.map((r) => parseExcelDateISO(r[colLmr])).filter(Boolean);
    const unique = [...new Set(lmrDates)];
    const fechaMayoritaria = computeFechaLmrMayoritaria(rows, colLmr);

    lmrSelect.innerHTML = "";
    if (fechaMayoritaria) {
      const opt = document.createElement("option");
      opt.value = fechaMayoritaria;
      opt.textContent = formatISOToDMY(fechaMayoritaria);
      lmrSelect.appendChild(opt);
      lmrSelect.value = fechaMayoritaria;
    } else {
      lmrSelect.innerHTML = `<option value="" selected>${htmlEscape(t("arandanoPt.lmrAutoDate"))}</option>`;
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
          ([f, count]) =>
            `${formatISOToDMY(f)}: <b>${count} registros</b>${f === fechaMayoritaria ? " (MAYORITARIA)" : ""}`
        )
        .join("<br>");

      showPtDialog({
        icon: "warning",
        title: "Múltiples fechas LMR detectadas",
        html: `<div class="agv-pt-dialog__html agv-pt-dialog__html--stacked">
          Se detectaron <b>${unique.length}</b> fechas LMR diferentes:<br><br>
          ${detalles}<br><br>
          Se usará la fecha mayoritaria. Las filas con fechas minoritarias se marcarán como error.
        </div>`,
        wide: true
      });
    } else {
      lmrSelect.classList.remove(`${this.shell.cls("input")}--warning`);
    }

    this.syncButtons();
  }

  getFilteredRows(cartilla, fecha) {
    const colFecha = this.profile.cols.fechaInspeccion;
    return this.dataRows.filter(
      (r) => r.__cartilla === cartilla && parseExcelDateISO(r[colFecha]) === fecha
    );
  }

  getLoadedCartillas() {
    return CARTILLA_ORDER.filter((cartilla) => this.cartillaStatus[cartilla]);
  }

  getRowsForExport(cartilla, fechaISO) {
    // Preferir el orden actual de la tabla (como lo dejó el usuario)
    if (
      this.currentFilteredRows?.length &&
      this.shell.refs.inspectionTypeSelect?.value === cartilla &&
      this.shell.refs.inspectionSelect?.value === fechaISO
    ) {
      markDuplicateLoteRows(this.currentFilteredRows, this.profile.cols.lote);
      return [...this.currentFilteredRows];
    }
    const rows = this.getFilteredRows(cartilla, fechaISO);
    markDuplicateLoteRows(rows, this.profile.cols.lote);
    return applyPtTableState(cartilla, fechaISO, rows, this.profile);
  }

  countInspectionsByCartilla(fechaISO) {
    const counts = {};
    this.getLoadedCartillas().forEach((cartilla) => {
      counts[cartilla] = this.getFilteredRows(cartilla, fechaISO).length;
    });
    return counts;
  }

  exportFormatHelpers() {
    return { formatISOToDMY, parseExcelDateISO };
  }

  buildExportSheet(cartilla, rows) {
    return buildPtExportSheetData(
      rows,
      cartilla,
      this.exportMetaByCartilla[cartilla],
      this.getReglas(cartilla)?.["configuracion-exportacion"] || {},
      this.exportFormatHelpers()
    );
  }

  exportFilteredCartilla(cartilla, fechaISO) {
    const rows = this.getRowsForExport(cartilla, fechaISO);
    if (!rows.length) {
      showPtDialog({ icon: "warning", title: "Sin datos", text: "No hay filas para exportar." });
      return;
    }

    exportPtFiltered({
      rows,
      cartilla,
      fechaLabel: formatISOToDMY(fechaISO),
      exportCfg: this.getReglas(cartilla)?.["configuracion-exportacion"] || {},
      exportMeta: this.exportMetaByCartilla[cartilla],
      helpers: this.exportFormatHelpers()
    });

    showPtDialog({
      icon: "success",
      title: t("arandanoPt.exportSuccess"),
      text: `${cartilla}: ${rows.length} filas exportadas.`,
      timer: 2200,
      showConfirmButton: false
    });
  }

  exportFilteredAllCartillas(fechaISO) {
    const loaded = this.getLoadedCartillas();
    const sheets = [];

    loaded.forEach((cartilla) => {
      const rows = this.getRowsForExport(cartilla, fechaISO);
      if (!rows.length) return;
      sheets.push({
        name: cartilla,
        data: this.buildExportSheet(cartilla, rows)
      });
    });

    if (!sheets.length) {
      showPtDialog({ icon: "warning", title: "Sin datos", text: "No hay filas para exportar." });
      return;
    }

    const fechaLabel = formatISOToDMY(fechaISO).replace(/\//g, "-");
    exportPtWorkbook({
      filename: `PT_Arandano_Filtrado_${fechaLabel}.xlsx`,
      sheets
    });

    showPtDialog({
      icon: "success",
      title: t("arandanoPt.exportSuccess"),
      text: `Exportadas ${sheets.length} cartilla(s) en un solo archivo.`,
      timer: 2200,
      showConfirmButton: false
    });
  }

  async promptExportFilteredChoice(cartilla, fechaISO) {
    const counts = this.countInspectionsByCartilla(fechaISO);
    const loaded = this.getLoadedCartillas();
    const fechaLabel = formatISOToDMY(fechaISO);

    if (loaded.length <= 1) {
      this.exportFilteredCartilla(cartilla, fechaISO);
      return;
    }

    const countsHtml = loaded
      .map(
        (c) =>
          `<li class="agv-mp-export-choice__item">
            <span class="agv-mp-export-choice__code">${htmlEscape(c)}</span>
            <span class="agv-mp-export-choice__count">${counts[c] ?? 0} inspecciones</span>
          </li>`
      )
      .join("");
    const totalDia = loaded.reduce((sum, c) => sum + (counts[c] ?? 0), 0);

    const result = await showPtExportChoiceDialog({
      title: "Exportar Excel filtrado",
      html: `<div class="agv-mp-export-choice">
        <div class="agv-mp-export-choice__meta">
          <span class="agv-mp-export-choice__meta-label">Fecha de inspección</span>
          <strong class="agv-mp-export-choice__meta-value">${htmlEscape(fechaLabel)}</strong>
        </div>
        <div class="agv-mp-export-choice__section">
          <span class="agv-mp-export-choice__section-label">Registros por cartilla</span>
          <ul class="agv-mp-export-choice__list">${countsHtml}</ul>
        </div>
        <p class="agv-mp-export-choice__total">
          <span>Total del día</span>
          <strong>${totalDia} inspecciones</strong>
        </p>
      </div>`,
      choices: [
        {
          id: "current",
          label: `Solo cartilla ${cartilla} (${counts[cartilla] ?? 0} registros)`
        },
        {
          id: "all",
          label: `Todas las cartillas cargadas (${loaded.length} hojas)`
        }
      ]
    });

    if (result.action === "current") {
      this.exportFilteredCartilla(cartilla, fechaISO);
    } else if (result.action === "all") {
      this.exportFilteredAllCartillas(fechaISO);
    }
  }

  async onExportFiltered() {
    if (!ensureXlsx()) return;
    const cartilla = this.shell.refs.inspectionTypeSelect?.value;
    const fecha = this.shell.refs.inspectionSelect?.value;

    if (!cartilla || !fecha) {
      showPtDialog({
        icon: "warning",
        title: "Datos incompletos",
        html: "Debes seleccionar <b>cartilla</b> y <b>fecha de inspección</b>."
      });
      return;
    }

    await this.promptExportFilteredChoice(cartilla, fecha);
  }

  async onRunReview() {
    const refs = this.shell.refs;
    const cartilla = refs.inspectionTypeSelect?.value;
    const fecha = refs.inspectionSelect?.value;
    if (!cartilla || !fecha) return;

    const colLmr = this.profile.cols.fechaLmr;
    const rowsPreview = this.getFilteredRows(cartilla, fecha);
    const lmrDates = [...new Set(rowsPreview.map((r) => parseExcelDateISO(r[colLmr])).filter(Boolean))];
    const lmrDisplay = formatISOToDMY(refs.lmrSelect?.value || lmrDates[0] || "");

    const confirm = await showPtConfirmDialog({
      icon: "info",
      title: t("arandanoPt.reviewConfirmTitle"),
      html: `<div class="agv-pt-dialog__html agv-pt-dialog__html--stacked">
        Se va a revisar:<br><br>
        <b>Cartilla:</b> ${htmlEscape(cartilla)}<br><br>
        <b>Fecha inspección:</b> ${htmlEscape(formatISOToDMY(fecha))}<br><br>
        <b>Fecha LMR (mayoritaria):</b> ${htmlEscape(lmrDisplay || "—")}
      </div>`,
      wide: true
    });

    if (!confirm?.isConfirmed) return;

    const rows = [...rowsPreview];
    rows.forEach((r, idx) => {
      r._filaNum = idx + 1;
    });

    const fechaLmrMayoritaria = computeFechaLmrMayoritaria(rows, colLmr);
    const { errorMap } = analyzePtRows(
      rows,
      this.getReglas(cartilla),
      this.catalogos,
      cartilla,
      fechaLmrMayoritaria
    );

    // Por ahora: en PT no se validan ni reportan datos SAP (sí Nota Condición y resto).
    this.errorMap = PT_SKIP_SAP_VALIDATION ? stripPtSapValidationErrors(errorMap) : errorMap;
    rows.forEach((r) => {
      const loteCol = this.profile.cols.lote + 1;
      const err = errorMap.get(r._filaNum)?.get(loteCol);
      if (err?.tipo === "duplicado") r.__duplicado = true;
    });
    markDuplicateLoteRows(rows, this.profile.cols.lote);
    this.currentFilteredRows = applyPtTableState(cartilla, fecha, rows, this.profile);

    // Mantener dataRows alineado con el orden/agrupación que el usuario dejó
    const colFechaSync = this.profile.cols.fechaInspeccion;
    const others = this.dataRows.filter(
      (r) => !(r.__cartilla === cartilla && parseExcelDateISO(r[colFechaSync]) === fecha)
    );
    this.dataRows = others.concat(this.currentFilteredRows);

    if (!this.searchBound) {
      bindTableSearch(refs.tableSearch, refs.resultsBody);
      this.searchBound = true;
    }
    if (!this.colMenuBound) {
      bindColumnContextMenu(refs.resultsTable, refs.colMenuEl);
      this.colMenuBound = true;
    }

    this.renderTable(this.currentFilteredRows, fecha);

    if (refs.resultsSection) {
      refs.resultsSection.hidden = false;
      refs.resultsSection.classList.add("is-visible");
    }
    if (refs.tableSearchWrap) refs.tableSearchWrap.hidden = false;
    if (refs.resultsTable) refs.resultsTable.hidden = false;
    refs.totalFilasDiv.textContent = `${this.currentFilteredRows.length} inspecciones`;
    this.syncButtons();

    const filasConError = this.currentFilteredRows.filter((r) => {
      const filaMap = this.errorMap?.get?.(r._filaNum);
      return Boolean((filaMap && filaMap.size > 0) || r.__duplicado);
    });
    const duplicateLotes = new Set(
      this.currentFilteredRows
        .filter((r) => r.__duplicado)
        .map((r) => String(r[this.profile.cols.lote] ?? "").trim())
        .filter(Boolean)
    );
    this.cartillaAnalysis?.present({
      rows: this.currentFilteredRows,
      filasConError,
      errorMap: this.errorMap || null,
      duplicateLotes,
      colLoteJs: this.profile?.cols?.lote ?? 9,
      columns: headersToAnalysisColumns(this.headers),
      cartilla: cartilla || "—",
      fechaLabel: formatISOToDMY(fecha),
      skipSapValidation: PT_SKIP_SAP_VALIDATION
    });
    this.schedulePersistTableState();
  }

  persistTableState() {
    const cartilla = this.shell.refs.inspectionTypeSelect?.value;
    const fecha = this.shell.refs.inspectionSelect?.value;
    if (!cartilla || !fecha || !this.profile) return;
    savePtTableState(cartilla, fecha, this.currentFilteredRows, this.profile);
  }

  schedulePersistTableState() {
    if (this._persistPtTimer) clearTimeout(this._persistPtTimer);
    this._persistPtTimer = setTimeout(() => {
      this._persistPtTimer = null;
      this.persistTableState();
    }, 0);
  }

  renderTable(rows, fecha) {
    renderPtTable({
      refs: this.shell.refs,
      rows,
      headers: this.headers,
      profile: this.profile,
      fechaInspeccion: fecha,
      errorMap: this.errorMap,
      onReorder: (from, to) => this.onRowReorder(from, to),
      onCopyReport: (row) => this.onCopyWhatsapp(row),
      onRowMark: () => this.schedulePersistTableState()
    });
  }

  onRowReorder(from, to) {
    if (from === to) return;
    const cartilla = this.shell.refs.inspectionTypeSelect?.value;
    const fecha = this.shell.refs.inspectionSelect?.value;
    const tbody = this.shell.refs.resultsBody;

    const moved = this.currentFilteredRows.splice(from, 1)[0];
    this.currentFilteredRows.splice(to, 0, moved);

    const colFecha = this.profile.cols.fechaInspeccion;
    const others = this.dataRows.filter(
      (r) => !(r.__cartilla === cartilla && parseExcelDateISO(r[colFecha]) === fecha)
    );
    this.dataRows = others.concat(this.currentFilteredRows);

    if (tbody?.children[from]) {
      const tr = tbody.children[from];
      tbody.removeChild(tr);
      const ref = to >= tbody.children.length ? null : tbody.children[to];
      tbody.insertBefore(tr, ref);
      [...tbody.children].forEach((rowEl, idx) => {
        rowEl.dataset.rowIndex = String(idx);
      });
    }

    this.schedulePersistTableState();
  }

  async onCopyWhatsapp(row) {
    const fecha = this.shell.refs.inspectionSelect?.value;
    const incidencias = collectWhatsappIncidents(row, this.profile, fecha);
    const msg = buildWhatsappReport(row, this.profile, incidencias);
    if (!msg) {
      showPtDialog({
        icon: "info",
        title: "WhatsApp",
        text: "Fila Correcta — Sin errores.",
        timer: 1200,
        showConfirmButton: false
      });
      return;
    }
    try {
      await navigator.clipboard.writeText(msg);
      showPtDialog({
        icon: "success",
        title: "Copiado",
        text: "Listo para WhatsApp.",
        timer: 1000,
        showConfirmButton: false
      });
    } catch {
      showPtDialog({ icon: "error", title: "Error", text: "No se pudo copiar al portapapeles." });
    }
  }

  onClear() {
    clearAllPtTableState();
    this.excelCabeceraByCartilla = {};
    this.exportMetaByCartilla = {};
    this.dataRows = [];
    this.headers = [];
    this.profile = null;
    this.reglasActivas = null;
    this.errorMap = new Map();
    this.cartillaStatus = { PTHPA: false, PTLPA: false, PTBPA: false };
    this.resetResultsUi();

    const refs = this.shell.refs;
    if (refs.fileInput) refs.fileInput.value = "";
    if (refs.inspectionTypeSelect) {
      refs.inspectionTypeSelect.innerHTML = `<option value="" disabled selected>${htmlEscape(t("arandanoPt.selectCartilla"))}</option>`;
      refs.inspectionTypeSelect.disabled = true;
    }
    if (refs.inspectionSelect) {
      refs.inspectionSelect.innerHTML = `<option value="" disabled selected>${htmlEscape(t("plagasArandano.selectDate"))}</option>`;
      refs.inspectionSelect.disabled = true;
    }
    if (refs.lmrSelect) {
      refs.lmrSelect.innerHTML = `<option value="" selected>${htmlEscape(t("arandanoPt.lmrAutoDate"))}</option>`;
      refs.lmrSelect.disabled = true;
    }

    this.shell.resetDashboard();
    this.shell.renderExcelInsightEmpty();
    this.shell.setLiveStatus(false);

    showPtDialog({
      icon: "success",
      title: t("plagasArandano.cleared"),
      text: t("plagasArandano.clearedText"),
      timer: 1200,
      showConfirmButton: false
    });
  }
}
