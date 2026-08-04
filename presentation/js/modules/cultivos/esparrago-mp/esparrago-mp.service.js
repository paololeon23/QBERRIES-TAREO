import { AGV_MP_SHELL_IDS } from "../shared/cartilla-shell.ids.js";
import { CartillaShellUi } from "../shared/cartilla-shell.ui.js";
import { cargarReglasDesdeRuta, analizarReporte } from "../../../../../engine/rule-engine.js";
import { hydrateLucideIcons } from "../../../utils/lucide-icon.util.js";
import { i18nService } from "../../../services/i18n.service.js";
import { appConfig } from "../../../config/app.config.js";
import { showMpDialog, showMpConfirmDialog, showMpExportChoiceDialog } from "../arandano-mp/arandano-mp-dialog.js";
import {
  buildFilteredSheetData,
  buildFullSheetDataWithErrors
} from "./esparrago-mp-export.js";
import { translateExcelHeader } from "../../../utils/excel-header-i18n.util.js";
import { refreshTranslatedHeaderRow } from "../../../utils/table-header-i18n.util.js";
import {
  createCartillaAnalysisController,
  filterFilasConErrorExcludingSapOnly
} from "./esparrago-mp-cartilla-analysis.js";
import {
  collectValidatedColumnIndexesJs,
  DEFAULT_MP_CONTEXT_COLS_JS,
  SAP_ZONE_FRONTEND_COLS_JS,
  resolveSapZoneHeader,
  SAP_ZONE_HEADER_LABELS_BY_JS
} from "../shared/mp-results-perf.util.js?v=2026072219";
import { expandMissingSapLayout } from "../shared/mp-sap-layout.util.js?v=2026072219";
import { applyDateDisplayFormatToRows } from "../shared/excel-date-format.util.js";
import { loadSapColumnasCatalog, getSapPerfil } from "../../../config/sap-columnas.registry.js";
import {
  applyMpColumnVisibility,
  bindMpColumnContextMenu
} from "../shared/mp-column-menu.util.js";
import {
  DEFAULT_MP_STICKY_WIDTHS,
  syncMpStickyOffsets
} from "../shared/mp-sticky-offsets.util.js";

const STICKY_COLUMNS = [0, 1, 6, 9]; // Id, Inspección código, Usuario, Lote
/** Anchos base sticky Espárrago MP (Usuario ancho por emails). */
const STICKY_COL_WIDTHS = { ...DEFAULT_MP_STICKY_WIDTHS };
/** Columnas de contexto siempre visibles en la tabla de errores (JS 0-based). */
const CONTEXT_COLUMNS_JS = DEFAULT_MP_CONTEXT_COLS_JS;
/** Productor…Peso Bruto (Excel 13–33). */
const SAP_ZONE_COLS_JS = SAP_ZONE_FRONTEND_COLS_JS || [
  12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32
];
const FILAS_SKIP = 5;

/** Columnas Excel (1-based) para verificación visual de sumas en tabla de revisión */
const ESP_MP_DEFECTOS_COLS = [
  35, 36, 37, 38, 41, 42, 43, 44, 45, 50, 58, 59, 62, 63, 64, 65, 66, 67, 68, 69, 80, 81
];

const ESP_MP_SUM_VERIFICATIONS = [
  {
    header: "SUMA DEFECTOS",
    sumCols: ESP_MP_DEFECTOS_COLS,
    displayOnly: true,
    mismatchHint: "Suma de defectos (sin Flácido ni Calidad punta C+)"
  },
  {
    header: "SUMA EXPORTABLE CALIBRES (g)",
    sumCols: [46, 53, 54, 70, 72, 73],
    compareCol: 77,
    compareLabel: "Muestra Exp. Calibres (g)",
    mismatchHint: "Debe coincidir con Muestra Exp. Calibres (g)"
  },
  {
    header: "SUMA EXPORTABLE LONGITUD",
    sumCols: [83, 84, 85, 86, 87],
    compareCol: 79,
    compareLabel: "Muestra Exp. Longitud",
    mismatchHint: "Debe coincidir con Muestra Exp. Longitud"
  },
  {
    header: "SUMA EXPORTABLE CALIBRES (und)",
    sumCols: [52, 55, 71, 74, 75, 89],
    compareCol: 78,
    compareLabel: "Muestra Exp. Calibres (und)",
    mismatchHint: "Debe coincidir con Muestra Exp. Calibres (und)"
  },
  {
    header: "SUMA DEFECTOS + MUESTRA EXP. CALIBRES (g)",
    sumCols: [...ESP_MP_DEFECTOS_COLS, 77],
    compareCol: 11,
    compareLabel: "Cant Muestra",
    mismatchHint: "Debe coincidir con Cant Muestra"
  }
];
const CARTILLA_ORDER = ["MPES"];
const REGLAS_PATH = "rules/modulos/esparrago-mp-mpes.rules.json";

function isPinnedColumn(index) {
  return STICKY_COLUMNS.includes(index);
}

function applyStickyColumnClasses(el, index) {
  if (!isPinnedColumn(index)) return;
  el.classList.add("agv-mp-sticky-col", `agv-mp-sticky-col-${index}`);
}

/** Anchos fijos + left acumulado. No tocar z-index inline (pisa thead sticky). */
function syncEsparragoMpStickyOffsets(tableEl) {
  syncMpStickyOffsets(tableEl, STICKY_COLUMNS, STICKY_COL_WIDTHS);
}

function excelColToJs(col) {
  return col - 1;
}

function rowNumericAtExcelCol(row, excelCol) {
  const raw = row[excelColToJs(excelCol)];
  if (raw === null || raw === undefined || raw === "") return 0;
  const n = Number.parseFloat(raw);
  return Number.isFinite(n) ? n : 0;
}

function sumExcelCols(row, excelCols) {
  return excelCols.reduce((acc, col) => acc + rowNumericAtExcelCol(row, col), 0);
}

function sumMatchesExpected(suma, expected, tolerance = 0) {
  return Math.abs(suma - expected) <= tolerance;
}

function getSumVerificationMeta(row, def) {
  const suma = sumExcelCols(row, def.sumCols);
  if (def.displayOnly) {
    return {
      value: String(suma),
      cellClass: "agv-mp-cell-calculated",
      title: def.mismatchHint
    };
  }
  const expected = rowNumericAtExcelCol(row, def.compareCol);
  const ok = sumMatchesExpected(suma, expected);
  const compareLabel = def.compareLabel || "valor esperado";
  return {
    value: String(suma),
    cellClass: ok ? "agv-mp-cell-calculated" : "agv-mp-cell-calculated agv-mp-cell-calculated--error",
    title: ok
      ? def.mismatchHint
      : `${def.mismatchHint}. Calculado: ${suma}, ${compareLabel}: ${expected}`
  };
}

function appendSumVerificationHeaders(headerRow) {
  ESP_MP_SUM_VERIFICATIONS.forEach((def) => {
    const th = document.createElement("th");
    th.className = "agv-mp-table__col-header agv-mp-sum-col-header";
    th.dataset.excelHeader = def.header;
    th.textContent = translateExcelHeader(def.header);
    th.title = def.mismatchHint;
    headerRow.appendChild(th);
  });
}

function appendSumVerificationCells(tr, row) {
  ESP_MP_SUM_VERIFICATIONS.forEach((def) => {
    const { value, cellClass, title } = getSumVerificationMeta(row, def);
    const td = document.createElement("td");
    td.className = cellClass;
    td.title = title;
    td.textContent = value;
    tr.appendChild(td);
  });
}

function htmlSumVerificationHeaders() {
  return ESP_MP_SUM_VERIFICATIONS.map(
    (def) =>
      `<th class="agv-mp-table__col-header agv-mp-sum-col-header" title="${htmlEscape(def.mismatchHint)}">${htmlEscape(def.header)}</th>`
  ).join("");
}

function htmlSumVerificationCells(row) {
  return ESP_MP_SUM_VERIFICATIONS.map((def) => {
    const { value, cellClass, title } = getSumVerificationMeta(row, def);
    const classAttr = cellClass ? ` class="${htmlEscape(cellClass)}"` : "";
    const titleAttr = title ? ` title="${htmlEscape(title)}"` : "";
    return `<td${classAttr}${titleAttr}>${htmlEscape(value)}</td>`;
  }).join("");
}

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

function parseExcelDateISO(valor) {
  const texto = String(valor ?? "").trim();
  if (!texto) return "";
  if (/^\d{8}$/.test(texto)) {
    return `${texto.slice(0, 4)}-${texto.slice(4, 6)}-${texto.slice(6, 8)}`;
  }
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {
    const [d, m, y] = texto.split("/");
    return `${y}-${m}-${d}`;
  }
  if (/^\d{2}-\d{2}-\d{4}$/.test(texto)) {
    const [d, m, y] = texto.split("-");
    return `${y}-${m}-${d}`;
  }
  const fecha = Date.parse(texto);
  return Number.isFinite(fecha) ? new Date(fecha).toISOString().slice(0, 10) : "";
}

function formatISOToDMY(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}-${m}-${y}`;
}

function valorCeldaParaMostrar(val) {
  if (val === null || val === undefined) return "";
  return String(val);
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

function buildCartillaSummaryHtml(cartillaStatus) {
  return `<div class="agv-mp-dialog__cartilla-grid">
    ${CARTILLA_ORDER.map((cartilla) => {
      const ok = cartillaStatus[cartilla];
      return `<article class="agv-mp-dialog__cartilla-card agv-mp-dialog__cartilla-card--${ok ? "ok" : "missing"}">
        <span class="agv-mp-dialog__cartilla-name">${htmlEscape(cartilla)}</span>
        <span class="agv-mp-dialog__cartilla-state">${ok ? "Tiene data" : "No tiene data"}</span>
      </article>`;
    }).join("")}
  </div>`;
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

function rowToRegistro(row, filaNum) {
  const _cols = {};
  const len = row.length;
  for (let idx = 0; idx < len; idx++) {
    _cols[String(idx + 1)] = row[idx] ?? "";
  }
  return { fila: filaNum, _cols };
}

/** Contexto + zona SAP (se valida) + columnas con reglas + columnas con error. */
function pickResultColumnsForCartilla(cartilla, allColumns, errorMap, filas, getReglas, headersByCartilla) {
  const baseCols = collectValidatedColumnIndexesJs(getReglas(cartilla), {
    contextColsJs: CONTEXT_COLUMNS_JS,
    stickyColsJs: STICKY_COLUMNS,
    includeSapZone: true
  });
  const needed = new Set([...CONTEXT_COLUMNS_JS, ...STICKY_COLUMNS, ...SAP_ZONE_COLS_JS, ...baseCols]);
  errorMap?.forEach((filaMap) => {
    filaMap?.forEach((_err, colNum) => {
      const js = Number(colNum) - 1;
      if (Number.isFinite(js) && js >= 0) needed.add(js);
    });
  });
  (filas || []).forEach((row) => {
    const filaMap = errorMap?.get(row._filaNum);
    if (!filaMap) return;
    filaMap.forEach((_err, colNum) => {
      const js = Number(colNum) - 1;
      if (Number.isFinite(js) && js >= 0) needed.add(js);
    });
  });

  const headers = headersByCartilla?.[cartilla] || [];
  const byIndex = new Map();
  (allColumns || []).forEach((col) => {
    if (!col || col.originalIndex == null) return;
    const idx = Number(col.originalIndex);
    if (!Number.isFinite(idx)) return;
    byIndex.set(idx, col);
  });
  needed.forEach((idx) => {
    const excelHeader = headers[idx] || byIndex.get(idx)?.header || "";
    const sapMeta = SAP_ZONE_HEADER_LABELS_BY_JS[idx]
      ? resolveSapZoneHeader(idx, excelHeader)
      : null;
    const header = sapMeta?.label || String(excelHeader || "").trim() || `Col ${idx + 1}`;
    const existing = byIndex.get(idx);
    byIndex.set(idx, {
      id: existing?.id || `col_${idx + 1}`,
      header,
      originalIndex: idx,
      headerTitle: sapMeta?.title || header,
      isSapZone: Boolean(sapMeta)
    });
  });

  return [...byIndex.values()]
    .filter((col) => needed.has(Number(col.originalIndex)))
    .sort((a, b) => Number(a.originalIndex) - Number(b.originalIndex));
}

function groupRowsByInspectionDate(rows, colFechaJs) {
  const byDate = new Map();
  rows.forEach((row) => {
    const iso = row._fechaInspeccionISO || parseExcelDateISO(row[colFechaJs]);
    if (!iso) return;
    if (!row._fechaInspeccionISO) row._fechaInspeccionISO = iso;
    if (!byDate.has(iso)) byDate.set(iso, []);
    byDate.get(iso).push(row);
  });
  return byDate;
}

function buildCompuestaColumnMap(reglas) {
  const map = new Map();
  (reglas?.["validaciones-compuestas"] || []).forEach((regla) => {
    const msg = regla["si-falla-mostrar"] || "";
    if (
      regla.tipo === "diferencia-maxima-columnas" ||
      regla.tipo === "igual-entre-columnas" ||
      regla.tipo === "fecha-no-mayor-que" ||
      regla.tipo === "fecha-menor-o-igual"
    ) {
      map.set(msg, [regla["columna-a"], regla["columna-b"]].filter(Boolean));
    }
    if (regla.tipo === "fecha-lmr-mayoritaria") {
      map.set(msg, [regla["columna-lmr"]].filter(Boolean));
    }
  });
  return map;
}

function buildErrorMap(resultado, compuestaColumnMap) {
  const map = new Map();

  const addError = (fila, colNum, tipo, problema) => {
    if (!colNum) return;
    if (!map.has(fila)) map.set(fila, new Map());
    const filaMap = map.get(fila);
    const existing = filaMap.get(colNum);
    if (!existing || tipo === "obligatorio") {
      filaMap.set(colNum, { tipo, problema });
    }
  };

  (resultado?.columnasDetalle || []).forEach((col) => {
    const colNum = col.numeroColumna;
    (col.detalle || []).forEach((item) => {
      if (col.esCompuesta) {
        const cols = compuestaColumnMap.get(item.problema) || compuestaColumnMap.get(col.nombreColumna);
        if (cols?.length) {
          cols.forEach((n) => addError(item.fila, n, item.tipo, item.problema));
        } else {
          addError(item.fila, colNum, item.tipo, item.problema);
        }
      } else {
        addError(item.fila, colNum, item.tipo, item.problema);
      }
    });
  });

  return map;
}

function detectDuplicateLotes(rows, colLoteJs) {
  const conteo = {};
  rows.forEach((row) => {
    const lote = String(row[colLoteJs] ?? "").trim();
    if (!lote) return;
    conteo[lote] = (conteo[lote] || 0) + 1;
  });
  return new Set(Object.keys(conteo).filter((lote) => conteo[lote] > 1));
}

function estiloExportCeldaError(cellClass) {
  const isEmpty = cellClass === "agv-mp-cell-error-empty";
  if (isEmpty) {
    return {
      fill: { patternType: "solid", fgColor: { rgb: "FFC94C4C" } },
      font: { color: { rgb: "FFFFFFFF" }, bold: true }
    };
  }
  return {
    fill: { patternType: "solid", fgColor: { rgb: "FFFEE2E2" } },
    font: { color: { rgb: "FFC94C4C" }, bold: true }
  };
}

export class EsparragoMpService {
  constructor() {
    this.shell = null;
    this.reglasByCartilla = {};
    this.reglas = null;
    this.cfgCartillas = null;
    this.rawRows = [];
    this.headersByCartilla = {};
    this.columnsByCartilla = {};
    this.excelCabeceraByCartilla = {};
    this.cartillaStatus = { MPES: false };
    this.notificationErrors = [];
    this.processedRows = [];
    this.lastErrorMap = new Map();
    this.duplicateLotes = new Set();
    this.excelLoaded = false;
    this.lastReviewKey = "";
    this.lastReviewAllKey = "";
    this.abortController = null;
    this.root = null;
    this.compuestaColumnMapByCartilla = {};
    this.cartillaAnalysis = null;
  }

  async init(appRoot) {
    this.root = appRoot;
    const [reglas] = await Promise.all([
      cargarReglasDesdeRuta(`${REGLAS_PATH}?v=${appConfig.cacheBustingVersion}`),
      loadSapColumnasCatalog(appConfig.cacheBustingVersion)
    ]);
    this.reglasByCartilla = { MPES: reglas };
    this.reglas = this.reglasByCartilla.MPES;
    this.cfgCartillas = this.reglas?.["configuracion-cartillas"] || {};
    CARTILLA_ORDER.forEach((cartilla) => {
      this.compuestaColumnMapByCartilla[cartilla] = buildCompuestaColumnMap(
        this.reglasByCartilla[cartilla]
      );
    });

    this.shell = new CartillaShellUi({
      root: appRoot,
      ids: AGV_MP_SHELL_IDS,
      cssPrefix: "agv-mp",
      i18nPrefix: "plagasEsparrago"
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

  getReglas(cartilla) {
    return this.reglasByCartilla[cartilla] || this.reglas;
  }

  getCfgCartilla(cartilla) {
    return this.getReglas(cartilla)?.["configuracion-cartillas"] || {};
  }

  colCartillaJsFor(cartilla) {
    return (this.getCfgCartilla(cartilla)["columna-cartilla"] || 2) - 1;
  }

  colFechaInspeccionJsFor(cartilla) {
    return (this.getCfgCartilla(cartilla)["columna-fecha-inspeccion"] || 41) - 1;
  }

  colFechaLmrJsFor(cartilla) {
    return (this.getCfgCartilla(cartilla)["columna-fecha-lmr"] || 51) - 1;
  }

  colLoteJsFor(cartilla) {
    return (this.getCfgCartilla(cartilla)["columna-lote"] || 10) - 1;
  }

  get colCartillaJs() {
    return this.colCartillaJsFor(this.getActiveCartilla() || "MPES");
  }

  get colFechaInspeccionJs() {
    return this.colFechaInspeccionJsFor(this.getActiveCartilla() || "MPES");
  }

  get colFechaLmrJs() {
    return this.colFechaLmrJsFor(this.getActiveCartilla() || "MPES");
  }

  get colLoteJs() {
    return this.colLoteJsFor(this.getActiveCartilla() || "MPES");
  }

  totalColumnasFor(cartilla) {
    return this.getReglas(cartilla)?.["total-columnas"] || 104;
  }

  getCodigosPermitidos() {
    return { MPES: "MPES" };
  }

  getActiveCartilla() {
    return "MPES";
  }

  compuestaMapFor(cartilla) {
    return this.compuestaColumnMapByCartilla[cartilla] || new Map();
  }

  get totalColumnas() {
    return this.totalColumnasFor("MPES");
  }

  readSheetCell(sheet, fila, columna) {
    const value = sheet[(fila ?? 1) - 1]?.[(columna ?? 1) - 1];
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  parseExcelCabecera(sheet) {
    const cfg = this.reglas?.["configuracion-cabecera-excel"];
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
    const selected = this.getActiveCartilla();
    if (selected && this.excelCabeceraByCartilla[selected]) return selected;
    return CARTILLA_ORDER.find((cartilla) => this.excelCabeceraByCartilla[cartilla]);
  }

  syncActionButtons() {
    const refs = this.shell?.refs;
    if (!refs) return;

    const cartilla = this.getActiveCartilla();
    const fecha = refs.inspectionSelect?.value || "";
    const reviewKey = cartilla && fecha ? `${cartilla}|${fecha}` : "";
    const hasCurrentReview =
      this.excelLoaded &&
      reviewKey &&
      this.lastReviewKey === reviewKey &&
      this.processedRows.length > 0;
    const canUseTodoActions =
      this.excelLoaded && Boolean(cartilla) && Boolean(this.cartillaStatus[cartilla]);
    const hasReviewedAll =
      canUseTodoActions && this.lastReviewAllKey === cartilla;

    if (refs.runReviewBtn) {
      refs.runReviewBtn.disabled = !canUseTodoActions || !fecha;
    }
    if (refs.exportBtn) {
      refs.exportBtn.disabled = !hasCurrentReview;
    }
    if (refs.reviewAllBtn) {
      refs.reviewAllBtn.disabled = !canUseTodoActions;
    }
    if (refs.exportExcelErroresBtn) {
      refs.exportExcelErroresBtn.disabled = !hasReviewedAll;
    }
  }

  getSelectedCartillaOrWarn() {
    const cartilla = this.getActiveCartilla();
    if (!this.excelLoaded || !this.rawRows.length) {
      showMpDialog({
        icon: "warning",
        title: t("plagasArandano.attention"),
        text: t("arandanoMp.noFile")
      });
      return "";
    }
    if (!this.excelLoaded || !this.cartillaStatus.MPES) {
      showMpDialog({
        icon: "warning",
        title: t("plagasArandano.attention"),
        text: t("arandanoMp.noFile")
      });
      return "";
    }
    return "MPES";
  }

  getCartillaRows(_cartilla) {
    return this.rawRows.map((row, idx) => {
      const copy = [...row];
      copy._filaNum = idx + 1;
      if (row._fechaInspeccionISO) copy._fechaInspeccionISO = row._fechaInspeccionISO;
      return copy;
    });
  }

  buildValidationStateForCartilla(cartilla) {
    const rows = this.getCartillaRows(cartilla);
    const errorMap = new Map();
    const duplicateRows = new Set();
    const colFechaInspeccionJs = this.colFechaInspeccionJsFor(cartilla);
    const colFechaLmrJs = this.colFechaLmrJsFor(cartilla);
    const colLoteJs = this.colLoteJsFor(cartilla);
    const reglas = this.getReglas(cartilla);
    const compuestaMap = this.compuestaMapFor(cartilla);
    const byDate = groupRowsByInspectionDate(rows, colFechaInspeccionJs);

    byDate.forEach((fechaRows) => {
      const fechaLmrMayoritaria = computeFechaLmrMayoritaria(fechaRows, colFechaLmrJs);
      const registros = fechaRows.map((row) => rowToRegistro(row, row._filaNum));
      const resultado = analizarReporte(
        { filas: registros, cultivo: "esparrago" },
        reglas,
        { cartilla, fechaLmrMayoritaria }
      );
      const partialMap = buildErrorMap(resultado, compuestaMap);

      partialMap.forEach((filaMap, filaNum) => {
        if (!errorMap.has(filaNum)) errorMap.set(filaNum, new Map());
        filaMap.forEach((err, colNum) => {
          errorMap.get(filaNum).set(colNum, err);
        });
      });

      const dupLotes = detectDuplicateLotes(fechaRows, colLoteJs);
      fechaRows.forEach((row) => {
        const lote = String(row[colLoteJs] ?? "").trim();
        if (lote && dupLotes.has(lote)) duplicateRows.add(row._filaNum);
      });
    });

    return { rows, errorMap, duplicateRows };
  }

  getCellExportMeta(row, colJs, errorMap, duplicateLotes, colLoteJs) {
    const colNum = colJs + 1;
    const val = valorCeldaParaMostrar(row[colJs]);
    const err = errorMap.get(row._filaNum)?.get(colNum);

    if (err) {
      if (err.tipo === "obligatorio") {
        return { val, cellClass: "agv-mp-cell-error-empty" };
      }
      return { val, cellClass: "agv-mp-cell-error-value" };
    }

    const lote = String(row[colLoteJs] ?? "").trim();
    if (colNum === colLoteJs + 1 && duplicateLotes.has(lote)) {
      return { val, cellClass: "agv-mp-cell-error-value" };
    }

    return { val, cellClass: "" };
  }

  getLoadedCartillas() {
    return CARTILLA_ORDER.filter((cartilla) => this.cartillaStatus[cartilla]);
  }

  getRowsForCartillaFecha(cartilla, fechaISO) {
    const colFechaJs = this.colFechaInspeccionJsFor(cartilla);
    return this.rawRows
      .filter((r) => (r._fechaInspeccionISO || parseExcelDateISO(r[colFechaJs])) === fechaISO)
      .map((row, idx) => {
        const copy = [...row];
        copy._filaNum = idx + 1;
        copy._fechaInspeccionISO = row._fechaInspeccionISO || fechaISO;
        return copy;
      });
  }

  countInspectionsByCartilla(fechaISO) {
    const counts = {};
    this.getLoadedCartillas().forEach((cartilla) => {
      counts[cartilla] = this.getRowsForCartillaFecha(cartilla, fechaISO).length;
    });
    return counts;
  }

  buildValidationStateForRows(cartilla, rows) {
    const colLoteJs = this.colLoteJsFor(cartilla);
    const colFechaLmrJs = this.colFechaLmrJsFor(cartilla);
    const fechaLmrMayoritaria = computeFechaLmrMayoritaria(rows, colFechaLmrJs);
    const registros = rows.map((row) => rowToRegistro(row, row._filaNum));
    const resultado = analizarReporte(
      { filas: registros, cultivo: "esparrago" },
      this.getReglas(cartilla),
      { cartilla, fechaLmrMayoritaria }
    );
    const errorMap = buildErrorMap(resultado, this.compuestaMapFor(cartilla));
    const duplicateLotes = detectDuplicateLotes(rows, colLoteJs);
    return { errorMap, duplicateLotes };
  }

  exportFormatHelpers() {
    return {
      formatISOToDMY,
      parseExcelDateISO,
      estiloExportCeldaError
    };
  }

  getExportConfig(cartilla) {
    return this.getReglas(cartilla)?.["configuracion-exportacion"] || {};
  }

  writeWorkbook(filename, sheets) {
    const wb = window.XLSX.utils.book_new();
    sheets.forEach(({ name, data }) => {
      const ws = window.XLSX.utils.aoa_to_sheet(data);
      window.XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    });
    window.XLSX.writeFile(wb, filename);
  }

  exportFilteredCartilla(cartilla, fechaISO) {
    if (!ensureXlsxLibrary()) return;

    const rows = this.getRowsForCartillaFecha(cartilla, fechaISO);
    if (!rows.length) {
      showMpDialog({
        icon: "info",
        title: t("plagasArandano.attention"),
        text: t("plagasArandano.errorArchivoVacio")
      });
      return;
    }

    const headers = this.headersByCartilla[cartilla] || [];
    const exportCfg = this.getExportConfig(cartilla);
    const wsData = buildFilteredSheetData(rows, cartilla, headers, exportCfg, this.exportFormatHelpers());
    const fechaLabel = formatISOToDMY(fechaISO).replaceAll("-", "");
    const nombre = `ESPARRAGO_${cartilla}_Filtrado_${fechaLabel}.xlsx`;

    this.writeWorkbook(nombre, [{ name: cartilla, data: wsData }]);

    showMpDialog({
      icon: "success",
      title: t("plagasArandano.exportGenerated"),
      text: `${cartilla}: ${rows.length} inspecciones exportadas.`,
      timer: 2200,
      showConfirmButton: false
    });
  }

  exportFilteredAllCartillas(fechaISO) {
    if (!ensureXlsxLibrary()) return;

    const sheets = [];
    const loaded = this.getLoadedCartillas();

    loaded.forEach((cartilla) => {
      const rows = this.getRowsForCartillaFecha(cartilla, fechaISO);
      if (!rows.length) return;
      const headers = this.headersByCartilla[cartilla] || [];
      const exportCfg = this.getExportConfig(cartilla);
      sheets.push({
        name: cartilla,
        data: buildFilteredSheetData(rows, cartilla, headers, exportCfg, this.exportFormatHelpers())
      });
    });

    if (!sheets.length) {
      showMpDialog({
        icon: "info",
        title: t("plagasArandano.attention"),
        text: t("plagasArandano.errorArchivoVacio")
      });
      return;
    }

    const fechaLabel = formatISOToDMY(fechaISO).replaceAll("-", "");
    const nombre = `ESPARRAGO_MP_Filtrado_${fechaLabel}.xlsx`;
    this.writeWorkbook(nombre, sheets);

    showMpDialog({
      icon: "success",
      title: t("plagasArandano.exportGenerated"),
      text: `Exportadas ${sheets.length} cartilla(s) en un solo archivo.`,
      timer: 2200,
      showConfirmButton: false
    });
  }

  exportExcelConErroresResaltados(cartilla, fechaISO) {
    if (!ensureXlsxLibrary()) return;

    const rows = fechaISO
      ? this.getRowsForCartillaFecha(cartilla, fechaISO)
      : this.getCartillaRows(cartilla);

    if (!rows.length) {
      showMpDialog({
        icon: "info",
        title: t("plagasArandano.attention"),
        text: t("plagasArandano.errorArchivoVacio")
      });
      return;
    }

    const headers = this.headersByCartilla[cartilla] || [];
    const totalCols = this.totalColumnasFor(cartilla);
    const { errorMap, duplicateLotes } = this.buildValidationStateForRows(cartilla, rows);
    const colLoteJs = this.colLoteJsFor(cartilla);
    const exportCfg = this.getExportConfig(cartilla);

    const wsData = buildFullSheetDataWithErrors(
      rows,
      headers,
      totalCols,
      exportCfg,
      (row, colJs) => this.getCellExportMeta(row, colJs, errorMap, duplicateLotes, colLoteJs),
      this.exportFormatHelpers()
    );

    const fechaSuffix = fechaISO ? `_${formatISOToDMY(fechaISO).replaceAll("-", "")}` : "_TodasFechas";
    const nombre = `ESPARRAGO_${cartilla}_Errores${fechaSuffix}.xlsx`;
    this.writeWorkbook(nombre, [{ name: `${cartilla}_Errores`, data: wsData }]);

    showMpDialog({
      icon: "success",
      title: t("arandanoMp.exportGenerated"),
      text: t("plagasArandano.exportGeneratedHighlight"),
      timer: 2200,
      showConfirmButton: false
    });
  }

  async promptExportFilteredChoice(cartilla, fechaISO) {
    const counts = this.countInspectionsByCartilla(fechaISO);
    const loaded = this.getLoadedCartillas();
    const fechaLabel = formatISOToDMY(fechaISO);

    const countsHtml = loaded
      .map(
        (c) =>
          `<li><b>${htmlEscape(c)}</b>: ${counts[c] ?? 0} inspecciones</li>`
      )
      .join("");
    const totalDia = loaded.reduce((sum, c) => sum + (counts[c] ?? 0), 0);

    if (loaded.length <= 1) {
      this.exportFilteredCartilla(cartilla, fechaISO);
      return;
    }

    this.exportFilteredCartilla(cartilla, fechaISO);
  }

  exportExcelCartillaCompleto(cartilla) {
    const fechaISO = this.shell.refs.inspectionSelect?.value || "";
    this.exportExcelConErroresResaltados(cartilla, fechaISO);
  }

  bindEvents() {
    this.abortController?.abort();
    this.abortController = new AbortController();
    const { signal } = this.abortController;
    const refs = this.shell.refs;

    refs.clearBtn?.addEventListener("click", () => this.onClear(), { signal });
    refs.fileInput?.addEventListener("change", (event) => this.onFileSelected(event), { signal });
    refs.inspectionTypeSelect?.addEventListener("change", () => this.onCartillaChange(), { signal });
    refs.inspectionSelect?.addEventListener("change", () => this.onInspectionDateChange(), { signal });
    refs.runReviewBtn?.addEventListener("click", () => this.onRunReview(), { signal });
    refs.reviewAllBtn?.addEventListener("click", () => this.onReviewAll(), { signal });
    refs.exportExcelErroresBtn?.addEventListener("click", () => this.onExportErrors(), { signal });
    refs.exportBtn?.addEventListener("click", () => this.onExportFiltered(), { signal });
    refs.notificationIcon?.addEventListener("click", () => this.onNotificationClick(), { signal });

    this.bindResultsColumnMenu();
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
      protectedColIndices: new Set(STICKY_COLUMNS),
      onVisibilityChange: () => syncEsparragoMpStickyOffsets(refs.resultsTable)
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
    this.headersByCartilla = {};
    this.columnsByCartilla = {};
    this.excelCabeceraByCartilla = {};
    this.cartillaStatus = { MPES: false };
    this.notificationErrors = [];
    this.processedRows = [];
    this.lastErrorMap = new Map();
    this.duplicateLotes = new Set();
    this.excelLoaded = false;
    this.lastReviewKey = "";
    this.lastReviewAllKey = "";
    this._sapLayoutNotice = null;
    this.cartillaAnalysis?.clear();
  }

  async onFileSelected(event) {
    const files = Array.from(event.target.files || []);
    const refs = this.shell.refs;

    if (!files.length) return;

    if (!ensureXlsxLibrary()) {
      if (refs.fileInput) refs.fileInput.value = "";
      return;
    }

    const maxArchivos = 1;

    if (files.length > maxArchivos) {
      showMpDialog({
        icon: "error",
        title: "Demasiados archivos",
        html: `Solo se permite <b>1 archivo</b> MPES.`
      });
      if (refs.fileInput) refs.fileInput.value = "";
      return;
    }

    this.resetData();
    this.shell.resetDashboard({ preserveFileInput: true });

    const permitidas = this.getCodigosPermitidos();
    const cartillasCargadas = new Set();

    try {
      for (const file of files) {
        const buffer = await file.arrayBuffer();
        const wb = window.XLSX.read(new Uint8Array(buffer), { type: "array" });
        const data = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
          header: 1,
          raw: false
        });

        const fila4 = data[3] || [];
        const estado = String(fila4[13] ?? "")
          .toUpperCase()
          .trim();
        const cartillaRaw = String(fila4[8] ?? "")
          .toUpperCase()
          .trim();
        const cartilla = permitidas[cartillaRaw];

        if (!cartilla) {
          showMpDialog({
            icon: "error",
            title: "Cartilla no válida",
            html: `La cartilla <b>${htmlEscape(cartillaRaw || "DESCONOCIDA")}</b> no está permitida.<br>
              Solo se acepta cartilla <b>MPES</b>.`
          });
          if (refs.fileInput) refs.fileInput.value = "";
          return;
        }

        if (cartillasCargadas.has(cartilla)) {
          showMpDialog({
            icon: "error",
            title: "Cartilla duplicada",
            html: `La cartilla <b>${htmlEscape(cartilla)}</b> ya fue cargada.<br>
              No se permiten cartillas repetidas.`
          });
          if (refs.fileInput) refs.fileInput.value = "";
          return;
        }

        if (estado !== "ENVIADA") {
          showMpDialog({
            icon: "error",
            title: "Estado incorrecto",
            html: `La cartilla <b>${htmlEscape(cartilla)}</b> debe estar en estado <b>ENVIADA</b>.`
          });
          if (refs.fileInput) refs.fileInput.value = "";
          return;
        }

        cartillasCargadas.add(cartilla);
        this.excelCabeceraByCartilla[cartilla] = this.parseExcelCabecera(data);

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

        // Si Nota Condición no está en col 28 → faltan SAP: insertar 15 + 5 vacías.
        const {
          headers,
          rows: layoutRows,
          expanded: sapLayoutExpanded,
          insertedSap15,
          insertedSap5
        } = expandMissingSapLayout(rawHeaders, rawDataRows, getSapPerfil("mp"));

        if (headers.length !== this.totalColumnasFor(cartilla)) {
          showMpDialog({
            icon: "error",
            title: "Estructura incorrecta",
            html: `El archivo de <b>${htmlEscape(cartilla)}</b> tiene <b>${headers.length}</b> columnas${
              sapLayoutExpanded
                ? ` (tras completar huecos SAP: +${insertedSap15 + insertedSap5})`
                : ""
            }.<br>
              Se requieren <b>${this.totalColumnasFor(cartilla)} columnas</b>.`
          });
          if (refs.fileInput) refs.fileInput.value = "";
          return;
        }

        this.headersByCartilla[cartilla] = headers;
        this.columnsByCartilla[cartilla] = headers.map((h, i) => ({
          id: `col_${i + 1}`,
          header: h || "",
          originalIndex: i
        }));

        const colFechaJs = this.colFechaInspeccionJsFor(cartilla);
        const dateColsExcel =
          this.getExportConfig(cartilla)?.["columnas-fecha"] || [20, 47, 48, 57];
        const filas = applyDateDisplayFormatToRows(layoutRows, headers, dateColsExcel).map((row) => {
          const copy = Array.isArray(row) ? row : [...row];
          copy._fechaInspeccionISO = parseExcelDateISO(copy[colFechaJs]);
          if (sapLayoutExpanded) copy._sapLayoutExpanded = true;
          return copy;
        });
        if (filas.length) {
          this.cartillaStatus[cartilla] = true;
          this.rawRows.push(...filas);
        }
        if (sapLayoutExpanded) {
          this._sapLayoutNotice = {
            cartilla,
            insertedSap15,
            insertedSap5
          };
        }
      }

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
      if (refs.fileInput) {
        refs.fileInput.title = files.map((file) => file.name).join(", ");
      }
      this.excelLoaded = true;
      this.shell.setLiveStatus(true);

      this.fillCartillaSelect();
      this.setNotification(this.detectMissingInspectionDates());
      this.renderExcelInsight();

      const cartillas = [...cartillasCargadas];
      const primeraCartilla = cartillas[0];
      const sapNote = this._sapLayoutNotice
        ? `<br><small>Se alineó el bloque SAP (+${this._sapLayoutNotice.insertedSap15} + ${this._sapLayoutNotice.insertedSap5}). Las columnas SAP vacías sí se validan como obligatorias.</small>`
        : "";
      showMpDialog({
        icon: "success",
        title: "Excel cargado",
        html: `Cartilla(s) <b>${htmlEscape(cartillas.join(", "))}</b> · <b>${this.rawRows.length}</b> registros · <b>${primeraCartilla ? this.totalColumnasFor(primeraCartilla) : 0}</b> columnas${sapNote}`,
        timer: sapNote ? 3200 : 1800,
        showConfirmButton: false
      });

      if (CARTILLA_ORDER.length > 1) await this.showCartillaSummary();

      this.syncActionButtons();
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

  fillCartillaSelect() {
    const select = this.shell.refs.inspectionTypeSelect;
    if (!select) {
      this.onCartillaChange();
      return;
    }

    select.innerHTML = `<option value="" disabled selected>${htmlEscape(t("arandanoMp.selectCartilla"))}</option>`;
    CARTILLA_ORDER.forEach((cartilla) => {
      if (!this.cartillaStatus[cartilla]) return;
      const opt = document.createElement("option");
      opt.value = cartilla;
      opt.textContent = cartilla;
      select.appendChild(opt);
    });
    select.disabled = false;

    const disponibles = CARTILLA_ORDER.filter((cartilla) => this.cartillaStatus[cartilla]);
    if (disponibles.length === 1) {
      select.value = disponibles[0];
      this.onCartillaChange();
      return;
    }

    this.syncActionButtons();
  }

  detectMissingInspectionDates() {
    const errors = [];
    const cartilla = "MPES";
    const regla = (this.getReglas(cartilla)?.["validaciones-carga"] || []).find(
      (v) => v.tipo === "aviso-fecha-inspeccion-faltante"
    );
    const colJs =
      (regla?.columna || this.getCfgCartilla(cartilla)["columna-fecha-inspeccion"] || 41) - 1;
    const colsMostrar = regla?.["columnas-mostrar"] || [1, 10];
    const idJs = (colsMostrar[0] || 1) - 1;
    const loteJs = (colsMostrar[1] || 10) - 1;

    this.rawRows.forEach((r) => {
      const iso = r._fechaInspeccionISO || parseExcelDateISO(r[colJs]);
      if (!iso) {
        errors.push({
          id: r[idJs] || "",
          lote: r[loteJs] || ""
        });
      }
    });

    return errors;
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
      title: "Falta fecha de inspección",
      html: `<div class="agv-mp-dialog__html--scroll">
        ${this.notificationErrors
          .map(
            (e) =>
              `• <b>ID:</b> ${htmlEscape(e.id)} &nbsp; <b>Lote:</b> ${htmlEscape(e.lote)}`
          )
          .join("<br>")}
      </div>`,
      wide: true
    });
  }

  async showCartillaSummary() {
    await showMpDialog({
      icon: "info",
      title: "Resumen de cartillas",
      html: buildCartillaSummaryHtml(this.cartillaStatus),
      wide: true,
      confirmButtonText: "Aceptar"
    });
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
                <linearGradient id="pmparInsightRingGrad" x1="0%" y1="0%" x2="100%" y2="100%">
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

  onCartillaChange() {
    const select = this.shell.refs.inspectionSelect;
    if (!select) return;

    const fechas = [
      ...new Set(
        this.rawRows
          .map((r) => r._fechaInspeccionISO || parseExcelDateISO(r[this.colFechaInspeccionJs]))
          .filter(Boolean)
      )
    ].sort();

    select.innerHTML = `<option value="" disabled selected>${htmlEscape(t("plagasArandano.selectDate"))}</option>`;
    fechas.forEach((iso) => {
      const opt = document.createElement("option");
      opt.value = iso;
      opt.textContent = formatISOToDMY(iso);
      select.appendChild(opt);
    });
    select.disabled = !fechas.length;

    if (this.shell.refs.lmrSelect) {
      this.shell.refs.lmrSelect.innerHTML =
        `<option value="" selected>${htmlEscape(t("arandanoMp.lmrAutoDate"))}</option>`;
      this.shell.refs.lmrSelect.disabled = true;
      this.shell.refs.lmrSelect.classList.remove(`${this.shell.cls("input")}--warning`);
    }

    this.renderExcelInsight();
    this.syncActionButtons();
  }

  onInspectionDateChange() {
    const cartilla = this.getActiveCartilla();
    const fechaISO = this.shell.refs.inspectionSelect?.value;
    const lmrSelect = this.shell.refs.lmrSelect;
    if (!cartilla || !fechaISO || !lmrSelect) return;

    const rows = this.rawRows.filter(
      (r) => (r._fechaInspeccionISO || parseExcelDateISO(r[this.colFechaInspeccionJs])) === fechaISO
    );

    const lmrDates = rows.map((r) => parseExcelDateISO(r[this.colFechaLmrJs])).filter(Boolean);
    const unique = [...new Set(lmrDates)];
    const fechaMayoritaria = computeFechaLmrMayoritaria(rows, this.colFechaLmrJs);

    lmrSelect.innerHTML = "";
    if (fechaMayoritaria) {
      const opt = document.createElement("option");
      opt.value = fechaMayoritaria;
      opt.textContent = formatISOToDMY(fechaMayoritaria);
      lmrSelect.appendChild(opt);
      lmrSelect.value = fechaMayoritaria;
    } else {
      lmrSelect.innerHTML = `<option value="" selected>${htmlEscape(t("arandanoMp.lmrAutoDate"))}</option>`;
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
        title: "Múltiples fechas LMR detectadas",
        html: `<div class="agv-mp-dialog__html--body">
          Se detectaron <b>${unique.length}</b> fechas LMR diferentes:<br><br>
          ${detalles}<br><br>
          Se usará la fecha mayoritaria. Las filas con fechas minoritarias se marcarán como error.
        </div>`,
        wide: true
      });
    } else {
      lmrSelect.classList.remove(`${this.shell.cls("input")}--warning`);
    }

    this.syncActionButtons();
  }

  buildReviewAllItems(cartilla) {
    const cartillaRows = this.getCartillaRows(cartilla);
    const colFechaInspeccionJs = this.colFechaInspeccionJsFor(cartilla);
    const colFechaLmrJs = this.colFechaLmrJsFor(cartilla);
    const colLoteJs = this.colLoteJsFor(cartilla);
    const reglas = this.getReglas(cartilla);
    const compuestaMap = this.compuestaMapFor(cartilla);
    const items = [];
    const byDate = groupRowsByInspectionDate(cartillaRows, colFechaInspeccionJs);
    const fechas = [...byDate.keys()].sort();

    fechas.forEach((fechaISO) => {
      const rows = byDate.get(fechaISO) || [];
      const fechaLmrMayoritaria = computeFechaLmrMayoritaria(rows, colFechaLmrJs);
      const registros = rows.map((row) => rowToRegistro(row, row._filaNum));
      const resultado = analizarReporte(
        { filas: registros, cultivo: "esparrago" },
        reglas,
        { cartilla, fechaLmrMayoritaria }
      );
      const errorMap = buildErrorMap(resultado, compuestaMap);
      const dupLotes = detectDuplicateLotes(rows, colLoteJs);
      const filasDetalle = rows.filter((row) => {
        const filaMap = errorMap.get(row._filaNum);
        if (filaMap?.size) return true;
        const lote = String(row[colLoteJs] ?? "").trim();
        return lote && dupLotes.has(lote);
      });

      items.push({
        cartilla,
        fecha: formatISOToDMY(fechaISO),
        fechaISO,
        totalFilas: rows.length,
        filasConError: filasDetalle.length,
        filasDetalle,
        lotesDuplicados: [...dupLotes],
        errorMap,
        duplicateLotes: dupLotes,
        tieneErrores: filasDetalle.length > 0
      });
    });

    return items;
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
    if (refs.resultsSection) {
      refs.resultsSection.classList.remove("is-visible", `${this.shell.cls("results")}--ok`, `${this.shell.cls("results")}--errors`);
    }
    if (refs.resultsSubtitleEl) refs.resultsSubtitleEl.textContent = "";
    if (refs.totalFilasDiv) refs.totalFilasDiv.textContent = "";
    this.cartillaAnalysis?.clear();
    this.processedRows = [];
    this.lastReviewKey = "";
    this.lastReviewAllKey = "";
    this.syncActionButtons();
  }

  rowHasErrorWithContext(row, errorMap, duplicateLotes) {
    const filaMap = errorMap.get(row._filaNum);
    if (filaMap && filaMap.size > 0) return true;
    const lote = String(row[this.colLoteJs] ?? "").trim();
    return Boolean(lote && duplicateLotes.has(lote));
  }

  getCellMetaWithContext(row, colJs, errorMap, duplicateLotes) {
    const colNum = colJs + 1;
    const val = valorCeldaParaMostrar(row[colJs]);
    const err = errorMap.get(row._filaNum)?.get(colNum);
    const lote = String(row[this.colLoteJs] ?? "").trim();

    if (err) {
      if (err.tipo === "obligatorio") {
        return { val, cellClass: "agv-mp-cell-error-empty", title: err.problema };
      }
      return { val, cellClass: "agv-mp-cell-error-value", title: err.problema };
    }

    if (colNum === this.colLoteJs + 1 && lote && duplicateLotes.has(lote)) {
      return { val, cellClass: "agv-mp-cell-error-value", title: t("plagasArandano.duplicateLots") };
    }

    return { val, cellClass: "", title: "" };
  }

  htmlTablaFilasConError(cartilla, filas, errorMap, duplicateLotes, options = {}) {
    const { titled = true } = options;
    if (!filas?.length) return "";

    const allColumns = this.columnsByCartilla[cartilla] || [];
    const columns = pickResultColumnsForCartilla(
      cartilla,
      allColumns,
      errorMap,
      filas,
      (c) => this.getReglas(c),
      this.headersByCartilla
    );
    const thead =
      columns
        .map((col) => {
          const sticky = isPinnedColumn(col.originalIndex)
            ? ` agv-mp-sticky-col agv-mp-sticky-col-${col.originalIndex}`
            : "";
          return `<th class="agv-mp-table__col-header${sticky}"${
            col.headerTitle ? ` title="${htmlEscape(col.headerTitle)}"` : ""
          }>${htmlEscape(col.header)}</th>`;
        })
        .join("") + htmlSumVerificationHeaders();

    const tbody = filas
      .map((row) => {
        const tds =
          columns
            .map((col) => {
              const { val, cellClass, title } = this.getCellMetaWithContext(
                row,
                col.originalIndex,
                errorMap,
                duplicateLotes
              );
              const sticky = isPinnedColumn(col.originalIndex)
                ? `agv-mp-sticky-col agv-mp-sticky-col-${col.originalIndex}`
                : "";
              const classes = [cellClass, sticky].filter(Boolean).join(" ");
              const classAttr = classes ? ` class="${htmlEscape(classes)}"` : "";
              const titleAttr = title ? ` title="${htmlEscape(title)}"` : "";
              return `<td${classAttr}${titleAttr}>${htmlEscape(val)}</td>`;
            })
            .join("") + htmlSumVerificationCells(row);
        return `<tr>${tds}</tr>`;
      })
      .join("");

    const titleBlock = titled
      ? `<p class="agv-mp-nested-table-title">${htmlEscape(t("plagasArandano.errorRowsTitle"))}</p>`
      : "";

    return `
      <div class="agv-mp-nested-table-wrap">
        ${titleBlock}
        <div class="agv-mp-table-scroll">
          <table class="agv-mp-table agv-mp-table--esparrago">
            <thead><tr>${thead}</tr></thead>
            <tbody>${tbody}</tbody>
          </table>
        </div>
      </div>`;
  }

  renderResumenTodasFechas(cartilla, items) {
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
        const estado = item.tieneErrores
          ? t("plagasArandano.statusWithIssues")
          : t("plagasArandano.statusOk");
        const rate = item.totalFilas
          ? Math.round((item.filasConError / item.totalFilas) * 100)
          : 0;
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

    const details = items
      .filter((item) => item.filasDetalle?.length)
      .map(
        (item, idx) => `
        <details class="agv-mp-date-detail" data-todo-detail="${idx}">
          <summary class="agv-mp-date-detail__head">
            <h4 class="agv-mp-date-detail__title">${htmlEscape(item.fecha)}</h4>
            <span class="agv-mp-date-detail__meta">${htmlEscape(
              t("plagasArandano.errorRowsCount", {
                errors: item.filasConError,
                total: item.totalFilas
              })
            )} · clic para ver</span>
          </summary>
          <div class="agv-mp-date-detail__body" data-todo-body="${idx}">
            <p class="agv-mp-date-detail__loading">Cargando tabla…</p>
          </div>
        </details>`
      )
      .join("");

    const detailsBlock = details
      ? `<div class="agv-mp-dashboard__details">
          <h3 class="agv-mp-dashboard__details-title">${htmlEscape(t("plagasArandano.errorsDetailHeading"))}</h3>
          ${details}
        </div>`
      : "";

    el.innerHTML = `
      <div class="agv-mp-dashboard">
        <div>
          <h3 class="agv-mp-dashboard__title">${htmlEscape(t("arandanoMp.reviewAllDialogTitle", { cartilla }))}</h3>
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

    const detailItems = items.filter((item) => item.filasDetalle?.length);
    el.querySelectorAll("details[data-todo-detail]").forEach((detailsEl) => {
      detailsEl.addEventListener("toggle", () => {
        if (!detailsEl.open || detailsEl.dataset.loaded === "1") return;
        const idx = Number(detailsEl.dataset.todoDetail);
        const item = detailItems[idx];
        const body = detailsEl.querySelector(`[data-todo-body="${idx}"]`);
        if (!item || !body) return;
        body.innerHTML = this.htmlTablaFilasConError(
          cartilla,
          item.filasDetalle,
          item.errorMap,
          item.duplicateLotes,
          { titled: false }
        );
        syncEsparragoMpStickyOffsets(body.querySelector(".agv-mp-table"));
        requestAnimationFrame(() => {
          syncEsparragoMpStickyOffsets(body.querySelector(".agv-mp-table"));
        });
        detailsEl.dataset.loaded = "1";
      });
    });
  }

  onReviewAll() {
    const cartilla = this.getSelectedCartillaOrWarn();
    if (!cartilla) return;

    const items = this.buildReviewAllItems(cartilla);

    if (!items.length) {
      showMpDialog({
        icon: "info",
        title: t("plagasArandano.noDates"),
        text: t("plagasArandano.noDatesText")
      });
      return;
    }

    this.hideSingleDateResults();
    this.renderResumenTodasFechas(cartilla, items);
    this.lastReviewAllKey = cartilla;
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
    const cartilla = this.getSelectedCartillaOrWarn();
    if (!cartilla) return;
    if (this.lastReviewAllKey !== cartilla) {
      showMpDialog({
        icon: "warning",
        title: "Revisión requerida",
        html: "Primero pulsa <b>Todo</b> para revisar todas las fechas; luego podrás descargar el Excel completo."
      });
      return;
    }
    this.exportExcelConErroresResaltados(cartilla, "");
  }

  async onExportFiltered() {
    if (this.shell.refs.exportBtn?.disabled) {
      showMpDialog({
        icon: "warning",
        title: "Revisión requerida",
        html: "Selecciona <b>cartilla</b> y <b>fecha de inspección</b>, luego ejecuta <b>Revisar Excel</b> antes de exportar."
      });
      return;
    }

    const cartilla = this.getActiveCartilla();
    const fechaISO = this.shell.refs.inspectionSelect?.value;

    if (!cartilla || !fechaISO) {
      showMpDialog({
        icon: "warning",
        title: "Datos incompletos",
        html: "Selecciona <b>cartilla</b> y <b>fecha de inspección</b>."
      });
      return;
    }

    await this.promptExportFilteredChoice(cartilla, fechaISO);
  }

  async onRunReview() {
    const cartilla = this.getActiveCartilla();
    const fechaISO = this.shell.refs.inspectionSelect?.value;

    if (!cartilla || !this.cartillaStatus[cartilla]) {
      showMpDialog({
        icon: "warning",
        title: "Sin cartilla",
        text: "Debes seleccionar una cartilla válida."
      });
      return;
    }

    if (!fechaISO) {
      showMpDialog({
        icon: "warning",
        title: "Falta fecha de inspección",
        html: "Debes seleccionar una <b>fecha de inspección</b> antes de ejecutar."
      });
      return;
    }

    const rows = this.rawRows
      .filter((r) => (r._fechaInspeccionISO || parseExcelDateISO(r[this.colFechaInspeccionJs])) === fechaISO)
      .map((row, idx) => {
        const copy = [...row];
        copy._filaNum = idx + 1;
        copy._fechaInspeccionISO = row._fechaInspeccionISO || fechaISO;
        return copy;
      });

    const fechaLmrMayoritaria = computeFechaLmrMayoritaria(rows, this.colFechaLmrJs);
    const lmrDisplay = this.shell.refs.lmrSelect?.value || fechaLmrMayoritaria;

    const confirm = await showMpConfirmDialog({
      icon: "info",
      title: "Revisión de fechas",
      html: `<div class="agv-mp-dialog__html--compact">
        Se va a revisar:<br><br>
        <b>Cartilla:</b> ${htmlEscape(cartilla)}<br><br>
        <b>Fecha inspección:</b> ${htmlEscape(formatISOToDMY(fechaISO))}<br><br>
        <b>Fecha LMR (mayoritaria):</b> ${htmlEscape(formatISOToDMY(lmrDisplay))}
      </div>`,
      confirmButtonText: "Continuar",
      cancelButtonText: "Cancelar",
      wide: true
    });

    if (!confirm?.isConfirmed) return;

    this.hideResumenTodasFechas();
    this.runValidationAndRender(rows, cartilla, fechaISO, fechaLmrMayoritaria);
  }

  runValidationAndRender(rows, cartilla, fechaISO, fechaLmrMayoritaria) {
    const registros = rows.map((row) => rowToRegistro(row, row._filaNum));
    const reglas = this.getReglas(cartilla);
    const resultado = analizarReporte(
      { filas: registros, cultivo: "esparrago" },
      reglas,
      { cartilla, fechaLmrMayoritaria }
    );

    this.processedRows = rows;
    this.lastErrorMap = buildErrorMap(resultado, this.compuestaMapFor(cartilla));
    this.duplicateLotes = detectDuplicateLotes(rows, this.colLoteJsFor(cartilla));

    const filasConErrorAll = rows.filter((row) => this.rowHasError(row));
    const colLoteJs = this.colLoteJsFor(cartilla);
    const filasConError = filterFilasConErrorExcludingSapOnly(filasConErrorAll, {
      errorMap: this.lastErrorMap || null,
      duplicateLotes: this.duplicateLotes || new Set(),
      colLoteJs,
      t: (k, v) => t(k, v),
      skipSapValidation: false
    });
    this.renderResultsTable(rows, filasConError, cartilla, fechaISO);
    this.lastReviewKey = `${cartilla}|${fechaISO}`;
    this.syncActionButtons();

    this.cartillaAnalysis?.present({
      rows,
      filasConError,
      errorMap: this.lastErrorMap || null,
      duplicateLotes: this.duplicateLotes || new Set(),
      colLoteJs,
      columns: this.columnsByCartilla?.[cartilla] || [],
      cartilla: cartilla || "—",
      fechaLabel: formatISOToDMY(fechaISO),
      translateHeader: translateExcelHeader
    });
  }

  rowHasError(row) {
    return this.rowHasErrorWithContext(row, this.lastErrorMap, this.duplicateLotes);
  }

  getCellMeta(row, colJs) {
    return this.getCellMetaWithContext(row, colJs, this.lastErrorMap, this.duplicateLotes);
  }

  renderResultsTable(allRows, filasConError, cartilla, fechaISO) {
    const refs = this.shell.refs;
    const headers = this.headersByCartilla[cartilla] || [];
    const allColumns = this.columnsByCartilla[cartilla] || [];
    const columns = pickResultColumnsForCartilla(
      cartilla,
      allColumns,
      this.lastErrorMap,
      filasConError,
      (c) => this.getReglas(c),
      this.headersByCartilla
    );

    if (refs.resultsHeader) refs.resultsHeader.innerHTML = "";
    if (refs.resultsBody) refs.resultsBody.innerHTML = "";

    const hasErrors = filasConError.length > 0;

    if (refs.resultsSection) {
      refs.resultsSection.classList.remove(`${this.shell.cls("results")}--ok`, `${this.shell.cls("results")}--errors`);
      refs.resultsSection.classList.add(hasErrors ? `${this.shell.cls("results")}--errors` : `${this.shell.cls("results")}--ok`);
      refs.resultsSection.classList.add("is-visible");
    }

    if (refs.resultsTitleEl) {
      refs.resultsTitleEl.textContent = hasErrors
        ? t("plagasArandano.errorRowsTitle")
        : t("plagasArandano.allCorrect");
    }

    if (refs.resultsSubtitleEl) {
      refs.resultsSubtitleEl.textContent = t("plagasArandano.resultsInspectionDate", {
        date: formatISOToDMY(fechaISO)
      });
    }

    if (refs.resultsIconEl) {
      refs.resultsIconEl.innerHTML = hasErrors
        ? '<i data-lucide="triangle-alert"></i>'
        : '<i data-lucide="circle-check"></i>';
    }

    if (refs.totalFilasDiv) {
      this._lastInspectionRowCount = allRows.length;
      refs.totalFilasDiv.textContent = t("esparragoMp.totalInspectionRows", { count: allRows.length });
    }

    if (!hasErrors) {
      const tr = document.createElement("tr");
      tr.className = "agv-mp-row-ok";
      const td = document.createElement("td");
      td.colSpan = Math.max(columns.length || headers.length, 1);
      td.textContent = t("esparragoMp.noInspectionErrors");
      tr.appendChild(td);
      refs.resultsBody?.appendChild(tr);
    } else {
      const headerFrag = document.createDocumentFragment();
      columns.forEach((col) => {
        const th = document.createElement("th");
        th.className = "agv-mp-table__col-header";
        th.dataset.colIndex = String(col.originalIndex);
        th.dataset.excelHeader = String(col.header ?? "");
        th.textContent = col.isSapZone
          ? col.header
          : translateExcelHeader(col.header, col.originalIndex);
        if (col.headerTitle) th.title = col.headerTitle;
        applyStickyColumnClasses(th, col.originalIndex);
        headerFrag.appendChild(th);
      });
      refs.resultsHeader?.appendChild(headerFrag);
      appendSumVerificationHeaders(refs.resultsHeader);

      const bodyFrag = document.createDocumentFragment();
      filasConError.forEach((row) => {
        const tr = document.createElement("tr");
        columns.forEach((col) => {
          const { val, cellClass, title } = this.getCellMeta(row, col.originalIndex);
          const td = document.createElement("td");
          td.dataset.colIndex = String(col.originalIndex);
          if (cellClass) td.className = cellClass;
          if (title) td.title = title;
          td.textContent = val;
          applyStickyColumnClasses(td, col.originalIndex);
          tr.appendChild(td);
        });
        appendSumVerificationCells(tr, row);
        bodyFrag.appendChild(tr);
      });
      refs.resultsBody?.appendChild(bodyFrag);
    }

    if (refs.resultsTable) refs.resultsTable.hidden = false;
    this.bindResultsColumnMenu();
    applyMpColumnVisibility(refs.resultsTable);
    requestAnimationFrame(() => {
      syncEsparragoMpStickyOffsets(refs.resultsTable);
    });
    if (refs.resultsIconEl) hydrateLucideIcons(refs.resultsIconEl);
  }

  onLanguageChange() {
    const refs = this.shell?.refs;
    if (!refs) return;
    if (refs.totalFilasDiv && this._lastInspectionRowCount != null) {
      refs.totalFilasDiv.textContent = t("esparragoMp.totalInspectionRows", {
        count: this._lastInspectionRowCount
      });
    }
    refreshTranslatedHeaderRow(refs.resultsHeader, (idx) => {
      const cartilla = refs.inspectionTypeSelect?.value;
      return (this.columnsByCartilla?.[cartilla] || []).find((col) => col.originalIndex === idx)?.header || "";
    });
    refs.resultsBody?.querySelectorAll("tr.agv-mp-row-ok td").forEach((td) => {
      td.textContent = t("esparragoMp.noInspectionErrors");
    });
    if (this.cartillaAnalysis?.getLast() && this.processedRows?.length) {
      const [cartilla, fechaISO] = String(this.lastReviewKey || "").split("|");
      if (cartilla && fechaISO) {
        const filasConError = this.processedRows.filter((row) => this.rowHasError(row));
        this.cartillaAnalysis.refreshPanel({
          rows: this.processedRows,
          filasConError,
          errorMap: this.lastErrorMap || null,
          duplicateLotes: this.duplicateLotes || new Set(),
          colLoteJs: this.colLoteJsFor(cartilla),
          columns: this.columnsByCartilla?.[cartilla] || [],
          cartilla,
          fechaLabel: formatISOToDMY(fechaISO),
          translateHeader: translateExcelHeader
        });
      }
    }
  }

  destroy() {
    this.abortController?.abort();
    this.abortController = null;
    this.root = null;
    this.shell = null;
  }
}
