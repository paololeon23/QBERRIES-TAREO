import { AGV_MP_SHELL_IDS } from "../shared/cartilla-shell.ids.js";
import { CartillaShellUi } from "../shared/cartilla-shell.ui.js";
import { cargarReglasDesdeRuta, analizarReporte } from "../../../../../engine/rule-engine.js";
import { hydrateLucideIcons } from "../../../utils/lucide-icon.util.js";
import { i18nService } from "../../../services/i18n.service.js";
import { appConfig } from "../../../config/app.config.js";
import { showMpDialog, showMpConfirmDialog, showMpExportChoiceDialog } from "./arandano-mp-dialog.js";
import {
  buildFilteredSheetData,
  buildFullSheetDataWithErrors
} from "./arandano-mp-export.js";
import {
  buildLazyDateDetailPlaceholders,
  bindLazyDateDetailTables,
  collectValidatedColumnIndexesJs,
  SAP_ZONE_FRONTEND_COLS_JS,
  resolveSapZoneHeader,
  SAP_ZONE_HEADER_LABELS_BY_JS
} from "../shared/mp-results-perf.util.js?v=2026072219";
import { expandMissingSapLayout } from "../shared/mp-sap-layout.util.js";
import { applyDateDisplayFormatToRows } from "../shared/excel-date-format.util.js";
import { normalizeMphaColumnLayout } from "./arandano-mp-mpha-layout.util.js";
import { loadSapColumnasCatalog, getSapPerfil } from "../../../config/sap-columnas.registry.js";
import { translateExcelHeader } from "../../../utils/excel-header-i18n.util.js";
import { createCartillaAnalysisController } from "../shared/cartilla-analysis.js";
import {
  applyMpColumnVisibility,
  bindMpColumnContextMenu,
  bindMpTableSearch
} from "../shared/mp-column-menu.util.js";

/** Excel 1,2,4,5,7,10 → JS (Fecha registro + Hora en orden natural). */
const STICKY_COLUMNS = [0, 1, 3, 4, 6, 9];
/** Contexto siempre visible en revisión (Excel 4–5 = JS 3–4). */
const CONTEXT_COLUMNS_JS = [0, 1, 3, 4, 6, 9, 10];
/** Encabezados cortos sticky (la columna es angosta). */
const STICKY_HEADER_SHORT_BY_JS = {
  0: "Id",
  1: "IC",
  3: "F. Reg.",
  4: "Hora",
  6: "Usuario",
  9: "Lote",
  10: "CM"
};
const STICKY_HEADER_TITLE_BY_JS = {
  0: "Id",
  1: "Inspección código",
  3: "Fecha registro",
  4: "Hora",
  6: "Usuario",
  9: "Lote",
  10: "Cant Muestra"
};
/** Productor…Peso Bruto (Excel 13–33), incluye Nota Condición para UI. */
const SAP_ZONE_COLS_JS = SAP_ZONE_FRONTEND_COLS_JS || [
  12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32
];
const FILAS_SKIP = 5;
/** Anchos fijos sticky Arándano MP (orden acumulado → left). */
const STICKY_COL_WIDTHS = {
  0: 95,
  1: 95,
  3: 110,
  4: 72,
  6: 230,
  9: 126
};

function isPinnedColumn(index) {
  return STICKY_COLUMNS.includes(index);
}

function formatResultColumnHeader(col) {
  const idx = Number(col?.originalIndex);
  if (STICKY_HEADER_SHORT_BY_JS[idx] != null) return STICKY_HEADER_SHORT_BY_JS[idx];
  if (col?.isSapZone && col.header) return col.header;
  if (SAP_ZONE_HEADER_LABELS_BY_JS[idx] != null) {
    return resolveSapZoneHeader(idx, col.header).label;
  }
  return translateExcelHeader(col?.header, idx);
}

function resultColumnHeaderTitle(col) {
  const idx = Number(col?.originalIndex);
  if (STICKY_HEADER_TITLE_BY_JS[idx]) return STICKY_HEADER_TITLE_BY_JS[idx];
  if (col?.headerTitle) return col.headerTitle;
  return formatResultColumnHeader(col);
}

function applyStickyColumnClasses(el, index) {
  if (!isPinnedColumn(index)) return;
  el.classList.add("agv-mp-sticky-col", `agv-mp-sticky-col-${index}`);
}

/** Recalcula left/width según columnas sticky realmente pintadas (evita huecos). */
function syncArandanoMpStickyOffsets(tableEl) {
  if (!tableEl) return;
  let left = 0;
  STICKY_COLUMNS.forEach((idx) => {
    const cells = tableEl.querySelectorAll(`.agv-mp-sticky-col-${idx}`);
    if (!cells.length) return;
    const width = STICKY_COL_WIDTHS[idx] ?? 90;
    cells.forEach((el) => {
      el.style.left = `${left}px`;
      el.style.width = `${width}px`;
      el.style.minWidth = `${width}px`;
      el.style.maxWidth = `${width}px`;
    });
    left += width;
  });
}
const CARTILLA_ORDER = ["MPHA", "MPBA", "MPGA"];
const REGLAS_POR_CARTILLA = {
  MPHA: "rules/modulos/arandano-mp-mpha.rules.json",
  MPBA: "rules/modulos/arandano-mp-mpbar.rules.json",
  MPGA: "rules/modulos/arandano-mp-mpgar.rules.json"
};

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
  if (/^\d{4}-\d{2}-\d{2}/.test(texto)) {
    return texto.slice(0, 10);
  }
  if (/^\d{2}[/-]\d{2}[/-]\d{4}$/.test(texto)) {
    const [d, m, y] = texto.split(/[/-]/);
    return `${y}-${m}-${d}`;
  }
  if (/^\d{2}[/-]\d{2}[/-]\d{2}$/.test(texto)) {
    const [d, m, y] = texto.split(/[/-]/);
    const fullY = Number(y) <= 50 ? `20${y}` : `19${y}`;
    return `${fullY}-${m}-${d}`;
  }
  const fecha = Date.parse(texto);
  return Number.isFinite(fecha) ? new Date(fecha).toISOString().slice(0, 10) : "";
}

/** Fecha visible/export: 21/06/2026 (con /). */
function formatISOToDMY(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

function fechaLabelParaArchivo(iso) {
  return formatISOToDMY(iso).replaceAll("/", "");
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
  row.forEach((val, idx) => {
    _cols[String(idx + 1)] = val ?? "";
  });
  return { fila: filaNum, _cols };
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
      // No pisar un error de valor (igualdad/rango/…) con «obligatorio».
      if (existing && existing.tipo !== "obligatorio" && tipo === "obligatorio") return;
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

/** Filas con error: primero lotes duplicados (agrupados), luego el resto. */
function sortErrorRowsByDuplicateLote(rows, colLoteJs, duplicateLotes) {
  if (!rows?.length) return [];
  if (!duplicateLotes?.size) return [...rows];

  const loteKey = (row) => String(row[colLoteJs] ?? "").trim();
  const seenLotes = new Set();
  const loteOrder = [];
  const rest = [];

  rows.forEach((row) => {
    const lote = loteKey(row);
    if (lote && duplicateLotes.has(lote)) {
      if (!seenLotes.has(lote)) {
        seenLotes.add(lote);
        loteOrder.push(lote);
      }
    } else {
      rest.push(row);
    }
  });

  const dups = [];
  loteOrder.forEach((lote) => {
    rows.filter((r) => loteKey(r) === lote).forEach((r) => dups.push(r));
  });

  return [...dups, ...rest];
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

export class ArandanoMpService {
  constructor() {
    this.shell = null;
    this.reglasByCartilla = {};
    this.reglas = null;
    this.cfgCartillas = null;
    this.rawRows = [];
    this.headersByCartilla = {};
    this.columnsByCartilla = {};
    this.excelCabeceraByCartilla = {};
    this.cartillaStatus = { MPHA: false, MPBA: false, MPGA: false };
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
    this.searchBound = false;
  }

  async init(appRoot) {
    this.root = appRoot;
    const [entradas] = await Promise.all([
      Promise.all(
        CARTILLA_ORDER.map(async (cartilla) => {
          const reglas = await cargarReglasDesdeRuta(
            `${REGLAS_POR_CARTILLA[cartilla]}?v=${appConfig.cacheBustingVersion}`
          );
          return [cartilla, reglas];
        })
      ),
      loadSapColumnasCatalog(appConfig.cacheBustingVersion)
    ]);
    this.reglasByCartilla = Object.fromEntries(entradas);
    this.reglas = this.reglasByCartilla.MPBA;
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
      i18nPrefix: "plagasArandano"
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
    return this.colCartillaJsFor(this.shell.refs.inspectionTypeSelect?.value || "MPBA");
  }

  get colFechaInspeccionJs() {
    return this.colFechaInspeccionJsFor(this.shell.refs.inspectionTypeSelect?.value || "MPBA");
  }

  get colFechaLmrJs() {
    return this.colFechaLmrJsFor(this.shell.refs.inspectionTypeSelect?.value || "MPBA");
  }

  get colLoteJs() {
    return this.colLoteJsFor(this.shell.refs.inspectionTypeSelect?.value || "MPBA");
  }

  totalColumnasFor(cartilla) {
    return this.getReglas(cartilla)?.["total-columnas"] || 104;
  }

  /**
   * Contexto + zona SAP (se valida) + columnas con reglas + columnas con error.
   */
  getValidatedColumnIndexesJs(cartilla) {
    return collectValidatedColumnIndexesJs(this.getReglas(cartilla), {
      contextColsJs: CONTEXT_COLUMNS_JS,
      stickyColsJs: STICKY_COLUMNS,
      includeSapZone: true
    });
  }

  pickColumnsForResults(cartilla, allColumns, errorMap, filas) {
    const needed = new Set([
      ...CONTEXT_COLUMNS_JS,
      ...STICKY_COLUMNS,
      ...SAP_ZONE_COLS_JS,
      ...this.getValidatedColumnIndexesJs(cartilla)
    ]);
    errorMap?.forEach((filaMap) => {
      filaMap?.forEach((_err, colNum) => {
        const js = Number(colNum) - 1;
        if (Number.isFinite(js) && js >= 0) needed.add(js);
      });
    });

    const headers = this.headersByCartilla[cartilla] || [];
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

  getCodigosPermitidos() {
    const map = {};
    CARTILLA_ORDER.forEach((cartilla) => {
      const reglas = this.getReglas(cartilla);
      const codigo =
        reglas?.["codigo-archivo-cabecera"] ||
        reglas?.["configuracion-cartillas"]?.["codigo-archivo"];
      if (codigo) map[String(codigo).toUpperCase()] = cartilla;
    });
    return map;
  }

  compuestaMapFor(cartilla) {
    return this.compuestaColumnMapByCartilla[cartilla] || new Map();
  }

  get totalColumnas() {
    return this.totalColumnasFor("MPBA");
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
    const selected = this.shell.refs.inspectionTypeSelect?.value;
    if (selected && this.excelCabeceraByCartilla[selected]) return selected;
    return CARTILLA_ORDER.find((cartilla) => this.excelCabeceraByCartilla[cartilla]);
  }

  syncActionButtons() {
    const refs = this.shell?.refs;
    if (!refs) return;

    const cartilla = refs.inspectionTypeSelect?.value || "";
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
    const cartilla = this.shell.refs.inspectionTypeSelect?.value || "";
    if (!this.excelLoaded || !this.rawRows.length) {
      showMpDialog({
        icon: "warning",
        title: t("plagasArandano.attention"),
        text: t("arandanoMp.noFile")
      });
      return "";
    }
    if (!cartilla || !this.cartillaStatus[cartilla]) {
      showMpDialog({
        icon: "warning",
        title: t("plagasArandano.attention"),
        text: t("arandanoMp.noCartilla")
      });
      return "";
    }
    return cartilla;
  }

  getCartillaRows(cartilla) {
    const colCartillaJs = this.colCartillaJsFor(cartilla);
    return this.rawRows
      .filter((r) => String(r[colCartillaJs] ?? "").toUpperCase() === cartilla)
      .map((row, idx) => {
        const copy = [...row];
        copy._filaNum = idx + 1;
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

    const fechas = [
      ...new Set(rows.map((r) => parseExcelDateISO(r[colFechaInspeccionJs])).filter(Boolean))
    ];

    fechas.forEach((fechaISO) => {
      const fechaRows = rows.filter(
        (r) => parseExcelDateISO(r[colFechaInspeccionJs]) === fechaISO
      );
      const fechaLmrMayoritaria = computeFechaLmrMayoritaria(fechaRows, colFechaLmrJs);
      const registros = fechaRows.map((row) => rowToRegistro(row, row._filaNum));
      const resultado = analizarReporte(
        { filas: registros, cultivo: "arandano" },
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
    const colCartillaJs = this.colCartillaJsFor(cartilla);
    const colFechaJs = this.colFechaInspeccionJsFor(cartilla);
    return this.rawRows
      .filter(
        (r) =>
          String(r[colCartillaJs] ?? "").toUpperCase() === cartilla &&
          parseExcelDateISO(r[colFechaJs]) === fechaISO
      )
      .map((row, idx) => {
        const copy = [...row];
        copy._filaNum = idx + 1;
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
      { filas: registros, cultivo: "arandano" },
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
    const fechaLabel = fechaLabelParaArchivo(fechaISO);
    const nombre = `ARANDANOS_${cartilla}_Filtrado_${fechaLabel}.xlsx`;

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

    const fechaLabel = fechaLabelParaArchivo(fechaISO);
    const nombre = `ARANDANOS_MP_Filtrado_${fechaLabel}.xlsx`;
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

    const fechaSuffix = fechaISO ? `_${fechaLabelParaArchivo(fechaISO)}` : "_TodasFechas";
    const nombre = `ARANDANOS_${cartilla}_Errores${fechaSuffix}.xlsx`;
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
          `<li class="agv-mp-export-choice__item">
            <span class="agv-mp-export-choice__code">${htmlEscape(c)}</span>
            <span class="agv-mp-export-choice__count">${counts[c] ?? 0} inspecciones</span>
          </li>`
      )
      .join("");
    const totalDia = loaded.reduce((sum, c) => sum + (counts[c] ?? 0), 0);

    if (loaded.length <= 1) {
      this.exportFilteredCartilla(cartilla, fechaISO);
      return;
    }

    const result = await showMpExportChoiceDialog({
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
    if (!this.searchBound) {
      bindMpTableSearch(refs.tableSearch, refs.resultsBody, {
        idColJs: 0,
        loteColJs: 9
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
      protectedColIndices: new Set(STICKY_COLUMNS),
      onVisibilityChange: () => syncArandanoMpStickyOffsets(refs.resultsTable)
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
    this.cartillaStatus = { MPHA: false, MPBA: false, MPGA: false };
    this.notificationErrors = [];
    this.processedRows = [];
    this.lastErrorMap = new Map();
    this.duplicateLotes = new Set();
    this.excelLoaded = false;
    this.lastReviewKey = "";
    this.lastReviewAllKey = "";
    this._sapLayoutNotice = null;
    this._mphaLayoutNotice = null;
    this._sapInsertedJsByCartilla = {};
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

    const maxArchivos =
      this.reglas?.["validaciones-archivo"]?.find((v) => v.tipo === "max-archivos-cartilla")?.valor ?? 3;

    if (files.length > maxArchivos) {
      showMpDialog({
        icon: "error",
        title: "Demasiados archivos",
        html: `Solo se permiten <b>${maxArchivos} cartillas</b> como máximo (MPHA, MPBA y MPGA).`
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
              Solo se aceptan: <b>MPHA (MPHPAR), MPBA (MPBAR) y MPGA (MPGAR)</b>.`
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

        // Si Nota Condición no está en col 28 → faltan SAP: insertar 15 + 5 vacías (Cond. Transporte → 34).
        let {
          headers,
          rows: layoutRows,
          expanded: sapLayoutExpanded,
          insertedSap15,
          insertedSap5,
          insertedJsIndexes
        } = expandMissingSapLayout(rawHeaders, rawDataRows, getSapPerfil("mp"));

        // MPHPAR: normalizar Pudrición con Larva / sumatorias al orden canónico (si ya viene así, no toca).
        let mphaLayoutNormalized = false;
        if (cartilla === "MPHA") {
          const mpha = normalizeMphaColumnLayout(headers, layoutRows);
          headers = mpha.headers;
          layoutRows = mpha.rows;
          mphaLayoutNormalized = mpha.normalized;
          if (mphaLayoutNormalized) {
            this._mphaLayoutNotice = { cartilla, steps: mpha.steps };
          }
        }

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
        this._sapInsertedJsByCartilla = this._sapInsertedJsByCartilla || {};
        this._sapInsertedJsByCartilla[cartilla] = insertedJsIndexes || [];
        this.columnsByCartilla[cartilla] = headers.map((h, i) => ({
          id: `col_${i + 1}`,
          header: h || "",
          originalIndex: i
        }));

        const colFechaJs = this.colFechaInspeccionJsFor(cartilla);
        const dateColsExcel = [
          ...new Set([
            ...(this.getExportConfig(cartilla)?.["columnas-fecha"] || [20, 41, 51]),
            this.getCfgCartilla(cartilla)["columna-fecha-cosecha"] || 20,
            this.getCfgCartilla(cartilla)["columna-fecha-inspeccion"] || 41,
            this.getCfgCartilla(cartilla)["columna-fecha-lmr"] || 51,
            this.getCfgCartilla(cartilla)["columna-fecha-produccion"] || 21
          ])
        ];
        const filas = applyDateDisplayFormatToRows(layoutRows, headers, dateColsExcel).map(
          (row) => {
            const copy = Array.isArray(row) ? row : [...row];
            copy._fechaInspeccionISO = parseExcelDateISO(copy[colFechaJs]);
            if (sapLayoutExpanded) copy._sapLayoutExpanded = true;
            return copy;
          }
        );
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
      const mphaNote = this._mphaLayoutNotice
        ? `<br><small>MPHPAR: se reordenó el bloque de defectos (Pudrición con Larva / sumatorias) al orden canónico.</small>`
        : "";
      showMpDialog({
        icon: "success",
        title: "Excel cargado",
        html: `Cartilla(s) <b>${htmlEscape(cartillas.join(", "))}</b> · <b>${this.rawRows.length}</b> registros · <b>${primeraCartilla ? this.totalColumnasFor(primeraCartilla) : 0}</b> columnas${sapNote}${mphaNote}`,
        timer: sapNote || mphaNote ? 3200 : 1800,
        showConfirmButton: false
      });

      await this.showCartillaSummary();

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
    if (!select) return;

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

    this.rawRows.forEach((r) => {
      const cartilla = String(r[this.colCartillaJsFor("MPBA")] ?? r[1] ?? "")
        .toUpperCase()
        .trim();
      if (!CARTILLA_ORDER.includes(cartilla)) return;

      const regla = (this.getReglas(cartilla)?.["validaciones-carga"] || []).find(
        (v) => v.tipo === "aviso-fecha-inspeccion-faltante"
      );
      const colJs =
        (regla?.columna || this.getCfgCartilla(cartilla)["columna-fecha-inspeccion"] || 41) - 1;
      const colsMostrar = regla?.["columnas-mostrar"] || [1, 10];

      if (!parseExcelDateISO(r[colJs])) {
        errors.push({
          id: r[(colsMostrar[0] || 1) - 1] || "",
          lote: r[(colsMostrar[1] || 10) - 1] || ""
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
    const cartilla = this.shell.refs.inspectionTypeSelect?.value;
    const select = this.shell.refs.inspectionSelect;
    if (!cartilla || !select) return;

    const fechas = [
      ...new Set(
        this.rawRows
          .filter((r) => String(r[this.colCartillaJs] ?? "").toUpperCase() === cartilla)
          .map((r) => parseExcelDateISO(r[this.colFechaInspeccionJs]))
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
    const cartilla = this.shell.refs.inspectionTypeSelect?.value;
    const fechaISO = this.shell.refs.inspectionSelect?.value;
    const lmrSelect = this.shell.refs.lmrSelect;
    if (!cartilla || !fechaISO || !lmrSelect) return;

    const rows = this.rawRows.filter(
      (r) =>
        String(r[this.colCartillaJs] ?? "").toUpperCase() === cartilla &&
        parseExcelDateISO(r[this.colFechaInspeccionJs]) === fechaISO
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
    const fechas = [
      ...new Set(
        cartillaRows.map((r) => parseExcelDateISO(r[colFechaInspeccionJs])).filter(Boolean)
      )
    ].sort();

    fechas.forEach((fechaISO) => {
      const rows = cartillaRows.filter(
        (r) => parseExcelDateISO(r[colFechaInspeccionJs]) === fechaISO
      );
      const fechaLmrMayoritaria = computeFechaLmrMayoritaria(rows, colFechaLmrJs);
      const registros = rows.map((row) => rowToRegistro(row, row._filaNum));
      const resultado = analizarReporte(
        { filas: registros, cultivo: "arandano" },
        reglas,
        { cartilla, fechaLmrMayoritaria }
      );
      const errorMap = buildErrorMap(resultado, compuestaMap);
      const dupLotes = detectDuplicateLotes(rows, colLoteJs);
      const filasDetalle = sortErrorRowsByDuplicateLote(
        rows.filter((row) => {
          const filaMap = errorMap.get(row._filaNum);
          if (filaMap?.size) return true;
          const lote = String(row[colLoteJs] ?? "").trim();
          return lote && dupLotes.has(lote);
        }),
        colLoteJs,
        dupLotes
      );

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
    if (refs.tableSearchWrap) refs.tableSearchWrap.hidden = true;
    if (refs.tableSearch) refs.tableSearch.value = "";
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
    const isLoteCol = colNum === this.colLoteJs + 1;

    // Duplicado: título claro (aunque el mapa traiga si-falla-mostrar genérico)
    if (isLoteCol && lote && duplicateLotes?.has?.(lote) && err?.tipo !== "longitud") {
      return {
        val,
        cellClass: "agv-mp-cell-error-value",
        title: `Lote ${lote} duplicado en esta fecha`
      };
    }

    if (err) {
      if (err.tipo === "obligatorio") {
        return { val, cellClass: "agv-mp-cell-error-empty", title: err.problema };
      }
      return { val, cellClass: "agv-mp-cell-error-value", title: err.problema };
    }

    if (isLoteCol && lote && duplicateLotes?.has?.(lote)) {
      return {
        val,
        cellClass: "agv-mp-cell-error-value",
        title: `Lote ${lote} duplicado en esta fecha`
      };
    }

    return { val, cellClass: "", title: "" };
  }

  htmlTablaFilasConError(cartilla, filas, errorMap, duplicateLotes, options = {}) {
    const { titled = true } = options;
    if (!filas?.length) return "";

    const allColumns = this.columnsByCartilla[cartilla] || [];
    const columns = this.pickColumnsForResults(cartilla, allColumns, errorMap, filas);
    const thead = columns
      .map((col) => {
        const sticky = isPinnedColumn(col.originalIndex)
          ? ` agv-mp-sticky-col agv-mp-sticky-col-${col.originalIndex}`
          : "";
        return `<th class="agv-mp-table__col-header${sticky}" title="${htmlEscape(
          resultColumnHeaderTitle(col)
        )}">${htmlEscape(formatResultColumnHeader(col))}</th>`;
      })
      .join("");

    const tbody = filas
      .map((row) => {
        const tds = columns
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
          .join("");
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
          <table class="agv-mp-table">
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

    const details = buildLazyDateDetailPlaceholders(items, htmlEscape, (item) =>
      t("plagasArandano.errorRowsCount", {
        errors: item.filasConError,
        total: item.totalFilas
      })
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
    bindLazyDateDetailTables(el, items, (item) => {
      const html = this.htmlTablaFilasConError(
        cartilla,
        item.filasDetalle,
        item.errorMap,
        item.duplicateLotes,
        { titled: false }
      );
      queueMicrotask(() => {
        el.querySelectorAll(".agv-mp-table-scroll .agv-mp-table").forEach((tableEl) => {
          syncArandanoMpStickyOffsets(tableEl);
        });
      });
      return html;
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
    // Botón de «Todo el Excel»: siempre todas las fechas de la cartilla.
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

    const cartilla = this.shell.refs.inspectionTypeSelect?.value;
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
    const cartilla = this.shell.refs.inspectionTypeSelect?.value;
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
      .filter(
        (r) =>
          String(r[this.colCartillaJs] ?? "").toUpperCase() === cartilla &&
          parseExcelDateISO(r[this.colFechaInspeccionJs]) === fechaISO
      )
      .map((row, idx) => {
        const copy = [...row];
        copy._filaNum = idx + 1;
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
      { filas: registros, cultivo: "arandano" },
      reglas,
      { cartilla, fechaLmrMayoritaria }
    );

    this.processedRows = rows;
    this.lastErrorMap = buildErrorMap(resultado, this.compuestaMapFor(cartilla));
    this.duplicateLotes = detectDuplicateLotes(rows, this.colLoteJsFor(cartilla));

    const colLoteJs = this.colLoteJsFor(cartilla);
    const filasConError = sortErrorRowsByDuplicateLote(
      rows.filter((row) => this.rowHasError(row)),
      colLoteJs,
      this.duplicateLotes
    );
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
    const columns = this.pickColumnsForResults(cartilla, allColumns, this.lastErrorMap, filasConError);

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
      refs.totalFilasDiv.textContent = t("arandanoMp.totalInspectionRows", {
        count: allRows.length
      });
    }

    if (!hasErrors) {
      const tr = document.createElement("tr");
      tr.className = "agv-mp-row-ok";
      const td = document.createElement("td");
      td.colSpan = Math.max(columns.length || headers.length, 1);
      td.textContent = t("arandanoMp.noInspectionErrors");
      tr.appendChild(td);
      refs.resultsBody?.appendChild(tr);
    } else {
      const headerFrag = document.createDocumentFragment();
      columns.forEach((col) => {
        const th = document.createElement("th");
        th.className = "agv-mp-table__col-header";
        th.dataset.colIndex = String(col.originalIndex);
        th.dataset.excelHeader = String(col.header ?? "");
        th.textContent = formatResultColumnHeader(col);
        th.title = `${resultColumnHeaderTitle(col)} — ${t("plagasArandano.hideColumnHint")}`;
        applyStickyColumnClasses(th, col.originalIndex);
        headerFrag.appendChild(th);
      });
      refs.resultsHeader?.appendChild(headerFrag);

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
        bodyFrag.appendChild(tr);
      });
      refs.resultsBody?.appendChild(bodyFrag);
    }

    if (refs.resultsTable) refs.resultsTable.hidden = false;
    if (refs.tableSearchWrap) refs.tableSearchWrap.hidden = false;
    this.bindResultsColumnMenu();
    applyMpColumnVisibility(refs.resultsTable);
    syncArandanoMpStickyOffsets(refs.resultsTable);
    if (refs.tableSearch?.value) {
      refs.tableSearch.dispatchEvent(new Event("input", { bubbles: true }));
    }
    if (refs.resultsIconEl) hydrateLucideIcons(refs.resultsIconEl);
  }

  destroy() {
    this.abortController?.abort();
    this.abortController = null;
    this.root = null;
    this.shell = null;
  }

  /** Solo UI: conserva datos/tabla al cambiar idioma. */
  onLanguageChange() {
    const refs = this.shell?.refs;
    if (!refs) return;
    if (refs.totalFilasDiv && this._lastInspectionRowCount != null) {
      refs.totalFilasDiv.textContent = t("arandanoMp.totalInspectionRows", {
        count: this._lastInspectionRowCount
      });
    }
    refs.resultsHeader?.querySelectorAll("th[data-excel-header], th[data-col-index]").forEach((th) => {
      const idx = Number(th.dataset.colIndex);
      if (Number.isFinite(idx) && STICKY_HEADER_SHORT_BY_JS[idx] != null) {
        th.textContent = STICKY_HEADER_SHORT_BY_JS[idx];
        th.title = STICKY_HEADER_TITLE_BY_JS[idx] || th.title;
        return;
      }
      const raw = th.dataset.excelHeader;
      if (raw != null && raw !== "") {
        th.textContent = translateExcelHeader(raw, idx);
        return;
      }
      if (Number.isFinite(idx)) {
        const cartilla = this.shell?.refs?.inspectionTypeSelect?.value;
        const col = (this.columnsByCartilla?.[cartilla] || []).find((c) => c.originalIndex === idx);
        if (col) th.textContent = formatResultColumnHeader(col);
      }
    });
    const okCell = refs.resultsBody?.querySelector("tr.agv-mp-row-ok td");
    if (okCell) okCell.textContent = t("arandanoMp.noInspectionErrors");
    syncArandanoMpStickyOffsets(refs.resultsTable);
  }
}
