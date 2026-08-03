import { i18nService } from "../../../services/i18n.service.js";
import { buildValidacionesPorColumnaFromReglas } from "../../../../../engine/cartilla-rules.adapter.js";

/**
 * Helpers de rendimiento para tablas de resultados MP.
 * Muestra columnas pedidas en contextColsJs (p. ej. todas con reglas) + columnas con error.
 */

export const DEFAULT_MP_CONTEXT_COLS_JS = [0, 1, 6, 9, 10];
export const DEFAULT_MP_STICKY_COLS_JS = [0, 1, 6, 9];

/**
 * Bloque SAP (Excel 13–27 y 29–33 → JS).
 * Nota Condición (Excel 28 / JS 27) NO es SAP (puede venir de WULUT) pero sí se muestra en UI.
 */
export const SAP_DATA_COLS_JS = [
  12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, // Productor…Linea
  28, 29, 30, 31, 32 // Tipo formato…Peso Bruto
];
export const NOTA_CONDICION_COL_JS = 27;
/** Zona a pintar en frontend: SAP + Nota Condición (WULUT). Excel 13–33 → JS 12–32. */
export const SAP_ZONE_FRONTEND_COLS_JS = Array.from({ length: 21 }, (_, i) => 12 + i);

/**
 * Por ahora: en PT de todos los cultivos no se validan ni se muestran datos SAP.
 * MP / Plagas no usan este flag. Nota Condición (JS 27) no es SAP y sí se mantiene.
 */
export const PT_SKIP_SAP_VALIDATION = true;

export function isPtSapDataColJs(colJs) {
  return SAP_DATA_COLS_JS.includes(Number(colJs));
}

/** Excel 1-based: 13–27 y 29–33 (excluye Nota Condición 28). */
export function isPtSapDataColNum(colNum) {
  const n = Number(colNum);
  return (n >= 13 && n <= 27) || (n >= 29 && n <= 33);
}

/** Quita del errorMap cualquier incidencia en columnas SAP (PT). */
export function stripPtSapValidationErrors(errorMap) {
  if (!PT_SKIP_SAP_VALIDATION || !errorMap?.size) return errorMap;
  [...errorMap.entries()].forEach(([filaNum, filaMap]) => {
    if (!filaMap?.size) return;
    [...filaMap.keys()].forEach((colNum) => {
      if (isPtSapDataColNum(colNum)) filaMap.delete(colNum);
    });
    if (!filaMap.size) errorMap.delete(filaNum);
  });
  return errorMap;
}

/** Filtra índices JS SAP de listas de columnas visibles PT. */
export function filterOutPtSapCols(cols = []) {
  if (!PT_SKIP_SAP_VALIDATION) return cols;
  return cols.filter((c) => typeof c !== "number" || !isPtSapDataColJs(c));
}

/**
 * Nombres canónicos del bloque (Excel 13–33).
 * Nota Condición (27) puede venir de WULUT; el resto son datos SAP.
 */
export const SAP_ZONE_HEADER_LABELS_BY_JS = {
  12: "Productor",
  13: "Guía de Remisión",
  14: "Etapa",
  15: "Campo",
  16: "Turno",
  17: "Fundo",
  18: "Variedad",
  19: "Fecha Cosecha",
  20: "Fecha de Producción",
  21: "Tecnologia de Postcosecha PT",
  22: "Calibre",
  23: "Tipo de Embalado",
  24: "Categoria",
  25: "Turno Linea",
  26: "Linea",
  27: "Nota Condición",
  28: "Tipo de formato",
  29: "Etiqueta",
  30: "Jaba",
  31: "Viaje",
  32: "Peso Bruto"
};

/**
 * Encabezado UI del bloque Productor…Peso Bruto / Nota Condición.
 * Sin etiqueta «(SAP)» en ningún cultivo: solo el nombre de columna (las celdas sí se pintan).
 * @param {number} colJs
 * @param {string} [excelHeader]
 * @returns {{ label: string, isSap: boolean, title: string }}
 */
export function resolveSapZoneHeader(colJs, excelHeader = "", options = {}) {
  const idx = Number(colJs);
  const known = SAP_ZONE_HEADER_LABELS_BY_JS[idx];
  if (!known) {
    const raw = String(excelHeader || "").trim();
    return { label: raw || `Col ${idx + 1}`, isSap: false, title: raw };
  }
  const hideSapTag = options.suppressSapLabel !== false; // default: ocultar «(SAP)»
  if (hideSapTag) {
    const isNota = idx === NOTA_CONDICION_COL_JS;
    return {
      label: known,
      isSap: false,
      title: isNota ? "Nota Condición" : known
    };
  }
  const isSap = idx !== NOTA_CONDICION_COL_JS;
  const label = isSap ? `${known} (SAP)` : `${known} (WULUT)`;
  const title = isSap ? known : "Nota Condición";
  return { label, isSap, title };
}

/**
 * Índices JS siempre útiles: sticky/contexto + columnas con reglas (+ compuestas/cruzadas) + zona SAP.
 * @param {object|null} reglas
 * @param {{ contextColsJs?: number[], stickyColsJs?: number[], includeSapZone?: boolean }} [options]
 * @returns {number[]}
 */
export function collectValidatedColumnIndexesJs(reglas, options = {}) {
  const needed = new Set([
    ...(options.contextColsJs || DEFAULT_MP_CONTEXT_COLS_JS),
    ...(options.stickyColsJs || DEFAULT_MP_STICKY_COLS_JS)
  ]);

  if (options.includeSapZone !== false) {
    SAP_ZONE_FRONTEND_COLS_JS.forEach((idx) => needed.add(idx));
  }

  buildValidacionesPorColumnaFromReglas(reglas).forEach((entry) => {
    if (Number.isFinite(entry?.indice_js)) needed.add(entry.indice_js);
  });

  // Columnas con cualquier regla activa en JSON (incluye debe-estar-vacio).
  (reglas?.columnas || []).forEach((col) => {
    const keys = Object.keys(col || {});
    const hasRule = keys.some(
      (k) => !["numero", "nombre-de-la-columna", "que-se-revisa", "aplica-a-cartilla"].includes(k)
    );
    if (!hasRule) return;
    const js = Number(col.numero) - 1;
    if (Number.isFinite(js) && js >= 0) needed.add(js);
  });

  const addExcelCol = (n) => {
    const js = Number(n) - 1;
    if (Number.isFinite(js) && js >= 0) needed.add(js);
  };

  [...(reglas?.["validaciones-compuestas"] || []), ...(reglas?.["validaciones-cruzadas"] || [])].forEach(
    (regla) => {
      addExcelCol(regla["columna-a"]);
      addExcelCol(regla["columna-b"]);
      addExcelCol(regla["igual-a-columna"]);
      (regla.columnas || regla["columnas-a-sumary"] || []).forEach(addExcelCol);
    }
  );

  return [...needed].sort((a, b) => a - b);
}

/**
 * @param {Array<{ originalIndex: number, header?: string }>} allColumns
 * @param {Map<number, Map<number, unknown>>|null} errorMap filaNum → (colExcel → err)
 * @param {Array<{ _filaNum?: number }>} filas
 * @param {number[]} [contextColsJs]
 */
export function pickResultColumns(
  allColumns,
  errorMap,
  filas,
  contextColsJs = DEFAULT_MP_CONTEXT_COLS_JS
) {
  const needed = new Set(contextColsJs);
  (filas || []).forEach((row) => {
    const filaMap = errorMap?.get(row._filaNum);
    if (!filaMap) return;
    filaMap.forEach((_err, colNum) => needed.add(colNum - 1));
  });
  return (allColumns || []).filter((col) => needed.has(col.originalIndex));
}

/**
 * Índices JS a mostrar cuando no hay lista de columnas tipada (ej. Palta).
 * @param {number} totalCols
 * @param {Map<number, Map<number, unknown>>|null} errorMap
 * @param {Array<{ _filaNum?: number }>} filas
 * @param {number[]} [contextColsJs]
 * @param {number[]} [stickyColsJs]
 */
export function pickResultColumnIndexes(
  totalCols,
  errorMap,
  filas,
  contextColsJs = DEFAULT_MP_CONTEXT_COLS_JS,
  stickyColsJs = [0, 1, 6, 9]
) {
  const needed = new Set([...contextColsJs, ...stickyColsJs]);
  (filas || []).forEach((row) => {
    const filaMap = errorMap?.get(row._filaNum);
    if (!filaMap) return;
    filaMap.forEach((_err, colNum) => {
      const js = colNum - 1;
      if (js >= 0 && js < totalCols) needed.add(js);
    });
  });
  return [...needed].filter((i) => i >= 0 && i < totalCols).sort((a, b) => a - b);
}

/**
 * Agrupa filas por fecha de inspección (usa cache _fechaInspeccionISO si existe).
 */
export function groupRowsByInspectionDate(rows, colFechaJs, parseExcelDateISO) {
  const byDate = new Map();
  (rows || []).forEach((row) => {
    const iso = row._fechaInspeccionISO || parseExcelDateISO(row[colFechaJs]);
    if (!iso) return;
    if (!row._fechaInspeccionISO) row._fechaInspeccionISO = iso;
    if (!byDate.has(iso)) byDate.set(iso, []);
    byDate.get(iso).push(row);
  });
  return byDate;
}

/**
 * Placeholders colapsables; la tabla se monta al abrir (toggle).
 * @param {Array} items con filasDetalle
 * @param {(item: object, idx: number) => string} metaHtml
 */
export function buildLazyDateDetailPlaceholders(items, htmlEscape, metaLabel) {
  return (items || [])
    .filter((item) => item.filasDetalle?.length)
    .map(
      (item, idx) => `
        <details class="agv-mp-date-detail" data-todo-detail="${idx}">
          <summary class="agv-mp-date-detail__head">
            <h4 class="agv-mp-date-detail__title">${htmlEscape(item.fecha)}</h4>
            <span class="agv-mp-date-detail__meta">${htmlEscape(metaLabel(item))} · clic para ver</span>
          </summary>
          <div class="agv-mp-date-detail__body" data-todo-body="${idx}">
            <p class="agv-mp-date-detail__loading">${htmlEscape(i18nService.translate("common.loadingTable"))}</p>
          </div>
        </details>`
    )
    .join("");
}

/**
 * Bind lazy load de tablas en details[data-todo-detail].
 */
export function bindLazyDateDetailTables(container, detailItems, renderTableHtml) {
  if (!container) return;
  const items = (detailItems || []).filter((item) => item.filasDetalle?.length);
  container.querySelectorAll("details[data-todo-detail]").forEach((detailsEl) => {
    detailsEl.addEventListener("toggle", () => {
      if (!detailsEl.open || detailsEl.dataset.loaded === "1") return;
      const idx = Number(detailsEl.dataset.todoDetail);
      const item = items[idx];
      const body = detailsEl.querySelector(`[data-todo-body="${idx}"]`);
      if (!item || !body) return;
      body.innerHTML = renderTableHtml(item, idx);
      detailsEl.dataset.loaded = "1";
    });
  });
}
