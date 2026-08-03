/** Orden y formato de exportación Excel filtrado — Espárrago MP (MPES). */

const VACIO = 0;

function range(from, to) {
  const out = [];
  for (let c = from; c <= to; c++) out.push(c);
  return out;
}

/** 1, vacío×4, 10, 11, 13, vacío, 14-19, vacío, 29, vacío, 30, vacío, 33, 28, 34-149, 153-167 */
export function buildExportOrderMPES() {
  return [
    1,
    VACIO,
    VACIO,
    VACIO,
    VACIO,
    10,
    11,
    13,
    VACIO,
    ...range(14, 19),
    VACIO,
    29,
    VACIO,
    30,
    VACIO,
    33,
    28,
    ...range(34, 149),
    ...range(153, 167),
  ];
}
export const EXPORT_ORDER_BY_CARTILLA = {
  MPES: buildExportOrderMPES(),
};

export function parseFlexibleNumber(val) {
  if (val === null || val === undefined) return NaN;
  if (typeof val === "object") {
    if (val.v != null) return parseFlexibleNumber(val.v);
    if (val.w != null) return parseFlexibleNumber(val.w);
  }
  const texto = String(val).trim().replace(/\s/g, "");
  if (!texto) return NaN;
  const normal = texto.includes(",") && !texto.includes(".")
    ? texto.replace(",", ".")
    : texto.replace(/,/g, "");
  const n = Number(normal);
  return Number.isFinite(n) ? n : NaN;
}

function formatExportDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

function rawValueForExportColumn(row, excelCol) {
  if (excelCol === VACIO) return "";
  return row[excelCol - 1];
}

export function formatExportValue(val, excelCol, exportCfg, formatISOToDMY, parseExcelDateISO, colSets, headers = []) {
  if (excelCol === VACIO) return "";

  const textCols = colSets?.textCols || new Set(exportCfg?.["columnas-texto"] || [10, 19]);
  const dateCols = colSets?.dateCols || new Set(exportCfg?.["columnas-fecha"] || [20, 47, 48, 57]);
  const header = String(headers[excelCol - 1] ?? "");
  const headerBlocks =
    header &&
    /linea|asp|formato|destino|lote|usuario|hora|productor|guia/i.test(header) &&
    !/fecha|lmr|date/i.test(header);

  if (dateCols.has(excelCol) && !headerBlocks) {
    const iso = parseExcelDateISO(val);
    if (iso) return formatExportDate(iso);
    const texto = val === null || val === undefined ? "" : String(val).trim();
    if (/^\d{2}-\d{2}-\d{4}$/.test(texto)) return texto.replace(/-/g, "/");
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) return texto;
    return "";
  }

  if (textCols.has(excelCol)) {
    return val === null || val === undefined ? "" : String(val);
  }

  const n = parseFlexibleNumber(val);
  if (!Number.isNaN(n)) return n;
  return val === null || val === undefined ? "" : String(val);
}

function cellForFilteredExport(val, excelCol, textCols) {
  if (excelCol === VACIO || val === "") return "";

  if (textCols.has(excelCol)) {
    return { v: String(val), t: "s" };
  }
  if (typeof val === "number") {
    return { v: val, t: "n" };
  }
  return { v: String(val), t: "s" };
}

export function headerForExportColumn(excelCol, headers) {
  if (excelCol === VACIO) return "";
  const js = excelCol - 1;
  return headers[js] ?? "";
}

function getExportColSets(exportCfg) {
  return {
    textCols: new Set(exportCfg?.["columnas-texto"] || [10, 19]),
    dateCols: new Set(exportCfg?.["columnas-fecha"] || [20, 47, 48, 57])
  };
}

export function buildFilteredSheetData(rows, cartilla, headers, exportCfg, helpers) {
  const order =
    EXPORT_ORDER_BY_CARTILLA[cartilla] ||
    exportCfg?.["reordenar-columnas-excel"] ||
    buildExportOrderMPES();

  const { formatISOToDMY, parseExcelDateISO } = helpers;
  const colSets = getExportColSets(exportCfg);
  const wsData = [];

  wsData.push(
    order.map((col) => ({
      v: headerForExportColumn(col, headers),
      t: "s",
      s: { font: { bold: true } },
    }))
  );

  rows.forEach((row) => {
    wsData.push(
      order.map((excelCol) =>
        cellForFilteredExport(
          formatExportValue(
            rawValueForExportColumn(row, excelCol),
            excelCol,
            exportCfg,
            formatISOToDMY,
            parseExcelDateISO,
            colSets,
            headers
          ),
          excelCol,
          colSets.textCols
        )
      )
    );
  });
  return wsData;
}

export function buildFullSheetDataWithErrors(
  rows,
  headers,
  totalCols,
  exportCfg,
  getCellMeta,
  helpers
) {
  const { formatISOToDMY, parseExcelDateISO, estiloExportCeldaError } = helpers;
  const colSets = getExportColSets(exportCfg);
  const wsData = [];

  wsData.push(
    headers.map((header) => ({
      v: header || "",
      t: "s",
      s: { font: { bold: true } },
    }))
  );

  rows.forEach((row) => {
    const linea = [];
    for (let js = 0; js < totalCols; js++) {
      const excelCol = js + 1;
      const val = formatExportValue(
        row[js],
        excelCol,
        exportCfg,
        formatISOToDMY,
        parseExcelDateISO,
        colSets,
        headers
      );
      const { cellClass } = getCellMeta(row, js);

      if (!cellClass) {
        if (val === "") linea.push("");
        else if (typeof val === "number") linea.push(val);
        else linea.push({ v: val, t: "s" });
      } else {
        linea.push({
          v: val === "" ? "" : val,
          t: typeof val === "number" ? "n" : "s",
          s: estiloExportCeldaError(cellClass),
        });
      }
    }
    wsData.push(linea);
  });

  return wsData;
}
