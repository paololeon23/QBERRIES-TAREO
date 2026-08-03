/** Orden y formato de exportación Excel filtrado — Arándano MP (104 cols). */

const VACIO = 0;

function range(from, to) {
  const out = [];
  for (let c = from; c <= to; c++) out.push(c);
  return out;
}

function commonExportPrefix() {
  // Lote…Campo, Turno, Fundo…Fecha Cosecha + huecos + Tipo formato…Peso Bruto
  return [10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, VACIO, 29, 30, 31, 32, 33, VACIO, VACIO];
}

/** MPBAR: prefijo + col 34 hasta el final en orden original. */
export function buildExportOrderMPBA() {
  return [...commonExportPrefix(), ...range(34, 104)];
}

/**
 * MPHAR: el layout ya se normaliza al cargar (orden canónico).
 * Export = orden natural 1..104 (prefijo común + resto).
 */
export function buildExportOrderMPHA() {
  return [...commonExportPrefix(), ...range(34, 104)];
}

/** MPGAR: prefijo + ORDER_MPGAR (104 cols). */
export function buildExportOrderMPGA() {
  return [
    ...commonExportPrefix(),
    ...range(34, 66),
    80,
    ...range(67, 72),
    81,
    ...range(73, 78),
    82,
    79,
    ...range(83, 86),
    100,
    ...range(87, 92),
    101,
    ...range(93, 98),
    102,
    99,
    103,
    104
  ];
}

export const EXPORT_ORDER_BY_CARTILLA = {
  MPBA: buildExportOrderMPBA(),
  MPHA: buildExportOrderMPHA(),
  MPGA: buildExportOrderMPGA()
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

export function formatExportValue(val, excelCol, exportCfg, formatISOToDMY, parseExcelDateISO, headers = []) {
  if (excelCol === VACIO) return "";

  const textCols = new Set(exportCfg?.["columnas-texto"] || [4, 5, 10, 19, 20, 41, 51]);
  const dateCols = new Set(exportCfg?.["columnas-fecha"] || [41, 51]);
  const header = String(headers[excelCol - 1] ?? "");
  const headerOk =
    !header ||
    /fecha|lmr|date/i.test(header) ||
    !/linea|asp|formato|destino|lote|usuario|hora|productor|guia/i.test(header);

  if (dateCols.has(excelCol) && headerOk) {
    const iso = parseExcelDateISO(val);
    return iso ? formatISOToDMY(iso) : "";
  }

  if (textCols.has(excelCol)) {
    return val === null || val === undefined ? "" : String(val);
  }

  const n = parseFlexibleNumber(val);
  if (!Number.isNaN(n)) return n;
  return val === null || val === undefined ? "" : String(val);
}

function cellForFilteredExport(val, excelCol, exportCfg) {
  if (excelCol === VACIO || val === "") return "";

  const textCols = new Set(exportCfg?.["columnas-texto"] || [4, 5, 10, 19, 20, 41, 51]);
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

export function buildFilteredSheetData(rows, cartilla, headers, exportCfg, helpers) {
  const order =
    EXPORT_ORDER_BY_CARTILLA[cartilla] ||
    exportCfg?.["reordenar-columnas-excel"] ||
    buildExportOrderMPBA();

  const { formatISOToDMY, parseExcelDateISO } = helpers;
  const wsData = [];

  wsData.push(
    order.map((col) => ({
      v: headerForExportColumn(col, headers),
      t: "s",
      s: { font: { bold: true } }
    }))
  );

  rows.forEach((row) => {
    wsData.push(
      order.map((excelCol) =>
        cellForFilteredExport(
          formatExportValue(row[excelCol - 1], excelCol, exportCfg, formatISOToDMY, parseExcelDateISO, headers),
          excelCol,
          exportCfg
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
  const wsData = [];

  wsData.push(
    headers.map((header) => ({
      v: header || "",
      t: "s",
      s: { font: { bold: true } }
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
        undefined,
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
          s: estiloExportCeldaError(cellClass)
        });
      }
    }
    wsData.push(linea);
  });

  return wsData;
}
