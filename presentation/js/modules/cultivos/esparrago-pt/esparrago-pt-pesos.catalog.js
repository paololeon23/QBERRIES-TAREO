/** Base mercado → cliente → formato → rangos de peso, embalaje y presentación */

export const DB_PESOS = {
  ASIA: {
    "VALLE FRESH": { "10X500": { min: 515, max: 520, tipo: "ATADO", presentacion: "no validar" } },
    "RAT TRADING": { "10X500": { min: 515, max: 520, tipo: "ATADO", presentacion: "no validar" } },
    OMSEM: { "10X500": { min: 515, max: 520, tipo: "ATADO", presentacion: "no validar" } }
  },
  EUROPA: {
    AEI: {
      "10X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "no validar" },
      "20X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "no validar" }
    },
    "BARFOOTS OF BOTLEY LTD": {
      "11X450": { min: 4950, max: 5049, tipo: "CAJA", presentacion: "no validar" },
      "20X250": { min: 5000, max: 5100, tipo: "CAJA", presentacion: "Box 5.00 Kg" },
      "25X200": { min: 5000, max: 5100, tipo: "CAJA", presentacion: "Box 5.00 Kg" },
      "24X125": { min: 3000, max: 3060, tipo: "CAJA", presentacion: "Box 3.00 Kg" },
      "40X100": { min: 4080, max: 4160, tipo: "CAJA", presentacion: "no validar" }
    },
    "IPL ASDA": {
      "27X180": { min: 180, max: 184, tipo: "ATADO", presentacion: "4.86 KG" },
      "40X100": { min: 100, max: 102, tipo: "ATADO", presentacion: "4.00 KG" }
    },
    "GARCIA MATEO": {
      "12X250": { min: 250, max: 258, tipo: "ATADO", presentacion: "Box 3.00 Kg" },
      "20X250": { min: 250, max: 258, tipo: "ATADO", presentacion: "Box 5.00 Kg" },
      "17X300": { min: 300, max: 306, tipo: "ATADO", presentacion: "Box 5.1 Kg" },
      "6X400": { min: 400, max: 412, tipo: "ATADO", presentacion: "no validar" },
      "8X420": { min: 420, max: 433, tipo: "ATADO", presentacion: "Box 3.36 Kg" }
    },
    "PAPAGALLO PRODUCE EU LLC": {
      "17X300": { min: 300, max: 306, tipo: "ATADO", presentacion: "no validar" },
      "20X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "no validar" }
    },
    "AMS-EUROPEAN": {
      "8X420": { min: 420, max: 428, tipo: "ATADO", presentacion: "no validar" },
      "12X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "no validar" }
    },
    "NATURE'S PRIDE B.V.": {
      "10X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "Box 2.50 Kg" },
      "20X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "Box 5.00 Kg" },
      "20X200": { min: 200, max: 204, tipo: "ATADO", presentacion: "Box 4.0 Kg" },
      "11X450": { min: 450, max: 459, tipo: "ATADO", presentacion: "Box 5.00 Kg" }
    },
    "SPECIAL FRUIT NV": {
      "10X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "no validar" },
      "20X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "no validar" },
      "11X450": { min: 450, max: 459, tipo: "ATADO", presentacion: "no validar" }
    },
    "EDEKA AG FRUCHTKONTOR WEST": {
      "10X250": { min: 255, max: 260, tipo: "ATADO", presentacion: "no validar" }
    },
    "VERDE IMPORT": {
      "6X400": { min: 400, max: 408, tipo: "ATADO", presentacion: "no validar" },
      "10X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "no validar" },
      "11X450": { min: 450, max: 459, tipo: "ATADO", presentacion: "no validar" }
    },
    TEBOZA: {
      "10X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "Box 2.50 Kg" },
      "20X200": { min: 200, max: 204, tipo: "ATADO", presentacion: "Box 4.0 Kg" },
      "20X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "Box 5.00 Kg" },
      "11X450": { min: 450, max: 459, tipo: "ATADO", presentacion: "Box 5.00 Kg" }
    },
    "COOP TRADING": {
      "10X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "no validar" }
    },
    "AARTSEN BREDA B.V.": {
      "11X450": { min: 450, max: 459, tipo: "ATADO", presentacion: "Box 5.00 Kg" },
      "10X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "2.5 KG" }
    },
    "BC NATURE QUALITY FRUITS": {
      "20X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "Box 2.50 Kg" }
    },
    "CHAMPIÑONES Y SETAS GIMENEZ": {
      "10X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "Box 2.50 Kg" },
      "20X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "Box 5.00 Kg" }
    },
    "S&F GLOBAL FRESH EXOTICS": {
      "20X250": { min: 250, max: 255, tipo: "ATADO", presentacion: "Box 5.00 Kg" },
      "11X454": { min: 454, max: 463, tipo: "ATADO", presentacion: "Box 5.00 Kg" }
    }
  },
  USA: {
    "FARM DIRECT SUPPLY, LLC": {
      "11X450": { min: 4950, max: 4960, tipo: "CAJA", presentacion: "Box 5.00 Kg" },
      "10X545": { min: 5450, max: 5559, tipo: "CAJA", presentacion: "5.45 Kg" },
      "11X454": { min: 4994, max: 5094, tipo: "CAJA", presentacion: "Box 5.00 Kg" }
    },
    "Square One Farms, Llc": {
      "11X450": { min: 4950, max: 4960, tipo: "CAJA", presentacion: "Box 5.00 Kg" },
      "10X544": { min: 544, max: 545, tipo: "ATADO", presentacion: "Box 12 Lb" }
    },
    "HARVEST SENSATIONS LLC.": {
      "11X450": { min: 4950, max: 4960, tipo: "CAJA", presentacion: "Box 5.00 Kg" }
    },
    "WALMART INC.": {
      "11X450": { min: 5000, max: 5010, tipo: "CAJA", presentacion: "no validar" }
    },
    "PRIME TIME INTERNATIONAL": {
      "11X450": { min: 4950, max: 4960, tipo: "CAJA", presentacion: "Box 5.00 Kg" }
    },
    "JMB PRODUCE": {
      "11X450": { min: 4950, max: 4960, tipo: "CAJA", presentacion: "Box 5.00 Kg" }
    },
    "ALPINE FRESH": {
      "11X450": { min: 4950, max: 4960, tipo: "CAJA", presentacion: "Box 5.00 Kg" }
    },
    "ALPINE FRESH – SMALL": {
      "11X450": { min: 450, max: 470, tipo: "ATADO", presentacion: "Box 5.00 Kg" }
    },
    "FRU VEG MARKETING INC.": {
      "11X450": { min: 4950, max: 4960, tipo: "CAJA", presentacion: "Box 5.00 Kg" },
      "15X325": { min: 325, max: 332, tipo: "ATADO", presentacion: "Box 4.8 Kg" },
      "28X453": { min: 450, max: 453, tipo: "ATADO", presentacion: "Box 28 Lb" }
    },
    METRO: {
      "15X325": { min: 4875, max: 4875, tipo: "CAJA", presentacion: "no validar" },
      "11X450": { min: 5170, max: 5170, tipo: "CAJA", presentacion: "no validar" }
    }
  }
};

/** Normaliza nombre de cliente: apóstrofos tipográficos (Excel), guiones, mayúsculas. */
function normalizeClienteNombre(value) {
  return (value || "")
    .toString()
    .trim()
    .replace(/[\u2018\u2019\u201A\u201B\u0060\u00B4]/g, "'")
    .replace(/[\u2013\u2014]/g, "-")
    .toUpperCase();
}

export function getRegisteredClientNames() {
  return Object.values(DB_PESOS).flatMap((mercado) => Object.keys(mercado));
}

export function isClientRegistered(clienteNombre) {
  const nombre = normalizeClienteNombre(clienteNombre);
  if (!nombre) return false;
  return getRegisteredClientNames().some((reg) => nombre.includes(normalizeClienteNombre(reg)));
}

export function resolveClienteKey(mercado, cliente, calibre) {
  const mercadoUpper = normalizeClienteNombre(mercado);
  const clienteUpper = normalizeClienteNombre(cliente);
  const calibreUpper = normalizeClienteNombre(calibre);
  const mercadoData = DB_PESOS[mercadoUpper];
  if (!mercadoData) return null;

  let clienteKey = Object.keys(mercadoData).find((k) =>
    clienteUpper.includes(normalizeClienteNombre(k))
  );
  if (clienteKey === "ALPINE FRESH" && calibreUpper === "SMALL") {
    clienteKey = "ALPINE FRESH – SMALL";
  }
  return clienteKey || null;
}

function normalizeFormato(value) {
  return (value || "").toString().trim().toUpperCase().replace(/\s/g, "");
}

function getFormatoEntry(mercado, cliente, formato, calibre) {
  const mercadoUpper = normalizeClienteNombre(mercado);
  const clienteKey = resolveClienteKey(mercado, cliente, calibre);
  if (!clienteKey) return null;
  return DB_PESOS[mercadoUpper]?.[clienteKey]?.[normalizeFormato(formato)] ?? null;
}

function normalizeFormatoEntries(entry) {
  if (!entry) return [];
  return Array.isArray(entry) ? entry : [entry];
}

export function getPesoRangosForFormato(mercado, cliente, formato, calibre) {
  return normalizeFormatoEntries(getFormatoEntry(mercado, cliente, formato, calibre));
}

export function getPesoRango(mercado, cliente, formato, calibre, presentacion = "") {
  const entries = getPesoRangosForFormato(mercado, cliente, formato, calibre);
  if (!entries.length) return null;

  const presUpper = (presentacion || "").toString().trim().toUpperCase();
  if (presUpper) {
    const byPresentacion = entries.find((entry) => entry.presentacion.toUpperCase() === presUpper);
    if (byPresentacion) return byPresentacion;
  }

  return entries[0];
}
