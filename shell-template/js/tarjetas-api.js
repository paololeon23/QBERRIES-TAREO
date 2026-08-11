/**
 * Cliente API Tarjeta Pallet (solo lectura).
 *
 * Proxy Netlify oficial en qpack (HTTPS, token solo en servidor).
 * NUNCA script.google.com ni API_TOKEN en el navegador.
 * Mismo API_TOKEN que Permisos; solo cambia el Web App / Sheet (APPS_SCRIPT_URL en qpack).
 */

/** Proxy oficial QPack · Tarjeta Pallet (lectura segura) */
export const TARJETAS_API_BASE = "https://qpack.netlify.app/api/tarjetas";

const REQUEST_TIMEOUT_MS = 18000;
const MAX_ATTEMPTS = 3;
const RETRY_DELAY_MS = 350;
const SOFT_CACHE_MS = 10000;
const DURABLE_CACHE_KEY = "qb_tarjetas_cache_v1";
const DURABLE_CACHE_MAX_ENTRIES = 12;

/** @type {Map<string, { at: number, data: any }>} */
const softCache = new Map();
/** @type {Map<string, Promise<any>>} */
const inflight = new Map();

function buildQuery(params = {}) {
  const qs = new URLSearchParams();
  Object.entries(params).forEach(([key, value]) => {
    if (value === undefined || value === null || value === "") return;
    if (value === false) return;
    qs.set(key, String(value));
  });
  return qs.toString();
}

function readDurableStore() {
  try {
    const raw = localStorage.getItem(DURABLE_CACHE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function writeDurableStore(store) {
  try {
    localStorage.setItem(DURABLE_CACHE_KEY, JSON.stringify(store));
  } catch {
    /* cuota / modo privado */
  }
}

function saveDurableCache(cacheKey, data) {
  if (!cacheKey || !data || data.ok === false) return;
  if (!String(cacheKey).includes("reporteTarjetas") && !String(cacheKey).includes("listarTarjetas")) {
    return;
  }
  const store = readDurableStore();
  store[cacheKey] = { at: Date.now(), data };
  const keys = Object.keys(store);
  if (keys.length > DURABLE_CACHE_MAX_ENTRIES) {
    keys
      .sort((a, b) => (store[a]?.at || 0) - (store[b]?.at || 0))
      .slice(0, keys.length - DURABLE_CACHE_MAX_ENTRIES)
      .forEach((k) => delete store[k]);
  }
  writeDurableStore(store);
}

function loadDurableCache(cacheKey) {
  const entry = readDurableStore()[cacheKey];
  if (!entry?.data) return null;
  return { at: Number(entry.at) || 0, data: entry.data };
}

function withCacheMeta(data, cachedAt) {
  return {
    ...data,
    fromCache: true,
    cachedAt: cachedAt || Date.now()
  };
}

function apiErrorMessage(data, fallback) {
  if (!data || typeof data !== "object") return fallback;
  const code = String(data.code || "").toUpperCase();
  if (code === "UNAUTHORIZED") {
    return "No se pudo autenticar con el servidor de tarjetas. Revisa la configuración del proxy Netlify.";
  }
  if (code === "NO_CONFIG" || code === "CONFIG") {
    return "Faltan variables de entorno en el proxy Netlify (APPS_SCRIPT_URL / API_TOKEN).";
  }
  return data.message || data.error || fallback;
}

function sleep(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function resolveApiBases() {
  if (typeof window !== "undefined" && window.QB_TARJETAS_PROXY) {
    return [String(window.QB_TARJETAS_PROXY)];
  }
  const bases = [];
  try {
    const host = window.location.hostname || "";
    // Si estamos en el mismo sitio del proxy → same-origin (más rápido, HTTPS)
    if (host === "qpack.netlify.app") {
      bases.push("/api/tarjetas");
    }
  } catch (_) {}
  bases.push(TARJETAS_API_BASE);
  return bases;
}

async function fetchOnce(url) {
  const controller = new AbortController();
  const timer = window.setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    const response = await fetch(url, {
      method: "GET",
      headers: { Accept: "application/json" },
      cache: "no-store",
      mode: "cors",
      credentials: "omit",
      signal: controller.signal
    });
    const text = await response.text();
    let data = null;
    try {
      data = JSON.parse(text);
    } catch {
      throw new Error("El servidor de tarjetas no devolvió JSON válido.");
    }
    if (data && data.ok === false) {
      throw new Error(apiErrorMessage(data, "Error en API de tarjetas"));
    }
    if (!response.ok) {
      throw new Error(apiErrorMessage(data, `API tarjetas respondió ${response.status}`));
    }
    return data;
  } finally {
    window.clearTimeout(timer);
  }
}

function isRetryableNetworkError(err) {
  const msg = String(err?.message || err || "").toLowerCase();
  if (err?.name === "AbortError") return true;
  if (msg.includes("failed to fetch")) return true;
  if (msg.includes("network")) return true;
  if (msg.includes("load failed")) return true;
  if (msg.includes("aborted")) return true;
  return false;
}

export async function callTarjetasApi(params = {}, options = {}) {
  const force = Boolean(options.force);
  const qs = buildQuery(params);
  const cacheKey = qs || "__ping__";

  const cached = softCache.get(cacheKey);
  if (!force && cached && Date.now() - cached.at < SOFT_CACHE_MS) {
    return cached.data;
  }

  const pending = inflight.get(cacheKey);
  if (pending) return pending;

  const job = (async () => {
    const bases = resolveApiBases();
    let lastErr = null;

    for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
      for (const base of bases) {
        const url = qs ? `${base}?${qs}` : base;
        try {
          const data = await fetchOnce(url);
          softCache.set(cacheKey, { at: Date.now(), data });
          saveDurableCache(cacheKey, data);
          return data;
        } catch (err) {
          lastErr = err;
          const msg = String(err?.message || "");
          if (/autoriz|config|variables de entorno/i.test(msg)) throw err;
        }
      }
      if (attempt < MAX_ATTEMPTS) await sleep(RETRY_DELAY_MS * attempt);
    }

    if (cached?.data) return withCacheMeta(cached.data, cached.at);

    const durable = loadDurableCache(cacheKey);
    if (durable?.data) {
      softCache.set(cacheKey, { at: Date.now(), data: durable.data });
      return withCacheMeta(durable.data, durable.at);
    }

    if (isRetryableNetworkError(lastErr) || lastErr?.name === "AbortError") {
      throw new Error("Sin conexión con el servidor de tarjetas. Revisa tu red e intenta de nuevo.");
    }
    throw lastErr instanceof Error
      ? lastErr
      : new Error(String(lastErr || "No se pudo consultar tarjetas"));
  })();

  inflight.set(cacheKey, job);
  try {
    return await job;
  } finally {
    inflight.delete(cacheKey);
  }
}

export async function pingTarjetas() {
  const data = await callTarjetasApi({ action: "ping" });
  if (!data?.ok) throw new Error(apiErrorMessage(data, "Ping de tarjetas falló"));
  return data;
}

/** Fecha calendario America/Lima → yyyy-MM-dd */
export function toLimaYmd(value) {
  if (value === undefined || value === null || value === "") return "";
  const raw = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

  const isoDay = raw.match(/^(\d{4}-\d{2}-\d{2})T/);
  if (isoDay) {
    const d = new Date(raw);
    if (!Number.isNaN(d.getTime())) {
      return new Intl.DateTimeFormat("en-CA", {
        timeZone: "America/Lima",
        year: "numeric",
        month: "2-digit",
        day: "2-digit"
      }).format(d);
    }
    return isoDay[1];
  }

  const m = raw.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
  if (m) {
    return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  }

  const d = new Date(raw);
  if (Number.isNaN(d.getTime())) return raw.slice(0, 10);
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Lima",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(d);
}

function normalizeTarjetaRow(row) {
  if (!row || typeof row !== "object") return row;
  return {
    ...row,
    fecha: toLimaYmd(row.fecha),
    lugar: row.lugar == null ? "" : String(row.lugar).trim(),
    variedad: row.variedad == null ? "" : String(row.variedad).trim(),
    correlativo: row.correlativo == null ? "" : String(row.correlativo).trim(),
    modulo: row.modulo == null ? "" : String(row.modulo).trim(),
    turno: row.turno == null ? "" : String(row.turno).trim(),
    lote: row.lote == null ? "" : String(row.lote).trim(),
    placa: row.placa == null ? "" : String(row.placa).trim().toUpperCase(),
    nGuia: row.nGuia == null ? "" : String(row.nGuia).trim(),
    dni: row.dni == null ? "" : String(row.dni).trim(),
    horaRegistrada: row.horaRegistrada == null ? "" : String(row.horaRegistrada).trim(),
    jarras: row.jarras === "" || row.jarras == null ? "" : Number(row.jarras) || 0,
    jabas: row.jabas === "" || row.jabas == null ? "" : Number(row.jabas) || 0,
    pesoBruto: row.pesoBruto === "" || row.pesoBruto == null ? "" : Number(row.pesoBruto) || 0,
    pesoNeto: row.pesoNeto === "" || row.pesoNeto == null ? "" : Number(row.pesoNeto) || 0
  };
}

function emptyKpis() {
  return {
    totalGuias: 0,
    totalFilas: 0,
    totalCorrelativos: 0,
    totalJarras: 0,
    totalJabas: 0,
    pesoBrutoTotal: 0,
    pesoNetoTotal: 0
  };
}

export function extractTarjetasKpis(payload = {}, rows = []) {
  const raw = payload.kpis && typeof payload.kpis === "object" ? payload.kpis : null;
  if (raw) {
    return {
      totalGuias: Number(raw.totalGuias) || 0,
      totalFilas: Number(raw.totalFilas) || rows.length || 0,
      totalCorrelativos: Number(raw.totalCorrelativos) || 0,
      totalJarras: Number(raw.totalJarras) || 0,
      totalJabas: Number(raw.totalJabas) || 0,
      pesoBrutoTotal: Number(raw.pesoBrutoTotal) || 0,
      pesoNetoTotal: Number(raw.pesoNetoTotal) || 0
    };
  }

  const guias = new Set();
  const corr = new Set();
  let jarras = 0;
  let jabas = 0;
  let bruto = 0;
  let neto = 0;
  rows.forEach((r) => {
    if (r.nGuia) guias.add(`${r.nGuia}|${r.fecha || ""}`);
    if (r.correlativo) corr.add(String(r.correlativo).toUpperCase());
    jarras += Number(r.jarras) || 0;
    jabas += Number(r.jabas) || 0;
    bruto += Number(r.pesoBruto) || 0;
    neto += Number(r.pesoNeto) || 0;
  });
  return {
    totalGuias: guias.size,
    totalFilas: rows.length,
    totalCorrelativos: corr.size,
    totalJarras: Math.round(jarras * 100) / 100,
    totalJabas: Math.round(jabas * 100) / 100,
    pesoBrutoTotal: Math.round(bruto * 100) / 100,
    pesoNetoTotal: Math.round(neto * 100) / 100
  };
}

function buildReporteParams(opts = {}) {
  const params = { action: "reporteTarjetas" };

  const fecha = opts.fecha != null ? String(opts.fecha).trim() : "";
  const todas = opts.todas === true || opts.todas === 1 || opts.todas === "1";

  if (todas) {
    params.todas = "1";
    params.limit = opts.limit ?? 2000;
  } else if (fecha) {
    params.fecha = fecha;
    params.limit = opts.limit ?? 500;
  } else {
    params.limit = opts.limit ?? 500;
  }

  if (opts.placa) params.placa = String(opts.placa).trim();
  if (opts.fundo) params.fundo = String(opts.fundo).trim();
  if (opts.variedad) params.variedad = String(opts.variedad).trim();
  if (opts.q) params.q = String(opts.q).trim();
  if (opts.soloKpis) params.soloKpis = "1";
  if (opts.offset != null) params.offset = opts.offset;

  return params;
}

function packReporteResponse(data) {
  if (!data?.ok) throw new Error(apiErrorMessage(data, "No se pudieron listar las tarjetas"));
  const rows = Array.isArray(data.data) ? data.data.map(normalizeTarjetaRow) : [];
  const kpis = extractTarjetasKpis(data, rows);
  return {
    ...data,
    data: rows,
    count: Number(data.count) || kpis.totalFilas || rows.length,
    kpis,
    fromCache: Boolean(data.fromCache),
    cachedAt: data.cachedAt || null
  };
}

/**
 * Lee soft/durable cache sin red (stale-while-revalidate).
 */
export function peekCachedReporteTarjetas(opts = {}) {
  const params = buildReporteParams(opts);
  const cacheKey = buildQuery(params);
  const soft = softCache.get(cacheKey);
  if (soft?.data?.ok !== false && soft?.data) {
    try {
      return packReporteResponse({ ...soft.data, fromCache: true, cachedAt: soft.at });
    } catch {
      /* ignore */
    }
  }
  const durable = loadDurableCache(cacheKey);
  if (durable?.data) {
    try {
      return packReporteResponse({ ...durable.data, fromCache: true, cachedAt: durable.at });
    } catch {
      /* ignore */
    }
  }
  return null;
}

/**
 * Reporte liviano (filas planas + KPIs).
 * - Default: HOY Lima (si el proxy no manda fecha, el script usa hoy).
 * - fecha=YYYY-MM-DD | todas=true
 *
 * @param {{ fecha?: string, todas?: boolean, placa?: string, fundo?: string, variedad?: string, q?: string, limit?: number, force?: boolean }} opts
 */
export async function reporteTarjetas(opts = {}) {
  const force = Boolean(opts.force);
  const data = await callTarjetasApi(buildReporteParams(opts), { force });
  return packReporteResponse(data);
}

/** Fecha de hoy en America/Lima como yyyy-MM-dd */
export function limaTodayYmd() {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Lima",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(new Date());
}

export function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

export { emptyKpis };
