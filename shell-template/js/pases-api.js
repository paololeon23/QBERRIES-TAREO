/**
 * Cliente API Pases de salida (solo lectura).
 *
 * Proxy Netlify oficial (HTTPS, token solo en servidor).
 * NUNCA script.google.com ni API_TOKEN en el navegador.
 */

/** Proxy oficial Q Berries (lectura segura) */
export const PERMISOS_API_BASE = "https://pasessalida-qberries.netlify.app/api/permisos";

const REQUEST_TIMEOUT_MS = 18000;
const MAX_ATTEMPTS = 3;
const RETRY_DELAY_MS = 350;
/** Reutiliza respuesta reciente (misma query) — evita GET de más en el poll */
const SOFT_CACHE_MS = 10000;
/** Caché durable en localStorage si falla la red (seguridad / continuidad) */
const DURABLE_CACHE_KEY = "qb_permisos_cache_v1";
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
    // cuota / modo privado: ignorar
  }
}

function saveDurableCache(cacheKey, data) {
  if (!cacheKey || !data || data.ok === false) return;
  // No guardar pings
  if (!String(cacheKey).includes("listarPermisos")) return;
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
    return "No se pudo autenticar con el servidor de pases. Revisa la configuración del proxy Netlify.";
  }
  if (code === "NO_CONFIG" || code === "CONFIG") {
    return "Faltan variables de entorno en el proxy Netlify (configuración del servidor).";
  }
  return data.message || data.error || fallback;
}

function sleep(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function resolveApiBases() {
  if (typeof window !== "undefined" && window.QB_PERMISOS_PROXY) {
    return [String(window.QB_PERMISOS_PROXY)];
  }
  const bases = [];
  try {
    const host = window.location.hostname || "";
    // Si estamos en el mismo sitio del proxy → same-origin (más rápido, HTTPS)
    if (host === "pasessalida-qberries.netlify.app") {
      bases.push("/api/permisos");
    }
  } catch (_) {}
  bases.push(PERMISOS_API_BASE);
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
      throw new Error("El servidor de pases no devolvió JSON válido.");
    }
    if (data && data.ok === false) {
      throw new Error(apiErrorMessage(data, "Error en API de permisos"));
    }
    if (!response.ok) {
      throw new Error(apiErrorMessage(data, `API permisos respondió ${response.status}`));
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

/**
 * GET JSON al proxy con:
 * - soft-cache corto (salta con force:true)
 * - dedupe in-flight
 * - reintentos en fallos de red
 * - caché durable (localStorage) si se cae la conexión
 */
export async function callPermisosApi(params = {}, options = {}) {
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
      throw new Error("Sin conexión con el servidor de pases. Revisa tu red e intenta de nuevo.");
    }
    throw lastErr instanceof Error
      ? lastErr
      : new Error(String(lastErr || "No se pudo consultar permisos"));
  })();

  inflight.set(cacheKey, job);
  try {
    return await job;
  } finally {
    inflight.delete(cacheKey);
  }
}

export async function pingPermisos() {
  const data = await callPermisosApi({ action: "ping" });
  if (!data?.ok) throw new Error(apiErrorMessage(data, "Ping de permisos falló"));
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

  const d = new Date(raw);
  if (Number.isNaN(d.getTime())) return raw;
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Lima",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(d);
}

/** Hora America/Lima → HH:mm:ss (o deja AM/PM del formulario si ya es hora legible) */
export function toLimaHms(value) {
  if (value === undefined || value === null || value === "") return "";
  const raw = String(value).trim();
  if (/^\d{1,2}:\d{2}(:\d{2})?(\s*[AaPp][Mm])?$/.test(raw)) return raw;

  // Sheets: tiempos serializados como 1899-12-30T HH:mm:ss.000Z → usar reloj del ISO
  const sheetTime = raw.match(/^(\d{4})-\d{2}-\d{2}T(\d{2}):(\d{2}):(\d{2})/);
  if (sheetTime && (sheetTime[1] === "1899" || sheetTime[1] === "1900")) {
    return `${sheetTime[2]}:${sheetTime[3]}:${sheetTime[4]}`;
  }

  const d = new Date(raw);
  if (Number.isNaN(d.getTime())) return raw;
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: "America/Lima",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  }).format(d);
}

function normalizePaseRow(row) {
  if (!row || typeof row !== "object") return row;
  return {
    ...row,
    fechaRegistro: toLimaYmd(row.fechaRegistro),
    fechaSalida: toLimaYmd(row.fechaSalida),
    fIngreso: toLimaYmd(row.fIngreso),
    horaRegistro: toLimaHms(row.horaRegistro),
    horaSalida: row.horaSalida == null ? "" : String(row.horaSalida).trim(),
    dni: row.dni == null ? "" : String(row.dni),
    dniResponsable: row.dniResponsable == null ? "" : String(row.dniResponsable),
    carnetVerificado: row.carnetVerificado == null ? "" : String(row.carnetVerificado).trim(),
    carnetDniEscaneado: row.carnetDniEscaneado == null ? "" : String(row.carnetDniEscaneado),
    carnetVerificadoAt: row.carnetVerificadoAt
      ? (() => {
          const d = new Date(row.carnetVerificadoAt);
          if (Number.isNaN(d.getTime())) return String(row.carnetVerificadoAt);
          return d.toLocaleString("es-PE", { timeZone: "America/Lima" });
        })()
      : ""
  };
}

/**
 * KPIs del GET: count, motivos, topMotivo, porDia, rango
 */
export function extractPermisosKpis(payload = {}, rows = []) {
  const totalRaw = payload.count ?? payload.total ?? payload.totalSalidas;
  const total = Number.isFinite(Number(totalRaw)) ? Number(totalRaw) : rows.length;

  let motivos = [];
  if (Array.isArray(payload.motivos)) {
    motivos = payload.motivos
      .map((item) => {
        if (!item || typeof item !== "object") return null;
        const motivo = String(item.motivo ?? item.nombre ?? "").trim();
        const count = Number(item.count ?? item.total ?? 0);
        if (!motivo) return null;
        return { motivo, count: Number.isFinite(count) ? count : 0 };
      })
      .filter(Boolean);
  }

  let topMotivo = null;
  if (payload.topMotivo && typeof payload.topMotivo === "object") {
    const motivo = String(payload.topMotivo.motivo ?? "").trim();
    const count = Number(payload.topMotivo.count ?? 0);
    if (motivo) topMotivo = { motivo, count: Number.isFinite(count) ? count : 0 };
  }

  if (!topMotivo && motivos.length) {
    topMotivo = [...motivos].sort((a, b) => b.count - a.count || a.motivo.localeCompare(b.motivo, "es"))[0];
  }

  if (!motivos.length && rows.length) {
    const map = new Map();
    rows.forEach((r) => {
      const key = String(r.motivo || "").trim();
      if (!key) return;
      map.set(key, (map.get(key) || 0) + 1);
    });
    motivos = [...map.entries()].map(([motivo, count]) => ({ motivo, count }));
    topMotivo =
      [...motivos].sort((a, b) => b.count - a.count || a.motivo.localeCompare(b.motivo, "es"))[0] || null;
  }

  const porDia = Array.isArray(payload.porDia)
    ? payload.porDia.map((item) => ({
        fecha: toLimaYmd(item?.fecha) || String(item?.fecha || "").trim(),
        count: Number(item?.count) || 0
      }))
    : [];

  return {
    total,
    motivos,
    topMotivo: topMotivo || { motivo: "", count: 0 },
    porDia,
    rango: payload.rango || null,
    fromApi: Number.isFinite(Number(totalRaw)) || Array.isArray(payload.motivos) || Boolean(payload.topMotivo)
  };
}

function buildListarParams(opts = {}) {
  const params = { action: "listarPermisos" };

  if (opts.dni) params.dni = String(opts.dni).trim();

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
  return params;
}

function packListarResponse(data) {
  if (!data?.ok) throw new Error(apiErrorMessage(data, "No se pudieron listar los pases"));
  const rows = Array.isArray(data.data) ? data.data.map(normalizePaseRow) : [];
  const kpis = extractPermisosKpis(data, rows);
  return {
    ...data,
    data: rows,
    count: kpis.total,
    kpis,
    fromCache: Boolean(data.fromCache),
    cachedAt: data.cachedAt || null
  };
}

/**
 * Lee soft/durable cache sin red (pintado inmediato / stale-while-revalidate).
 * @returns {ReturnType<typeof packListarResponse>|null}
 */
export function peekCachedListarPermisos(opts = {}) {
  const params = buildListarParams(opts);
  const cacheKey = buildQuery(params);
  const soft = softCache.get(cacheKey);
  if (soft?.data?.ok !== false && soft?.data) {
    try {
      return packListarResponse({ ...soft.data, fromCache: true, cachedAt: soft.at });
    } catch {
      /* ignore */
    }
  }
  const durable = loadDurableCache(cacheKey);
  if (durable?.data) {
    try {
      return packListarResponse({ ...durable.data, fromCache: true, cachedAt: durable.at });
    } catch {
      /* ignore */
    }
  }
  return null;
}

/**
 * Listar pases.
 * - Default (sin opts): HOY Lima (el proxy filtra solo hoy).
 * - fecha=YYYY-MM-DD: un día.
 * - todas=true: todas las fechas.
 *
 * @param {{ limit?: number, dni?: string, fecha?: string, todas?: boolean }} opts
 */
export async function listarPermisos(opts = {}) {
  const force = Boolean(opts.force);
  const data = await callPermisosApi(buildListarParams(opts), { force });
  return packListarResponse(data);
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
