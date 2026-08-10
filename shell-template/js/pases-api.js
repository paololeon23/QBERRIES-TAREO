/**
 * Cliente API Pases de salida (solo lectura).
 *
 * Seguridad: el navegador NUNCA ve API_TOKEN.
 *
 * Pruebas con Live Server:
 *   1) .env en la raíz (API_TOKEN + PERMISOS_SCRIPT_URL)
 *   2) node shell-template/scripts/permisos-local-proxy.mjs
 *   3) Live Server en 5500
 *
 * En Netlify / netlify dev: /api/permisos (Function).
 */

/** Ruta relativa (Netlify / netlify dev). */
export const PERMISOS_PROXY_PATH = "/api/permisos";

/** Proxy local seguro (script permisos-local-proxy.mjs). */
export const PERMISOS_PROXY_LOCAL = "http://127.0.0.1:8787/api/permisos";

/** Proxy publicado (solo tras deploy de la Function). */
export const PERMISOS_PROXY_REMOTE = "https://qberries-produccion.netlify.app/api/permisos";

function buildQuery(params = {}) {
  const qs = new URLSearchParams();
  Object.entries(params).forEach(([key, value]) => {
    if (value === undefined || value === null || value === "") return;
    if (value === false) return;
    qs.set(key, String(value));
  });
  return qs.toString();
}

function apiErrorMessage(data, fallback) {
  if (!data || typeof data !== "object") return fallback;
  return data.message || data.error || fallback;
}

function isLocalHost() {
  const host = window.location.hostname;
  if (!host || host === "localhost" || host === "127.0.0.1") return true;
  if (window.location.protocol === "file:") return true;
  return false;
}

/** Live Server típico: no tiene Functions → no molestar con 404 de /api. */
function isLiveServer() {
  if (!isLocalHost()) return false;
  const port = String(window.location.port || "");
  return port === "5500" || port === "5501" || port === "5502";
}

function proxyCandidates() {
  if (typeof window !== "undefined" && window.QB_PERMISOS_PROXY) {
    return [String(window.QB_PERMISOS_PROXY)];
  }
  if (isLiveServer()) {
    return [PERMISOS_PROXY_LOCAL, PERMISOS_PROXY_REMOTE];
  }
  if (isLocalHost()) {
    // netlify dev u otro puerto local
    return [PERMISOS_PROXY_PATH, PERMISOS_PROXY_LOCAL, PERMISOS_PROXY_REMOTE];
  }
  return [PERMISOS_PROXY_PATH];
}

async function fetchViaProxyUrl(basePath, params = {}) {
  const qs = buildQuery(params);
  const url = qs ? `${basePath}?${qs}` : basePath;
  let response;
  try {
    response = await fetch(url, {
      method: "GET",
      headers: { Accept: "application/json" },
      cache: "no-store"
    });
  } catch (err) {
    const e = new Error("PROXY_UNREACHABLE");
    e.cause = err;
    throw e;
  }

  if (response.status === 404 || response.status === 503 || response.status >= 500) {
    throw new Error(response.status === 404 ? "PROXY_404" : "PROXY_5XX");
  }

  const text = await response.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch {
    // Netlify a veces responde el HTML del SPA si la Function no está desplegada
    throw new Error("PROXY_NOT_JSON");
  }

  if (data?.ok === false) {
    throw new Error(apiErrorMessage(data, "Error en API de permisos"));
  }
  if (!response.ok) {
    throw new Error(`API permisos respondió ${response.status}`);
  }
  return data;
}

/**
 * Solo proxies (token never in browser).
 * Live Server → proxy local :8787 (recomendado) o Netlify si ya hay Function.
 */
export async function callPermisosApi(params = {}) {
  const bases = proxyCandidates();
  let lastErr = null;

  for (const base of bases) {
    try {
      return await fetchViaProxyUrl(base, params);
    } catch (err) {
      lastErr = err;
      const code = err?.message || "";
      if (
        code === "PROXY_404" ||
        code === "PROXY_UNREACHABLE" ||
        code === "PROXY_NOT_JSON" ||
        code === "PROXY_5XX"
      ) {
        continue;
      }
      throw err;
    }
  }

  const hint = isLiveServer()
    ? "En Live Server corre: node shell-template/scripts/permisos-local-proxy.mjs (con .env). El token NO va en la app."
    : "Publica la Function en Netlify o usa netlify dev. El token solo va en variables de entorno.";

  throw new Error(lastErr?.message?.startsWith("PROXY_") ? hint : lastErr?.message || hint);
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

/** Hora America/Lima → HH:mm:ss */
export function toLimaHms(value) {
  if (value === undefined || value === null || value === "") return "";
  const raw = String(value).trim();
  if (/^\d{1,2}:\d{2}(:\d{2})?(\s*[AaPp][Mm])?$/.test(raw)) return raw;

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
    dni: row.dni == null ? "" : String(row.dni),
    dniResponsable: row.dniResponsable == null ? "" : String(row.dniResponsable),
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
 * Extrae KPIs del GET:
 * { ok, count, motivos, topMotivo, porDia, data, rango }
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
  } else if (payload.conteoPorMotivo && typeof payload.conteoPorMotivo === "object") {
    motivos = Object.entries(payload.conteoPorMotivo).map(([motivo, count]) => ({
      motivo: String(motivo).trim(),
      count: Number(count) || 0
    }));
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

  const porDia = Array.isArray(payload.porDia) ? payload.porDia : [];

  return {
    total,
    motivos,
    topMotivo: topMotivo || { motivo: "", count: 0 },
    porDia,
    rango: payload.rango || null,
    fromApi: Number.isFinite(Number(totalRaw)) || Array.isArray(payload.motivos) || Boolean(payload.topMotivo)
  };
}

/**
 * Listado alineado con Code.gs:
 * - sin fecha → todas=1
 * - con fecha yyyy-MM-dd → filtro día
 *
 * @param {{ limit?: number, dni?: string, fecha?: string, todas?: boolean }} opts
 */
export async function listarPermisos(opts = {}) {
  const params = {
    action: "listarPermisos"
  };

  if (opts.dni) params.dni = String(opts.dni).trim();

  const fecha = opts.fecha != null ? String(opts.fecha).trim() : "";
  const todasExplicit = opts.todas === true || opts.todas === 1 || opts.todas === "1";

  if (todasExplicit || !fecha) {
    params.todas = "1";
    params.limit = opts.limit ?? 2000;
  } else {
    params.fecha = fecha;
    params.limit = opts.limit ?? 500;
  }

  const data = await callPermisosApi(params);
  if (!data?.ok) throw new Error(apiErrorMessage(data, "No se pudieron listar los pases"));

  const rows = Array.isArray(data.data) ? data.data.map(normalizePaseRow) : [];
  const kpis = extractPermisosKpis(data, rows);
  return {
    ...data,
    data: rows,
    count: kpis.total,
    kpis
  };
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
