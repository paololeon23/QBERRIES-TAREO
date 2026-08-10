/**
 * Cliente API Pases de salida (solo lectura).
 *
 * Usa ÚNICAMENTE el proxy Netlify (token oculto en servidor).
 * NUNCA script.google.com ni API_TOKEN en el navegador.
 */

/** Proxy oficial Q Berries */
export const PERMISOS_API_BASE = "https://pasessalida-qberries.netlify.app/api/permisos";

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
  const code = String(data.code || "").toUpperCase();
  if (code === "UNAUTHORIZED") {
    return "No se pudo autenticar con el servidor de pases. Revisa la configuración del proxy Netlify.";
  }
  if (code === "NO_CONFIG" || code === "CONFIG") {
    return "Faltan variables de entorno en el proxy Netlify (configuración del servidor).";
  }
  return data.message || data.error || fallback;
}

/**
 * GET JSON al proxy (también acepta el mismo shape si el servidor responde error).
 */
export async function callPermisosApi(params = {}) {
  const qs = buildQuery(params);
  const url = qs ? `${PERMISOS_API_BASE}?${qs}` : PERMISOS_API_BASE;

  let response;
  try {
    response = await fetch(url, {
      method: "GET",
      headers: { Accept: "application/json" },
      cache: "no-store"
    });
  } catch {
    throw new Error("Sin conexión con el servidor de pases. Revisa tu red e intenta de nuevo.");
  }

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

/**
 * Listar pases.
 * - Default (sin opts): HOY Lima (el proxy filtra solo hoy).
 * - fecha=YYYY-MM-DD: un día.
 * - todas=true: todas las fechas.
 *
 * @param {{ limit?: number, dni?: string, fecha?: string, todas?: boolean }} opts
 */
export async function listarPermisos(opts = {}) {
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
    // Sin fecha ni todas → el proxy usa HOY (America/Lima)
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
