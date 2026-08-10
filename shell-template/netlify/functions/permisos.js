/**
 * Proxy seguro → Google Apps Script (API permisos).
 *
 * - API_TOKEN y PERMISOS_SCRIPT_URL SOLO en env de Netlify (nunca en el JS).
 * - El navegador llama /api/permisos; esta función inyecta el token en el servidor.
 * - Panel shell: solo lectura. Acciones de escritura (crearPermiso) están bloqueadas aquí.
 * - CORS: mismo sitio + localhost/127.0.0.1 (Live Server de pruebas).
 *
 * Env:
 *   PERMISOS_SCRIPT_URL = https://script.google.com/macros/s/.../exec
 *   API_TOKEN            = mismo valor que Apps Script → Propiedades → API_TOKEN
 */

const ALLOWED_GET = new Set([
  "ping",
  "listarPermisos",
  "obtenerPermiso",
  "existePaseHoy"
]);

const ALLOWED_POST = new Set([
  "listarPermisos"
  // crearPermiso / guardarPermiso: NO en el proxy público del panel
]);

function getHeader(event, name) {
  const headers = event.headers || {};
  const want = name.toLowerCase();
  for (const key of Object.keys(headers)) {
    if (String(key).toLowerCase() === want) return String(headers[key] || "").trim();
  }
  return "";
}

function isLocalDevOrigin(origin) {
  if (!origin) return false;
  try {
    const o = new URL(origin);
    return o.hostname === "localhost" || o.hostname === "127.0.0.1";
  } catch {
    return false;
  }
}

function isAllowedOrigin(event) {
  const origin = getHeader(event, "origin");
  const host = getHeader(event, "host");
  if (!origin) return true; // same-origin navegación / curl
  if (isLocalDevOrigin(origin)) return true;
  try {
    return new URL(origin).host === host;
  } catch {
    return false;
  }
}

function corsHeaders(event) {
  const origin = getHeader(event, "origin");
  if (!origin) return {};
  if (!isLocalDevOrigin(origin)) {
    // same-site no necesita CORS; si viene Origin raro no lo reflejamos
    try {
      const host = getHeader(event, "host");
      if (new URL(origin).host === host) {
        return {
          "Access-Control-Allow-Origin": origin,
          Vary: "Origin"
        };
      }
    } catch (_) {}
    return {};
  }
  return {
    "Access-Control-Allow-Origin": origin,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Accept",
    Vary: "Origin"
  };
}

function jsonResponse(event, statusCode, obj) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "X-Content-Type-Options": "nosniff",
      ...corsHeaders(event)
    },
    body: JSON.stringify(obj)
  };
}

function textResponse(event, statusCode, text) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "X-Content-Type-Options": "nosniff",
      ...corsHeaders(event)
    },
    body: text
  };
}

function scriptConfig() {
  const scriptUrl = String(process.env.PERMISOS_SCRIPT_URL || "").trim();
  const token = String(process.env.API_TOKEN || process.env.PERMISOS_API_TOKEN || "").trim();
  return { scriptUrl, token };
}

/** Solo params seguros; nunca reenviar token/callback del cliente. */
function mergeQuery(event) {
  const qs = new URLSearchParams();
  const params = event.queryStringParameters || {};
  Object.keys(params).forEach((key) => {
    const k = String(key);
    if (/^(callback|token|apitoken|api_token)$/i.test(k)) return;
    const val = params[key];
    if (val === undefined || val === null || val === "") return;
    qs.set(k, String(val));
  });
  return qs;
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") {
    return {
      statusCode: 204,
      headers: {
        "Cache-Control": "no-store",
        ...corsHeaders(event)
      },
      body: ""
    };
  }

  if (!isAllowedOrigin(event)) {
    return jsonResponse(event, 403, {
      ok: false,
      code: "FORBIDDEN",
      message: "Origen no permitido"
    });
  }

  const { scriptUrl, token } = scriptConfig();
  if (!scriptUrl) {
    return jsonResponse(event, 500, {
      ok: false,
      code: "CONFIG",
      message: "Falta PERMISOS_SCRIPT_URL en variables de entorno de Netlify"
    });
  }
  if (!token) {
    return jsonResponse(event, 500, {
      ok: false,
      code: "CONFIG",
      message: "Falta API_TOKEN en variables de entorno de Netlify"
    });
  }

  try {
    if (event.httpMethod === "GET") {
      const qs = mergeQuery(event);
      const action = String(qs.get("action") || "ping").trim();
      if (!ALLOWED_GET.has(action)) {
        return jsonResponse(event, 403, {
          ok: false,
          code: "FORBIDDEN",
          message: "Acción no permitida en el proxy"
        });
      }
      qs.set("token", token);
      const url = `${scriptUrl}?${qs.toString()}`;
      const res = await fetch(url, {
        method: "GET",
        headers: { Accept: "application/json" },
        redirect: "follow"
      });
      const text = await res.text();
      return textResponse(event, 200, text);
    }

    if (event.httpMethod === "POST") {
      let body = {};
      try {
        body = JSON.parse(event.body || "{}");
      } catch {
        body = {};
      }
      if (!body || typeof body !== "object") body = {};

      const action = String(body.action || "").trim();
      if (!ALLOWED_POST.has(action)) {
        return jsonResponse(event, 403, {
          ok: false,
          code: "FORBIDDEN",
          message: "Escritura no permitida desde el panel público"
        });
      }

      delete body.token;
      delete body.apiToken;
      delete body.API_TOKEN;
      body.token = token;

      const res = await fetch(scriptUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json"
        },
        body: JSON.stringify(body),
        redirect: "follow"
      });
      const text = await res.text();
      return textResponse(event, 200, text);
    }

    return jsonResponse(event, 405, { ok: false, message: "Método no permitido" });
  } catch (err) {
    return jsonResponse(event, 502, {
      ok: false,
      code: "PROXY",
      message: String(err && err.message ? err.message : err)
    });
  }
};
