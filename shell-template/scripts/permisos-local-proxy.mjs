/**
 * Proxy LOCAL seguro para pruebas con Live Server.
 * Lee API_TOKEN y PERMISOS_SCRIPT_URL desde .env (nunca del navegador).
 *
 * Uso:
 *   1) Copia .env.example → .env y completa los valores
 *   2) node shell-template/scripts/permisos-local-proxy.mjs
 *   3) Sigue usando Live Server (127.0.0.1:5500)
 */

import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "../..");
const PORT = Number(process.env.PERMISOS_PROXY_PORT || 8787);

const ALLOWED_GET = new Set([
  "ping",
  "listarPermisos",
  "obtenerPermiso",
  "existePaseHoy"
]);

function loadEnvFile() {
  const envPath = path.join(ROOT, ".env");
  if (!fs.existsSync(envPath)) return;
  const text = fs.readFileSync(envPath, "utf8");
  text.split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) return;
    const i = trimmed.indexOf("=");
    if (i < 1) return;
    const key = trimmed.slice(0, i).trim();
    let val = trimmed.slice(i + 1).trim();
    if (
      (val.startsWith('"') && val.endsWith('"')) ||
      (val.startsWith("'") && val.endsWith("'"))
    ) {
      val = val.slice(1, -1);
    }
    if (!(key in process.env) || process.env[key] === "") {
      process.env[key] = val;
    }
  });
}

loadEnvFile();

const SCRIPT_URL = String(process.env.PERMISOS_SCRIPT_URL || "").trim();
const TOKEN = String(process.env.API_TOKEN || process.env.PERMISOS_API_TOKEN || "").trim();

function cors(req, res) {
  const origin = String(req.headers.origin || "");
  let allow = "http://127.0.0.1:5500";
  try {
    if (origin) {
      const host = new URL(origin).hostname;
      if (host === "localhost" || host === "127.0.0.1") allow = origin;
    }
  } catch (_) {}
  res.setHeader("Access-Control-Allow-Origin", allow);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Accept");
  res.setHeader("Cache-Control", "no-store");
}

function sendJson(req, res, status, obj) {
  cors(req, res);
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(obj));
}

const server = http.createServer(async (req, res) => {
  cors(req, res);

  if (req.method === "OPTIONS") {
    res.writeHead(204);
    res.end();
    return;
  }

  const url = new URL(req.url || "/", `http://127.0.0.1:${PORT}`);
  if (url.pathname !== "/api/permisos" && url.pathname !== "/api/permisos/") {
    sendJson(req, res, 404, { ok: false, message: "Usa /api/permisos" });
    return;
  }

  if (req.method !== "GET") {
    sendJson(req, res, 405, { ok: false, message: "Solo GET en proxy local" });
    return;
  }

  if (!SCRIPT_URL || !TOKEN) {
    sendJson(req, res, 500, {
      ok: false,
      code: "CONFIG",
      message:
        "Falta .env con API_TOKEN y PERMISOS_SCRIPT_URL en la raíz del repo (copia .env.example)."
    });
    return;
  }

  const action = String(url.searchParams.get("action") || "ping").trim();
  if (!ALLOWED_GET.has(action)) {
    sendJson(req, res, 403, {
      ok: false,
      code: "FORBIDDEN",
      message: "Acción no permitida en el proxy local"
    });
    return;
  }

  const qs = new URLSearchParams();
  url.searchParams.forEach((value, key) => {
    if (/^(callback|token|apitoken|api_token)$/i.test(key)) return;
    if (value === "") return;
    qs.set(key, value);
  });
  qs.set("token", TOKEN);

  try {
    const upstream = await fetch(`${SCRIPT_URL}?${qs.toString()}`, {
      method: "GET",
      headers: { Accept: "application/json" },
      redirect: "follow"
    });
    const text = await upstream.text();
    cors(req, res);
    res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
    res.end(text);
  } catch (err) {
    sendJson(req, res, 502, {
      ok: false,
      code: "PROXY",
      message: String(err && err.message ? err.message : err)
    });
  }
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`[permisos-local-proxy] http://127.0.0.1:${PORT}/api/permisos`);
  console.log(`[permisos-local-proxy] Live Server → usa este proxy (token solo en .env)`);
  if (!SCRIPT_URL || !TOKEN) {
    console.warn("[permisos-local-proxy] Configura .env antes de probar (ver .env.example)");
  }
});
