/**
 * AGV-MI — Service Worker siempre activo (offline fiable).
 * Cache-first same-origin · precache shell · warm bajo demanda.
 * No cachea: Gemini, Cloudinary, Google Fonts, jsDelivr.
 * Mantener CACHE_VERSION = appConfig.cacheBustingVersion.
 */
/* eslint-disable no-restricted-globals */

const CACHE_VERSION = "2026080501";
const SHELL_CACHE = `agv-mi-shell-${CACHE_VERSION}`;
const RUNTIME_CACHE = `agv-mi-runtime-${CACHE_VERSION}`;

const PRECACHE_URLS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./presentation/images/icono.png",
  "./presentation/images/logo.png",
  "./presentation/css/base/tokens.css",
  "./presentation/css/base/lucide-icons.css",
  "./presentation/css/base/reset.css",
  "./presentation/css/layout/app-shell.css",
  "./presentation/css/components/sidebar.css",
  "./presentation/css/components/topbar.css",
  "./presentation/css/components/ui/buttons.css",
  "./presentation/css/components/ui/language-selector.css",
  "./presentation/css/components/ui/skeleton.css",
  "./presentation/css/pages/bootstrap-error.css",
  "./presentation/css/pages/hero.css",
  "./presentation/css/components/confidentiality-gate.css",
  "./presentation/css/utilities/animations.css",
  "./presentation/css/utilities/display-quality.css",
  "./presentation/vendor/luxon.min.js",
  "./presentation/vendor/i18next.min.js",
  "./presentation/vendor/lucide.min.js",
  "./presentation/vendor/xlsx.full.min.js",
  "./presentation/js/main.js",
  "./presentation/js/config/app.config.js",
  "./presentation/js/config/routes.config.js",
  "./presentation/js/core/router.js",
  "./presentation/js/core/state-store.js",
  "./presentation/i18n/es-PE.json",
  "./presentation/i18n/en-US.json",
  "./presentation/i18n/fr-MA.json",
  "./presentation/i18n/zh-CN.json",
  "./engine/rule-engine.js",
  "./engine/cartilla-cell-validation.js",
  "./engine/cartilla-rules.adapter.js",
  "./engine/pt-rule-validators.js"
];

const BYPASS_HOSTS = [
  "generativelanguage.googleapis.com",
  "fonts.googleapis.com",
  "fonts.gstatic.com",
  "cdn.jsdelivr.net",
  "res.cloudinary.com"
];

function shouldBypass(url) {
  return BYPASS_HOSTS.some((host) => url.hostname === host || url.hostname.endsWith(`.${host}`));
}

function isSameOriginAsset(url) {
  if (url.origin !== self.location.origin) return false;
  const p = url.pathname;
  return (
    p.startsWith("/presentation/") ||
    p.startsWith("/rules/") ||
    p.startsWith("/engine/") ||
    p.startsWith("/ingestion/") ||
    p === "/" ||
    p.endsWith("/index.html") ||
    p.endsWith("/manifest.webmanifest")
  );
}

async function putOk(cache, request, response) {
  if (response && response.ok) {
    try {
      await cache.put(request, response.clone());
    } catch {
      /* quota / opaque */
    }
  }
  return response;
}

async function matchAny(request) {
  return (
    (await caches.match(request)) ||
    (await caches.match(request, { ignoreSearch: true })) ||
    null
  );
}

async function cacheUrls(urls, cacheName = RUNTIME_CACHE) {
  const cache = await caches.open(cacheName);
  await Promise.all(
    (urls || []).map(async (raw) => {
      const url = String(raw || "").trim();
      if (!url || url.startsWith("http")) return;
      try {
        const res = await fetch(url, { cache: "reload" });
        if (res.ok) await cache.put(url, res.clone());
      } catch {
        /* skip */
      }
    })
  );
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    (async () => {
      await cacheUrls(PRECACHE_URLS, SHELL_CACHE);
      await self.skipWaiting();
    })()
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(
        keys
          .filter((key) => key.startsWith("agv-mi-") && key !== SHELL_CACHE && key !== RUNTIME_CACHE)
          .map((key) => caches.delete(key))
      );
      await self.clients.claim();
    })()
  );
});

self.addEventListener("message", (event) => {
  const data = event.data || {};
  if (data.type === "SKIP_WAITING") {
    self.skipWaiting();
    return;
  }
  if (data.type === "WARM_URLS" && Array.isArray(data.urls)) {
    event.waitUntil(cacheUrls(data.urls, RUNTIME_CACHE));
  }
});

async function networkFirstNavigation(request) {
  try {
    const fresh = await fetch(request);
    const cache = await caches.open(SHELL_CACHE);
    await putOk(cache, "./index.html", fresh);
    return fresh;
  } catch {
    const cached =
      (await matchAny("./index.html")) ||
      (await matchAny("index.html")) ||
      (await matchAny("./")) ||
      (await matchAny(request));
    if (cached) return cached;
    return new Response(
      "<!doctype html><meta charset=utf-8><title>AGV-MI</title><p>AGV-MI offline. Abre una vez con internet para activar la caché.</p>",
      { status: 200, headers: { "Content-Type": "text/html; charset=utf-8" } }
    );
  }
}

/** Cache-first: offline siempre responde si ya se visitó el recurso. */
async function cacheFirst(request) {
  const cached = await matchAny(request);
  if (cached) {
    fetch(request)
      .then(async (res) => {
        if (!res || !res.ok) return;
        const cache = await caches.open(RUNTIME_CACHE);
        await putOk(cache, request, res);
      })
      .catch(() => {});
    return cached;
  }

  try {
    const fresh = await fetch(request);
    const cache = await caches.open(RUNTIME_CACHE);
    await putOk(cache, request, fresh);
    return fresh;
  } catch {
    return new Response("Recurso no disponible offline", {
      status: 503,
      headers: { "Content-Type": "text/plain; charset=utf-8" }
    });
  }
}

self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return;

  let url;
  try {
    url = new URL(request.url);
  } catch {
    return;
  }

  if (shouldBypass(url)) return;

  if (request.mode === "navigate") {
    event.respondWith(networkFirstNavigation(request));
    return;
  }

  if (!isSameOriginAsset(url)) return;
  if (url.pathname.endsWith("/sw.js")) return;

  event.respondWith(cacheFirst(request));
});
