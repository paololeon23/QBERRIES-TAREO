/**
 * Service Worker siempre activo: registro, claim, warm de rutas/reglas, auto-update.
 */
import { appConfig } from "../config/app.config.js";
import { routesConfig } from "../config/routes.config.js";

const RULES_WARM = [
  "rules/modulos/uva-mp.rules.json",
  "rules/modulos/uva-pt.rules.json",
  "rules/modulos/uva-plagas.rules.json",
  "rules/modulos/arandano-mp-mpbar.rules.json",
  "rules/modulos/arandano-mp-mpgar.rules.json",
  "rules/modulos/arandano-mp-mpha.rules.json",
  "rules/modulos/arandano-pt-ptbpar.rules.json",
  "rules/modulos/arandano-pt-pthpar.rules.json",
  "rules/modulos/arandano-pt-ptlpar.rules.json",
  "rules/modulos/arandano-plagas.rules.json",
  "rules/modulos/esparrago-mp-mpes.rules.json",
  "rules/modulos/esparrago-pt.rules.json",
  "rules/modulos/esparrago-plagas.rules.json",
  "rules/modulos/palta-mp.rules.json",
  "rules/modulos/palta-pt.rules.json",
  "rules/modulos/palta-plagas.rules.json"
];

function collectWarmUrls() {
  const urls = new Set(RULES_WARM);
  Object.values(routesConfig || {}).forEach((route) => {
    if (route.viewPath) urls.add(route.viewPath);
    if (route.modulePath) urls.add(route.modulePath);
    (route.stylesheets || []).forEach((href) => urls.add(href));
  });
  [
    "presentation/js/services/module-loader.service.js",
    "presentation/js/services/i18n.service.js",
    "presentation/js/controllers/shell.controller.js",
    "presentation/js/modules/cultivos/shared/cartilla-analysis.js",
    "presentation/js/modules/cultivos/shared/mp-results-perf.util.js",
    "presentation/js/utils/ensure-xlsx.util.js",
    "presentation/data/sap-columnas.json"
  ].forEach((u) => urls.add(u));
  return [...urls];
}

function postToWorker(worker, message) {
  try {
    worker?.postMessage?.(message);
  } catch {
    /* ignore */
  }
}

async function warmCaches(registration) {
  if (!navigator.onLine) return;
  const worker = registration?.active || registration?.waiting || registration?.installing;
  if (!worker) return;
  postToWorker(worker, { type: "WARM_URLS", urls: collectWarmUrls() });
}

function bindKeepAlive(registration) {
  // Actualiza el SW en cada foco / cada hora para no quedar en versión vieja.
  const check = () => {
    registration.update().catch(() => {});
  };
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") check();
  });
  window.addEventListener("online", () => {
    check();
    warmCaches(registration);
  });
  window.setInterval(check, 60 * 60 * 1000);

  registration.addEventListener("updatefound", () => {
    const installing = registration.installing;
    if (!installing) return;
    installing.addEventListener("statechange", () => {
      if (installing.state === "installed" && navigator.serviceWorker.controller) {
        postToWorker(installing, { type: "SKIP_WAITING" });
      }
      if (installing.state === "activated") {
        warmCaches(registration);
      }
    });
  });
}

export function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;
  const { protocol } = window.location;
  if (protocol !== "https:" && protocol !== "http:") return;

  const swUrl = `./sw.js?v=${appConfig.cacheBustingVersion}`;

  const start = async () => {
    try {
      const registration = await navigator.serviceWorker.register(swUrl, { scope: "./" });
      bindKeepAlive(registration);

      // Esperar controlador activo y precalentar módulos/reglas.
      if (navigator.serviceWorker.controller) {
        await warmCaches(registration);
      } else {
        await new Promise((resolve) => {
          const onController = () => {
            navigator.serviceWorker.removeEventListener("controllerchange", onController);
            resolve();
          };
          navigator.serviceWorker.addEventListener("controllerchange", onController);
          // Si ya hay active sin controller aún (primera install)
          window.setTimeout(resolve, 4000);
        });
        await warmCaches(registration);
      }
    } catch (err) {
      console.warn("[AGV-MI] Service Worker:", err?.message || err);
      // Reintento: el SW debe quedar funcionando.
      window.setTimeout(start, 8000);
    }
  };

  // Registrar ya (no esperar load) para activarse más rápido.
  if (document.readyState === "complete") start();
  else window.addEventListener("load", start, { once: true });
}
