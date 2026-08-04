import { appConfig } from "./config/app.config.js";
import { routesConfig } from "./config/routes.config.js";
import { initCropHectaresData } from "./config/crop-hectares.registry.js?v=20260800";
import { i18nService } from "./services/i18n.service.js";
import {
  renderApplicationShell,
  shellController,
  updateTopbarDatetime,
  updateTopbarTitle,
  updateActiveSidebarLink,
  updateBreadcrumbModule
} from "./controllers/shell.controller.js";
import { router } from "./core/router.js";
import { stateStore } from "./core/state-store.js";
import { moduleLoaderService } from "./services/module-loader.service.js";
import { hydrateLucideIcons } from "./utils/lucide-icon.util.js";
import { applyTranslationsToContainer } from "./utils/i18n-dom.util.js";
import { showConfidentialityGate } from "./components/confidentiality-gate.js";
import { getAppDocumentTitle, applyBrandPixelAssets } from "./utils/brand-pixel.util.js";
import { ensureStylesheets } from "./utils/ensure-stylesheet.util.js";
import { prefetchXlsxJs } from "./utils/ensure-xlsx.util.js";
import { registerServiceWorker } from "./utils/register-sw.util.js";

document.title = getAppDocumentTitle();
registerServiceWorker();

function renderBootstrapError(errorMessage) {
  const applicationRoot = document.getElementById("applicationRoot");
  applicationRoot.innerHTML = `
    <div class="bootstrap-error">
      <h1 class="bootstrap-error__title">${appConfig.appName}</h1>
      <p class="bootstrap-error__message">Error al iniciar la aplicación</p>
      <pre class="bootstrap-error__detail">${errorMessage}</pre>
      <p class="bootstrap-error__hint">Abre el proyecto con Live Server o cualquier servidor estático local.</p>
    </div>
  `;
}

function waitForGlobals(timeoutMs = 8000) {
  const start = Date.now();
  return new Promise((resolve, reject) => {
    const tick = () => {
      if (window.luxon && window.lucide && window.i18next) {
        resolve();
        return;
      }
      if (Date.now() - start > timeoutMs) {
        reject(new Error("Luxon / Lucide / i18next no cargaron a tiempo."));
        return;
      }
      requestAnimationFrame(tick);
    };
    tick();
  });
}

/**
 * Cambia idioma sin destruir el módulo activo (tabla/Excel cargados).
 * Solo reconstruye chrome (sidebar/topbar/footer) y re-traduce la vista actual.
 */
async function handleLanguageChange(selectedLanguage) {
  await i18nService.loadLanguage(selectedLanguage);
  stateStore.set({ currentLanguage: selectedLanguage });
  document.documentElement.lang = selectedLanguage.split("-")[0];

  const applicationRoot = document.getElementById("applicationRoot");
  const moduleInner = document.getElementById("dynamicModuleInner");
  const stash = document.createDocumentFragment();
  if (moduleInner) stash.appendChild(moduleInner);

  shellController.destroy();
  applicationRoot.innerHTML = renderApplicationShell(selectedLanguage);

  const freshInner = document.getElementById("dynamicModuleInner");
  const preserved = stash.firstChild;
  if (preserved && freshInner) {
    freshInner.replaceWith(preserved);
  }

  shellController.initialize(handleLanguageChange);
  hydrateLucideIcons(document.getElementById("sidebarNavigation"));
  hydrateLucideIcons(document.getElementById("applicationTopbar"));
  hydrateLucideIcons(document.getElementById("applicationFooter"));
  updateTopbarDatetime(selectedLanguage);

  const currentHash = stateStore.get().currentRoute || appConfig.defaultRoute;
  const routeDefinition =
    routesConfig[currentHash] ?? routesConfig[appConfig.defaultRoute];
  updateTopbarTitle(routeDefinition.titleKey);
  updateActiveSidebarLink(currentHash);
  updateBreadcrumbModule(currentHash);
  applyBrandPixelAssets();

  const activeModule = document.getElementById("dynamicModuleInner");
  applyTranslationsToContainer(activeModule, { hydrateIcons: false });
  await moduleLoaderService.applyLanguageChange(selectedLanguage);
  window.dispatchEvent(
    new CustomEvent("agv:language-changed", { detail: { language: selectedLanguage } })
  );
}

async function renderApplicationShellAndBind(languageCode) {
  shellController.destroy();
  const applicationRoot = document.getElementById("applicationRoot");
  applicationRoot.innerHTML = renderApplicationShell(languageCode);
  shellController.initialize(handleLanguageChange);
  hydrateLucideIcons(applicationRoot);
}

async function bootstrapApplication() {
  await waitForGlobals();

  // Flags + i18n en paralelo; hectáreas NO bloquean el primer paint.
  const hectaresPromise = initCropHectaresData().catch((err) => {
    console.warn("[AGV-MI] Hectáreas diferidas:", err);
  });

  await Promise.all([
    i18nService.initialize(appConfig.defaultLanguage),
    ensureStylesheets([
      "https://cdn.jsdelivr.net/npm/flag-icons@7.2.3/css/flag-icons.min.css",
      "presentation/css/pages/esparrago-pt.css"
    ])
  ]);

  stateStore.set({ currentLanguage: appConfig.defaultLanguage });
  document.documentElement.lang = appConfig.defaultLanguage.split("-")[0];

  await renderApplicationShellAndBind(appConfig.defaultLanguage);
  updateTopbarDatetime(appConfig.defaultLanguage);
  applyBrandPixelAssets();

  // Pintar Inicio de inmediato; hectáreas terminan en segundo plano.
  const routerPromise = router.start();
  // SheetJS en paralelo (listo antes de entrar a MP/PT/Plagas).
  prefetchXlsxJs();
  await Promise.all([routerPromise, hectaresPromise]);
  await showConfidentialityGate();
}

bootstrapApplication().catch((bootstrapError) => {
  console.error(`[${appConfig.appName}] Bootstrap error:`, bootstrapError);
  renderBootstrapError(bootstrapError.message);
});
