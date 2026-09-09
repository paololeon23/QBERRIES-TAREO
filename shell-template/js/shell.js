const ROUTE_TITLES = {
  inicio: "Inicio",
  tareo: "Tareo",
  "por-hora": "Tareo",
  validacion: "Tareo",
  historial: "Historial",
  recomendaciones: "Recomendaciones",
  "pases-de-salida": "Pases de salida",
  "tarjeta-pallet": "Tarjeta Pallet",
  produccion: "Producción",
  "reporte-trabajadores": "Reporte de Trabajadores",
  "trabajadores-jarras": "Trabajadores Jarras",
  "unificacion-jarras": "Unificación de jarras"
};

const ROUTE_VIEW = {
  inicio: "inicio",
  tareo: "tareo",
  "por-hora": "tareo",
  validacion: "tareo",
  historial: "historial",
  recomendaciones: "recomendaciones",
  "pases-de-salida": "pases-de-salida",
  "tarjeta-pallet": "tarjeta-pallet",
  produccion: "produccion",
  "reporte-trabajadores": "reporte-trabajadores",
  "trabajadores-jarras": "trabajadores-jarras",
  "unificacion-jarras": "unificacion-jarras"
};

function normalizeHash(hash) {
  const value = (hash || "#/inicio").replace(/^#\/?/, "");
  let key = value.split("/")[0] || "inicio";
  if (key === "" || key === "/") key = "inicio";
  if (key === "validacion" || key === "por-hora") key = "tareo";
  return ROUTE_TITLES[key] ? key : "inicio";
}

export function getCurrentRoute() {
  return normalizeHash(window.location.hash);
}

export function applyRoute(routeKey) {
  const route = ROUTE_TITLES[routeKey] ? routeKey : "inicio";
  const viewId = ROUTE_VIEW[route] || "inicio";

  document.querySelectorAll(".view").forEach((view) => {
    const match = view.dataset.view === viewId;
    view.hidden = !match;
    view.classList.toggle("is-active", match);
  });

  document.querySelectorAll(".sidebar-nav-link[data-route]").forEach((link) => {
    let linkRoute = link.dataset.route;
    if (linkRoute === "validacion" || linkRoute === "por-hora") linkRoute = "tareo";
    link.classList.toggle("is-active", linkRoute === viewId || link.dataset.route === route);
  });

  document.querySelectorAll(".sidebar__primary-panel, .sidebar__secondary-panel").forEach((panel) => {
    const active = Boolean(panel.querySelector(".sidebar-nav-link.is-active"));
    panel.classList.toggle("is-route-active", active);
  });

  const title = ROUTE_TITLES[route] || ROUTE_TITLES[viewId] || "Inicio";
  const titleEl = document.getElementById("txtPageTitle");
  const crumbEl = document.getElementById("txtBreadcrumbActive");
  if (titleEl) titleEl.textContent = title;
  if (crumbEl) crumbEl.textContent = title;

  window.dispatchEvent(new CustomEvent("qb:route-changed", { detail: { route: viewId } }));
}

function ensureProduccionNav() {
  if (document.querySelector('.sidebar-nav-link[data-route="produccion"]')) return;
  const tareo = document.querySelector('.sidebar-nav-link[data-route="tareo"]');
  const panel =
    tareo?.closest(".sidebar-panel__links") ||
    tareo?.closest(".sidebar__secondary-panel") ||
    document.querySelector('.sidebar__secondary-panel[data-panel="campo"] .sidebar-panel__links') ||
    document.querySelector(".sidebar__secondary-panel");
  if (!panel) return;
  const link = document.createElement("a");
  link.className = "sidebar-nav-link";
  link.href = "#/produccion";
  link.dataset.route = "produccion";
  link.dataset.sidebarTooltip = "Producción";
  link.innerHTML = `
    <span class="sidebar-nav-link__icon" aria-hidden="true">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M7 14v4"/><path d="M12 10v8"/><path d="M17 6v12"/></svg>
    </span>
    <span class="sidebar-nav-link__text">Producción</span>`;
  if (tareo?.nextSibling) panel.insertBefore(link, tareo.nextSibling);
  else panel.appendChild(link);
}

(() => {
  if (window.__QB_SHELL_INIT__) return;
  window.__QB_SHELL_INIT__ = true;

  const shell = document.getElementById("applicationRoot");
  const collapseBtn = document.getElementById("btnSidebarCollapse");
  const searchInput = document.getElementById("txtSidebarSearch");
  const backdrop = document.getElementById("sidebarDrawerBackdrop");
  const mobileMq = window.matchMedia("(max-width: 768px)");

  ensureProduccionNav();

  // QBerries: quitar cualquier Service Worker del origen (AGV cacheaba HTML viejo)
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.getRegistrations?.().then((regs) => {
      regs.forEach((reg) => {
        reg.unregister().catch(() => {});
      });
    });
    caches.keys?.().then((keys) => {
      keys
        .filter((k) => String(k).startsWith("agv-mi-"))
        .forEach((k) => caches.delete(k).catch(() => {}));
    });
  }

  function clearSidebarSearchFilter() {
    if (searchInput) searchInput.value = "";
    document.querySelectorAll(".sidebar-nav-link").forEach((el) => {
      el.hidden = false;
    });
  }

  function isMobile() {
    return mobileMq.matches;
  }

  function syncBackdrop(open) {
    if (!backdrop) return;
    backdrop.hidden = !open;
    backdrop.setAttribute("aria-hidden", open ? "false" : "true");
  }

  function syncCollapseButton(collapsed) {
    if (!collapseBtn) return;
    collapseBtn.classList.toggle("is-expand-state", collapsed);
    const label = collapsed ? "Expandir menú" : "Colapsar menú";
    collapseBtn.setAttribute("aria-label", label);
    collapseBtn.setAttribute("title", label);
    collapseBtn.setAttribute("aria-expanded", collapsed ? "false" : "true");
  }

  function setCollapsed(collapsed) {
    if (!shell) return;
    const next = Boolean(collapsed);

    if (isMobile()) {
      shell.classList.remove("is-sidebar-collapsed");
      if (next) {
        shell.classList.remove("is-sidebar-drawer-open");
        document.body.classList.remove("is-sidebar-drawer-open");
        syncBackdrop(false);
      } else {
        shell.classList.add("is-sidebar-drawer-open");
        document.body.classList.add("is-sidebar-drawer-open");
        syncBackdrop(true);
      }
      syncCollapseButton(next);
      return;
    }

    syncBackdrop(false);
    shell.classList.remove("is-sidebar-drawer-open");
    document.body.classList.remove("is-sidebar-drawer-open");
    shell.classList.add("is-sidebar-collapsing");
    shell.classList.toggle("is-sidebar-collapsed", next);
    syncCollapseButton(next);
    window.setTimeout(() => shell.classList.remove("is-sidebar-collapsing"), 220);
  }

  function closeAllPanelFlyouts() {
    document.querySelectorAll(".sidebar__primary-panel.is-flyout-open, .sidebar__secondary-panel.is-flyout-open").forEach((panel) => {
      panel.classList.remove("is-flyout-open");
      const rail = panel.querySelector(".sidebar-panel__rail");
      if (rail) rail.setAttribute("aria-expanded", "false");
    });
  }

  function closeMobileDrawer() {
    if (isMobile() && shell?.classList.contains("is-sidebar-drawer-open")) {
      setCollapsed(true);
    }
  }

  collapseBtn?.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!shell) return;
    closeAllPanelFlyouts();
    if (isMobile()) {
      setCollapsed(shell.classList.contains("is-sidebar-drawer-open"));
      return;
    }
    setCollapsed(!shell.classList.contains("is-sidebar-collapsed"));
  });

  document.querySelectorAll(".sidebar-panel__rail").forEach((rail) => {
    rail.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      if (!shell?.classList.contains("is-sidebar-collapsed") || isMobile()) return;
      const panel = rail.closest(".sidebar__primary-panel, .sidebar__secondary-panel");
      if (!panel) return;
      const wasOpen = panel.classList.contains("is-flyout-open");
      closeAllPanelFlyouts();
      if (!wasOpen) {
        panel.classList.add("is-flyout-open");
        rail.setAttribute("aria-expanded", "true");
      }
    });
  });

  document.addEventListener("click", (event) => {
    if (!shell?.classList.contains("is-sidebar-collapsed")) return;
    const panel = event.target.closest?.(".sidebar__primary-panel, .sidebar__secondary-panel");
    if (!panel) closeAllPanelFlyouts();
  });

  backdrop?.addEventListener("click", () => {
    closeMobileDrawer();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeAllPanelFlyouts();
      closeMobileDrawer();
    }
  });

  // Clic fuera del sidebar (main / topbar / footer) también cierra
  shell?.addEventListener("click", (event) => {
    if (!isMobile() || !shell.classList.contains("is-sidebar-drawer-open")) return;
    const sidebar = document.getElementById("sidebarNavigation");
    if (sidebar?.contains(event.target)) return;
    if (collapseBtn?.contains(event.target)) return;
    if (backdrop?.contains(event.target)) return;
    if (!sidebar?.contains(event.target)) {
      setCollapsed(true);
    }
  });

  searchInput?.addEventListener("input", () => {
    const query = searchInput.value.trim().toLowerCase();
    document.querySelectorAll(".sidebar-nav-link").forEach((el) => {
      if (!query) {
        el.hidden = false;
        return;
      }
      el.hidden = !(el.textContent || "").toLowerCase().includes(query);
    });
  });

  document.querySelectorAll(".sidebar-nav-link").forEach((link) => {
    link.addEventListener("click", () => {
      closeAllPanelFlyouts();
      if (isMobile()) setCollapsed(true);
    });
  });

  mobileMq.addEventListener?.("change", () => {
    if (isMobile()) {
      shell?.classList.remove("is-sidebar-collapsed");
      setCollapsed(true);
    } else {
      document.body.classList.remove("is-sidebar-drawer-open");
      shell?.classList.remove("is-sidebar-drawer-open");
      syncBackdrop(false);
    }
  });

  window.addEventListener("hashchange", () => {
    clearSidebarSearchFilter();
    applyRoute(getCurrentRoute());
  });

  window.addEventListener("pageshow", () => {
    ensureProduccionNav();
    clearSidebarSearchFilter();
  });

  if (
    !window.location.hash ||
    window.location.hash === "#/" ||
    window.location.hash === "#"
  ) {
    window.location.hash = "#/inicio";
  } else if (
    window.location.hash === "#/validacion" ||
    window.location.hash === "#/por-hora"
  ) {
    window.location.hash = "#/tareo";
  }

  applyRoute(getCurrentRoute());

  if (isMobile()) {
    shell?.classList.remove("is-sidebar-collapsed");
    setCollapsed(true);
  }

  // Fecha / hora en vivo (navbar empresarial)
  const dateEl = document.getElementById("txtTopbarDate");
  const timeEl = document.getElementById("txtTopbarTime");
  const liveBtn = document.getElementById("btnLive");

  function tickClock() {
    const now = new Date();
    if (dateEl) {
      dateEl.textContent = now.toLocaleDateString("es-PE", {
        weekday: "short",
        day: "2-digit",
        month: "short"
      });
    }
    if (timeEl) {
      timeEl.textContent = now.toLocaleTimeString("es-PE", {
        hour: "2-digit",
        minute: "2-digit"
      });
    }
  }

  tickClock();
  window.setInterval(tickClock, 1000);

  function syncOnlineUi() {
    if (!liveBtn) return;
    const online = typeof navigator !== "undefined" ? navigator.onLine !== false : true;
    liveBtn.classList.toggle("is-live", online);
    liveBtn.classList.toggle("is-offline", !online);
    liveBtn.title = online
      ? "Conexión en línea"
      : "Sin conexión — Pases/Tarjetas usarán caché local si existe";
    const txt = document.getElementById("txtLiveStatus");
    if (txt) txt.textContent = online ? "En vivo" : "Sin red";
  }
  syncOnlineUi();
  window.addEventListener("online", syncOnlineUi);
  window.addEventListener("offline", syncOnlineUi);

  liveBtn?.addEventListener("click", () => {
    syncOnlineUi();
  });
})();
