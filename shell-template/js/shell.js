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
  "reporte-trabajadores": "Reporte de Trabajadores"
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
  "reporte-trabajadores": "reporte-trabajadores"
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
  const panel = tareo?.closest(".sidebar__secondary-panel") || document.querySelector(".sidebar__secondary-panel");
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

  function setCollapsed(collapsed) {
    if (!shell) return;

    if (isMobile()) {
      if (collapsed) {
        shell.classList.remove("is-sidebar-drawer-open");
        document.body.classList.remove("is-sidebar-drawer-open");
        syncBackdrop(false);
      } else {
        shell.classList.add("is-sidebar-drawer-open");
        document.body.classList.add("is-sidebar-drawer-open");
        syncBackdrop(true);
      }
      collapseBtn?.classList.toggle("is-expand-state", !shell.classList.contains("is-sidebar-drawer-open"));
      return;
    }

    syncBackdrop(false);
    shell.classList.add("is-sidebar-collapsing");
    shell.classList.toggle("is-sidebar-collapsed", collapsed);
    collapseBtn?.classList.toggle("is-expand-state", collapsed);
    window.setTimeout(() => shell.classList.remove("is-sidebar-collapsing"), 220);
  }

  function closeMobileDrawer() {
    if (isMobile() && shell?.classList.contains("is-sidebar-drawer-open")) {
      setCollapsed(true);
    }
  }

  collapseBtn?.addEventListener("click", () => {
    if (isMobile()) {
      setCollapsed(shell.classList.contains("is-sidebar-drawer-open"));
      return;
    }
    setCollapsed(!shell.classList.contains("is-sidebar-collapsed"));
  });

  backdrop?.addEventListener("click", () => {
    closeMobileDrawer();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeMobileDrawer();
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

  liveBtn?.addEventListener("click", () => {
    liveBtn.classList.toggle("is-live");
    liveBtn.title = liveBtn.classList.contains("is-live") ? "Estado en vivo" : "Pausado";
  });
})();
