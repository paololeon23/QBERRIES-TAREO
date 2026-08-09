const ROUTE_TITLES = {
  "por-hora": "Por hora",
  validacion: "Por hora",
  historial: "Historial",
  recomendaciones: "Recomendaciones"
};

const ROUTE_VIEW = {
  "por-hora": "por-hora",
  validacion: "por-hora",
  historial: "historial",
  recomendaciones: "recomendaciones"
};

function normalizeHash(hash) {
  const value = (hash || "#/por-hora").replace(/^#\/?/, "");
  let key = value.split("/")[0] || "por-hora";
  if (key === "" || key === "/") key = "por-hora";
  if (key === "validacion") key = "por-hora";
  return ROUTE_TITLES[key] ? key : "por-hora";
}

export function getCurrentRoute() {
  return normalizeHash(window.location.hash);
}

export function applyRoute(routeKey) {
  const route = ROUTE_TITLES[routeKey] ? routeKey : "por-hora";
  const viewId = ROUTE_VIEW[route] || "por-hora";

  document.querySelectorAll(".view").forEach((view) => {
    const match = view.dataset.view === viewId;
    view.hidden = !match;
    view.classList.toggle("is-active", match);
  });

  document.querySelectorAll(".sidebar-nav-link[data-route]").forEach((link) => {
    const linkRoute = link.dataset.route === "validacion" ? "por-hora" : link.dataset.route;
    link.classList.toggle("is-active", linkRoute === viewId || link.dataset.route === route);
  });

  const title = ROUTE_TITLES[route];
  const titleEl = document.getElementById("txtPageTitle");
  const crumbEl = document.getElementById("txtBreadcrumbActive");
  if (titleEl) titleEl.textContent = title;
  if (crumbEl) crumbEl.textContent = title;

  window.dispatchEvent(new CustomEvent("qb:route-changed", { detail: { route: viewId } }));
}

(() => {
  const shell = document.getElementById("applicationRoot");
  const collapseBtn = document.getElementById("btnSidebarCollapse");
  const searchInput = document.getElementById("txtSidebarSearch");
  const backdrop = document.getElementById("sidebarDrawerBackdrop");
  const mobileMq = window.matchMedia("(max-width: 768px)");

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
    // Si el clic fue en el contenido (no en el panel), cerrar
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

  window.addEventListener("hashchange", () => applyRoute(getCurrentRoute()));

  if (
    !window.location.hash ||
    window.location.hash === "#/" ||
    window.location.hash === "#" ||
    window.location.hash === "#/validacion"
  ) {
    window.location.hash = "#/por-hora";
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

  // Buscador superior: filtra menú + enfoca búsqueda de tabla si existe
  const topSearch = document.getElementById("txtTopbarSearch");
  topSearch?.addEventListener("input", () => {
    const query = topSearch.value.trim().toLowerCase();
    document.querySelectorAll(".sidebar-nav-link").forEach((el) => {
      if (!query) {
        el.hidden = false;
        return;
      }
      el.hidden = !(el.textContent || "").toLowerCase().includes(query);
    });
    const tableSearch = document.getElementById("fltSearch");
    if (tableSearch && document.getElementById("validacionWorkspace") && !document.getElementById("validacionWorkspace").classList.contains("is-hidden")) {
      tableSearch.value = topSearch.value;
      tableSearch.dispatchEvent(new Event("input", { bubbles: true }));
    }
  });
})();
