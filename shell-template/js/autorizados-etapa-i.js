/**
 * Personal autorizado Etapa I (Licapa I).
 * La lista se usa para identificar sesión; NO bloquea menú ni carga de Excel.
 */

export const AUTORIZADOS_ETAPA_I = [
  { dni: "48268173", nombre: "NORIEGA PONTE, MICELY" },
  { dni: "77534125", nombre: "HERRERA ALBERCA, PAMELA" },
  { dni: "70192702", nombre: "PURIZAGA SAAVEDRA, PIERRE OSNAR" },
  { dni: "75141739", nombre: "ROJAS AREDO, YERSI YEN" },
  { dni: "48446147", nombre: "HILARIO ÁVALOS, EVELYN" },
  { dni: "48533707", nombre: "JULCA GAMBOA, DILMER ELICER" },
  { dni: "19085520", nombre: "RODRÍGUEZ GUTIÉRREZ, SANTOS RUBEN" },
  { dni: "77146080", nombre: "CHACÓN BERMÚDEZ, NADIA SARAHÍ" },
  { dni: "45372928", nombre: "PEÑA ROJAS, LAURA PATRICIA" },
  { dni: "71806261", nombre: "LEÓN TRIGOSO, JHONY ANDRÓNICO" },
  { dni: "70656198", nombre: "VARGAS DÍAZ, ALDRIN" }
];

const STORAGE_KEY = "qb.autorizado.etapaI.v1";
const GUEST = { dni: "", nombre: "QBerries" };
const BY_DNI = new Map(AUTORIZADOS_ETAPA_I.map((p) => [p.dni, p]));

export function normalizeDni(value) {
  return String(value || "").replace(/\D/g, "");
}

export function findAutorizado(dni) {
  const key = normalizeDni(dni);
  if (key.length < 8) return null;
  return BY_DNI.get(key) || null;
}

export function getSesionAutorizado() {
  try {
    const raw = sessionStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    const hit = findAutorizado(parsed?.dni);
    if (!hit) {
      sessionStorage.removeItem(STORAGE_KEY);
      return null;
    }
    return hit;
  } catch {
    return null;
  }
}

export function setSesionAutorizado(persona) {
  sessionStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({ dni: persona.dni, nombre: persona.nombre, at: Date.now() })
  );
}

export function clearSesionAutorizado() {
  sessionStorage.removeItem(STORAGE_KEY);
}

export function puedeSubirInformacion() {
  return true;
}

export function assertPuedeSubir() {
  return getSesionAutorizado() || GUEST;
}

function $(id) {
  return document.getElementById(id);
}

function initials(nombre) {
  const [apellidos, nombres] = String(nombre || "").split(",");
  const a = (apellidos || "").trim().split(/\s+/)[0] || "";
  const n = (nombres || "").trim().split(/\s+/)[0] || "";
  const pair = `${n.charAt(0)}${a.charAt(0)}`.toUpperCase();
  return pair.replace(/[^A-ZÁÉÍÓÚÑ]/g, "") || "QB";
}

function setError(msg) {
  const el = $("txtAccesoError");
  if (!el) return;
  if (!msg) {
    el.hidden = true;
    el.textContent = "";
    return;
  }
  el.hidden = false;
  el.textContent = msg;
}

function hideOverlay() {
  const overlay = $("overlayAccesoEtapaI");
  if (overlay) overlay.hidden = true;
}

function unlockShell() {
  document.body.classList.remove("is-acceso-locked");
  const root = $("applicationRoot");
  if (root) root.removeAttribute("inert");
  document.querySelectorAll('input[type="file"]').forEach((el) => {
    el.disabled = false;
  });
  hideOverlay();
}

function syncChrome(sesion) {
  unlockShell();

  const nameEl = $("sidebarUserName");
  const roleEl = $("sidebarUserRole");
  const avatarEl = $("sidebarUserAvatar");
  const chip = $("chipAccesoUser");
  const chipName = $("txtAccesoUser");
  const btnSalir = $("btnAccesoSalir");

  if (sesion) {
    if (nameEl) nameEl.textContent = sesion.nombre.replace(",", "");
    if (roleEl) roleEl.textContent = `Etapa I · ${sesion.dni}`;
    if (avatarEl) avatarEl.textContent = initials(sesion.nombre);
    if (chip) chip.hidden = false;
    if (chipName) chipName.textContent = sesion.nombre.split(",")[0].trim();
    if (btnSalir) btnSalir.hidden = false;
  } else {
    if (nameEl) nameEl.textContent = "QBerries";
    if (roleEl) roleEl.textContent = "Campo";
    if (avatarEl) avatarEl.textContent = "QB";
    if (chip) chip.hidden = true;
    if (btnSalir) btnSalir.hidden = true;
  }
}

export function showAccesoGate(message) {
  unlockShell();
  if (message) setError(message);
}

function entrar() {
  const input = $("txtAccesoDni");
  const dni = normalizeDni(input?.value);
  const hit = findAutorizado(dni);
  if (!hit) {
    setError("DNI no autorizado para Etapa I.");
    input?.select();
    return;
  }
  setSesionAutorizado(hit);
  setError("");
  if (input) input.value = "";
  syncChrome(hit);
}

function salir() {
  clearSesionAutorizado();
  setError("");
  const input = $("txtAccesoDni");
  if (input) input.value = "";
  syncChrome(null);
}

function bindAccesoUi() {
  $("btnAccesoEntrar")?.addEventListener("click", entrar);
  $("btnAccesoSalir")?.addEventListener("click", salir);
  $("btnAccesoCerrar")?.addEventListener("click", () => {
    setError("");
    hideOverlay();
  });
  $("formAccesoEtapaI")?.addEventListener("submit", (ev) => {
    ev.preventDefault();
    entrar();
  });
  $("txtAccesoDni")?.addEventListener("input", () => setError(""));
}

function initAccesoEtapaI() {
  unlockShell();
  bindAccesoUi();
  syncChrome(getSesionAutorizado());
  window.QBerriesAcceso = {
    AUTORIZADOS_ETAPA_I,
    findAutorizado,
    getSesionAutorizado,
    puedeSubirInformacion,
    assertPuedeSubir,
    showAccesoGate,
    salir
  };
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initAccesoEtapaI, { once: true });
} else {
  initAccesoEtapaI();
}
