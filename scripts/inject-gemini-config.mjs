/**
 * No-op de compatibilidad con el Build command de Netlify (UI).
 * Este repo (QBERRIES-TAREO) no usa Gemini; el sitio vive en shell-template/.
 * Sale 0 para no romper el deploy.
 */
console.log("[inject-gemini-config] skip — no Gemini config in this repo");
process.exit(0);
