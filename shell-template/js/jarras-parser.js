/**
 * Catálogos Trabajadores Jarras:
 * - Lotes: DATA LICAPA · H=lote · F=módulo · G=turno
 * - Grupos: Produccion_Licapa · columna Grupo (F) "Grupo LIC N"
 */

const LOTES_JSON = "./data/lotes-licapa.json";
const GRUPOS_JSON = "./data/grupos-licapa.json";

function clean(v) {
  return String(v ?? "").trim();
}

function colLetterToIndex(letter) {
  let n = 0;
  for (const ch of String(letter || "").toUpperCase()) {
    if (ch < "A" || ch > "Z") continue;
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

function cell(row, idx) {
  if (!row || idx < 0 || idx >= row.length) return "";
  return row[idx];
}

/** Parsea DATA LICAPA (matriz AOA) → lotes únicos. */
export function parseDataLicapaMatrix(matrix) {
  if (!matrix?.length) return [];
  const map = new Map();
  const start = clean(cell(matrix[0], colLetterToIndex("F"))).toLowerCase().includes("modulo")
    ? 1
    : 0;
  for (let i = start; i < matrix.length; i++) {
    const row = matrix[i];
    const lote = clean(cell(row, colLetterToIndex("H")));
    const modulo = clean(cell(row, colLetterToIndex("F")));
    const turno = clean(cell(row, colLetterToIndex("G")));
    if (!lote || !modulo) continue;
    const key = `${lote}|${modulo}|${turno}`;
    if (map.has(key)) continue;
    map.set(key, {
      lote,
      modulo,
      turno,
      cod: clean(cell(row, colLetterToIndex("C"))),
      label: clean(cell(row, colLetterToIndex("N")))
    });
  }
  return [...map.values()].sort((a, b) => {
    const na = Number(String(a.lote).replace(/\D/g, "")) || 0;
    const nb = Number(String(b.lote).replace(/\D/g, "")) || 0;
    return na - nb || String(a.lote).localeCompare(String(b.lote), "es");
  });
}

/** Extrae grupos únicos desde columna F (Grupo LIC …) o C si vienen ahí. */
export function parseGruposFromProduccionMatrix(matrix) {
  if (!matrix?.length) return [];
  const set = new Set();
  const idxF = colLetterToIndex("F");
  const idxC = colLetterToIndex("C");
  for (let i = 0; i < matrix.length; i++) {
    const row = matrix[i];
    for (const idx of [idxF, idxC]) {
      const raw = clean(cell(row, idx));
      if (!raw || /^grupo$/i.test(raw)) continue;
      const m = raw.match(/Grupo\s*LIC\s*\d+/i);
      if (m) set.add(m[0].replace(/\s+/g, " ").replace(/lic/i, "LIC"));
      else if (/^LIC\s*\d+/i.test(raw)) set.add(`Grupo ${raw.replace(/lic/i, "LIC")}`);
    }
  }
  return [...set].sort((a, b) => {
    const na = Number((a.match(/(\d+)\s*$/) || [])[1] || 0);
    const nb = Number((b.match(/(\d+)\s*$/) || [])[1] || 0);
    return na - nb || a.localeCompare(b, "es");
  });
}

export async function loadPackedLotes() {
  const res = await fetch(LOTES_JSON, { cache: "no-store" });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.json();
  return Array.isArray(data.lotes) ? data.lotes : [];
}

export async function loadPackedGrupos() {
  const res = await fetch(GRUPOS_JSON, { cache: "no-store" });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.json();
  return Array.isArray(data.grupos) ? data.grupos : [];
}

export function findLoteMeta(lotes, loteId) {
  const key = clean(loteId);
  if (!key) return null;
  return (
    lotes.find((l) => String(l.lote) === key) ||
    lotes.find((l) => String(l.label) === key) ||
    lotes.find((l) => String(l.cod) === key) ||
    null
  );
}

/** Texto visible del lote: sin número crudo (usa L003- / Q3). */
export function loteDisplay(l) {
  if (!l) return "";
  const label = clean(l.label).replace(/-$/, "");
  const cod = clean(l.cod);
  if (label) return label;
  if (cod) return cod;
  return clean(l.lote);
}

export function loteSearchBlob(l) {
  return `${l.label || ""} ${l.cod || ""} ${l.lote || ""} ${l.modulo || ""} ${l.turno || ""}`.toLowerCase();
}
