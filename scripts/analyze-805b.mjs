import fs from "fs";
import { createRequire } from "module";
import { pathToFileURL } from "url";

// Use installed xlsx + project's excel-parser if possible
const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const file =
  process.argv[2] ||
  "C:/Users/TatianaLeón/Downloads/Reporte_Horas - 2026-09-16T093338.274.xlsx";

// Minimal reimplementation matching excel-parser column mapping FALLBACK
// Documento col C (2), Trabajador D (3), Actividad M (12), Fundo N (13), Modulo O (14)
const buf = fs.readFileSync(file);
const wb = XLSX.read(buf, { type: "buffer", cellDates: true, raw: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });

function cleanText(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number" && Number.isFinite(v)) return String(Math.trunc(v));
  return String(v).replace(/\s+/g, " ").trim();
}

function normAct(a) {
  return String(a || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function isCosecha(act) {
  return normAct(act)
    .split("/")
    .map((p) => p.trim())
    .some((p) => p === "cosecha");
}

function isCosechaExact(act) {
  return normAct(act) === "cosecha";
}

function modOk(mod) {
  const m = String(mod || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
  return /MODULO\s*\d+/.test(m) && !m.includes("QBERRIES");
}

const docsResumen = new Set();
const docsModulos = new Set();
const people = [];

for (let r = 1; r < matrix.length; r++) {
  const row = matrix[r];
  if (!row) continue;
  const documento = cleanText(row[2]);
  const trabajador = cleanText(row[3]);
  const actividad = cleanText(row[12]);
  const fundo = cleanText(row[13]);
  const modulo = cleanText(row[14]);
  const supervisor = cleanText(row[22]);
  const macro = cleanText(row[8]);

  if (normAct(fundo) !== "licapa") continue;

  const rec = { r: r + 1, documento, trabajador, actividad, fundo, modulo, supervisor, macro, rawDoc: row[2] };

  if (isCosecha(actividad)) {
    if (documento) docsResumen.add(documento);
    else people.push({ ...rec, why: "resumen-sin-dni" });
  }
  if (isCosechaExact(actividad) && modOk(modulo)) {
    if (documento) docsModulos.add(documento);
    else people.push({ ...rec, why: "modulos-sin-dni" });
  }
}

console.log("Resumen COSECHA LICAPA únicos DNI:", docsResumen.size);
console.log("Módulos COSECHA LICAPA únicos DNI:", docsModulos.size);

const onlyR = [...docsResumen].filter((d) => !docsModulos.has(d));
const onlyM = [...docsModulos].filter((d) => !docsResumen.has(d));
console.log("Solo resumen:", onlyR.length, onlyR.slice(0, 10));
console.log("Solo módulos:", onlyM.length, onlyM.slice(0, 10));
console.log("Sin DNI hits:", people.length);
people.slice(0, 20).forEach((p) => console.log(p));

// Check if any documento becomes empty after cleanText but raw exists
let weird = 0;
const weirdList = [];
for (let r = 1; r < matrix.length; r++) {
  const row = matrix[r];
  if (!row) continue;
  if (normAct(cleanText(row[13])) !== "licapa") continue;
  if (!isCosechaExact(cleanText(row[12]))) continue;
  const raw = row[2];
  const cleaned = cleanText(raw);
  if (raw != null && raw !== "" && !cleaned) {
    weird++;
    weirdList.push({ r: r + 1, raw, type: typeof raw, name: cleanText(row[3]) });
  }
  // float DNIs
  if (typeof raw === "number" && !Number.isInteger(raw)) {
    weirdList.push({ r: r + 1, raw, type: "float", name: cleanText(row[3]), cleaned });
  }
}
console.log("weird docs:", weirdList.length);
weirdList.slice(0, 20).forEach((w) => console.log(w));

// Macro Partida = COSTO DE COSECHA filter?
const byMacro = new Map();
for (let r = 1; r < matrix.length; r++) {
  const row = matrix[r];
  if (normAct(cleanText(row[13])) !== "licapa") continue;
  if (!isCosecha(cleanText(row[12]))) continue;
  const macro = cleanText(row[8]) || "(vacío)";
  const doc = cleanText(row[2]);
  if (!doc) continue;
  if (!byMacro.has(macro)) byMacro.set(macro, new Set());
  byMacro.get(macro).add(doc);
}
console.log("\nPor Macro Partida (COSECHA LICAPA):");
[...byMacro.entries()]
  .sort((a, b) => b[1].size - a[1].size)
  .forEach(([m, s]) => console.log(`  ${s.size}  ${m}`));
