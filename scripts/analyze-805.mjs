import XLSX from "xlsx";

const file =
  process.argv[2] ||
  "C:/Users/TatianaLeón/Downloads/Reporte_Horas - 2026-09-16T093338.274.xlsx";

const wb = XLSX.readFile(file, { cellDates: true, raw: true });
const matrix = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
  header: 1,
  defval: "",
  raw: true
});
const headers = matrix[0];

const I = {
  cod: 1,
  doc: 2,
  name: 3,
  macro: 8,
  act: 12,
  fundo: 13,
  mod: 14,
  sup: 22,
  fecha: 24,
  hinicio: 27
};

function norm(s) {
  return String(s ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function docKey(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number" && Number.isFinite(v)) {
    // Excel numeric DNI
    return String(Math.trunc(v));
  }
  let s = String(v).trim();
  if (/e\+/i.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n)) return String(Math.trunc(n));
  }
  s = s.replace(/\.0$/, "");
  return s;
}

function modOk(mod) {
  const m = norm(mod);
  return /MODULO\s*\d+/.test(m) && !m.includes("QBERRIES");
}

function isCosechaExact(act) {
  return norm(act) === "COSECHA";
}

function isCosechaResumen(act) {
  const a = norm(act)
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
  // simulate resumen: nfd already upper - redo lower
  const low = String(act || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
  return low
    .split("/")
    .map((p) => p.trim())
    .some((p) => p === "cosecha");
}

const modsUnicos = new Map(); // doc -> first row
const allLicapaCosecha = new Map();

for (let r = 1; r < matrix.length; r++) {
  const row = matrix[r];
  if (!row) continue;
  const fundo = norm(row[I.fundo]);
  if (fundo !== "LICAPA") continue;
  const act = row[I.act];
  if (!isCosechaExact(act)) continue;
  const doc = docKey(row[I.doc]);
  const name = String(row[I.name] ?? "").trim();
  const mod = String(row[I.mod] ?? "").trim();
  const rec = {
    r: r + 1,
    doc,
    name,
    mod,
    act: String(act ?? ""),
    fundo: String(row[I.fundo] ?? ""),
    sup: String(row[I.sup] ?? "").trim(),
    fecha: row[I.fecha],
    macro: String(row[I.macro] ?? ""),
    rawDoc: row[I.doc],
    rawDocType: typeof row[I.doc]
  };

  if (!allLicapaCosecha.has(doc || `n:${norm(name)}`)) {
    allLicapaCosecha.set(doc || `n:${norm(name)}`, rec);
  }
  if (modOk(mod) && doc) {
    if (!modsUnicos.has(doc)) modsUnicos.set(doc, rec);
  }
}

console.log("Únicos COSECHA LICAPA (cualquier módulo):", allLicapaCosecha.size);
console.log("Únicos COSECHA LICAPA módulo válido:", modsUnicos.size);

// Simulate collapseToDayRows key: documento|fecha|macro
const dayMap = new Map();
for (let r = 1; r < matrix.length; r++) {
  const row = matrix[r];
  if (!row) continue;
  if (norm(row[I.fundo]) !== "LICAPA") continue;
  const doc = docKey(row[I.doc]);
  const fecha = row[I.fecha];
  const fechaStr =
    fecha instanceof Date
      ? fecha.toISOString().slice(0, 10)
      : String(fecha ?? "").trim();
  const macro = String(row[I.macro] ?? "");
  const id = doc || `n:${norm(row[I.name])}` || `row-${r}`;
  const key = `${id}|${fechaStr}|${macro}`;
  if (!dayMap.has(key)) {
    dayMap.set(key, {
      doc,
      name: String(row[I.name] ?? "").trim(),
      acts: new Set(),
      mods: new Set(),
      rows: []
    });
  }
  const g = dayMap.get(key);
  g.acts.add(String(row[I.act] ?? ""));
  g.mods.add(String(row[I.mod] ?? "").trim());
  g.rows.push(r + 1);
}

let resumenCosecha = 0;
const resumenDocs = new Set();
const dayWithCosechaNoModOk = [];
const dayWithCosecha = [];

dayMap.forEach((g, key) => {
  const actJoined = [...g.acts].join(" / ");
  if (!isCosechaResumen(actJoined) && ![...g.acts].some(isCosechaResumen)) return;
  // has cosecha
  const hasExact = [...g.acts].some(isCosechaExact);
  if (!hasExact && !isCosechaResumen(actJoined)) return;

  const pKey = g.doc ? g.doc : `n:${norm(g.name)}`;
  if (!g.doc) {
    // would have been skipped by OLD resumen
  }
  resumenDocs.add(pKey);
  dayWithCosecha.push({ key, ...g, actJoined });
  const anyModOk = [...g.mods].some(modOk);
  if (!anyModOk) dayWithCosechaNoModOk.push({ key, ...g, actJoined });
});

console.log("Resumen-style únicos con COSECHA (LICAPA):", resumenDocs.size);

// Diff: in modsUnicos but not resumenDocs
const onlyMods = [];
modsUnicos.forEach((rec, doc) => {
  if (!resumenDocs.has(doc)) onlyMods.push(rec);
});
const onlyResumen = [];
resumenDocs.forEach((doc) => {
  if (!modsUnicos.has(doc) && !String(doc).startsWith("n:")) onlyResumen.push(doc);
});

console.log("Solo en módulos (no resumen):", onlyMods.length);
onlyMods.slice(0, 20).forEach((p) =>
  console.log(" ", p.doc, p.name, p.mod, "fila", p.r, "rawDoc", p.rawDoc, p.rawDocType)
);
console.log("Solo en resumen (no módulos válidos):", onlyResumen.length);
onlyResumen.slice(0, 20).forEach((d) => console.log(" ", d));

// Check empty docs
const emptyDocs = [...allLicapaCosecha.values()].filter((p) => !p.doc);
console.log("Sin DNI en COSECHA LICAPA:", emptyDocs.length);
emptyDocs.forEach((p) => console.log(" ", p.name, p.mod, "fila", p.r));

// Find if any doc appears with different string forms
const forms = new Map();
for (let r = 1; r < matrix.length; r++) {
  const row = matrix[r];
  if (norm(row[I.fundo]) !== "LICAPA") continue;
  if (!isCosechaExact(row[I.act])) continue;
  const raw = row[I.doc];
  const k = docKey(raw);
  if (!k) continue;
  if (!forms.has(k)) forms.set(k, new Set());
  forms.get(k).add(`${typeof raw}:${JSON.stringify(raw)}`);
}
const multiForms = [...forms.entries()].filter(([, s]) => s.size > 1);
console.log("DNIs con múltiples formas raw:", multiForms.length);
multiForms.slice(0, 10).forEach(([k, s]) => console.log(k, [...s]));

// People with COSECHA but ONLY invalid modules
const onlyInvalidMod = [];
allLicapaCosecha.forEach((rec, key) => {
  if (modsUnicos.has(rec.doc)) return;
  // check all their modules
  let ok = false;
  let mods = new Set();
  for (let r = 1; r < matrix.length; r++) {
    const row = matrix[r];
    if (docKey(row[I.doc]) !== rec.doc && !( !rec.doc && norm(row[I.name]) === norm(rec.name))) continue;
    if (norm(row[I.fundo]) !== "LICAPA") continue;
    if (!isCosechaExact(row[I.act])) continue;
    mods.add(String(row[I.mod] ?? "").trim());
    if (modOk(row[I.mod])) ok = true;
  }
  if (!ok) onlyInvalidMod.push({ ...rec, mods: [...mods] });
});
console.log("COSECHA LICAPA solo en módulo inválido (QBERRIES etc):", onlyInvalidMod.length);
onlyInvalidMod.slice(0, 15).forEach((p) =>
  console.log(" ", p.doc, p.name, p.mods.join(" | "), "fila", p.r)
);

// Count unique docs with COSECHA LICAPA ignoring module filter (resumen style with doc only)
const resumenDocOnly = new Set();
dayMap.forEach((g) => {
  const has = [...g.acts].some((a) => isCosechaExact(a) || isCosechaResumen(a));
  if (!has) return;
  if (g.doc) resumenDocOnly.add(g.doc);
});
console.log("Resumen solo DNI (con COSECHA):", resumenDocOnly.size);
console.log("Diff módulos(805) - resumenDNI:", modsUnicos.size - resumenDocOnly.size);

const missingInResumen = [];
modsUnicos.forEach((rec, doc) => {
  if (!resumenDocOnly.has(doc)) missingInResumen.push(rec);
});
console.log("En módulos válidos pero NO en resumen DNI:", missingInResumen.length);
missingInResumen.forEach((p) =>
  console.log(" PROBLEM:", p.doc, p.name, p.mod, "fila", p.r, "fecha", p.fecha, "macro", p.macro)
);
