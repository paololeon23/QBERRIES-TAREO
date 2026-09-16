import XLSX from "xlsx";
import path from "path";

const file =
  process.argv[2] ||
  "C:/Users/TatianaLeón/Downloads/Reporte_Horas - 2026-09-16T093338.274.xlsx";

const wb = XLSX.readFile(file, { cellDates: true, raw: false });
console.log("Sheets:", wb.SheetNames);

function norm(s) {
  return String(s || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function findHeaderRow(matrix) {
  for (let i = 0; i < Math.min(15, matrix.length); i++) {
    const row = (matrix[i] || []).map((c) => norm(c));
    const joined = row.join("|");
    if (
      joined.includes("DOCUMENTO") ||
      joined.includes("DNI") ||
      (joined.includes("TRABAJADOR") && joined.includes("ACTIVIDAD"))
    ) {
      return i;
    }
  }
  return 0;
}

function colIndex(headers, candidates) {
  const h = headers.map((x) => norm(x));
  for (const c of candidates) {
    const n = norm(c);
    const i = h.findIndex((x) => x === n || x.includes(n));
    if (i >= 0) return i;
  }
  return -1;
}

for (const sn of wb.SheetNames) {
  const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sn], {
    header: 1,
    defval: "",
    raw: false
  });
  const hi = findHeaderRow(matrix);
  const headers = matrix[hi] || [];
  console.log("\n===", sn, "headerRow", hi, "===");
  console.log(
    "headers sample:",
    headers
      .map((h, i) => `${i}:${h}`)
      .filter(Boolean)
      .slice(0, 40)
      .join(" | ")
  );

  const iDoc = colIndex(headers, ["documento", "dni", "nro documento", "nro. documento"]);
  const iName = colIndex(headers, ["trabajador", "nombre", "apellidos y nombres"]);
  const iAct = colIndex(headers, ["actividad"]);
  const iFundo = colIndex(headers, ["fundo"]);
  const iMod = colIndex(headers, ["modulo", "módulo"]);
  const iSup = colIndex(headers, ["supervisor"]);
  console.log({ iDoc, iName, iAct, iFundo, iMod, iSup });

  if (iDoc < 0 || iAct < 0) continue;

  const cosecha = [];
  const sinDni = [];
  const byDoc = new Set();
  const byKey = new Set();

  for (let r = hi + 1; r < matrix.length; r++) {
    const row = matrix[r] || [];
    const act = norm(row[iAct]);
    if (act !== "COSECHA") continue;
    const fundo = iFundo >= 0 ? norm(row[iFundo]) : "";
    if (fundo && fundo !== "LICAPA") continue;
    const mod = iMod >= 0 ? String(row[iMod] || "").trim() : "";
    const modN = norm(mod);
    const modOk = /MODULO\s*\d+/.test(modN) && !modN.includes("QBERRIES");
    const doc = String(row[iDoc] ?? "")
      .trim()
      .replace(/\.0$/, "");
    const name = String(row[iName] ?? "").trim();
    const sup = iSup >= 0 ? String(row[iSup] ?? "").trim() : "";
    const rec = { excelRow: r + 1, doc, name, act: row[iAct], fundo: row[iFundo], mod, sup, modOk };

    if (modOk) {
      cosecha.push(rec);
      if (doc) byDoc.add(doc);
      const key = doc ? `d:${doc}` : name ? `n:${norm(name)}` : "";
      if (key) byKey.add(key);
      if (!doc) sinDni.push(rec);
    }
  }

  // unique by key for people with valid module
  const uniqPeople = new Map();
  for (const rec of cosecha) {
    const key = rec.doc ? `d:${rec.doc}` : rec.name ? `n:${norm(rec.name)}` : "";
    if (!key) continue;
    if (!uniqPeople.has(key)) uniqPeople.set(key, rec);
  }

  console.log("COSECHA LICAPA con módulo válido:");
  console.log("  filas", cosecha.length);
  console.log("  únicos por DNI", byDoc.size);
  console.log("  únicos DNI+nombre", byKey.size);
  console.log("  sin DNI (filas)", sinDni.length);
  console.log("  sin DNI (únicos)", [...uniqPeople.keys()].filter((k) => k.startsWith("n:")).length);

  if (sinDni.length) {
    console.log("\nPERSONAS SIN DNI:");
    const seen = new Set();
    for (const p of sinDni) {
      const k = norm(p.name) || `row${p.excelRow}`;
      if (seen.has(k)) continue;
      seen.add(k);
      console.log(
        `  - ${p.name || "(sin nombre)"} | fila Excel ${p.excelRow} | ${p.mod} | ${p.sup} | fundo=${p.fundo}`
      );
    }
  }

  // Also: DNIs that appear empty-looking (spaces, 0, etc.)
  const weirdDoc = [];
  for (const rec of cosecha) {
    const d = rec.doc;
    if (!d) continue;
    if (!/^\d{8,}$/.test(d.replace(/\D/g, "")) || d.replace(/\D/g, "").length < 8) {
      weirdDoc.push(rec);
    }
  }
  if (weirdDoc.length) {
    console.log("\nDNI raros / cortos:");
    const seen = new Set();
    for (const p of weirdDoc) {
      if (seen.has(p.doc)) continue;
      seen.add(p.doc);
      console.log(`  - doc="${p.doc}" name=${p.name} fila=${p.excelRow} mod=${p.mod}`);
    }
  }
}
