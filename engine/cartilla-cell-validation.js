/**
 * Evaluación de celdas desde validaciones_por_columna (configuration-driven).
 */

import { getReglaColumnaPorIndice, resolverMensajeRegla } from "./cartilla-rules.adapter.js";

export function cellDisplayValue(val) {
  if (val === null || val === undefined) return "";
  if (typeof val === "object" && val !== null) {
    if ("w" in val && val.w != null && String(val.w).trim() !== "") return String(val.w);
    if ("v" in val && val.v !== undefined && val.v !== null && val.v !== "") return String(val.v);
    if (Array.isArray(val.r)) {
      return val.r.map((x) => (x && x.w != null ? x.w : x.t) || "").join("");
    }
  }
  return String(val).trim();
}

export function matchesExactValue(val, expected) {
  const str = String(val).trim();
  const exp = String(expected).trim();
  if (str === exp) return true;
  const numVal = Number(str.replace(",", "."));
  const numExp = Number(exp.replace(",", "."));
  if (Number.isFinite(numVal) && Number.isFinite(numExp)) {
    return Math.abs(numVal - numExp) < 0.001;
  }
  return false;
}

export function parseFlexibleNumber(val) {
  const s = cellDisplayValue(val).replace(/\s/g, "").replace(",", ".");
  if (s === "") return NaN;
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

function ruleMessage(rule, fallback) {
  return rule.mensaje || rule.mensaje_error || fallback;
}

function runtimeRuleMessage(config, colIdx, rule, cellVal, tipoFallo, fallback) {
  const col = getReglaColumnaPorIndice(config?._reglasOrigen, colIdx);
  const extras = {
    valor: cellVal,
    texto: cellVal,
    lote: cellVal,
    detectado: cellVal.length,
    "longitud-detectada": String(cellVal.length)
  };
  if (col?.["si-falla-mostrar"]) {
    return resolverMensajeRegla(col, extras, tipoFallo);
  }
  const template = rule?.mensaje || rule?.mensaje_error;
  if (template && /\{\{/.test(template)) {
    return resolverMensajeRegla({ "si-falla-mostrar": template }, extras, tipoFallo);
  }
  return ruleMessage(rule, fallback);
}

function isValidHoraHHMM(val) {
  const s = cellDisplayValue(val);
  if (!/^\d{1,2}:\d{2}$/.test(s)) return false;
  const [h, m] = s.split(":").map(Number);
  return h >= 0 && h <= 23 && m >= 0 && m <= 59;
}

function compareColumnValues(val, other, ctx) {
  if (ctx?.normalizeDate) {
    const aIso = ctx.normalizeDate(val);
    const bIso = ctx.normalizeDate(other);
    // Solo comparar como fechas si ambas se normalizan; evita ISO vs texto suelto
    if (aIso && bIso) return aIso === bIso;
  }
  return cellDisplayValue(val) === cellDisplayValue(other);
}

export function getCellValidationIssues(idx, rawVal, ctx, config, options = {}) {
  const issues = [];
  const val = cellDisplayValue(rawVal);
  const colRules = (config?.validaciones_por_columna || []).find((c) => c.indice_js === idx);

  if (colRules) {
    const hasta = colRules.hasta_indice_js;
    const indices =
      hasta != null && hasta >= idx ? Array.from({ length: hasta - idx + 1 }, (_, i) => idx + i) : [idx];

    indices.forEach((colIdx) => {
      const cellVal = colIdx === idx ? val : cellDisplayValue(ctx.row?.[colIdx]);
      colRules.reglas.forEach((rule) => {
        if (rule.tipo === "obligatorio" && !cellVal) {
          issues.push({
            kind: "empty",
            message: runtimeRuleMessage(config, colIdx, rule, cellVal, "obligatorio", "Campo obligatorio"),
            colIdx
          });
        } else if (rule.tipo === "longitud_exacta" && cellVal && cellVal.length !== rule.valor) {
          issues.push({
            kind: "value",
            message: runtimeRuleMessage(
              config,
              colIdx,
              rule,
              cellVal,
              "longitud",
              `Debe tener ${rule.valor} caracteres`
            ),
            colIdx
          });
        } else if (rule.tipo === "valor_exacto" && cellVal && !matchesExactValue(cellVal, rule.valor)) {
          issues.push({
            kind: "value",
            message: runtimeRuleMessage(config, colIdx, rule, cellVal, "igualdad", ruleMessage(rule, "Valor incorrecto")),
            colIdx
          });
        } else if (
          rule.tipo === "texto_exacto" &&
          cellVal &&
          cellVal.toLowerCase() !== String(rule.valor).toLowerCase()
        ) {
          issues.push({
            kind: "value",
            message: runtimeRuleMessage(config, colIdx, rule, cellVal, "igualdad", ruleMessage(rule, "Texto incorrecto")),
            colIdx
          });
        } else if (
          rule.tipo === "contiene_texto" &&
          cellVal &&
          !cellVal.toUpperCase().includes(String(rule.valor).toUpperCase())
        ) {
          issues.push({ kind: "value", message: ruleMessage(rule, `Debe incluir ${rule.valor}`), colIdx });
        } else if (
          (rule.tipo === "duplicado_en_fecha" ||
            rule.tipo === "duplicado_en_cartilla" ||
            rule.tipo === "sin_duplicados_en_fecha") &&
          cellVal &&
          (ctx.duplicadosLote || ctx.duplicadosCartilla || []).includes(cellVal)
        ) {
          issues.push({
            kind: "value",
            message: ruleMessage(rule, `Lote ${cellVal} duplicado`).replace("{{lote}}", cellVal),
            colIdx
          });
        } else if (rule.tipo === "igual_a_columna" && cellVal) {
          const other = ctx.row?.[rule.columna_js];
          if (!cellDisplayValue(other)) {
            issues.push({ kind: "value", message: ruleMessage(rule, "Falta valor de comparación"), colIdx });
          } else if (!compareColumnValues(cellVal, other, ctx)) {
            issues.push({ kind: "value", message: ruleMessage(rule, "No coincide con columna de referencia"), colIdx });
          }
        } else if (rule.tipo === "igual_a_fecha_inspeccion" && cellVal) {
          const inspeccionIdx = config?.filtro_principal?.indice_js ?? ctx.fechaInspeccionIdx;
          const other = inspeccionIdx != null ? ctx.row?.[inspeccionIdx] : "";
          if (!cellDisplayValue(other)) {
            issues.push({ kind: "value", message: ruleMessage(rule, "Falta fecha de inspección"), colIdx });
          } else if (!compareColumnValues(cellVal, other, ctx)) {
            issues.push({ kind: "value", message: ruleMessage(rule, "Debe ser igual a fecha de inspección"), colIdx });
          }
        } else if (rule.tipo === "rango_numerico" && cellVal) {
          const n = parseFlexibleNumber(cellVal);
          if (!Number.isFinite(n)) {
            issues.push({ kind: "value", message: ruleMessage(rule, "Valor numérico inválido"), colIdx });
          } else {
            if (rule.minimo != null && n < rule.minimo) {
              issues.push({
                kind: "value",
                message: runtimeRuleMessage(config, colIdx, rule, cellVal, "rango", `Mínimo ${rule.minimo}`),
                colIdx
              });
            }
            if (rule.maximo != null && n > rule.maximo) {
              issues.push({
                kind: "value",
                message: runtimeRuleMessage(config, colIdx, rule, cellVal, "rango", `Máximo ${rule.maximo}`),
                colIdx
              });
            }
          }
        } else if (rule.tipo === "formato_hora" && cellVal && !isValidHoraHHMM(rawVal)) {
          issues.push({ kind: "value", message: ruleMessage(rule, "Formato HH:MM inválido"), colIdx });
        } else if (rule.tipo === "regex" && cellVal) {
          const patron = new RegExp(rule.patron);
          if (!patron.test(cellVal)) {
            issues.push({ kind: "value", message: ruleMessage(rule, "Formato inválido"), colIdx });
          }
        } else if (rule.tipo === "regex_opcional" && cellVal) {
          const patron = new RegExp(rule.patron);
          if (!patron.test(cellVal)) {
            issues.push({ kind: "value", message: ruleMessage(rule, "Formato inválido"), colIdx });
          }
        } else if (rule.tipo === "digito_unico" && cellVal && !/^[0-9]$/.test(cellVal)) {
          issues.push({ kind: "value", message: ruleMessage(rule, "Debe ser un dígito 0-9"), colIdx });
        } else if (rule.tipo === "lista_valores" && cellVal) {
          const permitidos = (rule.valores || []).map((v) => String(v).toUpperCase());
          if (!permitidos.includes(cellVal.toUpperCase())) {
            issues.push({ kind: "value", message: ruleMessage(rule, "Valor no permitido"), colIdx });
          }
        } else if (rule.tipo === "codigo_en_catalogo" && cellVal && options.catalogos) {
          const catalogo = options.catalogos[rule.catalogo];
          const map = catalogo?.entradas || catalogo || options.catalogoVariedades || {};
          if (!map[cellVal]) {
            issues.push({ kind: "value", message: ruleMessage(rule, "Código no válido en catálogo"), colIdx });
          }
        } else if (rule.tipo === "cruce_ipp_isp" && cellVal && options.cruceIppIsp) {
          const { tipo, lotesIPP, lotesISP } = options.cruceIppIsp;
          if (tipo === "IPP" && !lotesISP.includes(cellVal)) {
            issues.push({ kind: "value", message: `LOTE: ${cellVal} se tiene en IPP PERO NO EN ISP`, colIdx });
          } else if (tipo === "ISP" && !lotesIPP.includes(cellVal)) {
            issues.push({ kind: "value", message: `LOTE: ${cellVal} se tiene en ISP PERO NO EN IPP`, colIdx });
          }
        }
      });
    });
  }

  (config?.rangos_obligatorios || []).forEach((range) => {
    if (idx < range.desde_js || idx > range.hasta_js) return;
    if ((range.excepto_js || []).includes(idx)) return;
    if (!val) {
      issues.push({ kind: "empty", message: range.mensaje || "Campo obligatorio", colIdx: idx });
    }
  });

  return issues;
}

export function indicesToValidate(config) {
  const set = new Set(config?.columnas_a_revisar?.indices_js || []);
  (config?.columnas_visibles_frontend?.indices_js || []).forEach((i) => set.add(i));
  (config?.validaciones_por_columna || []).forEach((c) => {
    set.add(c.indice_js);
    if (c.hasta_indice_js != null) {
      for (let i = c.indice_js; i <= c.hasta_indice_js; i += 1) set.add(i);
    }
  });
  (config?.rangos_obligatorios || []).forEach((range) => {
    for (let i = range.desde_js; i <= range.hasta_js; i += 1) {
      if (!(range.excepto_js || []).includes(i)) set.add(i);
    }
  });
  return [...set];
}

export function findDuplicates(values) {
  const seen = new Set();
  const dups = new Set();
  values.forEach((v) => {
    if (!v) return;
    if (seen.has(v)) dups.add(v);
    else seen.add(v);
  });
  return [...dups];
}

export function applyReglasCompuestasFila(row, reglas, onIssue, helpers = {}) {
  if (!reglas) return;

  const cruzadas = reglas["validaciones-cruzadas"] || reglas["validaciones-compuestas"] || [];
  cruzadas.forEach((regla) => {
    const mensaje = regla["si-falla-mostrar"] || "Validación cruzada no cumplida";
    const tipo = regla.tipo;

    if (tipo === "suma-columnas") {
      const cols = regla.columnas || regla["columnas-a-sumary"] || [];
      const targetCol = regla["igual-a-columna"];
      const targetJs = typeof targetCol === "number" ? targetCol - 1 : null;
      const suma = cols.reduce((acc, colNum) => {
        const n = helpers.parseNumber?.(row[colNum - 1]) ?? parseFlexibleNumber(row[colNum - 1]);
        return acc + (Number.isFinite(n) ? n : 0);
      }, 0);
      const esperado = helpers.parseNumber?.(row[targetJs]) ?? parseFlexibleNumber(row[targetJs]);
      if (Number.isFinite(esperado) && Math.abs(suma - esperado) > 0.001) {
        cols.forEach((colNum) => onIssue(colNum - 1, { kind: "value", message: mensaje }));
        if (targetJs != null) onIssue(targetJs, { kind: "value", message: mensaje });
      }
    }

    if (tipo === "igual-entre-columnas") {
      const colA = (regla["columna-a"] ?? regla["columna-a-excel"]) - 1;
      const colB = (regla["columna-b"] ?? regla["columna-b-excel"]) - 1;
      const rawA = row[colA];
      const rawB = row[colB];
      if (!cellDisplayValue(rawA) && !cellDisplayValue(rawB)) return;
      const a = helpers.normalizeDate?.(rawA) || "";
      const b = helpers.normalizeDate?.(rawB) || "";
      const iguales =
        a && b ? a === b : cellDisplayValue(rawA) === cellDisplayValue(rawB);
      if (!iguales) {
        onIssue(colA, { kind: "value", message: mensaje });
        onIssue(colB, { kind: "value", message: mensaje });
      }
    }

    if (tipo === "igual-anio-mes-entre-columnas") {
      const colA = (regla["columna-a"] ?? regla["columna-a-excel"]) - 1;
      const colB = (regla["columna-b"] ?? regla["columna-b-excel"]) - 1;
      const rawA = row[colA];
      const rawB = row[colB];
      if (!cellDisplayValue(rawA) || !cellDisplayValue(rawB)) return;
      let a = helpers.normalizeDate?.(rawA) || "";
      let b = helpers.normalizeDate?.(rawB) || "";
      // normalizeDate PT a veces devuelve DD/MM/YYYY → convertir a ISO-like YM
      const toYm = (v, raw) => {
        const s = String(v || "").trim();
        if (/^\d{4}-\d{2}/.test(s)) return s.slice(0, 7);
        if (/^\d{2}[/-]\d{2}[/-]\d{4}$/.test(s)) {
          const [d, m, y] = s.split(/[/-]/);
          return `${y}-${m.padStart(2, "0")}`;
        }
        const fromHelper = helpers.normalizeDate?.(raw);
        if (fromHelper && /^\d{4}-\d{2}/.test(String(fromHelper))) return String(fromHelper).slice(0, 7);
        return "";
      };
      const ymA = toYm(a, rawA);
      const ymB = toYm(b, rawB);
      if (ymA && ymB && ymA !== ymB) {
        onIssue(colA, { kind: "value", message: mensaje });
        onIssue(colB, { kind: "value", message: mensaje });
      }
    }

    if (tipo === "fecha-menor-o-igual" || tipo === "fecha-no-mayor-que") {
      const colA = (regla["columna-a"] ?? regla["columna-a-excel"]) - 1;
      const colB = (regla["columna-b"] ?? regla["columna-b-excel"]) - 1;
      const a = helpers.normalizeDate?.(row[colA]) || cellDisplayValue(row[colA]);
      const b = helpers.normalizeDate?.(row[colB]) || cellDisplayValue(row[colB]);
      if (a && b && helpers.compareDates) {
        const cmp = helpers.compareDates(a, b);
        if (cmp !== null && cmp > 0) {
          onIssue(colA, { kind: "value", message: mensaje });
          onIssue(colB, { kind: "value", message: mensaje });
        }
      } else if (a && b && a > b) {
        // ISO yyyy-mm-dd: comparación lexicográfica = cronológica.
        onIssue(colA, { kind: "value", message: mensaje });
        onIssue(colB, { kind: "value", message: mensaje });
      }
    }
  });

  (reglas.columnas || []).forEach((col) => {
    if (col["no-mayor-que-columna"] == null) return;
    const idx = col.numero - 1;
    const refIdx =
      typeof col["no-mayor-que-columna"] === "number"
        ? col["no-mayor-que-columna"] - 1
        : Number(col["no-mayor-que-columna"].numero) - 1;
    const n = helpers.parseNumber?.(row[idx]) ?? parseFlexibleNumber(row[idx]);
    const ref = helpers.parseNumber?.(row[refIdx]) ?? parseFlexibleNumber(row[refIdx]);
    if (Number.isFinite(n) && Number.isFinite(ref) && n > ref) {
      onIssue(idx, {
        kind: "value",
        message: col["si-falla-mostrar"] || "No puede ser mayor que columna de referencia"
      });
    }
  });
}

export function paintCellValidation(td, issues, cssPrefix = "agv-mp") {
  let hasEmpty = false;
  let hasValue = false;
  const messages = [];

  issues.forEach((issue) => {
    if (issue.kind === "empty") hasEmpty = true;
    else hasValue = true;
    if (issue.message) messages.push(issue.message);
  });

  if (hasEmpty) td.classList.add(`${cssPrefix}-cell-error-empty`);
  if (hasValue) td.classList.add(`${cssPrefix}-cell-error-value`);

  if (messages.length) {
    td.title = [...new Set(messages)].join(" · ");
  }
}
