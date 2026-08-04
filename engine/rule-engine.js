/**
 * Capa 3 — Motor de dominio (CDA).
 * Evalúa reportes ingeridos según reglas JSON de rules/modulos/.
 */

import {
  evaluarCalibreArPt,
  evaluarClienteDestinoPt,
  evaluarTrazabilidadPt
} from "./pt-rule-validators.js";
import { resolverMensajeRegla } from "./cartilla-rules.adapter.js";

const CULTIVOS_VALIDOS = ["uva", "arandano", "esparrago", "palta"];

const CLAVES_METADATO_COLUMNA = new Set([
  "numero",
  "nombre-de-la-columna",
  "que-se-revisa",
  "aplica-a-cartilla"
]);

/** Cache en memoria: evita re-fetch de rules JSON en cada visita al módulo. */
const reglasCache = new Map();

export async function cargarReglasDesdeRuta(rutaRelativa) {
  const ruta = String(rutaRelativa || "").replace(/^\//, "");
  if (reglasCache.has(ruta)) {
    return reglasCache.get(ruta);
  }
  const response = await fetch(ruta);
  if (!response.ok) {
    throw new Error(`No se encontró el archivo de reglas: ${ruta}`);
  }
  const json = await response.json();
  reglasCache.set(ruta, json);
  return json;
}

export async function cargarReglas(cultivo, modulo) {
  const id = String(cultivo || "").toLowerCase();
  const moduloId = String(modulo || "").toLowerCase();

  if (!CULTIVOS_VALIDOS.includes(id)) {
    throw new Error(`No hay reglas definidas para el cultivo: ${cultivo}`);
  }

  if (!moduloId) {
    throw new Error("Debe indicar el módulo (mp, pt o plagas) para cargar las reglas");
  }

  return cargarReglasDesdeRuta(`rules/modulos/${id}-${moduloId}.rules.json`);
}

function valorVacio(valor) {
  return valor === null || valor === undefined || String(valor).trim() === "";
}

function normalizarTexto(valor) {
  return String(valor).trim();
}

function normalizarLista(valor) {
  return String(valor).trim().toUpperCase();
}

function esNumero(valor) {
  if (valorVacio(valor)) {
    return false;
  }
  const numero = Number(String(valor).replace(",", "."));
  return Number.isFinite(numero);
}

function aNumero(valor) {
  return Number(String(valor).replace(",", "."));
}

function crearDetalle(fila, valor, problema, tipo) {
  return {
    fila,
    valorEncontrado: valorVacio(valor) ? "" : String(valor),
    problema,
    tipo
  };
}

function resolverReferenciaColumna(ref) {
  if (typeof ref === "number" && Number.isFinite(ref)) {
    return { numero: ref };
  }
  if (typeof ref === "string" && /^\d+$/.test(ref.trim())) {
    return { numero: Number(ref.trim()) };
  }
  if (ref && typeof ref === "object") {
    if (ref.numero != null) {
      return { numero: Number(ref.numero), nombre: ref["nombre-de-la-columna"] || null };
    }
    if (ref["nombre-de-la-columna"]) {
      return { nombre: ref["nombre-de-la-columna"] };
    }
  }
  if (typeof ref === "string") {
    return { nombre: ref };
  }
  return {};
}

export function obtenerValorRegistro(registro, ref) {
  const { numero, nombre } = resolverReferenciaColumna(ref);

  if (numero != null && registro?._cols) {
    const porIndice = registro._cols[String(numero)];
    if (porIndice !== undefined) {
      return porIndice;
    }
  }

  if (nombre) {
    return registro?.[nombre];
  }

  return undefined;
}

export function tieneReglasActivas(reglaColumna) {
  if (!reglaColumna || typeof reglaColumna !== "object") {
    return false;
  }
  return Object.keys(reglaColumna).some((clave) => !CLAVES_METADATO_COLUMNA.has(clave));
}

function evaluarSumaColumnas(filas, reglaColumna, detalle, filasAfectadasSet) {
  const columnasSumar = reglaColumna["debe-ser-suma-de"];
  if (!Array.isArray(columnasSumar) || columnasSumar.length === 0) {
    return;
  }

  const tolerancia = reglaColumna.tolerancia ?? 0;
  const mensajeFallo = reglaColumna["si-falla-mostrar"] || "La sumatoria no coincide";
  const nombreResultado = reglaColumna["nombre-de-la-columna"] || `Columna ${reglaColumna.numero}`;

  filas.forEach((registro) => {
    const filaNum = registro.fila;
    const valorReportado = obtenerValorRegistro(registro, reglaColumna);
    const hayComponentes = columnasSumar.some((ref) => !valorVacio(obtenerValorRegistro(registro, ref)));

    if (valorVacio(valorReportado) && !hayComponentes) {
      return;
    }

    if (!esNumero(valorReportado)) {
      detalle.push(crearDetalle(filaNum, valorReportado, mensajeFallo, "suma"));
      filasAfectadasSet.add(filaNum);
      return;
    }

    let suma = 0;
    let componenteInvalido = false;

    columnasSumar.forEach((ref) => {
      const valor = obtenerValorRegistro(registro, ref);
      if (valorVacio(valor)) {
        return;
      }
      if (!esNumero(valor)) {
        componenteInvalido = true;
        return;
      }
      suma += aNumero(valor);
    });

    if (componenteInvalido) {
      detalle.push(
        crearDetalle(
          filaNum,
          valorReportado,
          `${mensajeFallo} (hay valores no numéricos en las columnas a sumar)`,
          "suma"
        )
      );
      filasAfectadasSet.add(filaNum);
      return;
    }

    const esperado = aNumero(valorReportado);
    if (Math.abs(suma - esperado) > tolerancia) {
      detalle.push(
        crearDetalle(
          filaNum,
          valorReportado,
          `${mensajeFallo}. ${nombreResultado}: ${esperado}, suma calculada: ${suma}`,
          "suma"
        )
      );
      filasAfectadasSet.add(filaNum);
    }
  });
}

function aplicaReglaACartilla(reglaColumna, contexto = {}) {
  const cartillas = reglaColumna["aplica-a-cartilla"];
  if (!Array.isArray(cartillas) || cartillas.length === 0) {
    return true;
  }
  return cartillas.includes(contexto.cartilla);
}

function parseFechaIso(valor) {
  // Serial Excel (número o texto) → ISO, para comparar aunque se haya formateado a DD/MM/YYYY.
  if (typeof valor === "number" && Number.isFinite(valor) && valor > 20000 && valor < 80000) {
    const fecha = new Date(Math.round((valor - 25569) * 86400 * 1000));
    if (Number.isNaN(fecha.getTime())) return "";
    const y = fecha.getUTCFullYear();
    const m = String(fecha.getUTCMonth() + 1).padStart(2, "0");
    const d = String(fecha.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  const texto = normalizarTexto(valor);
  if (!texto) return "";
  if (/^\d{5,6}(\.\d+)?$/.test(texto)) {
    const n = Number(texto);
    if (Number.isFinite(n) && n > 20000 && n < 80000) return parseFechaIso(n);
  }
  if (/^\d{8}$/.test(texto)) {
    return `${texto.slice(0, 4)}-${texto.slice(4, 6)}-${texto.slice(6, 8)}`;
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(texto)) {
    return texto.slice(0, 10);
  }
  if (/^\d{2}[/-]\d{2}[/-]\d{4}$/.test(texto)) {
    const [d, m, y] = texto.split(/[/-]/);
    return `${y}-${m}-${d}`;
  }
  if (/^\d{2}[/-]\d{2}[/-]\d{2}$/.test(texto)) {
    const [d, m, y] = texto.split(/[/-]/);
    const fullY = Number(y) <= 50 ? `20${y}` : `19${y}`;
    return `${fullY}-${m}-${d}`;
  }
  return "";
}

export function evaluarColumna(filas, reglaColumna, contexto = {}) {
  const nombreColumna = reglaColumna["nombre-de-la-columna"];
  if (!aplicaReglaACartilla(reglaColumna, contexto)) {
    return {
      nombreColumna,
      numeroColumna: reglaColumna.numero ?? null,
      estado: "ok",
      totalRevisadas: filas.length,
      totalFallidas: 0,
      filasAfectadas: [],
      detalle: []
    };
  }
  const esObligatorio = Boolean(reglaColumna["es-obligatorio"]);
  const msg = (tipo, valor, extras = {}) => {
    const texto = normalizarTexto(valor);
    return resolverMensajeRegla(
      reglaColumna,
      {
        valor,
        texto,
        detectado: texto.length,
        "longitud-detectada": texto.length,
        lote: texto,
        ...extras
      },
      tipo
    );
  };
  const detalle = [];
  const filasAfectadasSet = new Set();
  const conteoValores = new Map();

  filas.forEach((registro) => {
    const filaNum = registro.fila;
    const valor = obtenerValorRegistro(registro, reglaColumna);
    const texto = normalizarTexto(valor);

    if (valorVacio(valor)) {
      if (esObligatorio) {
        detalle.push(crearDetalle(filaNum, valor, msg("obligatorio", valor), "obligatorio"));
        filasAfectadasSet.add(filaNum);
      }
      return;
    }

    if (reglaColumna["debe-estar-vacio"] && !valorVacio(valor)) {
      detalle.push(crearDetalle(filaNum, valor, msg("vacio-requerido", valor), "vacio-requerido"));
      filasAfectadasSet.add(filaNum);
      return;
    }

    if (reglaColumna["igual-a-contexto"] === "fecha-lmr-mayoritaria") {
      const esperado = contexto.fechaLmrMayoritaria || "";
      if (parseFechaIso(valor) !== esperado) {
        detalle.push(crearDetalle(filaNum, valor, msg("igualdad", valor, { "valor-esperado": esperado }), "igualdad"));
        filasAfectadasSet.add(filaNum);
      }
    }

    if (reglaColumna["longitud-exacta"] != null && texto.length !== reglaColumna["longitud-exacta"]) {
      detalle.push(crearDetalle(filaNum, valor, msg("longitud", valor), "longitud"));
      filasAfectadasSet.add(filaNum);
    }

    if (reglaColumna["maximo-de-caracteres"] != null && texto.length > reglaColumna["maximo-de-caracteres"]) {
      detalle.push(crearDetalle(filaNum, valor, msg("longitud", valor), "longitud"));
      filasAfectadasSet.add(filaNum);
    }

    if (reglaColumna["patron-regex"]) {
      const patron = new RegExp(reglaColumna["patron-regex"]);
      if (!patron.test(texto)) {
        detalle.push(crearDetalle(filaNum, valor, msg("formato", valor), "formato"));
        filasAfectadasSet.add(filaNum);
      }
    }

    if (reglaColumna["igual-a-valor"] != null) {
      const esperado = String(reglaColumna["igual-a-valor"]).trim();
      const comparar = esNumero(valor) && esNumero(esperado) ? String(aNumero(valor)) : texto;
      const esperadoNorm = esNumero(esperado) ? String(aNumero(esperado)) : esperado;
      if (comparar !== esperadoNorm) {
        detalle.push(crearDetalle(filaNum, valor, msg("igualdad", valor), "igualdad"));
        filasAfectadasSet.add(filaNum);
      }
    }

    if (reglaColumna["igual-a-columna"] != null) {
      const otroValor = obtenerValorRegistro(registro, reglaColumna["igual-a-columna"]);
      const isoA = parseFechaIso(valor);
      const isoB = parseFechaIso(otroValor);
      const iguales =
        isoA && isoB
          ? isoA === isoB
          : normalizarTexto(valor) === normalizarTexto(otroValor);
      if (!iguales) {
        detalle.push(crearDetalle(filaNum, valor, msg("igualdad", valor), "igualdad"));
        filasAfectadasSet.add(filaNum);
      }
    }

    if (reglaColumna.minimo != null || reglaColumna.maximo != null) {
      if (!esNumero(valor)) {
        detalle.push(crearDetalle(filaNum, valor, msg("rango", valor), "rango"));
        filasAfectadasSet.add(filaNum);
      } else {
        const numero = aNumero(valor);
        if (reglaColumna.minimo != null && numero < reglaColumna.minimo) {
          detalle.push(crearDetalle(filaNum, valor, msg("rango", valor), "rango"));
          filasAfectadasSet.add(filaNum);
        }
        if (reglaColumna.maximo != null && numero > reglaColumna.maximo) {
          detalle.push(crearDetalle(filaNum, valor, msg("rango", valor), "rango"));
          filasAfectadasSet.add(filaNum);
        }
      }
    }

    if (Array.isArray(reglaColumna["valores-permitidos"])) {
      const permitidos = reglaColumna["valores-permitidos"].map((item) => normalizarLista(item));
      const comparar = esNumero(valor) ? String(aNumero(valor)) : normalizarLista(valor);
      const coincide = permitidos.some(
        (item) => item === comparar || item === normalizarLista(valor)
      );
      if (!coincide) {
        detalle.push(crearDetalle(filaNum, valor, msg("lista", valor), "lista"));
        filasAfectadasSet.add(filaNum);
      }
    }

    if (Array.isArray(reglaColumna["valores-prohibidos"])) {
      const comparar = esNumero(valor) ? String(aNumero(valor)) : texto;
      const prohibido = reglaColumna["valores-prohibidos"].some((item) => {
        const ref = String(item).trim();
        const refNorm = esNumero(ref) ? String(aNumero(ref)) : ref;
        return comparar === refNorm;
      });
      if (prohibido) {
        detalle.push(crearDetalle(filaNum, valor, msg("prohibido", valor), "prohibido"));
        filasAfectadasSet.add(filaNum);
      }
    }

    if (reglaColumna["debe-existir-en-catalogo"]) {
      const catalogoId = reglaColumna["debe-existir-en-catalogo"];
      const catalogo = contexto.catalogos?.[catalogoId];
      const entradas = catalogo?.entradas || {};
      const claves = Object.keys(entradas);
      const codigo = texto;
      const codigoPadded = /^\d+$/.test(codigo) ? codigo.padStart(2, "0") : codigo;
      const existe = claves.includes(codigo) || claves.includes(codigoPadded);
      if (!existe) {
        detalle.push(crearDetalle(filaNum, valor, msg("catalogo", valor), "catalogo"));
        filasAfectadasSet.add(filaNum);
      }
    }

    if (reglaColumna["no-puede-repetirse"]) {
      const agruparPor = reglaColumna["agrupar-por-columna"];
      const grupo = agruparPor != null ? normalizarTexto(obtenerValorRegistro(registro, agruparPor)) : "";
      const clave = `${grupo}::${normalizarLista(valor)}`;
      if (!conteoValores.has(clave)) {
        conteoValores.set(clave, []);
      }
      conteoValores.get(clave).push({ fila: filaNum, valor, grupo });
    }
  });

  if (reglaColumna["no-puede-repetirse"]) {
    conteoValores.forEach((apariciones) => {
      if (apariciones.length > 1) {
        apariciones.forEach(({ fila, valor }) => {
          detalle.push(crearDetalle(fila, valor, msg("duplicado", valor), "duplicado"));
          filasAfectadasSet.add(fila);
        });
      }
    });
  }

  evaluarSumaColumnas(filas, reglaColumna, detalle, filasAfectadasSet);

  const totalRevisadas = filas.length;
  const filasAfectadas = [...filasAfectadasSet].sort((a, b) => a - b);
  const totalFallidas = filasAfectadas.length;

  let estado = "ok";
  if (totalFallidas > 0) {
    estado = esObligatorio ? "critico" : "observado";
  }

  return {
    nombreColumna,
    numeroColumna: reglaColumna.numero ?? null,
    estado,
    totalRevisadas,
    totalFallidas,
    filasAfectadas,
    detalle
  };
}

function evaluarValidacionCompuesta(filas, regla, contexto = {}) {
  if (!aplicaReglaACartilla(regla, contexto)) {
    return {
      nombreColumna: regla.nombre || "Validación compuesta",
      numeroColumna: null,
      estado: "ok",
      totalRevisadas: filas.length,
      totalFallidas: 0,
      filasAfectadas: [],
      detalle: [],
      esCompuesta: true
    };
  }

  const tipo = regla.tipo;
  const mensajeFallo = regla["si-falla-mostrar"] || "Validación compuesta no cumplida";
  const detalle = [];
  const filasAfectadasSet = new Set();
  const nombre =
    regla.nombre || regla["nombre-de-la-columna"] || `Validación: ${tipo || "compuesta"}`;

  if (tipo === "igual-entre-columnas") {
    filas.forEach((registro) => {
      const valorA = obtenerValorRegistro(registro, regla["columna-a"]);
      const valorB = obtenerValorRegistro(registro, regla["columna-b"]);
      if (valorVacio(valorA) && valorVacio(valorB)) {
        return;
      }
      const isoA = parseFechaIso(valorA);
      const isoB = parseFechaIso(valorB);
      const iguales =
        isoA && isoB
          ? isoA === isoB
          : normalizarTexto(valorA) === normalizarTexto(valorB);
      if (!iguales) {
        detalle.push(crearDetalle(registro.fila, valorA, mensajeFallo, "compuesta"));
        filasAfectadasSet.add(registro.fila);
      }
    });
  }

  if (tipo === "fecha-no-mayor-que" || tipo === "fecha-menor-o-igual") {
    filas.forEach((registro) => {
      const valorA = obtenerValorRegistro(registro, regla["columna-a"]);
      const valorB = obtenerValorRegistro(registro, regla["columna-b"]);
      const isoA = parseFechaIso(valorA);
      const isoB = parseFechaIso(valorB);
      if (!isoA || !isoB) return;
      if (isoA > isoB) {
        detalle.push(crearDetalle(registro.fila, valorA, mensajeFallo, "compuesta"));
        filasAfectadasSet.add(registro.fila);
      }
    });
  }

  if (tipo === "diferencia-maxima-columnas") {
    const maximo = regla.maximo ?? 11.5;
    filas.forEach((registro) => {
      const valorA = obtenerValorRegistro(registro, regla["columna-a"]);
      const valorB = obtenerValorRegistro(registro, regla["columna-b"]);
      if (!esNumero(valorA) || !esNumero(valorB)) {
        return;
      }
      if (Math.abs(aNumero(valorB) - aNumero(valorA)) >= maximo) {
        detalle.push(crearDetalle(registro.fila, valorA, mensajeFallo, "compuesta"));
        filasAfectadasSet.add(registro.fila);
      }
    });
  }

  if (tipo === "suma-columnas") {
    const reglaSuma = {
      numero: regla["columna-resultado"],
      "nombre-de-la-columna": regla["nombre-de-la-columna"],
      "debe-ser-suma-de": regla["columnas-a-sumary"] || regla["debe-ser-suma-de"] || [],
      tolerancia: regla.tolerancia ?? 0,
      "si-falla-mostrar": mensajeFallo
    };
    evaluarSumaColumnas(filas, reglaSuma, detalle, filasAfectadasSet);
  }

  if (tipo === "cliente-destino-pt") {
    const result = evaluarClienteDestinoPt(filas, regla, contexto);
    detalle.push(...result.detalle);
    result.filasAfectadasSet.forEach((f) => filasAfectadasSet.add(f));
  }

  if (tipo === "calibre-ar-pt") {
    const result = evaluarCalibreArPt(filas, regla, contexto);
    detalle.push(...result.detalle);
    result.filasAfectadasSet.forEach((f) => filasAfectadasSet.add(f));
  }

  if (tipo === "trazabilidad-pt") {
    const result = evaluarTrazabilidadPt(filas, regla, contexto);
    detalle.push(...result.detalle);
    result.filasAfectadasSet.forEach((f) => filasAfectadasSet.add(f));
  }

  const filasAfectadas = [...filasAfectadasSet].sort((a, b) => a - b);
  const totalFallidas = filasAfectadas.length;

  return {
    nombreColumna: nombre,
    numeroColumna:
      regla["columna-resultado"] ?? regla["columna-lmr"] ?? regla["columna-a"] ?? null,
    estado: totalFallidas > 0 ? "observado" : "ok",
    totalRevisadas: filas.length,
    totalFallidas,
    filasAfectadas,
    detalle,
    esCompuesta: true
  };
}

export function evaluarValidacionesCompuestas(filas, reglasCompuestas = [], contexto = {}) {
  if (!Array.isArray(reglasCompuestas)) {
    return [];
  }
  return reglasCompuestas
    .filter((regla) => regla && typeof regla === "object")
    .map((regla) => evaluarValidacionCompuesta(filas, regla, contexto));
}

export function evaluarValidacionesArchivo(datos, reglas) {
  const validaciones = Array.isArray(reglas?.["validaciones-archivo"])
    ? reglas["validaciones-archivo"]
    : [];
  const detalle = [];

  validaciones.forEach((regla) => {
    const mensaje = regla["si-falla-mostrar"] || "Validación de archivo no cumplida";

    if (regla.tipo === "total-columnas") {
      const esperado = regla.valor ?? reglas["total-columnas"];
      const encontrado = datos?.meta?.totalColumnas ?? datos?.meta?.columnas?.length ?? 0;
      if (esperado != null && Number(encontrado) !== Number(esperado)) {
        detalle.push({
          fila: 0,
          valorEncontrado: String(encontrado),
          problema: `${mensaje} (esperado: ${esperado}, encontrado: ${encontrado})`,
          tipo: "archivo"
        });
      }
    }

    if (regla.tipo === "celda-cabecera") {
      const fila = Number(regla.fila);
      const columna = Number(regla.columna);
      const esperado = normalizarTexto(regla["valor-esperado"]);
      const encontrado = normalizarTexto(
        datos?.meta?.celdasCabecera?.[`${fila}:${columna}`]
      );
      if (esperado && encontrado !== esperado) {
        detalle.push({
          fila: 0,
          valorEncontrado: encontrado,
          problema: `${mensaje} (esperado: ${esperado}, encontrado: ${encontrado || "vacío"})`,
          tipo: "archivo"
        });
      }
    }
  });

  if (reglas?.["total-columnas"] != null && validaciones.length === 0) {
    const esperado = reglas["total-columnas"];
    const encontrado = datos?.meta?.totalColumnas ?? datos?.meta?.columnas?.length ?? 0;
    if (Number(encontrado) !== Number(esperado)) {
      detalle.push({
        fila: 0,
        valorEncontrado: String(encontrado),
        problema: `El archivo debe tener ${esperado} columnas (tiene ${encontrado})`,
        tipo: "archivo"
      });
    }
  }

  return detalle;
}

export async function cargarCatalogosParaReglas(reglas) {
  const ids = new Set();

  (reglas?.columnas || []).forEach((columna) => {
    if (columna["debe-existir-en-catalogo"]) {
      ids.add(columna["debe-existir-en-catalogo"]);
    }
  });

  (reglas?.["validaciones-compuestas"] || []).forEach((regla) => {
    ["catalogo", "catalogo-calibres", "catalogo-variedades"].forEach((key) => {
      if (regla?.[key]) ids.add(regla[key]);
    });
  });

  const catalogos = {};
  await Promise.all(
    [...ids].map(async (id) => {
      const ruta = id.includes("/")
        ? id
        : `presentation/data/catalogos/${id}.json`;
      const response = await fetch(ruta);
      if (response.ok) {
        catalogos[id] = await response.json();
      }
    })
  );

  return catalogos;
}

export function analizarReporte(datos, reglas, contexto = {}) {
  const filas = Array.isArray(datos?.filas) ? datos.filas : [];
  const columnasReglas = Array.isArray(reglas?.columnas) ? reglas.columnas : [];

  const erroresArchivo = evaluarValidacionesArchivo(datos, reglas);
  const columnasConReglas = columnasReglas.filter((regla) => tieneReglasActivas(regla));
  const columnasDetalle = columnasConReglas.map((regla) =>
    evaluarColumna(filas, regla, contexto)
  );
  const compuestasDetalle = evaluarValidacionesCompuestas(
    filas,
    reglas?.["validaciones-compuestas"],
    contexto
  );

  if (erroresArchivo.length > 0) {
    columnasDetalle.unshift({
      nombreColumna: "Archivo",
      numeroColumna: null,
      estado: "critico",
      totalRevisadas: 1,
      totalFallidas: erroresArchivo.length,
      filasAfectadas: [],
      detalle: erroresArchivo,
      esArchivo: true
    });
  }

  const columnasDetalleFinal = [...columnasDetalle, ...compuestasDetalle];

  const columnasOk = columnasDetalleFinal.filter((item) => item.estado === "ok").length;
  const columnasObservadas = columnasDetalleFinal.filter((item) => item.estado === "observado").length;
  const columnasCriticas = columnasDetalleFinal.filter((item) => item.estado === "critico").length;
  const totalColumnas = columnasDetalleFinal.length;
  const porcentajeAprobacion =
    totalColumnas === 0 ? 100 : Math.round((columnasOk / totalColumnas) * 100);

  return {
    cultivo: datos?.cultivo || reglas?.cultivo || "",
    totalColumnas,
    totalColumnasCatalogo: columnasReglas.length,
    columnasConReglas: columnasConReglas.length + compuestasDetalle.length,
    columnasOk,
    columnasObservadas,
    columnasCriticas,
    porcentajeAprobacion,
    columnasDetalle: columnasDetalleFinal,
    meta: datos?.meta || {}
  };
}
