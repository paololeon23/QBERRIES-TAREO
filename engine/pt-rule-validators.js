/** Validadores compuestos PT Arándano — usados desde rule-engine.js */

function normalizarTexto(valor) {
  return String(valor ?? "").trim();
}

function normalizarDestino(destino) {
  const map = {
    "MEDIO ORIENTE": "ASIA",
    ASIA: "ASIA",
    EUROPA: "EUROPA",
    CANADA: "CANADA",
    USA: "USA",
    BRASIL: "LATAM",
    COLOMBIA: "LATAM",
    LATAM: "LATAM"
  };
  const upper = normalizarTexto(destino).toUpperCase();
  return map[upper] || upper;
}

function esClienteEspecial(cliente) {
  const c = normalizarTexto(cliente);
  return c === "THE FRUITIST CO" || c === "COSTCO";
}

function usaCalibreEspecial(cliente, destino) {
  const c = normalizarTexto(cliente);
  const d = normalizarDestino(destino);
  if (d === "EUROPA" && (c === "COSTCO" || c === "COSTCO UK")) return false;
  return esClienteEspecial(c);
}

function esSubgrupoNA(val) {
  return normalizarTexto(val).toUpperCase() === "N/A";
}

function crearDetalle(fila, valor, problema, tipo) {
  return {
    fila,
    valorEncontrado: valor === null || valor === undefined ? "" : String(valor),
    problema,
    tipo
  };
}

function obtenerValor(registro, colNum) {
  return registro?._cols?.[String(colNum)];
}

function parseFechaJuliano(fechaTexto) {
  if (!fechaTexto) return null;
  const partes = String(fechaTexto).includes("/")
    ? fechaTexto.split("/")
    : fechaTexto.split("-").reverse();
  const [dd, mm, yyyy] = partes.map(Number);
  if (!yyyy || !mm || !dd) return null;
  const fechaObj = new Date(yyyy, mm - 1, dd);
  const start = new Date(yyyy, 0, 0);
  return String(Math.floor((fechaObj - start) / (1000 * 60 * 60 * 24))).padStart(3, "0");
}

function extraerTrazabilidad(codigo) {
  const texto = normalizarTexto(codigo);
  if (texto.length < 13) return null;
  return {
    variedad: texto.substring(8, 10),
    juliano: texto.substring(10),
    productor: texto.charAt(4)
  };
}

export function evaluarClienteDestinoPt(filas, regla, contexto) {
  const colCliente = regla["columna-cliente"];
  const colDestino = regla["columna-destino"];
  const catalogo = contexto.catalogos?.[regla.catalogo]?.entradas || [];
  const detalle = [];
  const filasAfectadasSet = new Set();
  const pares = Array.isArray(catalogo) ? catalogo : Object.values(catalogo);

  filas.forEach((registro) => {
    const cliente = normalizarTexto(obtenerValor(registro, colCliente));
    const destino = normalizarTexto(obtenerValor(registro, colDestino));
    if (!cliente || !destino) return;

    let invalido = false;
    if (cliente === "THE FRUITIST CO" && !["USA", "CANADA"].includes(destino)) invalido = true;
    else if (cliente === "FRUITIST SHANGHAI" && destino !== "ASIA") invalido = true;
    else if (cliente === "COSTCO" && !["USA", "CANADA", "EUROPA"].includes(destino)) invalido = true;
    else {
      const destinosPermitidos = pares
        .filter((p) => normalizarTexto(p.cliente).toUpperCase() === cliente.toUpperCase())
        .map((p) => normalizarDestino(p.destino));
      if (destinosPermitidos.length && !destinosPermitidos.includes(normalizarDestino(destino))) {
        invalido = true;
      }
    }

    if (invalido) {
      detalle.push(
        crearDetalle(
          registro.fila,
          destino,
          `${cliente} solo acepta destinos permitidos (actual: ${destino})`,
          "compuesta"
        )
      );
      filasAfectadasSet.add(registro.fila);
    }
  });

  return { detalle, filasAfectadasSet, colDestino };
}

export function evaluarCalibreArPt(filas, regla, contexto) {
  const colCalibre = regla["columna-calibre"];
  const colSubgrupo = regla["columna-subgrupo"];
  const colCliente = regla["columna-cliente"];
  const colDestino = regla["columna-destino"];
  const cat = contexto.catalogos?.[regla["catalogo-calibres"]] || {};
  const letras = cat.letras || {};
  const categorias = cat.categorias_frutist || {};
  const prohibidas = cat.palabras_prohibidas_regular || [];
  const detalle = [];
  const filasAfectadasSet = new Set();

  filas.forEach((registro) => {
    const cliente = normalizarTexto(obtenerValor(registro, colCliente));
    const destino = colDestino ? normalizarTexto(obtenerValor(registro, colDestino)) : "";
    const calibre = normalizarTexto(obtenerValor(registro, colCalibre));
    const subgrupo = normalizarTexto(obtenerValor(registro, colSubgrupo));
    if (!cliente) return;

    if (usaCalibreEspecial(cliente, destino)) {
      const cats = ["JUMBO", "MIXED", "REGULAR", "SUPER JUMBO"];
      if (!calibre) {
        detalle.push(crearDetalle(registro.fila, calibre, "Calibre AR vacío (obligatorio)", "compuesta"));
        filasAfectadasSet.add(registro.fila);
      } else if (!cats.includes(calibre.toUpperCase())) {
        detalle.push(crearDetalle(registro.fila, calibre, "Calibre debe ser JUMBO/MIXED/REGULAR/SUPER JUMBO", "compuesta"));
        filasAfectadasSet.add(registro.fila);
      }
      if (!subgrupo) {
        detalle.push(crearDetalle(registro.fila, subgrupo, "Subgrupo vacío (obligatorio)", "compuesta"));
        filasAfectadasSet.add(registro.fila);
      } else if (esSubgrupoNA(subgrupo)) {
        detalle.push(crearDetalle(registro.fila, subgrupo, "Subgrupo no puede ser N/A para cliente especial", "compuesta"));
        filasAfectadasSet.add(registro.fila);
      } else if (calibre && !categorias[calibre.toUpperCase()]?.includes(subgrupo)) {
        detalle.push(crearDetalle(registro.fila, subgrupo, `Subgrupo ${subgrupo} no pertenece a ${calibre}`, "compuesta"));
        filasAfectadasSet.add(registro.fila);
      }
    } else {
      if (!calibre) {
        detalle.push(crearDetalle(registro.fila, calibre, "Falta Calibre AR", "compuesta"));
        filasAfectadasSet.add(registro.fila);
      } else {
        const trad = letras[calibre] || calibre;
        if (prohibidas.includes(trad)) {
          detalle.push(crearDetalle(registro.fila, calibre, "Calibre no permitido para cliente regular", "compuesta"));
          filasAfectadasSet.add(registro.fila);
        }
      }
      if (!esSubgrupoNA(subgrupo)) {
        const msg = !subgrupo
          ? "Subgrupo vacío — debe ser N/A"
          : `Subgrupo debe ser N/A (valor incorrecto: ${subgrupo})`;
        detalle.push(crearDetalle(registro.fila, subgrupo, msg, "compuesta"));
        filasAfectadasSet.add(registro.fila);
      }
    }
  });

  return { detalle, filasAfectadasSet, cols: [colCalibre, colSubgrupo] };
}

export function evaluarTrazabilidadPt(filas, regla, contexto) {
  const colTraz = regla["columna-trazabilidad"];
  const colFecha = regla["columna-fecha-inspeccion"];
  const colCliente = regla["columna-cliente"];
  const variedades = contexto.catalogos?.[regla["catalogo-variedades"]]?.entradas || {};
  const detalle = [];
  const filasAfectadasSet = new Set();

  filas.forEach((registro) => {
    const traz = normalizarTexto(obtenerValor(registro, colTraz));
    const fecha = normalizarTexto(obtenerValor(registro, colFecha));
    const cliente = normalizarTexto(obtenerValor(registro, colCliente));
    if (!traz) return;

    const parsed = extraerTrazabilidad(traz);
    const julianoFecha = parseFechaJuliano(fecha);
    if (parsed && julianoFecha) {
      const actual = parseInt(julianoFecha, 10);
      const trazJul = parseInt(parsed.juliano, 10);
      const esDriscoll = cliente.toUpperCase().startsWith("DRISCOLL");
      const bad = esDriscoll ? trazJul !== actual && trazJul !== actual + 1 : trazJul !== actual;
      if (bad) {
        detalle.push(crearDetalle(registro.fila, traz, `Día juliano (${parsed.juliano}) no coincide con fecha`, "compuesta"));
        filasAfectadasSet.add(registro.fila);
      }
    }

    if (traz.length > 13) {
      detalle.push(crearDetalle(registro.fila, traz, "Trazabilidad excede 13 caracteres", "compuesta"));
      filasAfectadasSet.add(registro.fila);
    }

    const varKey = parsed?.variedad;
    const varEntry = variedades[varKey] || variedades[String(varKey).padStart(2, "0")];
    if (parsed && !varEntry) {
      detalle.push(crearDetalle(registro.fila, traz, `Código variedad (${varKey}) no existe`, "compuesta"));
      filasAfectadasSet.add(registro.fila);
    }

    // Desactivado por ahora: regla letra E en Sekoya Pop Orgánica ya no aplica
    // const varNombre = varEntry?.variedad || "";
    // if (varNombre === "Sekoya Pop Orgánica" && traz[4] !== "E") {
    //   detalle.push(crearDetalle(registro.fila, traz, "Falta letra E en trazabilidad (Sekoya Pop Orgánica)", "compuesta"));
    //   filasAfectadasSet.add(registro.fila);
    // }
  });

  return { detalle, filasAfectadasSet, cols: [colTraz] };
}

export function buildPtCompuestaColumnMap(reglasCompuestas = []) {
  const map = new Map();
  reglasCompuestas.forEach((regla) => {
    const msg = regla["si-falla-mostrar"] || "";
    if (regla.tipo === "cliente-destino-pt") {
      map.set(msg, [regla["columna-cliente"], regla["columna-destino"]]);
    }
    if (regla.tipo === "calibre-ar-pt") {
      map.set(msg, [regla["columna-calibre"], regla["columna-subgrupo"]]);
    }
    if (regla.tipo === "trazabilidad-pt") {
      map.set(msg, [regla["columna-trazabilidad"]]);
    }
    if (
      regla.tipo === "igual-anio-mes-entre-columnas" ||
      regla.tipo === "igual-entre-columnas" ||
      regla.tipo === "fecha-no-mayor-que" ||
      regla.tipo === "fecha-menor-o-igual"
    ) {
      map.set(msg, [regla["columna-a"], regla["columna-b"]].filter(Boolean));
    }
  });
  return map;
}
