function main() {
  ordenarYAgruparResultados();
}

function ordenarYAgruparResultados() {
  // Usar la planilla activa por defecto (más simple para ejecutar en Apps Script).
  // Si se define un ID, se usará ese documento en su lugar.
  const SPREADSHEET_ID = "";
  var spreadsheet = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  var hojaResultados = spreadsheet.getSheetByName("Resultados");
  var hojaPuntuacionEquipos = spreadsheet.getSheetByName("Puntuacion Equipos");

  hojaResultados.clear();
  hojaPuntuacionEquipos.clear();

  hojaPuntuacionEquipos.appendRow(["Equipo", "Puntuación Total"]);

  var datos = leerDatos(spreadsheet.getSheetByName("Carga"));
  var resultadosAgrupados = agruparYOrdenarResultados(datos);
  var puntuacionEquipos = procesarResultadosYEscribir(
    hojaResultados,
    resultadosAgrupados
  );

  escribirPuntuacionEquiposOrdenada(hojaPuntuacionEquipos, puntuacionEquipos);
}

var categoriasPorTiempo = {
  "3 - 50 mts Pecho":    
  [{cat: "a", min: 0, max: 39.99}, {cat: "b", min: 40, max: 44.99}, {cat: "c", min: 45, max: 49.99}, {cat: "d", min: 50, max: 54.99}, {cat: "e", min: 55, max: 59.99}, {cat: "f", min: 60, max: Infinity}],
  "2 - 50 mts Espalda":  
  [{cat: "a", min: 0, max: 37.99}, {cat: "b", min: 38, max: 42.99}, {cat: "c", min: 43, max: 47.99}, {cat: "d", min: 48, max: 52.99}, {cat: "e", min: 53, max: 57.99}, {cat: "f", min: 58, max: Infinity}],
  "1 - 50 mts Libre":    
  [{cat: "a", min: 0, max: 29.99}, {cat: "b", min: 30, max: 34.99}, {cat: "c", min: 35, max: 39.99}, {cat: "d", min: 40, max: 44.99}, {cat: "e", min: 45, max: 49.99}, {cat: "f", min: 50, max: Infinity}],
  "4 - 50 mts Mariposa":
   [{cat: "a", min: 0, max: 34.99}, {cat: "b", min: 35, max: 39.99}, {cat: "c", min: 40, max: 44.99}, {cat: "d", min: 45, max: 49.99}, {cat: "e", min: 50, max: 54.99}, {cat: "f", min: 55, max: Infinity}],
  "5 - 100 mts Medley":   
  [{cat: "a", min: 0, max: 79.99}, {cat: "b", min: 80, max: 87.99}, {cat: "c", min: 88, max: 95.99}, {cat: "d", min: 96, max: 103.99}, {cat: "e", min: 104, max: 111.99}, {cat: "f", min: 112, max: Infinity}]
};


// Asigna la categoría A–F basada en el tiempo y la prueba (solo individuales)
function asignarCategoria(prueba, tiempoTotal) {
  var rangos = categoriasPorTiempo[prueba];
  if (rangos) {
    for (var i = 0; i < rangos.length; i++) {
      var rango = rangos[i];
      if (tiempoTotal >= rango.min && tiempoTotal <= rango.max) {
        return rango.cat;
      }
    }
  }
  return "Sin categoría";
}

// 📌 Funciones getter para acceder a los valores de cada fila
function getEquipo(fila) { return /** @type {string} */ (fila[0]); }
function getPrueba(fila) { return /** @type {string} */ (fila[1]); }
function getNombre(fila) { return /** @type {string} */ (fila[2]); }
function getSerie(fila) { return /** @type {number} */ (fila[3]); }
function getAndarivel(fila) { return /** @type {number} */ (fila[4]); }
function getMinutos(fila) { return /** @type {number} */ (fila[5] || 0); }
function getSegundos(fila) { return /** @type {number} */ (fila[6] || 0); }
function getCentesimas(fila) { return /** @type {number} */ (fila[7] || 0); }
function getMetros(fila) { return /** @type {number} */ (fila[10] || 0); }
function getCategoria(fila) { return /** @type {string} */ (fila[9]); }

/**
 * @typedef {Object} Nadador
 * @property {string} equipo - Nombre del equipo
 * @property {string} prueba - Tipo de prueba (pecho, libre, etc.)
 * @property {string} nombre - Nombre del nadador
 * @property {string} categoria - Categoría (según planilla)
 * @property {number} serie - Número de serie
 * @property {number} andarivel - Número de andarivel
 * @property {number} minutos - Minutos registrados
 * @property {number} segundos - Segundos registrados
 * @property {number} centesimas - Centesimas registradas
 * @property {number} tiempoTotal - Tiempo total en segundos
 * @property {number|string} metros - Metros (solo para Americana)
 */

/**
 * Lee los datos de la hoja de cálculo y los convierte en un array de objetos Nadador
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaDatos - Hoja de Google Sheets con los tiempos
 * @returns {Nadador[]} - Lista de nadadores con sus datos estructurados
 */
function leerDatos(hojaDatos) {
  return hojaDatos.getRange(2, 1, hojaDatos.getLastRow() - 1, hojaDatos.getLastColumn())
    .getValues()
    .map(function (fila) {
      var prueba = getPrueba(fila);
      var minutos = getMinutos(fila);
      var segundos = getSegundos(fila);
      var centesimas = getCentesimas(fila);
      var tiempoTotal = minutos * 60 + segundos + centesimas / 100;
      var metros = Number(getMetros(fila)) || 0;

      var categoriaDefinida = getCategoria(fila);
      var categoria;

      if (esPosta(prueba)) {
        // Postas (4x50 Medley, Americana, etc.): la categoría viene definida por edades en la planilla
        categoria = categoriaDefinida || "Sin categoría";
      } else {
        // Pruebas individuales: la categoría se calcula siempre por tiempo según el reglamento
        categoria = asignarCategoria(prueba, tiempoTotal);
      }

      return {
        equipo: getEquipo(fila),
        prueba: prueba,
        nombre: getNombre(fila),
        categoria: categoria,
        serie: getSerie(fila),
        andarivel: getAndarivel(fila),
        minutos: minutos,
        segundos: segundos,
        centesimas: centesimas,
        tiempoTotal: prueba === "Americana" ? "" : tiempoTotal,
        metros: prueba === "Americana" ? metros : ""
      };
    })
    .filter(fila => fila.prueba === "Americana" ? (fila.metros || 0) > 0 : fila.tiempoTotal > 0);
}


const PUNTUACION_2026 = [9, 7, 6, 5, 4, 3, 2, 1];

function esPosta(prueba) {
  if (!prueba) return false;
  var p = prueba.toString().trim();
  if (p === "Americana") return true;
  if (/^\d+\s*x/i.test(p)) return true; // ej: 4x50 Medley
  var lower = p.toLowerCase();
  return lower.includes("posta") || lower.includes("relevo");
}

function calcularPuntuacion(prueba, posicion) {
  var base = PUNTUACION_2026[posicion - 1] || 0;
  return esPosta(prueba) ? base * 2 : base;
}

function escribirPuntuacionEquiposOrdenada(hoja, puntuacionEquipos) {
  var datos = Object.entries(puntuacionEquipos)
    .sort((a, b) => b[1] - a[1]); // Ordenar de mayor a menor puntuación

  if (datos.length > 0) {
    var rango = hoja.getRange(2, 1, datos.length, 2); // Seleccionar el rango exacto
    rango.setValues(datos); // Escribir todo de una vez
  }
}

function agruparYOrdenarResultados(resultados) {
  var agrupado = {};

  // Agrupar por prueba y categoría
  resultados.forEach(fila => {
    var key = fila.prueba + '-' + fila.categoria;
    if (!agrupado[key]) agrupado[key] = [];
    agrupado[key].push(fila);
  });

  // Ordenar las claves correctamente (prueba + categoría)
  var clavesOrdenadas = Object.keys(agrupado).sort((a, b) => {
    var indexA = a.lastIndexOf('-'); // Última aparición de '-'
    var indexB = b.lastIndexOf('-');

    var pruebaA = a.substring(0, indexA); // Extrae la parte de prueba
    var pruebaB = b.substring(0, indexB);
    var categoriaA = a.substring(indexA + 1); // Extrae la parte de categoría
    var categoriaB = b.substring(indexB + 1);

    // Ordenar primero por prueba, luego por categoría correctamente (a, b, c, d, e, f)
    if (pruebaA !== pruebaB) {
      return pruebaA.localeCompare(pruebaB);
    }
    return categoriaA.localeCompare(categoriaB, undefined, { numeric: true });
  });

  // Crear un nuevo objeto con las categorías ordenadas
  var agrupadoOrdenado = {};
  clavesOrdenadas.forEach(key => {
    agrupadoOrdenado[key] = agrupado[key].sort((a, b) => {
      if (a.prueba === "Americana") return (b.metros || 0) - (a.metros || 0); // desc por metros
      return (a.tiempoTotal || 0) - (b.tiempoTotal || 0); // asc por tiempo
    });
  });

  return agrupadoOrdenado;
}

/**
 * Procesa los resultados agrupados, asigna posiciones y puntuaciones, 
 * y escribe los datos en la hoja de resultados.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaResultados - Hoja donde se escribirán los resultados ordenados.
 * @param {Object} resultadosAgrupados - Resultados agrupados por prueba y categoría.
 * @returns {Object} - Objeto con la puntuación acumulada por equipo.
 */

function procesarResultadosYEscribir(hojaResultados, resultadosAgrupados) {
  hojaResultados.appendRow(['Equipo', 'Prueba', 'Nombre y Apellido', 'Categoría', 'Serie', 'Andarivel',
    'Minutos', 'Segundos', 'Centesimas', 'Metros', 'Tiempo Total', 'Puesto', 'Puntuación']);

  var puntuacionEquipos = {};
  var datosAEscribir = []; // Almacena todas las filas antes de escribirlas en la hoja
  var formatos = []; // Almacena los colores de fondo para cada fila
  var bordes = []; // Almacena las filas donde se aplicarán bordes

  var filaActual = 2; // Comienza en la segunda fila (la primera es el encabezado)

  for (var key in resultadosAgrupados) {
    var posicion = 1;
    var inicioGrupo = filaActual; // Marca el inicio del grupo de prueba/categoría

    resultadosAgrupados[key].forEach(function (fila, index) {
      var empate = false;
      if (index > 0) {
        var anterior = resultadosAgrupados[key][index - 1];
        empate = fila.prueba === "Americana"
          ? (fila.metros || 0) === (anterior.metros || 0)
          : (fila.tiempoTotal || 0) === (anterior.tiempoTotal || 0);
      }

      var posicionActual = empate ? posicion - 1 : posicion;
      var puntuacion = calcularPuntuacion(fila.prueba, posicionActual);

      // Definir color de fondo según la posición
      var color = "#FFFFFF"; // Blanco por defecto
      if (posicionActual === 1) color = "#FFD700"; // Oro
      if (posicionActual === 2) color = "#C0C0C0"; // Plata
      if (posicionActual === 3) color = "#CD7F32"; // Bronce

      // Agregar la fila y su color a los arrays
      datosAEscribir.push([
        fila.equipo, fila.prueba, fila.nombre, fila.categoria, fila.serie, fila.andarivel,
        fila.minutos, fila.segundos, fila.centesimas, fila.metros, formatearTiempo(fila.tiempoTotal), posicionActual, puntuacion
      ]);
      formatos.push(new Array(13).fill(color)); // Se aplica el color a toda la fila

      puntuacionEquipos[fila.equipo] = (puntuacionEquipos[fila.equipo] || 0) + puntuacion;
      if (!empate) posicion++;
      filaActual++;
    });

    // Almacenar la última fila del grupo para aplicar bordes
    bordes.push(filaActual - 1);
  }

  // Escribir los datos en la hoja en una sola llamada (más eficiente)
  if (datosAEscribir.length > 0) {
    var rangoDatos = hojaResultados.getRange(2, 1, datosAEscribir.length, datosAEscribir[0].length);
    rangoDatos.setValues(datosAEscribir);
    rangoDatos.setBackgrounds(formatos); // Aplica los colores de fondo

    // Forzar que la columna "Tiempo Total" (columna 11) se trate como texto (MM:SS:CC)
    var rangoTiempoTotal = hojaResultados.getRange(2, 11, datosAEscribir.length, 1);
    rangoTiempoTotal.setNumberFormat("@");
  }

  // Aplicar bordes gruesos al final de cada grupo de prueba/categoría
  bordes.forEach(fila => {
    hojaResultados.getRange(fila, 1, 1, 13).setBorder(false, false, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  return puntuacionEquipos;
}

// Convertir tiempo en segundos decimales a formato MM:SS:CC
function formatearTiempo(tiempoEnSegundos) {
  if (!tiempoEnSegundos || tiempoEnSegundos === "" || tiempoEnSegundos === 0) {
    return "";
  }

  var minutos = Math.floor(tiempoEnSegundos / 60);
  var segundos = Math.floor(tiempoEnSegundos % 60);
  var centesimas = Math.round((tiempoEnSegundos % 1) * 100);

  var minutosStr = minutos.toString();
  var segundosStr = segundos < 10 ? "0" + segundos : segundos.toString();
  var centesimasStr = centesimas < 10 ? "0" + centesimas : centesimas.toString();

  return minutosStr + ":" + segundosStr + ":" + centesimasStr;
}
