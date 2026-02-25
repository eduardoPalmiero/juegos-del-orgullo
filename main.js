function main() {
  // CONFIGURACIÓN DEL SPREADSHEET - REEMPLAZA EL ID CON EL TUYO
  const SPREADSHEET_ID = ""; // Reemplaza con el ID real
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hojaResultados = spreadsheet.getSheetByName('Resultados');
  var hojaPuntuacionEquipos = spreadsheet.getSheetByName('Puntuacion Equipos');

  // Limpiar hojas antes de escribir datos
  hojaResultados.clear();
  hojaPuntuacionEquipos.clear();

  hojaPuntuacionEquipos.appendRow(['Equipo', 'Puntuación Total']);

  // Procesar datos
  var datos = leerDatos(spreadsheet.getSheetByName("Carga"));
  var resultadosAgrupados = agruparYOrdenarResultados(datos);
  var puntuacionEquipos = procesarResultadosYEscribir(hojaResultados, resultadosAgrupados);

  // Escribir puntuación de equipos
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


// Asigna la categoría basada en el tiempo y la prueba
function asignarCategoria(prueba, tiempoTotal, categoriaDefinida) {
  if (prueba === "4x50 Medley") {
    return categoriaDefinida || "Sin categoría"; // Usa la categoría de la columna 9
  }
  
  if (categoriasPorTiempo[prueba]) {
    for (var i = 0; i < categoriasPorTiempo[prueba].length; i++) {
      var rango = categoriasPorTiempo[prueba][i];
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

/**
 * @typedef {Object} Nadador
 * @property {string} equipo - Nombre del equipo
 * @property {string} prueba - Tipo de prueba (pecho, libre, etc.)
 * @property {string} nombre - Nombre del nadador
 * @property {string} categoria - Categoría asignada según tiempos
 * @property {number} serie - Número de serie
 * @property {number} andarivel - Número de andarivel
 * @property {number} minutos - Minutos registrados
 * @property {number} segundos - Segundos registrados
 * @property {number} centesimas - Centesimas registradas
 * @property {number} tiempoTotal - Tiempo total en segundos
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
      var minutos = getMinutos(fila);
      var segundos = getSegundos(fila);
      var centesimas = getCentesimas(fila);
      var tiempoTotal = minutos * 60 + segundos + centesimas / 100;
      var categoriaDefinida = /** @type {string} */ (fila[9]); // Nueva columna con la categoría manual

      var categoria = asignarCategoria(getPrueba(fila), tiempoTotal, categoriaDefinida);

      return {
        equipo: getEquipo(fila),
        prueba: getPrueba(fila),
        nombre: getNombre(fila),
        categoria: categoria, // Ahora puede venir de la hoja o calcularse
        serie: getSerie(fila),
        andarivel: getAndarivel(fila),
        minutos: minutos,
        segundos: segundos,
        centesimas: centesimas,
        tiempoTotal: tiempoTotal
      };
    })
    .filter(fila => fila.tiempoTotal > 0);
}


const PUNTUACION_DEFAULT = [10, 8, 6, 5, 4, 3, 2, 1];
const PUNTUACION_4X50 = [20, 16, 12, 10, 8, 6, 4, 2];

function calcularPuntuacion(prueba, posicion) {
  return prueba === '4x50 Medley' 
    ? PUNTUACION_4X50[posicion - 1] || 0
    : PUNTUACION_DEFAULT[posicion - 1] || 0;
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
    agrupadoOrdenado[key] = agrupado[key].sort((a, b) => a.tiempoTotal - b.tiempoTotal);
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
    'Minutos', 'Segundos', 'Centesimas', 'Tiempo Total', 'Puesto', 'Puntuación']);

  var puntuacionEquipos = {};
  var datosAEscribir = []; // Almacena todas las filas antes de escribirlas en la hoja
  var formatos = []; // Almacena los colores de fondo para cada fila
  var bordes = []; // Almacena las filas donde se aplicarán bordes

  var filaActual = 2; // Comienza en la segunda fila (la primera es el encabezado)

  for (var key in resultadosAgrupados) {
    var posicion = 1;
    var inicioGrupo = filaActual; // Marca el inicio del grupo de prueba/categoría

    resultadosAgrupados[key].forEach(function (fila, index) {
      var puntuacion = calcularPuntuacion(fila.prueba, posicion);
      var empate = index > 0 && fila.tiempoTotal === resultadosAgrupados[key][index - 1].tiempoTotal;
      if (empate) posicion--;

      // Definir color de fondo según la posición
      var color = "#FFFFFF"; // Blanco por defecto
      if (posicion === 1) color = "#FFD700"; // 🥇 Oro (Dorado)
      if (posicion === 2) color = "#C0C0C0"; // 🥈 Plata (Plateado)
      if (posicion === 3) color = "#CD7F32"; // 🥉 Bronce (Bronceado)

      // Agregar la fila y su color a los arrays
      datosAEscribir.push([
        fila.equipo, fila.prueba, fila.nombre, fila.categoria, fila.serie, fila.andarivel,
        fila.minutos, fila.segundos, fila.centesimas, fila.tiempoTotal, posicion, puntuacion
      ]);
      formatos.push(new Array(12).fill(color)); // Se aplica el color a toda la fila

      puntuacionEquipos[fila.equipo] = (puntuacionEquipos[fila.equipo] || 0) + puntuacion;
      posicion++;
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
  }

  // Aplicar bordes gruesos al final de cada grupo de prueba/categoría
  bordes.forEach(fila => {
    hojaResultados.getRange(fila, 1, 1, 12).setBorder(false, false, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  return puntuacionEquipos;
}


