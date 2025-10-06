function ordenarYAgruparResultados() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaCarga = spreadsheet.getSheetByName('Carga'); // Hoja de datos originales
  var hojaResultados = spreadsheet.getSheetByName('Resultados'); // Hoja de resultados
  var hojaPuntuacionEquipos = spreadsheet.getSheetByName('Puntuacion Equipos'); // Hoja de puntuaciones por equipo

  // Limpiar hojas antes de escribir datos
  hojaResultados.clear();
  hojaPuntuacionEquipos.clear();

  // Escribir los encabezados en las hojas
  hojaResultados.appendRow(['Equipo', 'Prueba', 'Nombre y Apellido', 'Categoría', 'Género', 'Serie', 'Andarivel',
    'Minutos', 'Segundos', 'Centesimas', 'Tiempo Total', 'Metros', 'Posición', 'Puntuación']);
  hojaPuntuacionEquipos.appendRow(['Equipo', 'Puntuación Total']);

  // Leer y procesar los datos
  var datos = leerDatos(hojaCarga);
  var resultadosAgrupados = agruparYOrdenarResultados(datos);
  var puntuacionEquipos = procesarResultadosYEscribir(hojaResultados, resultadosAgrupados);

  // Escribir la puntuación por equipo, ordenada por puntuación de mayor a menor
  escribirPuntuacionEquiposOrdenada(hojaPuntuacionEquipos, puntuacionEquipos);
}

// Leer los datos de la hoja de entrada
function leerDatos(hojaDatos) {
  const desdeQueFila = 2;
  const desdeQueColumna = 1;
  return hojaDatos.getRange(desdeQueFila, desdeQueColumna, hojaDatos.getLastRow() - 1, hojaDatos.getLastColumn())
    .getValues()
    .map(function (fila) {
      var minutos = fila[7] || 0;       // Columna H (Minutos)
      var segundos = fila[8] || 0;      // Columna I (Segundos)
      var centesimas = fila[9] || 0;    // Columna J (Centesimas)
      var tiempoTotal = minutos * 60 + segundos + centesimas / 100;
      var metros = fila[11] || 0;       // Columna L (Metros)

      // Si es una prueba "Americana", usar "Metros"; de lo contrario, usar "Tiempo Total"
      if (fila[1] === 'Americana' && metros > 0) {
        return {
          equipo: fila[0],
          prueba: fila[1],
          nombre: fila[2],
          categoria: fila[3],
          genero: fila[4],
          serie: fila[5],
          andarivel: fila[6],
          minutos: '',
          segundos: '',
          centesimas: '',
          tiempoTotal: '',
          metros: metros
        };
      } else if (tiempoTotal > 0) {
        return {
          equipo: fila[0],
          prueba: fila[1],
          nombre: fila[2],
          categoria: fila[3],
          genero: fila[4],
          serie: fila[5],
          andarivel: fila[6],
          minutos: minutos,
          segundos: segundos,
          centesimas: centesimas,
          tiempoTotal: tiempoTotal,
          metros: ''
        };
      }
      return null; // Ignorar filas sin datos relevantes
    })
    .filter(function (fila) {
      return fila !== null;
    });
}

// Agrupar los datos por prueba, categoría y género, y luego ordenar por tiempo o metros según el tipo de prueba
function agruparYOrdenarResultados(resultados) {
  var agrupado = {};
  resultados.forEach(function (fila) {
    var key = fila.prueba + '-' + fila.categoria + '-' + fila.genero;
    if (!agrupado[key]) {
      agrupado[key] = [];
    }
    agrupado[key].push(fila);
  });

  // Ordenar cada grupo: por tiempo total o por metros en el caso de "Americana"
  for (var key in agrupado) {
    agrupado[key].sort(function (a, b) {
      if (a.prueba === 'Americana') {
        return b.metros - a.metros; // Ordenar de mayor a menor cantidad de metros
      } else {
        return a.tiempoTotal - b.tiempoTotal; // Ordenar por tiempo total de menor a mayor
      }
    });
  }

  return agrupado;
}

// Procesar resultados: asignar posiciones y puntuaciones, y escribir en la hoja
function procesarResultadosYEscribir(hojaResultados, resultadosAgrupados) {
  var puntuacionEquipos = {};

  for (var key in resultadosAgrupados) {
    var posicion = 1; // Posición inicial
    resultadosAgrupados[key].forEach(function (fila, index) {
      var puntuacion = calcularPuntuacion(fila.prueba, posicion); // Calcular puntuación según la posición y tipo de prueba

      // Determinar si hay empate
      var empate = false;
      if (index > 0) {
        var filaAnterior = resultadosAgrupados[key][index - 1];
        if (fila.prueba === 'Americana') {
          empate = fila.metros === filaAnterior.metros;
        } else {
          empate = fila.tiempoTotal === filaAnterior.tiempoTotal;
        }
      }

      if (empate) {
        hojaResultados.appendRow([fila.equipo, fila.prueba, fila.nombre, fila.categoria, fila.genero, fila.serie,
        fila.andarivel, fila.minutos, fila.segundos, fila.centesimas, fila.tiempoTotal,
        fila.metros, posicion - 1, calcularPuntuacion(fila.prueba, posicion - 1)]);
        puntuacion = calcularPuntuacion(fila.prueba, posicion - 1);
      } else {
        hojaResultados.appendRow([fila.equipo, fila.prueba, fila.nombre, fila.categoria, fila.genero, fila.serie,
        fila.andarivel, fila.minutos, fila.segundos, fila.centesimas, fila.tiempoTotal,
        fila.metros, posicion, puntuacion]);
        posicion++;
      }

      // Acumular la puntuación por equipo
      if (puntuacionEquipos[fila.equipo]) {
        puntuacionEquipos[fila.equipo] += puntuacion;
      } else {
        puntuacionEquipos[fila.equipo] = puntuacion;
      }
    });
  }

  return puntuacionEquipos;
}

// Escribir la puntuación total por equipo en la hoja "Puntuacion Equipos", ordenada por puntuación
function escribirPuntuacionEquiposOrdenada(hojaPuntuacionEquipos, puntuacionEquipos) {
  // Convertir el objeto en un array de [equipo, puntuacion]
  var equiposArray = Object.keys(puntuacionEquipos).map(function (equipo) {
    return [equipo, puntuacionEquipos[equipo]];
  });

  // Ordenar el array por puntuación de mayor a menor
  equiposArray.sort(function (a, b) {
    return b[1] - a[1];
  });

  // Escribir el array ordenado en la hoja
  equiposArray.forEach(function (fila) {
    hojaPuntuacionEquipos.appendRow(fila);
  });
}

// Calcular la puntuación en base a la posición y tipo de prueba
function calcularPuntuacion(prueba, posicion) {
  if (prueba === '4x50 Libres') {
    // Mapeo para 4x50 Libres
    switch (posicion) {
      case 1: return 20;
      case 2: return 16;
      case 3: return 12;
      case 4: return 10;
      case 5: return 8;
      case 6: return 6;
      case 7: return 4;
      case 8: return 2;
      default: return 0;
    }
  } else if (prueba === 'Americana') {
    // Mapeo para Americana
    switch (posicion) {
      case 1: return 30;
      case 2: return 24;
      case 3: return 18;
      case 4: return 15;
      case 5: return 12;
      case 6: return 9;
      case 7: return 6;
      case 8: return 3;
      default: return 0;
    }
  } else {
    // Mapeo estándar para otras pruebas
    switch (posicion) {
      case 1: return 10;
      case 2: return 8;
      case 3: return 6;
      case 4: return 5;
      case 5: return 4;
      case 6: return 3;
      case 7: return 2;
      case 8: return 1;
      default: return 0;
    }
  }
}

