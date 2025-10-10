function ordenarYAgruparResultados() {
  var tiempoInicio = new Date();
  console.log("🚀 Iniciando proceso...");
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaCarga = spreadsheet.getSheetByName("Carga"); // Hoja de datos originales
  var hojaResultados = spreadsheet.getSheetByName("Resultados"); // Hoja de resultados
  var hojaPuntuacionEquipos = spreadsheet.getSheetByName("Puntuacion Equipos"); // Hoja de puntuaciones por equipo

  // Limpiar hojas antes de escribir datos
  var tiempoLimpieza = new Date();
  hojaResultados.clear();
  hojaPuntuacionEquipos.clear();
  console.log("⏱️ Limpieza de hojas: " + (new Date() - tiempoLimpieza) + "ms");

  // Leer y procesar los datos
  var tiempoLectura = new Date();
  var datos = leerDatos(hojaCarga);
  console.log("⏱️ Lectura de datos (" + datos.length + " filas): " + (new Date() - tiempoLectura) + "ms");
  
  var tiempoAgrupacion = new Date();
  var resultadosAgrupados = agruparYOrdenarResultados(datos);
  console.log("⏱️ Agrupación y ordenamiento: " + (new Date() - tiempoAgrupacion) + "ms");
  
  var tiempoProcesamiento = new Date();
  var puntuacionEquipos = procesarResultadosYEscribir(
    hojaResultados,
    resultadosAgrupados
  );
  console.log("⏱️ Procesamiento y escritura de resultados: " + (new Date() - tiempoProcesamiento) + "ms");

  // Escribir la puntuación por equipo, ordenada por puntuación de mayor a menor
  var tiempoEquipos = new Date();
  escribirPuntuacionEquiposOrdenada(hojaPuntuacionEquipos, puntuacionEquipos);
  console.log("⏱️ Escritura de puntuación por equipos: " + (new Date() - tiempoEquipos) + "ms");
  
  var tiempoTotal = new Date() - tiempoInicio;
  console.log("✅ Proceso completado en: " + tiempoTotal + "ms (" + (tiempoTotal/1000).toFixed(2) + " segundos)");
}

// Leer los datos de la hoja de entrada
function leerDatos(hojaDatos) {
  // Optimización: usar getDataRange() en lugar de calcular rangos
  var todosLosDatos = hojaDatos.getDataRange().getValues();
  
  // Saltarse la primera fila (encabezados) y procesar el resto
  return todosLosDatos
    .slice(1) // Elimina la primera fila (encabezados)
    .map(function (fila) {
      var tipoPrueba = fila[2]; // Columna C (Tipo de Prueba)
      var minutos = fila[8] || 0; // Columna I (Minutos)
      var segundos = fila[9] || 0; // Columna J (Segundos)
      var centesimas = fila[10] || 0; // Columna K (Centesimas)
      var tiempoTotal = minutos * 60 + segundos + centesimas / 100;
      var metros = fila[12] || 0; // Columna M (Metros)

      if (fila[1] === "Americana" && metros > 0) {
        return {
          equipo: fila[0],
          prueba: fila[1],
          tipoPrueba: tipoPrueba,
          nombre: fila[3],
          categoria: fila[4],
          genero: fila[5],
          serie: fila[6],
          andarivel: fila[7],
          minutos: "",
          segundos: "",
          centesimas: "",
          tiempoTotal: "",
          metros: metros,
        };
      } else if (tiempoTotal > 0) {
        return {
          equipo: fila[0],
          prueba: fila[1],
          tipoPrueba: tipoPrueba,
          nombre: fila[3],
          categoria: fila[4],
          genero: fila[5],
          serie: fila[6],
          andarivel: fila[7],
          minutos: minutos,
          segundos: segundos,
          centesimas: centesimas,
          tiempoTotal: tiempoTotal,
          metros: "",
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
    var key =
      fila.prueba +
      "-" +
      fila.categoria +
      "-" +
      fila.genero +
      "-" +
      fila.tipoPrueba;
    if (!agrupado[key]) {
      agrupado[key] = [];
    }
    agrupado[key].push(fila);
  });

  // Ordenar cada grupo: por tiempo total o por metros en el caso de "Americana"
  for (var key in agrupado) {
    agrupado[key].sort(function (a, b) {
      if (a.prueba === "Americana") {
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
  var tiempoInicioProc = new Date();
  var puntuacionEquipos = {};
  var filasResultados = [];
  var coloresFondo = [];

  // Agregar encabezados como primera fila
  var encabezados = [
    "Equipo",
    "Prueba",
    "Tipo de Prueba",
    "Nombre y Apellido",
    "Categoría",
    "Género",
    "Serie",
    "Andarivel",
    "Minutos",
    "Segundos",
    "Centesimas",
    "Tiempo Total",
    "Metros",
    "Posición",
    "Puntuación",
  ];
  filasResultados.push(encabezados);
  
  // Color blanco para los encabezados
  var colorEncabezados = new Array(15).fill("#FFFFFF");
  coloresFondo.push(colorEncabezados);

  var tiempoConstruccion = new Date();
  for (var key in resultadosAgrupados) {
    var posicion = 1; // Posición inicial para cada grupo
    resultadosAgrupados[key].forEach(function (fila, index) {
      var empate = false;
      if (index > 0) {
        var filaAnterior = resultadosAgrupados[key][index - 1];
        if (fila.prueba === "Americana") {
          empate = fila.metros === filaAnterior.metros;
        } else {
          empate = fila.tiempoTotal === filaAnterior.tiempoTotal;
        }
      }

      var posicionActual = empate ? posicion - 1 : posicion;
      var puntuacion = 0;

      // Solo calcular puntuación si es competitiva
      if (fila.tipoPrueba === "Competitiva") {
        puntuacion = calcularPuntuacion(fila.prueba, posicionActual);
        if (puntuacionEquipos[fila.equipo]) {
          puntuacionEquipos[fila.equipo] += puntuacion;
        } else {
          puntuacionEquipos[fila.equipo] = puntuacion;
        }
      }

      filasResultados.push([
        fila.equipo,
        fila.prueba,
        fila.tipoPrueba,
        fila.nombre,
        fila.categoria,
        fila.genero,
        fila.serie,
        fila.andarivel,
        fila.minutos,
        fila.segundos,
        fila.centesimas,
        formatearTiempo(fila.tiempoTotal), // Aplicar formato de tiempo
        fila.metros,
        posicionActual,
        fila.tipoPrueba === "Competitiva" ? puntuacion : "",
      ]);

      // Definir color de fondo según la posición
      var color = "#FFFFFF"; // Blanco por defecto
      if (posicionActual === 1) color = "#FFD700"; // 🥇 Oro (Dorado)
      if (posicionActual === 2) color = "#C0C0C0"; // 🥈 Plata (Plateado)
      if (posicionActual === 3) color = "#CD7F32"; // 🥉 Bronce (Bronceado)
      
      // Crear array de colores para toda la fila (mismo color para todas las columnas)
      var colorFila = new Array(15).fill(color); // 15 columnas
      coloresFondo.push(colorFila);

      if (!empate) {
        posicion++;
      }
    });
  }
  console.log("  📊 Construcción de arrays (" + (filasResultados.length - 1) + " filas + encabezados): " + (new Date() - tiempoConstruccion) + "ms");

  // Escribir todas las filas (incluyendo encabezados) y aplicar colores en una sola operación
  var tiempoEscritura = new Date();
  if (filasResultados.length > 0) {
    var rango = hojaResultados.getRange(1, 1, filasResultados.length, filasResultados[0].length);
    var tiempoSetValues = new Date();
    rango.setValues(filasResultados);
    console.log("    ✏️ setValues (con encabezados): " + (new Date() - tiempoSetValues) + "ms");
    
    var tiempoSetColors = new Date();
    rango.setBackgrounds(coloresFondo);
    console.log("    🎨 setBackgrounds: " + (new Date() - tiempoSetColors) + "ms");
  }
  console.log("  📝 Escritura total en hoja: " + (new Date() - tiempoEscritura) + "ms");
  
  console.log("📊 Tiempo total procesarResultadosYEscribir: " + (new Date() - tiempoInicioProc) + "ms");
  return puntuacionEquipos;
}

// Escribir la puntuación total por equipo en la hoja "Puntuacion Equipos", ordenada por puntuación (uso de setValues)
function escribirPuntuacionEquiposOrdenada(
  hojaPuntuacionEquipos,
  puntuacionEquipos
) {
  // Convertir objeto en array [equipo, puntuacion] y ordenar desc
  var equiposArray = Object.keys(puntuacionEquipos)
    .map(function (equipo) {
      return [equipo, puntuacionEquipos[equipo]];
    })
    .sort(function (a, b) {
      return b[1] - a[1];
    });

  // Agregar encabezados como primera fila
  var datosCompletos = [["Equipo", "Puntuación Total"]];
  datosCompletos = datosCompletos.concat(equiposArray);

  if (datosCompletos.length > 1) {
    // Escribimos todo de una sola vez (encabezados + datos)
    hojaPuntuacionEquipos
      .getRange(1, 1, datosCompletos.length, 2)
      .setValues(datosCompletos);
  }
}

// Calcular la puntuación en base a la posición y tipo de prueba
function calcularPuntuacion(prueba, posicion) {
  if (prueba === "4x50 Libres") {
    // Mapeo para 4x50 Libres
    switch (posicion) {
      case 1:
        return 20;
      case 2:
        return 16;
      case 3:
        return 12;
      case 4:
        return 10;
      case 5:
        return 8;
      case 6:
        return 6;
      case 7:
        return 4;
      case 8:
        return 2;
      default:
        return 0;
    }
  } else if (prueba === "Americana") {
    // Mapeo para Americana
    switch (posicion) {
      case 1:
        return 30;
      case 2:
        return 24;
      case 3:
        return 18;
      case 4:
        return 15;
      case 5:
        return 12;
      case 6:
        return 9;
      case 7:
        return 6;
      case 8:
        return 3;
      default:
        return 0;
    }
  } else {
    // Mapeo estándar para otras pruebas
    switch (posicion) {
      case 1:
        return 10;
      case 2:
        return 8;
      case 3:
        return 6;
      case 4:
        return 5;
      case 5:
        return 4;
      case 6:
        return 3;
      case 7:
        return 2;
      case 8:
        return 1;
      default:
        return 0;
    }
  }
}

// Convertir tiempo en segundos decimales a formato MM:SS:CC
function formatearTiempo(tiempoEnSegundos) {
  if (!tiempoEnSegundos || tiempoEnSegundos === "" || tiempoEnSegundos === 0) {
    return "";
  }
  
  var minutos = Math.floor(tiempoEnSegundos / 60);
  var segundos = Math.floor(tiempoEnSegundos % 60);
  var centesimas = Math.round((tiempoEnSegundos % 1) * 100);
  
  // Formatear con ceros a la izquierda manualmente
  var minutosStr = minutos.toString();
  var segundosStr = segundos < 10 ? "0" + segundos : segundos.toString();
  var centesimasStr = centesimas < 10 ? "0" + centesimas : centesimas.toString();
  
  var resultado = minutosStr + ":" + segundosStr + ":" + centesimasStr;
    
  return resultado;
}
