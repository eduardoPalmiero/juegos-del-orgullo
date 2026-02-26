## Torneo Todes – Script de resultados 2026

Este proyecto contiene un script de **Google Apps Script** para procesar los resultados del Torneo Todes 2026 en una planilla de Google Sheets.  
La función de entrada es `main()`, que trabaja sobre la **planilla activa**.

---

## Estructura del spreadsheet

La planilla debe tener, como mínimo, estas hojas:

- `Carga`: hoja donde se cargan todos los resultados crudos.
- `Resultados`: hoja donde el script escribe los resultados ordenados.
- `Puntuacion Equipos`: hoja donde el script escribe la puntuación total por equipo.

### Hoja `Carga`

Cada fila (a partir de la fila 2) representa un resultado de una persona o equipo.  
Las columnas esperadas son:

1. **Equipo**: nombre del equipo (texto).
2. **Prueba**: nombre de la prueba. Ejemplos:
   - `"1 - 50 mts Libre"`
   - `"2 - 50 mts Espalda"`
   - `"3 - 50 mts Pecho"`
   - `"4 - 50 mts Mariposa"`
   - `"5 - 100 mts Medley"`
   - `"4x50 Medley"` (posta combinada)
   - `"Americana"` (posta americana)
3. **Nombre y Apellido**: de la persona (en postas puede ser el nombre del equipo / posta).
4. **Serie**: número de serie (entero).
5. **Andarivel**: número de andarivel (entero).
6. **Minutos**: parte en minutos del tiempo registrado (entero, puede ser 0).
7. **Segundos**: parte en segundos del tiempo registrado (entero, puede ser 0).
8. **Centesimas**: parte en centésimas del tiempo registrado (entero, puede ser 0).
9. **Categoría**:
   - Para **pruebas individuales** se puede dejar vacía: el script calculará la categoría A–F según los tiempos.
   - Para **postas** (por equipos o americana) se usa para las categorías por edades (A–D) y el script simplemente la copia.
10. **Metros** (índice 10, es decir, 11ª columna):
    - Solo se usa en la prueba `"Americana"`.
    - Representa los metros totales recorridos por el equipo.
    - Para el resto de las pruebas se deja vacío.

> Importante: la fila 1 debe contener los encabezados de estas columnas, pero el script no depende de los textos exactos de los títulos, sino de la posición de las columnas.
>
> **Muy importante sobre los nombres de las pruebas:** para que el script pueda reconocer la prueba y aplicar correctamente las categorías por tiempos, los nombres de las pruebas individuales deben escribirse **exactamente** así en la columna `Prueba` (incluyendo espacios y mayúsculas/minúsculas):
> - `"1 - 50 mts Libre"`
> - `"2 - 50 mts Espalda"`
> - `"3 - 50 mts Pecho"`
> - `"4 - 50 mts Mariposa"`
> - `"5 - 100 mts Medley"`  
> Si se usan otros textos (por ejemplo, `50 Pecho` o `3-50mts Pecho`), el script **no podrá calcular la categoría A–F** porque no encontrará coincidencia en su tabla interna.  
> Para las postas se recomienda usar exactamente `"4x50 Medley"` y `"Americana"` para mantener homogeneidad con el reglamento y el código.

### Hoja `Resultados`

Es generada por el script. Al ejecutar `main()`, se limpia y se vuelve a completar con estas columnas:

1. Equipo  
2. Prueba  
3. Nombre y Apellido  
4. Categoría  
5. Serie  
6. Andarivel  
7. Minutos  
8. Segundos  
9. Centesimas  
10. Metros  
11. Tiempo Total (formato texto `MM:SS:CC`)  
12. Puesto  
13. Puntuación  

Además:

- Las filas se agrupan por **Prueba + Categoría**.
- Dentro de cada grupo se ordena:
  - Por **tiempo total ascendente** en pruebas con tiempo.
  - Por **metros descendentes** en `"Americana"`.
- Se colorean las filas según el puesto:
  - 1.º: dorado.
  - 2.º: plateado.
  - 3.º: bronce.
- Se dibuja una línea de borde al final de cada grupo de Prueba/Categoría.

### Hoja `Puntuacion Equipos`

También se limpia y se completa en cada ejecución de `main()`.  
Tiene la forma:

- `Equipo`  
- `Puntuación Total`  

Los equipos se ordenan de mayor a menor puntuación, y los tres primeros puestos se destacan con colores (oro, plata y bronce).

---

## Reglas de categorías

### Pruebas individuales – Categorías por tiempos

Para las pruebas individuales, el script **calcula siempre** la categoría A–F en base al tiempo total en segundos, usando los rangos definidos en el reglamento:

- `"3 - 50 mts Pecho"`  
  - A: hasta 39.99 s  
  - B: 40.00 a 44.99 s  
  - C: 45.00 a 49.99 s  
  - D: 50.00 a 54.99 s  
  - E: 55.00 a 59.99 s  
  - F: 60.00 s o más  
- `"2 - 50 mts Espalda"`  
  - A: hasta 37.99 s  
  - B: 38.00 a 42.99 s  
  - C: 43.00 a 47.99 s  
  - D: 48.00 a 52.99 s  
  - E: 53.00 a 57.99 s  
  - F: 58.00 s o más  
- `"1 - 50 mts Libre"`  
  - A: hasta 29.99 s  
  - B: 30.00 a 34.99 s  
  - C: 35.00 a 39.99 s  
  - D: 40.00 a 44.99 s  
  - E: 45.00 a 49.99 s  
  - F: 50.00 s o más  
- `"4 - 50 mts Mariposa"`  
  - A: hasta 34.99 s  
  - B: 35.00 a 39.99 s  
  - C: 40.00 a 44.99 s  
  - D: 45.00 a 49.99 s  
  - E: 50.00 a 54.99 s  
  - F: 55.00 s o más  
- `"5 - 100 mts Medley"`  
  - A: hasta 79.99 s  
  - B: 80.00 a 87.99 s  
  - C: 88.00 a 95.99 s  
  - D: 96.00 a 103.99 s  
  - E: 104.00 a 111.99 s  
  - F: 112.00 s o más  

### Postas – Categorías por edades

Para las postas (`4x50 Medley`, `"Americana"`):

- La **categoría no se calcula por tiempo**, sino por suma de edades del equipo según el reglamento (A, B, C, D).
- Esa categoría se debe cargar manualmente en la columna **Categoría** de la hoja `Carga`.
- El script copia esa categoría a `Resultados` y la usa solo para agrupar y ordenar.

---

## Reglas de puntuación

La puntuación por puesto sigue el sistema FINA para 8 nadadores/equipos:

- 1.º: 9 puntos  
- 2.º: 7 puntos  
- 3.º: 6 puntos  
- 4.º: 5 puntos  
- 5.º: 4 puntos  
- 6.º: 3 puntos  
- 7.º: 2 puntos  
- 8.º: 1 punto  

Las **postas por equipos y americanas puntúan doble**:

- Cualquier prueba detectada como posta (`4x50 Medley`, `"Americana"`, u otras con formato `4x…`) suma el **doble de puntos** que los valores base.  
- La prueba `"Americana"` se ordena por metros y también puntúa doble.

La puntuación se acumula por equipo y el total se muestra en `Puntuacion Equipos`.

---

## Cómo usar el script

1. Crear o abrir el Google Sheet del torneo con las hojas `Carga`, `Resultados` y `Puntuacion Equipos`.
2. Abrir el editor de **Google Apps Script** desde el menú Extensiones → Apps Script.
3. Copiar el contenido de `main-torneo-todes.js` en el editor.
4. Guardar y ejecutar la función `main()` (la primera vez, autorizar permisos).
5. Verificar:
   - Que `Resultados` esté ordenada por prueba/categoría, con puestos y puntajes correctos.
   - Que `Puntuacion Equipos` muestre la suma total de puntos por equipo.
