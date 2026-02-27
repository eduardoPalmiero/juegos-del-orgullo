## Torneo Todes - Script de resultados 2026

Este proyecto contiene un script de Google Apps Script para procesar resultados del Torneo Todes 2026 en Google Sheets.

- Funcion de entrada: `main()`
- Trabaja sobre la planilla activa (o por ID si se configura `SPREADSHEET_ID`)

## Estructura del spreadsheet

La planilla debe tener estas hojas:

- `Carga`
- `Resultados`
- `Puntuacion Equipos`

### Hoja `Carga`

Cada fila (desde la fila 2) representa una persona o equipo:

1. Equipo
2. Prueba
3. Nombre y Apellido
4. Serie
5. Andarivel
6. Minutos
7. Segundos
8. Centesimas
9. Categoria (manual para postas)
10. (columna 11) Metros (solo para `Americana`)

## Nombres de prueba recomendados (2026)

Usar estos nombres literales en `Prueba`:

- `1 - 50 Libre`
- `2 - 50 Espalda`
- `3 - 50 Libres - Bautismo`
- `4 - 50 Pecho`
- `5 - 50 Mariposa`
- `6 - 50 Pecho - Bautismo`
- `7 - 100 Medley`
- `8 - 4x50 Medley`
- `Americana`

## Normalizacion interna de nombres

El script normaliza el texto de `Prueba` para categorizar individuales por tiempo:

- quita prefijos numericos (`1 -`, `2 -`, etc.)
- quita sufijo `- Bautismo`
- unifica variantes como `50 Libres` -> `50 Libre`
- tolera variantes con `mts` o `metros`

Esto permite mapear a claves internas:

- `50 Libre`
- `50 Espalda`
- `50 Pecho`
- `50 Mariposa`
- `100 Medley`

Igualmente, se recomienda mantener el formato listado arriba para evitar sorpresas en carga manual.

## Reglas de categorias

### Individuales

En pruebas individuales, la categoria A-F se calcula por tiempo.

### Postas

Para `4X50 Medley` y `Americana`, la categoria se toma de la columna `Categoria` (no se calcula por tiempo).

## Reglas de puntuacion

Base FINA para puestos 1 a 8:

- `9, 7, 6, 5, 4, 3, 2, 1`

Reglas adicionales:

- `4X50 Medley` y `Americana` puntuan doble.
- Toda prueba cuyo nombre incluya la palabra `Bautismo` puntua `0`.
- Las pruebas de Bautismo igual aparecen en `Resultados` con puesto y tiempo, pero no suman al total de equipo.
- `Americana` se ordena por metros (descendente).

## Salidas

### Hoja `Resultados`

Se limpia y se vuelve a generar con:

1. Equipo
2. Prueba
3. Nombre y Apellido
4. Categoria
5. Serie
6. Andarivel
7. Minutos
8. Segundos
9. Centesimas
10. Metros
11. Tiempo Total (`MM:SS:CC`)
12. Puesto
13. Puntuacion

Adicionalmente:

- agrupacion por `Prueba + Categoria`
- orden por tiempo ascendente (o metros descendente en `Americana`)
- coloreado de podio (oro, plata, bronce)

### Hoja `Puntuacion Equipos`

Se limpia y regenera con:

- `Equipo`
- `Puntuacion Total`

Orden descendente por puntaje total.

## Uso rapido

1. Cargar datos en `Carga`.
2. Abrir Apps Script (Extensiones -> Apps Script).
3. Copiar `main-torneo-todes.js`.
4. Ejecutar `main()`.
5. Revisar `Resultados` y `Puntuacion Equipos`.
