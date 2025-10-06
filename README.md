# Juegos del Orgullo - Script de Resultados de Natación

Este proyecto contiene un script de Google Apps Script para gestionar y procesar los resultados de un torneo de natación en Google Sheets.

## ¿Qué hace el script?
- **Lee los datos** desde la hoja "Carga" (participantes y resultados).
- **Agrupa y ordena** los resultados por Prueba, Categoría, Género y Tipo de Prueba.
- **Calcula posiciones y puntuaciones** solo para pruebas competitivas.
- **Escribe los resultados** en la hoja "Resultados" de forma eficiente (en bloque).
- **Calcula la puntuación total por equipo** y la muestra en la hoja "Puntuacion Equipos".

## Estructura de la hoja "Carga"
Las columnas deben estar en el siguiente orden:

1. Equipo
2. Prueba
3. Tipo de Prueba (Competitiva o Bautismo)
4. Nombre y Apellido
5. Categoría
6. Género
7. Serie
8. Andarivel
9. Minutos
10. Segundos
11. Centesimas
12. Tiempo Total
13. Metros

## Estructura de la hoja "Resultados"
Las columnas generadas son:

1. Equipo
2. Prueba
3. Tipo de Prueba
4. Nombre y Apellido
5. Categoría
6. Género
7. Serie
8. Andarivel
9. Minutos
10. Segundos
11. Centesimas
12. Tiempo Total
13. Metros
14. Posición
15. Puntuación

## Estructura de la hoja "Puntuacion Equipos"
- Equipo
- Puntuación Total

## ¿Cómo usar?
1. Copia el contenido de `main.js` en el editor de Google Apps Script vinculado a tu Google Sheet.
2. Asegúrate de que las hojas tengan los nombres correctos: "Carga", "Resultados" y "Puntuacion Equipos".
3. Ejecuta la función principal `ordenarYAgruparResultados`.

## Notas
- Solo las pruebas marcadas como "Competitiva" suman puntos y posición.
- Las pruebas de tipo "Bautismo" se muestran pero no suman puntos ni posición.
- El script está optimizado para escribir los resultados en bloque, mejorando el rendimiento.

---

Vibe coded by Eduardo Palmiero.
