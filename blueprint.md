# Blueprint: Automatización Académica MIA Capstone (Defensa Oral)

Este documento contiene las especificaciones lógicas, planos de datos y diseños de algoritmos acordados para la automatización y consolidación de calificaciones de Defensa Oral del proyecto Capstone.

---

## 1. Planos de Datos y Esquemas

### A. Estructura de Entrada (Microsoft Forms Crudo)
El reporte nativo de Excel exportado de MS Forms se compone de las siguientes columnas clave:
- **Id** (Columna 1, Entero): Identificador único del envío del formulario.
- **Seleccione su nombre (Evaluador)** (Columna 12, Texto): Nombre del jurado que evalúa.
- **Seleccione el nombre del Estudiante** (Columna 15, Texto): Nombre completo del alumno evaluado.
- **Criterios de Evaluación (Rúbrica)**:
  1. *Apertura: problema y objetivo* (Columna 18, Texto, Max: 20 pts)
  2. *Metodología (nivel adecuado)* (Columna 21, Texto, Max: 20 pts)
  3. *Resultados y evidencia* (Columna 24, Texto, Max: 20 pts)
  4. *Coherencia y manejo del tiempo* (Columna 27, Texto, Max: 10 pts)
  5. *Diapositivas como apoyo* (Columna 30, Texto, Max: 10 pts)
  6. *Respuestas a preguntas* (Columna 33, Texto, Max: 10 pts)
  7. *Cierre: aporte a próximos pasos* (Columna 36, Texto, Max: 10 pts)

### B. Estructura de Salida (Hoja `Calculo` / `Seguimiento`)
El consolidado final exportado tendrá la siguiente estructura plana:
- **Seleccione el nombre del Estudiante** (Texto): Nombre del alumno (clave única).
- **Cuenta de Seleccione su nombre (Evaluador)** (Entero): Cantidad de evaluaciones recibidas (generalmente 3, pero admite 2 o 1).
- **Promedio de Points - [Criterio]** (Decimal): Promedio aritmético para cada uno de los 7 criterios.
- **Rubrica - [Criterio]** (Texto): Etiqueta cualitativa traducida del promedio de ese criterio.
- **Nota ponderada** (Decimal): Suma total de los promedios de los 7 criterios (máximo 100).
- **Nota Rubrica Promedio** (Texto): Etiqueta cualitativa final sobre 100.

---

## 2. Reglas de Negocio y Lógica Algorítmica

### A. Extracción de Puntos (Regex)
Para extraer el puntaje de opciones como `'Excelente (20 puntos)'` se aplica el patrón:
$$\text{Regex Pattern: } \texttt{\textbackslash((\textbackslash d+)\textbackslash s*puntos?\textbackslash)}$$
- Retorna el primer grupo de captura convertido a entero.
- Si es numérico directo se conserva su valor. Fallbacks seguros a `0`.

### B. Normalización y Match de Nombres
Para cruces y agrupaciones consistentes se aplica la siguiente función de normalización:
1. Eliminar acentos y diacríticos (ej. `á` $\rightarrow$ `a`, `ñ` $\rightarrow$ `n`).
2. Pasar a mayúsculas sostenidas (`.upper()`).
3. Remover espacios en blanco adicionales y colapsar espacios internos a un único caracter de espacio.

### C. Traducción Cualitativa (Lógica de Rangos)
Debido a la precisión de punto flotante en cálculos y diferencias de IEEE 754 con Excel (que redondea a 15 dígitos significativos para comparaciones), la nota y los promedios se **redondean a 10 decimales** antes de clasificar:

- **Criterios de 20 pts**:
  - $\ge 18.0 \rightarrow$ Excelente
  - $\ge 14.0 \rightarrow$ Muy bueno
  - $\ge 10.0 \rightarrow$ Bueno
  - $\ge 6.0 \rightarrow$ Regular
  - $< 6.0 \rightarrow$ Insuficiente
- **Criterios de 10 pts**:
  - $\ge 9.0 \rightarrow$ Excelente
  - $\ge 7.0 \rightarrow$ Muy bueno
  - $\ge 5.0 \rightarrow$ Bueno
  - $\ge 3.0 \rightarrow$ Regular
  - $< 3.0 \rightarrow$ Insuficiente
- **Nota Final (100 pts)**:
  - $\ge 90.0 \rightarrow$ Excelente
  - $\ge 80.0 \rightarrow$ Muy Bueno *(Casing idéntico al histórico)*
  - $\ge 75.0 \rightarrow$ Bueno
  - $\ge 65.0 \rightarrow$ Regular
  - $< 65.0 \rightarrow$ Insuficiente

### D. Compliance de Tribunales
Cruce dinámico con la planificación del comités (`CALIFICACION-DOCENTE`):
- Se divide la cadena del grupo de estudiantes (delimitador: `,`) y se normaliza cada integrante.
- Se compara al docente planificado con el que envió la respuesta.
- Si existe una evaluación registrada del jurado asignado para ese estudiante, su estado es `Completado`, de lo contrario es `Pendiente`.
