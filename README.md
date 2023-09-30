# ExcelMonteCarlo
Macro en VBA para implementación y ejecución de modelos Monte Carlo en Excel

La macro permite el modelamiento de simulaciones tipo Montecarlo (Monte Carlo Simulations) desde Excel.
Generación automática de pseudo-aleatorios con posibilidad de selección de semilla (seed) para replicación de experimento.
Selección entre diferentes tipos de distribución de probabilidad: uniforme, triangular y normal.
Generación de números aleatorios de acuerdo a la selección de distribución de probabilidad.
Simulación de tipo 2D, permitiendo generar X cantidad de simulaciones Montecarlo generando nuevos pseudo-aleatorios para cada simulación externa.
Actualización en tiempo real del estado de la simulación para indicar porcentaje de avance en barra de estado de Excel.
Cálculo automático de estadísticas relevantes para cada simulación (mínimo, máximo, promedio, desviación estándar).
Funcionalidad para ajuste de distintos parámetros de simulación como lo son:
- Semilla para generación de pseudo-aleatorios.
- Distribución de probabilidad.
- Número de iteraciones por simulación.
- Número de simulaciones externas para simulaciones 2D.
- Parámetros propios de cada distribución de probabilidad (min, max, desv. estandar, etc.).
