## Descripción general de la pestaña Template_Mes
Explicación de la lógica de la pestaña Template_Mes, la cual se usa como base para la creación de nuevos periodos de facturación por mes y con la que se controlan las horas por trabajador, tanto de coste como horas facturables.

La pestaña Template_Mes se usa como plantilla de base para adaptar los periodos / meses que vamos facturando a nuestro cliente. El servicio se divide en 4 equipos: Bonificaciones, Subvenciones, Fondos de Reserva y Transversal. Cada uno de ellos lleva su propia planificación y ejecución de sprints, y los ciclos de facturación no se producen a nivel de mes, sino a nivel de finalización de sprint: Se podrá facturar mensualmente, todos los sprints que finalicen en el mes, aquellos que finalicen en el mes siguiente no se podrán facturar y se harán en el siguiente periodo. Es por eso que necesitamos controlar las horas de sprints finalizados en meses anteriores, horas que finalizan en sprints durante el mes, y horas de sprints que no se pueden facturar durante el mes de la pestaña en cuestión.

## Reglas para calcular los sprints.
Cada equipo tiene su criterio para organizar los sprints, teniendo en cuenta que solo consideramos días laborables de lunes a viernes:
- Bonificaciones: Duración de 10 días laborables empezando en martes y finalizando en lunes.
- Subvenciones: Duración de 16 días laborables.
- Fondos de Reserva: Divide el mes en dos sprints: del 1 al 15 del mes (ambos incluídos) y del 16 al último día del mes (ambos incluídos) por lo que este equipo siempre podrá facturar el 100% de las horas en el periodo que se ejecute.
- Transversales: un único sprint que abarca del 1 al último dia del mes (ambos incluídos) por lo que este equipo siempre podrá facturar el 100% de las horas en el periodo que se ejecute.

## Principales tablas para entender la estructura del documento:
### T_NOMBRE_EMPLEADO_COSTE Y T_NOMBRE_EMPLEADO_REVENUE
Muestran los nombres de los empleados por fila y están divididos en dos tablas para mostrar por día su información tanto de coste como de tarifa (revenue).
### T_TIMESHEET Y T_REVENUES_TIMESHEET
Reflejan las horas que supone cada empleado por día de coste y que se pueden facturar respectivamente
### T_SPRINTS
Representan los sprints que se realizan durante el mes natural. La fila 1 es para el equipo de bonificaciones, 2 para Subvenciones, 3 para Fondos de Reserva y 4 para Transversal. Se agrupan las celdas de cada fila en esta tabla para representar la duración del sprint. Gracias a la tabla T_CALENDARIO se puede ver cuándo inicia y finaliza el sprint y con la tabla T_TIMSHEET Y T_REVENUES_TIMESHEET sabremos cuántas horas de coste y de revenues tenemos para cada sprint.
### T_GAP_PERIODO_ANTERIOR
Muestra el acumulado de horas por persona del periodo anterior que no se pudieron facturar. Las fórmulas de cada celda son necesarias actualizarlas para cada periodo ya que dependen de la duración de los sprints del periodo anterior.
### T_REVENUES_MES_ACTUAL
Muestra el acumulado por persona de horas del periodo actual que se podrán facturar. Las fórmulas de cada celda son necesarias actualizarlas para cada periodo ya que dependen de la duración de los sprints del periodo actual.
### T_REVENUES_PERIODO_ACTUAL_NO_FACT
Muestra el acumulado por persona de horas del periodo actual que no se podrán facturar y que quedarán pendientes para el siguiente periodo. Las fórmulas de cada celda son necesarias actualizarlas para cada periodo ya que dependen de la duración de los sprints del periodo actual y que no se van a poder facturar.
