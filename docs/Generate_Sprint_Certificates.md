## Descripcion general

Este documento describe el comportamiento real implementado en el codigo para generar certificados de sprint a partir del forecast y la plantilla.

## Estructura del forecast que usa el codigo

1. Filas de equipos:
   - Se leen las filas 1, 2, 3 y 4 como equipos en este orden:
   - Bonificaciones
   - Subvenciones
   - Fondos de Reserva MRR
   - AccentureTransversal

2. Dias del mes:
   - Se leen de la fila 5 (valores numericos 1..31).

3. Identificacion de sprints:
   - Se detectan por celdas combinadas y tambien por celdas sueltas en filas 1..4.
   - Si el encabezado del sprint en la hoja del mes objetivo esta rayado (pattern no solido), ese sprint no se factura.
   - Para equipos normales se extrae el id con patron SP + numero.
   - Para Transversal se usa el id YYYY-MM del mes de la hoja.

4. Hojas consideradas:
   - Para detectar sprints facturables del mes objetivo se leen dos hojas: mes anterior y mes objetivo.
   - Para cada sprint se reconstruye la ventana completa entre fecha inicio y fecha fin uniendo segmentos de ambos meses cuando aplica.

5. Festivos:
   - Se detectan en fila 5 por relleno solido rojo o amarillo.
   - Solo se guardan los festivos dentro del rango de fechas del sprint.

6. Tecnicos y datos de equipo:
   - La unica columna valida para nombre de tecnico es la columna con cabecera Tecnico.
   - No se usan Nombre y Apellidos ni Usuario TFS para construir el listado.
   - Equipo se lee de cabecera Equipo.
   - Facturacion se lee de cabecera Facturacion si existe.
   - Categoria se lee de cabecera Perfil Facturable.
   - Si un tecnico aparece mas de una vez en el mismo equipo, se conserva la primera aparicion y se ignoran las siguientes.

7. Dias no laborables por tecnico:
   - Un dia de tecnico cuenta como no laborable si la celda del sprint esta gris (relleno gris por color o tema).
   - Si ese dia ademas es festivo detectado en fila 5, no suma horas libres.

## Reglas de facturacion implementadas

1. Un sprint es facturable si:
   - su fecha fin cae en el mes objetivo
   - y no esta rayado en la hoja del mes objetivo

2. Horas sprint:
   - Se calculan solo por los dias del rango del sprint comprendidos entre lunes y viernes (ambos incluidos).
   - El valor es comun para todos los tecnicos del sprint.
   - Horas por dia:
   - 7.5 entre el 15 de junio y el 14 de septiembre, ambos incluidos.
   - 8.5 el resto del anio.

3. Horas libres:
   - Se calculan sumando las horas diarias de cada dia no laborable del tecnico dentro del sprint, solo de lunes a viernes.
   - Horas por cada dia no laborable:
   - 7.5 entre el 15 de junio y el 14 de septiembre, ambos incluidos.
   - 8.5 el resto del anio.
   - Los dias en blanco/gris que caen en sabado o domingo no computan horas libres.
   - No se cuentan como horas libres las celdas grises que coinciden con un dia festivo.

## Escritura en plantilla implementada

Plantilla usada: plantilla de Inf_Certificacion 20251.xlsm

Para cada certificado:

1. Se copia la plantilla y se rellena la hoja Config:
   - A2: Fecha Inicio
   - B2: Fecha Fin
   - D2: identificador de sprint (en Transversal se escribe el nombre del mes)
   - G3: Product segun mapeo de equipo

2. Festivos:
   - Se escriben en la tabla TB_Festivos (columnas A y B).
   - Si hacen falta mas filas, la tabla se expande.

3. Equipo:
   - Se escribe en la tabla TB_Equipo (columnas F..J): Tecnico, Facturacion, Categoria, Horas Sprint, Horas libres.
   - Categoria se normaliza para que sea exactamente una opcion valida del desplegable (tabla TB_Perfiles).
   - Si una categoria no puede mapearse a una opcion valida, el proceso falla con error.
   - Si un tecnico viene duplicado en el payload, se escribe una sola vez.

4. ZIP final:
   - Se genera un zip con todos los certificados del mes.

5. Carpeta de salida:
   - La salida se genera siempre en `cert_automation/certificaciones/<YYYY-MM>`.
   - No se usa ruta de salida configurable por parametro.

## Nomenclatura de ficheros implementada

1. Equipo Transversal:
   - <sprint_id> Inf_Certificacion <anio> AccentureTranservsal.xlsm
   - Ejemplo: 2026-06 Inf_Certificacion 2026 AccentureTranservsal.xlsm

2. Resto de equipos:
   - Inf_Certificacion <anio> <equipo> <sprint_id>.xlsm
   - Ejemplo: Inf_Certificacion 2026 Subvenciones SP216.xlsm

## Restricciones operativas

1. No se modifica el fichero forecast de entrada.
2. El proceso requiere que existan las hojas FY esperadas por anio y mes.
3. La logica de tecnico exige cabecera Tecnico en la hoja usada para construir workloads.
4. La carpeta `certificaciones/` se considera artefacto generado y no debe versionarse.

