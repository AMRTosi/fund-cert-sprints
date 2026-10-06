## Descripcion general

Este documento define el requerimiento funcional del caso de uso 4: validacion de timesheets de Beeline para aprobar o rechazar gastos de personal subcontratado.

El proceso cruza un JSON de entrada (fuente externa) con un fichero Excel de control (fuente veraz), actualiza el estado visual en Excel y genera un JSON de resultado con el veredicto de cada elemento.

## Entradas

1. Fichero JSON de entrada con una coleccion de elementos a validar.
2. Fichero Excel de control con la hoja `Timecards Aprobadas`.

Referencia de workbook para pruebas funcionales/manuales:
- `./inputs/Delivery_Management_GestionIniciativas.xlsx`

## Contrato del JSON de entrada

Cada elemento del array de entrada debe contener todos estos campos, obligatorios y de tipo string:

1. `beelineRevieweeName`
2. `timePeriod`
3. `totalEstimatedAmount`

Reglas de validacion de contrato:

1. Si falta un campo obligatorio o viene vacio, se rechaza solo ese elemento y el proceso continua con el resto.
2. La comparacion de strings es exacta (sin normalizacion de mayusculas, tildes, espacios o puntuacion).

## Estructura esperada en Excel

Hoja objetivo: `Timecards Aprobadas`.

1. Columna A: nombres de reviewee (`beelineRevieweeName`).
2. Fila 3: periodos (`timePeriod`).
3. Interseccion reviewee + periodo: celda unica con el `totalEstimatedAmount` de control.

## Regla de estado (duplicado) por estilo visual

La deteccion de duplicado no depende de si la celda esta vacia o no, sino del formato de fuente:

1. Estado pendiente (no validado): fuente negra (estilo por defecto).
2. Estado ya validado: fuente verde y negrita.

Por tanto:

1. Si la celda de interseccion ya esta en verde + negrita, se considera duplicado.
2. Si la celda de interseccion esta en negro, se considera pendiente de validar.

## Flujo de validacion por elemento

Para cada objeto del JSON de entrada:

1. Validar contrato de entrada.
2. Buscar `beelineRevieweeName` en columna A de `Timecards Aprobadas` con comparacion exacta.
3. Si no existe reviewee:
	- `validationStatus`: `NotFound`
	- `message`: `Reviewee not Found.`
	- `beelineAction`: `NA`
4. Si existe reviewee, buscar `timePeriod` en fila 3.
5. La comparacion de periodos acepta equivalencia de formato de fecha (`10/5/2026` y `10/05/2026` se consideran el mismo dia).
6. Si la cabecera del periodo en Excel es formula, se usa su valor calculado cuando este disponible; en caso contrario, se deriva el rango semanal desde la columna anterior para intentar resolver el match.
7. Resolver celda de interseccion (reviewee, periodo).
8. Comprobar estilo de fuente en la celda:
	- Si esta en verde + negrita: duplicado.
	- Resultado:
	  - `validationStatus`: `Failed`
	  - `message`: `Timecard enviada por duplicado.`
	  - `beelineAction`: `Reject`
	- No modificar valor ni estilo.
9. Si esta en estado pendiente (fuente negra), leer el valor de la celda (importe veraz).
10. Si el valor de Excel no esta disponible:
	- `validationStatus`: `NotFound`
	- `message`: `totalEstimatedAmount no disponible en Excel. Revisar cálculo en el fichero.`
	- `beelineAction`: `NA`
11. Si hay valor, comparar `totalEstimatedAmount` como importe numerico contra la celda de Excel.
12. La comparacion debe tratar formatos numericos comunes (punto o coma decimal y separadores de miles) como equivalentes si representan el mismo valor.
13. Si coinciden:
	- `validationStatus`: `OK`
	- `message`: `` (string vacio)
	- `beelineAction`: `Approve`
	- Marcar la celda como validada cambiando estilo de fuente a verde + negrita.
14. Si no coinciden:
	- `validationStatus`: `Failed`
	- `message`: `Timecard incorrecta. Revisar y Rechzarla`
	- `beelineAction`: `Reject`
	- No modificar estilo ni valor.

## Contrato del JSON de salida

Salida: array plano de objetos (sin envoltorio), en el mismo orden del JSON de entrada.

Estructura por elemento:

```json
{
  "revieweeBeelineName": "de Manuel Ruiz, Alberto",
  "timePeriod": "9/28/2026–10/4/2026",
  "validationStatus": "OK",
  "message": "",
  "beelineAction": "Approve"
}
```

Catalogo cerrado:

1. `validationStatus`: `OK`, `Failed`, `NotFound`
2. `beelineAction`: `Approve`, `Reject`, `NA`

## Reglas de persistencia y efectos

1. El proceso lee el JSON de entrada.
2. El proceso lee el Excel de control.
3. El proceso escribe en `Timecards Aprobadas` unicamente para marcar como validadas las celdas aprobadas (estilo verde + negrita).
4. El proceso genera un JSON de salida de validacion.
5. Si un elemento falla, no bloquea el procesamiento del resto.

## Mensajes funcionales definidos

1. `Reviewee not Found.`
2. `Timecard enviada por duplicado.`
3. `Timecard incorrecta. Revisar y Rechzarla`
4. `totalEstimatedAmount no disponible en Excel. Revisar cálculo en el fichero.`

## Consideraciones de trazabilidad

1. El valor de Excel en la interseccion se considera la fuente veraz para el importe.
2. La marca de validacion (verde + negrita) es la unica senial de estado para detectar duplicados.
3. El valor de la celda no se sobrescribe por el proceso; solo puede cambiar el estilo en casos aprobados.
