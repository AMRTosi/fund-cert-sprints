# Caso de uso 5: Validacion de quincena myTE

## Objetivo y alcance

Comparar dos ficheros JSON: horas previstas extraidas de Excel y datos extraidos
de myTE. Se preservan las reglas del servicio fundae-gi-myte-validator.

No extrae datos de myTE, no modifica Excel ni los JSON de entrada y no envia
notificaciones. `sendNotification` es una indicacion para consumidores externos.
No hay menu numerico: el numero 5 identifica el caso de uso en el README.

## Ejecucion

```powershell
.\.venv\Scripts\python.exe -m sprint_cert_automation.cli validate-myte-fortnight `
  --excel-json ".\out\excel-current-period-cost-hours-excel\excel-cost-hours_2026-09.json" `
  --myte-json ".\inputs\myte_2026-09_q1.json" `
  --period-number 1 `
  --output-file-name "myte_validation_2026-09_q1"
```

| Argumento | Contrato |
|-----------|----------|
| `--excel-json` | Ruta obligatoria al JSON de horas previstas |
| `--myte-json` | Ruta obligatoria al JSON extraido de myTE |
| `--period-number` | Entero obligatorio, 1 o 2 |
| `--output-file-name` | Nombre obligatorio; admite extension .json opcional, no rutas |
| `--encoding` | Lectura y escritura; por defecto utf-8 |
| `--indent` | Indentacion de salida; por defecto 2 |
| `--pretty-log` | Resumen stdout como objeto JSON con total, passed y failed |

Salida fija: `out/myte-fortnight-validation/<OutputFileName>.json`, relativa al
proyecto, no al directorio de trabajo. No se admite `--dry-run`.

El resumen normal muestra filas evaluadas, Passed, Failed y ruta de salida.
`--pretty-log` conserva la forma del resumen original, sin envolver el informe.

## Contratos de entrada

Ambas entradas deben ser arrays de objetos. Todos los campos siguientes son
obligatorios y de tipo string; no se convierten automaticamente numeros o null.
Los campos adicionales no intervienen en la validacion.

### Excel

```json
[
  {
    "reviewee": "alice",
    "email": "alice@example.com",
    "subco": "0",
    "horasQuincena1": "80",
    "horasQuincena2": "72"
  }
]
```

El caso 3, `export-cost-hours-json`, produce este contrato sin transformaciones
adicionales. Debe garantizarse que `reviewee` representa el mismo identificador
en ambas fuentes.

### myTE

```json
[
  {
    "reviewee": "alice",
    "trStatus": "Processed",
    "chargeCode": "FUNDAE (COMK3001)",
    "totalHours": "67.2",
    "workingHours": "80",
    "percentageTotalHoursVsWorkingHours": "84%"
  }
]
```

`totalHours` es obligatorio como string, pero no se compara ni se parsea.
El porcentaje se toma de `percentageTotalHoursVsWorkingHours`, no se recalcula.

## Contrato de salida

Array plano, en el orden de Excel, excluyendo los trabajadores con subco=1.
Cada resultado tiene exactamente 13 propiedades obligatorias:

```json
[
  {
    "reviewee": "alice",
    "email": "alice@example.com",
    "validationStatus": "Passed",
    "horasQuincena1": "80",
    "horasQuincena2": "72",
    "trStatus": "Processed",
    "chargeCode": "FUNDAE (COMK3001)",
    "totalHours": "67.2",
    "workingHours": "80",
    "percentageTotalHoursVsWorkingHours": "84%",
    "comments": "",
    "sendNotification": false,
    "notificationMessage": ""
  }
]
```

- Se conservan exactamente `reviewee`, `email`, `horasQuincena1` y
  `horasQuincena2` de Excel. Ambas quincenas se incluyen, independientemente de
  `--period-number`.
- `trStatus`, `chargeCode`, `totalHours`, `workingHours` y
  `percentageTotalHoursVsWorkingHours` se copian del registro myTE correspondiente.
  Si no existe correspondencia, los cinco campos son strings vacios (`""`).
- Los datos de ambas fuentes mantienen sus strings originales, incluidos espacios,
  comas, decimales y `%`, sin normalizacion, redondeo ni recalculo para la salida.
  Se incluyen tambien en resultados Failed y validaciones cortocircuitadas,
  sin introducir parseos numericos adicionales.
- `subco` y los campos adicionales de las entradas no se incluyen en la salida.
- `validationStatus`: Passed o Failed; no se adapta a los estados de Beeline.
- `sendNotification`: booleano, true solo para New/Draft.
- `notificationMessage`: siempre string vacio.
- Arrays vacios o todos los trabajadores excluidos producen `[]`.

La ampliacion es aditiva y no cambia las reglas ni los contratos de entrada.
Los consumidores que exijan exactamente las seis propiedades anteriores deben
actualizar su esquema para aceptar las 13 propiedades.

## Reglas y prioridad

### Prevalidacion global

1. Validar que ambas entradas sean arrays y sus elementos objetos.
2. Validar todos los campos obligatorios y tipos, primero Excel y luego myTE.
3. Indexar por `reviewee.strip().lower()`, primero Excel y luego myTE.
4. Rechazar duplicados normalizados en cualquiera de las entradas.

Esta fase sucede **antes** de las exclusiones y descartes. Por tanto, un registro
con subco=1 o un registro myTE sin correspondencia tambien aborta si su contrato
es invalido o introduce duplicados. No se genera un resultado Failed por estos
errores: se aborta la ejecucion.

### Regla 0: subco

Si `subco.strip()` es `"1"`, se omite el trabajador de la salida. `"0"` y cualquier
otro string siguen el flujo normal. No se introduce una restriccion nueva a 0/1.

### Reglas 1 y 2: correspondencia

Los registros myTE que no existen en Excel no producen salida.
Para un trabajador de Excel sin myTE:

- Failed, sendNotification=false.
- Se conservan las dos horas quincenales de Excel; los cinco campos myTE son `""`.
- comments: `El reviewee registrado en el fichero excel no ha sido encontrado en myTE`

### Regla 3: estado

Se eliminan espacios exteriores de `trStatus`, pero se conserva la distincion
entre mayusculas y minusculas:

| Estado | Resultado | Notificacion | comments |
|--------|-----------|--------------|----------|
| New o Draft | Failed | true | `myTE no enviado. TR Status es New o Draft.` |
| Distinto de Processed | Failed | false | `Estado diferente de Processed. Revisar` |
| Processed | Continuar | false | Se aplican horas y porcentaje |

Los dos fallos detienen las validaciones del trabajador: no se parsean horas ni
porcentajes. Por ejemplo, `processed` no equivale a `Processed`.

### Regla 4: horas

- Quincena 1 selecciona `horasQuincena1`; quincena 2, `horasQuincena2`.
- Se comparan con `workingHours` tras aplicar `round(float_normalizado, 1)` a ambos.
- Si no coinciden: Failed, sendNotification=false.
- comments: `El numero de WorkingHours registradas en myTE es diferente al previsto en Excel.`
- El fallo corta el flujo y no se parsea el porcentaje.

Se mantiene `float` y el redondeo nativo de Python, no Decimal ni tolerancias.
El redondeo conserva los efectos de representacion binaria del motor original.

### Regla 5: porcentaje y chargeCode

El porcentaje normalizado no se redondea antes de aplicar la politica:

| Codigo contenido en chargeCode | Politica |
|--------------------------------|----------|
| COMK3002 | 91.0 a 93.0, ambos incluidos |
| COMK3001, reviewee normalizado patricia.c.martin o e.a.magana.ramirez | Exactamente 80.0 |
| COMK3001, resto | 82.0 a 86.0, ambos incluidos |
| Ninguno de los anteriores | Failed |

Los codigos distinguen mayusculas/minusculas. Si aparecen ambos, COMK3002 tiene
prioridad, conservando el comportamiento original.

Si cumple: Passed, sendNotification=false, comments vacio.
Si no cumple: Failed, sendNotification=false y:

`El porcentaje <valor_original> no cumple la politica definida para el chargeCode.`

Ejemplo: `El porcentaje <87.2%> no cumple la politica definida para el chargeCode.`
Se conserva el string original del porcentaje dentro de los signos `< >`.

### Normalizacion numerica

Para horas y porcentaje, se eliminan espacios exteriores y simbolos `%`, se
sustituye coma por punto y se convierte con `float`. El simbolo % es opcional.
Un valor no convertible aborta con:

`Valor numerico invalido en '<campo>' para reviewee '<reviewee>': <valor>`

Solo se parsean los campos alcanzados segun el flujo. No se valida numericamente
la quincena no seleccionada, totalHours ni registros excluidos o descartados.
Se conserva el parseo original, sin restricciones adicionales para valores
negativos, strings vacios de identidad o valores especiales aceptados por float.

## Errores, publicacion y codigos de salida

- 0: informe publicado correctamente, aunque todos sus resultados sean Failed.
- 1: error conocido de fichero, contrato, duplicados, numero o codificacion;
  mensaje explicito en stderr.
- 2: argumentos CLI ausentes o invalidos, gestionados por argparse.

La quincena se comprueba tambien desde App/servicio/dominio. Esta proteccion evita
que una llamada directa interprete una quincena invalida como quincena 2.

El informe se calcula completamente antes de escribir. Se publica mediante un
temporal en el mismo directorio y reemplazo del destino; los temporales se limpian
si falla. Un informe previo permanece intacto si falla la validacion, escritura
o reemplazo. Si existe un destino valido, una ejecucion correcta lo reemplaza.
No se permite que la ruta de salida coincida con ninguna entrada.

## Validacion automatizada

Las pruebas incluyen contratos exactos de 13 propiedades en todas las ramas,
copia literal de datos de ambas fuentes, campos myTE vacios sin correspondencia,
exclusiones, orden, normalizacion,
duplicados en ambos origenes, cortocircuitos, ambas quincenas, redondeo,
porcentajes limite, excepciones personales, codigos desconocidos, errores de IO,
publicacion atomica, opciones de codificacion y CLI.

La prueba de integracion crea un workbook sintetico, ejecuta el caso 3 y consume
directamente su salida en el caso 5 para ambas quincenas, sin modificar el Excel.
No se necesitan datos personales, acceso a myTE ni la carpeta de importacion.

```powershell
.\.venv\Scripts\python.exe -m pytest `
  tests\test_myte_validation.py `
  tests\test_myte_fortnight_validation_service.py `
  tests\test_cli.py -q
```
