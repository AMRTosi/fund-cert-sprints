# Arquitectura y diseño técnico

Documento de referencia técnica para futuras modificaciones del código.

## Visión general

El sistema implementa dos casos de uso principales como comandos CLI:

1. **generate** — Genera certificados de sprint (`.xlsm` + PDF) a partir del forecast.
2. **duplicate-sheet** — Duplica y configura una pestaña de periodo mensual en el forecast.

Ambos comparten la misma estructura de capas y convenciones de código.

## Capas de la arquitectura

```
CLI (cli.py)
  ↓ parsea argumentos
App (app.py)
  ↓ orquesta caso de uso
Services (services/)
  ↓ lógica de aplicación
Domain (domain/)
  ↓ modelos y reglas puras
Infrastructure (infrastructure/)
  ↓ adaptadores externos (Excel COM)
```

### CLI (`cli.py`)

- Punto de entrada: `python -m sprint_cert_automation.cli`
- Usa `argparse` con subcomandos: `generate` (implícito), `export-pdf`, `duplicate-sheet`
- Solo parsea argumentos y delega a `app.py`

### App (`app.py`)

- Capa de orquestación fina (thin orchestration)
- Cada función pública es un caso de uso: `generate_certificates()`, `export_certificates_to_pdf()`, `duplicate_period_sheet()`
- Instancia servicios y devuelve dataclasses de resultado

### Services (`services/`)

| Módulo | Responsabilidad |
|--------|----------------|
| `forecast_reader.py` | Lee hojas FY del forecast: sprints, festivos, técnicos, horas |
| `certificate_service.py` | Orquesta la generación de certificados end-to-end |
| `template_writer.py` | Escribe datos en la plantilla `.xlsm` (Config, tablas) |
| `sheet_duplicator.py` | Duplica pestaña Template_Mes y adapta contenido |
| `sprint_configurator.py` | Calcula y escribe sprints (T_SPRINTS filas 1-4) |
| `macro_export_service.py` | Ejecuta macro VBA y exporta a PDF vía COM |

### Domain (`domain/`)

- `models.py` — Dataclasses puras: `Sprint`, `Technician`, `SprintWindow`, etc.
- `rules.py` — Reglas de negocio sin dependencias externas (facturación, filtrado)

### Infrastructure (`infrastructure/`)

- `excel_com.py` — Interacción con Excel vía COM (pywin32) para macros y PDF

## Estructura del workbook Forecast

### Layout de pestaña de periodo (Template_Mes / FY sheets)

```
Filas 1-4:   T_SPRINTS (una fila por equipo)
Fila 5:      Números de día (1-31)
Fila 6:      Letras de día español (L,M,X,J,V,S,D) + cabeceras de columnas resumen
Filas 7-37:  T_COST_HOURS_ONLY (horas de coste por técnico)
Filas 41-71: T_REVENUES_TIMESHEET (horas de revenue por técnico)
```

### Columnas clave

| Rango | Contenido |
|-------|-----------|
| A-L (aprox.) | Datos fijos: Target, Nombre, Técnico, Facturación, Equipo, Tarifa... |
| M-AQ (cols 13-43) | Días 1-31 del calendario |
| A partir de col 44 | Columnas resumen: Horas myTE, Target, Factura, CCI, Gap Anterior, Mes Actual, etc. |

**Nota:** La posición de la columna día-1 varía entre hojas (col 10-13). El código detecta dinámicamente buscando `value==1` en fila 5.

### Columna de equipo

La columna que contiene el nombre del equipo (Transversal, Bonificaciones, Subvenciones, Fondos de Reserva) varía entre hojas:
- Puede ser D, E o F según la estructura de la hoja
- El código detecta dinámicamente buscando cabecera "Equipo" en fila 6, o escaneando valores conocidos en filas de datos

### Nomenclatura de pestañas FY

`FY{año_fiscal}_{abreviatura_mes}` — El año fiscal empieza en septiembre.

| Pestaña | Fecha |
|---------|-------|
| FY27_sept | Septiembre 2026 |
| FY27_ene | Enero 2027 |
| FY26_mayo | Mayo 2026 |

Conversión: meses sept-dic → año calendario = 2000+FY-1; meses ene-ago → año calendario = 2000+FY.

## Diseño del duplicado de periodo (`sheet_duplicator.py`)

### Flujo principal: `duplicate_sheet()`

```python
duplicate_sheet(workbook_path, source_sheet_name, new_sheet_name, year, month, previous_sheet_name, dry_run)
```

Pasos secuenciales:
1. Copia hoja fuente → `wb.copy_worksheet()`
2. Posiciona después de la fuente → `wb.move_sheet()`
3. Calendario → `update_calendar()` (fila 5: números, fila 6: letras)
4. Limpia grises → `remove_gray_fills()` (theme=2, tint<-0.05)
5. Fórmulas de coste → `fill_empty_cost_cells()`
6. Sprints → `configure_sprints()` (delegado a `sprint_configurator.py`)
7. Gap Periodo Anterior → `update_gap_anterior_formulas()`
8. Revenues Mes Actual → `update_revenues_mes_actual_formulas()`
9. Revenues No Facturable → `update_revenues_no_fact_formulas()`

### Constantes del layout

```python
FIRST_DAY_COL = 13       # Columna M = día 1
MAX_DAY_COL = 43         # Columna AQ = día 31
DAY_NUMBER_ROW = 5
DAY_LETTER_ROW = 6
COST_FIRST_ROW = 7
COST_LAST_ROW = 37
REVENUE_FIRST_ROW = 41
REVENUE_LAST_ROW = 71
REVENUE_OFFSET = 34      # revenue_row = cost_row + 34
```

### Detección de columnas resumen

Las columnas resumen (Gap Anterior, Mes Actual, etc.) no tienen posición fija. Se detectan escaneando la fila 6 desde la columna 44:

```python
_find_gap_anterior_columns(ws)    # Busca "Gap Mes Anterior"
_find_mes_actual_columns(ws)      # Busca "Mes Actual" (excluye "Gap Mes Actual")
_find_gap_mes_actual_columns(ws)  # Busca "Gap Mes Actual"
```

### Patrón de fórmulas generadas

Todas las fórmulas de resumen siguen el mismo patrón nested-IF por equipo:

```
=IF(team_col{row}="Transversal", expr_transv,
  IF(team_col{row}="Bonificaciones", expr_bonif,
    IF(team_col{row}="Subvenciones", expr_subv,
      IF(team_col{row}="Fondos de Reserva", expr_fdr, 0))))
```

Cada columna resumen tiene dos instancias: Horas y Factura (= Horas × Tarifa).

| Columna | Referencia | Rango |
|---------|-----------|-------|
| Gap Periodo Anterior | Revenue rows de pestaña **anterior** | Hatched tail del sprint anterior |
| Mes Actual | Revenue rows de pestaña **actual** | Día 1 → último día del último sprint sólido |
| Gap Mes Actual (No Fact) | Revenue rows de pestaña **actual** | Día después del último sprint sólido → último día del mes |

## Diseño del configurador de sprints (`sprint_configurator.py`)

### Modelo de datos

```python
@dataclass
class SprintSegment:
    name: str              # "Bonificaciones - SP264"
    sprint_number: int     # 264
    start_day: int         # día calendario (1-based)
    end_day: int           # día calendario (1-based)
    is_hatched: bool       # True si el sprint continúa en el siguiente mes
    team: str              # "bonificaciones", "subvenciones", "fdr", "transversal"

@dataclass
class TeamSprintInfo:
    segments: list[SprintSegment]
    # Propiedades: last_sprint_number, has_hatched_tail, hatched_tail
```

### Lectura de sprints: `read_sprints_from_sheet()`

- Detecta celdas merged y unmerged en filas 1-4
- Identifica equipo por fila (1=bonif, 2=subv, 3=fdr, 4=transv)
- Detecta hatched por `patternType` del fill (darkDown=bonif, lightDown=subv)
- Devuelve `dict[str, TeamSprintInfo]`

### Cálculo de sprints nuevos

Cada equipo tiene su función de cálculo:

| Función | Reglas |
|---------|--------|
| `calculate_bonif_sprints()` | 10 días laborables, carry-over, martes→lunes |
| `calculate_subv_sprints()` | 16 días laborables, carry-over |
| `calculate_fdr_sprints()` | Fijo: días 1-15, 16-último |
| `calculate_transversal_sprint()` | Mes completo, nombre = mes en español |

### Carry-over desde periodo anterior

Si el periodo anterior tiene un sprint hatched (no finalizado):
1. Se leen los días que ya ejecutó en el periodo anterior
2. Los días restantes se ejecutan al inicio del nuevo mes (sólido, sin hatched)
3. La numeración del nuevo sprint continúa desde el anterior

### Escritura: `write_sprints_to_sheet()`

- Limpia filas 1-4 (desmerge, borra valores y fills)
- Escribe cada segmento: merge de celdas, valor del nombre, fill correspondiente
- Aplica `Alignment(horizontal="center", vertical="center")`

### Fills por equipo

```python
BONIF_SOLID_FILL   = PatternFill("solid", fgColor=Color(theme=3, tint=0.8999...))
BONIF_HATCHED_FILL = PatternFill("darkDown", bgColor=Color(theme=3, tint=0.8999...))
SUBV_SOLID_FILL    = PatternFill("solid", fgColor=Color(theme=5, tint=0.7999...))
SUBV_HATCHED_FILL  = PatternFill("lightDown", bgColor=Color(theme=5, tint=0.7999...))
FDR_SOLID_FILL     = PatternFill("solid", fgColor=Color(rgb="FFC0F1C8"))
TRANSV_SOLID_FILL  = PatternFill("solid", fgColor=Color(theme=8, tint=0.7999...))
```

## Diseño de la generación de certificados

### Flujo: `CertificateGenerationService.run()`

1. `ForecastReader` lee las hojas FY (mes anterior + objetivo)
2. Detecta sprints facturables (fecha fin en mes objetivo, no hatched)
3. Para cada sprint facturable:
   - Calcula ventana de fechas, festivos, horas
   - Extrae técnicos del equipo con sus datos
   - `TemplateWriter` copia plantilla y rellena campos
4. Empaqueta todo en ZIP

### Reglas de facturación (`domain/rules.py`)

- Un sprint es facturable si su fecha fin cae en el mes objetivo Y no está rayado
- Horas sprint = días laborables × horas/día (excluyendo festivos)
- Días no laborables por técnico = celdas grises dentro del rango del sprint

## Convenciones de código

### Testing

- Tests unitarios en `tests/` con `pytest`
- Cada servicio tiene su fichero de test: `test_sheet_duplicator.py`, `test_sprint_configurator.py`
- Se usan workbooks en memoria (openpyxl.Workbook()) sin ficheros temporales cuando es posible
- Para tests que necesitan persistencia se usa `tmp_path` de pytest

### openpyxl — Notas importantes

- `load_workbook()` requiere `str(path)`, no objetos Path directamente
- `cell.value = None` para limpiar (no `ws.cell(row, col, value=None)`)
- `copy_worksheet()` crea la hoja al final; usar `move_sheet()` para reposicionar
- Merged cells: se leen via `ws.merged_cells.ranges`
- Fills: comparar `theme`, `tint`, `patternType` — no comparar objetos directamente

### Extensibilidad

Para añadir una nueva columna de fórmulas al duplicado:
1. Crear función `update_xxx_formulas(ws, ...)` en `sheet_duplicator.py`
2. Añadir función helper `_find_xxx_columns()` para detectar la columna
3. Conectar en `duplicate_sheet()` como paso N+1
4. Añadir tests en `test_sheet_duplicator.py`
5. Actualizar docstring de `duplicate_sheet()` con el nuevo paso

## Referencias cruzadas con documentación funcional

| Documento | Contenido |
|-----------|-----------|
| `docs/Generate Sprint Certificates.md` | Reglas de negocio de certificación: estructura del forecast, facturación, festivos, técnicos |
| `docs/Detalles de plantilla Template_Mes para duplicación de periodos.md` | Estructura de Template_Mes, reglas de sprints, fills, fórmulas, tablas |
| `README.md` | Guía de uso, comandos CLI, troubleshooting |
