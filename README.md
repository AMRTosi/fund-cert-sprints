# Sprint Certificate Automation

Herramienta CLI para automatizar la gestión de certificaciones de sprints y la preparación de periodos mensuales en el libro Excel de Forecast Fundae.

## Casos de uso

### 1. Generación de certificados de sprint

Genera ficheros `.xlsm` de certificación mensual a partir del libro Forecast y una plantilla de certificación.

**Flujo completo:**

```powershell
# 1. Dry-run para validar entradas
.\.venv\Scripts\python.exe -m sprint_cert_automation.cli `
  --forecast "./inputs/Fundae_Forecast_Copilot.xlsx" `
  --template "./inputs/plantilla de Inf_Certificacion 20251.xlsm" `
  --year 2026 --month 6 --dry-run

# 2. Generación real de certificados
.\.venv\Scripts\python.exe -m sprint_cert_automation.cli `
  --forecast "./inputs/Fundae_Forecast_Copilot.xlsx" `
  --template "./inputs/plantilla de Inf_Certificacion 20251.xlsm" `
  --year 2026 --month 6

# 3. Revisión manual en Excel de los ficheros generados

# 4. Exportación a PDF (ejecuta macro + exporta)
.\.venv\Scripts\python.exe -m sprint_cert_automation.cli export-pdf `
  --year 2026 --month 6
```

**Salida:** `certificaciones/2026-06/` con un `.xlsm` y `.pdf` por cada sprint facturable.

**Qué hace internamente:**
- Lee el forecast en modo solo lectura.
- Detecta ventanas de sprint desde las hojas FY (mes anterior + mes objetivo).
- Aplica reglas de facturación: un sprint es facturable si su fecha fin cae en el mes objetivo y no está rayado (hatched).
- Calcula festivos, horas de sprint y horas libres por técnico.
- Copia la plantilla `.xlsm`, rellena la hoja `Config` y las tablas de equipo.
- Genera un certificado por sprint facturable y empaqueta en ZIP.

---

### 2. Duplicado de periodo mensual

Crea una nueva pestaña de periodo en el libro Forecast a partir de `Template_Mes`, configurándola automáticamente para el mes indicado.

**Uso básico (sin periodo anterior):**

```powershell
.\.venv\Scripts\python.exe -m sprint_cert_automation.cli duplicate-sheet `
  --forecast "./inputs/Fundae_Forecast_Copilot_v2.xlsx" `
  --source "Template_Mes" `
  --target "FY27_dic" `
  --year 2026 --month 12 --dry-run
```

**Uso completo (con periodo anterior para carry-over de sprints):**

```powershell
.\.venv\Scripts\python.exe -m sprint_cert_automation.cli duplicate-sheet `
  --forecast "./inputs/Fundae_Forecast_Copilot_v2.xlsx" `
  --source "Template_Mes" `
  --target "FY27_dic" `
  --year 2026 --month 12 `
  --previous "FY27_nov"
```

**Pasos que ejecuta automáticamente:**

| Paso | Acción | Detalle |
|------|--------|---------|
| 1 | Copiar pestaña | Duplica la hoja fuente y la renombra |
| 2 | Posicionar | La sitúa inmediatamente después de la fuente |
| 3 | Calendario | Rellena números de día (fila 5) y letras de día en español (fila 6) |
| 4 | Limpiar grises | Elimina rellenos grises de T_COST_HOURS_ONLY |
| 5 | Fórmulas de coste | Rellena celdas vacías en días laborables con fórmulas de horas |
| 6 | Configurar sprints | Calcula T_SPRINTS (filas 1-4) con carry-over del periodo anterior |
| 7 | Gap Periodo Anterior | Fórmulas que suman horas de revenue del sprint no finalizado en el mes anterior |
| 8 | Revenues Mes Actual | Fórmulas que suman horas de revenue del rango facturable de cada equipo |
| 9 | Revenues No Facturable | Fórmulas que suman horas de revenue del rango NO facturable (sprint incompleto) |

**Reglas de sprints por equipo:**

| Equipo | Duración | Carry-over | Nomenclatura |
|--------|----------|------------|--------------|
| Bonificaciones | 10 días laborables | Sí (hatched=darkDown) | `Bonificaciones - SP{N}` |
| Subvenciones | 16 días laborables | Sí (hatched=lightDown) | `Subvenciones - SP{N}` |
| Fondos de Reserva | Días 1-15 y 16-último | No | `Fondos de Reserva - SP{N}` |
| Transversal | Mes completo | No | `Transversal - {mes_español}` |

---

## Estructura del proyecto

```
cert_automation/
├── config/                  # Mappings y configuración local
├── docs/                    # Documentación funcional
├── inputs/                  # Ficheros Excel (gitignored)
├── scripts/                 # Scripts de arranque Windows
├── src/sprint_cert_automation/
│   ├── cli.py              # Punto de entrada CLI (argparse)
│   ├── app.py              # Orquestación de casos de uso
│   ├── domain/             # Modelos y reglas de negocio
│   │   ├── models.py       # Dataclasses (Sprint, Technician, etc.)
│   │   └── rules.py        # Reglas de facturación
│   ├── services/           # Lógica de aplicación
│   │   ├── certificate_service.py   # Generación de certificados
│   │   ├── forecast_reader.py       # Lectura del forecast
│   │   ├── sheet_duplicator.py      # Duplicado de periodos
│   │   ├── sprint_configurator.py   # Cálculo y escritura de sprints
│   │   ├── template_writer.py       # Escritura en plantilla .xlsm
│   │   └── macro_export_service.py  # Export PDF via COM
│   ├── infrastructure/     # Adaptadores (Excel COM)
│   └── utils/              # Utilidades (fechas, etc.)
├── tests/                   # Tests unitarios
└── certificaciones/         # Salida generada (gitignored)
```

## Setup

```powershell
# Crear entorno virtual y dependencias
.\scripts\setup_env.ps1

# O manualmente:
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip install -e .

# Ejecutar tests
.\.venv\Scripts\python.exe -m pytest
```

El install editable (`-e .`) es necesario para que `python -m sprint_cert_automation.cli` resuelva el paquete desde `src/`.

## Notas

- Los ficheros Excel de entrada van en `inputs/` (ignorado por Git).
- La generación de certificados **nunca modifica** el workbook de forecast.
- El duplicado de periodos **sí modifica** el workbook de forecast (crea una nueva pestaña).
- Salida de certificados siempre en `certificaciones/<YYYY-MM>`.
- Usar `--dry-run` en cualquier comando para previsualizar sin efectos.

## Troubleshooting

| Error | Causa | Solución |
|-------|-------|----------|
| `Technician column 'Técnico' not found` | La hoja FY no tiene columna Técnico | Verificar cabecera en fila 6 |
| `Forecast sheet not found: FYxx_mon` | Año/mes no coincide con pestañas | Comprobar nombres de pestañas |
| `Category '...' not available in template` | Categoría sin mapeo en plantilla | Ajustar alias en `template_writer.py` |
| `No module named ...` | Entorno virtual roto | Ejecutar `.\scripts\setup_env.ps1` |
| `Cannot run the macro` | Seguridad de macros | Habilitar macros y ubicación de confianza |
| Warnings de openpyxl | Extensiones Excel no soportadas | Ignorar; validar salida en Excel |
