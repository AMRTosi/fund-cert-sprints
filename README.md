# Sprint Certificate Automation

Generate monthly sprint certification `.xlsm` files from a Forecast workbook and a certification template.

## Architecture

- Python business engine:
  - Reads forecast workbook in read-only mode.
  - Detects sprint windows from FY sheets.
  - Applies billing rules (sprint ends in target month and is not hatched).
  - Computes holidays, sprint hours, and free hours per technician.
- Template writer:
  - Copies the `.xlsm` template.
  - Fills `Config` sheet fields and team tables.
  - Writes one certificate per billable sprint.
- Packaging:
  - Produces one ZIP with all generated certificates.

## Folder layout

- `config/`: local mappings and rules config.
- `src/sprint_cert_automation/`: python source code.
- `scripts/`: launch scripts for Windows.
- `tests/`: unit tests for pure business rules.

## Use Every Time You Open VS Code

Use this checklist from inside `cert_automation`.

1. Open terminal in project folder:

```powershell
cd cert_automation
```

2. Ensure local environment exists and is up to date:

```powershell
.\scripts\setup_env.ps1
```

3. (Optional) Activate venv for interactive commands:

```powershell
.\.venv\Scripts\Activate.ps1
```

4. Run tests before generating files:

```powershell
.\.venv\Scripts\python.exe -m pytest
```

5. Run monthly generation in dry-run mode first:

```powershell
.\.venv\Scripts\python.exe -m sprint_cert_automation.cli `
  --forecast "./inputs/Fundae_Forecast_Copilot.xlsx" `
  --template "./inputs/plantilla de Inf_Certificacion 2025.xltm" `
  --year 2026 `
  --month 6 `
  --dry-run
```

Important: `--dry-run` does not write files. It only validates inputs and shows planned output paths.

Output path is fixed to:

`cert_automation/certificaciones/<YYYY-MM>`

For the command above, output folder is:

`cert_automation/certificaciones/2026-06`

6. Run real generation:

```powershell
.\.venv\Scripts\python.exe -m sprint_cert_automation.cli `
  --forecast "./inputs/Fundae_Forecast_Copilot.xlsx" `
  --template "./inputs/plantilla de Inf_Certificacion 2025.xltm" `
  --year 2026 `
  --month 6
```

## Quick start (first run)

1. Create and populate local virtual environment in `cert_automation/.venv`:

```powershell
.\scripts\setup_env.ps1
```

2. Optional: activate local environment for interactive work:

```powershell
.\.venv\Scripts\Activate.ps1
```

3. Install dependencies manually through local environment only (if needed):

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip install -e .
```

The editable install is required so `python -m sprint_cert_automation.cli` resolves the package from `src/`.

4. Run monthly generation (dry-run first):

```powershell
.\venv\Scripts\python.exe -m sprint_cert_automation.cli `
  --forecast "./inputs/Fundae_Forecast_Copilot.xlsx" `
  --template "./inputs/plantilla de Inf_Certificacion 2025.xltm" `
  --year 2026 `
  --month 6 `
  --dry-run
```

6. Run real generation:

```powershell
.\venv\Scripts\python.exe -m sprint_cert_automation.cli `
  --forecast "./inputs/Fundae_Forecast_Copilot.xlsx" `
  --template "./inputs/plantilla de Inf_Certificacion 2025.xltm" `
  --year 2026 `
  --month 6
```

## Notes

- Keep sensitive forecast/template files inside `cert_automation/inputs/`.
- `inputs/` and all Excel extensions (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) are ignored by Git.
- The process never modifies the forecast source workbook.
- If generation fails due to missing FY sheet names, verify year/month and workbook tabs.
- Generated files are always written under `cert_automation/certificaciones/<YYYY-MM>`.

## Troubleshooting

### 1) `Technician column 'Técnico' not found in forecast sheet`

Cause:
- The parser now requires `Técnico` header in the sheet used to build workloads.

Fix:
- Verify the target FY sheet has a `Técnico` column in header row 6.
- If month structure differs, update forecast source or parser mapping accordingly.

### 2) `Forecast sheet not found: FYxx_mon`

Cause:
- The requested `--year` and `--month` do not match existing FY tab names.

Fix:
- Check workbook sheet names (`FY26_mayo`, `FY26_jun`, etc.).
- Re-run with the correct year/month.

### 3) `Category '...' is not available in template dropdown options`

Cause:
- A category from forecast cannot be mapped to template table `TB_Perfiles`.

Fix:
- Confirm category exists in template (`Maestra` -> `TB_Perfiles`).
- Add/adjust alias mapping in `template_writer.py` if needed.

### 4) `No module named ...` or package/import errors

Cause:
- Virtual environment is missing or stale.

Fix:
- Recreate/update environment:

```powershell
.\scripts\setup_env.ps1
```

### 5) I need a custom output path

Cause:
- Current behavior enforces a fixed output location.

Fix:
- Use generated files from:

`cert_automation/certificaciones/<YYYY-MM>`

- The `-OutputDir` argument in `scripts/run_month.ps1` is currently ignored.

### 6) `scripts/run_month.ps1` warns that `-OutputDir` is ignored

Cause:
- This is expected after enforcing fixed output path.

Fix:
- Remove `-OutputDir` from your manual calls.
- Keep using `--year` and `--month` to control destination subfolder.

### 7) Warnings from `openpyxl` about unsupported extensions

Cause:
- The template contains Excel extensions (`Data Validation`, `Conditional Formatting`, etc.) that `openpyxl` warns about.

Fix:
- Warnings are expected in current implementation.
- Validate generated output in Excel; if strict fidelity is required, move template write path to Excel COM.

## Current status

Current implementation covers end-to-end generation:

- Forecast sprint detection and billable filtering.
- Holiday detection and team workload extraction.
- Category normalization against template dropdown values.
- Certificate writing and zip packaging.
