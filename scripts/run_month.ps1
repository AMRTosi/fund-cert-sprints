param(
  [Parameter(Mandatory = $true)] [string]$Forecast,
  [Parameter(Mandatory = $true)] [string]$Template,
  [Parameter(Mandatory = $true)] [int]$Year,
  [Parameter(Mandatory = $true)] [int]$Month,
  [string]$OutputDir,
  [switch]$DryRun
)

$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $pythonExe)) {
  throw "No se ha encontrado $pythonExe. Ejecuta scripts/setup_env.ps1 dentro de cert_automation para crear el entorno local."
}

if ($OutputDir) {
  Write-Warning "-OutputDir se ignora. La salida siempre se genera en cert_automation/certificaciones/YYYY-MM."
}

$cliArgs = @(
  "-m", "sprint_cert_automation.cli",
  "--forecast", "$Forecast",
  "--template", "$Template",
  "--year", "$Year",
  "--month", "$Month"
)

if ($DryRun) {
  $cliArgs += "--dry-run"
}

& $pythonExe @cliArgs
