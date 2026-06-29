param(
  [string]$PythonLauncher = "py"
)

$projectRoot = Split-Path -Parent $PSScriptRoot
$venvDir = Join-Path $projectRoot ".venv"
$pythonExe = Join-Path $venvDir "Scripts\python.exe"

Push-Location $projectRoot
try {
  if (-not (Test-Path $pythonExe)) {
    if (Get-Command $PythonLauncher -ErrorAction SilentlyContinue) {
      & $PythonLauncher -m venv $venvDir
    }
    elseif (Get-Command python -ErrorAction SilentlyContinue) {
      python -m venv $venvDir
    }
    else {
      throw "No se ha encontrado ni '$PythonLauncher' ni 'python' para crear el entorno virtual."
    }
  }

  & $pythonExe -m pip install --upgrade pip
  & $pythonExe -m pip install -r requirements.txt
  & $pythonExe -m pip install -e .
}
finally {
  Pop-Location
}