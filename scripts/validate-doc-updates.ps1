[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($env:DOC_HOOK_ALLOW_NO_DOCS -eq '1') {
    Write-Host '[doc-hook] Validation skipped by DOC_HOOK_ALLOW_NO_DOCS=1'
    exit 0
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Push-Location $repoRoot
try {
    $changed = git diff --name-only -- .
    if (-not $changed) {
        # Nothing changed in tracked files.
        exit 0
    }

    $files = @($changed -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })

    $codeChanged = $files | Where-Object {
        $_ -match '^src/' -or
        $_ -match '^tests/' -or
        $_ -match '^pyproject\.toml$' -or
        $_ -match '^requirements\.txt$'
    }

    if (-not $codeChanged) {
        exit 0
    }

    $docsChanged = $files | Where-Object {
        $_ -match '^README\.md$' -or
        $_ -match '^docs/.+\.md$'
    }

    if ($docsChanged) {
        Write-Host '[doc-hook] OK: code changes include documentation updates.'
        exit 0
    }

    Write-Error @'
[doc-hook] Blocking stop: detected code changes without documentation updates.
Update at least one of these files when behavior/flow changes:
- README.md
- docs/ARCHITECTURE.md
- docs/Generate_Sprint_Certificates.md
- docs/Template_Mes-Details.md
If this change intentionally does not require docs, rerun with DOC_HOOK_ALLOW_NO_DOCS=1.
'@
}
finally {
    Pop-Location
}
