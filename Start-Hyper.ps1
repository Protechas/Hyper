$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$hyperPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
$hyperReady = Join-Path $PSScriptRoot '.venv\hyper-ready'
if (-not (Test-Path -LiteralPath $hyperPython)) {
    # Prefer the working Python environment already shipped with this checkout.
    # The Windows Store app-execution alias can appear as `python.exe` even when
    # it cannot actually launch, so validate every fallback before using it.
    $bootstrapPython = Join-Path $PSScriptRoot 'env\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $bootstrapPython)) {
        $pythonCommand = Get-Command python.exe -CommandType Application -ErrorAction SilentlyContinue
        $bootstrapPython = if ($pythonCommand) { $pythonCommand.Source } else { $null }
    }

    $pythonWorks = $false
    if ($bootstrapPython) {
        try {
            & $bootstrapPython -c 'import sys' 2>$null
            $pythonWorks = ($LASTEXITCODE -eq 0)
        }
        catch {
            $pythonWorks = $false
        }
    }

    if (-not $pythonWorks) {
        throw 'Could not find a working Python installation. Install Python 3.11 or restore the repository env folder.'
    }

    & $bootstrapPython -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "Could not create the Python environment using $bootstrapPython." }
}
if (-not (Test-Path -LiteralPath $hyperReady)) {
    # Use the Windows certificate store so corporate/root certificates trusted
    # by the workstation are also honored by pip.
    & $hyperPython -m pip install --use-feature=truststore -r requirements.txt
    if ($LASTEXITCODE -ne 0) { throw 'Could not install Hyper dependencies.' }
    New-Item -ItemType File -Path $hyperReady -Force | Out-Null
}
& $hyperPython HyperWeb.py
if ($LASTEXITCODE -ne 0) { throw "Hyper closed with error code $LASTEXITCODE. Check the latest file in Documents\Hyper Logs for details." }
