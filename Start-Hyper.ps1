$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$hyperPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
$hyperReady = Join-Path $PSScriptRoot '.venv\hyper-ready'
if (-not (Test-Path -LiteralPath $hyperPython)) {
    python -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw 'Could not create the Python environment.' }
}
if (-not (Test-Path -LiteralPath $hyperReady)) {
    & $hyperPython -m pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) { throw 'Could not install Hyper dependencies.' }
    New-Item -ItemType File -Path $hyperReady -Force | Out-Null
}
& $hyperPython HyperWeb.py
if ($LASTEXITCODE -ne 0) { throw 'Hyper closed with an error.' }
