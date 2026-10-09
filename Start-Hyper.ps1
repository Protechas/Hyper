param(
    [switch]$ValidatePython
)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$hyperPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
$hyperReady = Join-Path $PSScriptRoot '.venv\hyper-ready'

function Test-HyperPython {
    param(
        [Parameter(Mandatory = $true)][string]$Command,
        [string[]]$PrefixArguments = @()
    )

    try {
        $testArguments = @($PrefixArguments) + @(
            '-c',
            'import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 11) else 3)'
        )
        & $Command @testArguments *> $null
        return ($LASTEXITCODE -eq 0)
    }
    catch {
        return $false
    }
}

$venvWorks = (Test-Path -LiteralPath $hyperPython) -and (Test-HyperPython -Command $hyperPython)
if (-not $venvWorks) {
    # Support both development machines with a standalone Python install and
    # managed workstations using Python 3.11 from the Microsoft Store.
    $candidates = @()
    $repositoryPython = Join-Path $PSScriptRoot 'env\Scripts\python.exe'
    if (Test-Path -LiteralPath $repositoryPython) {
        $candidates += [pscustomobject]@{
            Command = $repositoryPython
            PrefixArguments = @()
            Description = 'repository env'
        }
    }

    foreach ($candidateSpec in @(
        @{ Name = 'py.exe'; PrefixArguments = @('-3.11'); Description = 'Windows Python launcher' },
        @{ Name = 'python3.11.exe'; PrefixArguments = @(); Description = 'Python 3.11' },
        @{ Name = 'python.exe'; PrefixArguments = @(); Description = 'Python/Windows Store alias' },
        @{ Name = 'python3.exe'; PrefixArguments = @(); Description = 'Python 3 alias' }
    )) {
        $pythonCommands = @(Get-Command $candidateSpec.Name -CommandType Application -All -ErrorAction SilentlyContinue)
        foreach ($pythonCommand in $pythonCommands) {
            $candidates += [pscustomobject]@{
                Command = $pythonCommand.Source
                PrefixArguments = @($candidateSpec.PrefixArguments)
                Description = $candidateSpec.Description
            }
        }
    }

    $bootstrapPython = $null
    foreach ($candidate in $candidates) {
        if (Test-HyperPython -Command $candidate.Command -PrefixArguments $candidate.PrefixArguments) {
            $bootstrapPython = $candidate
            break
        }
    }

    if (-not $bootstrapPython) {
        throw 'Could not find Python 3.11. Install Python 3.11 from python.org or the Microsoft Store, then run Start-Hyper.cmd again.'
    }

    Write-Host "Creating Hyper environment with $($bootstrapPython.Description): $($bootstrapPython.Command)"
    $venvArguments = @($bootstrapPython.PrefixArguments) + @('-m', 'venv', '--clear', '.venv')
    & $bootstrapPython.Command @venvArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create the Python environment using $($bootstrapPython.Command)."
    }
    if (-not (Test-HyperPython -Command $hyperPython)) {
        throw 'The Python environment was created but could not be started.'
    }
}
if ($ValidatePython) {
    & $hyperPython -c "import sys; print('Hyper Python ready: {} ({})'.format(sys.executable, sys.version.split()[0]))"
    if ($LASTEXITCODE -ne 0) { throw 'The Hyper Python validation command failed.' }
    exit 0
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
