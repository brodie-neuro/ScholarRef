param(
    [string]$Python = "python",
    [string]$VenvDir = ".venv-release-smoke"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

function Invoke-CheckedCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Description,
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [string[]]$ArgumentList = @()
    )

    & $FilePath @ArgumentList
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $LASTEXITCODE."
    }
}

$venvPath = Join-Path $RepoRoot $VenvDir
if (Test-Path $venvPath) {
    Remove-Item $venvPath -Recurse -Force
}

Write-Host "Creating clean virtual environment at $venvPath"
Invoke-CheckedCommand -Description "virtual environment creation" -FilePath $Python -ArgumentList @("-m", "venv", $venvPath)

$venvPython = Join-Path $venvPath "Scripts\python.exe"
if (-not (Test-Path $venvPython)) {
    throw "Virtual environment python was not created: $venvPython"
}

Write-Host "Installing ScholarRef into clean virtual environment"
Invoke-CheckedCommand -Description "venv pip upgrade" -FilePath $venvPython -ArgumentList @("-m", "pip", "install", "--disable-pip-version-check", "--no-input", "--upgrade", "pip")
Invoke-CheckedCommand -Description "venv ScholarRef install" -FilePath $venvPython -ArgumentList @("-m", "pip", "install", "--disable-pip-version-check", "--no-input", ".[build]")

Write-Host "Running clean-environment validation"
Invoke-CheckedCommand -Description "venv pytest" -FilePath $venvPython -ArgumentList @("-m", "pytest", "-q")
Invoke-CheckedCommand -Description "venv GUI smoke test" -FilePath $venvPython -ArgumentList @("scholarref_gui.py", "--smoke-test")

Write-Host "Clean virtual environment validation passed"
