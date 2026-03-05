param(
    [string]$InstallerPath,
    [string]$InstallDir = (Join-Path $env:TEMP ("ScholarRef-InstallerSmoke-" + [guid]::NewGuid().ToString("N")))
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

function Invoke-CheckedProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Description,
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [string[]]$ArgumentList = @()
    )

    $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -Wait
    if ($process.ExitCode -ne 0) {
        throw "$Description failed with exit code $($process.ExitCode)."
    }
}

function Wait-ForRemoval {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathToCheck,
        [int]$TimeoutSeconds = 20
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Test-Path $PathToCheck) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 1
    }
}

if (-not $InstallerPath) {
    $InstallerPath = Get-ChildItem -Path (Join-Path $RepoRoot "dist") -Filter "ScholarRef-setup-*.exe" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1 -ExpandProperty FullName
}

if (-not $InstallerPath -or -not (Test-Path $InstallerPath)) {
    throw "Installer not found. Build the installer before running this smoke test."
}

$installRoot = Split-Path -Parent $InstallDir
if (-not (Test-Path $installRoot)) {
    New-Item -ItemType Directory -Path $installRoot | Out-Null
}

$installLog = Join-Path $env:TEMP ("ScholarRef-install-" + [guid]::NewGuid().ToString("N") + ".log")
$smokeStdout = Join-Path $env:TEMP ("ScholarRef-smoke-" + [guid]::NewGuid().ToString("N") + ".out.log")
$smokeStderr = Join-Path $env:TEMP ("ScholarRef-smoke-" + [guid]::NewGuid().ToString("N") + ".err.log")
$uninstallLog = Join-Path $env:TEMP ("ScholarRef-uninstall-" + [guid]::NewGuid().ToString("N") + ".log")

Write-Host "Running silent installer smoke test from $InstallerPath"
Invoke-CheckedProcess -Description "silent installer execution" -FilePath $InstallerPath -ArgumentList @(
    "/VERYSILENT",
    "/SUPPRESSMSGBOXES",
    "/NORESTART",
    "/CURRENTUSER",
    "/DIR=$InstallDir",
    "/LOG=$installLog"
)

$exePath = Join-Path $InstallDir "ScholarRef.exe"
if (-not (Test-Path $exePath)) {
    throw "Installed executable not found: $exePath"
}

Write-Host "Running installed executable smoke test"
if (Test-Path $smokeStdout) {
    Remove-Item $smokeStdout -Force
}
if (Test-Path $smokeStderr) {
    Remove-Item $smokeStderr -Force
}
$smokeProcess = Start-Process -FilePath $exePath -ArgumentList @("--smoke-test") -PassThru -Wait -RedirectStandardOutput $smokeStdout -RedirectStandardError $smokeStderr
if (Test-Path $smokeStdout) {
    Get-Content $smokeStdout
}
if (Test-Path $smokeStderr) {
    Get-Content $smokeStderr
}
if ($smokeProcess.ExitCode -ne 0) {
    throw "Installed executable smoke test failed with exit code $($smokeProcess.ExitCode)."
}

$uninstaller = Join-Path $InstallDir "unins000.exe"
if (Test-Path $uninstaller) {
    Write-Host "Running silent uninstall smoke test"
    Invoke-CheckedProcess -Description "silent uninstall" -FilePath $uninstaller -ArgumentList @(
        "/VERYSILENT",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
        "/LOG=$uninstallLog"
    )
    Wait-ForRemoval -PathToCheck $InstallDir
    if (Test-Path $InstallDir) {
        throw "Installer uninstall finished but the install directory still exists: $InstallDir`nInstall log: $installLog`nUninstall log: $uninstallLog"
    }
}

Write-Host "Installer smoke test passed"
