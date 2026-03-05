param(
    [string]$Python = "python",
    [switch]$SkipTests,
    [switch]$SkipDependencyInstall
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

function Get-ArtifactHashLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArtifactPath
    )

    $hash = (Get-FileHash $ArtifactPath -Algorithm SHA256).Hash.ToLower()
    return "$hash  $(Split-Path -Leaf $ArtifactPath)"
}

function Get-ReleaseManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ArtifactPaths
    )

    $commit = (& git rev-parse HEAD 2>$null).Trim()
    if (-not $commit) {
        $commit = "unknown"
    }

    return [ordered]@{
        app = "ScholarRef"
        version = (& $Python -c "import scholarref_runtime; print(scholarref_runtime.APP_VERSION)").Trim()
        build_time_utc = (Get-Date).ToUniversalTime().ToString("o")
        commit = $commit
        signed = [bool]$env:SCHOLARREF_SIGN_CERT_PATH
        artifacts = @(
            foreach ($artifactPath in $ArtifactPaths) {
                [ordered]@{
                    name = Split-Path -Leaf $artifactPath
                    sha256 = (Get-FileHash $artifactPath -Algorithm SHA256).Hash.ToLower()
                }
            }
        )
    }
}

if (-not $SkipDependencyInstall) {
    Write-Host "Installing build dependencies..."
    Invoke-CheckedCommand -Description "pip upgrade" -FilePath $Python -ArgumentList @("-m", "pip", "install", "--disable-pip-version-check", "--no-input", "--upgrade", "pip")
    Invoke-CheckedCommand -Description "build dependency install" -FilePath $Python -ArgumentList @("-m", "pip", "install", "--disable-pip-version-check", "--no-input", ".[build]")
}

if (-not $SkipTests) {
    Write-Host "Running tests..."
    Invoke-CheckedCommand -Description "pytest" -FilePath $Python -ArgumentList @("-m", "pytest", "-q")
}

Write-Host "Building Windows app..."
Invoke-CheckedCommand -Description "PyInstaller build" -FilePath $Python -ArgumentList @("-m", "PyInstaller", "--noconfirm", "--clean", "ScholarRef.spec")

$ExePath = Join-Path $RepoRoot "dist\ScholarRef\ScholarRef.exe"
if (-not (Test-Path $ExePath)) {
    throw "Expected packaged executable was not created: $ExePath"
}

if ($env:SCHOLARREF_SIGN_CERT_PATH -and $env:SCHOLARREF_SIGN_CERT_PASSWORD) {
    Write-Host "Signing packaged executable..."
    $timestampUrl = $env:SCHOLARREF_SIGN_TIMESTAMP_URL
    if (-not $timestampUrl) {
        $timestampUrl = "http://timestamp.digicert.com"
    }
    Invoke-CheckedCommand -Description "code signing" -FilePath "powershell" -ArgumentList @(
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        (Join-Path $PSScriptRoot "sign_windows_artifact.ps1"),
        "-Path",
        $ExePath,
        "-CertificatePath",
        $env:SCHOLARREF_SIGN_CERT_PATH,
        "-CertificatePassword",
        $env:SCHOLARREF_SIGN_CERT_PASSWORD,
        "-TimestampUrl",
        $timestampUrl
    )
}

Write-Host "Running packaged smoke test..."
Invoke-CheckedCommand -Description "packaged smoke test" -FilePath $ExePath -ArgumentList @("--smoke-test")

$ZipPath = Join-Path $RepoRoot "dist\ScholarRef-windows-x64.zip"
if (Test-Path $ZipPath) {
    Remove-Item $ZipPath -Force
}

Write-Host "Creating release archive..."
Compress-Archive -Path (Join-Path $RepoRoot "dist\ScholarRef\*") -DestinationPath $ZipPath

@(
    Get-ArtifactHashLine -ArtifactPath $ZipPath
) | Set-Content -Path (Join-Path $RepoRoot "dist\SHA256SUMS.txt") -Encoding ascii

Get-ReleaseManifest -ArtifactPaths @($ZipPath) |
    ConvertTo-Json -Depth 4 |
    Set-Content -Path (Join-Path $RepoRoot "dist\release-manifest.json") -Encoding utf8

Write-Host "Build complete: $ZipPath"
