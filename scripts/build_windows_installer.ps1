param(
    [string]$Python = "python",
    [switch]$SkipTests,
    [switch]$SkipAppBuild,
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

function Write-ReleaseManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ArtifactPaths,
        [Parameter(Mandatory = $true)]
        [string]$Version
    )

    $commit = (& git rev-parse HEAD 2>$null).Trim()
    if (-not $commit) {
        $commit = "unknown"
    }

    $manifest = [ordered]@{
        app = "ScholarRef"
        version = $Version
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

    $manifest |
        ConvertTo-Json -Depth 4 |
        Set-Content -Path (Join-Path $RepoRoot "dist\release-manifest.json") -Encoding utf8
}

if (-not $SkipAppBuild) {
    $buildArgs = @("-ExecutionPolicy", "Bypass", "-File", (Join-Path $PSScriptRoot "build_windows.ps1"), "-Python", $Python)
    if ($SkipTests) {
        $buildArgs += "-SkipTests"
    }
    if ($SkipDependencyInstall) {
        $buildArgs += "-SkipDependencyInstall"
    }
    Invoke-CheckedCommand -Description "Windows app build" -FilePath "powershell" -ArgumentList $buildArgs
}

$iscc = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if (-not $iscc) {
    $candidatePaths = @(
        (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe")
    ) | Where-Object { $_ }

    foreach ($candidate in $candidatePaths) {
        if (Test-Path $candidate) {
            $iscc = Get-Item $candidate
            break
        }
    }
}

if (-not $iscc) {
    throw "Inno Setup compiler not found. Install Inno Setup 6 and rerun."
}

$isccPath = $null
if ($iscc.PSObject.Properties["Source"]) {
    $isccPath = $iscc.Source
}
if (-not $isccPath -and $iscc.PSObject.Properties["FullName"]) {
    $isccPath = $iscc.FullName
}
if (-not $isccPath) {
    throw "Could not resolve the Inno Setup compiler path."
}

$version = (& $Python -c "import scholarref_runtime; print(scholarref_runtime.APP_VERSION)").Trim()
if (-not $version) {
    throw "Could not resolve ScholarRef version."
}

$env:SCHOLARREF_VERSION = $version
Invoke-CheckedCommand -Description "Inno Setup compile" -FilePath $isccPath -ArgumentList @((Join-Path $RepoRoot "installer\ScholarRef.iss"))

$InstallerPath = Join-Path $RepoRoot "dist\ScholarRef-setup-$version.exe"
if (-not (Test-Path $InstallerPath)) {
    throw "Expected installer was not created: $InstallerPath"
}

$zipPath = Join-Path $RepoRoot "dist\ScholarRef-windows-x64.zip"

if ($env:SCHOLARREF_SIGN_CERT_PATH -and $env:SCHOLARREF_SIGN_CERT_PASSWORD) {
    Write-Host "Signing installer..."
    $timestampUrl = $env:SCHOLARREF_SIGN_TIMESTAMP_URL
    if (-not $timestampUrl) {
        $timestampUrl = "http://timestamp.digicert.com"
    }
    Invoke-CheckedCommand -Description "installer signing" -FilePath "powershell" -ArgumentList @(
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        (Join-Path $PSScriptRoot "sign_windows_artifact.ps1"),
        "-Path",
        $InstallerPath,
        "-CertificatePath",
        $env:SCHOLARREF_SIGN_CERT_PATH,
        "-CertificatePassword",
        $env:SCHOLARREF_SIGN_CERT_PASSWORD,
        "-TimestampUrl",
        $timestampUrl
    )
}

$HashPath = Join-Path $RepoRoot "dist\SHA256SUMS.txt"
$lines = @()
if (Test-Path $zipPath) {
    $lines += Get-ArtifactHashLine -ArtifactPath $zipPath
}
$lines += Get-ArtifactHashLine -ArtifactPath $InstallerPath
Set-Content -Path $HashPath -Value $lines -Encoding ascii

if (Test-Path $zipPath) {
    Write-ReleaseManifest -ArtifactPaths @($zipPath, $InstallerPath) -Version $version
} else {
    Write-ReleaseManifest -ArtifactPaths @($InstallerPath) -Version $version
}

Write-Host "Installer complete: $InstallerPath"
Write-Host "Checksums written: $HashPath"
