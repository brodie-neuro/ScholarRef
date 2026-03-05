param(
    [Parameter(Mandatory = $true)]
    [string[]]$Path,
    [Parameter(Mandatory = $true)]
    [string]$CertificatePath,
    [Parameter(Mandatory = $true)]
    [string]$CertificatePassword,
    [string]$TimestampUrl = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Get-SignToolPath {
    $candidates = @(
        (Get-Command signtool.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue)
    ) | Where-Object { $_ }

    if (-not $candidates) {
        $kitsRoot = Join-Path ${env:ProgramFiles(x86)} "Windows Kits\10\bin"
        if (Test-Path $kitsRoot) {
            $candidates = Get-ChildItem -Path $kitsRoot -Filter signtool.exe -Recurse -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending |
                Select-Object -ExpandProperty FullName
        }
    }

    if (-not $candidates) {
        throw "signtool.exe was not found. Install the Windows SDK signing tools."
    }

    return $candidates[0]
}

$signTool = Get-SignToolPath

if (-not (Test-Path $CertificatePath)) {
    throw "Signing certificate not found: $CertificatePath"
}

foreach ($item in $Path) {
    if (-not (Test-Path $item)) {
        throw "Artifact to sign was not found: $item"
    }

    Write-Host "Signing $item"
    & $signTool sign /fd SHA256 /f $CertificatePath /p $CertificatePassword /tr $TimestampUrl /td SHA256 $item
    if ($LASTEXITCODE -ne 0) {
        throw "signtool failed for $item with exit code $LASTEXITCODE"
    }

    $signature = Get-AuthenticodeSignature -FilePath $item
    if ($signature.Status -eq "NotSigned" -or $null -eq $signature.SignerCertificate) {
        throw "Authenticode signature check failed for $item: $($signature.Status)"
    }
}
