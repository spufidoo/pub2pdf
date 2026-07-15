<#
.SYNOPSIS
    Build pub2pdf.zip for GitHub release upload.

.DESCRIPTION
    Maintainer-only. Packages the teacher-facing files into a flat pub2pdf.zip.
    End users download that ZIP from the GitHub Releases page — they never run this script.

    Requires: PowerShell 5.1+, and gh CLI if you use -Upload.

.EXAMPLE
    .\scripts\build-release.ps1

.EXAMPLE
    .\scripts\build-release.ps1 -Tag v1.0.0 -Upload
#>
[CmdletBinding()]
param(
    # Where to write the ZIP. Defaults to pub2pdf.zip in the repo root.
    [string]$OutputPath,

    # Existing release tag, e.g. v1.0.0. Required with -Upload.
    [string]$Tag,

    # Upload the ZIP to GitHub Releases (replaces an existing pub2pdf.zip asset).
    [switch]$Upload,

    [string]$Repo = "spufidoo/pub2pdf"
)

$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot

$ReleaseFiles = @(
    "pub2pdf.ps1",
    "pub2pdf-gui.ps1",
    "pub2pdf.ico",
    "Convert Publisher to PDF.bat",
    "Convert my Publisher files.bat",
    "README.md"
)

foreach ($name in $ReleaseFiles) {
    $path = Join-Path $RepoRoot $name
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Missing release file: $path"
    }
}

if (-not $OutputPath) {
    $OutputPath = Join-Path $RepoRoot "pub2pdf.zip"
}

$staging = Join-Path $env:TEMP ("pub2pdf-release-{0}" -f [Guid]::NewGuid().ToString("N"))

try {
    New-Item -ItemType Directory -Path $staging -Force | Out-Null

    foreach ($name in $ReleaseFiles) {
        Copy-Item -LiteralPath (Join-Path $RepoRoot $name) -Destination $staging
    }

    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }

    Compress-Archive -Path (Join-Path $staging "*") -DestinationPath $OutputPath -Force

    $sizeKb = [math]::Round((Get-Item -LiteralPath $OutputPath).Length / 1KB, 1)
    Write-Host "Built $OutputPath ($sizeKb KB)"
    Write-Host "Contents:"
    foreach ($name in $ReleaseFiles) {
        Write-Host "  - $name"
    }
}
finally {
    if (Test-Path -LiteralPath $staging) {
        Remove-Item -LiteralPath $staging -Recurse -Force
    }
}

if ($Upload) {
    if (-not $Tag) {
        throw "Specify -Tag (e.g. v1.0.0) when using -Upload."
    }

    $gh = Get-Command gh -ErrorAction SilentlyContinue
    if (-not $gh) {
        throw "gh CLI not found. Install GitHub CLI or upload $OutputPath manually."
    }

    Write-Host "Uploading to $Repo release $Tag ..."
    & gh release upload $Tag $OutputPath --repo $Repo --clobber
    Write-Host "Done. Check: https://github.com/$Repo/releases/tag/$Tag"
}
