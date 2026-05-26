# Build ZIP and publish GitHub Release (requires: gh auth login)
$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent
$Gh = "C:\Program Files\GitHub CLI\gh.exe"
if (-not (Test-Path $Gh)) { $Gh = "gh" }

Set-Location $ProjectDir
& "$ProjectDir\scripts\build_release_zip.ps1"

$Zip = Join-Path $ProjectDir "dist\docx-epub-converter-v2.1.zip"
$Notes = Join-Path $ProjectDir "docs\RELEASE_v2.1.md"

& $Gh auth status
& $gh release create v2.1 `
    --repo zetmar-collab/docx-epub-converter `
    --title "DOCX EPUB Converter v2.1" `
    --notes-file $Notes `
    $Zip

Write-Host "Release: https://github.com/zetmar-collab/docx-epub-converter/releases/tag/v2.1"
