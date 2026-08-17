# Build ZIP and publish GitHub Release (requires: gh auth login)
$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent
$Gh = "C:\Program Files\GitHub CLI\gh.exe"
if (-not (Test-Path $Gh)) { $Gh = "gh" }

Set-Location $ProjectDir
& "$ProjectDir\scripts\build_release_zip.ps1"

$ConstantsPath = Join-Path $ProjectDir "src\epub_converter\constants.py"
$VersionMatch = Select-String -Path $ConstantsPath -Pattern 'APP_VERSION\s*=\s*"([^"]+)"'
if (-not $VersionMatch) { throw "APP_VERSION not found in $ConstantsPath" }
$Version = $VersionMatch.Matches[0].Groups[1].Value

$Zip = Join-Path $ProjectDir "dist\docx-epub-converter-v$Version.zip"
$Notes = Join-Path $ProjectDir "docs\RELEASE_v$Version.md"

& $Gh auth status
& $gh release create "v$Version" `
    --repo zetmar-collab/docx-epub-converter `
    --title "DOCX EPUB Converter v$Version" `
    --notes-file $Notes `
    $Zip

Write-Host "Release: https://github.com/zetmar-collab/docx-epub-converter/releases/tag/v$Version"
