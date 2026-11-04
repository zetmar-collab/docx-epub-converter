# Builds dist/docx-epub-converter-v2.1.zip for GitHub Releases (no venv, no tests cache).
$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent

# Single source of truth for the version: src/epub_converter/constants.py
$ConstantsPath = Join-Path $ProjectDir "src\epub_converter\constants.py"
$VersionMatch = Select-String -Path $ConstantsPath -Pattern 'APP_VERSION\s*=\s*"([^"]+)"'
if (-not $VersionMatch) { throw "APP_VERSION not found in $ConstantsPath" }
$Version = $VersionMatch.Matches[0].Groups[1].Value
$ZipName = "docx-epub-converter-v$Version.zip"
$Staging = Join-Path $env:TEMP "docx-epub-converter-v$Version"
$DistDir = Join-Path $ProjectDir "dist"
$ZipPath = Join-Path $DistDir $ZipName

if (Test-Path $Staging) { Remove-Item -Recurse -Force $Staging }
New-Item -ItemType Directory -Path $Staging | Out-Null
if (-not (Test-Path $DistDir)) { New-Item -ItemType Directory -Path $DistDir | Out-Null }

$items = @(
    "src",
    "assets",
    "run_converter.py",
    "README.txt",
    "requirements.txt",
    "START_EPUB_CONVERTER.bat",
    "START_EPUB_CONVERTER.command",
    "install_windows.ps1",
    "install_macos.sh"
)
foreach ($item in $items) {
    $src = Join-Path $ProjectDir $item
    if (Test-Path $src) {
        Copy-Item -Path $src -Destination (Join-Path $Staging $item) -Recurse -Force
    }
}

Get-ChildItem -Path $Staging -Recurse -Directory -Filter "__pycache__" | Remove-Item -Recurse -Force
if (Test-Path $ZipPath) { Remove-Item -Force $ZipPath }
Compress-Archive -Path (Join-Path $Staging "*") -DestinationPath $ZipPath -CompressionLevel Optimal
Remove-Item -Recurse -Force $Staging

Write-Host "Release ZIP: $ZipPath"
Write-Host "Size: $((Get-Item $ZipPath).Length) bytes"
