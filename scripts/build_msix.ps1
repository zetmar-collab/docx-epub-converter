# Builds dist/DocxEpubConverter-<version>-x64.msix for Microsoft Store submission.
#
# Requirements:
#   - Python 3.10+ with the project's requirements installed
#   - pyinstaller           (pip install pyinstaller)
#   - Windows 10/11 SDK     (provides makeappx.exe / signtool.exe)
#
# The Store re-signs the package with the publisher certificate, so no signing
# is done here. Use -SelfSign only to test-install the package locally.
#
#   .\scripts\build_msix.ps1
#   .\scripts\build_msix.ps1 -SelfSign      # local sideload test build

[CmdletBinding()]
param(
    [switch]$SelfSign,
    [switch]$SkipAssets
)

$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$PackagingDir = Join-Path $ProjectDir "packaging"
$DistDir = Join-Path $ProjectDir "dist"
$LayoutDir = Join-Path $ProjectDir "build\msix-layout"

Set-Location $ProjectDir

# --- interpreter: prefer the project venv, fall back to PATH -----------------
$VenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
if (Test-Path $VenvPython) {
    $Python = $VenvPython
    Write-Host "Using project venv: $Python" -ForegroundColor DarkGray
} else {
    $Python = (Get-Command python -ErrorAction Stop).Source
    Write-Host "Using system Python: $Python" -ForegroundColor DarkGray
}

# --- version: single source of truth is src/epub_converter/constants.py ------
$ConstantsPath = Join-Path $ProjectDir "src\epub_converter\constants.py"
$VersionMatch = Select-String -Path $ConstantsPath -Pattern 'APP_VERSION\s*=\s*"([^"]+)"'
if (-not $VersionMatch) { throw "APP_VERSION not found in $ConstantsPath" }
$Version = $VersionMatch.Matches[0].Groups[1].Value
$Parts = $Version.Split(".")
while ($Parts.Count -lt 3) { $Parts += "0" }
$MsixVersion = "$($Parts[0]).$($Parts[1]).$($Parts[2]).0"
Write-Host "Version: $Version  ->  MSIX $MsixVersion" -ForegroundColor Cyan

# --- locate makeappx.exe from the Windows SDK --------------------------------
function Find-SdkTool([string]$Name) {
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $roots = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\bin",
        "$env:ProgramFiles\Windows Kits\10\bin"
    ) | Where-Object { Test-Path $_ }
    $found = $roots |
        ForEach-Object { Get-ChildItem -Path $_ -Recurse -Filter $Name -ErrorAction SilentlyContinue } |
        Where-Object { $_.FullName -match '\\x64\\' } |
        Sort-Object FullName -Descending |
        Select-Object -First 1
    if (-not $found) { throw "$Name not found. Install the Windows 10/11 SDK (App Certification Kit / MSIX Packaging Tools)." }
    return $found.FullName
}
$MakeAppx = Find-SdkTool "makeappx.exe"
Write-Host "makeappx: $MakeAppx"

# --- 1. Store assets ---------------------------------------------------------
if (-not $SkipAssets) {
    Write-Host "`n[1/4] Generating Store assets..." -ForegroundColor Green
    & $Python (Join-Path $PackagingDir "generate_store_assets.py")
    if ($LASTEXITCODE -ne 0) { throw "Asset generation failed." }
} else {
    Write-Host "`n[1/4] Skipping asset generation." -ForegroundColor DarkGray
}

# --- 2. PyInstaller ----------------------------------------------------------
Write-Host "`n[2/4] Building executable with PyInstaller..." -ForegroundColor Green
& $Python -m PyInstaller (Join-Path $PackagingDir "DocxEpubConverter.spec") --noconfirm --clean --distpath $DistDir --workpath (Join-Path $ProjectDir "build\pyinstaller")
if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed. Run: $Python -m pip install pyinstaller" }

$AppDir = Join-Path $DistDir "DocxEpubConverter"
if (-not (Test-Path (Join-Path $AppDir "DocxEpubConverter.exe"))) {
    throw "Expected executable not found in $AppDir"
}

# --- 3. Assemble the MSIX layout --------------------------------------------
Write-Host "`n[3/4] Assembling MSIX layout..." -ForegroundColor Green
if (Test-Path $LayoutDir) { Remove-Item -Recurse -Force $LayoutDir }
New-Item -ItemType Directory -Path $LayoutDir -Force | Out-Null

Copy-Item -Path (Join-Path $AppDir "*") -Destination $LayoutDir -Recurse -Force
Copy-Item -Path (Join-Path $PackagingDir "Assets") -Destination $LayoutDir -Recurse -Force

# An MSIX is itself an OPC package, so payload paths may not use the structures
# OPC reserves: "[Content_Types].xml", "_rels" directories, or "[" / "]" in a
# name. python-docx ships its default template both zipped (default.docx, the
# one it actually loads) and exploded (default-docx-template/), and the exploded
# copy trips every one of those rules -> makeappx fails with 0x8007007b.
# Strip any such payload; nothing at runtime reads it.
$reserved = @()
$reserved += Get-ChildItem $LayoutDir -Recurse -Directory -Force |
    Where-Object { $_.Name -eq "_rels" -or $_.Name -match '[\[\]]' }
$reserved += Get-ChildItem $LayoutDir -Recurse -File -Force |
    Where-Object { $_.Name -match '[\[\]]' }
foreach ($item in $reserved) {
    if (Test-Path -LiteralPath $item.FullName) {
        Write-Host "  removing OPC-reserved payload: $($item.FullName.Substring($LayoutDir.Length + 1))" -ForegroundColor DarkYellow
        Remove-Item -LiteralPath $item.FullName -Recurse -Force
    }
}
# The exploded template directory is now empty of its OPC parts; drop it whole.
$explodedTemplate = Join-Path $LayoutDir "_internal\docx\templates\default-docx-template"
if (Test-Path -LiteralPath $explodedTemplate) {
    Remove-Item -LiteralPath $explodedTemplate -Recurse -Force
}
if (-not (Test-Path (Join-Path $LayoutDir "_internal\docx\templates\default.docx"))) {
    Write-Warning "python-docx default.docx template is missing from the layout."
}

# Stamp the manifest version without mutating the source file.
$ManifestSrc = Join-Path $PackagingDir "AppxManifest.xml"
[xml]$Manifest = Get-Content $ManifestSrc
$Manifest.Package.Identity.Version = $MsixVersion
$Manifest.Save((Join-Path $LayoutDir "AppxManifest.xml"))

if ($Manifest.Package.Identity.Name -match "REPLACE") {
    throw "AppxManifest.xml still contains REPLACE placeholders - fill in Partner Center identity first."
}

# --- 4. Pack -----------------------------------------------------------------
Write-Host "`n[4/4] Packing MSIX..." -ForegroundColor Green
if (-not (Test-Path $DistDir)) { New-Item -ItemType Directory -Path $DistDir | Out-Null }
$MsixPath = Join-Path $DistDir "DocxEpubConverter-$MsixVersion-x64.msix"
if (Test-Path $MsixPath) { Remove-Item -Force $MsixPath }

& $MakeAppx pack /d $LayoutDir /p $MsixPath /o
if ($LASTEXITCODE -ne 0) { throw "makeappx pack failed." }

if ($SelfSign) {
    Write-Host "`nSelf-signing for local sideload test..." -ForegroundColor Yellow
    $SignTool = Find-SdkTool "signtool.exe"
    $Subject = $Manifest.Package.Identity.Publisher
    $Cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $Subject } | Select-Object -First 1
    if (-not $Cert) {
        Write-Host "Creating self-signed test certificate for $Subject"
        $Cert = New-SelfSignedCertificate -Type Custom -Subject $Subject `
            -KeyUsage DigitalSignature -FriendlyName "DOCX EPUB Converter Test" `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3", "2.5.29.19={text}")
        Write-Host "Install this cert into Local Machine > Trusted People before sideloading." -ForegroundColor Yellow
    }
    & $SignTool sign /fd SHA256 /sha1 $Cert.Thumbprint $MsixPath
    if ($LASTEXITCODE -ne 0) { throw "signtool sign failed." }
}

Write-Host "`nMSIX: $MsixPath" -ForegroundColor Cyan
Write-Host "Size: $([math]::Round((Get-Item $MsixPath).Length / 1MB, 2)) MB"
Write-Host "`nNext: upload to Partner Center (Packages step). Do NOT sign for Store submission." -ForegroundColor Cyan
