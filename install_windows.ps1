$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
$VenvPythonw = Join-Path $ProjectDir ".venv\Scripts\pythonw.exe"
$AppFile = Join-Path $ProjectDir "epub_converter.py"
$IconFile = Join-Path $ProjectDir "assets\app_icon.ico"

Set-Location $ProjectDir

if (-not (Test-Path $VenvPython)) {
    Write-Host "Tworzenie srodowiska Python..."
    python -m venv .venv
}

Write-Host "Instalowanie zaleznosci..."
& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r requirements.txt

$Desktop = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $Desktop "DOCX EPUB Converter.lnk"
$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $VenvPythonw
$Shortcut.Arguments = '"' + $AppFile + '"'
$Shortcut.WorkingDirectory = $ProjectDir
if (Test-Path $IconFile) {
    $Shortcut.IconLocation = $IconFile
}
$Shortcut.Description = "DOCX EPUB Converter - Autor: Marek Zettel"
$Shortcut.Save()

Write-Host "Gotowe. Skrot utworzony na pulpicie:"
Write-Host $ShortcutPath
