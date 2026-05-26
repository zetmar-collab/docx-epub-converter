@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Tworzenie srodowiska Python...
    python -m venv .venv
)
echo Instalowanie zaleznosci...
".venv\Scripts\python.exe" -m pip install -r requirements.txt --quiet
echo.
echo Uruchamianie EPUB Converter...
start "" ".venv\Scripts\pythonw.exe" "%CD%\run_converter.py"
