#!/bin/zsh
set -e

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$PROJECT_DIR"

if [ ! -x ".venv/bin/python" ]; then
    echo "Tworzenie srodowiska Python..."
    python3 -m venv .venv
fi

echo "Instalowanie zaleznosci..."
".venv/bin/python" -m pip install -r requirements.txt --quiet

echo "Uruchamianie EPUB Converter..."
".venv/bin/python" run_converter.py
