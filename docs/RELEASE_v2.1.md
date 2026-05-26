# DOCX EPUB Converter v2.1

Desktopowa konwersja DOCX → EPUB 3 (EAA / EPUB Accessibility 1.1 / WCAG 2.1 AA).

## Wymagania

- Python 3.10+
- Java (JRE) — opcjonalnie, do walidacji EpubCheck

## Instalacja (Windows)

1. Rozpakuj archiwum ZIP.
2. Uruchom `install_windows.ps1` (PowerShell) **lub** dwuklik `START_EPUB_CONVERTER.bat`.
3. Skrót na pulpicie: „DOCX EPUB Converter”.

## Instalacja (macOS)

```bash
chmod +x install_macos.sh START_EPUB_CONVERTER.command
./install_macos.sh
```

## Uruchomienie

- Windows: `START_EPUB_CONVERTER.bat` lub `run_converter.py`
- macOS/Linux: `python3 run_converter.py`

## Zmiany w v2.1

- Modułowa architektura (`src/epub_converter/`)
- Poprawka rozpoznawania stylów Heading (np. Heading 10 ≠ Heading 1)
- ISBN: suma kontrolna + `urn:isbn:` w metadanych OPUB
- Teksty alternatywne obrazów z DOCX lub domyślne w języku dokumentu
- Preflight: „Konwertuj mimo ostrzeżeń”
- Sprawdzenie Java przy starcie; batch z EpubCheck

Autor: Marek Zettel
