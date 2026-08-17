# DOCX EPUB Converter

[![Microsoft Store](https://img.shields.io/badge/Microsoft%20Store-pobierz-0078D4?logo=microsoftstore&logoColor=white)](https://apps.microsoft.com/detail/9NB4TQDS8M7Q)
[![Wersja](https://img.shields.io/badge/wersja-2.1.1-2f80d0)](https://github.com/zetmar-collab/docx-epub-converter/releases)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)

Desktopowa konwersja plików Word (`.docx`) do ebooków **EPUB 3** zgodnych
z Europejskim Aktem o Dostępności (EAA), EPUB Accessibility 1.1 i WCAG 2.1 AA.

Wszystko działa offline. Maszynopis nigdy nie opuszcza Twojego komputera —
bez konta, bez telemetrii, bez połączeń sieciowych.

**[➜ Zainstaluj z Microsoft Store](https://apps.microsoft.com/detail/9NB4TQDS8M7Q)**

---

## Instalacja

### Microsoft Store (zalecane)

[apps.microsoft.com/detail/9NB4TQDS8M7Q](https://apps.microsoft.com/detail/9NB4TQDS8M7Q) —
instalacja jednym kliknięciem, automatyczne aktualizacje, bez konfigurowania
Pythona.

### Z kodu źródłowego

Wymagany Python 3.10+.

**Windows**

```powershell
powershell -ExecutionPolicy Bypass -File .\install_windows.ps1
```

Instalator tworzy środowisko `.venv`, instaluje zależności i dodaje skrót na
pulpicie. Szybkie uruchomienie bez skrótu: `START_EPUB_CONVERTER.bat`.

**macOS**

```bash
chmod +x install_macos.sh START_EPUB_CONVERTER.command && ./install_macos.sh
```

Walidacja EpubCheck wymaga środowiska Java. Bez niej konwersja działa normalnie,
tylko bez etapu sprawdzania poprawności.

---

## Funkcje

**Konwersja**

- Nagłówki H1–H3 → struktura rozdziałów i nawigacyjny spis treści
- Pogrubienie, kursywa, podkreślenie, przekreślenie
- Listy wypunktowane i numerowane, cytaty blokowe
- Tabele (`w:tbl`) → poprawne tabele HTML
- Hiperłącza (`w:hyperlink`) → `<a href="...">`
- Obrazy inline (`a:blip`) wyodrębniane z DOCX i osadzane w EPUB
- Przypisy dolne i końcowe → `<sup>` z sekcją przypisów w rozdziale

**Dostępność**

- Semantyczne `epub:type` (chapter, frontmatter, backmatter)
- Teksty alternatywne obrazów (z `descr`/`title` lub domyślne dla języka)
- Metadane schema.org: `accessMode`, `accessibilityFeature`, `conformsTo`
- Zgodność z EPUB Accessibility 1.1 i WCAG 2.1 AA

**Praca z plikami**

- Kontrola wstępna DOCX — lista problemów struktury przed konwersją
- Podgląd okładki z ostrzeżeniem o proporcjach (2:3) i rozdzielczości (1200 px)
- Konwersja seryjna dowolnej liczby plików ze statusem per plik
- Interaktywny spis treści i pełny podgląd HTML w przeglądarce
- Walidacja EpubCheck po każdej konwersji

**Języki**

- Interfejs: polski, angielski
- Język dokumentu: 13 opcji (PL, EN, DE, FR, ES, IT, CS, SK, HU, RU, UK, NL, PT)

---

## Przygotowanie dokumentu DOCX

| Element | Styl w Wordzie |
|---|---|
| Rozdział | `Heading 1` / `Nagłówek 1` |
| Podrozdział | `Heading 2` / `Nagłówek 2`, `Heading 3` / `Nagłówek 3` |
| Lista | `List Bullet` / `List Number` |
| Cytat | `Quote` / `Cytat` |

Okładka: PNG lub JPG, proporcje 2:3 (np. 1600 × 2400 px), minimum 1200 px
szerokości.

---

## Rozwój

```bash
python -m venv .venv
.venv/Scripts/python.exe -m pip install -r requirements-dev.txt
.venv/Scripts/python.exe -m pytest tests -q
```

Budowanie pakietu MSIX dla Microsoft Store (wymaga PyInstallera i Windows SDK):

```powershell
.\scripts\build_msix.ps1
```

Wersja jest czytana z `APP_VERSION` w `src/epub_converter/constants.py` przez
wszystkie skrypty budujące — to jedyne miejsce, w którym się ją zmienia.

Struktura:

```
src/epub_converter/   kod aplikacji (parser, builder EPUB, GUI)
tests/                testy pytest
packaging/            manifest MSIX, ikony Store, polityka prywatności
scripts/              budowanie ZIP-a, MSIX-a i wydań GitHub
docs/                 notatki wydań
```

---

## Zależności

| Pakiet | Rola |
|---|---|
| `python-docx` | parsowanie plików DOCX |
| `Pillow` | obsługa obrazów i okładki |
| `epubcheck` | walidacja EPUB (wymaga Javy) |
| `tkinterdnd2` | opcjonalne — drag & drop |

---

## Historia zmian

Pełna lista w [`docs/`](docs/). Najnowsze wydania:
[v2.1.1](docs/RELEASE_v2.1.1.md) · [v2.1](docs/RELEASE_v2.1.md)

---

Autor: **Marek Zettel** · [Microsoft Store](https://apps.microsoft.com/detail/9NB4TQDS8M7Q)
