# DOCX EPUB Converter v2.1.1

Poprawka walidacji EpubCheck przy wgrywaniu EPUB do sklepów i czytników.

## Naprawione

- **cover_page.xhtml:** usunięto nieprawidłowe `role="doc-cover"` (EpubCheck ERROR). Zostaje `epub:type="cover"`.
- **Przypisy:** `role="doc-footnote"` zamienione na dozwolone `role="note"`.

## Instalacja

Jak w v2.1 — rozpakuj ZIP, uruchom `install_windows.ps1` lub `START_EPUB_CONVERTER.bat`.

Autor: Marek Zettel
