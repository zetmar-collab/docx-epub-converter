"""UI translations (Polish and English)."""

from epub_converter.constants import APP_AUTHOR

_EXTRA = {
    "pl": {
        "btn_preflight_continue": "Konwertuj mimo ostrzezen",
        "btn_preflight_cancel": "Anuluj",
        "java_warn_title": "Brak Java",
        "java_warn_msg": (
            "EpubCheck wymaga Java Runtime (JRE). Zainstaluj Java z https://adoptium.net/ "
            "— walidacja EPUB bedzie niedostepna do czasu instalacji."
        ),
        "err_isbn_checksum": "Suma kontrolna ISBN jest niepoprawna.",
        "epubcheck_skipped": "EpubCheck: pominiety (brak Java)",
        "batch_epubcheck_warn": " (EpubCheck: ostrzezenia)",
        "err_invalid_dnd": "Nieprawidlowy plik — przeciagnij istniejacy plik DOCX lub obraz PNG/JPG.",
    },
    "en": {
        "btn_preflight_continue": "Convert anyway",
        "btn_preflight_cancel": "Cancel",
        "java_warn_title": "Java not found",
        "java_warn_msg": (
            "EpubCheck requires a Java Runtime (JRE). Install Java from https://adoptium.net/ "
            "— EPUB validation will be unavailable until then."
        ),
        "err_isbn_checksum": "ISBN checksum is invalid.",
        "epubcheck_skipped": "EpubCheck: skipped (Java not installed)",
        "batch_epubcheck_warn": " (EpubCheck warnings)",
        "err_invalid_dnd": "Invalid file — drop an existing DOCX or PNG/JPG image.",
    },
}

TRANSLATIONS: dict[str, dict[str, str]] = {
    "pl": {
        "app_subtitle": "Desktopowa konwersja do EPUB 3 zgodnego z EAA / EPUB Accessibility 1.1 / WCAG 2.1 AA",
        "app_author_label": "Autor: " + APP_AUTHOR,
        "status_ready": "Gotowe. Wybierz pliki i uzupelnij metadane.",
        "btn_convert": "Konwertuj i zapisz EPUB",
        "section_metadata": "1. Metadane ebooka",
        "field_title": "Tytul *",
        "field_author": "Autor *",
        "field_publisher": "Wydawca",
        "field_year": "Rok wydania",
        "field_isbn": "ISBN (opcjonalnie)",
        "field_desc": "Krotki opis",
        "section_lang_doc": "Jezyk dokumentu",
        "section_files": "2. Pliki",
        "field_docx": "Plik DOCX *",
        "field_cover": "Okladka *",
        "no_cover_preview": "Brak podgladu okladki",
        "cover_load_error": "Nie mozna wczytac okladki: ",
        "section_save": "3. Zapis",
        "field_output": "Plik EPUB (sciezka zapisu) *",
        "btn_pick": "Wybierz",
        "btn_open_epub": "Otworz EPUB",
        "btn_open_folder": "Pokaz folder",
        "section_preview": "Podglad skonwertowanego pliku",
        "preview_after": "Po konwersji zobaczysz tutaj wynik walidacji i zapisany plik.",
        "toc_label": "Spis tresci",
        "preview_html_title": "Podglad EPUB",
        "author_label": "Autor",
        "err_isbn_title": "Niepoprawny ISBN",
        "err_isbn_msg": "ISBN musi zawierac 10 lub 13 cyfr (bez myslnikow) i poprawna sume kontrolna.",
        "isbn_help_title": "Co to jest ISBN i jak go uzyskac?",
        "isbn_help_text": (
            "ISBN (International Standard Book Number) to bezplatny, unikatowy numer\n"
            "identyfikujacy publikacje. Kazdy format (e-book, druk, audiobook) wymaga\n"
            "osobnego numeru.\n\n"
            "Format w tej aplikacji: 10 lub 13 cyfr, bez myslnikow ani spacji.\n"
            "Przyklad: 9788301234567\n\n"
            "Jak uzyskac ISBN w Polsce (bezplatnie):\n"
            "  1. Wejdz na e-isbn.pl (serwis Biblioteki Narodowej).\n"
            "  2. Zaloz konto wydawcy — takze jako osoba fizyczna.\n"
            "  3. Zloz wniosek o pule numerow (min. 10 sztuk).\n"
            "  4. Numery przyznawane w kilka dni roboczych.\n\n"
            "Pole jest opcjonalne — EPUB powstanie rowniez bez numeru ISBN."
        ),
        "preview_placeholder": (
            "Podglad pojawi sie po konwersji.\n\n"
            "W aplikacji widoczny jest tekst rozdzialow. "
            "Pelny podglad HTML mozna otworzyc przyciskiem ponizej."
        ),
        "btn_html_preview": "Otworz pelny podglad HTML",
        "converting": "Konwertuje...",
        "status_done": "Gotowe. Plik zapisany: ",
        "status_error": "Blad konwersji.",
        "dlg_pick_docx": "Wybierz plik DOCX",
        "dlg_pick_cover": "Wybierz okladke",
        "dlg_pick_output": "Zapisz plik EPUB",
        "warn_epubcheck": "EPUB zostal zbudowany, ale EpubCheck zglosil problemy:\n\n",
        "info_saved": "EPUB zostal zapisany:\n",
        "btn_ok": "OK",
        "err_missing_title": "Brak wymaganych danych",
        "err_missing_msg": "Uzupelnij wszystkie wymagane pola przed konwersja:\n\n- ",
        "err_year_title": "Niepoprawne metadane",
        "err_year_msg": "Rok wydania powinien miec format czterocyfrowy, np. 2026.",
        "err_cover_format_title": "Niepoprawna okladka",
        "err_cover_format_msg": "Okladka musi byc poprawnym plikiem PNG albo JPG/JPEG.",
        "err_cover_open_title": "Niepoprawna okladka",
        "err_cover_open_msg": "Nie mozna otworzyc pliku okladki:\n\n",
        "missing_title": "tytul",
        "missing_author": "autor",
        "missing_publisher": "wydawca",
        "missing_year": "rok wydania",
        "missing_desc": "krotki opis",
        "missing_docx": "plik DOCX",
        "missing_cover": "okladka",
        "missing_output": "sciezka zapisu EPUB",
        "preflight_title": "DOCX wymaga poprawek przed konwersja",
        "preflight_msg": "Dokument nie nadaje sie jeszcze do bezpiecznej konwersji EPUB.\n\n",
        "preflight_err_title": "Nie mozna sprawdzic DOCX",
        "preflight_err_msg": "Blad odczytu dokumentu:\n\n",
        "err_no_chapters": "Nie znaleziono rozdzialow. Uzyj stylu Heading 1 / Naglowek 1 dla tytulow rozdzialow.",
        "issue_empty_h1": "Jeden z naglowkow Heading 1 / Naglowek 1 jest pusty.",
        "issue_h3_without_h2": "Dokument zawiera naglowek Heading 3 / Naglowek 3 bez poprzedzajacego Heading 2.",
        "issue_no_chapters": "Brakuje rozdzialow oznaczonych stylem Heading 1 / Naglowek 1. EPUB wymaga poprawnej struktury rozdzialow.",
        "issue_text_before_h1": (
            "Przed pierwszym naglowkiem Heading 1 / Naglowek 1 znajduje sie tekst. "
            "Konwerter zaczyna EPUB od pierwszego rozdzialu, wiec ten tekst moglby zostac pominiety."
        ),
        "issue_empty_chapters": "Te rozdzialy nie maja tresci pod naglowkiem: ",
        "btn_copy": "Kopiuj do schowka",
        "btn_batch": "Konwertuj serie...",
        "batch_title": "Konwersja serii DOCX",
        "batch_select": "Wybierz pliki DOCX do konwersji seryjnej",
        "batch_need_author": "Uzupelnij pola Autor i Wydawca przed konwersja seryjna.",
        "batch_need_cover": "Wybierz okladke przed konwersja seryjna.",
        "batch_summary_ok": "Sukces: ",
        "batch_summary_err": ", Bledy: ",
        "cover_warn_ratio": "Uwaga: proporcje okladki odbiegaja od standardu 2:3 (sklepy ebookowe).",
        "cover_warn_res": "Uwaga: okladka moze byc za mala — zalecane minimum 1200 px szerokosci.",
        "epubcheck_messages": "komunikat(ow)",
        "label_file": "Plik:",
        "label_error": "Blad:",
    },
    "en": {
        "app_subtitle": "Desktop conversion to EPUB 3 compliant with EAA / EPUB Accessibility 1.1 / WCAG 2.1 AA",
        "app_author_label": "Author: " + APP_AUTHOR,
        "status_ready": "Ready. Select files and fill in metadata.",
        "btn_convert": "Convert and Save EPUB",
        "section_metadata": "1. eBook Metadata",
        "field_title": "Title *",
        "field_author": "Author *",
        "field_publisher": "Publisher",
        "field_year": "Year",
        "field_isbn": "ISBN (optional)",
        "field_desc": "Short description",
        "section_lang_doc": "Document language",
        "section_files": "2. Files",
        "field_docx": "DOCX file *",
        "field_cover": "Cover image *",
        "no_cover_preview": "No cover preview",
        "cover_load_error": "Cannot load cover: ",
        "section_save": "3. Save",
        "field_output": "EPUB file (save path) *",
        "btn_pick": "Browse",
        "btn_open_epub": "Open EPUB",
        "btn_open_folder": "Show folder",
        "section_preview": "Preview of converted file",
        "preview_after": "After conversion you will see the validation result and saved file here.",
        "toc_label": "Table of contents",
        "preview_html_title": "EPUB Preview",
        "author_label": "Author",
        "err_isbn_title": "Invalid ISBN",
        "err_isbn_msg": "ISBN must be 10 or 13 digits (no hyphens) with a valid checksum.",
        "isbn_help_title": "What is ISBN and how to get one?",
        "isbn_help_text": (
            "ISBN (International Standard Book Number) is a free, unique identifier\n"
            "for publications. Each format (e-book, print, audiobook) needs its own number.\n\n"
            "Format in this app: 10 or 13 digits, no hyphens or spaces.\n"
            "Example: 9788301234567\n\n"
            "How to get an ISBN in Poland (free of charge):\n"
            "  1. Go to e-isbn.pl (National Library of Poland).\n"
            "  2. Create a publisher account — individuals are welcome.\n"
            "  3. Apply for a block of numbers (minimum 10).\n\n"
            "This field is optional — EPUB will be created without an ISBN too."
        ),
        "preview_placeholder": (
            "Preview will appear after conversion.\n\n"
            "The application shows chapter text. "
            "Full HTML preview can be opened with the button below."
        ),
        "btn_html_preview": "Open full HTML preview",
        "converting": "Converting...",
        "status_done": "Done. File saved: ",
        "status_error": "Conversion error.",
        "dlg_pick_docx": "Select DOCX file",
        "dlg_pick_cover": "Select cover image",
        "dlg_pick_output": "Save EPUB file",
        "warn_epubcheck": "EPUB was built, but EpubCheck reported issues:\n\n",
        "info_saved": "EPUB has been saved:\n",
        "btn_ok": "OK",
        "err_missing_title": "Missing required data",
        "err_missing_msg": "Please fill in all required fields before conversion:\n\n- ",
        "err_year_title": "Invalid metadata",
        "err_year_msg": "Year must be a 4-digit number, e.g. 2026.",
        "err_cover_format_title": "Invalid cover image",
        "err_cover_format_msg": "Cover must be a valid PNG or JPG/JPEG file.",
        "err_cover_open_title": "Invalid cover image",
        "err_cover_open_msg": "Cannot open cover file:\n\n",
        "missing_title": "title",
        "missing_author": "author",
        "missing_publisher": "publisher",
        "missing_year": "year",
        "missing_desc": "short description",
        "missing_docx": "DOCX file",
        "missing_cover": "cover image",
        "missing_output": "EPUB save path",
        "preflight_title": "DOCX needs corrections before conversion",
        "preflight_msg": "The document is not yet suitable for safe EPUB conversion.\n\n",
        "preflight_err_title": "Cannot check DOCX",
        "preflight_err_msg": "Document read error:\n\n",
        "err_no_chapters": "No chapters found. Use Heading 1 / Nagłówek 1 style for chapter titles.",
        "issue_empty_h1": "One of the Heading 1 / Nagłówek 1 headings is empty.",
        "issue_h3_without_h2": "The document contains a Heading 3 / Nagłówek 3 without a preceding Heading 2.",
        "issue_no_chapters": "No chapters marked with Heading 1 / Nagłówek 1 style found. EPUB requires proper chapter structure.",
        "issue_text_before_h1": (
            "There is text before the first Heading 1 / Nagłówek 1. "
            "The converter starts the EPUB from the first chapter, so this text might be skipped."
        ),
        "issue_empty_chapters": "These chapters have no body text under their heading: ",
        "btn_copy": "Copy to clipboard",
        "btn_batch": "Batch convert...",
        "batch_title": "Batch DOCX Conversion",
        "batch_select": "Select DOCX files for batch conversion",
        "batch_need_author": "Fill in Author and Publisher fields before batch conversion.",
        "batch_need_cover": "Select a cover image before batch conversion.",
        "batch_summary_ok": "Success: ",
        "batch_summary_err": ", Errors: ",
        "cover_warn_ratio": "Warning: cover proportions differ from the 2:3 standard (required by most stores).",
        "cover_warn_res": "Warning: cover may be too small — recommended minimum width is 1200 px.",
        "epubcheck_messages": "message(s)",
        "label_file": "File:",
        "label_error": "Error:",
    },
}

for _lang, _extra in _EXTRA.items():
    TRANSLATIONS[_lang].update(_extra)
