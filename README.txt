DOCX EPUB Converter v2.1
Autor: Marek Zettel
Konwersja plikow DOCX do EPUB 3 zgodnych z EAA / EPUB Accessibility 1.1 / WCAG 2.1 AA

========================================================================
INSTALACJA
========================================================================

Windows:
1. Kliknij prawym przyciskiem myszy install_windows.ps1 i wybierz
   "Run with PowerShell".
   Jesli system blokuje skrypt, uruchom w PowerShell:
     powershell -ExecutionPolicy Bypass -File .\install_windows.ps1
2. Instalator utworzy srodowisko .venv, zainstaluje zaleznosci i doda
   skrot "DOCX EPUB Converter" na pulpicie.
3. Szybkie uruchomienie bez instalacji skrotu: START_EPUB_CONVERTER.bat

macOS:
1. W Terminalu przejdz do folderu projektu.
2. Uruchom:
     chmod +x install_macos.sh START_EPUB_CONVERTER.command
     ./install_macos.sh
3. Instalator utworzy aplikacje "DOCX EPUB Converter.app" na pulpicie.
4. Szybkie uruchomienie bez instalacji: START_EPUB_CONVERTER.command

Wymagania:
- Python 3.10 lub nowszy
- Java (wymagana przez EpubCheck do walidacji)

========================================================================
SPOSOB PRACY
========================================================================

Konwersja pojedynczego pliku:
1. Uzupelnij metadane ebooka (wszystkie pola oprocz ISBN sa wymagane).
2. Wybierz plik DOCX i okladke PNG/JPG.
3. Wspiaz sciezke zapisu EPUB.
4. Kliknij "Konwertuj i zapisz EPUB".

Konwersja seryjna (batch):
1. Uzupelnij autor, wydawca, rok i wybierz okladke (wspolne dla calej serii).
2. Kliknij "Konwertuj serie..." pod glownym przyciskiem konwersji.
3. Wybierz dowolna liczbe plikow DOCX.
4. Aplikacja konwertuje kazdy plik, wykrywajac tytul z pierwszego H1.
   Pliki EPUB zapisywane sa obok plikow DOCX z tym samym miejscem.

Przygotowanie DOCX:
- Rozdzialy musza byc oznaczone stylem "Heading 1" / "Naglowek 1".
- Podrozdzialy: "Heading 2" / "Naglowek 2", "Heading 3" / "Naglowek 3".
- Listy: "List Bullet" / "Lista Numerowana".
- Cytaty: "Quote" / "Cytat".
- Aplikacja sprawdza strukture przed konwersja i wyswietla liste
  ewentualnych problemow do poprawki.

Okladka:
- Format: PNG lub JPG/JPEG.
- Zalecane proporcje: 2:3 (np. 1600 x 2400 px) — standard sklepow ebook.
- Zalecane minimum: 1200 px szerokosci.
- Aplikacja wyswietla ostrzezenie przy blednych proporcjach lub za malej
  rozdzielczosci.

========================================================================
FUNKCJE APLIKACJI
========================================================================

Metadane i zapis:
  - Tytul, Autor, Wydawca, Rok, ISBN (opcjonalny, 10 lub 13 cyfr), Opis
  - Ostatnio uzyte katalogi zapamietywane miedzy sesjami
  - Profil metadanych (autor, wydawca, rok, jezyk) zapisywany po konwersji
    i automatycznie przywracany przy nastepnym uruchomieniu
  - Automatyczna propozycja nazwy pliku EPUB na podstawie tytulu
  - Przycisk "?" przy polu ISBN — informacja o bezplatnej rejestracji
    w serwisie e-isbn.pl (Biblioteka Narodowa)

Obsluga formatowania DOCX:
  - Pogrubienie, kursywa, podkreslenie, przekreslenie
  - Naglowki H1–H3, listy wypunktowane i numerowane, cytaty
  - Tabele (w:tbl) — konwertowane do HTML <table>
  - Hiperlinki (w:hyperlink) — konwertowane do <a href="...">
  - Obrazy inline (a:blip) — wyodrebniane z DOCX i osadzane w EPUB
  - Przypisy dolne / koncowe (w:footnoteReference) — konwertowane
    do znacznikow <sup>[N]</sup> z sekcja przypisow na koncu rozdzialu

Jezyki:
  - Interfejs uzytkownika: Polski / English (przelacznik w narozna)
  - Jezyk dokumentu EPUB: 13 opcji (PL, EN, DE, FR, ES, IT, CS, SK,
    HU, RU, UK, NL, PT)
  - Spis tresci (nav.xhtml) tlumaczony na wybrany jezyk dokumentu

Okladka:
  - Podglad miniatury po wyborze pliku
  - Ostrzezenie przy proporcjach odbiegajacych od standardu 2:3
  - Ostrzezenie przy szerokosci ponizej 1200 px
  - Drag & drop (wymaga biblioteki tkinterdnd2)

Konwersja seryjna (batch):
  - Konwersja dowolnej liczby plikow DOCX za jednym razem
  - Status per plik: [ ] / [>] / [OK] / [!!]
  - Pasek postepu i podsumowanie

Podglad i eksport:
  - Interaktywny spis tresci z podgladem tekstu rozdzialow
  - Pelny podglad HTML w przegladarce
  - Walidacja EpubCheck po kazdej konwersji

Dostepnosc EPUB:
  - Semantyczne znaczniki epub:type (chapter, frontmatter, backmatter)
  - Nawigacyjny spis tresci (nav.xhtml)
  - Teksty alternatywne dla obrazow
  - Metadane schema.org (accessMode, accessibilityFeature, conformsTo)
  - Zgodnosc z EPUB Accessibility 1.1 i WCAG 2.1 AA

========================================================================
HISTORIA ZMIAN
========================================================================

v2.1 (maj 2026)
  - Refaktoryzacja: kod w src/epub_converter/ (parser, EPUB, GUI, testy)
  - Poprawka: style Heading 10 nie sa juz mylone z Heading 1
  - ISBN: suma kontrolna, identyfikator urn:isbn: w OPF
  - Alt obrazow z DOCX (descr/title) lub domyslny tekst w jezyku dokumentu
  - Preflight: przycisk „Konwertuj mimo ostrzezen”
  - Sprawdzenie Java przy starcie (EpubCheck)
  - Batch: walidacja EpubCheck, opis z formularza
  - Uruchomienie: run_converter.py (skroty i instalatory zaktualizowane)

v2.0 (maj 2026)
  - Obrazy inline wyciagane z DOCX i osadzane w EPUB
  - Tabele (w:tbl) konwertowane do <table> HTML
  - Hiperlinki (w:hyperlink) konwertowane do <a href="...">
  - Wskaznik postepu konwersji (ttk.Progressbar indeterminate)
  - Drag & drop dla pol DOCX i okladki (opcjonalny, wymaga tkinterdnd2)
  - Automatyczna propozycja nazwy pliku EPUB na podstawie tytulu
  - Przewijany pulpit (canvas) przy malych ekranach
  - Tlumaczenie tytulu spisu tresci dla 13 jezykow
  - Jezyk UI (PL/EN) respektowany w podgladzie HTML i meta tagach
  - Podkreslenie i przekreslenie w para_to_html
  - Walidacja ISBN (regex 10 lub 13 cyfr)
  - Zapamietywanie ostatnio uzytego katalogu w oknach dialogowych
  - Przycisk "?" przy ISBN z informacja o rejestracji w e-isbn.pl
  - Zapis/odczyt profilu metadanych (autor, wydawca, rok, jezyk)
    — automatyczne wypelnianie formularza przy nastepnym uruchomieniu
  - Konwersja seryjna (batch) wielu plikow DOCX za jednym razem
  - Podglad okladki z ostrzezeniem o proporcjach (standard 2:3)
    i ostrzezeniem o za malej rozdzielczosci (< 1200 px)
  - Obsluga przypisow dolnych/koncowych (w:footnoteReference)
    z sekcja przypisow na koncu rozdzialu i CSS footnotes-section

  Poprawki v2.0 (bugfix):
  - Obrazy inline w podgladzie HTML osadzane jako base64 data URI
    (wczesniej sciezki ../images/ nie dzialaly w temp dir przegladarki)
  - Footnote <aside> uzupelniony o epub:type="footnote" role="doc-footnote"
    (wymagane przez EPUB3 do poprawnej identyfikacji przypisow)
  - Batch mode: preflight DOCX przed kazda konwersja; plik bez Heading 1
    oznaczany [!!] i pomijany zamiast generowac pusty EPUB
  - Usunieto hardkodowane stringi z show_result/show_error — wszystkie
    etykiety wynikow tlumaczone przez TRANSLATIONS (PL/EN)
  - a11y_summary w OPF: angielski jako fallback dla 11 jezykow innych niz PL
    (wczesniej dokument np. po niemiecku mial polski tekst dostepnosci)
  - requirements.txt: dodana informacja o opcjonalnej bibliotece tkinterdnd2

v1.0 (styczen 2026)
  - Pierwsza wersja publiczna
  - Konwersja DOCX -> EPUB 3 z metadanymi Dublin Core
  - Walidacja EpubCheck
  - Interfejs Tkinter (PL/EN), 13 jezykow dokumentu
  - Preflight DOCX (sprawdzenie struktury rozdzialow)
  - Podglad skonwertowanego pliku

========================================================================
ZALEZNOSCI PYTHON
========================================================================

  python-docx >= 1.1   -- parsowanie plikow DOCX
  epubcheck >= 0.4     -- walidacja plikow EPUB
  Pillow >= 10.0       -- obsluga obrazow i okladki
  tkinterdnd2          -- opcjonalne; drag & drop

Instalacja zaleznosci recznie:
  pip install -r requirements.txt
  pip install tkinterdnd2   # opcjonalnie

========================================================================
LICENCJA
========================================================================

Copyright (c) 2026 Marek Zettel. Wszelkie prawa zastrzezone.
