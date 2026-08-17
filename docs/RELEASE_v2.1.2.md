# DOCX EPUB Converter v2.1.2

Poprawki walidacji EpubCheck i wyodrębniania przypisów. Zmiany są już w kodzie —
przed wydaniem podnieś `APP_VERSION` w `src/epub_converter/constants.py` na
`2.1.2` (skrypty budujące czytają wersję stamtąd) i przebuduj pakiet.

> **Uwaga:** pakiet `2.1.1.0` wysłany do Microsoft Store zawiera jeszcze błędy
> opisane poniżej. Store nie przyjmie ponownie tego samego numeru wersji, więc
> aktualizacja wymaga podbicia `APP_VERSION`.

## Naprawione

- **EpubCheck wywracał się na każdym pliku.** Wrapper `epubcheck` uruchamia Javę
  bez `-Xss`, przez co kompilacja schematu RelaxNG kończyła się
  `StackOverflowError` na domyślnym stosie JVM. Wrapper przekierowuje stderr do
  `devnull`, więc awaria była niewidoczna i zwracał `valid=False` z pustą listą
  komunikatów — poprawny EPUB był zgłaszany jako wadliwy. Aplikacja wywołuje
  teraz `epubcheck.jar` bezpośrednio z `-Xss16m`.

- **Awaria walidatora była prezentowana jako wada książki.** `validate_epub()`
  zwraca teraz `(ran, valid, messages)`. Gdy walidator nie zdołał wydać
  werdyktu, `ran=False`, a interfejs informuje „nie zwalidowano" zamiast
  sugerować błąd w pliku użytkownika.

- **Okno wyników wywalało się przy prawdziwym błędzie walidacji.**
  `epubcheck.models.Message.__str__` rzuca `TypeError`, gdy EpubCheck nie poda
  pola `suggestion` (przypadek bardzo częsty), a interfejs formatuje komunikaty
  przez `str()`. Komunikaty są teraz budowane jako bezpieczne łańcuchy znaków.
  Błąd był zamaskowany przez poprzedni — walidacja nigdy nie działała, więc
  ścieżka błędu nigdy się nie wykonywała.

- **Przypisy nie były wyodrębniane z żadnego dokumentu.** `python-docx` nie
  rejestruje typu zawartości `footnotes+xml`, więc `word/footnotes.xml` ładuje
  się jako generyczny `Part` bez atrybutu `_element`. `extract_footnotes()`
  łapało wyjątek i zwracało pusty słownik: znaczniki `<sup>` powstawały, ale
  sekcja przypisów nigdy, a odnośniki `href="#fn-N"` wskazywały na nieistniejące
  kotwice. Dodano fallback do parsowania surowych bajtów części.

- **ISBN-10 z cyfrą kontrolną „X" był odrzucany.** `_check_isbn10()` od zawsze
  obsługiwał „X", ale regex formatu `\d{10}` blokował go wcześniej, czyniąc tę
  gałąź martwym kodem. Wzorzec zmieniony na `\d{9}[\dXx]|\d{13}`.

## Zmiany techniczne

- Ustawienia przeniesione z `~/.epub_converter_config.json` do
  `%LOCALAPPDATA%\DocxEpubConverter\config.json` (wymagane przy instalacji
  read-only z MSIX); stary plik jest migrowany przy pierwszym odczycie.
- Import `epubcheck` jest leniwy — brak pakietu lub środowiska Java nie
  przerywa startu aplikacji ani konwersji.
- Wersja czytana z `constants.py` przez wszystkie skrypty budujące zamiast być
  wpisana na sztywno w trzech miejscach.

## Testy

137 testów (`pytest tests`). Pokrycie: `docx_parser` 96%, `preflight` 100%,
`isbn` 100%, `styles` 100%, `config` 91%.

Autor: Marek Zettel
