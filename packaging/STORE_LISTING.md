# Microsoft Store listing — DOCX EPUB Converter

Copy-paste source for Partner Center. Character limits noted per field.

---

## Product identity (already in AppxManifest.xml)

| Field | Value |
|---|---|
| Package/Identity/Name | `MarekZettel-zetmar.DOCXEPUBConverter` |
| Package/Identity/Publisher | `CN=15A53D32-C868-48EE-B700-5DBB5449CA1B` |
| Package/Properties/PublisherDisplayName | `Marek Zettel - zetmar` |
| Version (MSIX) | `2.1.1.0` — bump the third part per submission, keep the 4th at `0` |

## Properties

| Field | Value |
|---|---|
| Category | Developer tools **or** Productivity (recommend **Productivity**) |
| Subcategory | — |
| Privacy policy | No hosting needed — Partner Center offers **"Provide privacy policy text"** as an alternative to a URL. The plain-text version of `PRIVACY.md` is pasted there. |
| Website | https://github.com/zetmar-collab/docx-epub-converter |
| Support contact | zetmar@gmail.com |
| Product declarations | Leave **"This product has been tested to meet accessibility guidelines" unticked.** The declaration is about the app's *own* Tkinter UI, not the accessible EPUBs it produces. The UI has had no screen-reader or keyboard-navigation testing, so ticking it would be an unverified claim. Tick it only after actually testing the UI. |
| System requirements | Leave unspecified. Declaring Keyboard/Mouse as minimum can show warnings to touch-only customers and block them from reviewing, with no benefit for a desktop utility. |

## Age ratings

No user-generated content, no ads, no data collection, no in-app purchases, no
network features → the IARC questionnaire yields the lowest rating (PEGI 3 /
ESRB Everyone).

## Pricing

Recommendation: **Free** for launch, or a one-time price. There is no in-app
purchase code in the product.

---

## Listing — English (en-US)

**Product name** (max 256)

```
DOCX EPUB Converter
```

**Short description** (max 1000)

```
Turn a Word manuscript into a bookstore-ready, accessible EPUB 3 ebook in one click — no subscription, no upload, fully offline.
```

**Description** (max 10000)

```
DOCX EPUB Converter takes the Word document you already wrote and produces a clean, standards-compliant EPUB 3 ebook that passes accessibility requirements — without sending your manuscript to any server.

Everything runs offline on your own computer. Your book never leaves your machine.

WHY ACCESSIBILITY MATTERS
Since the European Accessibility Act (EAA) came into force, ebooks sold in the EU must be accessible. This app builds that in by default: semantic HTML structure, a navigational table of contents, alternative text for images, and complete schema.org accessibility metadata — targeting EPUB Accessibility 1.1 and WCAG 2.1 AA.

WHAT IT CONVERTS
• Headings H1–H3 become real chapter structure
• Bold, italic, underline, strikethrough
• Bulleted and numbered lists
• Block quotes
• Tables, converted to proper HTML tables
• Hyperlinks
• Inline images, extracted and embedded
• Footnotes and endnotes, with a per-chapter notes section

BEFORE YOU CONVERT
A pre-flight check inspects your DOCX and lists exactly what needs fixing — missing heading styles, unstyled chapters, structural problems — before conversion, not after.

COVER HANDLING
Pick a PNG or JPG cover and see it previewed instantly. The app warns you if the aspect ratio strays from the 2:3 ebook-store standard or if the width is below 1200 px.

BATCH CONVERSION
Converting a whole series? Set the shared author, publisher, year and cover once, then select any number of DOCX files. Each one is converted with its title detected from the first H1, with per-file status and a progress summary.

PREVIEW AND VALIDATE
Browse an interactive table of contents with chapter text inside the app, or open a full HTML preview in your browser. If a Java runtime is present, every conversion is additionally validated with EpubCheck — the industry-standard EPUB validator.

MULTILINGUAL
Interface in Polish and English. Document language in 13 options: Polish, English, German, French, Spanish, Italian, Czech, Slovak, Hungarian, Russian, Ukrainian, Dutch and Portuguese — with the table of contents translated to match.

CONVENIENCE
Your author, publisher, year and language are remembered between sessions. File dialogs reopen where you left off. Output filenames are proposed automatically from the title. A built-in hint explains free ISBN registration through the Polish National Library.

PRIVACY
No account. No telemetry. No analytics. No ads. No network connections. Nothing is collected.

REQUIREMENTS
Windows 10 version 1809 or later, 64-bit. A Java runtime is optional and only needed for EpubCheck validation — conversion works without it.
```

**What's new in this version** (max 1500)

```
Version 2.1.1
• EpubCheck fix: removed an invalid role="doc-cover" attribute on the cover page
• Footnotes now use role="note" instead of role="doc-footnote"
• Settings moved to the standard per-user application data folder
• EPUB validation degrades gracefully when no Java runtime is present
```

**Search terms** (max 7 terms, 30 chars each)

```
docx to epub
epub converter
ebook
word to ebook
accessible epub
epub 3
self publishing
```

**Copyright and trademark info**

```
© 2026 Marek Zettel. All rights reserved.
```

**Additional licence terms**

```
Provided as-is under the MIT Licence. See the LICENSE file included with the app.
```

---

## Listing — Polski (pl-PL)

**Nazwa produktu**

```
DOCX EPUB Converter
```

**Krótki opis**

```
Zamień maszynopis z Worda w gotowy do sprzedaży, dostępny ebook EPUB 3 jednym kliknięciem — bez abonamentu, bez wysyłania plików, w pełni offline.
```

**Opis**

```
DOCX EPUB Converter bierze dokument Word, który już napisałeś, i tworzy czysty, zgodny ze standardami ebook EPUB 3, spełniający wymogi dostępności — bez wysyłania maszynopisu na jakikolwiek serwer.

Wszystko działa offline, na Twoim komputerze. Twoja książka nigdy go nie opuszcza.

DLACZEGO DOSTĘPNOŚĆ MA ZNACZENIE
Od wejścia w życie Europejskiego Aktu o Dostępności (EAA) ebooki sprzedawane w UE muszą być dostępne cyfrowo. Ta aplikacja zapewnia to domyślnie: semantyczna struktura HTML, nawigacyjny spis treści, teksty alternatywne obrazów i kompletne metadane dostępności schema.org — zgodnie z EPUB Accessibility 1.1 i WCAG 2.1 AA.

CO KONWERTUJE
• Nagłówki H1–H3 stają się prawdziwą strukturą rozdziałów
• Pogrubienie, kursywa, podkreślenie, przekreślenie
• Listy wypunktowane i numerowane
• Cytaty blokowe
• Tabele, konwertowane do poprawnych tabel HTML
• Hiperlinki
• Obrazy inline, wyodrębniane i osadzane
• Przypisy dolne i końcowe, z sekcją przypisów w każdym rozdziale

ZANIM SKONWERTUJESZ
Kontrola wstępna sprawdza plik DOCX i wypisuje dokładnie to, co wymaga poprawki — brakujące style nagłówków, rozdziały bez stylu, problemy strukturalne — przed konwersją, a nie po niej.

OKŁADKA
Wybierz okładkę PNG lub JPG i zobacz natychmiastowy podgląd. Aplikacja ostrzeże Cię, gdy proporcje odbiegają od standardu 2:3 księgarni ebooków lub gdy szerokość jest mniejsza niż 1200 px.

KONWERSJA SERYJNA
Konwertujesz całą serię? Ustaw wspólnego autora, wydawcę, rok i okładkę raz, a potem wskaż dowolną liczbę plików DOCX. Każdy zostanie skonwertowany z tytułem wykrytym z pierwszego nagłówka H1, ze statusem dla każdego pliku i podsumowaniem postępu.

PODGLĄD I WALIDACJA
Przeglądaj interaktywny spis treści z tekstem rozdziałów w aplikacji albo otwórz pełny podgląd HTML w przeglądarce. Jeśli w systemie jest środowisko Java, każda konwersja jest dodatkowo sprawdzana narzędziem EpubCheck — branżowym standardem walidacji EPUB.

WIELOJĘZYCZNOŚĆ
Interfejs po polsku i angielsku. Język dokumentu w 13 wariantach: polski, angielski, niemiecki, francuski, hiszpański, włoski, czeski, słowacki, węgierski, rosyjski, ukraiński, niderlandzki i portugalski — wraz z przetłumaczonym spisem treści.

WYGODA
Autor, wydawca, rok i język są zapamiętywane między sesjami. Okna wyboru plików otwierają się tam, gdzie skończyłeś. Nazwa pliku wynikowego jest proponowana automatycznie na podstawie tytułu. Wbudowana podpowiedź wyjaśnia bezpłatną rejestrację ISBN w Bibliotece Narodowej.

PRYWATNOŚĆ
Bez konta. Bez telemetrii. Bez analityki. Bez reklam. Bez połączeń sieciowych. Nic nie jest zbierane.

WYMAGANIA
Windows 10 w wersji 1809 lub nowszy, 64-bitowy. Środowisko Java jest opcjonalne i potrzebne wyłącznie do walidacji EpubCheck — konwersja działa bez niego.
```

**Nowości w tej wersji**

```
Wersja 2.1.1
• Poprawka EpubCheck: usunięto nieprawidłowy atrybut role="doc-cover" na stronie okładki
• Przypisy używają teraz role="note" zamiast role="doc-footnote"
• Ustawienia przeniesione do standardowego folderu danych aplikacji użytkownika
• Walidacja EPUB działa łagodnie, gdy w systemie brak środowiska Java
```

**Wyszukiwane terminy**

```
docx na epub
konwerter epub
ebook
word na ebook
dostępny epub
epub 3
self publishing
```

---

## Screenshots required

Minimum 1, maximum 10 per language. **1366 × 768** or **2160 × 1440** PNG.
Suggested set, captured from a real run:

1. Main window with metadata filled in and a cover preview loaded
2. Pre-flight check dialog listing DOCX structure issues
3. Interactive chapter preview with the table of contents
4. Batch conversion window mid-run, with per-file status
5. Successful result view showing the EpubCheck verdict

Store logo (300 × 300) and the optional 1240 × 600 promotional image are
generated into `packaging/Assets/` — see `PromoImage1240x600.png`.
