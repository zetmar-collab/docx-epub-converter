"""
DOCX EPUB Converter
EAA / EPUB Accessibility 1.1 / WCAG 2.1 AA
Autor: Marek Zettel
"""

from __future__ import annotations

import html
import io
import os
import re
import subprocess
import sys
import tempfile
import threading
import webbrowser
import zipfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

from docx import Document
from epubcheck import EpubCheck
from PIL import Image, ImageTk


APP_TITLE = "DOCX EPUB Converter"
APP_AUTHOR = "Marek Zettel"
NAVY_BG = "#071a33"
NAVY_PANEL = "#0f2747"
NAVY_PANEL_ALT = "#14365f"
NAVY_BORDER = "#27527f"
NAVY_INPUT = "#f4f8ff"
NAVY_TEXT = "#edf5ff"
NAVY_MUTED = "#b8cbe3"
NAVY_ACCENT = "#2f80d0"
NAVY_ACCENT_ACTIVE = "#4aa3ff"

DOC_LANGUAGES = [
    ("Polski (pl)", "pl"),
    ("English (en)", "en"),
    ("Deutsch (de)", "de"),
    ("Français (fr)", "fr"),
    ("Español (es)", "es"),
    ("Italiano (it)", "it"),
    ("Čeština (cs)", "cs"),
    ("Slovenčina (sk)", "sk"),
    ("Magyar (hu)", "hu"),
    ("Русский (ru)", "ru"),
    ("Українська (uk)", "uk"),
    ("Nederlands (nl)", "nl"),
    ("Português (pt)", "pt"),
]

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
        "issue_tables": "Dokument zawiera tabele. Ten konwerter nie przenosi tabel do EPUB, wiec tresc moglaby zostac utracona.",
        "issue_inline_shapes": "Dokument zawiera obrazy w tresci. Ten konwerter obsluguje osobna okladke, ale nie przenosi obrazow z DOCX.",
        "issue_empty_h1": "Jeden z naglowkow Heading 1 / Naglowek 1 jest pusty.",
        "issue_h3_without_h2": "Dokument zawiera naglowek Heading 3 / Naglowek 3 bez poprzedzajacego Heading 2.",
        "issue_no_chapters": "Brakuje rozdzialow oznaczonych stylem Heading 1 / Naglowek 1. EPUB wymaga poprawnej struktury rozdzialow.",
        "issue_text_before_h1": (
            "Przed pierwszym naglowkiem Heading 1 / Naglowek 1 znajduje sie tekst. "
            "Konwerter zaczyna EPUB od pierwszego rozdzialu, wiec ten tekst moglby zostac pominiety."
        ),
        "issue_empty_chapters": "Te rozdzialy nie maja tresci pod naglowkiem: ",
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
        "issue_tables": "The document contains tables. This converter does not transfer tables to EPUB, so content might be lost.",
        "issue_inline_shapes": "The document contains inline images. This converter handles a separate cover image but does not transfer inline images from DOCX.",
        "issue_empty_h1": "One of the Heading 1 / Nagłówek 1 headings is empty.",
        "issue_h3_without_h2": "The document contains a Heading 3 / Nagłówek 3 without a preceding Heading 2.",
        "issue_no_chapters": "No chapters marked with Heading 1 / Nagłówek 1 style found. EPUB requires proper chapter structure.",
        "issue_text_before_h1": (
            "There is text before the first Heading 1 / Nagłówek 1. "
            "The converter starts the EPUB from the first chapter, so this text might be skipped."
        ),
        "issue_empty_chapters": "These chapters have no body text under their heading: ",
    },
}


def xml_escape(text: str) -> str:
    return html.escape(text or "", quote=False)


def para_to_html(para) -> str:
    """Konwertuje akapit DOCX na HTML z obsluga podstawowego formatowania."""
    parts = []
    for run in para.runs:
        text = xml_escape(run.text)
        if not text:
            continue
        if run.bold and run.italic:
            text = f"<strong><em>{text}</em></strong>"
        elif run.bold:
            text = f"<strong>{text}</strong>"
        elif run.italic:
            text = f"<em>{text}</em>"
        parts.append(text)
    return "".join(parts)


def parse_docx(docx_bytes: bytes):
    doc = Document(io.BytesIO(docx_bytes))
    chapters = []
    cur_id = None
    cur_title = None
    cur_body = []
    in_ul = False
    in_ol = False
    chapter_counter = 0

    front_keywords = {"wstep", "wstęp", "przedmowa", "wprowadzenie", "preface", "introduction", "foreword", "prolog"}
    back_keywords = {"zakonczenie", "zakończenie", "epilog", "podsumowanie", "conclusion", "afterword", "epilogue"}

    def close_list(parts):
        nonlocal in_ul, in_ol
        if in_ul:
            parts.append("</ul>")
            in_ul = False
        if in_ol:
            parts.append("</ol>")
            in_ol = False

    def flush_chapter():
        if cur_title is None:
            return
        close_list(cur_body)
        key = html.unescape(cur_title).lower().strip()
        if key in front_keywords:
            epub_type, role = "frontmatter", "doc-preface"
        elif key in back_keywords:
            epub_type, role = "backmatter", "doc-conclusion"
        else:
            epub_type, role = "chapter", "doc-chapter"
        chapters.append((cur_id, cur_title, epub_type, role, "\n".join(cur_body)))

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        text = para.text.strip()

        is_h1 = any(name in style_name for name in ["Heading 1", "Nagłówek 1"])
        is_h2 = any(name in style_name for name in ["Heading 2", "Nagłówek 2"])
        is_h3 = any(name in style_name for name in ["Heading 3", "Nagłówek 3"])
        is_quote = any(name in style_name for name in ["Quote", "Cytat", "Blockquote", "Intense Quote", "Intensywny cytat"])
        is_bullet = any(name in style_name for name in ["List Bullet", "Lista wypunktowana", "List Paragraph"])
        is_number = any(name in style_name for name in ["List Number", "Lista numerowana"])

        if is_h1:
            flush_chapter()
            chapter_counter += 1
            cur_id = "ch" + str(chapter_counter).zfill(2)
            cur_title = xml_escape(text) if text else "Rozdzial"
            cur_body = []
            in_ul = False
            in_ol = False
        elif cur_title is None:
            continue
        elif is_h2:
            close_list(cur_body)
            cur_body.append("<h2>" + (para_to_html(para) or xml_escape(text)) + "</h2>")
        elif is_h3:
            close_list(cur_body)
            cur_body.append("<h3>" + (para_to_html(para) or xml_escape(text)) + "</h3>")
        elif is_quote:
            close_list(cur_body)
            cur_body.append("<blockquote><p>" + (para_to_html(para) or xml_escape(text)) + "</p></blockquote>")
        elif is_bullet:
            if not in_ul:
                if in_ol:
                    cur_body.append("</ol>")
                    in_ol = False
                cur_body.append("<ul>")
                in_ul = True
            cur_body.append("<li>" + (para_to_html(para) or xml_escape(text)) + "</li>")
        elif is_number:
            if not in_ol:
                if in_ul:
                    cur_body.append("</ul>")
                    in_ul = False
                cur_body.append("<ol>")
                in_ol = True
            cur_body.append("<li>" + (para_to_html(para) or xml_escape(text)) + "</li>")
        elif text:
            close_list(cur_body)
            cur_body.append("<p>" + (para_to_html(para) or xml_escape(text)) + "</p>")
        else:
            close_list(cur_body)

    flush_chapter()
    return chapters


def inspect_docx_for_epub(docx_bytes: bytes, tr: dict | None = None) -> list[str]:
    """Sprawdza, czy DOCX da sie bezpiecznie zamienic tym konwerterem na EPUB."""
    if tr is None:
        tr = TRANSLATIONS["pl"]
    issues = []
    doc = Document(io.BytesIO(docx_bytes))

    if doc.tables:
        issues.append(tr["issue_tables"])

    if doc.inline_shapes:
        issues.append(tr["issue_inline_shapes"])

    heading1_count = 0
    seen_heading1 = False
    nonempty_before_first_heading = []
    current_chapter_title = ""
    current_chapter_has_body = False
    empty_chapters = []
    last_heading_level = 0

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        text = para.text.strip()
        if not text:
            continue

        is_h1 = any(name in style_name for name in ["Heading 1", "Nagłówek 1"])
        is_h2 = any(name in style_name for name in ["Heading 2", "Nagłówek 2"])
        is_h3 = any(name in style_name for name in ["Heading 3", "Nagłówek 3"])

        if is_h1:
            if current_chapter_title and not current_chapter_has_body:
                empty_chapters.append(current_chapter_title)
            heading1_count += 1
            seen_heading1 = True
            current_chapter_title = text
            current_chapter_has_body = False
            last_heading_level = 1
            if not text:
                issues.append(tr["issue_empty_h1"])
            continue

        if not seen_heading1:
            nonempty_before_first_heading.append(text[:80])
            continue

        if is_h2:
            last_heading_level = 2
        elif is_h3:
            if last_heading_level < 2:
                issues.append(tr["issue_h3_without_h2"])
            last_heading_level = 3
        else:
            current_chapter_has_body = True

    if current_chapter_title and not current_chapter_has_body:
        empty_chapters.append(current_chapter_title)

    if heading1_count == 0:
        issues.append(tr["issue_no_chapters"])

    if nonempty_before_first_heading:
        issues.append(tr["issue_text_before_h1"])

    if empty_chapters:
        sample = ", ".join(empty_chapters[:5])
        issues.append(tr["issue_empty_chapters"] + sample)

    return issues


CSS = (
    "body{font-family:Georgia,'Times New Roman',serif;font-size:1em;line-height:1.8;"
    "margin:0 auto;max-width:40em;padding:1em 1.5em;color:#111;background:#fff}\n"
    "h1{font-size:1.6em;margin-top:2em;margin-bottom:.4em;color:#1a3a6b}\n"
    "h2{font-size:1.25em;margin-top:1.6em;margin-bottom:.3em;color:#2a5298}\n"
    "h3{font-size:1.05em;margin-top:1.2em;margin-bottom:.2em;font-weight:bold}\n"
    "p{margin:.5em 0;text-align:justify}\n"
    "ul,ol{margin:.4em 0 .4em 1.4em}li{margin:.2em 0}\n"
    "blockquote{margin:1em;padding:.5em 1em;border-left:4px solid #2a5298;"
    "background:#f4f7fc;font-style:italic;color:#333}\n"
    "strong{font-weight:bold}em{font-style:italic}\n"
)


def e(text: str) -> bytes:
    return text.encode("utf-8")


def make_container() -> bytes:
    return e(
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">\n'
        "  <rootfiles>\n"
        '    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>\n'
        "  </rootfiles>\n"
        "</container>\n"
    )


def make_cover_xhtml(cover_alt: str, cover_ext: str, lang: str = "pl") -> bytes:
    return e(
        '<?xml version="1.0" encoding="utf-8"?>\n'
        "<!DOCTYPE html>\n"
        '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops"'
        ' xml:lang="' + lang + '" lang="' + lang + '">\n'
        "<head>\n"
        '  <meta charset="utf-8"/>\n'
        "  <title>Cover</title>\n"
        "  <style>body{margin:0;padding:0;background:#000}.w{display:flex;align-items:center;"
        "justify-content:center;width:100%;height:100vh}.c{max-width:100%;max-height:100vh;"
        "object-fit:contain}</style>\n"
        "</head>\n"
        "<body>\n"
        '  <section epub:type="cover"><div class="w"><img class="c" src="../images/cover'
        + cover_ext
        + '" alt="'
        + xml_escape(cover_alt)
        + '"/></div></section>\n'
        "</body>\n"
        "</html>\n"
    )


def make_chapter_xhtml(title: str, epub_type: str, role: str, body: str, lang: str = "pl") -> bytes:
    return e(
        '<?xml version="1.0" encoding="utf-8"?>\n'
        "<!DOCTYPE html>\n"
        '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops"'
        ' xml:lang="' + lang + '" lang="' + lang + '">\n'
        "<head>\n"
        '  <meta charset="utf-8"/>\n'
        "  <title>" + title + "</title>\n"
        '  <link rel="stylesheet" type="text/css" href="../styles/main.css"/>\n'
        "</head>\n"
        "<body>\n"
        '  <section epub:type="' + epub_type + '" role="' + role + '">\n'
        "    <h1>" + title + "</h1>\n"
        + body
        + "\n  </section>\n"
        "</body>\n"
        "</html>\n"
    )


def make_nav(chapters, lang: str = "pl") -> bytes:
    items = ""
    for cid, title, _epub_type, _role, _body in chapters:
        items += '      <li><a href="content/' + cid + '.xhtml">' + title + "</a></li>\n"
    toc_title = "Table of Contents" if lang == "en" else "Spis tresci"
    return e(
        '<?xml version="1.0" encoding="utf-8"?>\n'
        "<!DOCTYPE html>\n"
        '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops"'
        ' xml:lang="' + lang + '" lang="' + lang + '">\n'
        '<head><meta charset="utf-8"/><title>' + toc_title + '</title></head>\n'
        "<body>\n"
        '  <nav epub:type="toc" id="toc" role="doc-toc" aria-label="' + toc_title + '">\n'
        "    <h1>" + toc_title + "</h1>\n"
        "    <ol>\n"
        + items
        + "    </ol>\n"
        "  </nav>\n"
        "</body>\n"
        "</html>\n"
    )


def make_opf(meta: dict, chapters, cover_mime: str, cover_ext: str, lang: str = "pl") -> bytes:
    manifest_items = (
        '    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>\n'
        '    <item id="css" href="styles/main.css" media-type="text/css"/>\n'
        '    <item id="cover-image" href="images/cover' + cover_ext + '" media-type="' + cover_mime + '" properties="cover-image"/>\n'
        '    <item id="cover-page" href="content/cover_page.xhtml" media-type="application/xhtml+xml"/>\n'
    )
    spine = '    <itemref idref="cover-page" linear="yes"/>\n'
    for cid, _title, _epub_type, _role, _body in chapters:
        manifest_items += '    <item id="' + cid + '" href="content/' + cid + '.xhtml" media-type="application/xhtml+xml"/>\n'
        spine += '    <itemref idref="' + cid + '"/>\n'

    today = date.today().isoformat()
    a11y_summary = (
        "Digitally accessible publication. Semantic HTML structure, navigational table of contents, "
        "alternative texts, complete metadata. Compliant with EPUB Accessibility 1.1 and WCAG 2.1 AA."
        if lang == "en"
        else "Publikacja dostepna cyfrowo. Semantyczna struktura HTML, nawigacyjny spis tresci, "
        "teksty alternatywne, kompletne metadane. Zgodna z EPUB Accessibility 1.1 i WCAG 2.1 AA."
    )
    return e(
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<package xmlns="http://www.idpf.org/2007/opf"\n'
        '         xmlns:dc="http://purl.org/dc/elements/1.1/"\n'
        '         xmlns:dcterms="http://purl.org/dc/terms/"\n'
        '         version="3.0" unique-identifier="uid" xml:lang="' + lang + '">\n'
        '  <metadata xmlns:schema="http://schema.org/">\n'
        '    <dc:identifier id="uid">' + xml_escape(meta["isbn"] or meta["title"]) + "-" + today + "</dc:identifier>\n"
        "    <dc:title>" + xml_escape(meta["title"]) + "</dc:title>\n"
        '    <dc:creator id="creator">' + xml_escape(meta["author"]) + "</dc:creator>\n"
        '    <meta refines="#creator" property="role" scheme="marc:relators">aut</meta>\n'
        "    <dc:language>" + xml_escape(lang) + "</dc:language>\n"
        "    <dc:publisher>" + xml_escape(meta["publisher"]) + "</dc:publisher>\n"
        "    <dc:date>" + xml_escape(meta["year"]) + "</dc:date>\n"
        "    <dc:description>" + xml_escape(meta["description"]) + "</dc:description>\n"
        "    <dc:subject>Ebook</dc:subject>\n"
        "    <dc:rights>Copyright " + xml_escape(meta["year"]) + " " + xml_escape(meta["author"]) + "</dc:rights>\n"
        '    <meta property="dcterms:modified">' + today + "T00:00:00Z</meta>\n"
        '    <meta property="schema:accessMode">textual</meta>\n'
        '    <meta property="schema:accessMode">visual</meta>\n'
        '    <meta property="schema:accessModeSufficient">textual</meta>\n'
        '    <meta property="schema:accessibilityFeature">structuralNavigation</meta>\n'
        '    <meta property="schema:accessibilityFeature">tableOfContents</meta>\n'
        '    <meta property="schema:accessibilityFeature">readingOrder</meta>\n'
        '    <meta property="schema:accessibilityFeature">alternativeText</meta>\n'
        '    <meta property="schema:accessibilityFeature">displayTransformability</meta>\n'
        '    <meta property="schema:accessibilityHazard">none</meta>\n'
        '    <meta property="schema:accessibilitySummary">' + xml_escape(a11y_summary) + '</meta>\n'
        '    <meta property="dcterms:conformsTo">EPUB Accessibility 1.1 - WCAG 2.1 Level AA</meta>\n'
        '    <link rel="dcterms:conformsTo" href="https://www.w3.org/TR/epub-a11y-11/#wcag-aa"/>\n'
        "  </metadata>\n"
        "  <manifest>\n"
        + manifest_items
        + "  </manifest>\n"
        "  <spine>\n"
        + spine
        + "  </spine>\n"
        "</package>\n"
    )


def build_epub(meta: dict, chapters, cover_bytes: bytes, cover_ext: str, lang: str = "pl") -> bytes:
    cover_mime = "image/jpeg" if cover_ext.lower() in (".jpg", ".jpeg") else "image/png"
    cover_alt = "Cover: " + meta["title"] + ", " + meta["author"] if lang == "en" else "Okladka: " + meta["title"] + ", autor " + meta["author"]

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        mimetype = zipfile.ZipInfo("mimetype")
        mimetype.compress_type = zipfile.ZIP_STORED
        zf.writestr(mimetype, "application/epub+zip")
        zf.writestr("META-INF/container.xml", make_container())
        zf.writestr("OEBPS/content.opf", make_opf(meta, chapters, cover_mime, cover_ext, lang))
        zf.writestr("OEBPS/nav.xhtml", make_nav(chapters, lang))
        zf.writestr("OEBPS/styles/main.css", e(CSS))
        zf.writestr("OEBPS/images/cover" + cover_ext, cover_bytes)
        zf.writestr("OEBPS/content/cover_page.xhtml", make_cover_xhtml(cover_alt, cover_ext, lang))

        for cid, title, epub_type, role, body in chapters:
            zf.writestr("OEBPS/content/" + cid + ".xhtml", make_chapter_xhtml(title, epub_type, role, body, lang))
    return buffer.getvalue()


def validate_epub(epub_bytes: bytes) -> tuple[bool, list]:
    with tempfile.NamedTemporaryFile(suffix=".epub", delete=False) as tmp:
        tmp.write(epub_bytes)
        tmp_path = tmp.name
    try:
        result = EpubCheck(tmp_path)
        return result.valid, result.messages
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def safe_filename(title: str) -> str:
    name = re.sub(r"[^\w\-]+", "-", title.lower(), flags=re.UNICODE).strip("-")
    return (name or "ebook") + ".epub"


def default_output_dir() -> Path:
    downloads = Path.home() / "Downloads"
    return downloads if downloads.exists() else Path.home()


def html_to_plain_text(fragment: str) -> str:
    text = re.sub(r"</(p|h1|h2|h3|li|blockquote)>", "\n", fragment)
    text = re.sub(r"<li>", "- ", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def open_path(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


@dataclass
class ConversionResult:
    meta: dict
    chapters: list
    epub_bytes: bytes
    cover_bytes: bytes
    cover_ext: str
    saved_path: Path
    valid: bool
    messages: list


class EpubConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x760")
        self.minsize(980, 680)

        self.ui_lang_var = tk.StringVar(value="pl")
        self.doc_lang_var = tk.StringVar(value="pl")

        self.docx_path = tk.StringVar()
        self.cover_path = tk.StringVar()
        self.output_path = tk.StringVar(value=str(default_output_dir() / "ebook.epub"))
        self.title_var = tk.StringVar()
        self.author_var = tk.StringVar()
        self.publisher_var = tk.StringVar()
        self.year_var = tk.StringVar(value=str(date.today().year))
        self.isbn_var = tk.StringVar()
        self.status_var = tk.StringVar(value=TRANSLATIONS["pl"]["status_ready"])

        self.cover_preview = None
        self.result: ConversionResult | None = None
        self.preview_html_path: Path | None = None
        self.convert_buttons: list[ttk.Button] = []
        self._tw: dict[str, tk.Widget] = {}
        self._doc_lang_codes: list[str] = [code for _, code in DOC_LANGUAGES]

        self._configure_style()
        self._build_ui()

    def t(self, key: str) -> str:
        return TRANSLATIONS[self.ui_lang_var.get()].get(key, key)

    def _configure_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        self.configure(background=NAVY_BG)
        style.configure("TFrame", background=NAVY_BG)
        style.configure("Card.TFrame", background=NAVY_PANEL, relief="solid", borderwidth=1)
        style.configure("TLabel", background=NAVY_BG, foreground=NAVY_TEXT, font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=NAVY_PANEL, foreground=NAVY_TEXT, font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=NAVY_BG, font=("Segoe UI", 18, "bold"), foreground=NAVY_TEXT)
        style.configure("Section.TLabel", background=NAVY_PANEL, font=("Segoe UI", 11, "bold"), foreground=NAVY_TEXT)
        style.configure(
            "TEntry",
            fieldbackground=NAVY_INPUT,
            foreground=NAVY_BG,
            bordercolor=NAVY_BORDER,
            lightcolor=NAVY_BORDER,
            darkcolor=NAVY_BORDER,
            insertcolor=NAVY_BG,
        )
        style.configure(
            "TButton",
            background=NAVY_PANEL_ALT,
            foreground=NAVY_TEXT,
            bordercolor=NAVY_BORDER,
            focusthickness=1,
            focuscolor=NAVY_ACCENT,
            padding=(10, 5),
        )
        style.map(
            "TButton",
            background=[("active", NAVY_ACCENT), ("disabled", "#1b3048")],
            foreground=[("disabled", "#788da6")],
        )
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            background=NAVY_ACCENT,
            foreground="#ffffff",
            bordercolor=NAVY_ACCENT_ACTIVE,
        )
        style.map("Primary.TButton", background=[("active", NAVY_ACCENT_ACTIVE), ("disabled", "#1b3048")])
        style.configure(
            "Lang.TButton",
            font=("Segoe UI", 9, "bold"),
            background=NAVY_PANEL_ALT,
            foreground=NAVY_TEXT,
            bordercolor=NAVY_BORDER,
            padding=(6, 3),
        )
        style.map("Lang.TButton", background=[("active", NAVY_ACCENT_ACTIVE)])
        style.configure("Status.TLabel", background=NAVY_PANEL_ALT, foreground=NAVY_TEXT, padding=8)
        style.configure("TSeparator", background=NAVY_BORDER)
        style.configure("Vertical.TScrollbar", background=NAVY_PANEL_ALT, troughcolor=NAVY_BG, bordercolor=NAVY_BORDER)
        style.configure(
            "TCombobox",
            fieldbackground=NAVY_INPUT,
            background=NAVY_PANEL_ALT,
            foreground=NAVY_BG,
            selectbackground=NAVY_ACCENT,
            selectforeground="#ffffff",
            bordercolor=NAVY_BORDER,
        )
        style.map("TCombobox", fieldbackground=[("readonly", NAVY_INPUT)], foreground=[("readonly", NAVY_BG)])

    def _build_ui(self):
        root = ttk.Frame(self, padding=16)
        root.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(root)
        header.pack(fill=tk.X, pady=(0, 12))

        title_group = ttk.Frame(header)
        title_group.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(title_group, text=APP_TITLE, style="Title.TLabel").pack(anchor=tk.W)
        self._tw["app_subtitle"] = ttk.Label(title_group, text=self.t("app_subtitle"))
        self._tw["app_subtitle"].pack(anchor=tk.W, pady=(4, 0))
        self._tw["app_author_label"] = ttk.Label(title_group, text=self.t("app_author_label"))
        self._tw["app_author_label"].pack(anchor=tk.W, pady=(4, 0))

        right_group = ttk.Frame(header)
        right_group.pack(side=tk.RIGHT)
        lang_btn = ttk.Button(right_group, text="EN", command=self._toggle_ui_lang, style="Lang.TButton", width=4)
        lang_btn.pack(anchor=tk.E, pady=(0, 6))
        self._tw["lang_toggle"] = lang_btn
        self._add_convert_button(right_group, anchor=tk.E, ipadx=18, ipady=6)

        main = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        main.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(main, style="Card.TFrame", padding=14)
        right = ttk.Frame(main, style="Card.TFrame", padding=14)
        main.add(left, weight=1)
        main.add(right, weight=2)

        self._build_form(left)
        self._build_preview(right)

        bottom = ttk.Frame(root)
        bottom.pack(fill=tk.X, pady=(12, 0))
        self._add_convert_button(bottom, side=tk.RIGHT, ipadx=24, ipady=7)

        ttk.Label(root, textvariable=self.status_var, style="Status.TLabel").pack(fill=tk.X, pady=(12, 0))

    def _add_convert_button(self, parent, **pack_options):
        button = ttk.Button(
            parent,
            text=self.t("btn_convert"),
            command=self.convert,
            style="Primary.TButton",
        )
        self.convert_buttons.append(button)
        if pack_options:
            button.pack(**pack_options)
        return button

    def _build_form(self, parent):
        self._tw["section_metadata"] = ttk.Label(parent, text=self.t("section_metadata"), style="Section.TLabel")
        self._tw["section_metadata"].pack(anchor=tk.W)

        self._tw["field_title"] = self._entry(parent, "field_title", self.title_var)
        self._tw["field_author"] = self._entry(parent, "field_author", self.author_var)
        self._tw["field_publisher"] = self._entry(parent, "field_publisher", self.publisher_var)
        self._tw["field_year"] = self._entry(parent, "field_year", self.year_var)
        self._tw["field_isbn"] = self._entry(parent, "field_isbn", self.isbn_var)

        self._tw["field_desc"] = ttk.Label(parent, text=self.t("field_desc"), style="Card.TLabel")
        self._tw["field_desc"].pack(anchor=tk.W, pady=(8, 2))
        self.desc_text = self._create_text_widget(parent, height=5)
        self.desc_text.pack(fill=tk.X)

        ttk.Separator(parent).pack(fill=tk.X, pady=14)
        self._tw["section_lang_doc"] = ttk.Label(parent, text=self.t("section_lang_doc"), style="Section.TLabel")
        self._tw["section_lang_doc"].pack(anchor=tk.W)
        lang_display_values = [name for name, _ in DOC_LANGUAGES]
        self.doc_lang_combo = ttk.Combobox(parent, values=lang_display_values, state="readonly")
        self.doc_lang_combo.current(0)
        self.doc_lang_combo.pack(fill=tk.X, pady=(4, 0))
        self.doc_lang_combo.bind("<<ComboboxSelected>>", self._on_doc_lang_change)

        ttk.Separator(parent).pack(fill=tk.X, pady=14)
        self._tw["section_files"] = ttk.Label(parent, text=self.t("section_files"), style="Section.TLabel")
        self._tw["section_files"].pack(anchor=tk.W)
        self._tw["field_docx"], self._tw["btn_docx"] = self._path_picker(
            parent, "field_docx", self.docx_path, self.pick_docx
        )
        self._tw["field_cover"], self._tw["btn_cover"] = self._path_picker(
            parent, "field_cover", self.cover_path, self.pick_cover
        )

        self.cover_label = ttk.Label(parent, text=self.t("no_cover_preview"), style="Card.TLabel")
        self.cover_label.pack(anchor=tk.W, pady=(8, 0))

        ttk.Separator(parent).pack(fill=tk.X, pady=14)
        self._tw["section_save"] = ttk.Label(parent, text=self.t("section_save"), style="Section.TLabel")
        self._tw["section_save"].pack(anchor=tk.W)
        self._tw["field_output"], self._tw["btn_output"] = self._path_picker(
            parent, "field_output", self.output_path, self.pick_output_path
        )

        form_button = self._add_convert_button(parent)
        form_button.pack(fill=tk.X, pady=(16, 6), ipady=7)

        actions = ttk.Frame(parent, style="Card.TFrame")
        actions.pack(fill=tk.X, pady=(8, 0))
        self.open_file_button = ttk.Button(
            actions, text=self.t("btn_open_epub"), command=self.open_epub, state=tk.DISABLED
        )
        self.open_file_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.open_folder_button = ttk.Button(
            actions, text=self.t("btn_open_folder"), command=self.open_output_folder, state=tk.DISABLED
        )
        self.open_folder_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

    def _build_preview(self, parent):
        self._tw["section_preview"] = ttk.Label(parent, text=self.t("section_preview"), style="Section.TLabel")
        self._tw["section_preview"].pack(anchor=tk.W)

        meta_frame = ttk.Frame(parent, style="Card.TFrame")
        meta_frame.pack(fill=tk.X, pady=(8, 10))
        self.result_label = ttk.Label(meta_frame, text=self.t("preview_after"), style="Card.TLabel")
        self.result_label.pack(anchor=tk.W)

        chapter_frame = ttk.Frame(parent, style="Card.TFrame")
        chapter_frame.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(chapter_frame, style="Card.TFrame")
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        self._tw["toc_label"] = ttk.Label(left, text=self.t("toc_label"), style="Section.TLabel")
        self._tw["toc_label"].pack(anchor=tk.W)
        self.chapter_list = tk.Listbox(
            left,
            width=28,
            activestyle="dotbox",
            exportselection=False,
            bg=NAVY_INPUT,
            fg=NAVY_BG,
            selectbackground=NAVY_ACCENT,
            selectforeground="#ffffff",
            highlightbackground=NAVY_BORDER,
            highlightcolor=NAVY_ACCENT,
            relief=tk.SOLID,
            borderwidth=1,
        )
        self.chapter_list.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.chapter_list.bind("<<ListboxSelect>>", self.on_chapter_select)

        text_frame = ttk.Frame(chapter_frame, style="Card.TFrame")
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.preview_text = self._create_text_widget(text_frame)
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_text.configure(yscrollcommand=scroll.set)
        self.preview_text.insert(tk.END, self.t("preview_placeholder"))
        self.preview_text.configure(state=tk.DISABLED)

        self.open_html_button = ttk.Button(
            parent, text=self.t("btn_html_preview"), command=self.open_html_preview, state=tk.DISABLED
        )
        self.open_html_button.pack(anchor=tk.E, pady=(10, 0))

    def _create_text_widget(self, parent, height=None):
        kwargs = {
            "wrap": tk.WORD,
            "relief": tk.SOLID,
            "borderwidth": 1,
            "bg": NAVY_INPUT,
            "fg": NAVY_BG,
            "insertbackground": NAVY_BG,
            "selectbackground": NAVY_ACCENT,
            "selectforeground": "#ffffff",
            "highlightbackground": NAVY_BORDER,
            "highlightcolor": NAVY_ACCENT,
            "font": ("Segoe UI", 10),
        }
        if height is not None:
            kwargs["height"] = height
        return tk.Text(parent, **kwargs)

    def _entry(self, parent, label_key: str, variable) -> ttk.Label:
        lbl = ttk.Label(parent, text=self.t(label_key), style="Card.TLabel")
        lbl.pack(anchor=tk.W, pady=(8, 2))
        ttk.Entry(parent, textvariable=variable).pack(fill=tk.X)
        return lbl

    def _path_picker(self, parent, label_key: str, variable, command) -> tuple[ttk.Label, ttk.Button]:
        lbl = ttk.Label(parent, text=self.t(label_key), style="Card.TLabel")
        lbl.pack(anchor=tk.W, pady=(8, 2))
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill=tk.X)
        ttk.Entry(row, textvariable=variable).pack(side=tk.LEFT, fill=tk.X, expand=True)
        btn = ttk.Button(row, text=self.t("btn_pick"), command=command)
        btn.pack(side=tk.LEFT, padx=(8, 0))
        return lbl, btn

    def _toggle_ui_lang(self):
        new_lang = "en" if self.ui_lang_var.get() == "pl" else "pl"
        self.ui_lang_var.set(new_lang)
        self._apply_language()

    def _apply_language(self):
        lang = self.ui_lang_var.get()
        tr = TRANSLATIONS[lang]

        self._tw["lang_toggle"].configure(text="EN" if lang == "pl" else "PL")

        label_keys = [
            "app_subtitle", "app_author_label",
            "section_metadata", "field_title", "field_author", "field_publisher",
            "field_year", "field_isbn", "field_desc", "section_lang_doc",
            "section_files", "field_docx", "field_cover",
            "section_save", "field_output",
            "section_preview", "toc_label",
        ]
        for key in label_keys:
            if key in self._tw:
                self._tw[key].configure(text=tr[key])

        for key in ("btn_docx", "btn_cover", "btn_output"):
            if key in self._tw:
                self._tw[key].configure(text=tr["btn_pick"])

        for btn in self.convert_buttons:
            btn.configure(text=tr["btn_convert"])

        self.open_file_button.configure(text=tr["btn_open_epub"])
        self.open_folder_button.configure(text=tr["btn_open_folder"])
        self.open_html_button.configure(text=tr["btn_html_preview"])

        if self.result is None:
            self.status_var.set(tr["status_ready"])
            self.result_label.configure(text=tr["preview_after"])
            self.preview_text.configure(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert(tk.END, tr["preview_placeholder"])
            self.preview_text.configure(state=tk.DISABLED)

        if not self.cover_path.get():
            self.cover_label.configure(text=tr["no_cover_preview"])

    def _on_doc_lang_change(self, _event):
        idx = self.doc_lang_combo.current()
        self.doc_lang_var.set(self._doc_lang_codes[idx])

    def pick_docx(self):
        path = filedialog.askopenfilename(
            title=self.t("dlg_pick_docx"), filetypes=[("DOCX", "*.docx")]
        )
        if path:
            self.docx_path.set(path)

    def pick_cover(self):
        path = filedialog.askopenfilename(
            title=self.t("dlg_pick_cover"), filetypes=[("Images / Obrazy", "*.png *.jpg *.jpeg")]
        )
        if path:
            self.cover_path.set(path)
            self.load_cover_preview(Path(path))

    def pick_output_path(self):
        current = self.output_path.get()
        try:
            init_dir = str(Path(current).parent) if current else str(default_output_dir())
            init_file = Path(current).name if current else "ebook.epub"
        except Exception:
            init_dir = str(default_output_dir())
            init_file = "ebook.epub"
        path = filedialog.asksaveasfilename(
            title=self.t("dlg_pick_output"),
            initialdir=init_dir,
            initialfile=init_file,
            defaultextension=".epub",
            filetypes=[("EPUB", "*.epub")],
        )
        if path:
            self.output_path.set(path)

    def load_cover_preview(self, path: Path):
        try:
            image = Image.open(path)
            image.thumbnail((160, 220))
            self.cover_preview = ImageTk.PhotoImage(image)
            self.cover_label.configure(image=self.cover_preview, text="")
        except Exception as exc:
            self.cover_label.configure(image="", text=self.t("cover_load_error") + str(exc))

    def validate_input(self) -> bool:
        missing = []
        if not self.title_var.get().strip():
            missing.append(self.t("missing_title"))
        if not self.author_var.get().strip():
            missing.append(self.t("missing_author"))
        if not self.publisher_var.get().strip():
            missing.append(self.t("missing_publisher"))
        if not self.year_var.get().strip():
            missing.append(self.t("missing_year"))
        if not self.desc_text.get("1.0", tk.END).strip():
            missing.append(self.t("missing_desc"))
        if not Path(self.docx_path.get()).is_file():
            missing.append(self.t("missing_docx"))
        if not Path(self.cover_path.get()).is_file():
            missing.append(self.t("missing_cover"))
        if not self.output_path.get().strip():
            missing.append(self.t("missing_output"))
        if missing:
            self.show_issue_window(
                self.t("err_missing_title"),
                self.t("err_missing_msg") + "\n- ".join(missing),
            )
            return False

        year = self.year_var.get().strip()
        if not re.fullmatch(r"\d{4}", year):
            self.show_issue_window(self.t("err_year_title"), self.t("err_year_msg"))
            return False

        try:
            with Image.open(Path(self.cover_path.get())) as image:
                if image.format not in {"PNG", "JPEG"}:
                    self.show_issue_window(
                        self.t("err_cover_format_title"),
                        self.t("err_cover_format_msg"),
                    )
                    return False
        except Exception as exc:
            self.show_issue_window(
                self.t("err_cover_open_title"),
                self.t("err_cover_open_msg") + str(exc),
            )
            return False

        return True

    def preflight_docx(self) -> bool:
        tr = TRANSLATIONS[self.ui_lang_var.get()]
        try:
            docx_bytes = Path(self.docx_path.get()).read_bytes()
            issues = inspect_docx_for_epub(docx_bytes, tr)
        except Exception as exc:
            self.show_issue_window(
                self.t("preflight_err_title"),
                self.t("preflight_err_msg") + str(exc),
            )
            return False

        if issues:
            self.show_issue_window(
                self.t("preflight_title"),
                self.t("preflight_msg")
                + "\n\n".join(f"{idx}. {issue}" for idx, issue in enumerate(issues, 1)),
            )
            return False
        return True

    def show_issue_window(self, title: str, message: str):
        window = tk.Toplevel(self)
        window.title(title)
        window.configure(background=NAVY_BG)
        window.geometry("620x420")
        window.minsize(520, 320)
        window.transient(self)
        window.grab_set()

        frame = ttk.Frame(window, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=title, style="Section.TLabel").pack(anchor=tk.W, pady=(0, 8))

        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        text = self._create_text_widget(text_frame)
        text.insert(tk.END, message)
        text.configure(state=tk.DISABLED)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text.configure(yscrollcommand=scroll.set)

        ttk.Button(frame, text=self.t("btn_ok"), command=window.destroy).pack(anchor=tk.E, pady=(12, 0))
        window.focus_set()

    def convert(self):
        if not self.validate_input():
            return
        if not self.preflight_docx():
            return
        self._set_convert_buttons_state(tk.DISABLED)
        self.status_var.set(self.t("converting"))
        self.result_label.configure(text=self.t("converting"))
        thread = threading.Thread(target=self._convert_worker, daemon=True)
        thread.start()

    def _set_convert_buttons_state(self, state):
        for button in self.convert_buttons:
            button.configure(state=state)

    def _convert_worker(self):
        try:
            docx_path = Path(self.docx_path.get())
            cover_path = Path(self.cover_path.get())
            doc_lang = self.doc_lang_var.get()

            saved_path = Path(self.output_path.get()).expanduser()
            if not saved_path.suffix:
                saved_path = saved_path.with_suffix(".epub")
            saved_path.parent.mkdir(parents=True, exist_ok=True)

            meta = {
                "title": self.title_var.get().strip(),
                "author": self.author_var.get().strip(),
                "publisher": self.publisher_var.get().strip() or self.author_var.get().strip(),
                "year": self.year_var.get().strip() or str(date.today().year),
                "isbn": self.isbn_var.get().strip(),
                "description": self.desc_text.get("1.0", tk.END).strip(),
            }

            chapters = parse_docx(docx_path.read_bytes())
            if not chapters:
                raise ValueError(self.t("err_no_chapters"))

            cover_bytes = cover_path.read_bytes()
            cover_ext = cover_path.suffix.lower()
            epub_bytes = build_epub(meta, chapters, cover_bytes, cover_ext, lang=doc_lang)
            saved_path.write_bytes(epub_bytes)
            valid, messages = validate_epub(epub_bytes)

            result = ConversionResult(meta, chapters, epub_bytes, cover_bytes, cover_ext, saved_path, valid, messages)
            preview_html_path = self.write_preview_html(result)
            self.after(0, lambda: self.show_result(result, preview_html_path))
        except Exception as exc:
            self.after(0, lambda: self.show_error(exc))

    def write_preview_html(self, result: ConversionResult) -> Path:
        parts = [
            "<!doctype html><html lang='pl'><head><meta charset='utf-8'><title>Podglad EPUB</title>",
            "<style>" + CSS + "</style></head><body>",
            "<h1>" + xml_escape(result.meta["title"]) + "</h1>",
            "<p><strong>Autor:</strong> " + xml_escape(result.meta["author"]) + "</p>",
        ]
        for _cid, title, _epub_type, _role, body in result.chapters:
            parts.append("<hr><h1>" + title + "</h1>" + body)
        parts.append("</body></html>")
        path = Path(tempfile.gettempdir()) / "docx_epub_converter_preview.html"
        path.write_text("\n".join(parts), encoding="utf-8")
        return path

    def show_result(self, result: ConversionResult, preview_html_path: Path):
        self.result = result
        self.preview_html_path = preview_html_path
        self._set_convert_buttons_state(tk.NORMAL)
        self.open_file_button.configure(state=tk.NORMAL)
        self.open_folder_button.configure(state=tk.NORMAL)
        self.open_html_button.configure(state=tk.NORMAL)

        lang = self.ui_lang_var.get()
        validation = (
            "EpubCheck: OK"
            if result.valid
            else f"EpubCheck: {len(result.messages)} " + ("message(s)" if lang == "en" else "komunikat(ow)")
        )
        self.status_var.set(self.t("status_done") + str(result.saved_path))
        self.result_label.configure(text=validation + "\n" + ("File: " if lang == "en" else "Plik: ") + str(result.saved_path))

        self.chapter_list.delete(0, tk.END)
        for index, (_cid, title, _epub_type, _role, _body) in enumerate(result.chapters, 1):
            self.chapter_list.insert(tk.END, f"{index}. {html.unescape(title)}")
        self.chapter_list.selection_set(0)
        self.update_preview(0)

        if not result.valid:
            details = "\n".join(str(msg) for msg in result.messages[:10])
            messagebox.showwarning(APP_TITLE, self.t("warn_epubcheck") + details)
        else:
            messagebox.showinfo(APP_TITLE, self.t("info_saved") + str(result.saved_path))

    def show_error(self, exc: Exception):
        self._set_convert_buttons_state(tk.NORMAL)
        self.status_var.set(self.t("status_error"))
        self.result_label.configure(text=("Error: " if self.ui_lang_var.get() == "en" else "Blad: ") + str(exc))
        messagebox.showerror(APP_TITLE, str(exc))

    def on_chapter_select(self, _event):
        selection = self.chapter_list.curselection()
        if selection:
            self.update_preview(selection[0])

    def update_preview(self, index: int):
        if not self.result:
            return
        _cid, title, _epub_type, _role, body = self.result.chapters[index]
        content = html.unescape(title) + "\n" + ("=" * len(html.unescape(title))) + "\n\n" + html_to_plain_text(body)
        self.preview_text.configure(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, content)
        self.preview_text.configure(state=tk.DISABLED)

    def open_epub(self):
        if self.result:
            open_path(self.result.saved_path)

    def open_output_folder(self):
        if self.result:
            open_path(self.result.saved_path.parent)

    def open_html_preview(self):
        if self.preview_html_path and self.preview_html_path.exists():
            webbrowser.open(self.preview_html_path.as_uri())


def main():
    app = EpubConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
