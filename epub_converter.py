"""
DOCX EPUB Converter v2.0
EAA / EPUB Accessibility 1.1 / WCAG 2.1 AA
Autor: Marek Zettel
"""

from __future__ import annotations

import base64
import html
import io
import json
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
from docx.oxml.ns import qn as _docx_qn
from epubcheck import EpubCheck
from PIL import Image, ImageTk

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except ImportError:
    _HAS_DND = False

_BaseApp = TkinterDnD.Tk if _HAS_DND else tk.Tk  # type: ignore[misc]

APP_TITLE = "DOCX EPUB Converter"
APP_VERSION = "2.0"
APP_AUTHOR = "Marek Zettel"
NAVY_BG = "#071a33"
NAVY_PANEL = "#0f2747"
NAVY_PANEL_ALT = "#14365f"
NAVY_BORDER = "#27527f"
NAVY_INPUT = "#f4f8ff"
NAVY_TEXT = "#edf5ff"
NAVY_ACCENT = "#2f80d0"
NAVY_ACCENT_ACTIVE = "#4aa3ff"

# OOXML tag constants (computed once at module load)
_W_P         = _docx_qn("w:p")
_W_TBL       = _docx_qn("w:tbl")
_W_R         = _docx_qn("w:r")
_W_T         = _docx_qn("w:t")
_W_HYPERLINK = _docx_qn("w:hyperlink")
_W_B         = _docx_qn("w:b")
_W_I         = _docx_qn("w:i")
_W_U         = _docx_qn("w:u")
_W_STRIKE    = _docx_qn("w:strike")
_W_RPR       = _docx_qn("w:rPr")
_A_BLIP  = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
_R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
_R_HYP          = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_W_FOOTNOTE_REF = _docx_qn("w:footnoteReference")
_W_FN_ID        = _docx_qn("w:id")
_FOOTNOTES_RT   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"

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

TOC_TITLES: dict[str, str] = {
    "pl": "Spis treści",
    "en": "Table of Contents",
    "de": "Inhaltsverzeichnis",
    "fr": "Table des matières",
    "es": "Tabla de contenidos",
    "it": "Indice",
    "cs": "Obsah",
    "sk": "Obsah",
    "hu": "Tartalomjegyzék",
    "ru": "Оглавление",
    "uk": "Зміст",
    "nl": "Inhoudsopgave",
    "pt": "Índice",
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
        "toc_title": "Spis treści",
        "preview_html_title": "Podglad EPUB",
        "author_label": "Autor",
        "err_isbn_title": "Niepoprawny ISBN",
        "err_isbn_msg": "ISBN musi zawierac dokladnie 10 lub 13 cyfr (bez myslnikow).",
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
            "     Potrzebne: nazwa, adres, NIP lub PESEL, e-mail.\n"
            "  3. Zloz wniosek o pule numerow (min. 10 sztuk).\n"
            "  4. Konto aktywowane zwykle w 1 dzien roboczy;\n"
            "     numery przyznawane w kilka dni roboczych.\n"
            "  5. W panelu e-ISBN przypisz konkretny numer do swojej publikacji.\n\n"
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
        "toc_title": "Table of Contents",
        "preview_html_title": "EPUB Preview",
        "author_label": "Author",
        "err_isbn_title": "Invalid ISBN",
        "err_isbn_msg": "ISBN must be exactly 10 or 13 digits (no hyphens).",
        "isbn_help_title": "What is ISBN and how to get one?",
        "isbn_help_text": (
            "ISBN (International Standard Book Number) is a free, unique identifier\n"
            "for publications. Each format (e-book, print, audiobook) needs its own number.\n\n"
            "Format in this app: 10 or 13 digits, no hyphens or spaces.\n"
            "Example: 9788301234567\n\n"
            "How to get an ISBN in Poland (free of charge):\n"
            "  1. Go to e-isbn.pl (National Library of Poland).\n"
            "  2. Create a publisher account — individuals are welcome.\n"
            "     Required: name, address, tax ID or national ID, e-mail.\n"
            "  3. Apply for a block of numbers (minimum 10).\n"
            "  4. Account activated usually within 1 business day;\n"
            "     numbers assigned within a few business days.\n"
            "  5. In the e-ISBN panel, assign a specific number to your publication.\n\n"
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


def xml_escape(text: str) -> str:
    return html.escape(text or "", quote=False)


# ---------------------------------------------------------------------------
# DOCX → HTML conversion helpers
# ---------------------------------------------------------------------------

def _run_elem_to_html(r_elem, images: dict[str, bytes], part) -> str:
    """Convert a single <w:r> XML element to HTML, extracting images if present."""
    blips = r_elem.findall(f".//{_A_BLIP}")
    if blips:
        r_embed = blips[0].get(_R_EMBED)
        if r_embed and r_embed in part.rels:
            try:
                img_part = part.rels[r_embed].target_part
                blob = img_part.blob
                ct = img_part.content_type
                ext = ".png" if "png" in ct else ".jpg"
                img_id = f"img{len(images) + 1:03d}{ext}"
                images[img_id] = blob
                return f'<img src="../images/{img_id}" alt=""/>'
            except Exception:
                pass
        return ""

    t_elems = r_elem.findall(_W_T)
    text = xml_escape("".join((t.text or "") for t in t_elems))
    if not text:
        return ""

    rPr = r_elem.find(_W_RPR)
    bold      = rPr is not None and rPr.find(_W_B)      is not None
    italic    = rPr is not None and rPr.find(_W_I)      is not None
    underline = rPr is not None and rPr.find(_W_U)      is not None
    strike    = rPr is not None and rPr.find(_W_STRIKE) is not None

    if bold and italic:
        text = f"<strong><em>{text}</em></strong>"
    elif bold:
        text = f"<strong>{text}</strong>"
    elif italic:
        text = f"<em>{text}</em>"
    if underline:
        text = f"<u>{text}</u>"
    if strike:
        text = f"<s>{text}</s>"
    return text


def para_to_html(
    para,
    images: dict[str, bytes],
    fn_refs: list | None = None,
    footnotes: dict[str, str] | None = None,
) -> str:
    """Convert a paragraph to HTML, handling formatting, inline images, hyperlinks, and footnotes."""
    parts = []
    part = para.part
    for child in para._p:
        if child.tag == _W_R:
            ref_elem = child.find(f".//{_W_FOOTNOTE_REF}")
            if ref_elem is not None:
                fn_id = ref_elem.get(_W_FN_ID, "")
                if fn_refs is not None and footnotes is not None and fn_id in footnotes:
                    fn_refs.append((fn_id, footnotes[fn_id]))
                parts.append(f'<sup><a epub:type="noteref" href="#fn-{fn_id}">[{fn_id}]</a></sup>')
            else:
                parts.append(_run_elem_to_html(child, images, part))
        elif child.tag == _W_HYPERLINK:
            r_id = child.get(_R_HYP)
            href = ""
            if r_id and r_id in part.rels:
                try:
                    href = xml_escape(part.rels[r_id].target_ref or "")
                except Exception:
                    pass
            inner = "".join(_run_elem_to_html(r, images, part) for r in child.findall(_W_R))
            if inner:
                parts.append(f'<a href="{href}">{inner}</a>' if href else inner)
    return "".join(parts)


def table_to_html(table) -> str:
    rows = ["<table>"]
    for i, row in enumerate(table.rows):
        rows.append("<tr>")
        for cell in row.cells:
            tag = "th" if i == 0 else "td"
            rows.append(f"  <{tag}>{xml_escape(cell.text.strip())}</{tag}>")
        rows.append("</tr>")
    rows.append("</table>")
    return "\n".join(rows)


def extract_footnotes(doc) -> dict[str, str]:
    """Return {id_str: plain_text} for all numbered footnotes in the document."""
    try:
        fn_part = doc.part.part_related_by(_FOOTNOTES_RT)
        result = {}
        for fn in fn_part._element.findall(_docx_qn("w:footnote")):
            fn_id = fn.get(_W_FN_ID, "")
            if fn_id in ("-1", "0"):
                continue
            texts = [xml_escape(t.text) for t in fn.findall(f".//{_W_T}") if t.text]
            result[fn_id] = "".join(texts)
        return result
    except Exception:
        return {}


def parse_docx(docx_bytes: bytes) -> tuple[list, dict[str, bytes]]:
    """Parse DOCX and return (chapters, images).

    chapters: list of (id, title, epub_type, role, body_html)
    images:   dict of filename → bytes for all inline images
    """
    doc = Document(io.BytesIO(docx_bytes))
    footnotes = extract_footnotes(doc)
    chapters: list = []
    images: dict[str, bytes] = {}
    cur_id: str | None = None
    cur_title: str | None = None
    cur_body: list[str] = []
    cur_fn_refs: list = []
    in_ul = False
    in_ol = False
    chapter_counter = 0

    front_kw = {"wstep", "wstęp", "przedmowa", "wprowadzenie", "preface", "introduction", "foreword", "prolog"}
    back_kw  = {"zakonczenie", "zakończenie", "epilog", "podsumowanie", "conclusion", "afterword", "epilogue"}

    para_map  = {id(p._p): p for p in doc.paragraphs}
    table_map = {id(t._tbl): t for t in doc.tables}

    def close_list(parts: list[str]) -> None:
        nonlocal in_ul, in_ol
        if in_ul:
            parts.append("</ul>")
            in_ul = False
        if in_ol:
            parts.append("</ol>")
            in_ol = False

    def flush_chapter() -> None:
        if cur_title is None:
            return
        close_list(cur_body)
        if cur_fn_refs:
            seen: set = set()
            items: list[str] = []
            for fn_id, fn_text in cur_fn_refs:
                if fn_id not in seen:
                    seen.add(fn_id)
                    items.append(
                        f'<aside id="fn-{fn_id}" epub:type="footnote" role="doc-footnote">'
                        f"<p><sup>[{fn_id}]</sup> {fn_text}</p></aside>"
                    )
            cur_body.append(
                '<section class="footnotes-section">\n' + "\n".join(items) + "\n</section>"
            )
        key = html.unescape(cur_title).lower().strip()
        if key in front_kw:
            epub_type, role = "frontmatter", "doc-preface"
        elif key in back_kw:
            epub_type, role = "backmatter", "doc-conclusion"
        else:
            epub_type, role = "chapter", "doc-chapter"
        chapters.append((cur_id, cur_title, epub_type, role, "\n".join(cur_body)))

    for child in doc.element.body:
        if child.tag == _W_P:
            para = para_map.get(id(child))
            if para is None:
                continue

            style_name = para.style.name if para.style else "Normal"
            text = para.text.strip()

            is_h1     = any(n in style_name for n in ["Heading 1", "Nagłówek 1"])
            is_h2     = any(n in style_name for n in ["Heading 2", "Nagłówek 2"])
            is_h3     = any(n in style_name for n in ["Heading 3", "Nagłówek 3"])
            is_quote  = any(n in style_name for n in ["Quote", "Cytat", "Blockquote", "Intense Quote", "Intensywny cytat"])
            is_bullet = any(n in style_name for n in ["List Bullet", "Lista wypunktowana", "List Paragraph"])
            is_number = any(n in style_name for n in ["List Number", "Lista numerowana"])

            inner    = para_to_html(para, images, cur_fn_refs if cur_title else None, footnotes)
            fallback = xml_escape(text)

            if is_h1:
                flush_chapter()
                chapter_counter += 1
                cur_id    = "ch" + str(chapter_counter).zfill(2)
                cur_title = xml_escape(text) if text else "Rozdzial"
                cur_body  = []
                cur_fn_refs = []
                in_ul = in_ol = False
            elif cur_title is None:
                continue
            elif is_h2:
                close_list(cur_body)
                cur_body.append("<h2>" + (inner or fallback) + "</h2>")
            elif is_h3:
                close_list(cur_body)
                cur_body.append("<h3>" + (inner or fallback) + "</h3>")
            elif is_quote:
                close_list(cur_body)
                cur_body.append("<blockquote><p>" + (inner or fallback) + "</p></blockquote>")
            elif is_bullet:
                if not in_ul:
                    if in_ol:
                        cur_body.append("</ol>")
                        in_ol = False
                    cur_body.append("<ul>")
                    in_ul = True
                cur_body.append("<li>" + (inner or fallback) + "</li>")
            elif is_number:
                if not in_ol:
                    if in_ul:
                        cur_body.append("</ul>")
                        in_ul = False
                    cur_body.append("<ol>")
                    in_ol = True
                cur_body.append("<li>" + (inner or fallback) + "</li>")
            elif text or inner:
                close_list(cur_body)
                cur_body.append("<p>" + (inner or fallback) + "</p>")
            else:
                close_list(cur_body)

        elif child.tag == _W_TBL and cur_title is not None:
            table = table_map.get(id(child))
            if table is not None:
                close_list(cur_body)
                cur_body.append(table_to_html(table))

    flush_chapter()
    return chapters, images


def inspect_docx_for_epub(docx_bytes: bytes, tr: dict | None = None) -> list[str]:
    if tr is None:
        tr = TRANSLATIONS["pl"]
    issues = []
    doc = Document(io.BytesIO(docx_bytes))

    heading1_count = 0
    seen_heading1 = False
    nonempty_before_first_heading: list[str] = []
    current_chapter_title = ""
    current_chapter_has_body = False
    empty_chapters: list[str] = []
    last_heading_level = 0

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        text = para.text.strip()
        if not text:
            continue

        is_h1 = any(n in style_name for n in ["Heading 1", "Nagłówek 1"])
        is_h2 = any(n in style_name for n in ["Heading 2", "Nagłówek 2"])
        is_h3 = any(n in style_name for n in ["Heading 3", "Nagłówek 3"])

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
        issues.append(tr["issue_empty_chapters"] + ", ".join(empty_chapters[:5]))

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
    "strong{font-weight:bold}em{font-style:italic}u{text-decoration:underline}s{text-decoration:line-through}\n"
    "table{border-collapse:collapse;width:100%;margin:1em 0}\n"
    "th,td{border:1px solid #ccc;padding:.4em .6em;text-align:left}\n"
    "th{background:#f0f4f8;font-weight:bold}\n"
    "img{max-width:100%;height:auto;display:block;margin:.5em auto}\n"
    "a{color:#2a5298}\n"
    ".footnotes-section{border-top:1px solid #ccc;margin-top:2em;padding-top:.5em;"
    "font-size:.85em;color:#444}\n"
    ".footnotes-section aside{margin:.4em 0}\n"
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
    toc_title = TOC_TITLES.get(lang, "Table of Contents")
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


def make_opf(
    meta: dict,
    chapters,
    images: dict[str, bytes],
    cover_mime: str,
    cover_ext: str,
    lang: str = "pl",
) -> bytes:
    manifest_items = (
        '    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>\n'
        '    <item id="css" href="styles/main.css" media-type="text/css"/>\n'
        '    <item id="cover-image" href="images/cover' + cover_ext + '" media-type="' + cover_mime + '" properties="cover-image"/>\n'
        '    <item id="cover-page" href="content/cover_page.xhtml" media-type="application/xhtml+xml"/>\n'
    )
    for img_name in images:
        ct = "image/png" if img_name.endswith(".png") else "image/jpeg"
        img_id = re.sub(r"[^a-z0-9]", "-", img_name)
        manifest_items += f'    <item id="{img_id}" href="images/{img_name}" media-type="{ct}"/>\n'

    spine = '    <itemref idref="cover-page" linear="yes"/>\n'
    for cid, _title, _epub_type, _role, _body in chapters:
        manifest_items += '    <item id="' + cid + '" href="content/' + cid + '.xhtml" media-type="application/xhtml+xml"/>\n'
        spine += '    <itemref idref="' + cid + '"/>\n'

    today = date.today().isoformat()
    _A11Y_PL = (
        "Publikacja dostepna cyfrowo. Semantyczna struktura HTML, nawigacyjny spis tresci, "
        "teksty alternatywne, kompletne metadane. Zgodna z EPUB Accessibility 1.1 i WCAG 2.1 AA."
    )
    _A11Y_EN = (
        "Digitally accessible publication. Semantic HTML structure, navigational table of contents, "
        "alternative texts, complete metadata. Compliant with EPUB Accessibility 1.1 and WCAG 2.1 AA."
    )
    # EN is used as a neutral fallback for all non-PL languages (DE, FR, ES, IT, …)
    a11y_summary = _A11Y_PL if lang == "pl" else _A11Y_EN
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


def build_epub(
    meta: dict,
    chapters,
    images: dict[str, bytes],
    cover_bytes: bytes,
    cover_ext: str,
    lang: str = "pl",
) -> bytes:
    cover_mime = "image/jpeg" if cover_ext.lower() in (".jpg", ".jpeg") else "image/png"
    # EN as fallback for all non-PL languages (no per-language cover text defined)
    cover_alt = (
        "Okladka: " + meta["title"] + ", autor " + meta["author"]
        if lang == "pl"
        else "Cover: " + meta["title"] + ", " + meta["author"]
    )

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        mimetype = zipfile.ZipInfo("mimetype")
        mimetype.compress_type = zipfile.ZIP_STORED
        zf.writestr(mimetype, "application/epub+zip")
        zf.writestr("META-INF/container.xml", make_container())
        zf.writestr("OEBPS/content.opf", make_opf(meta, chapters, images, cover_mime, cover_ext, lang))
        zf.writestr("OEBPS/nav.xhtml", make_nav(chapters, lang))
        zf.writestr("OEBPS/styles/main.css", e(CSS))
        zf.writestr("OEBPS/images/cover" + cover_ext, cover_bytes)
        zf.writestr("OEBPS/content/cover_page.xhtml", make_cover_xhtml(cover_alt, cover_ext, lang))
        for img_name, img_bytes in images.items():
            zf.writestr("OEBPS/images/" + img_name, img_bytes)
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


_CONFIG_FILE = Path.home() / ".epub_converter_config.json"


def safe_filename(title: str) -> str:
    name = re.sub(r"[^\w\-]+", "-", title.lower(), flags=re.UNICODE).strip("-")
    return (name or "ebook") + ".epub"


def default_output_dir() -> Path:
    downloads = Path.home() / "Downloads"
    return downloads if downloads.exists() else Path.home()


def html_to_plain_text(fragment: str) -> str:
    text = re.sub(r"</(p|h1|h2|h3|li|blockquote|tr)>", "\n", fragment)
    text = re.sub(r"</(table|ul|ol)>", "\n", text)
    text = re.sub(r"<li>", "- ", text)
    text = re.sub(r"<t[hd][^>]*>", "  ", text)
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
    images: dict
    epub_bytes: bytes
    cover_bytes: bytes
    cover_ext: str
    saved_path: Path
    valid: bool
    messages: list


class EpubConverterApp(_BaseApp):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("1120x760")
        self.minsize(900, 620)

        self.ui_lang_var  = tk.StringVar(value="pl")
        self.doc_lang_var = tk.StringVar(value="pl")

        self.docx_path   = tk.StringVar()
        self.cover_path  = tk.StringVar()
        self.output_path = tk.StringVar(value=str(default_output_dir() / "ebook.epub"))
        self.title_var     = tk.StringVar()
        self.author_var    = tk.StringVar()
        self.publisher_var = tk.StringVar()
        self.year_var      = tk.StringVar(value=str(date.today().year))
        self.isbn_var      = tk.StringVar()
        self.status_var    = tk.StringVar(value=TRANSLATIONS["pl"]["status_ready"])

        self.cover_preview: ImageTk.PhotoImage | None = None
        self.result: ConversionResult | None = None
        self.preview_html_path: Path | None = None
        self.convert_buttons: list[ttk.Button] = []
        self._tw: dict[str, tk.Widget] = {}
        self._doc_lang_codes: list[str] = [code for _, code in DOC_LANGUAGES]
        self._output_manually_set = False

        _cfg = self._load_config()
        self._last_docx_dir:  str = _cfg.get("last_docx_dir", "")
        self._last_cover_dir: str = _cfg.get("last_cover_dir", "")
        self._profile_doc_lang: str = _cfg.get("profile_doc_lang", "pl")

        if _cfg.get("profile_author"):
            self.author_var.set(_cfg["profile_author"])
        if _cfg.get("profile_publisher"):
            self.publisher_var.set(_cfg["profile_publisher"])
        if _cfg.get("profile_year"):
            self.year_var.set(_cfg["profile_year"])

        self.title_var.trace_add("write", self._on_title_change)

        self._configure_style()
        self._build_ui()

    # ------------------------------------------------------------------
    # Config persistence
    # ------------------------------------------------------------------

    def _load_config(self) -> dict:
        try:
            return json.loads(_CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_config(self, data: dict) -> None:
        try:
            _CONFIG_FILE.write_text(json.dumps(data), encoding="utf-8")
        except Exception:
            pass

    def _update_config(self, updates: dict) -> None:
        cfg = self._load_config()
        cfg.update(updates)
        self._save_config(cfg)

    def _save_profile(self) -> None:
        self._update_config({
            "profile_author":    self.author_var.get().strip(),
            "profile_publisher": self.publisher_var.get().strip(),
            "profile_year":      self.year_var.get().strip(),
            "profile_doc_lang":  self.doc_lang_var.get(),
        })

    # ------------------------------------------------------------------
    # Translations
    # ------------------------------------------------------------------

    def t(self, key: str) -> str:
        return TRANSLATIONS[self.ui_lang_var.get()].get(key, key)

    # ------------------------------------------------------------------
    # Styles
    # ------------------------------------------------------------------

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
        style.configure(
            "TProgressbar",
            troughcolor=NAVY_BG,
            background=NAVY_ACCENT,
            bordercolor=NAVY_BORDER,
            lightcolor=NAVY_ACCENT,
            darkcolor=NAVY_ACCENT,
        )

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        canvas = tk.Canvas(self, bg=NAVY_BG, highlightthickness=0)
        vscroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        root = ttk.Frame(canvas, padding=16)
        _root_id = canvas.create_window((0, 0), window=root, anchor=tk.NW)

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(_root_id, width=event.width)

        def _on_mousewheel(event):
            if not isinstance(event.widget, (tk.Text, tk.Listbox)):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        root.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        header = ttk.Frame(root)
        header.pack(fill=tk.X, pady=(0, 12))

        title_group = ttk.Frame(header)
        title_group.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(title_group, text=f"{APP_TITLE} v{APP_VERSION}", style="Title.TLabel").pack(anchor=tk.W)
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

        left  = ttk.Frame(main, style="Card.TFrame", padding=14)
        right = ttk.Frame(main, style="Card.TFrame", padding=14)
        main.add(left, weight=1)
        main.add(right, weight=2)

        self._build_form(left)
        self._build_preview(right)

        bottom = ttk.Frame(root)
        bottom.pack(fill=tk.X, pady=(12, 0))
        self._add_convert_button(bottom, side=tk.RIGHT, ipadx=24, ipady=7)

        self.progress = ttk.Progressbar(root, mode="indeterminate", length=100)
        self.progress.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(root, textvariable=self.status_var, style="Status.TLabel").pack(fill=tk.X, pady=(4, 0))

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

        self._tw["field_title"]     = self._entry(parent, "field_title",     self.title_var)
        self._tw["field_author"]    = self._entry(parent, "field_author",    self.author_var)
        self._tw["field_publisher"] = self._entry(parent, "field_publisher", self.publisher_var)
        self._tw["field_year"]      = self._entry(parent, "field_year",      self.year_var)
        _isbn_hdr = ttk.Frame(parent, style="Card.TFrame")
        _isbn_hdr.pack(fill=tk.X, pady=(8, 2))
        self._tw["field_isbn"] = ttk.Label(_isbn_hdr, text=self.t("field_isbn"), style="Card.TLabel")
        self._tw["field_isbn"].pack(side=tk.LEFT)
        self._tw["isbn_help_btn"] = ttk.Button(
            _isbn_hdr, text="?", width=2, style="Lang.TButton", command=self._show_isbn_help
        )
        self._tw["isbn_help_btn"].pack(side=tk.RIGHT)
        ttk.Entry(parent, textvariable=self.isbn_var).pack(fill=tk.X)

        self._tw["field_desc"] = ttk.Label(parent, text=self.t("field_desc"), style="Card.TLabel")
        self._tw["field_desc"].pack(anchor=tk.W, pady=(8, 2))
        self.desc_text = self._create_text_widget(parent, height=5)
        self.desc_text.pack(fill=tk.X)

        ttk.Separator(parent).pack(fill=tk.X, pady=14)
        self._tw["section_lang_doc"] = ttk.Label(parent, text=self.t("section_lang_doc"), style="Section.TLabel")
        self._tw["section_lang_doc"].pack(anchor=tk.W)
        lang_display_values = [name for name, _ in DOC_LANGUAGES]
        self.doc_lang_combo = ttk.Combobox(parent, values=lang_display_values, state="readonly")
        saved_lang = getattr(self, "_profile_doc_lang", "pl")
        if saved_lang in self._doc_lang_codes:
            self.doc_lang_combo.current(self._doc_lang_codes.index(saved_lang))
            self.doc_lang_var.set(saved_lang)
        else:
            self.doc_lang_combo.current(0)
        self.doc_lang_combo.pack(fill=tk.X, pady=(4, 0))
        self.doc_lang_combo.bind("<<ComboboxSelected>>", self._on_doc_lang_change)

        ttk.Separator(parent).pack(fill=tk.X, pady=14)
        self._tw["section_files"] = ttk.Label(parent, text=self.t("section_files"), style="Section.TLabel")
        self._tw["section_files"].pack(anchor=tk.W)
        self._tw["field_docx"], self._tw["btn_docx"] = self._path_picker(
            parent, "field_docx", self.docx_path, self.pick_docx, dnd_handler=self._dnd_docx
        )
        self._tw["field_cover"], self._tw["btn_cover"] = self._path_picker(
            parent, "field_cover", self.cover_path, self.pick_cover, dnd_handler=self._dnd_cover
        )

        self.cover_label = ttk.Label(parent, text=self.t("no_cover_preview"), style="Card.TLabel")
        self.cover_label.pack(anchor=tk.W, pady=(8, 0))
        self.cover_warn_label = ttk.Label(parent, text="", style="Card.TLabel", foreground="#e07000")
        self.cover_warn_label.pack(anchor=tk.W)

        ttk.Separator(parent).pack(fill=tk.X, pady=14)
        self._tw["section_save"] = ttk.Label(parent, text=self.t("section_save"), style="Section.TLabel")
        self._tw["section_save"].pack(anchor=tk.W)
        self._tw["field_output"], self._tw["btn_output"] = self._path_picker(
            parent, "field_output", self.output_path, self.pick_output_path
        )

        form_button = self._add_convert_button(parent)
        form_button.pack(fill=tk.X, pady=(16, 2), ipady=7)
        self._tw["btn_batch"] = ttk.Button(parent, text=self.t("btn_batch"), command=self.open_batch_dialog)
        self._tw["btn_batch"].pack(fill=tk.X, pady=(0, 6))

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

    def _path_picker(
        self, parent, label_key: str, variable, command, dnd_handler=None
    ) -> tuple[ttk.Label, ttk.Button]:
        lbl = ttk.Label(parent, text=self.t(label_key), style="Card.TLabel")
        lbl.pack(anchor=tk.W, pady=(8, 2))
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill=tk.X)
        entry = ttk.Entry(row, textvariable=variable)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if _HAS_DND and dnd_handler is not None:
            entry.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
            entry.dnd_bind("<<Drop>>", dnd_handler)  # type: ignore[attr-defined]
        btn = ttk.Button(row, text=self.t("btn_pick"), command=command)
        btn.pack(side=tk.LEFT, padx=(8, 0))
        return lbl, btn

    # ------------------------------------------------------------------
    # Language switching
    # ------------------------------------------------------------------

    def _toggle_ui_lang(self):
        self.ui_lang_var.set("en" if self.ui_lang_var.get() == "pl" else "pl")
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
            "section_preview", "toc_label", "btn_batch",
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
        self.doc_lang_var.set(self._doc_lang_codes[self.doc_lang_combo.current()])

    # ------------------------------------------------------------------
    # Auto output filename from title
    # ------------------------------------------------------------------

    def _show_isbn_help(self):
        self.show_issue_window(self.t("isbn_help_title"), self.t("isbn_help_text"))

    def _on_title_change(self, *_):
        if self._output_manually_set:
            return
        title = self.title_var.get().strip()
        if not title:
            return
        current = self.output_path.get()
        try:
            out_dir = Path(current).parent
        except Exception:
            out_dir = default_output_dir()
        self.output_path.set(str(out_dir / safe_filename(title)))

    # ------------------------------------------------------------------
    # File pickers & DnD
    # ------------------------------------------------------------------

    def pick_docx(self):
        init_dir = self._last_docx_dir or str(default_output_dir())
        path = filedialog.askopenfilename(
            title=self.t("dlg_pick_docx"),
            initialdir=init_dir,
            filetypes=[("DOCX", "*.docx")],
        )
        if path:
            self._set_docx(path)

    def _dnd_docx(self, event):
        path = event.data.strip().strip("{}")
        self._set_docx(path)

    def _set_docx(self, path: str):
        self.docx_path.set(path)
        self._last_docx_dir = str(Path(path).parent)
        self._update_config({"last_docx_dir": self._last_docx_dir})

    def pick_cover(self):
        init_dir = self._last_cover_dir or self._last_docx_dir or str(default_output_dir())
        path = filedialog.askopenfilename(
            title=self.t("dlg_pick_cover"),
            initialdir=init_dir,
            filetypes=[("Images / Obrazy", "*.png *.jpg *.jpeg")],
        )
        if path:
            self._set_cover(path)

    def _dnd_cover(self, event):
        path = event.data.strip().strip("{}")
        self._set_cover(path)

    def _set_cover(self, path: str):
        self.cover_path.set(path)
        self._last_cover_dir = str(Path(path).parent)
        self._update_config({"last_cover_dir": self._last_cover_dir})
        self.load_cover_preview(Path(path))

    def pick_output_path(self):
        current = self.output_path.get()
        try:
            init_dir  = str(Path(current).parent) if current else str(default_output_dir())
            init_file = Path(current).name if current else "ebook.epub"
        except Exception:
            init_dir  = str(default_output_dir())
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
            self._output_manually_set = True

    def load_cover_preview(self, path: Path):
        try:
            image = Image.open(path)
            orig_w, orig_h = image.size
            image.thumbnail((160, 220))
            self.cover_preview = ImageTk.PhotoImage(image)
            self.cover_label.configure(image=self.cover_preview, text="")
            warnings = []
            if orig_h > 0 and abs(orig_w / orig_h - 2 / 3) > 0.10:
                warnings.append(self.t("cover_warn_ratio"))
            if orig_w < 1200:
                warnings.append(self.t("cover_warn_res"))
            self.cover_warn_label.configure(text="\n".join(warnings))
        except Exception as exc:
            self.cover_label.configure(image="", text=self.t("cover_load_error") + str(exc))
            self.cover_warn_label.configure(text="")

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------

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

        if not re.fullmatch(r"\d{4}", self.year_var.get().strip()):
            self.show_issue_window(self.t("err_year_title"), self.t("err_year_msg"))
            return False

        isbn = self.isbn_var.get().strip()
        if isbn and not re.fullmatch(r"\d{10}(\d{3})?", isbn):
            self.show_issue_window(self.t("err_isbn_title"), self.t("err_isbn_msg"))
            return False

        try:
            with Image.open(Path(self.cover_path.get())) as image:
                if image.format not in {"PNG", "JPEG"}:
                    self.show_issue_window(self.t("err_cover_format_title"), self.t("err_cover_format_msg"))
                    return False
        except Exception as exc:
            self.show_issue_window(self.t("err_cover_open_title"), self.t("err_cover_open_msg") + str(exc))
            return False

        return True

    def preflight_docx(self) -> bool:
        tr = TRANSLATIONS[self.ui_lang_var.get()]
        try:
            docx_bytes = Path(self.docx_path.get()).read_bytes()
            issues = inspect_docx_for_epub(docx_bytes, tr)
        except Exception as exc:
            self.show_issue_window(self.t("preflight_err_title"), self.t("preflight_err_msg") + str(exc))
            return False
        if issues:
            self.show_issue_window(
                self.t("preflight_title"),
                self.t("preflight_msg") + "\n\n".join(f"{i}. {issue}" for i, issue in enumerate(issues, 1)),
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

        btn_row = ttk.Frame(frame)
        btn_row.pack(anchor=tk.E, pady=(12, 0))

        def _copy_to_clipboard():
            window.clipboard_clear()
            window.clipboard_append(message)
            window.update()

        ttk.Button(btn_row, text=self.t("btn_copy"), command=_copy_to_clipboard).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text=self.t("btn_ok"), command=window.destroy).pack(side=tk.LEFT)
        window.focus_set()

    # ------------------------------------------------------------------
    # Conversion
    # ------------------------------------------------------------------

    def convert(self):
        if not self.validate_input():
            return
        if not self.preflight_docx():
            return
        self._set_convert_buttons_state(tk.DISABLED)
        self.status_var.set(self.t("converting"))
        self.result_label.configure(text=self.t("converting"))
        self.progress.start(10)
        threading.Thread(target=self._convert_worker, daemon=True).start()

    def _set_convert_buttons_state(self, state):
        for button in self.convert_buttons:
            button.configure(state=state)

    def _convert_worker(self):
        try:
            docx_path  = Path(self.docx_path.get())
            cover_path = Path(self.cover_path.get())
            doc_lang   = self.doc_lang_var.get()

            saved_path = Path(self.output_path.get()).expanduser()
            if not saved_path.suffix:
                saved_path = saved_path.with_suffix(".epub")
            saved_path.parent.mkdir(parents=True, exist_ok=True)

            meta = {
                "title":       self.title_var.get().strip(),
                "author":      self.author_var.get().strip(),
                "publisher":   self.publisher_var.get().strip() or self.author_var.get().strip(),
                "year":        self.year_var.get().strip() or str(date.today().year),
                "isbn":        self.isbn_var.get().strip(),
                "description": self.desc_text.get("1.0", tk.END).strip(),
            }

            chapters, images = parse_docx(docx_path.read_bytes())
            if not chapters:
                raise ValueError(self.t("err_no_chapters"))

            cover_bytes = cover_path.read_bytes()
            cover_ext   = cover_path.suffix.lower()
            epub_bytes  = build_epub(meta, chapters, images, cover_bytes, cover_ext, lang=doc_lang)
            saved_path.write_bytes(epub_bytes)
            valid, messages = validate_epub(epub_bytes)

            result = ConversionResult(meta, chapters, images, epub_bytes, cover_bytes, cover_ext, saved_path, valid, messages)
            preview_html_path = self.write_preview_html(result)
            self.after(0, lambda: self.show_result(result, preview_html_path))
        except Exception as exc:
            self.after(0, lambda: self.show_error(exc))

    def write_preview_html(self, result: ConversionResult) -> Path:
        lang = self.ui_lang_var.get()
        tr   = TRANSLATIONS.get(lang, TRANSLATIONS["en"])

        # Pre-build data URIs so inline images render in the browser (temp dir has no images folder)
        data_uris: dict[str, str] = {}
        for img_name, img_bytes in result.images.items():
            mime = "image/png" if img_name.endswith(".png") else "image/jpeg"
            data_uris[img_name] = f"data:{mime};base64,{base64.b64encode(img_bytes).decode()}"

        def _fix_img_src(body: str) -> str:
            def _replace(m: re.Match) -> str:
                key = m.group(1).split("/")[-1]
                return f'src="{data_uris.get(key, m.group(1))}"'
            return re.sub(r'src="(\.\./images/[^"]+)"', _replace, body)

        parts = [
            f"<!doctype html><html lang='{lang}'><head><meta charset='utf-8'>"
            f"<title>{tr['preview_html_title']}</title>",
            "<style>" + CSS + "</style></head><body>",
            "<h1>" + xml_escape(result.meta["title"]) + "</h1>",
            "<p><strong>" + tr["author_label"] + ":</strong> " + xml_escape(result.meta["author"]) + "</p>",
        ]
        for _cid, title, _epub_type, _role, body in result.chapters:
            parts.append("<hr><h1>" + title + "</h1>" + _fix_img_src(body))
        parts.append("</body></html>")
        path = Path(tempfile.gettempdir()) / "docx_epub_converter_preview.html"
        path.write_text("\n".join(parts), encoding="utf-8")
        return path

    def show_result(self, result: ConversionResult, preview_html_path: Path):
        self.result = result
        self.preview_html_path = preview_html_path
        self._save_profile()
        self.progress.stop()
        self._set_convert_buttons_state(tk.NORMAL)
        self.open_file_button.configure(state=tk.NORMAL)
        self.open_folder_button.configure(state=tk.NORMAL)
        self.open_html_button.configure(state=tk.NORMAL)

        validation = (
            "EpubCheck: OK"
            if result.valid
            else f"EpubCheck: {len(result.messages)} " + self.t("epubcheck_messages")
        )
        self.status_var.set(self.t("status_done") + str(result.saved_path))
        self.result_label.configure(
            text=validation + "\n" + self.t("label_file") + " " + str(result.saved_path)
        )

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
        self.progress.stop()
        self._set_convert_buttons_state(tk.NORMAL)
        self.status_var.set(self.t("status_error"))
        self.result_label.configure(text=self.t("label_error") + " " + str(exc))
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

    # ------------------------------------------------------------------
    # Batch conversion
    # ------------------------------------------------------------------

    def open_batch_dialog(self):
        if not self.author_var.get().strip():
            messagebox.showwarning(APP_TITLE, self.t("batch_need_author"))
            return
        if not Path(self.cover_path.get()).is_file():
            messagebox.showwarning(APP_TITLE, self.t("batch_need_cover"))
            return
        paths = filedialog.askopenfilenames(
            title=self.t("batch_select"),
            initialdir=self._last_docx_dir or str(default_output_dir()),
            filetypes=[("DOCX", "*.docx")],
        )
        if paths:
            self._run_batch(list(paths))

    def _run_batch(self, docx_paths: list[str]):
        win = tk.Toplevel(self)
        win.title(self.t("batch_title"))
        win.configure(background=NAVY_BG)
        win.geometry("680x500")
        win.transient(self)

        frame = ttk.Frame(win, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=self.t("batch_title"), style="Section.TLabel").pack(anchor=tk.W, pady=(0, 8))

        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        status_list = tk.Listbox(
            list_frame, width=72, height=16,
            bg=NAVY_INPUT, fg=NAVY_BG, font=("Segoe UI", 9),
            selectbackground=NAVY_ACCENT, selectforeground="#ffffff",
            relief=tk.SOLID, borderwidth=1,
        )
        status_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=status_list.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        status_list.configure(yscrollcommand=sb.set)
        for p in docx_paths:
            status_list.insert(tk.END, f"[ ] {Path(p).name}")

        progress = ttk.Progressbar(frame, mode="determinate", maximum=len(docx_paths))
        progress.pack(fill=tk.X, pady=(8, 4))
        summary_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=summary_var, style="Card.TLabel").pack(anchor=tk.W)
        close_btn = ttk.Button(frame, text=self.t("btn_ok"), command=win.destroy, state=tk.DISABLED)
        close_btn.pack(anchor=tk.E, pady=(8, 0))

        def upd(idx, symbol, text):
            status_list.delete(idx)
            status_list.insert(idx, f"{symbol} {text}")

        def worker():
            ok = err = 0
            cover_path = Path(self.cover_path.get())
            cover_ext = cover_path.suffix.lower()
            try:
                cover_bytes = cover_path.read_bytes()
            except Exception as exc:
                win.after(0, lambda: summary_var.set(f"Cover error: {exc}"))
                win.after(0, lambda: close_btn.configure(state=tk.NORMAL))
                return
            doc_lang = self.doc_lang_var.get()

            for i, docx_str in enumerate(docx_paths):
                p = Path(docx_str)
                win.after(0, lambda idx=i, n=p.name: upd(idx, "[>]", n))
                try:
                    docx_bytes = p.read_bytes()
                    issues = inspect_docx_for_epub(docx_bytes, TRANSLATIONS[self.ui_lang_var.get()])
                    if issues:
                        raise ValueError(issues[0])
                    chapters, images = parse_docx(docx_bytes)
                    title = html.unescape(chapters[0][1]) if chapters else p.stem
                    meta = {
                        "title":       title,
                        "author":      self.author_var.get().strip(),
                        "publisher":   self.publisher_var.get().strip() or self.author_var.get().strip(),
                        "year":        self.year_var.get().strip() or str(date.today().year),
                        "isbn":        "",
                        "description": "",
                    }
                    epub_bytes = build_epub(meta, chapters, images, cover_bytes, cover_ext, lang=doc_lang)
                    out_path = p.with_suffix(".epub")
                    out_path.write_bytes(epub_bytes)
                    ok += 1
                    win.after(0, lambda idx=i, n=p.name: upd(idx, "[OK]", n))
                except Exception as exc:
                    err += 1
                    msg = f"{p.name} — {exc}"
                    win.after(0, lambda idx=i, m=msg: upd(idx, "[!!]", m))
                win.after(0, lambda v=i + 1: progress.configure(value=v))

            summary = self.t("batch_summary_ok") + str(ok)
            if err:
                summary += self.t("batch_summary_err") + str(err)
            win.after(0, lambda: summary_var.set(summary))
            win.after(0, lambda: close_btn.configure(state=tk.NORMAL))

        threading.Thread(target=worker, daemon=True).start()


def main():
    app = EpubConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
