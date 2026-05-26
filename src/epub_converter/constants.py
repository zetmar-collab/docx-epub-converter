"""Application and OOXML constants."""

from docx.oxml.ns import qn as _docx_qn

APP_TITLE = "DOCX EPUB Converter"
APP_VERSION = "2.1"
APP_AUTHOR = "Marek Zettel"

NAVY_BG = "#071a33"
NAVY_PANEL = "#0f2747"
NAVY_PANEL_ALT = "#14365f"
NAVY_BORDER = "#27527f"
NAVY_INPUT = "#f4f8ff"
NAVY_TEXT = "#edf5ff"
NAVY_ACCENT = "#2f80d0"
NAVY_ACCENT_ACTIVE = "#4aa3ff"

_W_P = _docx_qn("w:p")
_W_TBL = _docx_qn("w:tbl")
_W_R = _docx_qn("w:r")
_W_T = _docx_qn("w:t")
_W_HYPERLINK = _docx_qn("w:hyperlink")
_W_B = _docx_qn("w:b")
_W_I = _docx_qn("w:i")
_W_U = _docx_qn("w:u")
_W_STRIKE = _docx_qn("w:strike")
_W_RPR = _docx_qn("w:rPr")
_A_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
_R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
_R_HYP = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_W_FOOTNOTE_REF = _docx_qn("w:footnoteReference")
_W_FN_ID = _docx_qn("w:id")
_FOOTNOTES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
_WP_DOC_PR = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"

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

A11Y_SUMMARIES: dict[str, str] = {
    "pl": (
        "Publikacja dostepna cyfrowo. Semantyczna struktura HTML, nawigacyjny spis tresci, "
        "teksty alternatywne obrazow, kompletne metadane. Zgodna z EPUB Accessibility 1.1 i WCAG 2.1 AA."
    ),
    "en": (
        "Digitally accessible publication. Semantic HTML structure, navigational table of contents, "
        "image alternative text, complete metadata. Compliant with EPUB Accessibility 1.1 and WCAG 2.1 AA."
    ),
    "de": (
        "Digital zugangliches Publikation. Semantische HTML-Struktur, Navigationsinhalt, "
        "Alternativtexte fur Bilder. Entspricht EPUB Accessibility 1.1 und WCAG 2.1 AA."
    ),
    "fr": (
        "Publication numerique accessible. Structure HTML semantique, table des matieres de navigation, "
        "textes alternatifs des images. Conforme a EPUB Accessibility 1.1 et WCAG 2.1 AA."
    ),
    "es": (
        "Publicacion digital accesible. Estructura HTML semantica, tabla de contenidos de navegacion, "
        "textos alternativos de imagenes. Conforme con EPUB Accessibility 1.1 y WCAG 2.1 AA."
    ),
    "it": (
        "Pubblicazione digitale accessibile. Struttura HTML semantica, indice di navigazione, "
        "testi alternativi delle immagini. Conforme a EPUB Accessibility 1.1 e WCAG 2.1 AA."
    ),
}

DEFAULT_IMAGE_ALT: dict[str, str] = {
    "pl": "Ilustracja",
    "en": "Illustration",
    "de": "Abbildung",
    "fr": "Illustration",
    "es": "Ilustracion",
    "it": "Illustrazione",
    "cs": "Ilustrace",
    "sk": "Ilustracia",
    "hu": "Illusztracio",
    "ru": "Иллюстрация",
    "uk": "Ілюстрація",
    "nl": "Illustratie",
    "pt": "Ilustracao",
}

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
