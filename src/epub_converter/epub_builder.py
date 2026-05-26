"""EPUB 3 package assembly."""

from __future__ import annotations

import io
import re
import uuid
import zipfile
from datetime import date

from epub_converter.constants import A11Y_SUMMARIES, CSS, TOC_TITLES
from epub_converter.isbn import normalize_isbn
from epub_converter.utils import e, xml_escape


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
        f' xml:lang="{lang}" lang="{lang}">\n'
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
        f' xml:lang="{lang}" lang="{lang}">\n'
        "<head>\n"
        '  <meta charset="utf-8"/>\n'
        f"  <title>{title}</title>\n"
        '  <link rel="stylesheet" type="text/css" href="../styles/main.css"/>\n'
        "</head>\n"
        "<body>\n"
        f'  <section epub:type="{epub_type}" role="{role}">\n'
        f"    <h1>{title}</h1>\n"
        + body
        + "\n  </section>\n"
        "</body>\n"
        "</html>\n"
    )


def make_nav(chapters, lang: str = "pl") -> bytes:
    items = ""
    for cid, title, _epub_type, _role, _body in chapters:
        items += f'      <li><a href="content/{cid}.xhtml">{title}</a></li>\n'
    toc_title = TOC_TITLES.get(lang, "Table of Contents")
    return e(
        '<?xml version="1.0" encoding="utf-8"?>\n'
        "<!DOCTYPE html>\n"
        '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops"'
        f' xml:lang="{lang}" lang="{lang}">\n'
        f"<head><meta charset=\"utf-8\"/><title>{toc_title}</title></head>\n"
        "<body>\n"
        f'  <nav epub:type="toc" id="toc" role="doc-toc" aria-label="{toc_title}">\n'
        f"    <h1>{toc_title}</h1>\n"
        "    <ol>\n"
        + items
        + "    </ol>\n"
        "  </nav>\n"
        "</body>\n"
        "</html>\n"
    )


def _identifier_block(meta: dict) -> str:
    isbn = normalize_isbn(meta.get("isbn", ""))
    if isbn:
        return f'    <dc:identifier id="uid">urn:isbn:{xml_escape(isbn)}</dc:identifier>\n'
    return f'    <dc:identifier id="uid">urn:uuid:{uuid.uuid4()}</dc:identifier>\n'


def _a11y_summary(lang: str) -> str:
    return A11Y_SUMMARIES.get(lang, A11Y_SUMMARIES.get("en", A11Y_SUMMARIES["en"]))


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
        f'    <item id="cover-image" href="images/cover{cover_ext}" media-type="{cover_mime}" properties="cover-image"/>\n'
        '    <item id="cover-page" href="content/cover_page.xhtml" media-type="application/xhtml+xml"/>\n'
    )
    for img_name in images:
        ct = "image/png" if img_name.endswith(".png") else "image/jpeg"
        img_id = re.sub(r"[^a-z0-9]", "-", img_name)
        manifest_items += f'    <item id="{img_id}" href="images/{img_name}" media-type="{ct}"/>\n'

    spine = '    <itemref idref="cover-page" linear="yes"/>\n'
    for cid, _title, _epub_type, _role, _body in chapters:
        manifest_items += (
            f'    <item id="{cid}" href="content/{cid}.xhtml" media-type="application/xhtml+xml"/>\n'
        )
        spine += f'    <itemref idref="{cid}"/>\n'

    today = date.today().isoformat()
    isbn = normalize_isbn(meta.get("isbn", ""))
    isbn_meta = ""
    if isbn:
        isbn_meta = f'    <meta property="identifier-type" refines="#uid" scheme="onix:codelist5">15</meta>\n'

    cover_alt_pl = f"Okladka: {meta['title']}, autor {meta['author']}"
    cover_alt_en = f"Cover: {meta['title']}, {meta['author']}"

    return e(
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<package xmlns="http://www.idpf.org/2007/opf"\n'
        '         xmlns:dc="http://purl.org/dc/elements/1.1/"\n'
        '         xmlns:dcterms="http://purl.org/dc/terms/"\n'
        f'         version="3.0" unique-identifier="uid" xml:lang="{lang}">\n'
        '  <metadata xmlns:schema="http://schema.org/">\n'
        + _identifier_block(meta)
        + isbn_meta
        + f"    <dc:title>{xml_escape(meta['title'])}</dc:title>\n"
        f'    <dc:creator id="creator">{xml_escape(meta["author"])}</dc:creator>\n'
        '    <meta refines="#creator" property="role" scheme="marc:relators">aut</meta>\n'
        f"    <dc:language>{xml_escape(lang)}</dc:language>\n"
        f"    <dc:publisher>{xml_escape(meta['publisher'])}</dc:publisher>\n"
        f"    <dc:date>{xml_escape(meta['year'])}</dc:date>\n"
        f"    <dc:description>{xml_escape(meta['description'])}</dc:description>\n"
        "    <dc:subject>Ebook</dc:subject>\n"
        f"    <dc:rights>Copyright {xml_escape(meta['year'])} {xml_escape(meta['author'])}</dc:rights>\n"
        f'    <meta property="dcterms:modified">{today}T00:00:00Z</meta>\n'
        '    <meta property="schema:accessMode">textual</meta>\n'
        '    <meta property="schema:accessMode">visual</meta>\n'
        '    <meta property="schema:accessModeSufficient">textual</meta>\n'
        '    <meta property="schema:accessibilityFeature">structuralNavigation</meta>\n'
        '    <meta property="schema:accessibilityFeature">tableOfContents</meta>\n'
        '    <meta property="schema:accessibilityFeature">readingOrder</meta>\n'
        '    <meta property="schema:accessibilityFeature">alternativeText</meta>\n'
        '    <meta property="schema:accessibilityFeature">displayTransformability</meta>\n'
        '    <meta property="schema:accessibilityHazard">none</meta>\n'
        f'    <meta property="schema:accessibilitySummary">{xml_escape(_a11y_summary(lang))}</meta>\n'
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
    cover_alt = (
        f"Okladka: {meta['title']}, autor {meta['author']}"
        if lang == "pl"
        else f"Cover: {meta['title']}, {meta['author']}"
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
            zf.writestr(
                "OEBPS/content/" + cid + ".xhtml",
                make_chapter_xhtml(title, epub_type, role, body, lang),
            )
    return buffer.getvalue()
