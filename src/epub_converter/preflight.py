"""DOCX structure checks before EPUB conversion."""

from __future__ import annotations

import io

from docx import Document

from epub_converter.i18n import TRANSLATIONS
from epub_converter.styles import is_heading


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

        if is_heading(style_name, 1):
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

        if not text:
            continue

        if not seen_heading1:
            nonempty_before_first_heading.append(text[:80])
            continue

        if is_heading(style_name, 2):
            last_heading_level = 2
        elif is_heading(style_name, 3):
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
