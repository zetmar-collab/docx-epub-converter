"""HTML preview export for converted EPUB."""

from __future__ import annotations

import base64
import re
import tempfile
from pathlib import Path

from epub_converter.constants import CSS
from epub_converter.i18n import TRANSLATIONS
from epub_converter.models import ConversionResult
from epub_converter.utils import xml_escape


def write_preview_html(result: ConversionResult, ui_lang: str) -> Path:
    tr = TRANSLATIONS.get(ui_lang, TRANSLATIONS["en"])
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
        f"<!doctype html><html lang='{ui_lang}'><head><meta charset='utf-8'>"
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
