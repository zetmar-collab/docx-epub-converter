from epub_converter.epub_builder import make_cover_xhtml, make_opf
from epub_converter.utils import e


def test_opf_uses_urn_isbn_when_provided():
    meta = {
        "title": "Test",
        "author": "Author",
        "publisher": "Pub",
        "year": "2026",
        "isbn": "9780143007234",
        "description": "Desc",
    }
    opf = make_opf(meta, [], {}, "image/jpeg", ".jpg", lang="en").decode("utf-8")
    assert 'id="uid">urn:isbn:9780143007234<' in opf


def test_cover_page_has_no_invalid_doc_cover_role():
    html = make_cover_xhtml("Okladka test", ".jpg", lang="pl").decode("utf-8")
    assert 'epub:type="cover"' in html
    assert "doc-cover" not in html
    assert 'role="doc-cover"' not in html


def test_opf_uuid_when_no_isbn():
    meta = {
        "title": "Test",
        "author": "Author",
        "publisher": "Pub",
        "year": "2026",
        "isbn": "",
        "description": "Desc",
    }
    opf = make_opf(meta, [], {}, "image/jpeg", ".jpg", lang="en").decode("utf-8")
    assert "urn:uuid:" in opf
    assert "urn:isbn:" not in opf.split("uid")[1].split("</dc:identifier>")[0]
