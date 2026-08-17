"""Tests for the pre-conversion DOCX structure check."""

import docx_factory as F
from epub_converter.i18n import TRANSLATIONS
from epub_converter.preflight import inspect_docx_for_epub

PL = TRANSLATIONS["pl"]
EN = TRANSLATIONS["en"]


def test_well_formed_document_reports_no_issues():
    issues = inspect_docx_for_epub(
        F.simple_doc(
            [
                ("Heading 1", "Rozdzial pierwszy"),
                ("Normal", "Tresc rozdzialu."),
                ("Heading 2", "Podrozdzial"),
                ("Normal", "Dalsza tresc."),
                ("Heading 1", "Rozdzial drugi"),
                ("Normal", "Tresc."),
            ]
        )
    )
    assert issues == []


def test_missing_h1_is_reported():
    issues = inspect_docx_for_epub(
        F.simple_doc([("Normal", "Tekst bez zadnego naglowka")])
    )
    assert PL["issue_no_chapters"] in issues


def test_text_before_first_h1_is_reported():
    issues = inspect_docx_for_epub(
        F.simple_doc(
            [
                ("Normal", "Sierocy akapit"),
                ("Heading 1", "Rozdzial"),
                ("Normal", "Tresc"),
            ]
        )
    )
    assert PL["issue_text_before_h1"] in issues


def test_empty_h1_is_reported():
    issues = inspect_docx_for_epub(
        F.simple_doc([("Heading 1", ""), ("Normal", "Tresc")])
    )
    assert PL["issue_empty_h1"] in issues


def test_chapter_without_body_is_reported_by_name():
    issues = inspect_docx_for_epub(
        F.simple_doc(
            [
                ("Heading 1", "Pusty rozdzial"),
                ("Heading 1", "Pelny rozdzial"),
                ("Normal", "Tresc"),
            ]
        )
    )
    reported = [i for i in issues if i.startswith(PL["issue_empty_chapters"])]
    assert reported
    assert "Pusty rozdzial" in reported[0]
    assert "Pelny rozdzial" not in reported[0]


def test_last_chapter_without_body_is_also_reported():
    """The final chapter is only checked after the paragraph loop ends."""
    issues = inspect_docx_for_epub(
        F.simple_doc(
            [
                ("Heading 1", "Pelny"),
                ("Normal", "Tresc"),
                ("Heading 1", "Ostatni pusty"),
            ]
        )
    )
    reported = [i for i in issues if i.startswith(PL["issue_empty_chapters"])]
    assert reported and "Ostatni pusty" in reported[0]


def test_empty_chapter_list_is_capped_at_five_names():
    issues = inspect_docx_for_epub(
        F.simple_doc([("Heading 1", f"Pusty {i}") for i in range(1, 8)])
    )
    reported = [i for i in issues if i.startswith(PL["issue_empty_chapters"])][0]
    names = reported[len(PL["issue_empty_chapters"]) :].split(", ")
    assert len(names) == 5


def test_h3_without_preceding_h2_is_reported():
    issues = inspect_docx_for_epub(
        F.simple_doc(
            [
                ("Heading 1", "Rozdzial"),
                ("Heading 3", "Za gleboko"),
                ("Normal", "Tresc"),
            ]
        )
    )
    assert PL["issue_h3_without_h2"] in issues


def test_h3_after_h2_is_accepted():
    issues = inspect_docx_for_epub(
        F.simple_doc(
            [
                ("Heading 1", "Rozdzial"),
                ("Heading 2", "Podrozdzial"),
                ("Heading 3", "Pod-pod"),
                ("Normal", "Tresc"),
            ]
        )
    )
    assert PL["issue_h3_without_h2"] not in issues


def test_heading_level_resets_for_each_new_chapter():
    """H2 in chapter one must not license a bare H3 in chapter two."""
    issues = inspect_docx_for_epub(
        F.simple_doc(
            [
                ("Heading 1", "Pierwszy"),
                ("Heading 2", "Podrozdzial"),
                ("Normal", "Tresc"),
                ("Heading 1", "Drugi"),
                ("Heading 3", "Za gleboko"),
                ("Normal", "Tresc"),
            ]
        )
    )
    assert PL["issue_h3_without_h2"] in issues


def test_issues_use_the_supplied_translation():
    issues = inspect_docx_for_epub(F.simple_doc([("Normal", "Bez naglowkow")]), EN)
    assert EN["issue_no_chapters"] in issues
    assert PL["issue_no_chapters"] not in issues


def test_blank_paragraphs_do_not_count_as_chapter_body():
    issues = inspect_docx_for_epub(
        F.simple_doc([("Heading 1", "Rozdzial"), ("Normal", "   ")])
    )
    reported = [i for i in issues if i.startswith(PL["issue_empty_chapters"])]
    assert reported and "Rozdzial" in reported[0]
