"""Tests for EpubCheck invocation and result reporting.

The distinction these lock down: a validator that could not run must never be
reported as a defect in the user's book.
"""

import subprocess

import pytest

from epub_converter import validation


# ---------------------------------------------------------------------------
# Availability
# ---------------------------------------------------------------------------


def test_epubcheck_unavailable_without_java(monkeypatch):
    monkeypatch.setattr(validation, "_java_executable", lambda: None)
    assert validation.epubcheck_available() is False


def test_epubcheck_unavailable_without_jar(monkeypatch):
    monkeypatch.setattr(validation, "_epubcheck_jar", lambda: None)
    assert validation.epubcheck_available() is False


def test_java_available_follows_executable_lookup(monkeypatch):
    monkeypatch.setattr(validation, "_java_executable", lambda: r"C:\java\bin\java.exe")
    assert validation.java_available() is True
    monkeypatch.setattr(validation, "_java_executable", lambda: None)
    assert validation.java_available() is False


# ---------------------------------------------------------------------------
# "Did not run" must never look like "invalid"
# ---------------------------------------------------------------------------


def test_missing_java_reports_did_not_run(monkeypatch):
    monkeypatch.setattr(validation, "_java_executable", lambda: None)
    ran, valid, messages = validation.validate_epub(b"whatever")
    assert ran is False
    assert valid is True
    assert messages == []


def test_missing_jar_reports_did_not_run(monkeypatch):
    monkeypatch.setattr(validation, "_epubcheck_jar", lambda: None)
    ran, valid, messages = validation.validate_epub(b"whatever")
    assert (ran, valid, messages) == (False, True, [])


def _fake_run(stdout: bytes):
    def run(*_args, **_kwargs):
        return subprocess.CompletedProcess(args=[], returncode=1, stdout=stdout, stderr=b"")

    return run


def test_crashed_validator_reports_did_not_run(monkeypatch):
    """A StackOverflowError leaves no parseable JSON. That is a tool failure,
    not a broken book, and must not be shown to the user as a defect."""
    monkeypatch.setattr(validation, "_java_executable", lambda: "java")
    monkeypatch.setattr(validation, "_epubcheck_jar", lambda: "epubcheck.jar")
    monkeypatch.setattr(validation.subprocess, "run", _fake_run(b"java.lang.StackOverflowError\n"))
    ran, valid, messages = validation.validate_epub(b"epub")
    assert ran is False
    assert valid is True
    assert messages == []


def test_empty_output_reports_did_not_run(monkeypatch):
    monkeypatch.setattr(validation, "_java_executable", lambda: "java")
    monkeypatch.setattr(validation, "_epubcheck_jar", lambda: "epubcheck.jar")
    monkeypatch.setattr(validation.subprocess, "run", _fake_run(b""))
    assert validation.validate_epub(b"epub") == (False, True, [])


def test_timeout_reports_did_not_run(monkeypatch):
    def boom(*_args, **_kwargs):
        raise subprocess.TimeoutExpired(cmd="java", timeout=1)

    monkeypatch.setattr(validation, "_java_executable", lambda: "java")
    monkeypatch.setattr(validation, "_epubcheck_jar", lambda: "epubcheck.jar")
    monkeypatch.setattr(validation.subprocess, "run", boom)
    assert validation.validate_epub(b"epub") == (False, True, [])


def test_launch_failure_reports_did_not_run(monkeypatch):
    def boom(*_args, **_kwargs):
        raise OSError("cannot start java")

    monkeypatch.setattr(validation, "_java_executable", lambda: "java")
    monkeypatch.setattr(validation, "_epubcheck_jar", lambda: "epubcheck.jar")
    monkeypatch.setattr(validation.subprocess, "run", boom)
    assert validation.validate_epub(b"epub") == (False, True, [])


# ---------------------------------------------------------------------------
# Verdicts from real JSON
# ---------------------------------------------------------------------------


def _json(n_fatal=0, n_error=0, n_warning=0, messages=None):
    import json

    return json.dumps(
        {
            "checker": {
                "filename": "book.epub",
                "nFatal": n_fatal,
                "nError": n_error,
                "nWarning": n_warning,
            },
            "messages": messages or [],
        }
    ).encode()


def _patched(monkeypatch, payload):
    monkeypatch.setattr(validation, "_java_executable", lambda: "java")
    monkeypatch.setattr(validation, "_epubcheck_jar", lambda: "epubcheck.jar")
    monkeypatch.setattr(validation.subprocess, "run", _fake_run(payload))


def test_clean_report_is_valid(monkeypatch):
    _patched(monkeypatch, _json())
    assert validation.validate_epub(b"epub") == (True, True, [])


def test_errors_make_it_invalid(monkeypatch):
    _patched(monkeypatch, _json(n_error=2))
    ran, valid, _ = validation.validate_epub(b"epub")
    assert (ran, valid) == (True, False)


def test_fatals_make_it_invalid(monkeypatch):
    _patched(monkeypatch, _json(n_fatal=1))
    assert validation.validate_epub(b"epub")[1] is False


def test_warnings_alone_stay_valid(monkeypatch):
    _patched(monkeypatch, _json(n_warning=5))
    ran, valid, _ = validation.validate_epub(b"epub")
    assert (ran, valid) == (True, True)


# ---------------------------------------------------------------------------
# Message formatting — the UI calls str() on each of these
# ---------------------------------------------------------------------------


def test_message_without_suggestion_formats_without_crashing(monkeypatch):
    """EpubCheck routinely omits 'suggestion'. epubcheck.models.Message raises
    TypeError on those, which used to crash the result dialog."""
    _patched(
        monkeypatch,
        _json(
            n_error=1,
            messages=[
                {
                    "ID": "RSC-005",
                    "severity": "ERROR",
                    "message": "Error while parsing file",
                    "suggestion": None,
                    "locations": [{"path": "OEBPS/ch01.xhtml", "line": 12, "column": 3}],
                }
            ],
        ),
    )
    _, _, messages = validation.validate_epub(b"epub")
    rendered = "\n".join(str(m) for m in messages)
    assert "RSC-005" in rendered
    assert "OEBPS/ch01.xhtml:12:3" in rendered
    assert "None" not in rendered


def test_message_without_locations_formats(monkeypatch):
    _patched(
        monkeypatch,
        _json(
            n_error=1,
            messages=[
                {
                    "ID": "PKG-008",
                    "severity": "FATAL",
                    "message": "Unable to read file",
                    "suggestion": None,
                    "locations": [],
                }
            ],
        ),
    )
    rendered = str(validation.validate_epub(b"epub")[2][0])
    assert rendered.startswith("FATAL PKG-008")
    assert "Unable to read file" in rendered


def test_suggestion_is_included_when_present(monkeypatch):
    _patched(
        monkeypatch,
        _json(
            n_error=1,
            messages=[
                {
                    "ID": "OPF-003",
                    "severity": "WARNING",
                    "message": "Item not declared",
                    "suggestion": "Add it to the manifest",
                    "locations": [{"path": "OEBPS/content.opf", "line": -1, "column": -1}],
                }
            ],
        ),
    )
    rendered = str(validation.validate_epub(b"epub")[2][0])
    assert "Item not declared" in rendered
    assert "Add it to the manifest" in rendered
    assert ":-1" not in rendered


def test_one_entry_per_location(monkeypatch):
    _patched(
        monkeypatch,
        _json(
            n_error=1,
            messages=[
                {
                    "ID": "RSC-012",
                    "severity": "ERROR",
                    "message": "Fragment identifier is not defined",
                    "suggestion": None,
                    "locations": [
                        {"path": "OEBPS/ch01.xhtml", "line": 5, "column": 1},
                        {"path": "OEBPS/ch02.xhtml", "line": 9, "column": 2},
                    ],
                }
            ],
        ),
    )
    rendered = str(validation.validate_epub(b"epub")[2][0])
    assert "OEBPS/ch01.xhtml:5:1" in rendered
    assert "OEBPS/ch02.xhtml:9:2" in rendered


# ---------------------------------------------------------------------------
# Java invocation
# ---------------------------------------------------------------------------


def test_java_is_invoked_with_a_larger_stack(monkeypatch):
    """Without -Xss the RelaxNG compiler overflows the default JVM stack and
    a valid EPUB is reported as broken."""
    captured = {}

    def capture(cmd, *_args, **_kwargs):
        captured["cmd"] = cmd
        return subprocess.CompletedProcess(args=cmd, returncode=0, stdout=_json(), stderr=b"")

    monkeypatch.setattr(validation, "_java_executable", lambda: "java")
    monkeypatch.setattr(validation, "_epubcheck_jar", lambda: "epubcheck.jar")
    monkeypatch.setattr(validation.subprocess, "run", capture)
    validation.validate_epub(b"epub")

    cmd = captured["cmd"]
    assert any(str(arg).startswith("-Xss") for arg in cmd)
    assert "--json" in cmd and "-jar" in cmd


def test_temp_file_is_removed_even_when_java_fails(monkeypatch, tmp_path):
    created = []
    real_named = validation.tempfile.NamedTemporaryFile

    def tracking(*args, **kwargs):
        handle = real_named(*args, **kwargs)
        created.append(handle.name)
        return handle

    monkeypatch.setattr(validation.tempfile, "NamedTemporaryFile", tracking)
    monkeypatch.setattr(validation, "_java_executable", lambda: "java")
    monkeypatch.setattr(validation, "_epubcheck_jar", lambda: "epubcheck.jar")
    monkeypatch.setattr(
        validation.subprocess, "run", lambda *a, **k: (_ for _ in ()).throw(OSError("nope"))
    )

    validation.validate_epub(b"epub")
    import os

    assert created and not any(os.path.exists(p) for p in created)


# ---------------------------------------------------------------------------
# End-to-end against the real jar (skipped when unavailable)
# ---------------------------------------------------------------------------


@pytest.mark.skipif(
    not validation.epubcheck_available(), reason="requires Java and the epubcheck jar"
)
def test_real_epubcheck_accepts_a_generated_epub():
    import docx_factory as F
    from epub_converter.docx_parser import parse_docx
    from epub_converter.epub_builder import build_epub
    from PIL import Image
    import io

    docx = F.simple_doc(
        [
            ("Heading 1", "Rozdzial pierwszy"),
            ("Normal", "Tresc rozdzialu."),
            ("Heading 2", "Podrozdzial"),
            ("List Bullet", "Punkt listy"),
            ("Quote", "Cytat"),
        ]
    )
    chapters, images = parse_docx(docx, "pl")
    buf = io.BytesIO()
    Image.new("RGB", (1600, 2400), (7, 26, 51)).save(buf, "PNG")
    meta = {
        "title": "Ksiazka testowa",
        "author": "Autor",
        "publisher": "Wydawca",
        "year": "2026",
        "isbn": "",
        "description": "Opis",
    }
    epub = build_epub(meta, chapters, images, buf.getvalue(), ".png", lang="pl")

    ran, valid, messages = validation.validate_epub(epub)
    assert ran is True, "EpubCheck should run when Java and the jar are present"
    assert valid is True, "generated EPUB should pass EpubCheck:\n" + "\n".join(
        str(m) for m in messages
    )


@pytest.mark.skipif(
    not validation.epubcheck_available(), reason="requires Java and the epubcheck jar"
)
def test_real_epubcheck_rejects_garbage():
    ran, valid, messages = validation.validate_epub(b"this is not an epub archive")
    assert ran is True
    assert valid is False
    assert messages
    "\n".join(str(m) for m in messages)  # must not raise
