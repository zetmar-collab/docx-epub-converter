"""Tests for shared helpers: filenames, HTML flattening, dropped paths, config."""

import pytest

from epub_converter import config as config_module
from epub_converter.utils import (
    html_to_plain_text,
    safe_filename,
    sanitize_dropped_path,
    xml_escape,
)


# ---------------------------------------------------------------------------
# safe_filename
# ---------------------------------------------------------------------------


def test_safe_filename_lowercases_and_hyphenates():
    assert safe_filename("Moja Wielka Ksiazka") == "moja-wielka-ksiazka.epub"


def test_safe_filename_strips_punctuation():
    assert safe_filename("Tytul: podtytul, cz. 2!") == "tytul-podtytul-cz-2.epub"


def test_safe_filename_keeps_unicode_letters():
    assert safe_filename("Zażółć gęślą jaźń") == "zażółć-gęślą-jaźń.epub"


def test_safe_filename_has_no_leading_or_trailing_hyphen():
    name = safe_filename("!!! Tytul !!!")
    assert name == "tytul.epub"


@pytest.mark.parametrize("title", ["", "   ", "???", "---"])
def test_safe_filename_falls_back_for_unusable_titles(title):
    assert safe_filename(title) == "ebook.epub"


def test_safe_filename_keeps_existing_hyphens_verbatim():
    """Hyphens are in the keep-set, so only the runs *between* them collapse.
    'Tytul - podtytul' therefore yields '-3-' rather than a single separator -
    cosmetic, and the resulting name is still valid and unique."""
    assert safe_filename("a   ---   b") == "a-----b.epub"
    assert safe_filename("Tytul - podtytul") == "tytul---podtytul.epub"


# ---------------------------------------------------------------------------
# xml_escape
# ---------------------------------------------------------------------------


def test_xml_escape_handles_ampersand_and_angle_brackets():
    assert xml_escape("a & <b>") == "a &amp; &lt;b&gt;"


def test_xml_escape_leaves_quotes_alone():
    """quote=False - attribute values are escaped by their own call sites."""
    assert xml_escape('say "hi"') == 'say "hi"'


def test_xml_escape_accepts_none():
    assert xml_escape(None) == ""


# ---------------------------------------------------------------------------
# html_to_plain_text
# ---------------------------------------------------------------------------


def test_plain_text_strips_tags():
    assert html_to_plain_text("<p>Ala ma <strong>kota</strong></p>") == "Ala ma kota"


def test_plain_text_unescapes_entities():
    assert html_to_plain_text("<p>Kowalski &amp; Syn</p>") == "Kowalski & Syn"


def test_plain_text_breaks_lines_on_block_elements():
    result = html_to_plain_text("<p>Jeden</p><p>Dwa</p>")
    assert result == "Jeden\nDwa"


def test_plain_text_marks_list_items_with_dashes():
    result = html_to_plain_text("<ul><li>raz</li><li>dwa</li></ul>")
    assert "- raz" in result and "- dwa" in result


def test_plain_text_collapses_excess_blank_lines():
    assert "\n\n\n" not in html_to_plain_text("<p>a</p><p></p><p></p><p>b</p>")


def test_plain_text_separates_table_cells():
    result = html_to_plain_text("<table><tr><td>A</td><td>B</td></tr></table>")
    assert "A" in result and "B" in result
    assert "AB" not in result


# ---------------------------------------------------------------------------
# sanitize_dropped_path
# ---------------------------------------------------------------------------


def test_dropped_path_accepts_an_existing_file(tmp_path):
    target = tmp_path / "ksiazka.docx"
    target.write_bytes(b"x")
    assert sanitize_dropped_path(str(target)) == target.resolve()


def test_dropped_path_strips_tk_braces_and_quotes(tmp_path):
    target = tmp_path / "z odstepem.docx"
    target.write_bytes(b"x")
    assert sanitize_dropped_path("{" + str(target) + "}") == target.resolve()
    assert sanitize_dropped_path('"' + str(target) + '"') == target.resolve()


def test_dropped_path_rejects_a_directory(tmp_path):
    assert sanitize_dropped_path(str(tmp_path)) is None


def test_dropped_path_rejects_a_missing_file(tmp_path):
    assert sanitize_dropped_path(str(tmp_path / "nie_ma.docx")) is None


def test_dropped_path_rejects_empty_input():
    assert sanitize_dropped_path("") is None


# ---------------------------------------------------------------------------
# config persistence
# ---------------------------------------------------------------------------


@pytest.fixture
def isolated_config(tmp_path, monkeypatch):
    """Point config at a temp file and disable the legacy fallback."""
    cfg = tmp_path / "config.json"
    monkeypatch.setattr(config_module, "CONFIG_FILE", cfg)
    monkeypatch.setattr(config_module, "_LEGACY_CONFIG_FILE", tmp_path / "absent.json")
    return cfg


def test_config_round_trips(isolated_config):
    config_module.save_config({"profile_author": "Marek Zettel"})
    assert config_module.load_config()["profile_author"] == "Marek Zettel"


def test_config_creates_missing_parent_directory(tmp_path, monkeypatch):
    nested = tmp_path / "a" / "b" / "config.json"
    monkeypatch.setattr(config_module, "CONFIG_FILE", nested)
    monkeypatch.setattr(config_module, "_LEGACY_CONFIG_FILE", tmp_path / "absent.json")
    config_module.save_config({"k": "v"})
    assert nested.is_file()


def test_config_preserves_non_ascii(isolated_config):
    config_module.save_config({"profile_publisher": "Wydawnictwo Żółw"})
    assert config_module.load_config()["profile_publisher"] == "Wydawnictwo Żółw"


def test_update_config_merges_instead_of_replacing(isolated_config):
    config_module.save_config({"profile_author": "A", "profile_year": "2026"})
    config_module.update_config({"profile_year": "2027"})
    cfg = config_module.load_config()
    assert cfg == {"profile_author": "A", "profile_year": "2027"}


def test_missing_config_loads_as_empty(isolated_config):
    assert config_module.load_config() == {}


def test_corrupt_config_loads_as_empty(isolated_config):
    isolated_config.write_text("{ this is not json", encoding="utf-8")
    assert config_module.load_config() == {}


def test_non_dict_config_loads_as_empty(isolated_config):
    isolated_config.write_text('["a", "b"]', encoding="utf-8")
    assert config_module.load_config() == {}


def test_legacy_config_is_migrated_forward(tmp_path, monkeypatch):
    legacy = tmp_path / "legacy.json"
    legacy.write_text('{"profile_author": "Stary Profil"}', encoding="utf-8")
    new = tmp_path / "new" / "config.json"
    monkeypatch.setattr(config_module, "CONFIG_FILE", new)
    monkeypatch.setattr(config_module, "_LEGACY_CONFIG_FILE", legacy)

    assert config_module.load_config()["profile_author"] == "Stary Profil"
    assert new.is_file(), "legacy config should be copied to the new location"
    assert legacy.is_file(), "legacy config should be left in place"


def test_new_config_wins_over_legacy(tmp_path, monkeypatch):
    legacy = tmp_path / "legacy.json"
    legacy.write_text('{"profile_author": "Stary"}', encoding="utf-8")
    new = tmp_path / "config.json"
    new.write_text('{"profile_author": "Nowy"}', encoding="utf-8")
    monkeypatch.setattr(config_module, "CONFIG_FILE", new)
    monkeypatch.setattr(config_module, "_LEGACY_CONFIG_FILE", legacy)

    assert config_module.load_config()["profile_author"] == "Nowy"
