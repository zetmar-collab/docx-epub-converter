from epub_converter.isbn import is_valid_isbn, normalize_isbn


def test_normalize_strips_hyphens():
    assert normalize_isbn("978-83-01-23456-7") == "9788301234567"


def test_valid_isbn13():
    assert is_valid_isbn("9780143007234") is True


def test_invalid_checksum():
    assert is_valid_isbn("9780143007230") is False


def test_valid_isbn10():
    assert is_valid_isbn("0306406152") is True
