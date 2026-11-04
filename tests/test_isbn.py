from epub_converter.isbn import is_valid_isbn, normalize_isbn


def test_normalize_strips_hyphens():
    assert normalize_isbn("978-83-01-23456-7") == "9788301234567"


def test_valid_isbn13():
    assert is_valid_isbn("9780143007234") is True


def test_invalid_checksum():
    assert is_valid_isbn("9780143007230") is False


def test_valid_isbn10():
    assert is_valid_isbn("0306406152") is True


def test_isbn10_with_x_check_digit_is_accepted():
    """Check digit 10 is written as 'X'; _check_isbn10 has always handled it."""
    assert is_valid_isbn("080442957X") is True
    assert is_valid_isbn("0-8044-2957-X") is True


def test_isbn10_x_check_digit_is_case_insensitive():
    assert is_valid_isbn("080442957x") is True


def test_x_is_rejected_outside_the_check_digit_position():
    assert is_valid_isbn("08X442957X") is False


def test_isbn13_may_not_end_in_x():
    assert is_valid_isbn("978014300723X") is False


def test_normalize_strips_internal_whitespace():
    assert normalize_isbn(" 978 83 01 23456 7 ") == "9788301234567"


def test_wrong_length_is_rejected():
    assert is_valid_isbn("12345") is False
    assert is_valid_isbn("97801430072345") is False


def test_empty_and_non_numeric_input_is_rejected():
    assert is_valid_isbn("") is False
    assert is_valid_isbn("brak numeru") is False
