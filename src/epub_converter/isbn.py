"""ISBN validation with checksum."""

import re


def normalize_isbn(raw: str) -> str:
    return re.sub(r"[\s\-]", "", (raw or "").strip())


def is_valid_isbn_format(isbn: str) -> bool:
    # ISBN-10 may end in 'X' (check digit 10); ISBN-13 is always 13 digits.
    return bool(re.fullmatch(r"\d{9}[\dXx]|\d{13}", isbn))


def _check_isbn10(isbn: str) -> bool:
    total = 0
    for i, ch in enumerate(isbn[:9]):
        total += int(ch) * (10 - i)
    check = (11 - (total % 11)) % 11
    last = "X" if check == 10 else str(check)
    return isbn[9].upper() == last


def _check_isbn13(isbn: str) -> bool:
    total = sum(int(ch) * (1 if i % 2 == 0 else 3) for i, ch in enumerate(isbn[:12]))
    return int(isbn[12]) == (10 - (total % 10)) % 10


def is_valid_isbn(isbn: str) -> bool:
    normalized = normalize_isbn(isbn)
    if not is_valid_isbn_format(normalized):
        return False
    if len(normalized) == 10:
        return _check_isbn10(normalized)
    return _check_isbn13(normalized)
