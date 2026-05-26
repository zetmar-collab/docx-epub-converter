from epub_converter.styles import is_heading, is_bullet_list


def test_heading_10_is_not_heading_1():
    assert not is_heading("Heading 10", 1)
    assert is_heading("Heading 1", 1)
    assert is_heading("Heading 1 Char", 1)


def test_heading_2_vs_20():
    assert not is_heading("Heading 20", 2)
    assert is_heading("Heading 2", 2)


def test_polish_headings():
    assert is_heading("Nagłówek 1", 1)
    assert not is_heading("Nagłówek 10", 1)


def test_list_bullet_prefix():
    assert is_bullet_list("List Bullet")
    assert is_bullet_list("List Bullet 2")
