"""Word style name matching without false positives (e.g. Heading 10 vs Heading 1)."""


def matches_style_prefix(style_name: str, *base_names: str) -> bool:
    """Match exact style or Word variants like 'Heading 1 Char'."""
    for base in base_names:
        if style_name == base:
            return True
        if style_name.startswith(base + " "):
            return True
    return False


def is_heading(style_name: str, level: int) -> bool:
    return matches_style_prefix(style_name, f"Heading {level}", f"Nagłówek {level}")


def is_quote_style(style_name: str) -> bool:
    return matches_style_prefix(
        style_name,
        "Quote",
        "Cytat",
        "Blockquote",
        "Intense Quote",
        "Intensywny cytat",
    )


def is_bullet_list(style_name: str) -> bool:
    return matches_style_prefix(style_name, "List Bullet", "Lista wypunktowana", "List Paragraph")


def is_numbered_list(style_name: str) -> bool:
    return matches_style_prefix(style_name, "List Number", "Lista numerowana")
