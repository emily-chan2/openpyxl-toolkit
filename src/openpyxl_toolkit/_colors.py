"""Reading and writing the hex colours the toolkit accepts."""

from ._sentinel import UNCHANGED


def has_color(value):
    """True when a colour carries something the user chose.

    A cell that was never coloured reports the ARGB default rather than nothing,
    so an absent colour and an explicit one are only distinguishable by value.
    """
    if value is None:
        return False
    if isinstance(value, str):
        return value.upper() not in ("", "00000000")
    if getattr(value, "type", None) != "rgb":
        return True  # a theme, indexed or automatic colour is a real choice
    return (value.rgb or "").upper() not in ("", "00000000")


def normalize_color(color):
    """Turn '#f4d2d3' or 'f4d2d3' into the 8-digit ARGB openpyxl stores."""
    if color is UNCHANGED:
        return UNCHANGED
    if color is None:
        # None clears the colour; UNCHANGED is how "leave it alone" is spelled.
        return None
    hex_color = color.lstrip("#").upper()
    if len(hex_color) == 6:
        hex_color = f"FF{hex_color}"
    return hex_color
