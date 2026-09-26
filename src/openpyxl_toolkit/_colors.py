"""Reading and writing the hex colors the toolkit accepts.

A color is checked here rather than left to openpyxl. openpyxl rejects a bad one with
``Colors must be aRGB hex values``, which names neither the argument it came from nor
the value that was wrong, and which describes the eight-digit form it stores rather than
the six-digit one most callers write.
"""

from __future__ import annotations

import re
from typing import Final, overload

from openpyxl.styles.colors import Color

from ._sentinel import UNCHANGED, Unchanged

# Six digits for RGB, or eight with an alpha channel in front. openpyxl stores
# only the eight-digit form; the six-digit one is widened on the way through.
_HEX_COLOR: Final = re.compile("[0-9A-F]{6}|[0-9A-F]{8}")

_HEX_DIGITS: Final = frozenset("0123456789ABCDEF")


def has_color(value: Color | str | None) -> bool:
    """True when a color carries something the user chose.

    A cell that was never colored reports the ARGB default rather than nothing,
    so an absent color and an explicit one are only distinguishable by value.
    """
    if value is None:
        return False
    if isinstance(value, str):
        return value.upper() not in ("", "00000000")
    if getattr(value, "type", None) != "rgb":
        return True  # a theme, indexed or automatic color is a real choice
    return (value.rgb or "").upper() not in ("", "00000000")


@overload
def normalize_color(color: Unchanged, name: str = ...) -> Unchanged: ...
@overload
def normalize_color(color: None, name: str = ...) -> None: ...
@overload
def normalize_color(color: str, name: str = ...) -> str: ...


def normalize_color(color: str | None | Unchanged, name: str = "color") -> str | None | Unchanged:
    """Turn '#f4d2d3' or 'f4d2d3' into the 8-digit ARGB openpyxl stores.

    ``name`` is the argument the color arrived in, so a value that is not a color is
    reported against what the caller wrote rather than against whichever cell openpyxl
    reached first.

    Raises
    ------
    ValueError
        If ``color`` is not six or eight hex digits, with or without a leading ``#``.
        Case and surrounding whitespace do not matter, since neither can mean anything
        else.
    """
    if color is UNCHANGED:
        return UNCHANGED
    if color is None:
        # None clears the color; UNCHANGED is how "leave it alone" is spelled.
        return None

    hex_color = color.strip().lstrip("#").upper()
    if not _HEX_COLOR.fullmatch(hex_color):
        # How many digits is only useful to someone who wrote digits. Saying it
        # to someone who wrote 'red' answers a question they did not ask.
        hint = (
            " A hex color has six digits, or eight with an alpha channel in front."
            if hex_color and set(hex_color) <= _HEX_DIGITS
            else ""
        )
        raise ValueError(f"{name} must be a hex color such as '#f4d2d3', got {color!r}.{hint}")

    if len(hex_color) == 6:
        hex_color = f"FF{hex_color}"
    return hex_color
