"""What a cell displays, how wide that is, and how many lines it takes.

Excel shows a stored value through the cell's number format, and a column is
fitted to what the reader sees rather than to the value underneath.

Fitting a row runs the measurement the other way. A column's width says how many
pixels of text a cell has room for, and the text is broken into lines against
that; the row is then as tall as the lines it holds.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterable, Mapping
from datetime import date, datetime, time
from typing import TYPE_CHECKING, Any

from openpyxl.styles import DEFAULT_FONT

from . import _metrics
from .columns import column_letter

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet

    from ._types import AnyCell, FontSpec

# Splits on runs of whitespace and keeps them, because the spaces between two
# words that stay on one line take up room. re.split with a group is the only
# spelling that returns both halves.
_WHITESPACE = re.compile(r"(\s+)")

# Excel date tokens, longest first so "yyyy" is matched before "yy". Each one maps
# to a function of the value, rather than to a strftime directive: the no-padding
# directives (%-d and friends) are a glibc and BSD extension that Windows rejects,
# and falling back on those platforms silently mis-measured the column. The
# name-based parts still go through strftime, which is portable and gives the
# locale's own month and day names.
DATE_TOKENS: tuple[tuple[str, Callable[[Any], str]], ...] = (
    ("yyyy", lambda v: f"{v.year:04d}"),
    ("yy", lambda v: f"{v.year % 100:02d}"),
    ("mmmm", lambda v: v.strftime("%B")),
    ("mmm", lambda v: v.strftime("%b")),
    ("dddd", lambda v: v.strftime("%A")),
    ("ddd", lambda v: v.strftime("%a")),
    ("dd", lambda v: f"{v.day:02d}"),
    ("d", lambda v: str(v.day)),
    ("ss", lambda v: f"{v.second:02d}"),
    ("s", lambda v: str(v.second)),
    ("am/pm", lambda v: v.strftime("%p")),
)


def _month_or_minute(value: Any, after_hour: bool, padded: bool) -> str | None:
    """What an "m" token stands for, or None when the value does not carry it.

    An "m" is minutes once an hour token has been seen and a month otherwise.
    A time has no month and a date has no minute, so a format asking for the
    part the value does not hold gets no guess, the same as the other tokens.
    """
    number = getattr(value, "minute" if after_hour else "month", None)
    if number is None:
        return None
    return f"{number:02d}" if padded else str(number)


def displayed_text(cell: AnyCell) -> str:
    """The text Excel shows in a cell, as far as it can be worked out cheaply.

    Only dates and times are translated. Excel's number formats are a language of
    their own, and a partial implementation of the rest would mis-measure in ways
    that are harder to notice than the raw value, so everything else is measured
    as it is stored.
    """
    value = cell.value
    if value is None:
        return ""
    if not isinstance(value, (datetime, date, time)):
        return str(value)

    code = (cell.number_format or "").lower()
    if not code or code == "general":
        return str(value)
    # A bare date has no time parts and a bare time has no date parts; a format
    # asking for what the value does not carry is not worth guessing at.
    needs = {"y": "year", "d": "day", "h": "hour", "s": "second"}
    if any(not hasattr(value, attr) for letter, attr in needs.items() if letter in code):
        return str(value)

    # An hour is written 12-hour when the code also carries AM/PM.
    twelve_hour = "am/pm" in code

    def hour(value: Any, padded: bool) -> str:
        shown = value.hour
        if twelve_hour:
            shown = shown % 12 or 12
        return f"{shown:02d}" if padded else str(shown)

    # A minute token looks identical to a month token; Excel tells them apart by
    # whether an hour came first, so track that while walking the code.
    out, index, after_hour = [], 0, False
    while index < len(code):
        # The table first, so the month-name tokens mmmm and mmm are taken whole
        # rather than having their first two characters eaten as a numeric month.
        for token, render in DATE_TOKENS:
            if code.startswith(token, index):
                out.append(render(value))
                index += len(token)
                break
        else:
            if code.startswith("hh", index) or code.startswith("h", index):
                padded = code.startswith("hh", index)
                after_hour = True
                out.append(hour(value, padded))
                index += 2 if padded else 1
            elif code.startswith("mm", index) or code.startswith("m", index):
                padded = code.startswith("mm", index)
                part = _month_or_minute(value, after_hour, padded)
                if part is None:
                    return str(value)
                out.append(part)
                index += 2 if padded else 1
            else:
                out.append(code[index])
                index += 1

    return "".join(out)


def measure_text(text: str, font: FontSpec, normal_font: FontSpec) -> float:
    """Column width that fits ``text``, from the character widths of its font."""
    name, size = font
    base_name, base_size = normal_font
    return _metrics.column_width(text, size, base_size, name, base_name)


def workbook_normal_font(worksheet: Worksheet) -> FontSpec:
    """The workbook's normal font, as ``(name, point_size)``.

    A cell with no font record of its own inherits this, and Excel's column
    width unit is defined by it rather than by whatever the cell uses.
    """
    fonts = getattr(worksheet.parent, "_fonts", None)
    normal = fonts[0] if fonts else None
    # DEFAULT_FONT is annotated as carrying an optional name and size, though it
    # always has both; the last fallback is what keeps the return type honest.
    name = getattr(normal, "name", None) or DEFAULT_FONT.name or _metrics.DEFAULT_FONT_NAME
    size = getattr(normal, "sz", None) or DEFAULT_FONT.sz or 11.0
    return (str(name), float(size))


def measure_row_height(lines: int, font: FontSpec) -> float:
    """Height in points of a cell holding ``lines`` lines of ``font``."""
    _, size = font
    return _metrics.row_height(lines, size)


def column_width_of(worksheet: Worksheet, column: int) -> float:
    """The width a column already has, as Excel reads it.

    A column nobody has sized carries no width in the file at all, and takes the
    sheet's default; a sheet naming none takes Excel's. openpyxl instead reports
    13.0 for such a column, which is neither of those and is not written to the
    file either, so reading its ``width`` straight through would size a wrapped
    cell against a column width that exists nowhere but in openpyxl.
    """
    letter = column_letter(column)
    dimensions = worksheet.column_dimensions
    # Membership rather than a lookup. column_dimensions creates whatever entry
    # it is asked for, so looking one up is enough to make a column that was
    # never touched report a deliberate width of 13.0 from then on.
    if letter in dimensions:
        width = dimensions[letter].width
        if width:
            return float(width)
    default = worksheet.sheet_format.defaultColWidth
    return float(default) if default else _metrics.DEFAULT_COLUMN_WIDTH


def available_pixels(worksheet: Worksheet, columns: Iterable[int], normal_font: FontSpec) -> float:
    """Pixels of text that fit across ``columns``.

    A width of N stands for N of the workbook's widest digit, and all of it is
    room to wrap into. The five pixels of padding in the column width formula
    are not deducted here: measured against Excel's own best fit, a cell wraps
    at the full width of its column, and taking the padding off again counts a
    line more than Excel draws on text that only just fits.

    More than one column is a cell merged across them, which wraps across the
    whole merge.
    """
    base_name, base_size = normal_font
    base_widths, _ = _metrics.widths_for(base_name, fallback=_metrics.DEFAULT_FONT_NAME)
    unit = _metrics.max_digit_pixels(base_size, base_widths)
    total = sum(column_width_of(worksheet, column) for column in columns)
    return total * unit


def _paragraph_lines(
    text: str, point_size: float, widths: Mapping[str, float], pixels: float
) -> int:
    """How many lines one run of text takes when wrapped to ``pixels``."""
    lines = 1
    used = 0.0
    for token in _WHITESPACE.split(text):
        if not token:
            continue
        measured = _metrics.text_pixels(token, point_size, widths)
        if token.isspace():
            # A break swallows the spaces it falls on, so whitespace never
            # pushes a line over by itself. It only takes room ahead of a word
            # that ends up on the same line.
            used += measured
            continue
        if used and used + measured > pixels:
            lines += 1
            used = 0.0
        if measured <= pixels:
            used += measured
            continue
        # A single word wider than the line. Excel breaks it where it runs out
        # rather than letting it spill, so walk it a character at a time.
        for character in token:
            advance = _metrics.text_pixels(character, point_size, widths)
            if used and used + advance > pixels:
                lines += 1
                used = 0.0
            used += advance
    return lines


def count_lines(text: str, font: FontSpec, pixels: float) -> int:
    """How many lines ``text`` takes in a cell with ``pixels`` of room.

    A line break in the value always starts a new line. What is left over is
    wrapped the way Excel wraps it: at a space where there is one, and inside
    the word where there is not.
    """
    name, size = font
    widths, _ = _metrics.widths_for(name)
    # Excel writes a line break as \n; a value that came from somewhere else may
    # carry the pair.
    paragraphs = text.replace("\r\n", "\n").split("\n")
    if pixels < _metrics.max_digit_pixels(size, widths):
        # A column too narrow to hold one character. Breaking against a width
        # that small would hand back a line per character and a row hundreds of
        # points tall, where what the reader gets is a column showing nothing.
        return len(paragraphs)
    return sum(_paragraph_lines(line, size, widths, pixels) for line in paragraphs)
