"""What a cell displays, and how wide that is.

Excel shows a stored value through the cell's number format, and a column is
fitted to what the reader sees rather than to the value underneath.
"""

from datetime import date, datetime, time

from openpyxl.styles import DEFAULT_FONT

from . import _metrics

# Excel date tokens, longest first so "yyyy" is matched before "yy". Each one maps
# to a function of the value, rather than to a strftime directive: the no-padding
# directives (%-d and friends) are a glibc and BSD extension that Windows rejects,
# and falling back on those platforms silently mis-measured the column. The
# name-based parts still go through strftime, which is portable and gives the
# locale's own month and day names.
DATE_TOKENS = (
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


def _month_or_minute(value, after_hour, padded):
    """What an "m" token stands for, or None when the value does not carry it.

    An "m" is minutes once an hour token has been seen and a month otherwise.
    A time has no month and a date has no minute, so a format asking for the
    part the value does not hold gets no guess, the same as the other tokens.
    """
    number = getattr(value, "minute" if after_hour else "month", None)
    if number is None:
        return None
    return f"{number:02d}" if padded else str(number)


def displayed_text(cell):
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

    def hour(value, padded):
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


def measure_text(text, font, normal_font):
    """Column width that fits ``text``, from the real advances of its font."""
    name, size = font
    base_name, base_size = normal_font
    return _metrics.column_width(text, size, base_size, name, base_name)


def workbook_normal_font(worksheet):
    """The workbook's normal font, as ``(name, point_size)``.

    A cell with no font record of its own inherits this, and Excel's column
    width unit is defined by it rather than by whatever the cell uses.
    """
    fonts = getattr(worksheet.parent, "_fonts", None)
    normal = fonts[0] if fonts else None
    return (
        getattr(normal, "name", None) or DEFAULT_FONT.name,
        getattr(normal, "sz", None) or DEFAULT_FONT.sz,
    )
