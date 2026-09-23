"""The font metrics behind set_column_best_fit."""

import pytest

from openpyxl_toolkit import _metrics

BUILT_IN = [
    "calibri",
    "aptos",
    "arial",
    "times new roman",
    "courier new",
    "cambria",
    "verdana",
    "georgia",
    "tahoma",
    "futura",
]


@pytest.mark.parametrize("font", BUILT_IN)
@pytest.mark.parametrize("count", [1, 10, 20])
def test_a_width_of_n_fits_n_digits(font, count):
    """The definition of the unit: a column of width N holds N digits, plus padding."""
    widths, _ = _metrics.widths_for(font)
    padding = _metrics.CELL_PADDING_PIXELS / _metrics.max_digit_pixels(11, widths)

    width = _metrics.column_width("0" * count, 11, 11, font, font)

    assert width == pytest.approx(count + padding, abs=0.01)


def test_calibri_max_digit_width_matches_the_documented_value():
    """Excel documents 7 pixels for Calibri 11 at 96 DPI."""
    widths, _ = _metrics.widths_for("calibri")
    assert _metrics.max_digit_pixels(11, widths) == 7


def test_helvetica_and_arial_share_a_table():
    """Arial was drawn to Helvetica's advances; one table serves both."""
    assert _metrics.widths_for("helvetica")[0] is _metrics.widths_for("arial")[0]


def test_font_names_are_matched_case_and_space_insensitively():
    for name in ("Times New Roman", "  times new roman  ", "TIMES NEW ROMAN"):
        widths, matched = _metrics.widths_for(name)
        assert matched
        assert widths is _metrics.widths_for("times new roman")[0]


def test_an_unknown_font_falls_back_to_the_wide_face_and_says_so():
    widths, matched = _metrics.widths_for("Comic Sans MS")
    assert not matched
    assert widths is _metrics.widths_for("verdana")[0]


def test_no_font_named_is_not_an_unknown_font():
    """None means the workbook named none, so Excel's default applies instead."""
    widths, matched = _metrics.widths_for(None)
    assert matched
    assert widths is _metrics.widths_for("calibri")[0]


def test_the_fallback_face_is_the_widest_table_there_is():
    """Only sound while it holds. A wider table added later should prompt a rethink."""
    mean = {
        font: sum(_metrics.widths_for(font)[0].values()) / len(_metrics.widths_for(font)[0])
        for font in BUILT_IN
    }

    assert max(mean, key=lambda font: mean[font]) == _metrics.UNKNOWN_FONT_NAME


def test_an_unknown_text_font_is_measured_wider_than_calibri_would_be():
    """Too narrow hides what a column holds; too wide only looks untidy."""
    unknown = _metrics.column_width("Representative", 11, 11, "Segoe UI", "Calibri")
    as_calibri = _metrics.column_width("Representative", 11, 11, "Calibri", "Calibri")

    assert unknown > as_calibri


def test_an_unknown_unit_font_does_not_narrow_the_column():
    """The unit falls back to Calibri, not to the wide face.

    A wider unit divides by more and hands back a narrower column, so using the
    fallback here would undo what it is for.
    """
    unknown_unit = _metrics.column_width("Representative", 11, 11, "Georgia", "Segoe UI")
    calibri_unit = _metrics.column_width("Representative", 11, 11, "Georgia", "Calibri")

    assert unknown_unit == calibri_unit


def test_every_built_in_font_matches_its_own_table():
    """A built-in that stopped matching would be measured as Verdana without a word."""
    assert [font for font in BUILT_IN if not _metrics.widths_for(font)[1]] == []


def test_courier_new_is_monospaced():
    widths, _ = _metrics.widths_for("courier new")
    assert len({widths[c] for c in "iWm0.@"}) == 1


def test_proportional_fonts_are_not_monospaced():
    widths, _ = _metrics.widths_for("calibri")
    assert widths["W"] > widths["i"] * 3


def test_a_larger_font_needs_more_of_the_same_unit():
    """The unit is the workbook's font, so bigger cell text takes more units."""
    small = _metrics.column_width("hello world", 11, 11, "calibri", "calibri")
    large = _metrics.column_width("hello world", 22, 11, "calibri", "calibri")
    assert large > small * 1.8


def test_every_table_covers_printable_ascii():
    for font in BUILT_IN:
        widths, _ = _metrics.widths_for(font)
        missing = [chr(c) for c in range(32, 127) if chr(c) not in widths]
        assert missing == [], f"{font} is missing {missing}"
