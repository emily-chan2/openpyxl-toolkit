"""Behavior of WorksheetToolkit.set_font, asserted after a real save/reload cycle."""

from copy import copy

import pytest
from openpyxl.styles import Font

from openpyxl_toolkit import WorksheetToolkit


@pytest.fixture
def grid(sheet):
    """A populated 3x3 block, so row and column selection hit more than one cell."""
    for row in range(1, 4):
        for column in range(1, 4):
            sheet.cell(row=row, column=column, value=f"r{row}c{column}")
    return sheet


def bold_map(ws):
    """The bold flag of every cell in the 3x3 block, as rows of three."""
    return [[ws.cell(row=r, column=c).font.bold for c in (1, 2, 3)] for r in (1, 2, 3)]


def test_size_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_font(size=18)
    assert roundtrip(sheet)["A1"].font.size == 18


def test_bold_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_font(bold=True)
    assert roundtrip(sheet)["A1"].font.bold is True


def test_italic_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_font(italic=True)
    assert roundtrip(sheet)["A1"].font.italic is True


def test_underline_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_font(underline="double")
    assert roundtrip(sheet)["A1"].font.underline == "double"


def test_strike_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_font(strike=True)
    assert roundtrip(sheet)["A1"].font.strike is True


def test_typeface_name_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_font(name="Georgia")
    assert roundtrip(sheet)["A1"].font.name == "Georgia"


def test_a_later_call_keeps_attributes_set_by_an_earlier_one(sheet, roundtrip):
    """Setting italic must not undo bold, size and name."""
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_font(bold=True, size=14, name="Georgia")
    toolkit.set_font(italic=True)

    font = roundtrip(sheet)["A1"].font
    assert (font.bold, font.size, font.name, font.italic) == (True, 14, "Georgia", True)


def test_a_later_call_keeps_a_color_set_by_an_earlier_one(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_font(color="#00FF00")
    toolkit.set_font(bold=True)

    assert roundtrip(sheet)["A1"].font.color.rgb == "FF00FF00"


def test_row_selection_reaches_every_column_of_that_row_only(grid, roundtrip):
    WorksheetToolkit(grid).set_font(rows=[2], bold=True)

    assert bold_map(roundtrip(grid)) == [
        [False, False, False],
        [True, True, True],
        [False, False, False],
    ]


def test_column_selection_reaches_every_row_of_that_column_only(grid, roundtrip):
    WorksheetToolkit(grid).set_font(columns=[3], bold=True)

    assert bold_map(roundtrip(grid)) == [
        [False, False, True],
        [False, False, True],
        [False, False, True],
    ]


def test_intersection_selection_reaches_only_the_shared_cells(grid, roundtrip):
    WorksheetToolkit(grid).set_font(rows=[1, 2], columns=[1, 2], bold=True)

    assert bold_map(roundtrip(grid)) == [
        [True, True, False],
        [True, True, False],
        [False, False, False],
    ]


def test_union_selection_reaches_the_whole_row_and_the_whole_column(grid, roundtrip):
    """With intersections_only=False the row and the column are both swept in full."""
    WorksheetToolkit(grid).set_font(rows=[1], columns=[1], intersections_only=False, bold=True)

    assert bold_map(roundtrip(grid)) == [
        [True, True, True],
        [True, False, False],
        [True, False, False],
    ]


def test_set_font_returns_the_toolkit_for_chaining(sheet):
    toolkit = WorksheetToolkit(sheet)
    assert toolkit.set_font(bold=True) is toolkit


def test_lowercase_hex_color_is_stored_as_upper_case_argb(sheet, roundtrip):
    WorksheetToolkit(sheet).set_font(color="#ff0000")
    assert roundtrip(sheet)["A1"].font.color.rgb == "FFFF0000"


def test_color_none_clears_a_previously_set_color(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_font(color="#FF0000")
    toolkit.set_font(color=None)

    assert roundtrip(sheet)["A1"].font.color is None


def test_a_call_with_no_style_arguments_leaves_the_font_untouched(sheet, roundtrip):
    """set_font() with nothing to set must produce the same file as never calling it.

    Both sides are copied out of their StyleProxy first: two proxies never compare
    equal to each other, even when they wrap identical fonts.
    """
    sheet["A1"] = "value"
    untouched = copy(roundtrip(sheet)["A1"].font)

    WorksheetToolkit(sheet).set_font()

    assert copy(roundtrip(sheet)["A1"].font) == untouched


def test_vert_align_survives_a_change_to_bold_alone(sheet, roundtrip):
    sheet["A1"].font = Font(vertAlign="superscript")

    WorksheetToolkit(sheet).set_font(bold=True)

    assert roundtrip(sheet)["A1"].font.vertAlign == "superscript"


def test_scheme_survives_a_change_to_bold_alone(sheet, roundtrip):
    sheet["A1"].font = Font(scheme="major")

    WorksheetToolkit(sheet).set_font(bold=True)

    assert roundtrip(sheet)["A1"].font.scheme == "major"


def test_name_none_puts_the_cell_back_on_the_default_face(sheet, roundtrip):
    """None clears, as it does for color: openpyxl's own default font has no name."""
    sheet["A1"].font = Font(name="Georgia", bold=True)

    WorksheetToolkit(sheet).set_font(name=None)

    reloaded = roundtrip(sheet)["A1"].font
    assert reloaded.name is None
    assert reloaded.bold is True


def test_size_none_puts_the_cell_back_on_the_default_size(sheet, roundtrip):
    sheet["A1"].font = Font(sz=18, bold=True)

    WorksheetToolkit(sheet).set_font(size=None)

    reloaded = roundtrip(sheet)["A1"].font
    assert reloaded.sz is None
    assert reloaded.bold is True
