"""Behaviour of ``WorksheetToolkit.set_fill``, asserted on the reloaded workbook."""

import pytest
from openpyxl.styles import Color, GradientFill, PatternFill

from openpyxl_toolkit import WorksheetToolkit


def _grid(ws, rows=3, columns=3):
    for row in range(1, rows + 1):
        for column in range(1, columns + 1):
            ws.cell(row=row, column=column, value=f"r{row}c{column}")
    return ws


def _painted(ws):
    """Coordinates of every cell carrying a pattern of any kind."""
    return {
        ws.cell(row=row, column=column).coordinate
        for row in range(1, ws.max_row + 1)
        for column in range(1, ws.max_column + 1)
        if ws.cell(row=row, column=column).fill.fill_type is not None
    }


def test_solid_fill_survives_the_round_trip(sheet, roundtrip):
    _grid(sheet)
    WorksheetToolkit(sheet).set_fill(rows=1, columns=1, fill_type="solid", start_color="#f4d2d3")

    fill = roundtrip(sheet)["A1"].fill
    assert (fill.fill_type, fill.start_color.rgb.upper()) == ("solid", "FFF4D2D3")


def test_fill_type_survives_a_later_colour_only_call(sheet, roundtrip):
    """Changing the colour must not throw away the pattern set by an earlier call."""
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_fill(rows=1, columns=1, fill_type="solid", start_color="#ff0000")
    toolkit.set_fill(rows=1, columns=1, start_color="#00ff00")

    fill = roundtrip(sheet)["A1"].fill
    assert (fill.fill_type, fill.start_color.rgb.upper()) == ("solid", "FF00FF00")


def test_start_colour_survives_a_later_fill_type_only_call(sheet, roundtrip):
    """Changing the pattern must not throw away the colour set by an earlier call."""
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_fill(rows=1, columns=1, fill_type="solid", start_color="#ff0000")
    toolkit.set_fill(rows=1, columns=1, fill_type="lightGrid")

    fill = roundtrip(sheet)["A1"].fill
    assert (fill.fill_type, fill.start_color.rgb.upper()) == ("lightGrid", "FFFF0000")


def test_end_colour_survives_a_later_start_colour_call(sheet, roundtrip):
    """The pattern background is the third attribute that must be merged, not reset."""
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_fill(
        rows=1, columns=1, fill_type="lightGrid", start_color="#ff0000", end_color="#0000ff"
    )
    toolkit.set_fill(rows=1, columns=1, start_color="#00ff00")

    fill = roundtrip(sheet)["A1"].fill
    assert fill.end_color.rgb.upper() == "FF0000FF"


def test_intersections_only_paints_the_intersection(sheet, roundtrip):
    _grid(sheet)
    WorksheetToolkit(sheet).set_fill(
        rows=[1, 2],
        columns=[1, 3],
        intersections_only=True,
        fill_type="solid",
        start_color="#ff0000",
    )

    assert _painted(roundtrip(sheet)) == {"A1", "A2", "C1", "C2"}


def test_union_selection_paints_whole_rows_and_whole_columns(sheet, roundtrip):
    _grid(sheet)
    WorksheetToolkit(sheet).set_fill(
        rows=[1], columns=[1], intersections_only=False, fill_type="solid", start_color="#ff0000"
    )

    assert _painted(roundtrip(sheet)) == {"A1", "B1", "C1", "A2", "A3"}


def test_row_selection_without_columns_spans_the_used_range(sheet, roundtrip):
    _grid(sheet)
    WorksheetToolkit(sheet).set_fill(rows=2, fill_type="solid", start_color="#ff0000")

    assert _painted(roundtrip(sheet)) == {"A2", "B2", "C2"}


def test_empty_selection_paints_every_used_cell(sheet, roundtrip):
    _grid(sheet, rows=2, columns=2)
    WorksheetToolkit(sheet).set_fill(fill_type="solid", start_color="#ff0000")

    assert _painted(roundtrip(sheet)) == {"A1", "A2", "B1", "B2"}


def test_existing_theme_colour_is_kept_when_only_the_pattern_changes(sheet, roundtrip):
    _grid(sheet)
    sheet["A1"].fill = PatternFill(
        patternType="solid", fgColor=Color(theme=9, tint=0.6), bgColor=Color(indexed=64)
    )

    WorksheetToolkit(sheet).set_fill(rows=1, columns=1, fill_type="solid")

    colour = roundtrip(sheet)["A1"].fill.start_color
    assert (colour.type, colour.theme, colour.tint) == ("theme", 9, pytest.approx(0.6))


def test_existing_indexed_colour_is_kept_when_only_the_pattern_changes(sheet, roundtrip):
    _grid(sheet)
    sheet["A1"].fill = PatternFill(
        patternType="solid", fgColor=Color(indexed=5), bgColor=Color(indexed=64)
    )

    WorksheetToolkit(sheet).set_fill(rows=1, columns=1, fill_type="solid")

    colour = roundtrip(sheet)["A1"].fill.start_color
    assert (colour.type, colour.indexed) == ("indexed", 5)


def test_existing_automatic_colour_is_kept_when_only_the_pattern_changes(sheet, roundtrip):
    _grid(sheet)
    sheet["A1"].fill = PatternFill(
        patternType="solid", fgColor=Color(auto=True), bgColor=Color(indexed=64)
    )

    WorksheetToolkit(sheet).set_fill(rows=1, columns=1, fill_type="solid")

    colour = roundtrip(sheet)["A1"].fill.start_color
    assert (colour.type, colour.auto) == ("auto", True)


def test_a_gradient_filled_cell_can_be_given_a_solid_fill(sheet, roundtrip):
    _grid(sheet)
    sheet["A1"].fill = GradientFill(stop=("FFFFFFFF", "FFFF0000"))

    WorksheetToolkit(sheet).set_fill(rows=1, columns=1, fill_type="solid", start_color="#00ff00")

    fill = roundtrip(sheet)["A1"].fill
    assert (fill.fill_type, fill.start_color.rgb.upper()) == ("solid", "FF00FF00")


@pytest.mark.xfail(reason="a colour without a fill_type produces an invisible fill", strict=True)
def test_a_colour_on_its_own_produces_a_visible_fill(sheet, roundtrip):
    """The docstring's own example: fill row 1 with pink, no fill_type given."""
    _grid(sheet)
    WorksheetToolkit(sheet).set_fill(rows=1, start_color="#f4d2d3")

    assert roundtrip(sheet)["A1"].fill.fill_type == "solid"


def test_a_failing_call_leaves_no_cell_in_the_range_painted(sheet, roundtrip, monkeypatch):
    """A range is formatted as a unit: if one cell cannot be filled, none is."""
    _grid(sheet)
    attempts = []

    def fail_on_the_third_cell(*args, **kwargs):
        attempts.append(None)
        if len(attempts) == 3:
            raise RuntimeError("cell could not be filled")
        return PatternFill(*args, **kwargs)

    monkeypatch.setattr("openpyxl_toolkit.worksheet_toolkit.PatternFill", fail_on_the_third_cell)

    with pytest.raises(RuntimeError):
        WorksheetToolkit(sheet).set_fill(
            rows=1, columns=[1, 2, 3], fill_type="solid", start_color="#ff0000"
        )

    assert _painted(roundtrip(sheet)) == set()


def test_a_cell_does_not_share_a_colour_object_with_the_fill_it_was_read_from(sheet, roundtrip):
    """A Color is stored by reference; sharing one lets a later mutation repaint cells."""
    _grid(sheet)
    header = PatternFill(patternType="solid", fgColor=Color(rgb="FF3366CC"))
    sheet["A1"].fill = header

    WorksheetToolkit(sheet).set_fill(rows=1, columns=1, fill_type="lightGrid")
    header.start_color.rgb = "FFFF0000"

    assert roundtrip(sheet)["A1"].fill.start_color.rgb == "FF3366CC"


def test_a_gradient_survives_a_call_that_asks_for_nothing(sheet, roundtrip):
    """set_fill with no style arguments must not flatten a gradient into a blank fill."""
    _grid(sheet)
    sheet["A1"].fill = GradientFill(stop=("FFFFFFFF", "FFFF0000"))

    WorksheetToolkit(sheet).set_fill(rows=1, columns=1)

    assert roundtrip(sheet)["A1"].fill.tagname == "gradientFill"


def test_clearing_a_colour_is_rejected_rather_than_silently_repainting(sheet):
    """None has no meaning for a pattern fill: its default foreground is opaque black."""
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)
    for kwargs in ({"start_color": None}, {"end_color": None}):
        with pytest.raises(ValueError, match="cannot clear a colour"):
            toolkit.set_fill(rows=1, columns=1, **kwargs)
