"""Behaviour of ``WorksheetToolkit.set_alignment``, asserted after a real round trip."""

import pytest
from openpyxl.styles import Alignment

from openpyxl_toolkit import WorksheetToolkit


def fill(sheet, rows, columns):
    """Give the sheet a used range, so max_row/max_column are well defined."""
    for row in range(1, rows + 1):
        for column in range(1, columns + 1):
            sheet.cell(row=row, column=column, value=f"r{row}c{column}")
    return sheet


def horizontals(sheet, rows, columns):
    return {
        (row, column): sheet.cell(row=row, column=column).alignment.horizontal
        for row in range(1, rows + 1)
        for column in range(1, columns + 1)
    }


def test_horizontal_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_alignment(rows=[1], columns=[1], horizontal="center")

    assert roundtrip(sheet)["A1"].alignment.horizontal == "center"


def test_vertical_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_alignment(rows=[1], columns=[1], vertical="top")

    assert roundtrip(sheet)["A1"].alignment.vertical == "top"


def test_wrap_text_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_alignment(rows=[1], columns=[1], wrap_text=True)

    assert roundtrip(sheet)["A1"].alignment.wrapText is True


def test_shrink_to_fit_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_alignment(rows=[1], columns=[1], shrink_to_fit=True)

    assert roundtrip(sheet)["A1"].alignment.shrinkToFit is True


def test_indent_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_alignment(rows=[1], columns=[1], indent=3)

    assert roundtrip(sheet)["A1"].alignment.indent == 3


def test_text_rotation_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_alignment(rows=[1], columns=[1], text_rotation=45)

    assert roundtrip(sheet)["A1"].alignment.textRotation == 45


def test_reading_order_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_alignment(rows=[1], columns=[1], reading_order=2)

    assert roundtrip(sheet)["A1"].alignment.readingOrder == 2


def test_changing_one_attribute_preserves_every_other_one(sheet, roundtrip):
    """A later call changes only the attributes it names."""
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_alignment(
        rows=[1],
        columns=[1],
        horizontal="center",
        vertical="top",
        wrap_text=True,
        shrink_to_fit=True,
        indent=2,
        text_rotation=30,
        reading_order=1,
    )

    toolkit.set_alignment(rows=[1], columns=[1], vertical="bottom")

    alignment = roundtrip(sheet)["A1"].alignment
    assert (
        alignment.horizontal,
        alignment.vertical,
        alignment.wrapText,
        alignment.shrinkToFit,
        alignment.indent,
        alignment.textRotation,
        alignment.readingOrder,
    ) == ("center", "bottom", True, True, 2, 30, 1)


def test_attributes_set_by_separate_calls_accumulate(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_alignment(rows=[1], columns=[1], horizontal="right")
    toolkit.set_alignment(rows=[1], columns=[1], wrap_text=True)
    toolkit.set_alignment(rows=[1], columns=[1], indent=4)

    alignment = roundtrip(sheet)["A1"].alignment
    assert (alignment.horizontal, alignment.wrapText, alignment.indent) == ("right", True, 4)


def test_no_selection_applies_to_every_cell_in_the_used_range(sheet, roundtrip):
    fill(sheet, rows=2, columns=2)

    WorksheetToolkit(sheet).set_alignment(horizontal="center")

    assert set(horizontals(roundtrip(sheet), 2, 2).values()) == {"center"}


def test_rows_alone_select_whole_rows(sheet, roundtrip):
    fill(sheet, rows=3, columns=3)

    WorksheetToolkit(sheet).set_alignment(rows=[1, 3], horizontal="center")

    selected = {
        cell for cell, value in horizontals(roundtrip(sheet), 3, 3).items() if value == "center"
    }
    assert selected == {(1, 1), (1, 2), (1, 3), (3, 1), (3, 2), (3, 3)}


def test_columns_alone_select_whole_columns(sheet, roundtrip):
    fill(sheet, rows=3, columns=3)

    WorksheetToolkit(sheet).set_alignment(columns=[2], horizontal="center")

    selected = {
        cell for cell, value in horizontals(roundtrip(sheet), 3, 3).items() if value == "center"
    }
    assert selected == {(1, 2), (2, 2), (3, 2)}


def test_intersections_only_true_selects_the_grid_of_intersections(sheet, roundtrip):
    fill(sheet, rows=3, columns=3)

    WorksheetToolkit(sheet).set_alignment(
        rows=[1, 2], columns=[1, 2], intersections_only=True, horizontal="center"
    )

    selected = {
        cell for cell, value in horizontals(roundtrip(sheet), 3, 3).items() if value == "center"
    }
    assert selected == {(1, 1), (1, 2), (2, 1), (2, 2)}


def test_intersections_only_false_selects_the_union_of_rows_and_columns(sheet, roundtrip):
    """Union semantics cell by cell, including the corners left alone."""
    fill(sheet, rows=3, columns=3)

    WorksheetToolkit(sheet).set_alignment(
        rows=[2], columns=[2], intersections_only=False, horizontal="center"
    )

    assert horizontals(roundtrip(sheet), 3, 3) == {
        (1, 1): None,
        (1, 2): "center",
        (1, 3): None,
        (2, 1): "center",
        (2, 2): "center",
        (2, 3): "center",
        (3, 1): None,
        (3, 2): "center",
        (3, 3): None,
    }


def test_intersections_only_false_is_ignored_when_only_rows_are_given(sheet, roundtrip):
    fill(sheet, rows=2, columns=2)

    WorksheetToolkit(sheet).set_alignment(rows=[2], intersections_only=False, horizontal="center")

    selected = {
        cell for cell, value in horizontals(roundtrip(sheet), 2, 2).items() if value == "center"
    }
    assert selected == {(2, 1), (2, 2)}


def test_returns_the_toolkit_for_chaining(sheet):
    toolkit = WorksheetToolkit(sheet)

    assert toolkit.set_alignment(rows=[1], columns=[1], horizontal="center") is toolkit


def test_justify_last_line_survives_an_unrelated_alignment_change(sheet, roundtrip):
    sheet["A1"].alignment = Alignment(horizontal="left", justifyLastLine=True)

    WorksheetToolkit(sheet).set_alignment(rows=[1], columns=[1], horizontal="center")

    assert roundtrip(sheet)["A1"].alignment.justifyLastLine is True


def test_relative_indent_survives_an_unrelated_alignment_change(sheet, roundtrip):
    sheet["A1"].alignment = Alignment(horizontal="left", relativeIndent=2)

    WorksheetToolkit(sheet).set_alignment(rows=[1], columns=[1], horizontal="center")

    assert roundtrip(sheet)["A1"].alignment.relativeIndent == 2


def test_call_with_no_keyword_arguments_changes_nothing(sheet, roundtrip):
    original = Alignment(
        horizontal="left",
        vertical="bottom",
        wrap_text=True,
        indent=1,
        text_rotation=10,
        readingOrder=1,
        justifyLastLine=True,
        relativeIndent=2,
    )
    sheet["A1"].alignment = original

    WorksheetToolkit(sheet).set_alignment()

    assert roundtrip(sheet)["A1"].alignment == original


@pytest.mark.parametrize(
    ("argument", "attribute", "set_to"),
    [("indent", "indent", 3), ("reading_order", "readingOrder", 2)],
)
def test_none_clears_a_numeric_alignment_back_to_zero(
    sheet, roundtrip, argument, attribute, set_to
):
    """Neither has a null state, so None is translated to 0 rather than rejected."""
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_alignment(**{argument: set_to})
    toolkit.set_alignment(**{argument: None})

    assert getattr(roundtrip(sheet)["A1"].alignment, attribute) == 0


def test_clearing_an_indent_leaves_the_other_alignments_alone(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_alignment(horizontal="center", wrap_text=True, indent=4)
    toolkit.set_alignment(indent=None)

    alignment = roundtrip(sheet)["A1"].alignment

    assert (alignment.indent, alignment.horizontal, alignment.wrapText) == (0, "center", True)
