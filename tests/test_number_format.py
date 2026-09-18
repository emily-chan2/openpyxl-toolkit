"""set_number_format: the code Excel uses to display a value."""

import pytest

from openpyxl_toolkit import WorksheetToolkit


def _grid(sheet, rows=3, columns=3):
    for row in range(1, rows + 1):
        for column in range(1, columns + 1):
            sheet.cell(row=row, column=column, value=row * 100 + column)


def test_a_format_survives_the_round_trip(sheet, roundtrip):
    _grid(sheet)

    WorksheetToolkit(sheet).set_number_format(number_format='"$"#,##0.00', cells="A1")

    assert roundtrip(sheet)["A1"].number_format == '"$"#,##0.00'


def test_a_range_takes_the_format_and_nothing_outside_it_does(sheet, roundtrip):
    _grid(sheet)

    WorksheetToolkit(sheet).set_number_format(number_format="0.00%", cells="A1:B2")

    ws = roundtrip(sheet)
    assert [ws[c].number_format for c in ("A1", "B1", "A2", "B2")] == ["0.00%"] * 4
    assert ws["C1"].number_format == "General"
    assert ws["A3"].number_format == "General"


def test_rows_and_columns_select_the_intersection(sheet, roundtrip):
    _grid(sheet)

    WorksheetToolkit(sheet).set_number_format(
        number_format="0.00", rows=[1, 2], columns=[2], intersections_only=True
    )

    ws = roundtrip(sheet)
    assert ws["B1"].number_format == "0.00"
    assert ws["B2"].number_format == "0.00"
    assert ws["A1"].number_format == "General"
    assert ws["B3"].number_format == "General"


def test_a_later_call_replaces_the_earlier_format(sheet, roundtrip):
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)

    toolkit.set_number_format(number_format="0.00", cells="A1")
    toolkit.set_number_format(number_format="0%", cells="A1")

    assert roundtrip(sheet)["A1"].number_format == "0%"


def test_setting_a_format_leaves_the_font_alone(sheet, roundtrip):
    """Formatting is merged here as everywhere else, not replaced."""
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_font(cells="A1", bold=True, size=14)

    toolkit.set_number_format(number_format="0.00", cells="A1")

    reloaded = roundtrip(sheet)["A1"]
    assert (reloaded.font.bold, reloaded.font.sz) == (True, 14.0)
    assert reloaded.number_format == "0.00"


def test_general_puts_a_cell_back_to_the_default(sheet, roundtrip):
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_number_format(number_format="0.00", cells="A1")

    toolkit.set_number_format(number_format="General", cells="A1")

    assert roundtrip(sheet)["A1"].number_format == "General"


def test_semicolons_hide_the_contents_without_losing_the_value(sheet, roundtrip):
    _grid(sheet)

    WorksheetToolkit(sheet).set_number_format(number_format=";;;", cells="A1")

    reloaded = roundtrip(sheet)["A1"]
    assert reloaded.number_format == ";;;"
    assert reloaded.value == 101


@pytest.mark.parametrize("bad", [None, 0, 14, 3.5, ["0.00"]])
def test_a_format_that_is_not_a_string_is_rejected(sheet, bad):
    """openpyxl takes the assignment, and the workbook then cannot be saved.

    The failure surfaces from the stylesheet writer with no mention of the cell
    or the value that caused it.
    """
    _grid(sheet)

    with pytest.raises(TypeError, match="number_format must be a string"):
        WorksheetToolkit(sheet).set_number_format(number_format=bad, cells="A1")


def test_the_rejected_workbook_would_really_not_have_saved(sheet, roundtrip):
    """The reason for the check above, exercised through openpyxl directly."""
    _grid(sheet)
    sheet["A1"].number_format = None

    with pytest.raises(TypeError):
        roundtrip(sheet)


def test_a_merged_range_takes_the_format_on_every_cell(sheet, roundtrip):
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)
    toolkit.merge_cells(cells="A1:B1")

    toolkit.set_number_format(number_format="0.00", cells="A1:B1")

    ws = roundtrip(sheet)
    assert ws["A1"].number_format == "0.00"


def test_it_returns_the_toolkit_for_chaining(sheet):
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)

    assert toolkit.set_number_format(number_format="0.00", cells="A1") is toolkit


def test_the_format_is_keyword_only(sheet):
    _grid(sheet)

    with pytest.raises(TypeError):
        WorksheetToolkit(sheet).set_number_format("0.00")


def test_a_format_is_required(sheet):
    _grid(sheet)

    with pytest.raises(TypeError, match="number_format"):
        WorksheetToolkit(sheet).set_number_format(cells="A1")
