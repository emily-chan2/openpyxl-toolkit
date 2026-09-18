"""The cells= range argument, shared by every method that names a range."""

import pytest
from openpyxl.styles import Side

from openpyxl_toolkit import WorksheetToolkit


def _grid(worksheet, rows=4, columns=4):
    for row in range(1, rows + 1):
        for column in range(1, columns + 1):
            worksheet.cell(row=row, column=column, value=f"r{row}c{column}")
    return worksheet


def _bold(worksheet, rows=4, columns=4):
    return {
        worksheet.cell(row=row, column=column).coordinate
        for row in range(1, rows + 1)
        for column in range(1, columns + 1)
        if worksheet.cell(row=row, column=column).font.bold
    }


@pytest.mark.parametrize(
    ("cells", "expected"),
    [
        ("A1:B2", {"A1", "A2", "B1", "B2"}),
        ("C3", {"C3"}),
        ("B:B", {"B1", "B2", "B3", "B4"}),
        ("B:C", {"B1", "B2", "B3", "B4", "C1", "C2", "C3", "C4"}),
        ("2:2", {"A2", "B2", "C2", "D2"}),
        ("2:3", {"A2", "B2", "C2", "D2", "A3", "B3", "C3", "D3"}),
        ("b2:a1", {"A1", "A2", "B1", "B2"}),
    ],
)
def test_the_three_range_spellings(sheet, roundtrip, cells, expected):
    """A block, whole columns and whole rows, plus a reversed and lowercased range."""
    _grid(sheet)

    WorksheetToolkit(sheet).set_font(cells=cells, bold=True)

    assert _bold(roundtrip(sheet)) == expected


def test_an_unbounded_side_stops_at_the_used_range(sheet, roundtrip):
    """'B:B' means column B as far as the sheet goes, not a million rows."""
    _grid(sheet, rows=2, columns=2)

    WorksheetToolkit(sheet).set_font(cells="B:B", bold=True)

    assert (sheet.max_row, sheet.max_column) == (2, 2)
    assert _bold(roundtrip(sheet)) == {"B1", "B2"}


@pytest.mark.parametrize("method", ["set_font", "set_fill", "set_border", "set_alignment"])
def test_every_cell_setter_accepts_it(sheet, method):
    _grid(sheet)
    arguments = {
        "set_font": {"bold": True},
        "set_fill": {"fill_type": "solid", "start_color": "#FFD966"},
        "set_border": {"style": "thin"},
        "set_alignment": {"horizontal": "center"},
    }[method]

    getattr(WorksheetToolkit(sheet), method)(cells="A1:B2", **arguments)


def test_cells_and_rows_together_are_rejected(sheet):
    _grid(sheet)
    with pytest.raises(ValueError, match="not both"):
        WorksheetToolkit(sheet).set_font(cells="A1:B2", rows=[1], bold=True)


@pytest.mark.parametrize("cells", ["", "   ", "nonsense", "A1:", "1A"])
def test_a_range_that_cannot_be_parsed_is_rejected(sheet, cells):
    _grid(sheet)
    with pytest.raises(ValueError):
        WorksheetToolkit(sheet).set_font(cells=cells, bold=True)


def test_merge_cells_accepts_it(sheet, roundtrip):
    sheet["A1"] = "heading"

    WorksheetToolkit(sheet).merge_cells(cells="A1:C1")

    assert [str(r) for r in roundtrip(sheet).merged_cells.ranges] == ["A1:C1"]


def test_unmerge_cells_accepts_it(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    sheet["A1"] = "heading"
    toolkit.merge_cells(cells="A1:C1")

    toolkit.unmerge_cells(cells="A1:C1")

    assert list(roundtrip(sheet).merged_cells.ranges) == []


def test_range_string_is_gone(sheet):
    """Replaced by cells. openpyxl still uses the name; the toolkit does not."""
    with pytest.raises(TypeError, match="range_string"):
        WorksheetToolkit(sheet).merge_cells(range_string="A1:C1")


def test_outside_border_accepts_it(sheet, roundtrip):
    _grid(sheet, rows=3, columns=3)

    WorksheetToolkit(sheet).set_outside_border(cells="A1:C3", style="thin")

    reloaded = roundtrip(sheet)
    assert reloaded["A1"].border.top.style == "thin"
    assert reloaded["C3"].border.right.style == "thin"
    interior = reloaded["B2"].border
    assert [
        side.style for side in (interior.top, interior.bottom, interior.left, interior.right)
    ] == [None] * 4


def test_outside_border_still_accepts_coordinates(sheet, roundtrip):
    _grid(sheet, rows=3, columns=3)

    WorksheetToolkit(sheet).set_outside_border(
        start_row=1, end_row=3, start_column=1, end_column=3, style="thin"
    )

    assert roundtrip(sheet)["A1"].border.top == Side(style="thin")


def test_outside_border_rejects_both_forms_at_once(sheet):
    _grid(sheet)
    with pytest.raises(ValueError, match="not both"):
        WorksheetToolkit(sheet).set_outside_border(
            cells="A1:C3", start_row=1, end_row=3, start_column=1, end_column=3, style="thin"
        )


def test_one_spelling_reaches_five_methods(sheet):
    """The same range string is accepted by every method that names a range."""
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)

    toolkit.set_font(cells="A1:C3", bold=True)
    toolkit.set_fill(cells="A1:C3", fill_type="solid", start_color="#EEEEEE")
    toolkit.set_border(cells="A1:C3", style="thin")
    toolkit.set_alignment(cells="A1:C3", horizontal="center")
    toolkit.set_outside_border(cells="A1:C3", style="medium")
    toolkit.merge_cells(cells="A1:C1")


@pytest.mark.parametrize("method", ["merge_cells", "unmerge_cells"])
def test_merge_rejects_both_forms_at_once(sheet, method):
    """Neither form wins: openpyxl silently discards the coordinates, this does not."""
    _grid(sheet)
    with pytest.raises(ValueError, match="not both"):
        getattr(WorksheetToolkit(sheet), method)(
            cells="A1:B2", start_row=3, start_column=3, end_row=4, end_column=4
        )


@pytest.mark.parametrize("method", ["merge_cells", "unmerge_cells"])
def test_even_one_stray_coordinate_is_rejected(sheet, method):
    _grid(sheet)
    with pytest.raises(ValueError, match="end_column"):
        getattr(WorksheetToolkit(sheet), method)(cells="A1:B2", end_column=4)


def test_every_range_method_agrees_on_mixing_the_two_forms(sheet):
    """set_font, set_outside_border and the merge pair behave the same way."""
    _grid(sheet)
    toolkit = WorksheetToolkit(sheet)
    calls = [
        lambda: toolkit.set_font(cells="A1:B2", rows=[3], bold=True),
        lambda: toolkit.set_outside_border(
            cells="A1:B2", start_row=3, end_row=4, start_column=3, end_column=4, style="thin"
        ),
        lambda: toolkit.merge_cells(
            cells="A1:B2", start_row=3, start_column=3, end_row=4, end_column=4
        ),
        lambda: toolkit.unmerge_cells(
            cells="A1:B2", start_row=3, start_column=3, end_row=4, end_column=4
        ),
    ]
    for call in calls:
        with pytest.raises(ValueError, match="not both"):
            call()
