"""The rows= and columns= selection arguments, shared by the methods that format cells.

None and an empty list are different things. None is an axis the caller did not
filter on, so it stands for every row or column in use. An empty list is a filter
that found nothing, and selects nothing. Reading the two as the same thing is how a
computed selection that came out empty formats the whole sheet instead.
"""

import pytest

from openpyxl_toolkit import WorksheetToolkit


def _grid(ws, rows=3, columns=3):
    for row in range(1, rows + 1):
        for column in range(1, columns + 1):
            ws.cell(row=row, column=column, value=f"r{row}c{column}")
    return ws


def _bold(ws):
    """Coordinates of every bold cell."""
    return {
        ws.cell(row=row, column=column).coordinate
        for row in range(1, ws.max_row + 1)
        for column in range(1, ws.max_column + 1)
        if ws.cell(row=row, column=column).font.b
    }


# Every method that takes rows=, with the arguments it needs to do something visible.
CELL_METHODS = [
    ("set_alignment", {"horizontal": "center"}),
    ("set_border", {"sides": "left", "style": "thin"}),
    ("set_fill", {"fill_type": "solid", "start_color": "#f4d2d3"}),
    ("set_font", {"bold": True}),
    ("set_number_format", {"number_format": "0.00"}),
]


@pytest.mark.parametrize(("method", "kwargs"), CELL_METHODS)
def test_an_empty_row_selection_touches_nothing(sheet, roundtrip, method, kwargs):
    _grid(sheet)
    before = {c.coordinate: c._style for row in sheet.iter_rows() for c in row}

    getattr(WorksheetToolkit(sheet), method)(rows=[], **kwargs)

    assert {c.coordinate: c._style for row in sheet.iter_rows() for c in row} == before


@pytest.mark.parametrize(("method", "kwargs"), CELL_METHODS)
def test_an_empty_column_selection_touches_nothing(sheet, method, kwargs):
    _grid(sheet)
    before = {c.coordinate: c._style for row in sheet.iter_rows() for c in row}

    getattr(WorksheetToolkit(sheet), method)(columns=[], **kwargs)

    assert {c.coordinate: c._style for row in sheet.iter_rows() for c in row} == before


def test_an_empty_row_selection_is_not_every_row(sheet, roundtrip):
    """The defect this guards: [] read as None formatted the whole sheet."""
    _grid(sheet)

    WorksheetToolkit(sheet).set_font(rows=[], bold=True)

    assert _bold(roundtrip(sheet)) == set()


def test_an_empty_selection_beats_a_named_one_on_the_other_axis(sheet, roundtrip):
    """Nothing crossed with column 1 is still nothing."""
    _grid(sheet)

    WorksheetToolkit(sheet).set_font(rows=[], columns=[1], bold=True)

    assert _bold(roundtrip(sheet)) == set()


def test_a_union_still_takes_the_axis_that_named_something(sheet, roundtrip):
    """Under the cross reading, no rows plus column 1 is column 1."""
    _grid(sheet)

    WorksheetToolkit(sheet).set_font(rows=[], columns=[1], intersections_only=False, bold=True)

    assert _bold(roundtrip(sheet)) == {"A1", "A2", "A3"}


@pytest.mark.parametrize(
    ("rows", "columns", "expected"),
    [
        ([1], None, {"A1", "B1", "C1"}),
        (None, [1], {"A1", "A2", "A3"}),
        ([1], [1], {"A1"}),
        (None, None, {"A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"}),
    ],
)
def test_none_still_means_every_one_in_use(sheet, roundtrip, rows, columns, expected):
    """An axis left as None is unfiltered, which the empty-list fix must not change."""
    _grid(sheet)

    WorksheetToolkit(sheet).set_font(rows=rows, columns=columns, bold=True)

    assert _bold(roundtrip(sheet)) == expected


def test_an_empty_selection_leaves_row_heights_alone(sheet, roundtrip):
    _grid(sheet)

    WorksheetToolkit(sheet).set_row_height(rows=[], height=40)

    # row_dimensions is a defaultdict, so asking for a row creates its entry.
    # Membership is the only question that does not answer itself.
    reloaded = roundtrip(sheet)
    assert [row for row in (1, 2, 3) if row in reloaded.row_dimensions] == []


@pytest.mark.parametrize(
    ("method", "kwargs"),
    [("set_column_width", {"width": 30}), ("set_column_best_fit", {})],
)
def test_an_empty_selection_leaves_column_widths_alone(sheet, roundtrip, method, kwargs):
    _grid(sheet)

    getattr(WorksheetToolkit(sheet), method)(columns=[], **kwargs)

    # column_dimensions is a defaultdict too, and its default width is 13.0, so
    # reading one back manufactures the entry the assertion would then find.
    reloaded = roundtrip(sheet)
    assert [letter for letter in "ABC" if letter in reloaded.column_dimensions] == []


def test_a_range_string_is_unaffected(sheet, roundtrip):
    """cells= resolves to real rows and columns, so it never reaches the empty case."""
    _grid(sheet)

    WorksheetToolkit(sheet).set_font(cells="A1:B2", bold=True)

    assert _bold(roundtrip(sheet)) == {"A1", "B1", "A2", "B2"}
