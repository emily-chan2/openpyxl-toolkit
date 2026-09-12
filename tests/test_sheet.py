"""Sheet-level settings and the cell-selection mechanics behind every style setter."""

import pytest

from openpyxl_toolkit import WorksheetToolkit


def _grid(worksheet, rows, columns):
    """Give the worksheet a real used range of ``rows`` x ``columns`` populated cells."""
    for row in range(1, rows + 1):
        for column in range(1, columns + 1):
            worksheet.cell(row=row, column=column, value=f"r{row}c{column}")
    return worksheet


def _bold_cells(worksheet, through_row=5, through_column=5):
    """Coordinates of every bold cell in the top-left window, as written to the file."""
    return {
        worksheet.cell(row=row, column=column).coordinate
        for row in range(1, through_row + 1)
        for column in range(1, through_column + 1)
        if worksheet.cell(row=row, column=column).font.bold
    }


# --- sheet-level settings ----------------------------------------------------


def test_freeze_panes_at_b2_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).freeze_panes("B2")

    assert roundtrip(sheet).freeze_panes == "B2"


def test_merged_range_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).merge_cells(range_string="A1:C2")

    assert [str(cell_range) for cell_range in roundtrip(sheet).merged_cells.ranges] == ["A1:C2"]


def test_merging_by_coordinates_produces_the_same_range_as_the_range_string(sheet, roundtrip):
    WorksheetToolkit(sheet).merge_cells(start_row=1, start_column=1, end_row=2, end_column=3)

    assert [str(cell_range) for cell_range in roundtrip(sheet).merged_cells.ranges] == ["A1:C2"]


def test_zoom_scale_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_zoom_scale(85)

    assert roundtrip(sheet).sheet_view.zoomScale == 85


@pytest.mark.parametrize("zoom_scale", [10, 400])
def test_zoom_scale_bounds_are_inclusive(sheet, roundtrip, zoom_scale):
    WorksheetToolkit(sheet).set_zoom_scale(zoom_scale)

    assert roundtrip(sheet).sheet_view.zoomScale == zoom_scale


@pytest.mark.parametrize("zoom_scale", [9, 401, 0, -100])
def test_zoom_scale_outside_the_allowed_span_is_rejected(sheet, zoom_scale):
    with pytest.raises(ValueError):
        WorksheetToolkit(sheet).set_zoom_scale(zoom_scale)


def test_sheet_level_methods_return_the_toolkit_for_chaining(sheet):
    toolkit = WorksheetToolkit(sheet)

    chained = toolkit.freeze_panes("B2").merge_cells(range_string="A1:B1").set_zoom_scale(120)

    assert chained is toolkit


# --- cell selection, observed through set_font -------------------------------


def test_selecting_rows_only_touches_every_column_of_those_rows(sheet, roundtrip):
    _grid(sheet, 3, 3)

    WorksheetToolkit(sheet).set_font(rows=2, bold=True)

    assert _bold_cells(roundtrip(sheet)) == {"A2", "B2", "C2"}


def test_selecting_columns_only_touches_every_row_of_those_columns(sheet, roundtrip):
    _grid(sheet, 3, 3)

    WorksheetToolkit(sheet).set_font(columns=2, bold=True)

    assert _bold_cells(roundtrip(sheet)) == {"B1", "B2", "B3"}


def test_selecting_neither_rows_nor_columns_touches_the_whole_used_range(sheet, roundtrip):
    _grid(sheet, 2, 2)

    WorksheetToolkit(sheet).set_font(bold=True)

    assert _bold_cells(roundtrip(sheet)) == {"A1", "A2", "B1", "B2"}


def test_intersection_selection_touches_only_the_crossings(sheet, roundtrip):
    _grid(sheet, 3, 3)

    WorksheetToolkit(sheet).set_font(
        rows=[1, 2], columns=[1, 3], intersections_only=True, bold=True
    )

    assert _bold_cells(roundtrip(sheet)) == {"A1", "A2", "C1", "C2"}


def test_union_selection_touches_whole_rows_and_whole_columns(sheet, roundtrip):
    _grid(sheet, 3, 3)

    WorksheetToolkit(sheet).set_font(rows=1, columns=1, intersections_only=False, bold=True)

    assert _bold_cells(roundtrip(sheet)) == {"A1", "B1", "C1", "A2", "A3"}


def test_duplicate_and_unordered_selectors_touch_each_cell_once(sheet, roundtrip):
    _grid(sheet, 3, 3)

    WorksheetToolkit(sheet).set_font(rows=[3, 1, 1], columns=[2, 2], bold=True)

    assert _bold_cells(roundtrip(sheet)) == {"B1", "B3"}


def test_a_later_font_attribute_does_not_reset_an_earlier_one(sheet, roundtrip):
    """Overlapping selections must merge: the second call keeps the first call's bold."""
    _grid(sheet, 2, 3)
    toolkit = WorksheetToolkit(sheet)

    toolkit.set_font(rows=[1, 2], columns=[1, 2], bold=True)
    toolkit.set_font(rows=1, columns=[1, 2, 3], size=16)

    reloaded = roundtrip(sheet)
    assert [
        (reloaded.cell(row=1, column=c).font.bold, reloaded.cell(row=1, column=c).font.size)
        for c in (1, 2, 3)
    ] == [(True, 16.0), (True, 16.0), (False, 16.0)]


def test_a_later_selection_does_not_reset_cells_outside_it(sheet, roundtrip):
    """Cells the second call never selects keep everything the first call gave them."""
    _grid(sheet, 2, 3)
    toolkit = WorksheetToolkit(sheet)

    toolkit.set_font(rows=[1, 2], columns=[1, 2], bold=True)
    toolkit.set_font(rows=1, columns=[1, 2, 3], size=16)

    reloaded = roundtrip(sheet)
    assert [
        (reloaded.cell(row=2, column=c).font.bold, reloaded.cell(row=2, column=c).font.size)
        for c in (1, 2)
    ] == [(True, 11.0), (True, 11.0)]


# --- known defects -----------------------------------------------------------


@pytest.mark.xfail(reason="styling materialises cells and inflates the used range", strict=True)
def test_styling_a_range_does_not_grow_the_used_range(sheet):
    """The used range is what set_column_best_fit and the selection defaults read.

    Formatting must not make the sheet claim it holds more data than it does.
    """
    sheet["A1"] = "only cell"
    WorksheetToolkit(sheet).set_font(rows=[1, 2], columns=[1, 2], bold=True)

    assert (sheet.max_row, sheet.max_column) == (1, 1)


@pytest.mark.xfail(
    reason="union row sweep grows max_row, so the column sweep over-reaches", strict=True
)
def test_union_selection_does_not_reach_rows_the_caller_never_asked_for(sheet, roundtrip):
    """Row 3 is neither in the requested rows nor in the sheet's data, so it must stay plain."""
    _grid(sheet, 2, 2)

    WorksheetToolkit(sheet).set_font(rows=4, columns=1, intersections_only=False, bold=True)

    assert _bold_cells(roundtrip(sheet)) == {"A1", "A2", "A4", "B4"}


@pytest.mark.xfail(reason="only the merge anchor keeps its fill through a save", strict=True)
def test_a_fill_across_a_merged_range_survives_on_every_cell(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    sheet["A1"] = "heading"
    toolkit.merge_cells(range_string="A1:C1")

    toolkit.set_fill(rows=1, columns=[1, 2, 3], fill_type="solid", start_color="#FFD966")

    reloaded = roundtrip(sheet)
    assert [
        (
            reloaded.cell(row=1, column=c).fill.fill_type,
            reloaded.cell(row=1, column=c).fill.start_color.rgb,
        )
        for c in (1, 2, 3)
    ] == [("solid", "FFFFD966")] * 3


@pytest.mark.xfail(reason="unfreezing sets A1 and leaves orphan pane selections", strict=True)
def test_unfreezing_leaves_no_pane_state_behind(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.freeze_panes("B2")

    toolkit.freeze_panes(None)

    reloaded = roundtrip(sheet)
    assert not reloaded.freeze_panes
    assert [selection.pane for selection in reloaded.sheet_view.selection] == [None]
