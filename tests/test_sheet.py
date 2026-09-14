"""Sheet-level settings and the cell-selection mechanics behind every style setter."""

import re
import zipfile

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
    WorksheetToolkit(sheet).set_zoom_scale(zoom_scale=85)

    assert roundtrip(sheet).sheet_view.zoomScale == 85


@pytest.mark.parametrize("zoom_scale", [10, 400])
def test_zoom_scale_bounds_are_inclusive(sheet, roundtrip, zoom_scale):
    WorksheetToolkit(sheet).set_zoom_scale(zoom_scale=zoom_scale)

    assert roundtrip(sheet).sheet_view.zoomScale == zoom_scale


@pytest.mark.parametrize("zoom_scale", [9, 401, 0, -100])
def test_zoom_scale_outside_the_allowed_span_is_rejected(sheet, zoom_scale):
    with pytest.raises(ValueError):
        WorksheetToolkit(sheet).set_zoom_scale(zoom_scale=zoom_scale)


def test_sheet_level_methods_return_the_toolkit_for_chaining(sheet):
    toolkit = WorksheetToolkit(sheet)

    chained = (
        toolkit.freeze_panes("B2").merge_cells(range_string="A1:B1").set_zoom_scale(zoom_scale=120)
    )

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


def test_naming_a_row_outside_the_data_styles_it(sheet, roundtrip):
    """Cells named explicitly are styled even where the sheet holds no data yet.

    This does extend the used range, and that is the intended trade: the only way
    to grow the sheet is to ask for it by name. Skipping those cells instead would
    make the call quietly do less than it was asked to.
    """
    sheet["A1"] = "only cell"

    WorksheetToolkit(sheet).set_font(rows=[1, 2], columns=[1, 2], bold=True)

    assert _bold_cells(roundtrip(sheet)) == {"A1", "A2", "B1", "B2"}


def test_the_default_selection_does_not_extend_the_sheet(sheet):
    """rows=None means the used range, and reading it must not grow it.

    set_column_best_fit and the selection defaults both read the used range, so an
    automatic call that quietly enlarged it would compound on every later call.
    """
    sheet["A1"] = "only cell"

    WorksheetToolkit(sheet).set_font(bold=True)

    assert (sheet.max_row, sheet.max_column) == (1, 1)


def test_union_selection_does_not_reach_rows_the_caller_never_asked_for(sheet, roundtrip):
    """Row 3 is neither in the requested rows nor in the sheet's data, so it must stay plain."""
    _grid(sheet, 2, 2)

    WorksheetToolkit(sheet).set_font(rows=4, columns=1, intersections_only=False, bold=True)

    assert _bold_cells(roundtrip(sheet)) == {"A1", "A2", "A4", "B4"}


def test_a_fill_across_a_merged_range_reaches_every_cell_in_the_file(sheet, tmp_path):
    """Checked against the saved file, not a reload.

    openpyxl replaces the non-anchor cells of a merge with MergedCell objects when it
    reads a workbook, and drops their style while doing so. The style is in the file:
    reading it back is what loses it, so a round-trip assertion would be testing
    openpyxl rather than this library.
    """
    toolkit = WorksheetToolkit(sheet)
    sheet["A1"] = "heading"
    toolkit.merge_cells(range_string="A1:C1")

    toolkit.set_fill(rows=1, columns=[1, 2, 3], fill_type="solid", start_color="#FFD966")

    path = tmp_path / "book.xlsx"
    sheet.parent.save(path)
    with zipfile.ZipFile(path) as archive:
        worksheet = archive.read("xl/worksheets/sheet1.xml").decode()
        styles = archive.read("xl/styles.xml").decode()

    row = re.search(r'<row r="1".*?</row>', worksheet, re.S).group()
    used = re.findall(r'<c r="([A-C]1)" s="(\d+)"', row)
    assert [coordinate for coordinate, _ in used] == ["A1", "B1", "C1"]
    assert len({style for _, style in used}) == 1, "the three cells disagree on style"

    # cellStyleXfs precedes cellXfs, so the cell's style index counts within cellXfs
    cell_formats = re.search(r"<cellXfs.*?</cellXfs>", styles, re.S).group()
    used_format = re.findall(r"<xf [^>]*/>", cell_formats)[int(used[0][1])]
    fill_index = int(re.search(r'fillId="(\d+)"', used_format).group(1))
    assert "FFFFD966" in re.findall(r"<fill>.*?</fill>", styles, re.S)[fill_index]


def test_unfreezing_leaves_no_pane_state_behind(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.freeze_panes("B2")

    toolkit.freeze_panes(None)

    reloaded = roundtrip(sheet)
    assert not reloaded.freeze_panes
    assert [selection.pane for selection in reloaded.sheet_view.selection] == [None]


def test_a_column_beyond_the_grid_is_rejected(sheet):
    """openpyxl creates such a cell unvalidated, and the workbook can then never be saved."""
    with pytest.raises(ValueError, match="column 20000"):
        WorksheetToolkit(sheet).set_font(rows=1, columns=20000, bold=True)


def test_a_row_beyond_the_grid_is_rejected(sheet):
    with pytest.raises(ValueError, match="row 2000000"):
        WorksheetToolkit(sheet).set_font(rows=2000000, columns=1, bold=True)


def test_a_column_of_zero_is_rejected(sheet):
    with pytest.raises(ValueError, match="column 0"):
        WorksheetToolkit(sheet).set_font(rows=1, columns=0, bold=True)


def test_rejecting_a_bad_index_leaves_the_workbook_saveable(sheet, roundtrip):
    """The point of rejecting early: a materialised out-of-range cell is unrecoverable."""
    _grid(sheet, 2, 2)
    toolkit = WorksheetToolkit(sheet)

    with pytest.raises(ValueError):
        toolkit.set_font(rows=1, columns=20000, bold=True)

    assert roundtrip(sheet)["A1"].value == "r1c1"


def test_the_dimension_setters_are_bounds_checked_too(sheet):
    """16385 is past Excel's limit but inside get_column_letter's, so only our check catches it."""
    toolkit = WorksheetToolkit(sheet)
    for call in (
        lambda: toolkit.set_column_width(columns=16385, width=12),
        lambda: toolkit.set_column_best_fit(columns=16385),
    ):
        with pytest.raises(ValueError, match="outside the worksheet"):
            call()


def test_set_row_height_is_bounds_checked(sheet):
    with pytest.raises(ValueError, match="outside the worksheet"):
        WorksheetToolkit(sheet).set_row_height(rows=2000000, height=20)


def test_best_fit_rejects_a_bad_column_before_materialising_it(sheet, roundtrip):
    """get_column_letter raises only after ws.cell has already made the sheet unsaveable."""
    _grid(sheet, 2, 2)
    with pytest.raises(ValueError):
        WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 20000])

    assert roundtrip(sheet)["A1"].value == "r1c1"


def test_a_fractional_index_is_rejected(sheet):
    """A float index writes a cell reference like A1.5 that openpyxl cannot reload."""
    toolkit = WorksheetToolkit(sheet)
    with pytest.raises(TypeError, match="must be an integer"):
        toolkit.set_font(rows=1.5, columns=1, bold=True)
    with pytest.raises(TypeError, match="must be an integer"):
        toolkit.set_row_height(rows=1.5, height=20)


def test_the_toolkit_rejects_anything_that_is_not_a_worksheet(sheet):
    """Passing the workbook is the easy slip, and it used to fail much later."""
    with pytest.raises(TypeError, match="expected a Worksheet, got Workbook"):
        WorksheetToolkit(sheet.parent)
    with pytest.raises(TypeError, match="got str"):
        WorksheetToolkit("Sheet1")


def test_unmerge_cells_undoes_a_merge(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    sheet["A1"] = "heading"
    toolkit.merge_cells(range_string="A1:C1")

    toolkit.unmerge_cells(range_string="A1:C1")

    assert list(roundtrip(sheet).merged_cells.ranges) == []


def test_unmerge_cells_accepts_coordinates_too(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    sheet["A1"] = "heading"
    toolkit.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)

    toolkit.unmerge_cells(start_row=1, start_column=1, end_row=1, end_column=3)

    assert list(roundtrip(sheet).merged_cells.ranges) == []


def test_unmerging_a_range_that_is_not_merged_is_reported(sheet):
    """Ignoring it would hide a typo in the range."""
    with pytest.raises(ValueError):
        WorksheetToolkit(sheet).unmerge_cells(range_string="A5:C5")


@pytest.mark.parametrize("method", ["merge_cells", "unmerge_cells"])
def test_a_range_with_no_arguments_says_what_is_missing(sheet, method):
    """openpyxl's own message for this is 'expected <class int>'."""
    with pytest.raises(ValueError, match="start_row, start_column, end_row, end_column"):
        getattr(WorksheetToolkit(sheet), method)()


def test_a_partial_range_says_which_corner_is_missing(sheet):
    with pytest.raises(ValueError, match="end_column"):
        WorksheetToolkit(sheet).merge_cells(start_row=1, start_column=1, end_row=1)


def test_the_unchanged_sentinel_reads_as_words(sheet):
    """It appears in every signature help() and an editor tooltip render."""
    import inspect

    from openpyxl_toolkit.worksheet_toolkit import _UNCHANGED

    assert repr(_UNCHANGED) == "<unchanged>"
    assert "<unchanged>" in str(inspect.signature(WorksheetToolkit.set_font))
    assert "object object at" not in str(inspect.signature(WorksheetToolkit.set_font))


def test_zoom_scale_is_keyword_only(sheet):
    with pytest.raises(TypeError):
        WorksheetToolkit(sheet).set_zoom_scale(85)
