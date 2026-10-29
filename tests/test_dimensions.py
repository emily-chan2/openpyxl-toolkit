"""Column widths, row heights and best-fit sizing, as observed in the saved file."""

from datetime import date, datetime, time

import pytest
from openpyxl.styles import Alignment, Font

from openpyxl_toolkit import WorksheetToolkit
from openpyxl_toolkit._text import displayed_text

EXCEL_MAX_COLUMN_WIDTH = 255


# --------------------------------------------------------------------------
# Characterisation: behavior that is currently correct and must stay correct
# --------------------------------------------------------------------------


def test_column_width_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_column_width(columns=[1], width=21.5)

    assert roundtrip(sheet).column_dimensions["A"].width == 21.5


def test_row_height_survives_the_round_trip(sheet, roundtrip):
    WorksheetToolkit(sheet).set_row_height(rows=[1], height=44.5)

    assert roundtrip(sheet).row_dimensions[1].height == 44.5


def test_one_column_is_still_written_as_a_list(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_width(columns=[1], width=21)
    toolkit.set_column_width(columns=[2], width=21)

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == ws.column_dimensions["B"].width


def test_one_row_is_still_written_as_a_list(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_row_height(rows=[1], height=33)
    toolkit.set_row_height(rows=[2], height=33)

    ws = roundtrip(sheet)
    assert ws.row_dimensions[1].height == ws.row_dimensions[2].height


def test_width_with_no_columns_applies_to_every_used_column(sheet, roundtrip):
    sheet["A1"] = "a"
    sheet["B1"] = "b"
    sheet["C1"] = "c"

    WorksheetToolkit(sheet).set_column_width(width=17)

    ws = roundtrip(sheet)
    assert [ws.column_dimensions[letter].width for letter in "ABC"] == [17, 17, 17]


def test_height_with_no_rows_applies_to_every_used_row(sheet, roundtrip):
    sheet["A1"] = "a"
    sheet["A2"] = "b"
    sheet["A3"] = "c"

    WorksheetToolkit(sheet).set_row_height(height=27)

    ws = roundtrip(sheet)
    assert [ws.row_dimensions[row].height for row in (1, 2, 3)] == [27, 27, 27]


def test_best_fit_makes_a_long_column_wider_than_a_short_one(sheet, roundtrip):
    sheet["A1"] = "a very long piece of text indeed"
    sheet["B1"] = "hi"

    WorksheetToolkit(sheet).set_column_best_fit()

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width > ws.column_dimensions["B"].width


def test_best_fit_skips_rows_listed_in_ignore_rows(sheet, roundtrip):
    """Column A's oversized header is ignored, so it fits like control column B."""
    sheet["A1"] = "H" * 40
    sheet["A2"] = "short"
    sheet["B2"] = "short"

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2], ignore_rows=[1])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == ws.column_dimensions["B"].width


def test_best_fit_skips_formula_cells_by_default(sheet, roundtrip):
    """Column A is sized by its literal cell only, like control column B."""
    sheet["A1"] = '=REPT("x",100)'
    sheet["A2"] = "abc"
    sheet["B2"] = "abc"

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == ws.column_dimensions["B"].width


def test_best_fit_measures_formula_text_when_ignore_formulas_is_false(sheet, roundtrip):
    sheet["A1"] = '=REPT("x",100)'
    sheet["A2"] = "abc"
    sheet["B2"] = "abc"

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2], ignore_formulas=False)

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width > ws.column_dimensions["B"].width


def test_best_fit_adds_padding_on_top_of_the_measured_width(sheet, roundtrip):
    sheet["A1"] = "abcde"
    sheet["B1"] = "abcde"

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_best_fit(columns=[1], padding=0)
    toolkit.set_column_best_fit(columns=[2], padding=10)

    ws = roundtrip(sheet)
    difference = ws.column_dimensions["B"].width - ws.column_dimensions["A"].width
    assert difference == pytest.approx(10)


def test_setting_a_row_height_leaves_an_earlier_column_width_intact(sheet, roundtrip):
    """A later call adds to the layout rather than replacing it."""
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_width(columns=[1], width=25)
    toolkit.set_row_height(rows=[1], height=40)

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == 25


def test_setting_one_column_width_leaves_another_column_width_intact(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_width(columns=[1], width=25)
    toolkit.set_column_width(columns=[2], width=9)

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == 25


def test_best_fit_leaves_an_explicit_width_on_an_untouched_column_intact(sheet, roundtrip):
    sheet["A1"] = "x"
    sheet["B1"] = "a much longer value"

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_width(columns=[1], width=42)
    toolkit.set_column_best_fit(columns=[2])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == 42


def test_dimension_setters_return_the_toolkit_for_chaining(sheet):
    toolkit = WorksheetToolkit(sheet)

    assert toolkit.set_column_width(columns=[1], width=10) is toolkit
    assert toolkit.set_row_height(rows=[1], height=10) is toolkit
    assert toolkit.set_column_best_fit(columns=[1]) is toolkit
    assert toolkit.set_row_best_fit(rows=[1]) is toolkit


# --------------------------------------------------------------------------
# Known defects: each test is what should pass once the bug is fixed
# --------------------------------------------------------------------------


def test_column_width_applies_to_non_anchor_merged_columns(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.merge_cells(cells="A1:D1")
    toolkit.set_column_width(columns=[1, 2, 3], width=15)

    ws = roundtrip(sheet)
    assert [ws.column_dimensions[letter].width for letter in "ABC"] == [15, 15, 15]


def test_best_fit_applies_to_non_anchor_merged_columns(sheet, roundtrip):
    sheet["A2"] = "a" * 20
    sheet["B2"] = "b" * 5
    sheet["C2"] = "c" * 10

    toolkit = WorksheetToolkit(sheet)
    toolkit.merge_cells(cells="A1:D1")
    toolkit.set_column_best_fit(columns=[1, 2, 3])

    ws = roundtrip(sheet)
    widths = [ws.column_dimensions[letter].width for letter in "ABC"]
    assert widths[0] > widths[2] > widths[1]


def test_best_fit_skips_a_wrapped_cell(sheet, roundtrip):
    """Wrapping exists so text conforms to the column, not the other way round."""
    sheet["A1"] = "word " * 40
    sheet["A1"].alignment = Alignment(wrap_text=True)
    sheet["A2"] = "Region"

    WorksheetToolkit(sheet).set_column_best_fit()

    fitted = roundtrip(sheet).column_dimensions["A"].width
    assert fitted < 12


def test_best_fit_measures_a_wrapped_cell_when_ignore_wrapped_is_false(sheet, roundtrip):
    sheet["A1"] = "word " * 40
    sheet["A1"].alignment = Alignment(wrap_text=True)
    sheet["A2"] = "Region"

    WorksheetToolkit(sheet).set_column_best_fit(ignore_wrapped=False)

    assert roundtrip(sheet).column_dimensions["A"].width > 80


def test_best_fit_skips_a_cell_merged_across_columns(sheet, roundtrip):
    """The banner is shown across the merge; only the top-left cell holds it."""
    sheet["A1"] = "Regional sales, third quarter 2026, by close date"
    sheet["A2"] = "North"
    toolkit = WorksheetToolkit(sheet)
    toolkit.merge_cells(cells="A1:G1")

    toolkit.set_column_best_fit(columns=[1])

    fitted = roundtrip(sheet).column_dimensions["A"].width
    assert fitted < 12


def test_best_fit_measures_a_merge_that_stays_within_one_column(sheet, roundtrip):
    """A vertical merge shows its text in that column, so it still sets the width."""
    sheet["A1"] = "A tall merged label"
    toolkit = WorksheetToolkit(sheet)
    toolkit.merge_cells(cells="A1:A3")

    toolkit.set_column_best_fit(columns=[1])

    assert roundtrip(sheet).column_dimensions["A"].width > 14


def test_best_fit_measures_a_merged_anchor_when_ignore_merged_is_false(sheet, roundtrip):
    sheet["A1"] = "Regional sales, third quarter 2026, by close date"
    sheet["A2"] = "North"
    toolkit = WorksheetToolkit(sheet)
    toolkit.merge_cells(cells="A1:G1")

    toolkit.set_column_best_fit(columns=[1], ignore_merged=False)

    assert roundtrip(sheet).column_dimensions["A"].width > 35


def test_best_fit_leaves_a_width_alone_when_every_cell_is_skipped(sheet, roundtrip):
    """Skipped is not the same as empty: a column of skipped cells is not shrunk."""
    sheet["A1"] = "word " * 40
    sheet["A1"].alignment = Alignment(wrap_text=True)
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_width(width=30, columns=[1])

    toolkit.set_column_best_fit(columns=[1])

    assert roundtrip(sheet).column_dimensions["A"].width == 30


def test_best_fit_handles_a_font_with_no_explicit_size(sheet, roundtrip):
    """A font inheriting its size fits like the same text at the default size."""
    sheet["A1"] = "hello world"
    sheet["A1"].font = Font(bold=True)
    sheet["B1"] = "hello world"

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == pytest.approx(
        ws.column_dimensions["B"].width, rel=0.3
    )


def test_best_fit_leaves_a_deliberate_width_on_an_empty_column_alone(sheet, roundtrip):
    sheet["A1"] = "data"

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_width(columns=[3], width=30)
    toolkit.set_column_best_fit(columns=[1, 3])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["C"].width == 30


def test_best_fit_sizes_a_date_from_its_displayed_form(sheet, roundtrip):
    sheet["A1"] = datetime(2026, 9, 10)
    sheet["A1"].number_format = "yyyy-mm-dd"
    sheet["B1"] = "2026-09-10"

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == pytest.approx(
        ws.column_dimensions["B"].width, abs=1.0
    )


def test_column_width_is_clamped_to_the_excel_maximum(sheet, roundtrip):
    WorksheetToolkit(sheet).set_column_width(columns=[1], width=300)

    assert roundtrip(sheet).column_dimensions["A"].width == EXCEL_MAX_COLUMN_WIDTH


def test_zero_column_width_collapses_the_column_in_the_saved_file(sheet, roundtrip):
    WorksheetToolkit(sheet).set_column_width(columns=[1], width=0)

    dimension = roundtrip(sheet).column_dimensions["A"]
    assert dimension.width == 0 or dimension.hidden


def test_zero_row_height_collapses_the_row_in_the_saved_file(sheet, roundtrip):
    WorksheetToolkit(sheet).set_row_height(rows=[1], height=0)

    dimension = roundtrip(sheet).row_dimensions[1]
    assert dimension.height == 0 or dimension.hidden


def test_negative_column_width_is_rejected(sheet):
    with pytest.raises((ValueError, TypeError)):
        WorksheetToolkit(sheet).set_column_width(columns=[1], width=-5)


def test_negative_row_height_is_rejected(sheet):
    with pytest.raises((ValueError, TypeError)):
        WorksheetToolkit(sheet).set_row_height(rows=[1], height=-5)


def test_row_zero_is_rejected(sheet):
    with pytest.raises((ValueError, TypeError)):
        WorksheetToolkit(sheet).set_row_height(rows=[0], height=20)


def test_column_width_without_a_width_is_rejected(sheet):
    with pytest.raises((TypeError, ValueError)):
        WorksheetToolkit(sheet).set_column_width(columns=[1])


def test_row_height_without_a_height_is_rejected(sheet):
    with pytest.raises((TypeError, ValueError)):
        WorksheetToolkit(sheet).set_row_height(rows=[1])


def test_best_fit_measures_an_unstyled_cell_at_the_workbook_default_size(sheet, roundtrip):
    """A cell with no font inherits the workbook's size, not a hardcoded 11."""
    sheet.parent._fonts[0] = Font(name="Calibri", sz=22)
    sheet["A1"] = "hello world"
    sheet["B1"] = "hello world"
    sheet["B1"].font = Font(sz=22)
    sheet["C1"] = "hello world"
    sheet["C1"].font = Font(sz=11)

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2, 3])

    ws = roundtrip(sheet)
    inherited = ws.column_dimensions["A"].width
    assert inherited == pytest.approx(ws.column_dimensions["B"].width)
    assert inherited > ws.column_dimensions["C"].width


def test_best_fit_widens_for_a_cell_font_larger_than_the_workbook_default(sheet, roundtrip):
    """The width unit is the workbook's font, so a bigger cell font needs more units."""
    sheet.parent._fonts[0] = Font(name="Calibri", sz=11)
    sheet["A1"] = "hello world"
    sheet["B1"] = "hello world"
    sheet["B1"].font = Font(sz=22)

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["B"].width > ws.column_dimensions["A"].width * 1.8


def test_best_fit_renders_a_time_in_twelve_hour_form_when_the_format_says_so(sheet, roundtrip):
    """'2:07 PM' is shorter than '14:07:03', and shorter still than the stored value."""
    sheet["A1"] = datetime(2026, 9, 5, 14, 7, 3)
    sheet["A1"].number_format = "h:mm am/pm"
    sheet["B1"] = "2:07 PM"

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == pytest.approx(
        ws.column_dimensions["B"].width, abs=1.0
    )


def test_best_fit_measures_a_general_formatted_value_as_stored(sheet, roundtrip):
    """Only dates are translated; anything else is measured as it is stored."""
    sheet["A1"] = 1234.5678
    sheet["B1"] = "1234.5678"

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == pytest.approx(ws.column_dimensions["B"].width)


def test_row_height_is_clamped_to_the_excel_maximum(sheet, roundtrip):
    WorksheetToolkit(sheet).set_row_height(rows=[1], height=500)

    assert roundtrip(sheet).row_dimensions[1].height == 409


def test_best_fit_matches_the_definition_of_a_column_width_unit(sheet, roundtrip):
    """A width of N means N digits fit, plus Excel's 5 pixels of cell padding."""
    for count in (1, 10, 20):
        sheet.cell(row=1, column=count, value="0" * count)

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 10, 20])

    ws = roundtrip(sheet)
    for count, letter in ((1, "A"), (10, "J"), (20, "T")):
        assert ws.column_dimensions[letter].width == pytest.approx(count + 5 / 7, abs=0.01)


def test_best_fit_distinguishes_narrow_from_wide_characters(sheet, roundtrip):
    sheet["A1"] = "i" * 10
    sheet["B1"] = "W" * 10

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["B"].width > ws.column_dimensions["A"].width * 3


def test_best_fit_honours_min_and_max_width(sheet, roundtrip):
    sheet["A1"] = "x"
    sheet["B1"] = "a very long piece of text indeed"

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2], min_width=12, max_width=20)

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == 12
    assert ws.column_dimensions["B"].width == 20


def test_best_fit_uses_the_metrics_of_the_cell_font(sheet, roundtrip):
    """Courier New is monospaced and wider than Calibri, so text needs more room."""
    sheet["A1"] = "iiiiiiiiii"
    sheet["B1"] = "iiiiiiiiii"
    sheet["B1"].font = Font(name="Courier New")

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["B"].width > ws.column_dimensions["A"].width * 1.5


@pytest.mark.parametrize(
    ("number_format", "shown"),
    [
        ("yyyy-mm-dd", "2026-09-05"),
        ("mm/dd/yyyy", "09/05/2026"),
        ("d/m/yyyy", "5/9/2026"),
        ("mmm d, yyyy", "Sep 5, 2026"),
        ("mmmm yyyy", "September 2026"),
        ("dddd", "Saturday"),
        ("hh:mm:ss", "14:07:03"),
        ("h:mm am/pm", "2:07 PM"),
        ("m/d/yy h:mm", "9/5/26 14:07"),
    ],
)
def test_a_date_is_measured_as_its_number_format_renders_it(sheet, roundtrip, number_format, shown):
    """The width must match the displayed text, whatever platform this runs on.

    The month-name tokens are the ones that break if a bare mm is matched first,
    and the unpadded hour is the one that breaks on Windows, where strftime has no
    %-I. Both are covered here rather than left to a manual check.
    """
    sheet["A1"] = datetime(2026, 9, 5, 14, 7, 3)
    sheet["A1"].number_format = number_format
    sheet["B1"] = shown

    WorksheetToolkit(sheet).set_column_best_fit(columns=[1, 2])

    ws = roundtrip(sheet)
    assert ws.column_dimensions["A"].width == pytest.approx(ws.column_dimensions["B"].width)


@pytest.mark.parametrize("number_format", ["mm", "m", "mm:ss;@"])
def test_a_time_with_a_month_format_falls_back_to_the_stored_value(sheet, number_format):
    """A lone m means month, and a time has none, so there is nothing to render.

    Every other token checks the value carries the part it asks for before
    rendering it. The m token could not, because it means minute after an hour
    token and month otherwise, and it read .month off a time and raised.
    """
    sheet["A1"] = time(14, 30)
    sheet["A1"].number_format = number_format

    assert displayed_text(sheet["A1"]) == str(time(14, 30))


def test_a_date_with_a_month_format_still_renders_the_month(sheet):
    sheet["A1"] = date(2026, 9, 15)
    sheet["A1"].number_format = "mm"

    assert displayed_text(sheet["A1"]) == "09"


# --------------------------------------------------------------------------
# Row best fit: a row is as tall as the lines its cells take
# --------------------------------------------------------------------------

SENTENCE = "the quick brown fox jumps over the lazy dog"

# Measured in Excel: wrapped in a column of width 30, this cell best fits to 64.0,
# which is four lines of 16.0.
MEASURED_NOTE = (
    "Shared with the cricket club until noon; kick-off moved back an hour and the "
    "far changing rooms are unavailable for the whole morning."
)


def _wrapped(cell, value):
    cell.value = value
    cell.alignment = Alignment(wrap_text=True)
    return cell


def test_row_best_fit_gives_one_line_the_height_excel_best_fits_it_to(sheet, roundtrip):
    """16.0 for a line of the default font, measured from Excel's own best fit.

    Not the 15.0 Excel uses as a default ROW height. That 15.0 is a row holding
    one line, and it is the one case that says nothing about what a second line
    costs; scaling it gives 13.5 a line and comes out short of every wrapped
    cell. See LINE_SPACING.
    """
    sheet["A1"] = "hello"

    WorksheetToolkit(sheet).set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == 16.0


def test_row_best_fit_gives_a_twenty_point_line_the_height_excel_does(sheet, roundtrip):
    """Excel best fits a single line of 20 point text to 26.0.

    The second of the two measurements the line height rests on, and the one that
    shows a line is not proportional to its size: scaling 11 point's 16.0 up to 20
    gives 29.09, three points more than Excel draws. Part of a line is fixed and
    does not grow with the type.
    """
    sheet["A1"] = "Fixture notes"
    sheet["A1"].font = Font(sz=20)

    WorksheetToolkit(sheet).set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == 26.0


def test_row_best_fit_interpolates_a_size_between_the_measured_ones(sheet, roundtrip):
    """A size nobody measured lands between the two that were, not on a round number.

    14 point is three ninths of the way from 11 to 20, so its line is 16.0 plus
    three ninths of the 10.0 between them. Rounding a height to a whole point or a
    whole pixel would lose the third.
    """
    sheet["A1"] = "hello"
    sheet["A1"].font = Font(sz=14)

    WorksheetToolkit(sheet).set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == pytest.approx(16.0 + 10.0 / 3)


def test_row_best_fit_is_taller_for_a_larger_font(sheet, roundtrip):
    """A 16 point title needs more room than the 11 point cell under it."""
    sheet["A1"] = "title"
    sheet["A1"].font = Font(sz=16)
    sheet["A2"] = "body"

    WorksheetToolkit(sheet).set_row_best_fit()

    ws = roundtrip(sheet)
    assert ws.row_dimensions[1].height > ws.row_dimensions[2].height


def test_row_best_fit_takes_the_tallest_cell_rather_than_the_first(sheet, roundtrip):
    """Row 1 is sized by B1, which it reaches second, like control row 2."""
    sheet["A1"] = "body"
    sheet["B1"] = "title"
    sheet["B1"].font = Font(sz=16)
    sheet["A2"] = "title"
    sheet["A2"].font = Font(sz=16)

    WorksheetToolkit(sheet).set_row_best_fit()

    ws = roundtrip(sheet)
    assert ws.row_dimensions[1].height == ws.row_dimensions[2].height


def test_row_best_fit_handles_a_font_with_no_explicit_size(sheet, roundtrip):
    """Font(bold=True) carries no size, and the workbook's is what it inherits."""
    sheet.parent._fonts[0] = Font(name="Calibri", sz=22)
    sheet["A1"] = "hello"
    sheet["A1"].font = Font(bold=True)
    sheet["A2"] = "hello"
    sheet["A2"].font = Font(sz=22)

    WorksheetToolkit(sheet).set_row_best_fit()

    ws = roundtrip(sheet)
    assert ws.row_dimensions[1].height == ws.row_dimensions[2].height


def test_row_best_fit_grows_a_wrapped_cell_as_its_column_narrows(sheet, roundtrip):
    """The same sentence takes more lines in less room, and more lines is taller."""
    toolkit = WorksheetToolkit(sheet)
    for row, (column, width) in enumerate([(1, 60), (2, 20), (3, 10)], start=1):
        _wrapped(sheet.cell(row=row, column=column), SENTENCE)
        toolkit.set_column_width(columns=[column], width=width)

    toolkit.set_row_best_fit()

    ws = roundtrip(sheet)
    heights = [ws.row_dimensions[row].height for row in (1, 2, 3)]
    assert heights[0] < heights[1] < heights[2]


def test_row_best_fit_leaves_a_wrapped_cell_on_one_line_at_its_fitted_width(sheet, roundtrip):
    """The width set_column_best_fit chose is the width this measures against."""
    _wrapped(sheet["A1"], "a moderately long label")

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_best_fit(columns=[1], ignore_wrapped=False)
    toolkit.set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == 16.0


def test_row_best_fit_wraps_once_the_column_is_narrower_than_the_fitted_width(sheet, roundtrip):
    """Five of the pixels a width stands for are the cell's padding, not room for text.

    Counting them as room rather than as padding would leave this on one line, and
    the row too short for what the reader is shown.
    """
    _wrapped(sheet["A1"], "a moderately long label")

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_best_fit(columns=[1], ignore_wrapped=False)
    toolkit.set_column_width(columns=[1], width=sheet.column_dimensions["A"].width - 1)
    toolkit.set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == 32.0


def test_row_best_fit_wraps_at_the_full_width_of_the_column(sheet, roundtrip):
    """Excel best fits this cell to 64.0 at a column width of 30: four lines.

    The five pixels of padding in the column width formula are not room taken
    away from the text. Deducting them here breaks the line one word earlier --
    "hour" lands at 208 against a margin of 205 rather than 210 -- and makes
    this five lines, a quarter taller than Excel draws it.
    """
    _wrapped(sheet["A1"], MEASURED_NOTE)

    WorksheetToolkit(sheet).set_column_width(columns=[1], width=30).set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == 64.0


def test_row_best_fit_counts_a_line_break_in_a_wrapped_value(sheet, roundtrip):
    """Three lines of text that would otherwise fit on one."""
    _wrapped(sheet["A1"], "one\ntwo\nthree")

    WorksheetToolkit(sheet).set_column_width(columns=[1], width=40).set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == 48.0


def test_row_best_fit_adds_nothing_to_a_cell_beyond_its_lines(sheet, roundtrip):
    """Four lines is exactly four one-line rows: Excel charges no cell padding.

    Pinning the shape as well as the number. A model that adds a fixed amount per
    cell, or that multiplies a padded one-line height, agrees at one line and
    drifts further out with every line after it.
    """
    _wrapped(sheet["A1"], "one\ntwo\nthree\nfour")
    sheet["A3"] = "one"

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_width(columns=[1], width=40)
    toolkit.set_row_best_fit()

    ws = roundtrip(sheet)
    one_line = ws.row_dimensions[3].height
    assert ws.row_dimensions[1].height == 64.0
    assert ws.row_dimensions[1].height == 4 * one_line


def test_row_best_fit_draws_an_unwrapped_cell_on_one_line(sheet, roundtrip):
    """Excel spills an unwrapped cell sideways, line breaks and all, not down."""
    sheet["A1"] = "one\ntwo\nthree"
    sheet["A2"] = "x" * 200

    WorksheetToolkit(sheet).set_row_best_fit()

    ws = roundtrip(sheet)
    assert [ws.row_dimensions[row].height for row in (1, 2)] == [16.0, 16.0]


def test_row_best_fit_breaks_a_word_too_long_for_the_line(sheet, roundtrip):
    """No space to break at, so the break falls where the line runs out.

    A width of 10 is 65 pixels of text, and an 11 point Calibri x is 6 of them:
    ten to a line, so sixty take six lines, at 16.0 each.
    """
    _wrapped(sheet["A1"], "x" * 60)

    WorksheetToolkit(sheet).set_column_width(columns=[1], width=10).set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == 96.0


def test_row_best_fit_counts_the_spaces_between_two_words(sheet, roundtrip):
    """A run of spaces takes room on its line, so it can push the next word down."""
    _wrapped(sheet["A1"], "aaaa bbbb")
    _wrapped(sheet["A2"], "aaaa" + " " * 12 + "bbbb")
    sheet["B1"] = "aaaa bbbb"

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_best_fit(columns=[2])
    toolkit.set_column_width(columns=[1], width=sheet.column_dimensions["B"].width)
    toolkit.set_row_best_fit()

    ws = roundtrip(sheet)
    assert [ws.row_dimensions[row].height for row in (1, 2)] == [16.0, 32.0]


def test_row_best_fit_measures_an_unsized_column_at_excels_default_width(sheet, roundtrip):
    """A column with no width of its own takes 8.43, which is what the file means.

    openpyxl reports 13.0 for one, a placeholder it does not write to the file, and
    a cell wrapped against it is given room no reader has.
    """
    for row, column in ((1, 1), (2, 2), (3, 3)):
        _wrapped(sheet.cell(row=row, column=column), SENTENCE)

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_width(columns=[2], width=8.43)
    toolkit.set_column_width(columns=[3], width=13.0)
    toolkit.set_row_best_fit()

    ws = roundtrip(sheet)
    unsized, excel_default, openpyxl_placeholder = (
        ws.row_dimensions[row].height for row in (1, 2, 3)
    )
    assert unsized == excel_default
    assert unsized > openpyxl_placeholder


def test_row_best_fit_leaves_a_column_too_narrow_for_text_on_one_line(sheet, roundtrip):
    """A width with no room at all would otherwise wrap to a line per character."""
    _wrapped(sheet["A1"], "x" * 50)

    WorksheetToolkit(sheet).set_column_width(columns=[1], width=0.5).set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == 16.0


def test_row_best_fit_adds_padding_on_top_of_the_measured_height(sheet, roundtrip):
    sheet["A1"] = "abcde"
    sheet["A2"] = "abcde"

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_row_best_fit(rows=[1], padding=0)
    toolkit.set_row_best_fit(rows=[2], padding=6)

    ws = roundtrip(sheet)
    difference = ws.row_dimensions[2].height - ws.row_dimensions[1].height
    assert difference == pytest.approx(6)


def test_row_best_fit_skips_columns_listed_in_ignore_columns(sheet, roundtrip):
    """Row 1's oversized label is ignored, so it fits like control row 2."""
    sheet["A1"] = "title"
    sheet["A1"].font = Font(sz=24)
    sheet["B1"] = "body"
    sheet["B2"] = "body"

    WorksheetToolkit(sheet).set_row_best_fit(ignore_columns=[1])

    ws = roundtrip(sheet)
    assert ws.row_dimensions[1].height == ws.row_dimensions[2].height


def test_row_best_fit_skips_formula_cells_by_default(sheet, roundtrip):
    """Row 1 is sized by its literal cell only, like control row 2."""
    sheet["A1"] = "=A2"
    sheet["A1"].font = Font(sz=24)
    sheet["B1"] = "body"
    sheet["A2"] = "body"

    WorksheetToolkit(sheet).set_row_best_fit()

    ws = roundtrip(sheet)
    assert ws.row_dimensions[1].height == ws.row_dimensions[2].height


def test_row_best_fit_measures_formula_text_when_ignore_formulas_is_false(sheet, roundtrip):
    sheet["A1"] = "=A2"
    sheet["A1"].font = Font(sz=24)
    sheet["A2"] = "body"

    WorksheetToolkit(sheet).set_row_best_fit(ignore_formulas=False)

    ws = roundtrip(sheet)
    assert ws.row_dimensions[1].height > ws.row_dimensions[2].height


def test_row_best_fit_skips_a_cell_merged_down_rows(sheet, roundtrip):
    """The height a merge needs is shared by its rows, so no one of them grows."""
    _wrapped(sheet["A1"], "a" * 60)
    sheet["B1"] = "x"

    toolkit = WorksheetToolkit(sheet)
    toolkit.merge_cells(cells="A1:A3")
    toolkit.set_column_width(columns=[1], width=8)
    toolkit.set_row_best_fit(rows=[1])

    assert roundtrip(sheet).row_dimensions[1].height == 16.0


def test_row_best_fit_measures_a_merged_anchor_when_ignore_merged_is_false(sheet, roundtrip):
    _wrapped(sheet["A1"], "a" * 60)
    sheet["B1"] = "x"

    toolkit = WorksheetToolkit(sheet)
    toolkit.merge_cells(cells="A1:A3")
    toolkit.set_column_width(columns=[1], width=8)
    toolkit.set_row_best_fit(rows=[1], ignore_merged=False)

    assert roundtrip(sheet).row_dimensions[1].height > 16.0


def test_row_best_fit_wraps_a_merge_at_the_width_of_the_whole_merge(sheet, roundtrip):
    """Three merged columns are three columns of room, not the one holding the value."""
    _wrapped(sheet["A1"], SENTENCE)
    _wrapped(sheet["A3"], SENTENCE)

    toolkit = WorksheetToolkit(sheet)
    toolkit.merge_cells(cells="A1:C1")
    toolkit.set_column_width(columns=[1, 2, 3], width=10)
    toolkit.set_row_best_fit()

    ws = roundtrip(sheet)
    assert ws.row_dimensions[1].height < ws.row_dimensions[3].height


def test_row_best_fit_honours_min_and_max_height(sheet, roundtrip):
    sheet["A1"] = "x"
    _wrapped(sheet["A2"], SENTENCE)

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_column_width(columns=[1], width=8)
    toolkit.set_row_best_fit(min_height=24, max_height=40)

    ws = roundtrip(sheet)
    assert ws.row_dimensions[1].height == 24
    assert ws.row_dimensions[2].height == 40


def test_row_best_fit_is_clamped_to_the_excel_maximum(sheet, roundtrip):
    _wrapped(sheet["A1"], "word " * 400)

    WorksheetToolkit(sheet).set_column_width(columns=[1], width=8).set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == 409


def test_row_best_fit_leaves_a_deliberate_height_on_an_empty_row_alone(sheet, roundtrip):
    sheet["A2"] = "x"

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_row_height(rows=[1], height=40)
    toolkit.set_row_best_fit()

    assert roundtrip(sheet).row_dimensions[1].height == 40


def test_row_best_fit_leaves_the_rows_it_was_not_given_alone(sheet, roundtrip):
    sheet["A1"] = "title"
    sheet["A1"].font = Font(sz=24)
    sheet["A2"] = "title"
    sheet["A2"].font = Font(sz=24)

    toolkit = WorksheetToolkit(sheet)
    toolkit.set_row_height(rows=[1, 2], height=40)
    toolkit.set_row_best_fit(rows=[2])

    ws = roundtrip(sheet)
    assert ws.row_dimensions[1].height == 40
    assert ws.row_dimensions[2].height != 40


def test_row_best_fit_rejects_a_single_row_number(sheet):
    with pytest.raises(TypeError):
        WorksheetToolkit(sheet).set_row_best_fit(rows=3)


def test_row_best_fit_rejects_a_row_outside_the_grid_before_sizing_anything(sheet, roundtrip):
    """Every row is checked before the first cell is made, not as each one is reached.

    openpyxl raises for the bad row either way, but only once the rows ahead of it
    have been resized and their cells materialised, which cannot be undone.
    """
    sheet["A1"] = "title"
    sheet["A1"].font = Font(sz=24)
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_row_height(rows=[1], height=40)

    with pytest.raises(ValueError):
        toolkit.set_row_best_fit(rows=[1, 1_048_577])

    assert roundtrip(sheet).row_dimensions[1].height == 40
