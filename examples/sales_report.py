"""Build a small sales report with openpyxl-toolkit.

Run it from the repository root::

    python examples/sales_report.py

It writes ``sales_report.xlsx`` next to this file.
"""

from datetime import date
from pathlib import Path

from openpyxl import Workbook

from openpyxl_toolkit import WorksheetToolkit, column_letter

INK = "#1f2933"
PAPER = "#ffffff"
BAND = "#f4f6f8"
RULE = "#d7dde3"
HEADER_BG = "#1d3557"
HEADER_INK = "#ffffff"

STATUS_COLORS = {
    "Closed won": ("#e7f5ec", "#1b6b3a"),
    "In review": ("#fdf3dc", "#8a6100"),
    "At risk": ("#fdeaea", "#a12626"),
}

HEADINGS = ("Region", "Representative", "Closed", "Units", "Unit price", "Revenue", "Status")

DEALS = [
    ("North", "Priya Raghunathan", date(2026, 7, 6), 1200, 18.50, "Closed won"),
    ("North", "Tom Ek", date(2026, 7, 21), 340, 42.00, "In review"),
    ("South", "Marisol Delgado-Ruiz", date(2026, 7, 29), 2750, 9.25, "Closed won"),
    ("South", "Wen Li", date(2026, 8, 3), 95, 310.00, "At risk"),
    ("East", "Amara Okonkwo", date(2026, 8, 11), 1810, 16.75, "Closed won"),
    ("East", "Jonas Brandt", date(2026, 8, 19), 620, 27.40, "In review"),
    ("West", "Kiri Tamatea", date(2026, 8, 30), 4400, 6.10, "Closed won"),
    ("West", "Diego Sarmiento", date(2026, 9, 2), 150, 188.00, "At risk"),
    ("West", "Hana Sato", date(2026, 9, 9), 980, 23.90, "Closed won"),
]

TITLE_ROW = 1
SUBTITLE_ROW = 2
HEADING_ROW = 3
FIRST_DATA_ROW = HEADING_ROW + 1
LAST_DATA_ROW = FIRST_DATA_ROW + len(DEALS) - 1
TOTAL_ROW = LAST_DATA_ROW + 1
NOTE_ROW = TOTAL_ROW + 2


def write_values(sheet):
    """Put the text, numbers and formulas in place before any formatting."""
    sheet[f"A{TITLE_ROW}"] = "Regional sales"
    sheet[f"A{SUBTITLE_ROW}"] = "Third quarter 2026, by close date"

    for column, heading in enumerate(HEADINGS, start=1):
        sheet.cell(row=HEADING_ROW, column=column, value=heading)

    for offset, (region, rep, closed, units, price, status) in enumerate(DEALS):
        row = FIRST_DATA_ROW + offset
        sheet.cell(row=row, column=1, value=region)
        sheet.cell(row=row, column=2, value=rep)
        sheet.cell(row=row, column=3, value=closed).number_format = "dd mmm yyyy"
        sheet.cell(row=row, column=4, value=units).number_format = "#,##0"
        sheet.cell(row=row, column=5, value=price).number_format = '"$"#,##0.00'
        sheet.cell(row=row, column=6, value=f"=D{row}*E{row}").number_format = '"$"#,##0.00'
        sheet.cell(row=row, column=7, value=status)

    sheet[f"A{TOTAL_ROW}"] = "Total"
    units = sheet.cell(row=TOTAL_ROW, column=4, value=f"=SUM(D{FIRST_DATA_ROW}:D{LAST_DATA_ROW})")
    units.number_format = "#,##0"
    revenue = sheet.cell(row=TOTAL_ROW, column=6, value=f"=SUM(F{FIRST_DATA_ROW}:F{LAST_DATA_ROW})")
    revenue.number_format = '"$"#,##0.00'

    sheet[f"A{NOTE_ROW}"] = (
        "Revenue is units multiplied by unit price. Deals marked at risk are excluded "
        "from the committed forecast."
    )


def build(path):
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Q3 Sales"
    write_values(sheet)

    toolkit = WorksheetToolkit(sheet)
    last_column_index = 7
    last_column_letter = column_letter(last_column_index)

    # First 3 rows are banners across the table
    for row in (TITLE_ROW, SUBTITLE_ROW, NOTE_ROW):
        toolkit.merge_cells(
            start_row=row, end_row=row, start_column=1, end_column=last_column_index
        )

    # Title and subtitle fonts and alignment
    # Every setter can take a range string and calls can chain
    (
        toolkit.set_font(cells=f"A{TITLE_ROW}", size=18, bold=True, color=INK)
        .set_font(cells=f"A{SUBTITLE_ROW}", size=11, italic=True, color="#6b7280")
        .set_alignment(cells=f"A{TITLE_ROW}:{last_column_letter}{SUBTITLE_ROW}", vertical="center")
    )

    # The heading row fill, font, and alignment
    heading_range = f"A{HEADING_ROW}:{last_column_letter}{HEADING_ROW}"
    (
        toolkit.set_fill(cells=heading_range, fill_type="solid", start_color=HEADER_BG)
        .set_font(cells=heading_range, bold=True, size=11, color=HEADER_INK)
        .set_alignment(cells=heading_range, horizontal="left", vertical="center")
    )

    # Alternate the colors of the data rows
    banded = list(range(FIRST_DATA_ROW + 1, LAST_DATA_ROW + 1, 2))
    toolkit.set_fill(rows=banded, fill_type="solid", start_color=BAND)
    toolkit.set_fill(
        rows=[row for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1) if row not in banded],
        fill_type="solid",
        start_color=PAPER,
    )

    # Data row alignment
    data_rows = list(range(FIRST_DATA_ROW, LAST_DATA_ROW + 1))
    toolkit.set_alignment(
        rows=data_rows, columns=[4, 5, 6], intersections_only=True, horizontal="right"
    )
    toolkit.set_alignment(rows=data_rows, columns=[3], intersections_only=True, horizontal="center")

    # Status fills, fonts, and alignment
    for offset, deal in enumerate(DEALS):
        background, ink = STATUS_COLORS[deal[5]]
        cell = f"G{FIRST_DATA_ROW + offset}"
        toolkit.set_fill(cells=cell, fill_type="solid", start_color=background)
        toolkit.set_font(cells=cell, bold=True, size=10, color=ink)
        toolkit.set_alignment(cells=cell, horizontal="center")

    # A border under the headings and a border above the total
    toolkit.set_border(cells=heading_range, sides="bottom", style="medium", color=HEADER_BG)
    toolkit.set_border(
        cells=f"A{LAST_DATA_ROW}:{last_column_letter}{LAST_DATA_ROW}",
        sides="bottom",
        style="thin",
        color=RULE,
    )

    # The total row, bold on a tint. D:F is right aligned so the sums line up under
    # the figures above, and set_outside_border draws the perimeter of the range
    # rather than a box around every cell in it
    total_range = f"A{TOTAL_ROW}:{last_column_letter}{TOTAL_ROW}"
    (
        toolkit.set_font(cells=total_range, bold=True, color=INK)
        .set_fill(cells=total_range, fill_type="solid", start_color="#eef2f6")
        .set_alignment(cells=f"D{TOTAL_ROW}:F{TOTAL_ROW}", horizontal="right")
        .set_outside_border(cells=total_range, style="thin", color=HEADER_BG)
    )

    # The footnote, small and quiet. It shares the banner merge across all seven
    # columns, so wrap_text is what holds the text inside that width, and top
    # alignment keeps it against the rows above as the row grows to fit
    toolkit.set_font(cells=f"A{NOTE_ROW}", size=9, italic=True, color="#6b7280")
    toolkit.set_alignment(cells=f"A{NOTE_ROW}", wrap_text=True, vertical="top")

    # Fit the columns to what the reader sees. The banner rows are left out, or
    # column A would be sized to the whole title
    toolkit.set_column_best_fit(
        ignore_rows=[TITLE_ROW, SUBTITLE_ROW, NOTE_ROW],
        padding=1.5,
        min_width=9,
    )
    # The revenue column holds formulas, which best fit skips: the formula text is
    # not what the reader sees. Its width comes from the widest result instead
    toolkit.set_column_width(columns=[6], width=14)

    toolkit.freeze_panes(f"A{FIRST_DATA_ROW}")
    toolkit.set_zoom_scale(zoom_scale=120)
    sheet.sheet_view.showGridLines = False

    workbook.save(path)
    return path


if __name__ == "__main__":
    written = build(Path(__file__).with_name("sales_report.xlsx"))
    print(f"wrote {written}")
