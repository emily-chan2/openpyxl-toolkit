# openpyxl-toolkit

[![PyPI](https://img.shields.io/pypi/v/openpyxl-toolkit?style=flat-square&color=1d3557)](https://pypi.org/project/openpyxl-toolkit/)
[![Python versions](https://img.shields.io/pypi/pyversions/openpyxl-toolkit?style=flat-square&color=1d3557)](https://pypi.org/project/openpyxl-toolkit/)
[![CI](https://img.shields.io/github/actions/workflow/status/emily-chan2/openpyxl-toolkit/ci.yml?branch=main&style=flat-square&label=ci&color=1d3557)](https://github.com/emily-chan2/openpyxl-toolkit/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT-1d3557?style=flat-square)](https://github.com/emily-chan2/openpyxl-toolkit/blob/main/LICENSE)

openpyxl formatting without the boilerplate.

## Purpose

It is our opinion that formatting and styling worksheets with openpyxl is clunky
and non-intuitive. We wanted to build something to make the process less painful
and the end result more readable.

For example, openpyxl stores a cell's style as an immutable object. This means
changing one attribute results in building an entirely new `Font` object and
assigning it, which drops the size, the color, and anything else that was set:

```python
>>> from copy import copy
>>> from openpyxl import Workbook
>>> from openpyxl.styles import Font

>>> wb = Workbook()
>>> ws = wb.active
>>> ws["A1"] = "My cell text"

>>> ws["A1"].font = Font(size=14, color="FF1D3557") # openpyxl wants ARGB
>>> ws["A1"].font = Font(bold=True) # the size and the color are now gone
```

If you try `ws["A1"].font.bold = True`, you get `Style objects are immutable and
cannot be changed. Reassign the style with a copy`. So if you wanted to modify
just one part of a font, you would have to make a copy, modify it, and assign
it:

```python
>>> font = copy(ws["A1"].font)
>>> font.bold = True
>>> ws["A1"].font = font
```

This toolkit allows you to change attributes one by one without having to copy
previously set attributes.

```python
>>> from openpyxl_toolkit import WorksheetToolkit

>>> toolkit = WorksheetToolkit(ws)
>>> toolkit.set_font(cells="A1", size=14, color="#1d3557")
>>> toolkit.set_font(cells="A1", bold=True) # size and color are not discarded
```

## Install

```bash
pip install openpyxl-toolkit
```

Python 3.10 or newer, and openpyxl 3.1 or newer.

## Methods

### Cell formatting

| Method | What it sets |
| --- | --- |
| `set_alignment` | horizontal, vertical, text_rotation, wrap_text, shrink_to_fit, indent, reading_order |
| `set_border` | style and color, on any of six sides |
| `set_outside_border` | one border around the edge of a block, leaving the inside alone |
| `set_fill` | fill_type, start_color, end_color |
| `set_font` | name, size, bold, italic, underline, strike, color |
| `set_number_format` | how a value is displayed: currency, percent, dates, decimal places |

### Row and column formatting

| Method | What it sets |
| --- | --- |
| `set_column_width` / `set_row_height` | an explicit size; 0 hides the column or row |
| `set_column_best_fit` | a column width that fits the widest cell |
| `merge_cells` / `unmerge_cells` | a merged range |

### Sheet-level controls

| Method | What it sets |
| --- | --- |
| `freeze_panes` | rows above and columns left of a cell stay visible |
| `set_autofilter` | Excel's filter controls on a range |
| `set_gridline_visibility` | whether the grid between cells is drawn on screen |
| `set_sheet_visibility` | whether the sheet is shown, hidden, or hidden from the unhide list too |
| `set_tab_color` | the color of the sheet's tab |
| `set_zoom_scale` | the zoom level when the reader opens the file: 10 to 400 |

```python
>>> toolkit.set_tab_color("#1d3557")
>>> toolkit.set_gridline_visibility(visible=False)
>>> toolkit.set_autofilter(cells="A1:C3")
>>> toolkit.set_sheet_visibility(state="hidden")
```

Each method returns the toolkit, so calls can chain. Most parameters are
keyword-only.

Anything not named is left as is. `None` is not the same as leaving out an
argument: `set_font(color=None)` clears the color, while omitting `color`
keeps whatever was there.

## Cell selection

On every method that formats cells, there are 2 methods of selecting which cells
the action should be applied to:
1) `cells`
2) `rows` and `columns`

```python
>>> toolkit.set_fill(cells="A1:C3", fill_type="solid", start_color="#f4f6f8")
>>> toolkit.set_fill(cells="A:C", fill_type="solid", start_color="#f4f6f8")
>>> toolkit.set_fill(rows=[1, 2], columns=[1, 2], intersections_only=True,
...                  fill_type="solid", start_color="#ffff00")
```

`cells` takes a block (`"A1:C3"`), whole columns (`"B:D"`), whole rows (`"2:5"`)
or one cell (`"C3"`). Case does not matter, a reversed range such as `"C3:A1"`
is normalized, and an unbounded side is filled in from the used range, so
`"B:B"` means column B as far as the sheet goes (rather than all 1,048,576
rows).

`rows` and `columns` take lists.

Given both rows and columns, `intersections_only` decides what they mean:

| | Cells formatted |
| --- | --- |
| `intersections_only=True` (default) | where the rows and the columns cross — a block |
| `intersections_only=False` | every cell in those rows **and** every cell in those columns — a cross |

Giving both `cells` and `rows`/`columns` raises.

### Column letters and numbers

Because column letters can be hard to work with, we have provided a pair of
functions for conversion. openpyxl has its own versions in `openpyxl.utils`, but
under names that are more difficult to remember.

```python
>>> from openpyxl_toolkit import column_letter, column_index

>>> column_letter(3)    # 'C'
>>> column_index("C")   # 3
>>> toolkit.set_column_width(width=18, columns=[column_index("D")])
```

## Type hints

The package ships `py.typed`. Alignment names, border styles, fill patterns and
underline styles are `Literal` types, so an editor completes them and a checker
catches a typo before the code runs:

```
error: Argument "horizontal" to "set_alignment" has incompatible type "Literal['centre']"
```

openpyxl rejects the same typo, but only once the call runs; the `Literal` types
move it into the editor. Elsewhere the annotations really are stricter than
openpyxl is at runtime, which matters more than it sounds: openpyxl accepts
`bold="no"` and stores `True`, and `name=42` and stores `"42"`.

## License

MIT.
