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
assigning it, which silently drops the size, the color, and anything else that
was set. This toolkit allows you to change attributes one by one without having
to copy previously set attributes.

```python
>>> from openpyxl import Workbook
>>> from openpyxl_toolkit import WorksheetToolkit

>>> workbook = Workbook()
>>> sheet = workbook.active
>>> sheet["A1"] = "Region"

>>> toolkit = WorksheetToolkit(sheet)
>>> toolkit.set_font(cells="A1", size=14, color="#1d3557")
>>> toolkit.set_font(cells="A1", bold=True) # size and color are not discarded
```

## Install

```bash
pip install openpyxl-toolkit
```

Python 3.10 or newer, and openpyxl 3.1 or newer.

## Choosing cells

Three ways, on every method that formats cells.

```python
>>> toolkit.set_fill(cells="A1:C3", fill_type="solid", start_color="#f4f6f8")
>>> toolkit.set_fill(cells="A:C", fill_type="solid", start_color="#f4f6f8")
>>> toolkit.set_fill(rows=[1, 2], columns=[1, 2], intersections_only=True,
...                  fill_type="solid", start_color="#ffff00")
```

`cells` takes a block (`"A1:C3"`), whole columns (`"B:D"`), whole rows (`"2:5"`)
or one cell (`"C3"`). Case does not matter, a reversed range such as `"C3:A1"`
is normalized, and an unbounded side is filled in from the used range, so
`"B:B"` means column B as far as the sheet goes rather than all 1,048,576 rows.

`rows` and `columns` take lists.

Given both rows and columns, `intersections_only` decides what they mean:

| | Cells formatted |
| --- | --- |
| `intersections_only=True` (default) | where the rows and the columns cross — a block |
| `intersections_only=False` | every cell in those rows **and** every cell in those columns — a cross |

Giving both `cells` and `rows`/`columns` raises.

### Column letters and numbers

`rows` and `columns` take numbers, so there is a pair for converting either way.

```python
>>> from openpyxl_toolkit import column_letter, column_index

>>> column_letter(3)    # 'C'
>>> column_index("C")   # 3
>>> toolkit.set_column_width(width=18, columns=[column_index("D")])
```

openpyxl has its own pair of methods in `openpyxl.utils`, but they run off the
end of the grid. `get_column_letter(16385)` returns `"XFE"`, and a workbook
using that column will not open. These stop at `XFD`, column 16,384.

## Methods

| Method | What it sets |
| --- | --- |
| `set_font` | name, size, bold, italic, underline, strike, color |
| `set_fill` | fill_type, start_color, end_color |
| `set_alignment` | horizontal, vertical, text_rotation, wrap_text, shrink_to_fit, indent, reading_order |
| `set_number_format` | how a value is displayed: currency, percent, dates, decimal places |
| `set_border` | style and color, on any of six sides |
| `set_outside_border` | one border around the edge of a block, leaving the inside alone |
| `set_column_width` / `set_row_height` | an explicit size; 0 hides the column or row |
| `set_column_best_fit` | a width that fits the widest cell |
| `merge_cells` / `unmerge_cells` | a merged range; unmerging one that is not merged does nothing |
| `freeze_panes` | rows above and columns left of a cell stay visible |
| `set_zoom_scale` | 10 to 400 |

Each returns the toolkit, so calls chain. Every parameter is keyword-only, apart
from `freeze_panes`. Anything not named is left as is. `None` is not the same as
leaving it out: `set_font(color=None)` clears the color, while omitting `color`
keeps whatever was there.

## Fitting columns

`set_column_best_fit` measures every character at its own width in the font
being used, so a column of i's does not come out as wide as a column of W's.
Excel's own formula turns the total into a width:

```
width = (pixels of text + 5 padding pixels) / max digit width
```

```python
>>> toolkit.set_column_best_fit(ignore_rows=[1], padding=1.5, min_width=9)
```

Built-in metrics cover Aptos, Arial, Calibri, Cambria, Courier New, Futura,
Georgia, Helvetica, Tahoma, Times New Roman, and Verdana. Any other face is
measured with Verdana's character widths unless `measure` is given. Verdana is
the widest of them, so a face with no metrics errs wide: a column that is too
narrow hides what it holds, while one that is too wide only looks untidy.

Three items to note:

- Text is assumed to be on one line. Wrapped text is not accounted for.
- Only dates and times are rendered as Excel displays them. Other number
  formats, including the ones `set_number_format` writes, are measured as the
  value is stored, so a currency column can come out narrower than it needs to
  be. `min_width` is the answer.
- Formula cells are skipped by default. Set those column widths directly, or
  pass `ignore_formulas=False` to override this behavior.

A merged title in row 1 will size column A to the whole title, because a merged
range stores its value in the top-left cell. Use `ignore_rows` in these cases.

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
