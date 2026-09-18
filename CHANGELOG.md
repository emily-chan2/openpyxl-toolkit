# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `set_number_format` sets the code Excel uses to display a value: `'0.00'`, `'"$"#,##0.00'`,
  `'0.00%'`, `'yyyy-mm-dd'`, `';;;'` to show nothing. It takes the same `cells`, `rows` and
  `columns` arguments as the other setters. A format that is not a string raises `TypeError`:
  openpyxl accepts the assignment and the workbook then cannot be saved, with the failure
  surfacing from the stylesheet writer with no mention of the cell that caused it.
  `set_column_best_fit` still measures these cells as the value is stored rather than as they
  display.

### Fixed

- `set_fill` no longer raises `TypeError` on a cell whose existing fill uses a theme,
  indexed or automatic colour. Reading `.rgb` off such a colour returns openpyxl's
  descriptor rather than a string, so the `Color` itself is used instead -- copied, not
  shared, since a `Color` is stored by reference and sharing one lets a later mutation
  repaint every cell that inherited it.
- `set_fill` no longer raises `AttributeError` on a cell carrying a `GradientFill`. A
  call that asks for no change leaves the gradient alone; one that does asks replaces it.
- `set_fill` builds every fill before assigning any of them, so a failure part-way
  through a range no longer leaves the worksheet half-formatted.
- `set_column_width` and `set_column_best_fit` no longer raise `AttributeError` when
  row 1 of a target column belongs to a merged range. They looked the column letter up
  through a cell, and a non-anchor `MergedCell` has no `column_letter`.
- `set_column_best_fit` no longer raises `TypeError` on a cell whose font has no explicit
  size, which `Font(bold=True)` alone is enough to produce. It falls back to the
  workbook's own default font size rather than assuming 11pt.
- `set_font(color=None)` clears the font colour instead of raising `AttributeError`;
  `_UNCHANGED` remains the way to say "leave it alone". `set_fill` rejects a `None`
  colour with `ValueError`, because a pattern fill has no colourless state -- pass
  `fill_type=None` to remove the fill instead.
- Row and column indexes outside Excel's grid are rejected, in the dimension setters as
  well as the style setters. openpyxl created such cells without complaint and the
  workbook could then never be saved. Non-integer indexes raise `TypeError`: a float
  wrote a cell reference such as `A1.5` that openpyxl could not read back.

### Changed

- `cells=` accepts a range string on every method that names a range:
  `set_font`, `set_fill`, `set_border`, `set_alignment`, `merge_cells`,
  `unmerge_cells` and `set_outside_border`. It takes a block (`"A1:C3"`), whole
  columns (`"B:D"`), whole rows (`"2:5"`) or a single cell (`"C3"`), and an
  unbounded side is filled in from the used range. Mixing it with `rows=`,
  `columns=` or the four coordinates raises.
- **Breaking:** `range_string` is removed from `merge_cells` and
  `unmerge_cells`. Use `cells=` instead; it is the same string.

- `set_zoom_scale` is keyword-only, like every other method. `set_zoom_scale(85)`
  becomes `set_zoom_scale(zoom_scale=85)`.
- `set_column_width`, `set_row_height` and `set_outside_border` take their required
  argument first. Every method is keyword-only, so this changes no call.
- `WorksheetToolkit` raises `TypeError` if given anything other than a worksheet.
  Passing the workbook used to surface much later as an `AttributeError` from
  inside a private method.
- `merge_cells` with no arguments says which coordinates are missing instead of
  letting openpyxl report `expected <class 'int'>`.

### Added

- `unmerge_cells`, taking the same arguments as `merge_cells`. Calling it on a
  range that is not merged does nothing; a range that cannot be parsed still raises.

- `set_column_best_fit` measures text with the real per-character advances of the
  font instead of treating every character as the same width, and uses Excel's own
  formula to convert that to a column width. Text columns come out 10-25% narrower,
  and a column of narrow letters is no longer as wide as one of capitals. Built-in
  metrics cover Aptos, Calibri, Arial, Helvetica, Times New Roman, Courier New,
  Cambria, Verdana, Georgia, Tahoma and Futura; anything else falls back to Calibri
  or can be handled with the new `measure` argument.
- `set_column_best_fit`'s `padding` now defaults to 0. The old default of 2 was
  making up for the formula not counting Excel's own 5 pixels of cell padding,
  which it now does.

### Added

- `min_width` and `max_width` on `set_column_best_fit`, and a `measure` hook for
  supplying metrics for a face with no built-in table.

- Packaged the project for distribution: `src/` layout, `pyproject.toml` (hatchling),
  and `openpyxl_toolkit` as the import name. `WorksheetToolkit` is re-exported from the
  package root, so `from openpyxl_toolkit import WorksheetToolkit` works.
- Test suite covering the save/reload round trip, and continuous integration across
  Python 3.10-3.14.

[Unreleased]: https://github.com/emily-chan2/openpyxl-toolkit/compare/HEAD
