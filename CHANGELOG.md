# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- **Breaking:** `set_column_best_fit` skips a cell whose text does not determine the
  width of the column holding it. Two kinds were being measured and should not have
  been. A wrapped cell was sized to its unwrapped line, which guaranteed it never
  wrapped: a 105 character note took a column 89 units wide. And a cell merged across
  columns was sized as if its text sat in the top-left cell alone, so a banner merged
  over seven columns widened the first one to hold all of it. Both are now left out,
  and `ignore_wrapped=False` and `ignore_merged=False` put either back. A merge that
  stays within one column is measured as before, its text being in that column.
- A column whose every cell was skipped keeps the width it had, rather than being cut
  to the padding. Skipped is not the same as empty, and the existing rule for an empty
  column now covers both.
- `ignore_rows` is no longer the answer to a merged banner, which was the reason the
  README gave for it. The argument is unchanged and still takes a row out for any
  reason of your own; it simply is not needed for this one. The example that used it
  for exactly that no longer does, and produces identical widths without it.
- `set_zoom_scale` takes its value positionally again, so `set_zoom_scale(85)` works
  alongside `set_zoom_scale(zoom_scale=85)`. 0.1.0 made it keyword-only for uniformity
  with the setters, but the name already says what the number is. Nothing breaks:
  allowing a positional only widens what is accepted. `freeze_panes` was always
  positional for the same reason, and the two are now consistent with each other.

### Removed

- **Breaking:** `set_column_best_fit` no longer takes a `measure` argument. It was a
  hook for supplying your own text measurement, meant for a font with no built-in
  table, and it did not earn its place. The callable was invoked for every cell rather
  than only the unmeasured ones, so overriding one face meant handling all sixteen, and
  the built-in measurement it would have delegated to was private, so there was no way
  to hand the rest back. A face with no table has three better answers: open an issue to
  have it added, set that column with `set_column_width`, or bound the result with
  `min_width` and `max_width`.
- The trade that removal makes: a licensed or in-house face that cannot be contributed
  upstream loses its one exact option, and falls back to setting widths directly. That
  is judged the right trade at sixteen faces and rising, but it is a trade rather than a
  free removal. Measurement itself is unchanged, and widths are identical for anyone who
  was not passing `measure`.

## [0.2.0] - 2026-09-23

### Added

- `column_letter` and `column_index`, exported from the package root, convert between a
  column's number and its letter: `column_letter(3)` is `'C'`, `column_index('C')` is 3.
  openpyxl has a pair of these in `openpyxl.utils`, but both run past the end of the
  grid -- `get_column_letter(16385)` returns `'XFE'` and
  `column_index_from_string('XFE')` returns 16385. Workbooks built on columns that far
  out will not open. These new methods stop at XFD, and report an index outside the grid
  the same way every other method does. A letter is read in either case and with
  surrounding whitespace, since neither can mean anything else; a cell reference such as
  `'C1'` raises rather than being read as its column.
- Six more faces in the built-in metrics for `set_column_best_fit`: Garamond, Inter,
  Open Sans, Palatino, Roboto and Segoe UI. Book Antiqua and Selawik are recognized as
  well, sharing the table of the face each matches. Sixteen faces are now measured
  exactly and twenty-seven names resolve to one of them; anything else still falls back
  to Verdana.
- Segoe UI is measured through Selawik, Microsoft's own openly licensed substitute for
  it, rather than from the proprietary file. The match was checked rather than taken on
  trust: all 188 characters Selawik carries have the advance Segoe UI gives them. Book
  Antiqua was checked the same way against Palatino, and shares every advance.

### Changed

- **Breaking:** `set_column_best_fit` measures a face it has no metrics for with Verdana's
  character widths rather than Calibri's. Calibri is the narrowest table there is and
  under-measured every other one, which made an unknown face the case most likely to be cut
  off; Verdana is the widest. The two errors are not equally bad -- a column measured too
  narrow hides what it holds, while one measured too wide only looks untidy -- so an
  unmeasurable face now errs wide. A column of such text comes out about a quarter wider
  than before. Faces with metrics are unaffected, as are `min_width` and `measure`.
- The face defining the width unit stays Calibri, and is now a separate constant from the
  one an unknown face falls back to. Excel's unit is one widest-digit of the workbook's
  normal font, so widening it divides by more and hands back a narrower column. A workbook
  whose normal font has no metrics would otherwise have been measured about a fifth
  narrower by the change above, rather than wider.

## [0.1.0] - 2026-09-22

The first release. Until now the project was one file, `worksheet_toolkit.py`, sitting
at the root of this repository to be copied into a project. The entries below are
written against that file, since it is the only version anyone has had.

### Added

- Packaged for distribution: `src/` layout, `pyproject.toml` (hatchling), and
  `openpyxl_toolkit` as the import name. `WorksheetToolkit` is re-exported from the
  package root, so `from openpyxl_toolkit import WorksheetToolkit` works. Python 3.10 or
  newer and openpyxl 3.1 or newer are required; the loose file had no floor of its own,
  so anyone on 3.9 stays on the copied version.
- Type hints across the public surface, with a `py.typed` marker, without which a
  checker ignores the annotations in an installed copy. Alignment names, border styles,
  fill patterns and underline styles are `Literal` types, so an editor completes them
  and reports a typo such as `horizontal='centre'` where openpyxl would have raised at
  the call. The rest are stricter than openpyxl is at runtime, which takes `bold='no'`
  and stores `True`, and `name=42` and stores `'42'`.
- `set_number_format` sets the code Excel uses to display a value: `'0.00'`,
  `'"$"#,##0.00'`, `'0.00%'`, `'yyyy-mm-dd'`, `';;;'` to show nothing. It takes the same
  `cells`, `rows` and `columns` arguments as the other setters. A format that is not a
  string raises `TypeError`: openpyxl accepts the assignment and the workbook then
  cannot be saved, with the failure surfacing from the stylesheet writer with no mention
  of the cell that caused it.
- `unmerge_cells`, taking the same arguments as `merge_cells`. Calling it on a range
  that is not merged does nothing; a range that cannot be parsed still raises.
- `min_width` and `max_width` on `set_column_best_fit`, for what the measurement cannot
  see: a currency column is measured from the stored value and comes out too narrow,
  while one long cell makes a column too wide to sit beside the others. A `measure` hook
  takes metrics for a face with no built-in table.
- `WorksheetToolkit` has a `__repr__`, so a chained call at a prompt echoes
  `<WorksheetToolkit 'Sheet1'>` rather than a memory address. The "leave this alone"
  default every styling parameter carries prints as `<unchanged>` for the same reason:
  `help()` and an editor's hover panel showed `<object object at 0x100d64a70>`, a
  different address each run.
- Test suite covering the save/reload round trip, and continuous integration across
  Python 3.10-3.14. The round trip is what the suite is built around, because several of
  the bugs below are invisible until the file is written out and read back.

### Changed

- **Breaking:** `set_column_best_fit` measures text with the real per-character advances
  of the font instead of treating every character as the same width, and uses Excel's own
  formula to convert that to a column width. A column of narrow letters is no longer as
  wide as one of capitals: ten `i`s and ten `W`s used to be given identical widths. Every
  existing call produces different widths. Built-in metrics cover Aptos, Calibri, Arial,
  Helvetica, Times New Roman, Courier New, Cambria, Verdana, Georgia, Tahoma and Futura;
  anything else falls back to Calibri or can be handled with the new `measure` argument.
- **Breaking:** `set_column_best_fit`'s `padding` defaults to 0. The old default of 2 was
  making up for the formula not counting Excel's own 5 pixels of cell padding, which it
  now does. Between this and the per-character measurement, a column of ordinary text
  sized with the defaults comes out 10-25% narrower than before. A column of wide
  capitals goes the other way, since its characters were the ones being under-measured.
- **Breaking:** `rows` and `columns` take a list. `rows=3` was accepted as a shorthand
  for `rows=[3]`, on the style setters and the dimension setters alike, and now raises
  `TypeError` naming the list form. A count and a single index are close enough to each
  other that accepting both hid the typo. Tuples, ranges, sets and generators still work.
- **Breaking:** an empty `rows` or `columns` selects every row or column on the style
  setters, whichever form the argument takes. An empty list already did; an empty
  iterator selected nothing, so a generator that filtered everything out formatted
  nothing and now formats every row. The dimension setters are unaffected: an empty
  selection there still does nothing.
- **Breaking:** `width`, `height` and `set_outside_border`'s `style` are required.
  Omitting a width or a height did nothing at all and said nothing about it. Omitting
  the style drew no border, but a `color` passed alongside it was still written, which
  repainted a border already there.
- **Breaking:** `set_zoom_scale` is keyword-only, like every other setter. It was the
  last one taking its value positionally, so `set_zoom_scale(85)` becomes
  `set_zoom_scale(zoom_scale=85)`. `freeze_panes(cell)` is not a setter and stays
  positional.
- **Breaking:** a foreground color with nothing to show it on is rejected rather than
  recorded. `set_fill(start_color=...)` on a cell with no pattern used to store a color
  that never appeared, while the identical call on an already-filled cell worked, so one
  line behaved differently depending on what was in the cell. `set_border(color=...)`
  with no style gets the same rule. The mirror case raises too: `fill_type='solid'` with
  no color painted the cell black, an unset foreground being ARGB `00000000`.
  `fill_type=None` still removes a fill and trips neither check.
- **Breaking:** an unknown border side is rejected. `sides=('lft',)` drew nothing and
  raised nothing.
- **Breaking:** a negative width or height raises instead of being written to the file,
  and a width or height of 0 hides the column or row. openpyxl's writer drops a falsy
  dimension, so hiding is the only way to honour what was asked for, and it survives a
  round trip.
- **Breaking:** row and column indexes outside Excel's grid are rejected, in the
  dimension setters as well as the style setters, against Excel's maximums of 1,048,576
  rows and 16,384 columns. openpyxl enforces only the row maximum, and only when creating
  a cell: an out-of-grid column, or an out-of-grid row reached through `set_row_height`,
  went into the file without complaint. A non-integer index raises `TypeError`: a float
  wrote a cell reference such as `A1.5`, and openpyxl could not read that workbook back.
- **Breaking:** `merge_cells` with missing coordinates says which ones instead of letting
  openpyxl report `expected <class 'int'>`, and raises `ValueError` where openpyxl raised
  `TypeError`. Code catching `TypeError` around a merge stops catching it; a stray
  `range_string` is still a `TypeError`, since Python raises that one.
- `cells` accepts a range string on every method that names a range: `set_font`,
  `set_fill`, `set_border`, `set_alignment`, `merge_cells`, `unmerge_cells` and
  `set_outside_border`. On the style setters and `set_outside_border` it takes a block
  (`"A1:C3"`), whole columns (`"B:D"`), whole rows (`"2:5"`) or a single cell (`"C3"`),
  and an unbounded side is filled in from the used range. `merge_cells` and
  `unmerge_cells` take a block or a single cell; an unbounded range reaches openpyxl and
  raises there. Mixing `cells` with `rows`, `columns` or the four coordinates raises,
  naming the arguments that would have been ignored.
- `set_alignment(indent=None)` and `set_alignment(reading_order=None)` set the value
  back to 0 instead of raising `TypeError` from inside openpyxl. Neither has a null
  state to be put back to, so None is translated rather than refused: the rest of the
  class treats None as clearing the thing it names, and these two were the exception.
- `WorksheetToolkit` raises `TypeError` if given anything other than a worksheet.
  Passing the workbook was accepted, then surfaced on the first method call as an
  `AttributeError` about a missing `cell` or `sheet_view` -- except `freeze_panes`, which
  raised nothing and set the attribute on the workbook, where nothing reads it.
- `set_column_width`, `set_row_height` and `set_outside_border` take their required
  argument first. All three are keyword-only, so this changes no call.

### Removed

- **Breaking:** `range_string` is gone from `merge_cells`. Use `cells` instead; it is the
  same string. A call still passing `range_string` raises `TypeError` rather than having
  the argument quietly ignored.

### Fixed

- `set_font` and `set_alignment` no longer reset the attributes they do not expose.
  They built a new `Font` or `Alignment` from their own arguments, which cleared
  `vertAlign`, `scheme`, `family`, `charset`, `outline`, `shadow`, `condense` and
  `extend` on a font, and `justifyLastLine` and `relativeIndent` on an alignment -- the
  opposite of the merging the library exists to do. Every call was affected, not only
  one on an unusual font: Excel's default carries `family` and `scheme`, so
  `set_font(bold=True)` on an untouched cell stripped both. Both methods copy the
  existing object and set only what they were given.
- `freeze_panes(None)` removes the split selections openpyxl leaves behind when a frozen
  pane goes away -- three of them after a freeze at a cell such as `B2`. A sheet left
  with those was what made Excel offer to repair the workbook. The pane itself was
  already being cleared: the old code passed `"A1"`, which openpyxl reads as unfreeze.
- `set_border` writes the color through the shared normalizer. It stripped the `#`
  itself instead, so a six-digit `'#ff0000'` reached openpyxl with no alpha byte and was
  stored as ARGB `00ff0000`, where the same input through `set_font` stored `FFff0000`:
  one color, two values in the same file. `set_outside_border` goes through `set_border`
  and was fixed with it. Every color the toolkit writes is now upper-cased as well, so
  the same input gives the same bytes whichever setter wrote it.
- `set_border` no longer rewrites both diagonal flags on a call touching either, so
  adding `diagonal_down` leaves an existing `diagonal_up` on, and `style=None` on a
  named diagonal takes that direction away as it does on a straight side. A cell holds
  one diagonal line and a flag per direction, so taking one direction away leaves the
  other still drawing that line, and the line goes only once neither is drawn. Naming a
  diagonal with no style no longer raises the flag on a cell that draws nothing.
- `set_fill` no longer raises `TypeError` on a cell whose existing fill uses a theme,
  indexed or automatic color. Reading `.rgb` off such a color returns openpyxl's
  descriptor rather than a string, so the `Color` itself is used instead -- copied, not
  shared, since a `Color` is stored by reference and sharing one lets a later mutation
  repaint every cell that inherited it.
- `set_fill` no longer raises `AttributeError` on a cell carrying a `GradientFill`. A
  call that asks for no change leaves the gradient alone, one naming both a `fill_type`
  and a color replaces it, and `fill_type=None` clears it. A color or a pattern on its
  own raises `ValueError`, a gradient having neither to supply the half left out.
- `set_fill` and `set_border` build every style object before assigning any of them, so
  a failure part-way through a range no longer leaves the worksheet half-formatted.
- `set_column_width` and `set_column_best_fit` no longer raise `AttributeError` when
  row 1 of a target column belongs to a merged range. They looked the column letter up
  through a cell, and a non-anchor `MergedCell` has no `column_letter`.
- `set_column_width` no longer extends the sheet's used range. Looking the letter up
  through `ws.cell()` created that cell, so widening an empty column pushed `max_column`
  out to it, and every later call defaulting to "everything in use" then covered the
  columns in between.
- `set_column_best_fit` no longer raises `TypeError` on a cell whose font has no explicit
  size, which `Font(bold=True)` alone is enough to produce. It falls back to the
  workbook's own default font size rather than assuming 11pt.
- `set_column_best_fit` sizes a date or time from what Excel displays rather than from
  the stored value, so a cell holding `2026-09-05 14:07:03` under `yyyy-mm-dd` is sized
  for ten characters rather than nineteen. Other number formats, including the ones
  `set_number_format` writes, are still measured as stored, so a currency column can come
  out narrower than it needs to be.
- `set_column_best_fit` skips a column it measured nothing in -- one that is empty, one
  holding only formulas under the default `ignore_formulas=True`, and one whose every row
  is in `ignore_rows`. All three measured zero and were written down to the bare padding,
  wiping a width set deliberately beforehand.
- Widths and heights are held to Excel's own maximums of 255 and 409. A long enough
  value sized a column past what Excel accepts.
- `set_font(color=None)` clears the font color instead of raising `AttributeError`.
  Omitting an argument is how to leave a value alone; `None` is a value in its own right
  and clears the thing it names. `set_fill` rejects a `None` color with `ValueError`,
  because a pattern fill has no colorless state -- pass `fill_type=None` to remove the
  fill instead.
- Naming rows and columns with `intersections_only=False` no longer formats rows nobody
  asked for. The sheet bounds are read once before the sweep; `ws.cell()` creates a cell
  that does not exist, so a named row past the end of the sheet grew `max_row` during
  the row pass and the column pass then swept down to it.

[0.2.0]: https://github.com/emily-chan2/openpyxl-toolkit/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/emily-chan2/openpyxl-toolkit/releases/tag/v0.1.0
