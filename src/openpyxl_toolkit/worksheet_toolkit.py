from copy import copy
from datetime import date, datetime, time
from numbers import Number

from openpyxl.styles import DEFAULT_FONT, Border, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import _metrics

_UNCHANGED = object()

# Excel's grid limits. openpyxl will create a cell outside them without complaint,
# but the workbook can then never be saved, so they are rejected at the door.
# get_column_letter is not a substitute: it accepts up to ZZZ (18278).
_MAX_ROW = 1_048_576
_MAX_COLUMN = 16_384

# Excel refuses a column wider than this and clips a row taller than it.
_MAX_COLUMN_WIDTH = 255
_MAX_ROW_HEIGHT = 409

_BORDER_SIDES = ("top", "bottom", "left", "right", "diagonal_up", "diagonal_down")

# Excel date tokens, longest first so "yyyy" is matched before "yy". Anything not
# listed here is copied through, which is right for separators and literal text.
_DATE_TOKENS = (
    ("yyyy", "%Y"),
    ("yy", "%y"),
    ("mmmm", "%B"),
    ("mmm", "%b"),
    ("dddd", "%A"),
    ("ddd", "%a"),
    ("dd", "%d"),
    ("d", "%-d"),
    ("hh", "%H"),
    ("h", "%-H"),
    ("ss", "%S"),
    ("s", "%-S"),
    ("am/pm", "%p"),
)


def _check_indexes(indexes, limit, label):
    """Reject indexes openpyxl would accept but Excel cannot store."""
    for index in indexes:
        # bool is an int, and a float index writes a cell reference like "A1.5"
        # that openpyxl itself can no longer parse when reloading the file.
        if not isinstance(index, int) or isinstance(index, bool):
            raise TypeError(f"{label} index must be an integer, got {index!r}")
        if not 1 <= index <= limit:
            raise ValueError(
                f"{label} {index} is outside the worksheet: {label}s run from 1 to {limit}"
            )


def _check_bounds(rows=(), columns=()):
    _check_indexes(rows, _MAX_ROW, "row")
    _check_indexes(columns, _MAX_COLUMN, "column")


def _displayed_text(cell):
    """The text Excel shows in a cell, as far as it can be worked out cheaply.

    Only dates and times are translated. Excel's number formats are a language of
    their own, and a partial implementation of the rest would mis-measure in ways
    that are harder to notice than the raw value, so everything else is measured
    as it is stored.
    """
    value = cell.value
    if value is None:
        return ""
    if not isinstance(value, (datetime, date, time)):
        return str(value)

    code = (cell.number_format or "").lower()
    if not code or code == "general":
        return str(value)

    # An hour is written 12-hour when the code also carries AM/PM.
    twelve_hour = "am/pm" in code
    # A minute token looks identical to a month token; Excel tells them apart by
    # whether an hour came first, so track that while walking the code.
    pattern, index, after_hour = "", 0, False
    while index < len(code):
        for token, directive in _DATE_TOKENS:
            if code.startswith(token, index):
                if token in ("hh", "h"):
                    after_hour = True
                    if twelve_hour:
                        directive = "%I" if token == "hh" else "%-I"
                pattern += directive
                index += len(token)
                break
        else:
            if code.startswith("mm", index):
                pattern += "%M" if after_hour else "%m"
                index += 2
            elif code.startswith("m", index):
                pattern += "%-M" if after_hour else "%-m"
                index += 1
            else:
                pattern += code[index]
                index += 1

    try:
        return value.strftime(pattern)
    except ValueError:
        return str(value)


def _has_color(value):
    """True when a colour carries something the user chose.

    A cell that was never coloured reports the ARGB default rather than nothing,
    so an absent colour and an explicit one are only distinguishable by value.
    """
    if value is None:
        return False
    if isinstance(value, str):
        return value.upper() not in ("", "00000000")
    if getattr(value, "type", None) != "rgb":
        return True  # a theme, indexed or automatic colour is a real choice
    return (value.rgb or "").upper() not in ("", "00000000")


def _normalize_color(color):
    if color is _UNCHANGED:
        return _UNCHANGED
    if color is None:
        # None clears the colour; _UNCHANGED is how "leave it alone" is spelled.
        return None
    hex_color = color.lstrip("#").upper()
    if len(hex_color) == 6:
        hex_color = f"FF{hex_color}"
    return hex_color


class WorksheetToolkit:
    """
    WorksheetToolkit
    ----------------
    A toolkit for formatting openpyxl worksheets.

    Most methods return self to allow method chaining.
    """

    def __init__(self, worksheet):
        """Initialize a WorksheetToolkit instance for a single worksheet.

        Parameters
        ----------
        worksheet : openpyxl.worksheet.worksheet.Worksheet
            The openpyxl Worksheet object to operate on. All formatting and
            utility methods in this instance will apply to this worksheet.
        """
        self.worksheet = worksheet

    def freeze_panes(self, cell=None):
        """Freeze panes so that rows above and columns to the left of the given cell remain visible.
        Pass None or an empty string to unfreeze panes.

        Parameters
        ----------
        cell : str, optional
            The top-left cell to freeze panes at, e.g. 'B2'. If None, panes are unfrozen.

        Warning
        -------
        Due to the way Excel stores pane information, some versions of Excel may display a repair
        warning when opening a workbook, even if this method was used correctly.
        This is an Excel-specific behavior and not a bug in the toolkit.

        If you encounter such a warning, you can safely:
        - Re-save the workbook in Excel
        - Or create a new worksheet/workbook and reapply formatting using this toolkit
        """
        if cell is None or cell == "":
            cell = "A1"
        self.worksheet.freeze_panes = cell
        return self

    def merge_cells(
        self, *, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None
    ):
        """Merge a rectangular range of cells in the worksheet. You can either provide
        `range_string` or all four start/end coordinates.

        Parameters
        ----------
        range_string : str, optional
            A string representing the range to merge, e.g., 'A1:C3'.
            If provided, start/end row/column are ignored.
        start_row : int, optional
            Row number of the top-left cell of the range.
        start_column : int, optional
            Column number of the top-left cell of the range.
        end_row : int, optional
            Row number of the bottom-right cell of the range.
        end_column : int, optional
            Column number of the bottom-right cell of the range.
        """
        self.worksheet.merge_cells(
            range_string=range_string,
            start_row=start_row,
            end_row=end_row,
            start_column=start_column,
            end_column=end_column,
        )
        return self

    def set_alignment(
        self,
        *,
        rows=None,
        columns=None,
        intersections_only=True,
        horizontal=_UNCHANGED,
        vertical=_UNCHANGED,
        text_rotation=_UNCHANGED,
        wrap_text=_UNCHANGED,
        shrink_to_fit=_UNCHANGED,
        indent=_UNCHANGED,
        reading_order=_UNCHANGED,
    ):
        """Set the alignment of cells defined by rows and columns. If rows and columns are empty or
        None, applies to all cells.

        Parameters
        ----------
        rows : Union[List[int], int], optional
        columns : Union[List[int], int], optional
        intersections_only : bool
            If True, then alignment is applied to only the cells that have a row number in the rows
            argument and a column number in the columns argument.
            If False, then the alignment is applied to any cell in the given rows and any cell in
            the given columns.
            Only considered if both rows and columns are provided.
        horizontal : str, optional
            Horizontal alignment: 'general', 'left', 'center', 'right', 'fill', 'justify',
            'centerContinuous', or 'distributed'.
        vertical : str, optional
            Vertical alignment: 'top', 'center', 'bottom', 'justify', 'distributed'.
        text_rotation : int, optional
            Rotation of the text in degrees (0-180).
        wrap_text : bool, optional
        shrink_to_fit : bool, optional
        indent : int, optional
            Number of spaces to indent text.
        reading_order : int, optional
            Text direction: 0 = Context (default), 1 = Left-to-right, 2 = Right-to-left.

        Examples
        --------
        >>> # Center all text horizontally and vertically
        >>> toolkit.set_alignment(horizontal='center', vertical='center')
        >>>
        >>> # Wrap text in rows 1-2 only
        >>> toolkit.set_alignment(rows=[1, 2], wrap_text=True)
        >>>
        >>> # Rotate text 45 degrees for intersection of row 1-2 and columns 1-2
        >>> toolkit.set_alignment(rows=[1, 2], columns=[1, 2], intersections_only=True,
        >>>                       text_rotation=45)
        """
        parameters = {
            "horizontal": horizontal,
            "vertical": vertical,
            "text_rotation": text_rotation,
            "wrap_text": wrap_text,
            "shrink_to_fit": shrink_to_fit,
            "indent": indent,
            "readingOrder": reading_order,
        }

        updates = []
        for cell in self._iter_cells(rows, columns, intersections_only):
            # Copy and override for the same reason as set_font: rebuilding from
            # the arguments drops justifyLastLine and relativeIndent.
            new_alignment = copy(cell.alignment)
            for attribute, value in parameters.items():
                if value is not _UNCHANGED:
                    setattr(new_alignment, attribute, value)
            updates.append((cell, new_alignment))

        for cell, alignment in updates:
            cell.alignment = alignment
        return self

    def set_border(
        self,
        *,
        rows=None,
        columns=None,
        intersections_only=True,
        sides=("top", "bottom", "left", "right"),
        style=_UNCHANGED,
        color=_UNCHANGED,
    ):
        """Set borders for specified cells.

        Parameters
        ----------
        rows : int or list of int, optional
        columns : int or list of int, optional
        intersections_only : bool, optional
        sides : str or tuple of str, optional
            Sides to modify. Can include:
            'top', 'bottom', 'left', 'right', 'diagonal_up', 'diagonal_down'.
        style : str, optional
            Border style. Can be:
            'hair', 'thin','medium', 'thick', 'double', 'dotted', 'dashed','mediumDashed',
            'mediumDashDot', 'mediumDashDotDot', 'dashDot', 'dashDotDot', 'slantDashDot'.
        color : str, optional
            Hex color code, e.g., '#ff0000'.

        Examples
        --------
        >>> # Thin red top and bottom borders for rows 1-2
        >>> toolkit.set_border(rows=[1,2], sides=('top','bottom'), style='thin', color='#ff0000')
        >>>
        >>> # Draw a diagonal line from bottom-left to top-right in cell B2
        >>> toolkit.set_border(rows=2, columns=2, intersections_only=True, sides=('diagonal_up'),
        >>>                    style='thin', color='#00ff00')
        """
        if isinstance(sides, str):
            sides = (sides,)
        unknown = [side for side in sides if side not in _BORDER_SIDES]
        if unknown:
            raise ValueError(f"unknown border side(s) {unknown}. Choose from {list(_BORDER_SIDES)}")

        # Built first and assigned afterwards, so a cell that cannot be given the
        # requested border does not leave the rest of the range half-drawn.
        updates = []
        for cell in self._iter_cells(rows, columns, intersections_only):
            current = cell.border
            border_kwargs = {}

            # Standard sides
            for side_name in ("left", "right", "top", "bottom"):
                if side_name in sides:
                    new_style = (
                        style if style is not _UNCHANGED else getattr(current, side_name).style
                    )
                    if color is not _UNCHANGED and new_style is None:
                        raise ValueError(
                            f"{cell.coordinate} has no {side_name} border, so a colour on its "
                            f"own would not show. Pass style='thin' as well."
                        )
                    border_kwargs[side_name] = Side(
                        style=new_style,
                        color=_normalize_color(color)
                        if color is not _UNCHANGED
                        else getattr(current, side_name).color,
                    )
                else:
                    border_kwargs[side_name] = getattr(current, side_name)

            # Diagonal side logic
            if "diagonal_up" in sides or "diagonal_down" in sides:
                new_style = style if style is not _UNCHANGED else current.diagonal.style
                if color is not _UNCHANGED and new_style is None:
                    raise ValueError(
                        f"{cell.coordinate} has no diagonal border, so a colour on its own "
                        f"would not show. Pass style='thin' as well."
                    )
                border_kwargs["diagonal"] = Side(
                    style=new_style,
                    color=_normalize_color(color)
                    if color is not _UNCHANGED
                    else current.diagonal.color,
                )
                border_kwargs["diagonalUp"] = "diagonal_up" in sides
                border_kwargs["diagonalDown"] = "diagonal_down" in sides
            else:
                border_kwargs["diagonal"] = current.diagonal
                border_kwargs["diagonalUp"] = current.diagonalUp
                border_kwargs["diagonalDown"] = current.diagonalDown

            updates.append((cell, Border(**border_kwargs)))

        for cell, border in updates:
            cell.border = border

        return self

    def set_outside_border(
        self, *, start_row, end_row, start_column, end_column, style, color=_UNCHANGED
    ):
        """Set a border only on the outside edges of a rectangular block of cells.

        Parameters
        ----------
        start_row : int
            First row of the block.
        end_row : int
            Last row of the block.
        start_column : int
            First column of the block.
        end_column : int
            Last column of the block.
        style : str, optional
            Border style. Can be:
            'hair', 'thin','medium', 'thick', 'double', 'dotted', 'dashed','mediumDashed',
            'mediumDashDot', 'mediumDashDotDot', 'dashDot', 'dashDotDot', 'slantDashDot'.
        color : str, optional
            Hex color code, e.g., '#ff0000'.

        Example
        -------
        >>> # Add a thin black border around rows 1-3 and columns 1-4
        >>> toolkit.set_outside_border(start_row=1, end_row=3, start_column=1, end_column=4,
                                       style='thin', color='#000000')
        """
        rows = list(range(start_row, end_row + 1))
        columns = list(range(start_column, end_column + 1))

        # Top border
        self.set_border(
            rows=[start_row],
            columns=columns,
            intersections_only=True,
            sides=("top",),
            style=style,
            color=color,
        )
        # Bottom border
        self.set_border(
            rows=[end_row],
            columns=columns,
            intersections_only=True,
            sides=("bottom",),
            style=style,
            color=color,
        )
        # Left border
        self.set_border(
            rows=rows,
            columns=[start_column],
            intersections_only=True,
            sides=("left",),
            style=style,
            color=color,
        )
        # Right border
        self.set_border(
            rows=rows,
            columns=[end_column],
            intersections_only=True,
            sides=("right",),
            style=style,
            color=color,
        )

        return self

    def set_column_best_fit(
        self,
        *,
        columns=None,
        padding=0,
        ignore_rows=None,
        ignore_formulas=True,
        min_width=None,
        max_width=None,
        measure=None,
    ):
        """Fit columns to their widest cell.

        Widths come from the real per-character advances of Calibri, so a column of
        narrow letters is not given the same width as one of capitals. Excel's own
        formula is used to turn that into a column width:

            width = (pixels of text + 5 padding pixels) / max digit width

        A column with nothing in it is left alone rather than shrunk.

        Parameters
        ----------
        columns : int or list of int, optional
            Column numbers to fit. Defaults to every column in use.
        padding : float, optional
            Extra width on top of the fitted value. Defaults to 0; Excel's own
            5 pixels of cell padding are already part of the formula.
        ignore_rows : list of int, optional
            Row numbers to leave out when measuring.
        ignore_formulas : bool, optional
            If True, cells holding formulas are not measured. The formula text is
            not what the reader sees, so measuring it oversizes the column.
        min_width : float, optional
            Lower bound on the result.
        max_width : float, optional
            Upper bound. Defaults to Excel's own maximum of 255.
        measure : callable, optional
            ``measure(text, font, normal_font) -> width``, where each font is a
            ``(name, point_size)`` pair: the cell's own font, and the workbook's
            normal font that Excel's width unit is defined in. Supply this for a
            face with no built-in table.

        Notes
        -----
        - Text is assumed to be on one line; wrapped text is not accounted for.
        - Only dates and times are rendered as Excel displays them. Other number
          formats are measured as the value is stored, so a currency or percentage
          column may come out narrower than it needs to be.
        - Built-in metrics cover Aptos, Calibri, Arial, Helvetica, Times New Roman,
          Courier New, Cambria, Verdana, Georgia, Tahoma and Futura. Any other face
          is measured with Calibri's advances unless ``measure`` is given.
        """
        ws = self.worksheet
        if columns is None:
            columns = range(1, ws.max_column + 1)
        if isinstance(columns, Number):
            columns = [columns]
        columns = list(columns)
        # Before any ws.cell() call: materialising an out-of-grid cell makes the
        # workbook permanently unsaveable, so a later failure would come too late.
        _check_bounds(columns=columns)
        ignore_rows = set() if ignore_rows is None else set(ignore_rows)
        normal_font = self._normal_font()

        measure = measure or self._measure_text
        for col in columns:
            excel_width = 0
            measured_anything = False
            for row in range(1, ws.max_row + 1):
                if row in ignore_rows:
                    continue

                cell = ws.cell(row=row, column=col)

                if ignore_formulas and cell.data_type == "f":
                    continue

                if cell.value is None:
                    continue

                measured_anything = True
                # A cell that inherits the workbook font reports neither name nor
                # size of its own, which Font(bold=True) alone is enough to produce.
                font = (
                    cell.font.name or normal_font[0],
                    normal_font[1] if cell.font.sz is None else cell.font.sz,
                )
                excel_width = max(measure(_displayed_text(cell), font, normal_font), excel_width)

            if not measured_anything:
                # Nothing to fit. Leaving the column alone matters because a width
                # set deliberately beforehand would otherwise be cut to the padding.
                continue

            width = excel_width + padding
            if min_width is not None:
                width = max(width, min_width)
            ws.column_dimensions[get_column_letter(col)].width = min(
                width, _MAX_COLUMN_WIDTH if max_width is None else max_width
            )

        return self

    def set_column_width(self, *, columns=None, width):
        """
        Set the width of one or more columns.

        Parameters
        ----------
        columns : int or list of int
            Column number(s) to modify.
        width : float
            Column width in Excel character units, not pixels. Required: there is no
            existing value to leave alone, so omitting it cannot mean anything.
            A width of 0 hides the column, which is how Excel stores a zero.
        """
        if width < 0:
            raise ValueError(f"width must not be negative, got {width}")

        if isinstance(columns, Number):
            columns = [columns]
        if columns is None:
            columns = list(range(1, self.worksheet.max_column + 1))

        columns = list(columns)
        _check_bounds(columns=columns)
        width = min(width, _MAX_COLUMN_WIDTH)
        for col in columns:
            # get_column_letter rather than a cell lookup: row 1 of the column
            # may be a MergedCell, which has no column_letter at all.
            dimension = self.worksheet.column_dimensions[get_column_letter(col)]
            # openpyxl cannot persist a zero width -- the writer drops any falsy
            # dimension -- so the only way to honour it is to hide the column.
            if width == 0:
                dimension.hidden = True
            else:
                dimension.width = width
        return self

    def set_row_height(self, *, rows=None, height):
        """
        Set the height of one or more rows.

        Parameters
        ----------
        rows : int or list of int
            Row number(s) to modify.
        height : float
            Row height in points. Required, for the same reason as the column width.
            A height of 0 hides the row.
        """
        if height < 0:
            raise ValueError(f"height must not be negative, got {height}")

        if isinstance(rows, Number):
            rows = [rows]
        if rows is None:
            rows = list(range(1, self.worksheet.max_row + 1))

        rows = list(rows)
        _check_bounds(rows=rows)
        height = min(height, _MAX_ROW_HEIGHT)
        for row in rows:
            dimension = self.worksheet.row_dimensions[row]
            if height == 0:
                dimension.hidden = True
            else:
                dimension.height = height
        return self

    def set_fill(
        self,
        *,
        rows=None,
        columns=None,
        intersections_only=True,
        fill_type=_UNCHANGED,
        start_color=_UNCHANGED,
        end_color=_UNCHANGED,
    ):
        """Set the fill of cells defined by rows and columns. If rows and columns are empty or None,
        applies to all cells.

        Parameters
        ----------
        rows : Union[List[int], int], optional
        columns : Union[List[int], int], optional
        intersections_only : bool, optional
            If True, then fill is applied to only the cells that have a row number in the rows
            argument and a column number in the columns argument.
            If False, then the fill is applied to any cell in the given rows and any cell in the
            given columns.
            Only considered if both rows and columns are provided.
        fill_type : str, optional
            Type of fill/pattern. Common values: 'solid', 'gray125', 'darkGrid', etc. Use None for
            transparent fill.
        start_color : str, optional
            Hex color code for the primary fill color (foreground). For solid fills,
            this is the visible background color.
        end_color : str, optional
            Hex color code for the secondary fill color (background). This is only
            relevant for patterned fills (e.g., 'trellis', 'cross', stripes).
            For solid fills, this value is ignored by Excel and usually does not
            need to be provided.

        Examples
        --------
        >>> # Fill row 1 with pink
        >>> toolkit.set_fill(rows=1, start_color='#f4d2d3')
        >>>
        >>> # Fill intersection of row 1-2 and col 1-2 with yellow
        >>> toolkit.set_fill(rows=[1,2], columns=[1,2], intersections_only=True, start_color='#ffff00')
        """
        if start_color is None or end_color is None:
            # A pattern fill has no colourless state: its default foreground is an
            # opaque black, so None here would repaint rather than clear.
            raise ValueError(
                "set_fill cannot clear a colour on its own -- a pattern fill always carries "
                "one. Pass fill_type=None to remove the fill, or give an explicit hex colour."
            )

        requested = any(arg is not _UNCHANGED for arg in (fill_type, start_color, end_color))

        # Every fill is built before any is assigned, so a failure part-way through
        # leaves the worksheet exactly as it was found rather than half-formatted.
        updates = []
        for cell in self._iter_cells(rows, columns, intersections_only):
            current = cell.fill
            # cell.fill is a StyleProxy, so isinstance against PatternFill never
            # matches; the wrapped object's tagname is what distinguishes them.
            if getattr(current, "tagname", None) == "patternFill":
                current_type = current.fill_type
                # copy(), not the object itself: a Color is stored by reference, so
                # passing the existing one through would leave this cell sharing a
                # mutable Color with whatever it was read from -- including the
                # process-global default that PatternFill uses for an absent bgColor.
                current_start = copy(current.start_color)
                current_end = copy(current.end_color)
            elif not requested:
                # Nothing was asked for, so leave the gradient alone rather than
                # flattening it into a blank pattern fill.
                continue
            else:
                # A GradientFill has no pattern or start/end colour to merge with,
                # so anything left unspecified falls back to PatternFill's default.
                current_type = current_start = current_end = None

            new_type = current_type if fill_type is _UNCHANGED else fill_type
            new_start = (
                current_start if start_color is _UNCHANGED else _normalize_color(start_color)
            )

            if start_color is not _UNCHANGED and new_type is None:
                raise ValueError(
                    f"{cell.coordinate} has no fill pattern, so a colour on its own would "
                    f"not show. Pass fill_type='solid' as well."
                )
            if fill_type not in (_UNCHANGED, None) and not _has_color(new_start):
                raise ValueError(
                    f"{cell.coordinate} has no fill colour, so fill_type={fill_type!r} alone "
                    f"would paint it black. Pass start_color as well."
                )

            updates.append(
                (
                    cell,
                    PatternFill(
                        fill_type=new_type,
                        # The Color object is passed through rather than its .rgb:
                        # for a theme, indexed or automatic colour that attribute is
                        # the descriptor itself, which PatternFill rejects.
                        start_color=new_start,
                        end_color=current_end
                        if end_color is _UNCHANGED
                        else _normalize_color(end_color),
                    ),
                )
            )

        for cell, fill in updates:
            cell.fill = fill
        return self

    def set_font(
        self,
        *,
        rows=None,
        columns=None,
        intersections_only=True,
        name=_UNCHANGED,
        size=_UNCHANGED,
        bold=_UNCHANGED,
        italic=_UNCHANGED,
        underline=_UNCHANGED,
        strike=_UNCHANGED,
        color=_UNCHANGED,
    ):
        """Set the font of cells defined by rows and columns. If rows and columns are empty or None,
        applies to all cells.

        Parameters
        ----------
        rows : Union[List[int], int], optional
        columns : Union[List[int], int], optional
        intersections_only : bool
            If True, then font is applied to only the cells that have a row number in the rows
            argument and a column number in the columns argument.
            If False, then the font is applied to any cell in the given rows and any cell in the
            given columns.
            Only considered if both rows and columns are provided.
        name : str, optional
            Name of font face
        size : int, optional
        bold : bool, optional
        italic : bool, optional
        underline : str, optional
            'single', 'double', 'singleAccounting', 'doubleAccounting', or None
        strike : bool, optional
        color : str, optional
            Hex color code, e.g., '#f4d2d3'

        Example
        -------
        >>> # Make rows 1 and 2 bold
        >>> set_font(rows=[1, 2], bold=True)
        >>>
        >>> # Set the font size for the following cells only: row1/col1, row2/col1, row1/col2, row2/col2
        >>> set_font(rows=[1, 2], columns=[1, 2], intersections_only=True, size=16)
        >>>
        >>> # Set the font color for all of row 1 and 2, and all of columns 1 and 2
        >>> set_font(rows=[1, 2], columns=[1, 2], color='#000000')
        """
        parameters = {
            "name": name,
            "size": size,
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "strike": strike,
            "color": _normalize_color(color),
        }

        updates = []
        for cell in self._iter_cells(rows, columns, intersections_only):
            # Copy and override, rather than build a new Font from the arguments:
            # a fresh Font would silently reset every attribute this method does
            # not expose, such as vertAlign and scheme.
            new_font = copy(cell.font)
            for attribute, value in parameters.items():
                if value is not _UNCHANGED:
                    setattr(new_font, attribute, value)
            updates.append((cell, new_font))

        for cell, font in updates:
            cell.font = font
        return self

    def set_zoom_scale(self, zoom_scale=100):
        """Set the zoom scale.

        Parameters
        ----------
        zoom_scale : int, optional
            Zoom percentage (10-400). Defaults to 100.
        """
        if not 10 <= zoom_scale <= 400:
            raise ValueError("zoom_scale must be between 10 and 400")
        self.worksheet.sheet_view.zoomScale = zoom_scale
        return self

    @staticmethod
    def _measure_text(text, font, normal_font):
        """Column width that fits ``text``, from the real advances of its font."""
        name, size = font
        base_name, base_size = normal_font
        return _metrics.column_width(text, size, base_size, name, base_name)

    def _normal_font(self):
        """The workbook's normal font, as ``(name, point_size)``.

        A cell with no font record of its own inherits this, and Excel's column
        width unit is defined by it rather than by whatever the cell uses.
        """
        fonts = getattr(self.worksheet.parent, "_fonts", None)
        normal = fonts[0] if fonts else None
        return (
            getattr(normal, "name", None) or DEFAULT_FONT.name,
            getattr(normal, "sz", None) or DEFAULT_FONT.sz,
        )

    def _iter_cells(self, rows=None, columns=None, intersections_only=True):
        ws = self.worksheet

        # Normalize rows and columns
        if rows is None:
            rows = []
        if columns is None:
            columns = []
        if isinstance(rows, Number):
            rows = [rows]
        if isinstance(columns, Number):
            columns = [columns]
        if not rows:
            intersections_only = True
            rows = list(range(1, ws.max_row + 1))
        if not columns:
            intersections_only = True
            columns = list(range(1, ws.max_column + 1))
        rows = sorted(set(rows))
        columns = sorted(set(columns))
        _check_bounds(rows, columns)

        # --- Intersection case ---
        if intersections_only:
            for r in rows:
                for c in columns:
                    yield ws.cell(row=r, column=c)
            return

        # --- Union case ---
        seen = set()

        # Row sweep (top to bottom, left to right)
        for r in rows:
            for c in range(1, ws.max_column + 1):
                key = (r, c)
                if key not in seen:
                    seen.add(key)
                    yield ws.cell(row=r, column=c)

        # Column sweep (left to right, top to bottom)
        for c in columns:
            for r in range(1, ws.max_row + 1):
                key = (r, c)
                if key not in seen:
                    seen.add(key)
                    yield ws.cell(row=r, column=c)
