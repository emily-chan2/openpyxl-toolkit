from numbers import Number
from openpyxl.styles import Alignment, Color, Font, PatternFill

_UNCHANGED = object()

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
            The top-left cell to freeze panes at, e.g. "B2". If None, panes are unfrozen.

        Warning
        -------
        Due to the way Excel stores pane information, some versions of Excel may display a repair
        warning when opening a workbook, even if this method was used correctly.
        This is an Excel-specific behavior and not a bug in the toolkit.

        If you encounter such a warning, you can safely:
        - Re-save the workbook in Excel
        - Or create a new worksheet/workbook and reapply formatting using this toolkit
        """
        if cell is None or cell == '':
            cell = 'A1'
        self.worksheet.freeze_panes = cell
        return self

    def set_alignment(self, *, rows=None, columns=None, intersections_only=False,
                      horizontal=_UNCHANGED, vertical=_UNCHANGED, text_rotation=_UNCHANGED,
                      wrap_text=_UNCHANGED, shrink_to_fit=_UNCHANGED, indent=_UNCHANGED,
                      reading_order=_UNCHANGED):
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
        for cell in self._iter_cells(rows, columns, intersections_only):
            current = cell.alignment
            cell.alignment = Alignment(
                horizontal=current.horizontal if horizontal is _UNCHANGED else horizontal,
                vertical=current.vertical if vertical is _UNCHANGED else vertical,
                text_rotation=current.textRotation if text_rotation is _UNCHANGED else text_rotation,
                wrap_text=current.wrapText if wrap_text is _UNCHANGED else wrap_text,
                shrink_to_fit=current.shrinkToFit if shrink_to_fit is _UNCHANGED else shrink_to_fit,
                indent=current.indent if indent is _UNCHANGED else indent,
                readingOrder=current.readingOrder if reading_order is _UNCHANGED else reading_order,
            )
        return self

    def set_column_best_fit(self, *, columns=None, padding=2, ignore_rows=None, ignore_formulas=True):
        """Auto-fit columns to their contents for Calibri font.

        Uses the approximation:
            width = 0.09903846 * font_size + 0.00186808
        multiplied by the length of the longest cell in each column.

        Parameters
        ----------
        columns : int or list of int, optional
            Column numbers to auto-fit. Defaults to all columns.
        padding : int, optional
            Extra width added to prevent clipping. Defaults to 2.
        ignore_rows : list of int, optional
            Row numbers to ignore when calculating maximum cell length.
        ignore_formulas : bool, optional
            If True, cells containing formulas are ignored in width calculation.

        Notes
        -----
        - This method assumes text is in a single line.
        - Works best for Calibri; other fonts may appear slightly off.
        """
        ws = self.worksheet
        if columns is None:
            columns = range(1, ws.max_column + 1)
        if isinstance(columns, int):
            columns = [columns]
        if ignore_rows is None:
            ignore_rows = set()
        else:
            ignore_rows = set(ignore_rows)

        for col in columns:
            max_len = 0
            font_size = 11  # default Calibri size

            for row in range(1, ws.max_row + 1):
                if row in ignore_rows:
                    continue

                cell = ws.cell(row=row, column=col)

                if ignore_formulas and cell.data_type == 'f':
                    continue

                value = str(cell.value) if cell.value is not None else ''
                max_len = max(max_len, len(value))

                # Update font size if cell has a specific font
                if cell.font and cell.font.sz:
                    font_size = max(font_size, cell.font.sz)

            # Approximate Excel width using Calibri formula
            excel_width = (0.09903846 * font_size + 0.00186808) * max_len
            ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = excel_width + padding

        return self

    def set_fill(self, *, rows=None, columns=None, intersections_only=False, fill_type=_UNCHANGED,
                 start_color=_UNCHANGED, end_color=_UNCHANGED):
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
        >>> toolkit.set_fill(rows=1, start_color="#f4d2d3")
        >>>
        >>> # Fill intersection of row 1-2 and col 1-2 with yellow
        >>> toolkit.set_fill(rows=[1,2], columns=[1,2], intersections_only=True, start_color="#ffff00")
        """
        for cell in self._iter_cells(rows, columns, intersections_only):
            current = cell.fill
            cell.fill = PatternFill(
                fill_type=current.fill_type if fill_type is _UNCHANGED else fill_type,
                start_color=current.start_color.rgb if start_color is _UNCHANGED else self._normalize_color(start_color),
                end_color=current.end_color.rgb if end_color is _UNCHANGED else self._normalize_color(end_color),
            )
        return self

    def set_font(self, *, rows=None, columns=None, intersections_only=False, name=_UNCHANGED,
                 size=_UNCHANGED, bold=_UNCHANGED, italic=_UNCHANGED, underline=_UNCHANGED,
                 strike=_UNCHANGED, color=_UNCHANGED):
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
        for cell in self._iter_cells(rows, columns, intersections_only):
            current_font = cell.font
            cell.font = Font(name=current_font.name if name is _UNCHANGED else name,
                            size=current_font.size if size is _UNCHANGED else size,
                            bold=current_font.bold if bold is _UNCHANGED else bold,
                            italic=current_font.italic if italic is _UNCHANGED else italic,
                            underline=current_font.underline if underline is _UNCHANGED else underline,
                            strike=current_font.strike if strike is _UNCHANGED else strike,
                            color=current_font.color if color is _UNCHANGED else Color(rgb=self._normalize_color(color)))
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

    def _iter_cells(self, rows=None, columns=None, intersections_only=False):
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

    def _normalize_color(self, color):
        if color is _UNCHANGED:
            return _UNCHANGED
        hex_color = color.lstrip('#')
        if len(hex_color) == 6:
            hex_color = f'FF{hex_color}'
        return hex_color