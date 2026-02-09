from numbers import Number
from openpyxl.styles import Color, Font

_UNCHANGED = object()

class WorksheetToolkit:
    def __init__(self, worksheet):
        self.worksheet = worksheet

    def set_font(self, *, rows=None, columns=None, intersections_only=False, name=_UNCHANGED,
                 size=_UNCHANGED, bold=_UNCHANGED, italic=_UNCHANGED, underline=_UNCHANGED,
                 strike=_UNCHANGED, color=_UNCHANGED):
        """Sets the font of the cells defined by the lists of rows and columns. If rows and columns
        are both empty lists, then applies the font to all cells.

        Parameters
        ----------
        - rows : Optional[List[int], int]
        - columns : Optional[List[int], int]
        - intersections_only : bool
            - If True, then font is applied to only the cells that have a row number in the rows
              argument and a column number in the columns argument. If False, then the font is
              applied to any cell in the given rows and any cell in the given columns.
            - Only considered if both rows and columns are provided
        - name : Optional[str]
        - size : Optional[int]
        - bold : Optional[bool]
        - italic : Optional[bool]
        - underline : Optional[str]
            - 'single', 'double', 'singleAccounting', 'doubleAccounting', or None
        - strike : Optional[bool]
        - color : Optional[str]
            - Hex color code

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
            cell.font = self._merge_font(cell, name=name, size=size, bold=bold, italic=italic, underline=underline, strike=strike, color=color)
        return self

    def _merge_font(self, cell, name=_UNCHANGED, size=_UNCHANGED, bold=_UNCHANGED, italic=_UNCHANGED,
                    underline=_UNCHANGED, strike=_UNCHANGED, color=_UNCHANGED):
        current_font = cell.font

        font_name = current_font.name if name is _UNCHANGED else name
        font_size = current_font.size if size is _UNCHANGED else size
        font_bold = current_font.bold if bold is _UNCHANGED else bold
        font_italic = current_font.italic if italic is _UNCHANGED else italic
        font_underline = current_font.underline if underline is _UNCHANGED else underline
        font_strike = current_font.strike if strike is _UNCHANGED else strike
        if color is _UNCHANGED:
            font_color = current_font.color
        else:
            font_color = Color(rgb=color.lstrip("#"))
        new_font = Font(
            name=font_name,
            size=font_size,
            bold=font_bold,
            italic=font_italic,
            underline=font_underline,
            strike=font_strike,
            color=font_color,
        )
        return new_font

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
            rows = list(range(1, ws.max_row + 1))
        if not columns:
            columns = list(range(1, ws.max_column + 1))
        rows = sorted(set(rows))
        columns = sorted(set(columns))

        # Intersection only
        if intersections_only:
            for r in rows:
                for c in columns:
                    yield ws.cell(row=r, column=c)
            return

        # Union case
        seen = set()

        # Row sweep (top → bottom, left → right)
        for r in rows:
            for c in range(1, ws.max_column + 1):
                key = (r, c)
                if key not in seen:
                    seen.add(key)
                    yield ws.cell(row=r, column=c)

        # Column sweep (left → right, top → bottom)
        for c in columns:
            for r in range(1, ws.max_row + 1):
                key = (r, c)
                if key not in seen:
                    seen.add(key)
                    yield ws.cell(row=r, column=c)
