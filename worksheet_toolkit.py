
class WorksheetToolkit:
    def __init__(self, worksheet):
        self.worksheet = worksheet

    def set_font(self, *, rows=[], columns=[], intersections_only=False, name=None, size=None,
                 bold=None, italic=None, underline=None, strike=None, color=None):
        """Sets the font of the cells defined by the lists of rows and columns. If rows and columns
        are both empty lists, then applies the font to all cells.

        Parameters
        ----------
        - ws : Worksheet
        - font : Font
        - rows : Union[List[int], int]
        - columns : Union[List[int], int]
        - intersections_only : bool
            - If True, then font is applied to only the cells that have a row number in the rows
              argument and a column number in the columns argument. If False, then the font is
              applied to any cell in the given rows and any cell in the given columns.
            - Only considered if both rows and columns are provided
        - name : Optional[str]
        - size : Optional[int]
        - bold : Optional[bool]
        - italic : Optional[bool]
        - underline : Optional[bool, str]
            - 'none', 'single', 'double', 'singleAccounting', or 'doubleAccounting'
            - Can also use True or False. True means 'single', False means 'none'
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
        ws = self.worksheet

        if isinstance(rows, Number):
            rows = [rows]
        if isinstance(columns, Number):
            columns = [columns]
        if len(rows) == 0 and len(columns) == 0:
            rows = range(1, ws.max_row+1)
            columns = range(1, ws.max_column+1)

        if (not intersections_only) or len(rows) == 0 or len(columns) == 0:
            for row in rows:
                for col in range(1, ws.max_column+1):
                    cell = ws.cell(row=row, column=col)
                    cell.font = self._merge_font(cell, name=name, size=size, bold=bold, italic=italic, underline=underline, strike=strike, color=color)
            for col in columns:
                for row in range(1, ws.max_row+1):
                    cell = ws.cell(row=row, column=col)
                    cell.font = self._merge_font(cell, name=name, size=size, bold=bold, italic=italic, underline=underline, strike=strike, color=color)
        else:
            for row in rows:
                for col in columns:
                    cell = ws.cell(row=row, column=col)
                    cell.font = self._merge_font(cell, name=name, size=size, bold=bold, italic=italic, underline=underline, strike=strike, color=color)

    def _merge_font(self, cell, name=None, size=None, bold=None, italic=None, underline=None,
                    strike=None, color=None):
        current_font = cell.font

        font_name = current_font.name if name is None else name
        font_size = current_font.size if size is None else size
        font_bold = current_font.bold if bold is None else bold
        font_italic = current_font.italic if italic is None else italic
        font_underline = current_font.underline if underline is None else underline
        if font_underline is True:
            font_underline = 'single'
        elif font_underline is False:
            font_underline = 'none'
        font_strike = current_font.strike if strike is None else strike
        font_color = current_font.color if color is None else color
        if font_color.startswith('#'):
            font_color = font_color[1:]

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
