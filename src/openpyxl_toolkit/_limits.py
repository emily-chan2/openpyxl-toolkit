"""The sizes Excel will not exceed.

openpyxl enforces none of these. It will create a cell outside the grid or accept
a width of 1000 without complaint, and the resulting workbook then fails to open.
"""

from typing import Final

# get_column_letter is not a substitute for MAX_COLUMN: it accepts up to ZZZ (18278).
MAX_ROW: Final = 1_048_576
MAX_COLUMN: Final = 16_384

# Excel refuses a column wider than this and clips a row taller than it.
MAX_COLUMN_WIDTH: Final = 255
MAX_ROW_HEIGHT: Final = 409
