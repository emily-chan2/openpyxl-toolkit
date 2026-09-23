"""Converting between a column's letter and its number.

openpyxl has a pair of these in ``openpyxl.utils``, but both run off the end of
the grid: ``get_column_letter(16385)`` hands back ``'XFE'`` and
``column_index_from_string('XFE')`` hands back 16385, neither of which Excel can
open. These stop at XFD, and report the same way as every other index the
package checks.
"""

from __future__ import annotations

from openpyxl.utils import column_index_from_string, get_column_letter

from ._limits import MAX_COLUMN
from ._ranges import check_indexes

__all__ = ["column_index", "column_letter"]

#: The letter of the last column Excel has, worked out rather than spelled out.
LAST_COLUMN = get_column_letter(MAX_COLUMN)


def column_letter(index: int) -> str:
    """Convert a column number to the letter Excel shows in its header.

    Parameters
    ----------
    index : int
        Column number, counting from 1.

    Returns
    -------
    str
        The column letter, in upper case.

    Raises
    ------
    TypeError
        If ``index`` is not an integer. A float is rejected rather than rounded:
        openpyxl raises ``list indices must be integers`` from inside itself.
    ValueError
        If ``index`` falls outside Excel's grid.

    Examples
    --------
    >>> column_letter(1)
    'A'
    >>> column_letter(27)
    'AA'
    >>> column_letter(16384)
    'XFD'
    """
    check_indexes([index], MAX_COLUMN, "column")
    return get_column_letter(index)


def column_index(letter: str) -> int:
    """Convert a column letter to the number openpyxl counts it by.

    Parameters
    ----------
    letter : str
        Column letter. Case does not matter, and surrounding whitespace is
        ignored, since neither can mean anything else.

    Returns
    -------
    int
        The column number, counting from 1.

    Raises
    ------
    TypeError
        If ``letter`` is not a string.
    ValueError
        If ``letter`` is not made of letters, or names a column outside Excel's
        grid. A cell reference such as ``'C1'`` is rejected rather than read as
        its column.

    Examples
    --------
    >>> column_index("A")
    1
    >>> column_index("aa")
    27
    >>> column_index("XFD")
    16384
    """
    if not isinstance(letter, str):
        raise TypeError(f"column letter must be a string, got {letter!r}")

    name = letter.strip().upper()
    if not name.isascii() or not name.isalpha():
        # A cell reference is the likely mistake, but only say so when one was passed.
        hint = (
            " A column letter carries no row number: write 'C' rather than 'C1'."
            if any(character.isdigit() for character in name)
            else ""
        )
        raise ValueError(f"{letter!r} is not a column letter.{hint}")
    # Compared as text, which works because both are upper case ASCII: a longer
    # name is always a later column, and equal lengths sort the same way.
    if (len(name), name) > (len(LAST_COLUMN), LAST_COLUMN):
        raise ValueError(
            f"column {name} is outside the worksheet: columns run from A to {LAST_COLUMN}"
        )
    return column_index_from_string(name)
