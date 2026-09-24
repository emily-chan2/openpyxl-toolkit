"""Converting between a column's letter and its number.

Neither function touches a worksheet, so neither has an opinion about how wide
Excel's grid is. They convert, and the methods that create cells are where an
index too far out is caught. openpyxl's pair in ``openpyxl.utils`` stops at
``ZZZ``, three letters, which is neither Excel's limit nor no limit at all.
"""

from __future__ import annotations

__all__ = ["column_index", "column_letter"]

_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def column_letter(index: int) -> str:
    """Convert a column number to its letter.

    Parameters
    ----------
    index : int
        Column number, counting from 1. There is no upper limit: the number does
        not have to be a column Excel has, since converting one is not the same
        as writing to it. Formatting a column past the end of the grid is caught
        by the method that formats it.

    Returns
    -------
    str
        The column letter, in upper case.

    Raises
    ------
    TypeError
        If ``index`` is not an integer. A float is rejected rather than rounded.
    ValueError
        If ``index`` is below 1. Columns are counted from 1, as Excel counts
        them, so 0 and negatives name nothing.

    Examples
    --------
    >>> column_letter(1)
    'A'
    >>> column_letter(27)
    'AA'
    >>> column_letter(16384)
    'XFD'
    """
    if not isinstance(index, int) or isinstance(index, bool):
        raise TypeError(f"column index must be an integer, got {index!r}")
    if index < 1:
        raise ValueError(f"column index must be 1 or more, got {index}")

    letters = []
    while index:
        index, remainder = divmod(index - 1, 26)
        letters.append(_LETTERS[remainder])
    return "".join(reversed(letters))


def column_index(letter: str) -> int:
    """Convert a column letter to its number.

    Parameters
    ----------
    letter : str
        Column letter. Case does not matter, and surrounding whitespace is
        ignored, since neither can mean anything else. There is no upper limit,
        for the same reason there is none on :func:`column_letter`.

    Returns
    -------
    int
        The column number, counting from 1.

    Raises
    ------
    TypeError
        If ``letter`` is not a string.
    ValueError
        If ``letter`` is not made of letters. A cell reference such as ``'C1'``
        is rejected rather than read as its column, since reading it that way
        would accept a range where a column was meant.

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

    index = 0
    for character in name:
        index = index * 26 + (_LETTERS.index(character) + 1)
    return index
