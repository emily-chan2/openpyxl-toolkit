"""Chainable formatting helpers for openpyxl worksheets.

The public surface is :class:`WorksheetToolkit`, which wraps a single
``openpyxl`` worksheet and merges style changes into whatever formatting the
cell already carries, rather than replacing it wholesale. :func:`column_letter`
and :func:`column_index` convert between the two ways of naming a column.
"""

from .columns import column_index, column_letter
from .worksheet_toolkit import WorksheetToolkit

__all__ = ["WorksheetToolkit", "column_index", "column_letter"]
__version__ = "0.4.0"
