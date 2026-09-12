"""Chainable formatting helpers for openpyxl worksheets.

The public surface is :class:`WorksheetToolkit`, which wraps a single
``openpyxl`` worksheet and merges style changes into whatever formatting the
cell already carries, rather than replacing it wholesale.
"""

from .worksheet_toolkit import WorksheetToolkit

__all__ = ["WorksheetToolkit"]
__version__ = "0.1.0"
