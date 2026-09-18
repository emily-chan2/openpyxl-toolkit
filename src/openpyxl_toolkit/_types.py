"""The names for what the public methods accept.

The string sets come from openpyxl's own validators, read off ``Alignment``,
``Font``, ``Side`` and ``PatternFill``. Spelling them as ``Literal`` is what puts
the choices in an editor's completion list and catches a typo such as "centre"
before the workbook is written. openpyxl still validates at runtime, so a value
it gains and this file has not caught up with raises rather than corrupts.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
from typing import TYPE_CHECKING, Literal, TypeAlias

if TYPE_CHECKING:
    from openpyxl.cell.cell import Cell, MergedCell

#: What ``worksheet.cell()`` hands back. Every cell of a merged range except
#: the top-left one is a MergedCell, which carries style but not a value.
AnyCell: TypeAlias = "Cell | MergedCell"

#: Row or column numbers, or None for every one in use.
IndexSelection: TypeAlias = "Iterable[int] | None"

#: A font as ``(name, point size)``.
FontSpec: TypeAlias = "tuple[str, float]"

#: ``measure(text, font, normal_font) -> column width``.
MeasureText: TypeAlias = Callable[[str, FontSpec, FontSpec], float]

HorizontalAlignment: TypeAlias = Literal[
    "general",
    "left",
    "center",
    "right",
    "fill",
    "justify",
    "centerContinuous",
    "distributed",
]

VerticalAlignment: TypeAlias = Literal[
    "top",
    "center",
    "bottom",
    "justify",
    "distributed",
]

#: 0 leaves the direction to the text, 1 is left-to-right, 2 is right-to-left.
ReadingOrder: TypeAlias = Literal[0, 1, 2]

Underline: TypeAlias = Literal[
    "single",
    "double",
    "singleAccounting",
    "doubleAccounting",
]

BorderStyle: TypeAlias = Literal[
    "hair",
    "thin",
    "medium",
    "thick",
    "double",
    "dotted",
    "dashed",
    "mediumDashed",
    "mediumDashDot",
    "mediumDashDotDot",
    "dashDot",
    "dashDotDot",
    "slantDashDot",
]

BorderSide: TypeAlias = Literal[
    "top",
    "bottom",
    "left",
    "right",
    "diagonal_up",
    "diagonal_down",
]

FillType: TypeAlias = Literal[
    "solid",
    "darkDown",
    "darkGray",
    "darkGrid",
    "darkHorizontal",
    "darkTrellis",
    "darkUp",
    "darkVertical",
    "gray0625",
    "gray125",
    "lightDown",
    "lightGray",
    "lightGrid",
    "lightHorizontal",
    "lightTrellis",
    "lightUp",
    "lightVertical",
    "mediumGray",
]
