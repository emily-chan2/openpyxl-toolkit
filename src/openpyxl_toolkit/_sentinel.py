"""The marker for an argument the caller left out.

None is already a meaningful value for several parameters -- a fill type of None
removes the fill, an underline of None removes the underline -- so it cannot also
stand for "leave this as it is".

The marker is a one-member enum rather than a plain object, because that is the
only spelling a type checker narrows. Against a plain class, ``if value is not
UNCHANGED`` leaves the sentinel in the type and every use of the value is then an
error; against ``Literal[_UnchangedType.TOKEN]`` the check removes it.
"""

from __future__ import annotations

import enum
from typing import Literal, TypeAlias


class _UnchangedType(enum.Enum):
    TOKEN = enum.auto()

    def __repr__(self) -> str:
        # help() and editor tooltips print the default value of every styling
        # parameter, and "<Unchanged.TOKEN: 1>" reads as an implementation leak.
        return "<unchanged>"


UNCHANGED = _UnchangedType.TOKEN

#: The type to write in a signature: ``str | None | Unchanged``.
Unchanged: TypeAlias = Literal[_UnchangedType.TOKEN]
