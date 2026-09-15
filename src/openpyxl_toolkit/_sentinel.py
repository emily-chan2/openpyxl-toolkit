"""The marker for an argument the caller left out.

None is already a meaningful value for several parameters -- a fill type of None
removes the fill, an underline of None removes the underline -- so it cannot also
stand for "leave this as it is".
"""


class Unchanged:
    """Marker for a parameter the caller did not pass."""

    __slots__ = ()

    def __repr__(self):
        return "<unchanged>"


UNCHANGED = Unchanged()
