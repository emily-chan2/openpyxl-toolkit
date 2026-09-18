"""The examples in the docstrings have to keep working.

They are written the way a reader would type them, without the value each call
hands back, so doctest cannot check them: it would compare that value against an
empty expected output. Running them is what matters. All of them were broken
before stage 6 -- one called a bare function, one passed a colour with no
pattern, one spelled a tuple `('diagonal_up')`, which is a string.

The README is checked by doctest instead, since the output there is the point.
"""

import doctest
import inspect

import pytest
from openpyxl import Workbook

from openpyxl_toolkit import WorksheetToolkit

DOCUMENTED = sorted(
    name
    for name, member in vars(WorksheetToolkit).items()
    if inspect.isfunction(member) and ">>>" in (member.__doc__ or "")
)
# The class carries examples of its own, and they are the first ones anyone sees.
if ">>>" in (WorksheetToolkit.__doc__ or ""):
    DOCUMENTED.append("WorksheetToolkit")


def _populated_sheet():
    workbook = Workbook()
    sheet = workbook.active
    for row in range(1, 4):
        for column in range(1, 5):
            sheet.cell(row=row, column=column, value=f"r{row}c{column}")
    return sheet


@pytest.mark.parametrize("name", DOCUMENTED)
def test_the_examples_in_a_docstring_run(name):
    sheet = _populated_sheet()
    namespace = {
        "workbook": sheet.parent,
        "sheet": sheet,
        "toolkit": WorksheetToolkit(sheet),
    }
    owner = WorksheetToolkit if name == "WorksheetToolkit" else getattr(WorksheetToolkit, name)
    examples = doctest.DocTestParser().get_examples(owner.__doc__)
    source = "".join(example.source for example in examples)

    exec(compile(source, f"<{name} docstring>", "exec"), namespace)


def test_the_documented_methods_are_the_ones_expected():
    """An empty list would make the test above pass by testing nothing."""
    assert DOCUMENTED == [
        "set_alignment",
        "set_border",
        "set_fill",
        "set_font",
        "set_number_format",
        "set_outside_border",
        "WorksheetToolkit",
    ]
