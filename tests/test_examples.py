"""The examples in the docstrings and the README have to keep working.

They are written the way a reader would type them, without the value each call
hands back, so doctest cannot check them: it would compare that value against an
empty expected output. Running them is what matters. All of the docstring ones
were broken before stage 6 -- one called a bare function, one passed a color
with no pattern, one spelled a tuple `('diagonal_up')`, which is a string.

What this catches is the way examples rot: a renamed method, a changed
signature, an argument that no longer exists, a call that now raises. What it
cannot catch is a comment beside an example going out of date, since a comment
is never run. Claims worth pinning belong in a test of their own.

The column functions are the exception. They hand back a value rather than the
toolkit, so their examples are written with the answer under them and doctest
can check it, which is stricter than merely running them.
"""

import doctest
import inspect
from pathlib import Path

import pytest
from openpyxl import Workbook

import openpyxl_toolkit.columns
from openpyxl_toolkit import WorksheetToolkit

README = Path(__file__).resolve().parent.parent / "README.md"

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


def test_the_examples_in_the_readme_run():
    """One script, not one per block: the README builds a sheet and then uses it."""
    examples = doctest.DocTestParser().get_examples(README.read_text(encoding="utf-8"))
    source = "".join(example.source for example in examples)

    exec(compile(source, "<README.md>", "exec"), {})


def test_the_readme_still_carries_examples():
    """Nothing to run would make the test above pass without reading anything."""
    examples = doctest.DocTestParser().get_examples(README.read_text(encoding="utf-8"))
    assert len(examples) >= 10


def test_the_examples_in_the_column_functions_give_the_answers_they_claim():
    """Checked against the printed output, not just run."""
    results = doctest.testmod(openpyxl_toolkit.columns)

    assert results.failed == 0
    assert results.attempted >= 6


def test_the_documented_methods_are_the_ones_expected():
    """An empty list would make the test above pass by testing nothing."""
    assert DOCUMENTED == [
        "set_alignment",
        "set_autofilter",
        "set_border",
        "set_fill",
        "set_font",
        "set_gridline_visibility",
        "set_number_format",
        "set_outside_border",
        "set_sheet_visibility",
        "set_tab_color",
        "WorksheetToolkit",
    ]
