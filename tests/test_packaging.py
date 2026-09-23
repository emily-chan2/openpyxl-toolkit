"""Checks on the installed package metadata."""

from importlib import resources
from importlib.metadata import version
from typing import get_args

import pytest
from openpyxl.styles import Alignment, Font, PatternFill, Side

import openpyxl_toolkit
from openpyxl_toolkit import _types


def test_version_is_single_sourced():
    assert openpyxl_toolkit.__version__ == version("openpyxl-toolkit")


def test_public_surface():
    assert openpyxl_toolkit.__all__ == ["WorksheetToolkit", "column_index", "column_letter"]
    assert openpyxl_toolkit.WorksheetToolkit.__module__ == "openpyxl_toolkit.worksheet_toolkit"
    assert openpyxl_toolkit.column_letter.__module__ == "openpyxl_toolkit.columns"
    assert openpyxl_toolkit.column_index.__module__ == "openpyxl_toolkit.columns"


def test_py_typed_ships_with_the_package():
    """Without the marker, a type checker ignores the annotations once installed."""
    marker = resources.files("openpyxl_toolkit") / "py.typed"
    assert marker.is_file()


@pytest.mark.parametrize(
    ("alias", "owner", "attribute"),
    [
        ("HorizontalAlignment", Alignment, "horizontal"),
        ("VerticalAlignment", Alignment, "vertical"),
        ("Underline", Font, "u"),
        ("BorderStyle", Side, "style"),
        ("FillType", PatternFill, "patternType"),
    ],
)
def test_the_literal_types_match_what_openpyxl_accepts(alias, owner, attribute):
    """openpyxl validates these at runtime; the Literals only mirror it for editors.

    A value openpyxl gains and the Literal has not caught up with would be
    reported as a type error on a call that works, which is worse than no hint.
    """
    allowed = {value for value in getattr(owner, attribute).values if value is not None}

    assert set(get_args(getattr(_types, alias))) == allowed
