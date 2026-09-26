"""Behavior of the hex colors every color argument accepts.

openpyxl turns a bad color down with ``Colors must be aRGB hex values``, naming
neither the argument nor the value and describing a form almost nobody writes.
The toolkit checks first, so what is asserted here is mostly the refusal and the
wording of it.
"""

import re

import pytest
from openpyxl.styles import PatternFill

from openpyxl_toolkit import WorksheetToolkit
from openpyxl_toolkit._colors import normalize_color

ACCEPTED = [
    ("#f4d2d3", "FFF4D2D3"),
    ("f4d2d3", "FFF4D2D3"),
    ("#F4D2D3", "FFF4D2D3"),
    ("FFF4D2D3", "FFF4D2D3"),
    ("#fff4d2d3", "FFF4D2D3"),
    ("  #F4D2D3  ", "FFF4D2D3"),
    ("80f4d2d3", "80F4D2D3"),
]

# Everything here reaches openpyxl as a color it will not store, or -- in the case
# of the three-digit form -- as one it would store as a different color.
REJECTED = ["red", "#f00", "#f00a", "#12345", "ff0000ff00", "", "#", "0x00FF00", "gggggg"]


@pytest.mark.parametrize(("written", "stored"), ACCEPTED)
def test_an_accepted_color_survives_the_round_trip(sheet, roundtrip, written, stored):
    WorksheetToolkit(sheet).set_font(cells="A1", color=written)

    assert roundtrip(sheet)["A1"].font.color.rgb.upper() == stored


def test_six_digits_are_opaque_rather_than_transparent(sheet, roundtrip):
    """Handed six digits, openpyxl stores an alpha of 00, which shows nothing."""
    WorksheetToolkit(sheet).set_fill(cells="A1", fill_type="solid", start_color="f4d2d3")

    assert roundtrip(sheet)["A1"].fill.start_color.rgb.upper().startswith("FF")


@pytest.mark.parametrize("color", REJECTED)
def test_a_value_that_is_not_a_color_is_rejected(color):
    with pytest.raises(ValueError, match="must be a hex color"):
        normalize_color(color)


@pytest.mark.parametrize("color", REJECTED)
def test_the_rejected_value_is_quoted_back(color):
    with pytest.raises(ValueError, match=re.escape(f"got {color!r}")):
        normalize_color(color)


@pytest.mark.parametrize("color", ["#f00", "#12345", "ff0000ff00"])
def test_a_hex_value_of_the_wrong_length_is_told_how_many_digits(color):
    with pytest.raises(ValueError, match="six digits"):
        normalize_color(color)


@pytest.mark.parametrize("color", ["red", "0x00FF00", "gggggg", ""])
def test_a_value_that_is_not_hex_is_given_no_digit_count(color):
    """Digit counts answer a question the writer of 'red' did not ask."""
    with pytest.raises(ValueError, match="must be a hex color") as raised:
        normalize_color(color)

    assert "six digits" not in str(raised.value)


def test_each_argument_is_rejected_under_its_own_name(sheet):
    toolkit = WorksheetToolkit(sheet)
    calls = {
        "color": [
            lambda: toolkit.set_font(cells="A1", color="red"),
            lambda: toolkit.set_border(cells="A1", sides="left", style="thin", color="red"),
            lambda: toolkit.set_outside_border(cells="A1:B2", style="thin", color="red"),
            lambda: toolkit.set_tab_color("red"),
        ],
        "start_color": [lambda: toolkit.set_fill(cells="A1", fill_type="solid", start_color="red")],
        "end_color": [
            lambda: toolkit.set_fill(
                cells="A1", fill_type="solid", start_color="#f4d2d3", end_color="red"
            )
        ],
    }
    for name, group in calls.items():
        for call in group:
            with pytest.raises(ValueError, match=f"^{name} must be a hex color"):
                call()


def test_an_empty_start_color_is_turned_down_as_a_bad_color_not_a_missing_one(sheet):
    """An empty string reads as absent, which named the wrong problem."""
    with pytest.raises(ValueError, match="^start_color must be a hex color"):
        WorksheetToolkit(sheet).set_fill(cells="A1", fill_type="solid", start_color="")


@pytest.mark.parametrize(
    "call",
    [
        lambda toolkit: toolkit.set_border(rows=[], sides="left", style="thin", color="red"),
        lambda toolkit: toolkit.set_fill(rows=[], fill_type="solid", start_color="red"),
    ],
)
def test_a_selection_holding_no_cells_still_rejects_the_color(sheet, call):
    """The fault is in the call, so it does not depend on the range finding a cell."""
    with pytest.raises(ValueError, match="must be a hex color"):
        call(WorksheetToolkit(sheet))


def test_a_rejected_color_leaves_every_cell_as_it_was(sheet, roundtrip):
    for row in range(1, 4):
        for column in range(1, 4):
            sheet.cell(row=row, column=column, value="x")
    sheet["A1"].fill = PatternFill("solid", start_color="FF00FF00")
    toolkit = WorksheetToolkit(sheet)

    for call in (
        lambda: toolkit.set_fill(rows=[1, 2, 3], fill_type="solid", start_color="red"),
        lambda: toolkit.set_border(rows=[1, 2, 3], sides="left", style="thin", color="red"),
        lambda: toolkit.set_font(rows=[1, 2, 3], color="red"),
    ):
        with pytest.raises(ValueError):
            call()

    reloaded = roundtrip(sheet)
    assert reloaded["A1"].fill.start_color.rgb.upper() == "FF00FF00"
    assert reloaded["B2"].fill.fill_type is None
    assert reloaded["B2"].border.left.style is None


def test_none_still_clears_rather_than_being_rejected(sheet, roundtrip):
    toolkit = WorksheetToolkit(sheet)
    toolkit.set_tab_color("#1d3557")
    toolkit.set_tab_color(None)

    assert roundtrip(sheet).sheet_properties.tabColor is None
