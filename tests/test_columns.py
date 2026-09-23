"""Behavior of ``column_letter`` and ``column_index``.

The pair openpyxl ships accepts columns up to ZZZ, which is 1894 columns past the
end of the grid. What is asserted here is mostly the refusal.
"""

import pytest

from openpyxl_toolkit import column_index, column_letter
from openpyxl_toolkit._limits import MAX_COLUMN


@pytest.mark.parametrize(
    ("index", "letter"),
    [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (702, "ZZ"), (703, "AAA"), (16384, "XFD")],
)
def test_a_column_number_becomes_its_letter(index, letter):
    assert column_letter(index) == letter


@pytest.mark.parametrize(
    ("letter", "index"),
    [("A", 1), ("Z", 26), ("AA", 27), ("AZ", 52), ("ZZ", 702), ("AAA", 703), ("XFD", 16384)],
)
def test_a_column_letter_becomes_its_number(letter, index):
    assert column_index(letter) == index


def test_every_column_in_the_grid_survives_a_round_trip():
    """Off-by-one errors in base-26 hide at the 26 and 702 boundaries."""
    assert [column_index(column_letter(i)) for i in range(1, MAX_COLUMN + 1)] == list(
        range(1, MAX_COLUMN + 1)
    )


def test_lower_case_is_accepted():
    assert column_index("aa") == 27


def test_surrounding_whitespace_is_ignored():
    assert column_index("  c  ") == 3


@pytest.mark.parametrize("index", [MAX_COLUMN + 1, 18278, 100000])
def test_a_number_past_the_last_column_is_rejected(index):
    """openpyxl returns XFE and beyond, which Excel then refuses to open."""
    with pytest.raises(ValueError, match="outside the worksheet"):
        column_letter(index)


@pytest.mark.parametrize("index", [0, -1])
def test_a_number_below_the_first_column_is_rejected(index):
    with pytest.raises(ValueError, match="outside the worksheet"):
        column_letter(index)


@pytest.mark.parametrize("letter", ["XFE", "ZZZ", "AAAA"])
def test_a_letter_past_the_last_column_is_rejected(letter):
    with pytest.raises(ValueError, match="outside the worksheet"):
        column_index(letter)


@pytest.mark.parametrize("index", [1.0, True, None, "3"])
def test_a_column_number_that_is_not_an_integer_is_rejected(index):
    """A float reaches openpyxl as a list index and raises from inside it."""
    with pytest.raises(TypeError, match="must be an integer"):
        column_letter(index)


@pytest.mark.parametrize("letter", [None, 3, ["A"]])
def test_a_column_letter_that_is_not_a_string_is_rejected(letter):
    with pytest.raises(TypeError, match="must be a string"):
        column_index(letter)


@pytest.mark.parametrize("letter", ["", "!", "A-", "Ä"])
def test_something_that_is_not_letters_is_rejected(letter):
    with pytest.raises(ValueError, match="is not a column letter"):
        column_index(letter)


def test_a_cell_reference_is_rejected_rather_than_read_as_its_column():
    """Reading it as column C would accept a range where a column was meant."""
    with pytest.raises(ValueError, match="carries no row number"):
        column_index("C1")


def test_the_row_number_hint_is_left_off_when_no_row_number_was_given():
    with pytest.raises(ValueError) as caught:
        column_index("")

    assert "row number" not in str(caught.value)


def test_the_number_of_a_column_matches_what_openpyxl_uses_for_it(sheet):
    """The number has to be the one openpyxl indexes cells by, not merely reversible."""
    cell = sheet.cell(row=1, column=column_index("D"), value="x")

    assert cell.coordinate == "D1"


def test_the_letter_of_a_column_matches_what_openpyxl_calls_it(sheet):
    assert column_letter(4) == sheet.cell(row=1, column=4, value="x").column_letter
