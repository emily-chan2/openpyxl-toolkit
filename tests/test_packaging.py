"""The package metadata itself is part of the contract."""

from importlib.metadata import version

import openpyxl_toolkit


def test_version_is_single_sourced():
    assert openpyxl_toolkit.__version__ == version("openpyxl-toolkit")


def test_public_surface():
    assert openpyxl_toolkit.__all__ == ["WorksheetToolkit"]
    assert openpyxl_toolkit.WorksheetToolkit.__module__ == "openpyxl_toolkit.worksheet_toolkit"
