# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Packaged the project for distribution: `src/` layout, `pyproject.toml` (hatchling),
  and `openpyxl_toolkit` as the import name. `WorksheetToolkit` is re-exported from the
  package root, so `from openpyxl_toolkit import WorksheetToolkit` works.
- Test suite covering the save/reload round trip, and continuous integration across
  Python 3.10-3.14.

[Unreleased]: https://github.com/emily-chan2/openpyxl-toolkit/compare/HEAD
