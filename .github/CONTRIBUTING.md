# Contributing

One maintainer, so this is short and practical.

## Scope

openpyxl-toolkit formats cells on a worksheet that already exists, merging what it
is given into the formatting a cell already carries. It does not write values, and
it does nothing at workbook level: charts, images, pivot tables, data validation
and print setup are all outside it for now. Something outside that scope is likelier
to be declined than built, so it is worth raising before writing it.

## Reporting something

Bugs and feature requests go through the issue forms. They ask for the things that
turn out to matter here: where the workbook came from, and whether the problem
survives a save and a reload.

Never open a public issue for a security problem. Use **Report a vulnerability** on
the Security tab.

## Changing something

A bug fix, a typo or a docstring correction can go straight to a pull request.
Anything that adds an argument, adds a method, or changes what an existing call
does wants an issue first, so the shape can be settled before the work.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -e . --group dev
```

Upgrading pip is not ceremony: `--group` needs pip 25.1, and a virtualenv built
with Python 3.10 or 3.11 still ships pip 24. The install has to be editable and has
to include the project itself — the dependency group carries the tools but not
openpyxl, and one test reads the installed package's own metadata.

## Checks

All four run from the repository root, and together take about three seconds.

```bash
pytest
ruff check .
ruff format --check .
mypy
```

`README.md` is part of the suite: its examples run as doctests, so editing the prose
around them can turn the run red. `mypy` reads a relative path out of
`pyproject.toml` and only works from the root.

## Tests

Every behavioural test saves a real `.xlsx` and loads it back, through the
`roundtrip` fixture. A style that is right in memory and wrong in the file is the
failure this library exists to prevent, so an assertion against the in-memory cell
does not count for much.

A fix wants a test that fails without it. Revert the change and watch it fail
before opening the pull request.

## Changelog

Add a line to `CHANGELOG.md` under `Unreleased`, in whichever section fits. If the
change is invisible to someone using the library, say so in the pull request and
leave the changelog alone.

## Generated code

Code written with an AI assistant is fine, as long as you have read it, run it and
can explain it. Pull requests that read as machine-generated and unreviewed are
closed without a detailed reply.
