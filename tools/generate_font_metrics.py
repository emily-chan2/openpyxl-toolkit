"""Regenerate src/openpyxl_toolkit/_font_widths.py from font files.

Not part of the package. Run it when adding a font or refreshing the tables:

    pip install fonttools
    python tools/generate_font_metrics.py

Widths are stored as a fraction of the em, which makes them independent of point
size. Where an openly licensed font exists that was drawn to match a Microsoft
one advance-for-advance, that is the one measured, so the shipped numbers do not
come from a proprietary file. The match is verified, not assumed -- see
tests/test_metrics.py.
"""

import sys
from collections import defaultdict
from pathlib import Path

try:
    from fontTools.ttLib import TTCollection, TTFont
except ImportError:
    sys.exit("fonttools is needed to regenerate the tables: pip install fonttools")

# Faces measured from openly licensed fonts drawn to match a proprietary one
# advance-for-advance. Filename is relative to the directory passed on the command
# line; download them from github.com/google/fonts.
OPEN_SOURCES = {
    "calibri": ("Carlito-Regular.ttf", "Carlito, metrically identical to Calibri (OFL)"),
    "arial": ("Arimo.ttf", "Arimo, metrically identical to Arial and Helvetica (OFL)"),
    "times new roman": ("Tinos.ttf", "Tinos, metrically identical to Times New Roman (OFL)"),
    "courier new": ("Cousine.ttf", "Cousine, metrically identical to Courier New (OFL)"),
    "cambria": ("Caladea.ttf", "Caladea, metrically identical to Cambria (OFL)"),
}

# Faces with no open equivalent, measured from the font files installed on the
# machine that generated this table. Only the advance widths are taken: a table of
# numbers describing how wide each character is, not any part of the font itself.
# Absolute paths, and a face name where the file is a collection.
SYSTEM_SOURCES = {
    "aptos": (
        "/Applications/Microsoft Excel.app/Contents/Resources/DFonts/Aptos.ttf",
        None,
        "Aptos, the current Excel default, from a Microsoft Office installation",
    ),
    "verdana": (
        "/System/Library/Fonts/Supplemental/Verdana.ttf",
        None,
        "Verdana, from macOS",
    ),
    "georgia": (
        "/System/Library/Fonts/Supplemental/Georgia.ttf",
        None,
        "Georgia, from macOS",
    ),
    "tahoma": (
        "/System/Library/Fonts/Supplemental/Tahoma.ttf",
        None,
        "Tahoma, from macOS",
    ),
    "futura": (
        "/System/Library/Fonts/Supplemental/Futura.ttc",
        "Futura Medium",
        "Futura Medium, macOS's regular weight of Futura",
    ),
}

# faces that share another face's advances exactly
ALIASES = {
    "helvetica": "arial",
    "liberation sans": "arial",
    "arimo": "arial",
    "carlito": "calibri",
    "liberation serif": "times new roman",
    "tinos": "times new roman",
    "liberation mono": "courier new",
    "cousine": "courier new",
    "caladea": "cambria",
}

CODEPOINTS = list(range(32, 127)) + list(range(160, 256))


def open_face(path, face_name=None):
    """The font, picking a named face when the file is a collection."""
    if path.suffix != ".ttc":
        return TTFont(str(path))
    faces = TTCollection(str(path)).fonts
    if face_name is None:
        return faces[0]
    for face in faces:
        if (face["name"].getDebugName(4) or "").strip() == face_name:
            return face
    sys.exit(f"{path} has no face named {face_name!r}")


def advances(path, face_name=None):
    """Character -> advance width as a fraction of the em."""
    font = open_face(path, face_name)
    upem = font["head"].unitsPerEm
    cmap, hmtx = font.getBestCmap(), font["hmtx"]
    widths = {}
    for code in CODEPOINTS:
        glyph = cmap.get(code)
        if glyph is not None:
            widths[chr(code)] = round(hmtx[glyph][0] / upem, 4)
    return widths


def as_groups(widths):
    """Characters sharing a width, so the table shows the shape of the font."""
    grouped = defaultdict(list)
    for character, width in widths.items():
        grouped[width].append(character)
    return [(w, "".join(sorted(grouped[w]))) for w in sorted(grouped, reverse=True)]


def main(font_dir):
    font_dir = Path(font_dir)
    wanted = [(name, font_dir / f, None, note) for name, (f, note) in OPEN_SOURCES.items()]
    wanted += [(name, Path(p), face, note) for name, (p, face, note) in SYSTEM_SOURCES.items()]

    blocks = []
    for name, path, face, note in wanted:
        if not path.exists():
            print(f"  skipping {name}: {path} not present", file=sys.stderr)
            continue
        groups = as_groups(advances(path, face))
        rows = "\n".join(f"        {w!r}: {chars!r}," for w, chars in groups)
        blocks.append(f'    # {note}\n    "{name}": {{\n{rows}\n    }},')

    body = "\n".join(blocks)
    aliases = "\n".join(f'    "{k}": "{v}",' for k, v in sorted(ALIASES.items()))
    out = Path("src/openpyxl_toolkit/_font_widths.py")
    out.write_text(
        '"""Per-character advance widths, as a fraction of the em.\n\n'
        "Generated by tools/generate_font_metrics.py. Do not edit by hand.\n\n"
        "Each table is keyed by width so that characters sharing one stay visible\n"
        "rather than being repeated two hundred times.\n"
        '"""\n\n'
        f"WIDTH_GROUPS = {{\n{body}\n}}\n\n"
        f"ALIASES = {{\n{aliases}\n}}\n\n"
        "FONT_WIDTHS = {\n"
        "    name: {character: width for width, characters in groups.items() "
        "for character in characters}\n"
        "    for name, groups in WIDTH_GROUPS.items()\n"
        "}\n"
    )
    print(f"wrote {out} for {len(blocks)} fonts and {len(ALIASES)} aliases")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "/tmp/carlito")
