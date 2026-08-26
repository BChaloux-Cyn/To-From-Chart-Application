"""Sheet layout constants and builders — the single source of truth for addresses."""
from __future__ import annotations

VISIBLE = -1
VERY_HIDDEN = 2

# (tab name, VBA code name, visibility)
SHEETS = [
    ("Home", "shHome", VISIBLE),
    ("Harness", "shHarness", VISIBLE),
    ("Connectors", "shConnectors", VISIBLE),
    ("Check", "shCheck", VISIBLE),
    ("_Snapshot", "shSnapshot", VERY_HIDDEN),
    ("_Lists", "shLists", VERY_HIDDEN),
    ("_State", "shState", VERY_HIDDEN),
]


def build_sheets(wb, set_codename) -> dict:
    """Create every sheet in order, then remove Excel's defaults.

    Visibility is applied last: a sheet cannot be hidden while it is the only
    visible sheet, and Excel refuses to delete the last visible sheet.
    """
    originals = [wb.Worksheets(i + 1) for i in range(wb.Worksheets.Count)]

    sheets = {}
    anchor = originals[-1]
    for tab, codename, _ in SHEETS:
        sheet = wb.Worksheets.Add(After=anchor)
        sheet.Name = tab
        set_codename(wb, sheet, codename)
        sheets[tab] = sheet
        anchor = sheet

    for sheet in originals:
        sheet.Delete()

    for tab, _, visibility in SHEETS:
        sheets[tab].Visible = visibility

    return sheets


COLORS = [
    "Black", "White", "Red", "Green", "Blue", "Yellow", "Orange",
    "Brown", "Violet", "Gray", "Pink", "Tan", "Light Blue",
    "Light Green", "Other",
]

AWGS = ["24", "22", "20", "18", "16", "14", "12", "10", "8"]

TERMINATIONS = [
    "Crimp Pin", "Crimp Socket", "Ring Terminal", "Spade Terminal",
    "Butt Splice", "Ferrule", "Solder Cup", "Quick Disconnect",
    "Bare Tinned", "None",
]

LIST_COLUMNS = [
    ("Color", 1, COLORS),
    ("AWG", 2, AWGS),
    ("Termination", 3, TERMINATIONS),
]


def _dynamic_range(sheet: str, column_letter: str) -> str:
    """An OFFSET range that grows with its column and never has height 0."""
    return (
        f"=OFFSET('{sheet}'!${column_letter}$2,0,0,"
        f"MAX(1,COUNTA('{sheet}'!${column_letter}:${column_letter})-1),1)"
    )


LIST_NAMES = {
    "ListColor": _dynamic_range("_Lists", "A"),
    "ListAWG": _dynamic_range("_Lists", "B"),
    "ListTermination": _dynamic_range("_Lists", "C"),
    "ListRefDes": _dynamic_range("Connectors", "A"),
}


def build_lists(sheets) -> None:
    sheet = sheets["_Lists"]
    for header, column, values in LIST_COLUMNS:
        sheet.Cells(1, column).Value = header
        for offset, value in enumerate(values):
            # Text format keeps AWG sizes as strings rather than numbers.
            cell = sheet.Cells(offset + 2, column)
            cell.NumberFormat = "@"
            cell.Value = value


def build_names(wb) -> None:
    for name, refers_to in LIST_NAMES.items():
        wb.Names.Add(Name=name, RefersTo=refers_to)
