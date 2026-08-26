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
