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


CHART_HEADER_ROW = 6
CHART_FIRST_ROW = 7
CHART_LAST_ROW = 1006

CHART_HEADERS = [
    "From Conn", "From Pin", "From Term", "Signal", "Color", "AWG",
    "Length (in)", "To Term", "To Conn", "To Pin", "Notes",
]

CHART_COLUMN_WIDTHS = [11, 9, 15, 18, 13, 7, 12, 15, 11, 9, 30]

# (label text, label cell, value cell)
TITLE_BLOCK = [
    ("Harness Name", "A2", "B2"),
    ("Harness Number", "D2", "E2"),
    ("Revision", "G2", "H2"),
    ("Student", "A3", "B3"),
    ("Class / Project", "D3", "E3"),
    ("Date", "G3", "H3"),
    ("Description", "A4", "B4"),
    ("Length Units", "G4", "H4"),
]

TB_NAMES = {
    "TB_Name": "B2",
    "TB_Number": "E2",
    "TB_Rev": "H2",
    "TB_Student": "B3",
    "TB_Class": "E3",
    "TB_Date": "H3",
    "TB_Desc": "B4",
    "TB_Units": "H4",
}

XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1


def build_harness(sheets) -> None:
    sheet = sheets["Harness"]

    sheet.Range("A1").Value = "WIRE HARNESS TO-FROM CHART"
    sheet.Range("A1").Font.Size = 16
    sheet.Range("A1").Font.Bold = True

    for label, label_cell, value_cell in TITLE_BLOCK:
        sheet.Range(label_cell).Value = label
        sheet.Range(label_cell).Font.Bold = True
        sheet.Range(value_cell).Interior.Color = 0xF2F2F2

    sheet.Range(TB_NAMES["TB_Units"]).Value = "in"
    units = sheet.Range(TB_NAMES["TB_Units"]).Validation
    units.Delete()
    units.Add(
        Type=XL_VALIDATE_LIST,
        AlertStyle=XL_VALID_ALERT_STOP,
        Operator=XL_BETWEEN,
        Formula1="in,mm",
    )
    units.IgnoreBlank = True
    units.InCellDropdown = True

    for index, header in enumerate(CHART_HEADERS, start=1):
        cell = sheet.Cells(CHART_HEADER_ROW, index)
        cell.Value = header
        cell.Font.Bold = True
        cell.Interior.Color = 0xD9D9D9
        sheet.Columns(index).ColumnWidth = CHART_COLUMN_WIDTHS[index - 1]


def build_title_block_names(wb, sheets) -> None:
    for name, cell in TB_NAMES.items():
        wb.Names.Add(Name=name, RefersTo=f"='Harness'!${cell[0]}${cell[1:]}")
