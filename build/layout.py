"""Sheet layout constants and builders — the single source of truth for addresses."""
from __future__ import annotations

VISIBLE = -1
VERY_HIDDEN = 2

import library_layout

# (tab name, VBA code name, visibility)
SHEETS = [
    ("Home", "shHome", VISIBLE),
    ("Harness", "shHarness", VISIBLE),
    ("Connectors", "shConnectors", VISIBLE),
    ("Check", "shCheck", VISIBLE),
    ("_Snapshot", "shSnapshot", VERY_HIDDEN),
    ("_Edit", "shEdit", VERY_HIDDEN),
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

# A and G are widened past their chart-only sizing so the title-block labels
# that also live in those columns (Harness Name/Description; Length Units)
# don't clip.
CHART_COLUMN_WIDTHS = [14, 9, 15, 18, 13, 7, 13, 15, 11, 9, 30]

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

# Widens each free-text title-block value past its single narrow chart-grid
# column so the gray fill (and the text) doesn't stop short when a value -
# Harness Name and Description especially - is longer than one column.
# H4 (Length Units) is a short controlled value and is left unmerged.
TB_MERGE_SPANS = {
    "B2": "C2",
    "E2": "F2",
    "H2": "I2",
    "B3": "C3",
    "E3": "F3",
    "H3": "I3",
    "B4": "F4",
}

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
XL_CONTINUOUS = 1
XL_THIN = 2


def build_harness(sheets) -> None:
    sheet = sheets["Harness"]

    sheet.Range("A1").Value = "WIRE HARNESS TO-FROM CHART"
    sheet.Range("A1").Font.Size = 16
    sheet.Range("A1").Font.Bold = True

    for label, label_cell, value_cell in TITLE_BLOCK:
        sheet.Range(label_cell).Value = label
        sheet.Range(label_cell).Font.Bold = True

        span = TB_MERGE_SPANS.get(value_cell)
        value_range = sheet.Range(f"{value_cell}:{span}") if span else sheet.Range(value_cell)
        value_range.Interior.Color = 0xF2F2F2
        if span:
            value_range.Merge()
        value_range.Borders.LineStyle = XL_CONTINUOUS
        value_range.Borders.Weight = XL_THIN
        value_range.Borders.Color = 0x000000

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


XL_VALIDATE_DECIMAL = 2
XL_GREATER = 5

# (column index, validation type, Formula1)
CHART_VALIDATION = [
    (1, XL_VALIDATE_LIST, "=ListRefDes"),
    (3, XL_VALIDATE_LIST, "=ListTermination"),
    (5, XL_VALIDATE_LIST, "=ListColor"),
    (6, XL_VALIDATE_LIST, "=ListAWG"),
    (7, XL_VALIDATE_DECIMAL, "0"),
    (8, XL_VALIDATE_LIST, "=ListTermination"),
    (9, XL_VALIDATE_LIST, "=ListRefDes"),
]


def build_chart_validation(sheets) -> None:
    sheet = sheets["Harness"]
    for column, vtype, formula in CHART_VALIDATION:
        target = sheet.Range(
            sheet.Cells(CHART_FIRST_ROW, column),
            sheet.Cells(CHART_LAST_ROW, column),
        )
        target.Validation.Delete()
        if vtype == XL_VALIDATE_DECIMAL:
            target.Validation.Add(
                Type=vtype,
                AlertStyle=XL_VALID_ALERT_STOP,
                Operator=XL_GREATER,
                Formula1=formula,
            )
        else:
            target.Validation.Add(
                Type=vtype,
                AlertStyle=XL_VALID_ALERT_STOP,
                Operator=XL_BETWEEN,
                Formula1=formula,
            )
            target.Validation.InCellDropdown = True
        target.Validation.IgnoreBlank = True


CONN_HEADERS = ["Ref Des", "ConnectorID", "Name", "Part Number", "Type", "Pin Count"]
CONN_COLUMN_WIDTHS = [10, 16, 28, 18, 12, 10]
CONN_FIRST_ROW = 2

CHECK_HEADERS = ["Row", "Severity", "Message"]
CHECK_COLUMN_WIDTHS = [8, 12, 80]

STATE_KEYS = ["BuildVersion", "HarnessPath", "Dirty", "LengthUnits", "TestMode"]
STATE_DEFAULTS = {
    "HarnessPath": "",
    "Dirty": "FALSE",
    "LengthUnits": "in",
    "TestMode": "FALSE",
}


def _build_header_sheet(sheet, headers, widths) -> None:
    for index, header in enumerate(headers, start=1):
        cell = sheet.Cells(1, index)
        cell.Value = header
        cell.Font.Bold = True
        cell.Interior.Color = 0xD9D9D9
        sheet.Columns(index).ColumnWidth = widths[index - 1]


def build_connectors(sheets) -> None:
    _build_header_sheet(sheets["Connectors"], CONN_HEADERS, CONN_COLUMN_WIDTHS)


def build_check(sheets) -> None:
    _build_header_sheet(sheets["Check"], CHECK_HEADERS, CHECK_COLUMN_WIDTHS)


def build_state(sheets, version: str) -> None:
    sheet = sheets["_State"]
    sheet.Cells(1, 1).Value = "Key"
    sheet.Cells(1, 2).Value = "Value"
    values = dict(STATE_DEFAULTS, BuildVersion=version)
    for offset, key in enumerate(STATE_KEYS):
        row = offset + 2
        sheet.Cells(row, 1).Value = key
        sheet.Cells(row, 2).NumberFormat = "@"
        sheet.Cells(row, 2).Value = values[key]


MSO_SHAPE_ROUNDED_RECTANGLE = 5

HOME_TEXT = [
    ("A1", "WIRE HARNESS CREATOR"),
    ("A3", "This workbook is the editor. It is not a drawing."),
    ("A4", "Use it to build harness files, which are saved separately as .xlsx."),
    ("A6", "1. Add the connectors your harness uses on the Connectors sheet."),
    ("A7", "2. Fill in the to-from chart on the Harness sheet, one row per wire."),
    ("A8", "3. Pick From Conn first - the From Pin list is built from that connector."),
    ("A10", "New Harness clears everything and starts over."),
]

HOME_BUTTONS = [
    # (caption, macro, left, top, width, height)
    ("New Harness", "modChart.NewHarness", 20, 220, 120, 32),
    ("Save Harness", "modHarnessUI.SaveHarness", 150, 220, 120, 32),
    ("Save Harness As", "modHarnessUI.SaveHarnessAs", 280, 220, 120, 32),
    ("Open Harness", "modHarnessUI.OpenHarness", 410, 220, 120, 32),
    ("Add Connector", "modConnectorUI.ShowAddConnector", 20, 260, 120, 32),
    ("Remove Connector", "modConnectorUI.ShowRemoveConnector", 150, 260, 130, 32),
    ("Manage Library", "modConnectorUI.ShowManageLibrary", 20, 300, 120, 32),
]


def build_home(sheets) -> None:
    sheet = sheets["Home"]
    for cell, text in HOME_TEXT:
        sheet.Range(cell).Value = text
    sheet.Range("A1").Font.Size = 16
    sheet.Range("A1").Font.Bold = True
    sheet.Columns(1).ColumnWidth = 90

    for caption, macro, left, top, width, height in HOME_BUTTONS:
        shape = sheet.Shapes.AddShape(
            MSO_SHAPE_ROUNDED_RECTANGLE, left, top, width, height
        )
        shape.TextFrame2.TextRange.Text = caption
        shape.OnAction = macro


SNAP_CONN_FIRST_ROW = 2
SNAP_CONN_LAST_ROW = 201
SNAP_PINS_HEADER_ROW = 210
SNAP_PINS_FIRST_ROW = 211
SNAP_PINS_LAST_ROW = 2210


def build_snapshot(sheets) -> None:
    sheet = sheets["_Snapshot"]
    for index, header in enumerate(library_layout.CONN_HEADERS, start=1):
        sheet.Cells(1, index).Value = header
    for index, header in enumerate(library_layout.PIN_HEADERS, start=1):
        sheet.Cells(SNAP_PINS_HEADER_ROW, index).Value = header
