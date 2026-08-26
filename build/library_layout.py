"""Sheet layout for ConnectorLibrary.xlsx — macro-free, three tables."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from layout import VISIBLE

SHEET_NAMES = ["Connectors", "Pins", "Photos"]

CONN_HEADERS = [
    "ConnectorID", "Name", "Manufacturer", "PartNumber", "Type",
    "PinCount", "Notes", "PhotoShapeName", "CreatedUtc", "ModifiedUtc", "Origin",
]
CONN_COLUMN_WIDTHS = [16, 24, 16, 18, 12, 10, 24, 20, 20, 20, 14]

PIN_HEADERS = ["ConnectorID", "PinNumber", "PinLabel", "NormX", "NormY", "LabelX", "LabelY"]
PIN_COLUMN_WIDTHS = [16, 10, 16, 10, 10, 10, 10]


def _build_header_sheet(sheet, headers, widths) -> None:
    for index, header in enumerate(headers, start=1):
        cell = sheet.Cells(1, index)
        cell.Value = header
        cell.Font.Bold = True
        cell.Interior.Color = 0xD9D9D9
        sheet.Columns(index).ColumnWidth = widths[index - 1]


def build_library_sheets(wb) -> dict:
    """Create Connectors, Pins, and Photos, then remove Excel's defaults.

    No VBA project is touched here - no code names, nothing imported. That
    is what keeps the saved .xlsx macro-free.
    """
    originals = [wb.Worksheets(i + 1) for i in range(wb.Worksheets.Count)]

    sheets = {}
    anchor = originals[-1]
    for name in SHEET_NAMES:
        sheet = wb.Worksheets.Add(After=anchor)
        sheet.Name = name
        sheet.Visible = VISIBLE
        sheets[name] = sheet
        anchor = sheet

    for sheet in originals:
        sheet.Delete()

    _build_header_sheet(sheets["Connectors"], CONN_HEADERS, CONN_COLUMN_WIDTHS)
    _build_header_sheet(sheets["Pins"], PIN_HEADERS, PIN_COLUMN_WIDTHS)
    sheets["Photos"].Range("A1").Value = "Connector Photos"
    sheets["Photos"].Range("A1").Font.Bold = True

    return sheets
