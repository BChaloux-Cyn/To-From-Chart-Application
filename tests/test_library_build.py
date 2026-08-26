import pytest


def test_library_artifact_is_produced(library_artifact):
    assert library_artifact.exists()
    assert library_artifact.suffix == ".xlsx"


def test_library_has_three_sheets_in_order(library_wb):
    names = [library_wb.Worksheets(i + 1).Name for i in range(library_wb.Worksheets.Count)]
    assert names == ["Connectors", "Pins", "Photos"]


CONN_HEADERS = [
    "ConnectorID", "Name", "Manufacturer", "PartNumber", "Type",
    "PinCount", "Notes", "PhotoShapeName", "CreatedUtc", "ModifiedUtc", "Origin",
]


@pytest.mark.parametrize("index,header", list(enumerate(CONN_HEADERS, start=1)))
def test_connectors_sheet_headers(library_wb, index, header):
    assert library_wb.Worksheets("Connectors").Cells(1, index).Value == header


PIN_HEADERS = ["ConnectorID", "PinNumber", "PinLabel", "NormX", "NormY", "LabelX", "LabelY"]


@pytest.mark.parametrize("index,header", list(enumerate(PIN_HEADERS, start=1)))
def test_pins_sheet_headers(library_wb, index, header):
    assert library_wb.Worksheets("Pins").Cells(1, index).Value == header


def test_library_workbook_has_no_vba_modules(library_wb):
    # Every open workbook exposes implicit document modules (ThisWorkbook +
    # one per sheet, vbext_ct_Document = 100) regardless of macros - that's
    # Excel's in-memory object model, not evidence of an imported module.
    # A real standard/class module (StdModule=1, ClassModule=2) would mean
    # this workbook is not macro-free.
    VBEXT_CT_DOCUMENT = 100
    types = {c.Type for c in library_wb.VBProject.VBComponents}
    assert types <= {VBEXT_CT_DOCUMENT}
