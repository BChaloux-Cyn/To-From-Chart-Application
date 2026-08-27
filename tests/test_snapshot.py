def test_snapshot_connectors_header_matches_the_library_schema(wb):
    sheet = wb.Worksheets("_Snapshot")
    conn_headers = [
        "ConnectorID", "Name", "Manufacturer", "PartNumber", "Type",
        "PinCount", "Notes", "PhotoShapeName", "CreatedUtc", "ModifiedUtc", "Origin",
    ]
    for index, header in enumerate(conn_headers, start=1):
        assert sheet.Cells(1, index).Value == header


def test_snapshot_pins_header_matches_the_library_schema(wb):
    sheet = wb.Worksheets("_Snapshot")
    pin_headers = ["ConnectorID", "PinNumber", "PinLabel", "NormX", "NormY", "LabelX", "LabelY"]
    for index, header in enumerate(pin_headers, start=1):
        assert sheet.Cells(210, index).Value == header
