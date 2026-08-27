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


from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def seed_library_connector(wb, library_wb, tmp_path, connector_id="DTM-04P", pins=((1, "A"), (2, "B"))):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws_conn = library_wb.Worksheets("Connectors")
    ws_pins = library_wb.Worksheets("Pins")
    ws_photos = library_wb.Worksheets("Photos")

    fields = (connector_id, "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
              len(pins), "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields)
    for pin_number, label in pins:
        run(wb, "modLibrary.WritePin", ws_pins, 2, 100000, (connector_id, pin_number, label, 0.1, 0.1, 0.1, 0.1))
    shape_name = run(wb, "modLibrary.EmbedConnectorPhoto", ws_photos, connector_id, str(photo_path))
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields[:7] + (shape_name,) + fields[8:])
    return ws_conn, ws_pins, ws_photos


def test_snapshot_connector_copies_the_full_definition(wb, library_wb, tmp_path):
    ws_conn, ws_pins, ws_photos = seed_library_connector(wb, library_wb, tmp_path)
    wsnap = wb.Worksheets("_Snapshot")

    ok = run(wb, "modSnapshot.SnapshotConnector", wsnap, ws_conn, ws_pins, ws_photos, "DTM-04P")
    assert ok is True

    result = run(wb, "modLibrary.ReadConnector", wsnap, 2, 201, "DTM-04P")
    assert result[1] == "Deutsch DTM 4-way"

    pins = run(wb, "modLibrary.ReadPinsForConnector", wsnap, 211, 2210, "DTM-04P")
    assert [row[1] for row in pins] == [1, 2]

    assert wsnap.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"


def test_snapshot_connector_is_idempotent(wb, library_wb, tmp_path):
    ws_conn, ws_pins, ws_photos = seed_library_connector(wb, library_wb, tmp_path)
    wsnap = wb.Worksheets("_Snapshot")

    run(wb, "modSnapshot.SnapshotConnector", wsnap, ws_conn, ws_pins, ws_photos, "DTM-04P")
    run(wb, "modSnapshot.SnapshotConnector", wsnap, ws_conn, ws_pins, ws_photos, "DTM-04P")

    assert wsnap.Cells(3, 1).Value is None  # not duplicated on a second call


def test_snapshot_connector_for_unknown_id_returns_false(wb, library_wb):
    wsnap = wb.Worksheets("_Snapshot")
    ws_conn = library_wb.Worksheets("Connectors")
    ws_pins = library_wb.Worksheets("Pins")
    ws_photos = library_wb.Worksheets("Photos")

    assert run(wb, "modSnapshot.SnapshotConnector", wsnap, ws_conn, ws_pins, ws_photos, "NO-SUCH-ID") is False
