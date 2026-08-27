import library_layout

from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_export_connector_copies_record_pins_and_photo(wb, library_wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws_src_conn = library_wb.Worksheets("Connectors")
    ws_src_pins = library_wb.Worksheets("Pins")
    ws_src_photos = library_wb.Worksheets("Photos")

    fields = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
              2, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws_src_conn, 2, 100000, fields)
    run(wb, "modLibrary.WritePin", ws_src_pins, 2, 100000, ("DTM-04P", 1, "A", 0.1, 0.1, 0.1, 0.1))
    run(wb, "modLibrary.EmbedConnectorPhoto", ws_src_photos, "DTM-04P", str(photo_path))

    dest_wb = app.Workbooks.Add()
    try:
        library_layout.build_library_sheets(dest_wb)
        ws_dest_conn = dest_wb.Worksheets("Connectors")
        ws_dest_pins = dest_wb.Worksheets("Pins")
        ws_dest_photos = dest_wb.Worksheets("Photos")

        ok = run(wb, "modLibraryTransfer.ExportConnector", ws_src_conn, ws_src_pins, ws_src_photos,
                 ws_dest_conn, ws_dest_pins, ws_dest_photos, "DTM-04P")
        assert ok is True

        result = run(wb, "modLibrary.ReadConnector", ws_dest_conn, 2, 100000, "DTM-04P")
        assert result[1] == "Deutsch DTM 4-way"
        pins = run(wb, "modLibrary.ReadPinsForConnector", ws_dest_pins, 2, 100000, "DTM-04P")
        assert len(pins) == 1
        assert ws_dest_photos.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"
    finally:
        dest_wb.Close(SaveChanges=False)


def test_export_unknown_connector_returns_false(wb, library_wb, app):
    dest_wb = app.Workbooks.Add()
    try:
        library_layout.build_library_sheets(dest_wb)
        ok = run(
            wb, "modLibraryTransfer.ExportConnector",
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"),
            dest_wb.Worksheets("Connectors"), dest_wb.Worksheets("Pins"), dest_wb.Worksheets("Photos"),
            "NO-SUCH-ID",
        )
        assert ok is False
    finally:
        dest_wb.Close(SaveChanges=False)
