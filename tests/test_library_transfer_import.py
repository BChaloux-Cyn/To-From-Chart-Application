import library_layout

from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def seed_source_connector(wb, app, tmp_path, connector_id="DTM-04P", name="Deutsch DTM 4-way"):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    src_wb = app.Workbooks.Add()
    library_layout.build_library_sheets(src_wb)
    ws_conn = src_wb.Worksheets("Connectors")
    ws_pins = src_wb.Worksheets("Pins")
    ws_photos = src_wb.Worksheets("Photos")

    fields = (connector_id, name, "Deutsch", "DTM06-4S", "Connector",
              1, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields)
    run(wb, "modLibrary.WritePin", ws_pins, 2, 100000, (connector_id, 1, "A", 0.1, 0.1, 0.1, 0.1))
    run(wb, "modLibrary.EmbedConnectorPhoto", ws_photos, connector_id, str(photo_path))
    return src_wb, ws_conn, ws_pins, ws_photos


def test_import_with_no_collision_keeps_its_id(wb, library_wb, app, tmp_path):
    src_wb, ws_src_conn, ws_src_pins, ws_src_photos = seed_source_connector(wb, app, tmp_path)
    try:
        dest_id = run(
            wb, "modLibraryTransfer.ImportConnector", ws_src_conn, ws_src_pins, ws_src_photos,
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"),
            "DTM-04P", "other_library.xlsx",
        )
        assert dest_id == "DTM-04P"

        result = run(wb, "modLibrary.ReadConnector", library_wb.Worksheets("Connectors"), 2, 100000, "DTM-04P")
        assert result[10] == "other_library.xlsx"
    finally:
        src_wb.Close(SaveChanges=False)


def test_import_colliding_with_a_local_id_is_renamed_and_the_original_kept(wb, library_wb, app, tmp_path):
    ws_dest_conn = library_wb.Worksheets("Connectors")
    existing = ("DTM-04P", "A different local part", "", "", "Connector", 1, "", "", "", "", "Local")
    run(wb, "modLibrary.WriteConnector", ws_dest_conn, 2, 100000, existing)

    src_wb, ws_src_conn, ws_src_pins, ws_src_photos = seed_source_connector(wb, app, tmp_path)
    try:
        dest_id = run(
            wb, "modLibraryTransfer.ImportConnector", ws_src_conn, ws_src_pins, ws_src_photos,
            ws_dest_conn, library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"),
            "DTM-04P", "other_library.xlsx",
        )

        assert dest_id == "DTM-04P-2"
        original = run(wb, "modLibrary.ReadConnector", ws_dest_conn, 2, 100000, "DTM-04P")
        assert original[1] == "A different local part"
        imported = run(wb, "modLibrary.ReadConnector", ws_dest_conn, 2, 100000, "DTM-04P-2")
        assert imported[1] == "Deutsch DTM 4-way"
    finally:
        src_wb.Close(SaveChanges=False)


def test_import_pins_are_rewritten_under_the_renamed_id(wb, library_wb, app, tmp_path):
    ws_dest_conn = library_wb.Worksheets("Connectors")
    existing = ("DTM-04P", "A different local part", "", "", "Connector", 1, "", "", "", "", "Local")
    run(wb, "modLibrary.WriteConnector", ws_dest_conn, 2, 100000, existing)

    src_wb, ws_src_conn, ws_src_pins, ws_src_photos = seed_source_connector(wb, app, tmp_path)
    try:
        dest_id = run(
            wb, "modLibraryTransfer.ImportConnector", ws_src_conn, ws_src_pins, ws_src_photos,
            ws_dest_conn, library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"),
            "DTM-04P", "other_library.xlsx",
        )

        pins = run(wb, "modLibrary.ReadPinsForConnector", library_wb.Worksheets("Pins"), 2, 100000, dest_id)
        assert len(pins) == 1
        assert pins[0][0] == dest_id
    finally:
        src_wb.Close(SaveChanges=False)


def test_import_unknown_source_id_returns_empty_string(wb, library_wb, app):
    src_wb = app.Workbooks.Add()
    try:
        library_layout.build_library_sheets(src_wb)
        result = run(
            wb, "modLibraryTransfer.ImportConnector",
            src_wb.Worksheets("Connectors"), src_wb.Worksheets("Pins"), src_wb.Worksheets("Photos"),
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"),
            "NO-SUCH-ID", "other.xlsx",
        )
        assert result == ""
    finally:
        src_wb.Close(SaveChanges=False)
