import shutil

from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_full_connector_definition_round_trips_through_the_saved_file(
    wb, app, library_artifact, tmp_path
):
    photo_path = write_sample_photo(tmp_path / "photo.png")

    # A private copy: this test's Save() must not mutate the session-scoped
    # library_artifact that every other test's library_wb fixture reopens.
    working_copy = tmp_path / "ConnectorLibrary.xlsx"
    shutil.copyfile(library_artifact, working_copy)

    lib = app.Workbooks.Open(str(working_copy))
    try:
        ws_conn = lib.Worksheets("Connectors")
        ws_pins = lib.Worksheets("Pins")
        ws_photos = lib.Worksheets("Photos")

        connector_id = run(wb, "modLibrary.SlugifyConnectorID", "DTM06-4S", "")
        assert connector_id == "DTM06-4S"

        fields = (
            connector_id, "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
            4, "", "PHOTO_" + connector_id, "2026-08-26T00:00:00Z",
            "2026-08-26T00:00:00Z", "Local",
        )
        assert run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields) is True

        for pin_number, label, x, y in [
            (1, "+12V", 0.1, 0.1), (2, "GND", 0.9, 0.1),
            (3, "SIG", 0.1, 0.9), (4, "", 0.9, 0.9),
        ]:
            pin_fields = (connector_id, pin_number, label, x, y, x, y)
            assert run(wb, "modLibrary.WritePin", ws_pins, 2, 100000, pin_fields) is True

        shape_name = run(wb, "modLibrary.EmbedConnectorPhoto", ws_photos, connector_id, str(photo_path))
        assert shape_name == "PHOTO_" + connector_id

        lib.Save()
    finally:
        lib.Close(SaveChanges=False)

    # Reopen as a wholly separate file handle - proves the data actually
    # persisted to disk, not just to the in-memory COM object.
    reopened = app.Workbooks.Open(str(working_copy))
    try:
        ws_conn = reopened.Worksheets("Connectors")
        ws_pins = reopened.Worksheets("Pins")
        ws_photos = reopened.Worksheets("Photos")

        result = run(wb, "modLibrary.ReadConnector", ws_conn, 2, 100000, connector_id)
        # Excel's COM Value property reads numbers back as float (4 -> 4.0)
        # and an empty string written to a cell reads back as None (Excel
        # treats "" as clearing the cell to blank) - see the equivalent note
        # in test_library_connectors.py.
        expected = fields[:5] + (4.0, None) + fields[7:]
        assert tuple(result) == expected

        pins = run(wb, "modLibrary.ReadPinsForConnector", ws_pins, 2, 100000, connector_id)
        assert [int(row[1]) for row in pins] == [1, 2, 3, 4]

        assert ws_photos.Shapes("PHOTO_" + connector_id).Name == "PHOTO_" + connector_id
    finally:
        reopened.Close(SaveChanges=False)
