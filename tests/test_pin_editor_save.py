from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_save_writes_the_connector_record(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws_conn = library_wb.Worksheets("Connectors")
    ws_pins = library_wb.Worksheets("Pins")
    ws_photos = library_wb.Worksheets("Photos")
    ws_scratch = wb.Worksheets("_Edit")

    ok = run(
        wb, "modPinEditor.SaveConnector", ws_conn, ws_pins, ws_photos, ws_scratch,
        "DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector", 4,
        "", str(photo_path), "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )
    assert ok is True

    result = run(wb, "modLibrary.ReadConnector", ws_conn, 2, 100000, "DTM-04P")
    assert result[1] == "Deutsch DTM 4-way"
    assert result[7] == "PHOTO_DTM-04P"
    assert ws_photos.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"


def test_save_writes_the_scratch_pins_into_the_library(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws_conn = library_wb.Worksheets("Connectors")
    ws_pins = library_wb.Worksheets("Pins")
    ws_photos = library_wb.Worksheets("Photos")
    ws_scratch = wb.Worksheets("_Edit")

    run(wb, "modPinEditor.PlacePin", ws_scratch, "DTM-04P", 1, "+12V", 0.1, 0.1)
    run(wb, "modPinEditor.PlacePin", ws_scratch, "DTM-04P", 2, "GND", 0.9, 0.1)

    run(
        wb, "modPinEditor.SaveConnector", ws_conn, ws_pins, ws_photos, ws_scratch,
        "DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector", 2,
        "", str(photo_path), "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )

    pins = run(wb, "modLibrary.ReadPinsForConnector", ws_pins, 2, 100000, "DTM-04P")
    assert [row[1] for row in pins] == [1, 2]


def test_save_overwriting_a_connector_replaces_its_old_pins(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws_conn = library_wb.Worksheets("Connectors")
    ws_pins = library_wb.Worksheets("Pins")
    ws_photos = library_wb.Worksheets("Photos")
    ws_scratch = wb.Worksheets("_Edit")

    run(wb, "modPinEditor.PlacePin", ws_scratch, "DTM-04P", 1, "A", 0.1, 0.1)
    run(wb, "modPinEditor.PlacePin", ws_scratch, "DTM-04P", 2, "B", 0.2, 0.2)
    run(
        wb, "modPinEditor.SaveConnector", ws_conn, ws_pins, ws_photos, ws_scratch,
        "DTM-04P", "First", "", "", "Connector", 2, "", str(photo_path),
        "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )

    run(wb, "modPinEditor.ClearScratchPins", ws_scratch)
    run(wb, "modPinEditor.PlacePin", ws_scratch, "DTM-04P", 1, "OnlyOne", 0.5, 0.5)
    run(
        wb, "modPinEditor.SaveConnector", ws_conn, ws_pins, ws_photos, ws_scratch,
        "DTM-04P", "Second", "", "", "Connector", 1, "", str(photo_path),
        "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )

    pins = run(wb, "modLibrary.ReadPinsForConnector", ws_pins, 2, 100000, "DTM-04P")
    assert len(pins) == 1
    assert pins[0][2] == "OnlyOne"
