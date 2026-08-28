from pathlib import Path

from tests.conftest import run, run_action
from tests.fixtures.sample_photo import write_sample_photo


def seed(wb, library_wb, tmp_path, connector_id="DTM-04P"):
    photo = write_sample_photo(tmp_path / f"{connector_id}.png")
    ws_conn = library_wb.Worksheets("Connectors")
    fields = (connector_id, "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
              2, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields)
    run(wb, "modLibrary.WritePin", library_wb.Worksheets("Pins"), 2, 100000,
        (connector_id, 1, "Pin 1", 0.1, 0.1, 0.1, 0.1))
    shape = run(wb, "modLibrary.EmbedConnectorPhoto",
                library_wb.Worksheets("Photos"), connector_id, str(photo))
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000,
        fields[:7] + (shape,) + fields[8:])


def test_delete_removes_the_row_the_pins_and_the_cache_file(wb, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)
    cache = run(wb, "modLibrary.CachePhotoPath", str(tmp_path), "DTM-04P", "jpg")
    Path(cache).write_bytes(b"cached preview")

    result = run_action(
        wb, "modManageActions.DeleteFromLibrary",
        library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
        library_wb.Worksheets("Photos"), str(tmp_path), "DTM-04P",
    )
    assert (result.ok, result.outcome, result.payload) == (True, "CONNECTOR_DELETED", "DTM-04P")

    assert run(wb, "modLibrary.ReadConnector",
               library_wb.Worksheets("Connectors"), 2, 100000, "DTM-04P") is None
    assert run(wb, "modLibrary.ReadPinsForConnector",
               library_wb.Worksheets("Pins"), 2, 100000, "DTM-04P") is None
    assert not Path(cache).exists()


def test_delete_succeeds_when_no_cache_file_was_ever_written(wb, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)
    result = run_action(
        wb, "modManageActions.DeleteFromLibrary",
        library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
        library_wb.Worksheets("Photos"), str(tmp_path), "DTM-04P",
    )
    assert result.outcome == "CONNECTOR_DELETED"


def test_export_builds_the_sheets_and_copies_the_record(wb, app, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)
    dest = app.Workbooks.Add()
    try:
        result = run_action(
            wb, "modManageActions.ExportToWorkbook",
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"), dest, "DTM-04P",
        )
        assert (result.ok, result.outcome, result.payload) == (True, "EXPORTED", "DTM-04P")

        row = run(wb, "modLibrary.ReadConnector", dest.Worksheets("Connectors"), 2, 100000, "DTM-04P")
        assert row[1] == "Deutsch DTM 4-way"
    finally:
        dest.Close(SaveChanges=False)


def test_exporting_an_unknown_connector_fails(wb, app, library_wb):
    dest = app.Workbooks.Add()
    try:
        result = run_action(
            wb, "modManageActions.ExportToWorkbook",
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"), dest, "NOPE",
        )
        assert (result.ok, result.outcome, result.payload) == (False, "EXPORT_FAILED", "NOPE")
    finally:
        dest.Close(SaveChanges=False)
