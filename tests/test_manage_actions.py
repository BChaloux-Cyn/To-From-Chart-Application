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


def test_import_copies_every_connector_and_reports_photo_status(wb, app, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path, "DTM-04P")
    export_path = tmp_path / "export.xlsx"
    dest = app.Workbooks.Add()
    try:
        run_action(wb, "modManageActions.ExportToWorkbook",
                   library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
                   library_wb.Worksheets("Photos"), dest, "DTM-04P")
        dest.SaveAs(Filename=str(export_path), FileFormat=51)
    finally:
        dest.Close(SaveChanges=False)

    src = app.Workbooks.Open(str(export_path))
    try:
        # Import into a library that does not yet hold this connector.
        run(wb, "modLibrary.DeleteConnector", library_wb.Worksheets("Connectors"),
            2, 100000, "DTM-04P")

        result = run_action(
            wb, "modManageActions.ImportAllFromWorkbook", src,
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"),
        )
        assert (result.ok, result.outcome) == (True, "IMPORTED")
        assert [row[0] for row in result.payload] == ["DTM-04P"]
        assert isinstance(result.payload[0][1], bool)
    finally:
        src.Close(SaveChanges=False)


def test_import_of_an_empty_workbook_reports_an_empty_table(wb, app, library_wb, tmp_path):
    empty_path = tmp_path / "empty.xlsx"
    dest = app.Workbooks.Add()
    try:
        run(wb, "modLibraryTransfer.BuildExportSheets", dest)
        dest.SaveAs(Filename=str(empty_path), FileFormat=51)
    finally:
        dest.Close(SaveChanges=False)

    src = app.Workbooks.Open(str(empty_path))
    try:
        result = run_action(
            wb, "modManageActions.ImportAllFromWorkbook", src,
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"),
        )
        assert result.outcome == "IMPORTED"
        assert run(wb, "modContract.TableRowCount", result.payload) == 0
    finally:
        src.Close(SaveChanges=False)


def test_attach_replacement_photo_embeds_the_chosen_file(wb, library_wb, tmp_path):
    photo = write_sample_photo(tmp_path / "replacement.png")
    result = run_action(wb, "modManageActions.AttachReplacementPhoto",
                        library_wb.Worksheets("Photos"), "DTM-04P", str(photo))
    assert (result.ok, result.outcome, result.payload) == (True, "PHOTO_ATTACHED", "DTM-04P")


def test_attach_replacement_photo_reports_a_missing_file(wb, library_wb, tmp_path):
    result = run_action(wb, "modManageActions.AttachReplacementPhoto",
                        library_wb.Worksheets("Photos"), "DTM-04P",
                        str(tmp_path / "does-not-exist.png"))
    assert (result.ok, result.outcome, result.payload) == (False, "PHOTO_FAILED", "DTM-04P")
