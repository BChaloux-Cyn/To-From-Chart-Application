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


def test_delete_cascades_to_placed_instances_of_that_connector(wb, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)  # J1
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)  # J2

    result = run_action(
        wb, "modManageActions.DeleteFromLibrary",
        library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
        library_wb.Worksheets("Photos"), str(tmp_path), "DTM-04P",
    )
    assert result.outcome == "CONNECTOR_DELETED_CASCADED"
    assert list(result.payload) == ["J1", "J2"]
    assert run(wb, "modConnectors.PinCountFor", "J1") == 0
    assert run(wb, "modConnectors.PinCountFor", "J2") == 0


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


def export_dtm_04p(wb, app, library_wb, tmp_path, filename="export.xlsx"):
    """Exports the library's DTM-04P to its own workbook and reopens it,
    the way an import source file always looks."""
    export_path = tmp_path / filename
    dest = app.Workbooks.Add()
    try:
        run_action(wb, "modManageActions.ExportToWorkbook",
                   library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
                   library_wb.Worksheets("Photos"), dest, "DTM-04P")
        dest.SaveAs(Filename=str(export_path), FileFormat=51)
    finally:
        dest.Close(SaveChanges=False)
    return app.Workbooks.Open(str(export_path))


def test_import_one_connector_inserts_fresh_under_its_own_id(wb, app, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path, "DTM-04P")
    src = export_dtm_04p(wb, app, library_wb, tmp_path)
    try:
        # Import into a library that does not yet hold this connector.
        run(wb, "modLibrary.DeleteConnector", library_wb.Worksheets("Connectors"),
            2, 100000, "DTM-04P")

        result = run_action(
            wb, "modManageActions.ImportOneConnector", src,
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"), "DTM-04P", src.Name,
        )
        assert (result.ok, result.outcome, result.payload) == (True, "CONNECTOR_IMPORTED", "DTM-04P")

        row = run(wb, "modLibrary.ReadConnector", library_wb.Worksheets("Connectors"), 2, 100000, "DTM-04P")
        assert row[1] == "Deutsch DTM 4-way"
    finally:
        src.Close(SaveChanges=False)


def test_import_one_connector_overwrites_without_cascading_when_pin_count_is_unchanged(wb, app, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path, "DTM-04P")
    src = export_dtm_04p(wb, app, library_wb, tmp_path)
    try:
        run(wb, "modConnectors.AddConnectorInstance",
            "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 2)  # J1

        result = run_action(
            wb, "modManageActions.ImportOneConnector", src,
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"), "DTM-04P", src.Name,
        )
        assert (result.ok, result.outcome, result.payload) == (True, "CONNECTOR_IMPORTED", "DTM-04P")
        assert run(wb, "modConnectors.PinCountFor", "J1") == 2  # untouched
    finally:
        src.Close(SaveChanges=False)


def test_import_one_connector_cascades_when_the_overwrite_changes_pin_count(wb, app, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path, "DTM-04P")  # exported source: PinCount 2
    src = export_dtm_04p(wb, app, library_wb, tmp_path)
    try:
        # The library's own copy has since been edited to a different pin
        # count than what the source file holds.
        changed = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
                   4, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
        run(wb, "modLibrary.WriteConnector", library_wb.Worksheets("Connectors"), 2, 100000, changed)
        run(wb, "modConnectors.AddConnectorInstance",
            "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)  # J1
        run(wb, "modConnectors.AddConnectorInstance",
            "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)  # J2

        result = run_action(
            wb, "modManageActions.ImportOneConnector", src,
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"), "DTM-04P", src.Name,
        )
        assert result.outcome == "CONNECTOR_IMPORTED_CASCADED"
        assert list(result.payload) == ["J1", "J2"]
        assert run(wb, "modConnectors.PinCountFor", "J1") == 0
        assert run(wb, "modConnectors.PinCountFor", "J2") == 0
    finally:
        src.Close(SaveChanges=False)


def test_imported_photo_ok_reflects_whether_the_shape_exists(wb, library_wb, tmp_path):
    assert run(wb, "modManageActions.ImportedPhotoOk",
               library_wb.Worksheets("Photos"), "NO-SUCH-ID") is False

    seed(wb, library_wb, tmp_path, "DTM-04P")
    assert run(wb, "modManageActions.ImportedPhotoOk",
               library_wb.Worksheets("Photos"), "DTM-04P") is True


def test_export_library_copies_every_connector(wb, app, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path, "DTM-04P")
    seed(wb, library_wb, tmp_path, "DTM-08P")
    dest = app.Workbooks.Add()
    try:
        result = run_action(
            wb, "modManageActions.ExportLibraryToWorkbook",
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"), dest,
        )
        assert (result.ok, result.outcome, result.payload) == (True, "LIBRARY_EXPORTED", 2)

        ids = {row[1] for row in run(wb, "modLibrary.ConnectorIndex", dest.Worksheets("Connectors"))}
        assert ids == {"DTM-04P", "DTM-08P"}
    finally:
        dest.Close(SaveChanges=False)


def test_export_library_is_empty_but_succeeds_when_the_library_is_empty(wb, app, library_wb):
    dest = app.Workbooks.Add()
    try:
        result = run_action(
            wb, "modManageActions.ExportLibraryToWorkbook",
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"), dest,
        )
        assert (result.ok, result.outcome, result.payload) == (True, "LIBRARY_EXPORTED", 0)
    finally:
        dest.Close(SaveChanges=False)


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
