from tests.conftest import run, run_action
from tests.fixtures.sample_photo import write_sample_photo


def seed(wb, library_wb, tmp_path, connector_id="DTM-04P", pin_count=2):
    photo = write_sample_photo(tmp_path / "photo.png")
    ws_conn = library_wb.Worksheets("Connectors")
    fields = (connector_id, "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
              pin_count, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields)
    for n in range(1, pin_count + 1):
        run(wb, "modLibrary.WritePin", library_wb.Worksheets("Pins"), 2, 100000,
            (connector_id, n, f"Pin {n}", 0.1, 0.1, 0.1, 0.1))
    shape = run(wb, "modLibrary.EmbedConnectorPhoto",
                library_wb.Worksheets("Photos"), connector_id, str(photo))
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000,
        fields[:7] + (shape,) + fields[8:])


def add(wb, library_wb, connector_id):
    return run_action(
        wb, "modPickerActions.AddFromLibrary",
        wb.Worksheets("_Snapshot"), library_wb.Worksheets("Connectors"),
        library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"), connector_id,
    )


def test_add_from_library_creates_an_instance_and_snapshots_it(wb, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)

    result = add(wb, library_wb, "DTM-04P")
    assert result.ok is True
    assert result.outcome == "ADDED"
    assert result.payload == "J1"

    snapshot = run(wb, "modLibrary.ReadConnector", wb.Worksheets("_Snapshot"), 2, 201, "DTM-04P")
    assert snapshot[1] == "Deutsch DTM 4-way"


def test_adding_twice_allocates_sequential_ref_designators(wb, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)
    assert add(wb, library_wb, "DTM-04P").payload == "J1"
    assert add(wb, library_wb, "DTM-04P").payload == "J2"


def test_adding_an_unknown_connector_reports_not_found(wb, library_wb):
    result = add(wb, library_wb, "NOPE")
    assert (result.ok, result.outcome, result.payload) == (False, "CONNECTOR_NOT_FOUND", "NOPE")
