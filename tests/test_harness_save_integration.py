from tests.conftest import run, run_action
from tests.fixtures.sample_photo import write_sample_photo


def test_full_harness_round_trips_through_a_saved_file(wb, app, tmp_path):
    wsHarness = wb.Worksheets("Harness")
    wsHarness.Range("B2").Value = "Dash Harness"
    wsHarness.Range("E2").Value = "HN-100"
    wsHarness.Range("H2").Value = "A"
    wsHarness.Cells(7, 1).Value = "J1"
    wsHarness.Cells(7, 2).Value = 1
    wsHarness.Cells(7, 4).Value = "12V_SW"
    wsHarness.Cells(7, 5).Value = "Red"
    wsHarness.Cells(7, 6).Value = "18"
    wsHarness.Cells(7, 7).Value = 24
    wsHarness.Cells(7, 9).Value = "J2"
    wsHarness.Cells(7, 10).Value = 1

    wsSnap = wb.Worksheets("_Snapshot")
    fields = (
        "DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
        4, "", "PHOTO_DTM-04P", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )
    run(wb, "modLibrary.WriteConnector", wsSnap, 2, 201, fields)
    run(wb, "modLibrary.WritePin", wsSnap, 211, 2210, ("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1))
    photo_path = write_sample_photo(tmp_path / "photo.png")
    run(wb, "modLibrary.EmbedConnectorPhoto", wsSnap, "DTM-04P", str(photo_path))

    # Excel's COM Value property always reads numbers back as float (4 -> 4.0)
    # and an empty string written to a cell reads back as None (Excel treats
    # "" as clearing the cell to blank) - see test_library_connectors.py.
    fields_round_tripped = fields[:5] + (4.0, None) + fields[7:]

    dest = app.Workbooks.Add()
    saved_path = tmp_path / "HN-100.xlsx"
    try:
        result = run_action(wb, "modHarnessActions.SaveHarness", dest)
        assert result.ok is True
        dest.SaveAs(Filename=str(saved_path), FileFormat=51)
    finally:
        dest.Close(SaveChanges=False)

    # Reopen as a wholly separate file handle.
    reopened = app.Workbooks.Open(str(saved_path))
    try:
        ws = reopened.Worksheets("Harness")
        assert ws.Range("B2").Value == "Dash Harness"
        assert ws.Cells(7, 1).Value == "J1"
        assert ws.Cells(7, 4).Value == "12V_SW"
        assert ws.Cells(7, 12).Value == "J1|1"
        assert ws.Cells(7, 13).Value == "J2|1"

        wsSnapDest = reopened.Worksheets("_Snapshot")
        assert wsSnapDest.Visible == 2
        result2 = run(wb, "modLibrary.ReadConnector", wsSnapDest, 2, 201, "DTM-04P")
        assert tuple(result2) == fields_round_tripped
        assert wsSnapDest.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"

        assert reopened.HasVBProject is False  # macro-free
    finally:
        reopened.Close(SaveChanges=False)
