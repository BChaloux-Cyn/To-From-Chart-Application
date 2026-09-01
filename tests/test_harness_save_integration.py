from pathlib import Path

from tests.conftest import run, run_action
from tests.fixtures.sample_photo import write_sample_photo


def test_full_harness_round_trips_through_a_saved_file(wb, app, tmp_path):
    wsHarness = wb.Worksheets("Harness")
    wsHarness.Range("B2").Value = "Dash Harness"
    wsHarness.Range("E2").Value = "HN-100"
    wsHarness.Range("H2").Value = "A"
    wsHarness.Cells(7, 1).Value = "J1"
    wsHarness.Cells(7, 2).Value = 1
    wsHarness.Cells(7, 3).Value = "Crimp Pin"      # From Term
    wsHarness.Cells(7, 4).Value = "12V_SW"
    wsHarness.Cells(7, 5).Value = "Red"
    wsHarness.Cells(7, 6).Value = "18"
    wsHarness.Cells(7, 7).Value = 24
    wsHarness.Cells(7, 8).Value = "Ring Terminal"  # To Term
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

    wsConn = wb.Worksheets("Connectors")
    wsConn.Cells(2, 1).Value = "J1"
    wsConn.Cells(2, 2).Value = "DTM-04P"
    wsConn.Cells(2, 3).Value = "Deutsch DTM 4-way"
    wsConn.Cells(2, 5).Value = "Connector"
    wsConn.Cells(2, 6).Value = 4

    wsConn.Cells(3, 1).Value = "J2"
    wsConn.Cells(3, 2).Value = "DTM-04P"
    wsConn.Cells(3, 3).Value = "Deutsch DTM 4-way"
    wsConn.Cells(3, 5).Value = "Connector"
    wsConn.Cells(3, 6).Value = 4

    run(wb, "modLibrary.WritePin", wsSnap, 211, 2210, ("DTM-04P", 2, "GND", 0.9, 0.1, 0.9, 0.1))

    # BuildConnectorPages looks in LibraryFolder() (ThisWorkbook.Path, not
    # tmp_path) for the photo cache - a real path alongside the built
    # artifact, not test-isolated - see test_snapshot.py's identical seeding.
    library_folder = run(wb, "modSnapshot.LibraryFolder")
    cache_path = Path(library_folder) / "Photos" / "DTM-04P.png"
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    write_sample_photo(cache_path)

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
        cache_path.unlink(missing_ok=True)

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

        page = reopened.Worksheets("CONN_J1")
        assert page.Cells(1, 27).Value == "DTM-04P"  # metadata cell, for 3e
        assert page.Columns(27).Hidden is True
        assert page.Shapes("PIN_1").Name == "PIN_1"
        assert page.Cells(1, 10).Value == "Pin"
        assert page.Cells(2, 10).Value == 1
        assert page.Cells(2, 11).Value == "+12V"

        j1 = reopened.Worksheets("CONN_J1")
        assert j1.Cells(2, 13).Value == "12V_SW"   # Signal, matched as From
        assert j1.Cells(2, 14).Value == "Red"       # Color
        assert j1.Cells(2, 15).Value == 18           # AWG (Excel coerces the numeric-looking string to a number)
        assert j1.Cells(2, 17).Value == 24           # Length
        assert j1.Cells(2, 12).Value == "J2-1"      # Wire To, matched as From
        assert j1.Cells(3, 12).Value == ""          # pin 2, unwired: blank, not an error

        j2 = reopened.Worksheets("CONN_J2")
        assert j2.Cells(2, 13).Value == "12V_SW"    # same wire, matched as To this time
        assert j2.Cells(2, 12).Value == "J1-1"      # Wire To points back at J1
        assert j1.Cells(2, 16).Value == "Crimp Pin"       # From side sees From Term
        assert j2.Cells(2, 16).Value == "Ring Terminal"   # To side sees To Term

        assert reopened.HasVBProject is False  # macro-free
    finally:
        reopened.Close(SaveChanges=False)
