import shutil
from pathlib import Path

from tests.conftest import run, run_action
from tests.fixtures.sample_photo import write_sample_photo


def test_full_harness_round_trips_through_a_saved_file(wb, app, artifact, tmp_path):
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

    # A second wired row, on a third (shared) pin, that intentionally leaves
    # Color and To Pin blank - proves Finding 1: a matched-but-blank field
    # renders blank rather than 0 (LookupFormula) or a partial dash
    # (WireToFormula), without needing a whole new connector.
    wsHarness.Cells(8, 1).Value = "J1"
    wsHarness.Cells(8, 2).Value = 3
    wsHarness.Cells(8, 3).Value = "Crimp Pin"
    wsHarness.Cells(8, 4).Value = "SIG3"
    wsHarness.Cells(8, 5).Value = ""            # Color intentionally blank
    wsHarness.Cells(8, 6).Value = "22"
    wsHarness.Cells(8, 7).Value = 30
    wsHarness.Cells(8, 8).Value = "Ring Terminal"
    wsHarness.Cells(8, 9).Value = "J2"
    wsHarness.Cells(8, 10).Value = ""           # To Pin intentionally blank

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
    run(wb, "modLibrary.WritePin", wsSnap, 211, 2210, ("DTM-04P", 3, "SIG3", 0.5, 0.5, 0.5, 0.5))

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
        assert j1.Cells(3, 13).Value == ""          # pin 2, unwired: LookupFormula-driven column also blank

        j2 = reopened.Worksheets("CONN_J2")
        assert j2.Cells(2, 13).Value == "12V_SW"    # same wire, matched as To this time
        assert j2.Cells(2, 12).Value == "J1-1"      # Wire To points back at J1
        assert j2.Cells(2, 15).Value == 18           # AWG, matched as To
        assert j2.Cells(2, 17).Value == 24           # Length, matched as To
        assert j1.Cells(2, 16).Value == "Crimp Pin"       # From side sees From Term
        assert j2.Cells(2, 16).Value == "Ring Terminal"   # To side sees To Term

        # Finding 1: a matched-but-blank field renders blank, not 0 or a
        # partial dash. Pin 3 is wired (J1 pin 3 -> J2 pin 3) with Color and
        # To Pin intentionally left blank on the chart row.
        assert j1.Cells(4, 14).Value == ""   # Color: LookupFormula-driven, matched but blank -> not 0
        assert j1.Cells(4, 12).Value == ""   # Wire To: To Pin blank -> not a partial "J2-" dash

        assert reopened.HasVBProject is False  # macro-free

        wsHarnessReopened = reopened.Worksheets("Harness")
        wsHarnessReopened.Cells(7, 5).Value = "Blue"  # hand-edit Color on the open, macro-free file
        assert j1.Cells(2, 14).Value == "Blue"
        assert j2.Cells(2, 14).Value == "Blue"

        harness_ps = ws.PageSetup
        assert harness_ps.Orientation == 2  # xlLandscape
        assert harness_ps.PrintTitleRows == "$6:$6"
        assert "HN-100" in harness_ps.CenterFooter

        page_ps = j1.PageSetup
        assert page_ps.Orientation == 1  # xlPortrait
        assert page_ps.PrintTitleRows == ""
        assert "HN-100" in page_ps.CenterFooter
    finally:
        reopened.Close(SaveChanges=False)

    # Load half: a genuinely independent Creator instance proves the save/load
    # pair reconstructs state, rather than passing on in-memory state the save
    # half already had. Opening `artifact` a second time would NOT give one -
    # Excel returns the same Workbook object for an already-open path
    # (Workbooks.Count stays 1 and edits are shared through both handles), so
    # the artifact is copied to a fresh path first. Relocating it is safe
    # because nothing LoadHarness calls reads ThisWorkbook.Path.
    creator2_path = tmp_path / "HarnessCreator2.xlsm"
    shutil.copyfile(artifact, creator2_path)

    wb2 = app.Workbooks.Open(str(creator2_path))
    try:
        srcWb = app.Workbooks.Open(str(saved_path))
        try:
            load_result = run_action(wb2, "modHarnessActions.LoadHarness", srcWb)
            assert load_result.ok is True
            assert load_result.outcome == "HARNESS_LOADED"
        finally:
            srcWb.Close(SaveChanges=False)

        loaded_harness = wb2.Worksheets("Harness")
        assert loaded_harness.Range("B2").Value == "Dash Harness"
        assert loaded_harness.Cells(7, 1).Value == "J1"
        assert loaded_harness.Cells(7, 4).Value == "12V_SW"

        loaded_conn = wb2.Worksheets("Connectors")
        assert loaded_conn.Cells(2, 1).Value == "J1"
        assert loaded_conn.Cells(2, 2).Value == "DTM-04P"

        # fields_round_tripped, not fields: the snapshot reaches this point
        # through the same Excel value coercion as the save half asserted
        # above (4 -> 4.0, "" -> None).
        loaded_snap = wb2.Worksheets("_Snapshot")
        loaded_fields = run(wb2, "modLibrary.ReadConnector", loaded_snap, 2, 201, "DTM-04P")
        assert tuple(loaded_fields) == fields_round_tripped

        assert run(wb2, "modState.GetState", "HarnessPath") == str(saved_path)
        assert run(wb2, "modState.GetState", "Dirty") == "FALSE"
    finally:
        wb2.Close(SaveChanges=False)
