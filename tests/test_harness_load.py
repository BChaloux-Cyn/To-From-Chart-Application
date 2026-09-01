from tests.conftest import run, run_action
from tests.fixtures.sample_photo import write_sample_photo


def test_copy_snapshot_into_replaces_prior_contents(wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    srcWb = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", srcWb)
        wsSrcSnap = srcWb.Worksheets("_Snapshot")
        fields = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
                  4, "", "PHOTO_DTM-04P", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
        run(wb, "modLibrary.WriteConnector", wsSrcSnap, 2, 201, fields)
        run(wb, "modLibrary.WritePin", wsSrcSnap, 211, 2210, ("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1))
        run(wb, "modLibrary.EmbedConnectorPhoto", wsSrcSnap, "DTM-04P", str(photo_path))

        wsDestSnap = wb.Worksheets("_Snapshot")
        # Stale data from a previous session, which the copy must overwrite/replace.
        run(wb, "modLibrary.WriteConnector", wsDestSnap, 2, 201,
            ("OLD-ID", "Old Connector", "", "", "Connector", 2, "", "", "", "", "Local"))
        run(wb, "modLibrary.EmbedConnectorPhoto", wsDestSnap, "OLD-ID", str(photo_path))

        run(wb, "modHarnessLoad.CopySnapshotInto", wsSrcSnap, wsDestSnap)

        # Excel's COM Value property always reads numbers back as float (4 -> 4.0)
        # and an empty string written to a cell reads back as None (Excel treats
        # "" as clearing the cell to blank) - see test_harness_build.py.
        fields_round_tripped = fields[:5] + (4.0, None) + fields[7:]

        result = run(wb, "modLibrary.ReadConnector", wsDestSnap, 2, 201, "DTM-04P")
        assert tuple(result) == fields_round_tripped
        assert wsDestSnap.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"
        assert wsDestSnap.Shapes.Count == 1  # the stale OLD-ID photo is gone
    finally:
        srcWb.Close(SaveChanges=False)


def test_copy_title_block_values_updates_length_units_header(wb, app):
    srcWb = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", srcWb)
        wsSrc = srcWb.Worksheets("Harness")
        wsSrc.Range("B2").Value = "Loaded Harness"
        wsSrc.Range("E2").Value = "HN-200"
        wsSrc.Range("H4").Value = "mm"

        wsDest = wb.Worksheets("Harness")
        run(wb, "modHarnessLoad.CopyTitleBlockValues", wsSrc, wsDest)

        assert wsDest.Range("B2").Value == "Loaded Harness"
        assert wsDest.Range("E2").Value == "HN-200"
        assert wsDest.Cells(6, 7).Value == "Length (mm)"
    finally:
        srcWb.Close(SaveChanges=False)


def test_copy_chart_values_round_trips_and_counts_used_rows(wb, app):
    srcWb = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", srcWb)
        wsSrc = srcWb.Worksheets("Harness")
        wsSrc.Cells(7, 1).Value = "J1"
        wsSrc.Cells(7, 9).Value = "J2"
        wsSrc.Cells(10, 1).Value = "J1"
        wsSrc.Cells(10, 9).Value = "J2"

        wsDest = wb.Worksheets("Harness")
        n = run(wb, "modHarnessLoad.CopyChartValues", wsSrc, wsDest)

        assert n == 2
        assert wsDest.Cells(7, 1).Value == "J1"
        assert wsDest.Cells(10, 9).Value == "J2"
    finally:
        srcWb.Close(SaveChanges=False)


def test_rebuild_connector_instances_reads_ref_des_from_sheet_names(wb, app, tmp_path):
    srcWb = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", srcWb)
        wsSnap = srcWb.Worksheets("_Snapshot")
        fields = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
                  4, "", "PHOTO_DTM-04P", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
        run(wb, "modLibrary.WriteConnector", wsSnap, 2, 201, fields)

        page = srcWb.Worksheets.Add(After=srcWb.Worksheets(srcWb.Worksheets.Count))
        page.Name = "CONN_J1"
        run(wb, "modConnectorPage.WriteMetadata", page, "DTM-04P")

        wsDestSnap = wb.Worksheets("_Snapshot")
        run(wb, "modHarnessLoad.CopySnapshotInto", wsSnap, wsDestSnap)

        wsDestConn = wb.Worksheets("Connectors")
        n = run(wb, "modHarnessLoad.RebuildConnectorInstances", srcWb, wsDestSnap, wsDestConn)

        assert n == 1
        assert wsDestConn.Cells(2, 1).Value == "J1"
        assert wsDestConn.Cells(2, 2).Value == "DTM-04P"
        assert wsDestConn.Cells(2, 3).Value == "Deutsch DTM 4-way"
        assert int(wsDestConn.Cells(2, 6).Value) == 4
    finally:
        srcWb.Close(SaveChanges=False)


def test_load_harness_rejects_a_file_with_no_harness_sheet(wb, app):
    srcWb = app.Workbooks.Add()
    try:
        result = run_action(wb, "modHarnessActions.LoadHarness", srcWb)
        assert result.ok is False
        assert result.outcome == "HARNESS_LOAD_FAILED"
    finally:
        srcWb.Close(SaveChanges=False)


def test_load_harness_reconstructs_creator_state(wb, app, tmp_path):
    srcWb = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", srcWb)
        wsSrcHarness = srcWb.Worksheets("Harness")
        wsSrcHarness.Range("B2").Value = "Loaded Harness"
        wsSrcHarness.Range("E2").Value = "HN-300"
        wsSrcHarness.Cells(7, 1).Value = "J1"
        wsSrcHarness.Cells(7, 2).Value = 1
        wsSrcHarness.Cells(7, 9).Value = "J2"
        wsSrcHarness.Cells(7, 10).Value = 1

        wsSrcSnap = srcWb.Worksheets("_Snapshot")
        fields = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
                  4, "", "PHOTO_DTM-04P", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
        run(wb, "modLibrary.WriteConnector", wsSrcSnap, 2, 201, fields)

        page = srcWb.Worksheets.Add(After=srcWb.Worksheets(srcWb.Worksheets.Count))
        page.Name = "CONN_J1"
        run(wb, "modConnectorPage.WriteMetadata", page, "DTM-04P")

        result = run_action(wb, "modHarnessActions.LoadHarness", srcWb)
        assert result.ok is True
        assert result.outcome == "HARNESS_LOADED"
        assert result.payload == 1

        wsDestHarness = wb.Worksheets("Harness")
        assert wsDestHarness.Range("B2").Value == "Loaded Harness"
        assert wsDestHarness.Cells(7, 1).Value == "J1"

        wsDestConn = wb.Worksheets("Connectors")
        assert wsDestConn.Cells(2, 1).Value == "J1"
        assert wsDestConn.Cells(2, 2).Value == "DTM-04P"

        # Pin dropdown validation rebuilt against the reconstructed Connectors sheet.
        validation_formula = wsDestHarness.Cells(7, 2).Validation.Formula1
        assert validation_formula == "1,2,3,4"

        assert run(wb, "modState.GetState", "Dirty") == "FALSE"
    finally:
        srcWb.Close(SaveChanges=False)
