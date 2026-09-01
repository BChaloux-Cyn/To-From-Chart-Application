from tests.conftest import run
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
