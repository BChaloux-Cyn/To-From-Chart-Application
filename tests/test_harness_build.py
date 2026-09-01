from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_build_harness_sheets_creates_harness_and_snapshot(wb, app):
    dest = app.Workbooks.Add()
    try:
        ok = run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        assert ok is True
        names = [dest.Worksheets(i + 1).Name for i in range(dest.Worksheets.Count)]
        assert names == ["Harness", "_Snapshot"]
    finally:
        dest.Close(SaveChanges=False)


def test_snapshot_sheet_is_very_hidden(wb, app):
    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        assert dest.Worksheets("_Snapshot").Visible == 2  # xlSheetVeryHidden
    finally:
        dest.Close(SaveChanges=False)


def test_build_harness_sheets_rejects_a_non_fresh_workbook(wb, app):
    dest = app.Workbooks.Add()
    try:
        dest.Worksheets.Add()  # now has 2 sheets, no longer "fresh"
        ok = run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        assert ok is False
        assert dest.Worksheets.Count == 2  # untouched
    finally:
        dest.Close(SaveChanges=False)


def _fill_title_block(wb):
    ws = wb.Worksheets("Harness")
    ws.Range("B2").Value = "Test Harness"
    ws.Range("E2").Value = "HN-001"
    ws.Range("H2").Value = "A"
    ws.Range("B3").Value = "A Student"
    ws.Range("E3").Value = "Shop 1"
    ws.Range("H3").Value = "2026-09-01"
    ws.Range("B4").Value = "A test harness"
    ws.Range("H4").Value = "in"


def test_copy_title_block_round_trips_every_field(wb, app):
    _fill_title_block(wb)
    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        run(wb, "modHarnessBuild.CopyTitleBlock", dest.Worksheets("Harness"))
        ws = dest.Worksheets("Harness")
        assert ws.Range("B2").Value == "Test Harness"
        assert ws.Range("E2").Value == "HN-001"
        assert ws.Range("H2").Value == "A"
        assert ws.Range("B3").Value == "A Student"
        assert ws.Range("E3").Value == "Shop 1"
        assert ws.Range("H4").Value == "in"
        assert ws.Range("A1").Value == "WIRE HARNESS TO-FROM CHART"
        assert ws.Cells(6, 1).Value == "From Conn"  # chart header row rendered
    finally:
        dest.Close(SaveChanges=False)


def test_copy_chart_rows_round_trips_values_and_counts_used_rows(wb, app):
    _fill_title_block(wb)
    wsSrc = wb.Worksheets("Harness")
    wsSrc.Cells(7, 1).Value = "J1"
    wsSrc.Cells(7, 2).Value = 1
    wsSrc.Cells(7, 9).Value = "J2"
    wsSrc.Cells(7, 10).Value = 1
    wsSrc.Cells(8, 1).Value = "J1"
    wsSrc.Cells(8, 2).Value = 2
    wsSrc.Cells(8, 9).Value = "J2"
    wsSrc.Cells(8, 10).Value = 2

    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        n = run(wb, "modHarnessBuild.CopyChartRows", dest.Worksheets("Harness"))
        assert n == 2

        ws = dest.Worksheets("Harness")
        assert ws.Cells(7, 1).Value == "J1"
        assert ws.Cells(7, 10).Value == 1
        assert ws.Cells(9, 1).Value is None  # untouched beyond the used rows
    finally:
        dest.Close(SaveChanges=False)


def test_copy_chart_rows_writes_live_join_key_formulas(wb, app):
    _fill_title_block(wb)
    wsSrc = wb.Worksheets("Harness")
    wsSrc.Cells(7, 1).Value = "J1"
    wsSrc.Cells(7, 2).Value = 3
    wsSrc.Cells(7, 9).Value = "J2"
    wsSrc.Cells(7, 10).Value = 4

    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        run(wb, "modHarnessBuild.CopyChartRows", dest.Worksheets("Harness"))
        ws = dest.Worksheets("Harness")
        assert ws.Cells(7, 12).Value == "J1|3"
        assert ws.Cells(7, 13).Value == "J2|4"
        assert ws.Cells(8, 12).Value == ""  # blank row: formula present, resolves empty
        assert ws.Columns(12).Hidden is True
        assert ws.Columns(13).Hidden is True

        # A hand edit to the chart after saving keeps the join key correct -
        # this is what makes 3c's pin tables react to it with no macro.
        ws.Cells(7, 2).Value = 9
        assert ws.Cells(7, 12).Value == "J1|9"
    finally:
        dest.Close(SaveChanges=False)


def test_copy_snapshot_round_trips_connectors_pins_and_photos(wb, app, tmp_path):
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
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        run(wb, "modHarnessBuild.CopySnapshot", dest.Worksheets("_Snapshot"))

        wsDest = dest.Worksheets("_Snapshot")
        result = run(wb, "modLibrary.ReadConnector", wsDest, 2, 201, "DTM-04P")
        assert tuple(result) == fields_round_tripped

        pins = run(wb, "modLibrary.ReadPinsForConnector", wsDest, 211, 2210, "DTM-04P")
        assert [int(row[1]) for row in pins] == [1]

        assert wsDest.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"
    finally:
        dest.Close(SaveChanges=False)
