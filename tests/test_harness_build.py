from tests.conftest import run


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
