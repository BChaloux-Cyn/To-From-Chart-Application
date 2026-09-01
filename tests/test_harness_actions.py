from tests.conftest import run_action


def test_save_harness_succeeds_against_a_fresh_workbook(wb, app):
    wsSrc = wb.Worksheets("Harness")
    wsSrc.Range("B2").Value = "Test Harness"
    wsSrc.Cells(7, 1).Value = "J1"
    wsSrc.Cells(7, 2).Value = 1
    wsSrc.Cells(7, 9).Value = "J2"
    wsSrc.Cells(7, 10).Value = 1

    dest = app.Workbooks.Add()
    try:
        result = run_action(wb, "modHarnessActions.SaveHarness", dest)
        assert result.ok is True
        assert result.outcome == "HARNESS_SAVED"
        assert result.payload == 1
        assert dest.Worksheets("Harness").Range("B2").Value == "Test Harness"
    finally:
        dest.Close(SaveChanges=False)


def test_save_harness_rejects_a_non_fresh_workbook(wb, app):
    dest = app.Workbooks.Add()
    try:
        dest.Worksheets.Add()
        result = run_action(wb, "modHarnessActions.SaveHarness", dest)
        assert result.ok is False
        assert result.outcome == "HARNESS_SAVE_FAILED"
    finally:
        dest.Close(SaveChanges=False)
