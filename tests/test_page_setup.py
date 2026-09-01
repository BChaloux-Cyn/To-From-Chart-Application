from tests.conftest import run

XL_LANDSCAPE = 2
XL_PORTRAIT = 1


def test_last_used_chart_row_finds_the_last_populated_row(wb, app):
    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        ws = dest.Worksheets("Harness")
        ws.Cells(7, 1).Value = "J1"
        ws.Cells(10, 9).Value = "J2"  # gap between row 7 and row 10 is realistic

        n = run(wb, "modPageSetup.LastUsedChartRow", ws)
        assert n == 10
    finally:
        dest.Close(SaveChanges=False)


def test_last_used_chart_row_falls_back_to_the_header_row_when_empty(wb, app):
    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        ws = dest.Worksheets("Harness")
        n = run(wb, "modPageSetup.LastUsedChartRow", ws)
        assert n == 6
    finally:
        dest.Close(SaveChanges=False)


def test_apply_harness_page_setup(wb, app):
    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        ws = dest.Worksheets("Harness")
        ws.Cells(7, 1).Value = "J1"

        run(wb, "modPageSetup.ApplyHarnessPageSetup", ws, "HN-100", "A")

        ps = ws.PageSetup
        assert ps.Orientation == XL_LANDSCAPE
        assert ps.FitToPagesWide == 1
        assert ps.PrintArea == "$A$1:$K$7"
        assert ps.PrintTitleRows == "$6:$6"
        assert "HN-100" in ps.CenterFooter
        assert "A" in ps.CenterFooter
        assert "&P" in ps.CenterFooter and "&N" in ps.CenterFooter
    finally:
        dest.Close(SaveChanges=False)


def test_apply_connector_page_setup(wb, app):
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        pins = (("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1),)
        run(wb, "modConnectorPage.WriteTableSkeleton", ws, pins)

        run(wb, "modPageSetup.ApplyConnectorPageSetup", ws, "HN-100", "A")

        ps = ws.PageSetup
        assert ps.Orientation == XL_PORTRAIT
        assert ps.FitToPagesWide == 1
        assert ps.FitToPagesTall == 1
        assert ps.PrintArea == "$A$1:$Q$30"
        assert ps.PrintTitleRows == ""
        assert "HN-100" in ps.CenterFooter
    finally:
        dest.Close(SaveChanges=False)


def test_apply_connector_page_setup_grows_past_the_minimum_for_a_large_pin_count(wb, app):
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        pins = tuple(("DTM-64P", n, "", 0.1, 0.1, 0.1, 0.1) for n in range(1, 41))
        run(wb, "modConnectorPage.WriteTableSkeleton", ws, pins)

        run(wb, "modPageSetup.ApplyConnectorPageSetup", ws, "HN-100", "A")

        assert ws.PageSetup.PrintArea == "$A$1:$Q$41"  # header + 40 pins
    finally:
        dest.Close(SaveChanges=False)
