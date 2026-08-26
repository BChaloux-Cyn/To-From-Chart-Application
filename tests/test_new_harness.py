import pytest

from tests.conftest import run

TB_NAMES = [
    "TB_Name", "TB_Number", "TB_Rev", "TB_Student",
    "TB_Class", "TB_Date", "TB_Desc",
]


def populate(wb):
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(7, 2).Value = 1
    sheet.Cells(7, 4).Value = "+12V Batt"
    for name in TB_NAMES:
        wb.Names(name).RefersToRange.Value = "seeded"
    wb.Worksheets("Check").Cells(2, 3).Value = "stale finding"


def test_new_harness_clears_the_chart(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    sheet = wb.Worksheets("Harness")
    assert sheet.Cells(7, 1).Value is None
    assert sheet.Cells(7, 4).Value is None


def test_new_harness_clears_the_connector_list(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    assert wb.Worksheets("Connectors").Cells(2, 1).Value is None


def test_new_harness_clears_the_title_block(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    for name in TB_NAMES:
        assert wb.Names(name).RefersToRange.Value is None


def test_new_harness_clears_stale_check_results(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    assert wb.Worksheets("Check").Cells(2, 3).Value is None


def test_new_harness_resets_units_and_path(wb):
    populate(wb)
    wb.Names("TB_Units").RefersToRange.Value = "mm"
    run(wb, "modState.SetState", "HarnessPath", r"C:\temp\x.xlsx")
    run(wb, "modChart.NewHarness")
    assert wb.Worksheets("Harness").Cells(6, 7).Value == "Length (in)"
    assert run(wb, "modState.GetState", "HarnessPath") == ""


def test_new_harness_leaves_the_workbook_clean(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    assert run(wb, "modState.IsDirty") is False


def test_new_harness_drops_stale_pin_validation(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    with pytest.raises(Exception):
        _ = wb.Worksheets("Harness").Cells(7, 2).Validation.Type


def test_home_sheet_has_a_new_harness_button(wb):
    shapes = wb.Worksheets("Home").Shapes
    actions = [shapes(i + 1).OnAction for i in range(shapes.Count)]
    assert "modChart.NewHarness" in actions
