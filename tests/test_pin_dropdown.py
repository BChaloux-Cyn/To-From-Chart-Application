import pytest

from tests.conftest import run


def add_four_way(wb):
    return run(wb, "modConnectors.AddConnectorInstance",
               "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)


def test_typing_a_connector_builds_its_pin_list(wb):
    add_four_way(wb)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    assert sheet.Cells(7, 2).Validation.Formula1 == "1,2,3,4"


def test_to_connector_drives_the_to_pin_column(wb):
    add_four_way(wb)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 9).Value = "J1"
    assert sheet.Cells(7, 10).Validation.Formula1 == "1,2,3,4"


def test_changing_the_connector_clears_a_stale_pin(wb):
    add_four_way(wb)
    run(wb, "modConnectors.AddConnectorInstance",
        "GND-STUD", "Chassis ground stud", "", "Stud", 1)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(7, 2).Value = 4
    sheet.Cells(7, 1).Value = "ST1"
    assert sheet.Cells(7, 2).Value is None
    assert sheet.Cells(7, 2).Validation.Formula1 == "1"


def test_unknown_connector_leaves_no_pin_validation(wb):
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J99"
    with pytest.raises(Exception):
        _ = sheet.Cells(7, 2).Validation.Type


def test_large_connector_falls_back_to_a_numeric_range(wb):
    # 1,2,...,120 exceeds Excel's 255-character Formula1 limit.
    run(wb, "modConnectors.AddConnectorInstance",
        "BIG-120", "120 way", "", "Connector", 120)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    assert sheet.Cells(7, 2).Validation.Type == 1  # xlValidateWholeNumber
    assert sheet.Cells(7, 2).Validation.Formula1 == "1"
    assert sheet.Cells(7, 2).Validation.Formula2 == "120"


def test_each_row_gets_its_own_pin_list(wb):
    add_four_way(wb)
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-12P", "12 way", "", "Connector", 12)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(8, 1).Value = "J2"
    assert sheet.Cells(7, 2).Validation.Formula1 == "1,2,3,4"
    assert sheet.Cells(8, 2).Validation.Formula1 == "1,2,3,4,5,6,7,8,9,10,11,12"


def test_editing_the_chart_marks_the_workbook_dirty(wb):
    run(wb, "modState.ClearDirty")
    wb.Worksheets("Harness").Cells(7, 4).Value = "+12V Batt"
    assert run(wb, "modState.IsDirty") is True


def test_switching_units_rewrites_the_length_header(wb):
    sheet = wb.Worksheets("Harness")
    wb.Names("TB_Units").RefersToRange.Value = "mm"
    assert sheet.Cells(6, 7).Value == "Length (mm)"
    assert run(wb, "modState.GetState", "LengthUnits") == "mm"
