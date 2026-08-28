from tests.conftest import run, run_action

# modChart.bas declarations: the chart occupies rows 7 to 1006, with the
# From-connector column at 1 and the To-connector column at 9.
CHART_FIRST_ROW, CHART_LAST_ROW = 7, 1006
COL_FROM_CONN, COL_TO_CONN = 1, 9


def test_a_bulk_edit_rebuilds_validation_for_the_clamped_row_range(wb):
    sheet = wb.Worksheets("Harness")
    # 3000 cells spanning rows 1 to 1000, so it exceeds the bulk threshold
    # and starts above the chart. Only rows 7 to 1000 should be rebuilt.
    target = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1000, 3))

    result = run_action(wb, "modChart.ApplyHarnessEdit", sheet, target)
    assert result.outcome == "BULK_REBUILT"
    assert result.payload == 1000 - CHART_FIRST_ROW + 1


def test_a_single_connector_cell_edit_rebuilds_that_row(wb):
    sheet = wb.Worksheets("Harness")
    target = sheet.Cells(CHART_FIRST_ROW, COL_FROM_CONN)

    result = run_action(wb, "modChart.ApplyHarnessEdit", sheet, target)
    assert (result.ok, result.outcome, result.payload) == (True, "CELLS_REBUILT", 1)


def test_an_edit_outside_the_connector_columns_rebuilds_nothing(wb):
    sheet = wb.Worksheets("Harness")
    target = sheet.Cells(CHART_FIRST_ROW, COL_TO_CONN - 1)

    result = run_action(wb, "modChart.ApplyHarnessEdit", sheet, target)
    assert (result.ok, result.outcome) == (False, "NO_OP")


def test_a_rename_is_applied_and_reported(wb):
    sheet = wb.Worksheets("Connectors")
    run(wb, "modConnectors.AddConnectorInstance", "DTM-04P", "DTM 4-way",
        "DTM06-4S", "Connector", 4)
    row = 2
    sheet.Cells(row, 1).Value = "J9"

    result = run_action(wb, "modConnectors.ApplyConnectorEdit",
                        sheet.Cells(row, 1), "J1", row)
    assert (result.ok, result.outcome, result.payload) == (True, "RENAMED", "J9")


def test_a_rename_onto_an_existing_ref_des_is_rejected_with_the_value_to_restore(wb):
    sheet = wb.Worksheets("Connectors")
    run(wb, "modConnectors.AddConnectorInstance", "DTM-04P", "DTM 4-way",
        "DTM06-4S", "Connector", 4)
    run(wb, "modConnectors.AddConnectorInstance", "DTM-04P", "DTM 4-way",
        "DTM06-4S", "Connector", 4)
    sheet.Cells(2, 1).Value = "J2"

    result = run_action(wb, "modConnectors.ApplyConnectorEdit",
                        sheet.Cells(2, 1), "J1", 2)
    assert (result.ok, result.outcome, result.payload) == (False, "RENAME_REJECTED", "J1")


def test_an_edit_that_is_not_a_rename_reports_no_rename(wb):
    sheet = wb.Worksheets("Connectors")
    run(wb, "modConnectors.AddConnectorInstance", "DTM-04P", "DTM 4-way",
        "DTM06-4S", "Connector", 4)

    result = run_action(wb, "modConnectors.ApplyConnectorEdit",
                        sheet.Cells(2, 3), "J1", 2)
    assert (result.ok, result.outcome) == (False, "NO_RENAME")
