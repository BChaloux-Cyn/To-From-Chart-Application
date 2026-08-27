from tests.conftest import run


def add(wb, connector_id="DTM-04P"):
    return run(wb, "modConnectors.AddConnectorInstance",
               connector_id, "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)


def test_rename_rewrites_every_chart_reference(wb):
    add(wb)  # J1
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(8, 9).Value = "J1"

    conn_sheet = wb.Worksheets("Connectors")
    conn_sheet.Cells(2, 1).Value = "J99"  # simulates the cell already having been edited

    ok = run(wb, "modConnectors.RenameRefDes", "J1", "J99")
    assert ok is True
    assert sheet.Cells(7, 1).Value == "J99"
    assert sheet.Cells(8, 9).Value == "J99"


def test_rename_leaves_unrelated_rows_alone(wb):
    add(wb)  # J1
    add(wb)  # J2
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(8, 1).Value = "J2"

    conn_sheet = wb.Worksheets("Connectors")
    conn_sheet.Cells(2, 1).Value = "J99"

    run(wb, "modConnectors.RenameRefDes", "J1", "J99")
    assert sheet.Cells(8, 1).Value == "J2"


def test_rename_colliding_with_an_existing_ref_des_is_rejected(wb):
    add(wb)  # J1
    add(wb)  # J2

    conn_sheet = wb.Worksheets("Connectors")
    conn_sheet.Cells(2, 1).Value = "J2"  # simulates renaming J1 to the already-used J2

    ok = run(wb, "modConnectors.RenameRefDes", "J1", "J2")
    assert ok is False


def test_rename_to_the_same_value_is_a_no_op(wb):
    add(wb)  # J1
    assert run(wb, "modConnectors.RenameRefDes", "J1", "J1") is False


def test_renaming_via_the_sheet_reverts_the_cell_on_collision(wb):
    add(wb)  # J1
    add(wb)  # J2
    conn_sheet = wb.Worksheets("Connectors")

    # Select requires the sheet to be active first - under headless COM
    # automation the workbook opens with whichever sheet was active at
    # build time (Home), not Connectors, so Select alone raises "Select
    # method of Range class failed."
    conn_sheet.Activate()
    conn_sheet.Cells(2, 1).Select()  # caches "J1" as the prior value - () matters: a bare `.Select`
                                      # is a Python attribute access on the COM method, never calls it
    conn_sheet.Cells(2, 1).Value = "J2"  # collides - shConnectors.evt must revert this

    assert conn_sheet.Cells(2, 1).Value == "J1"
