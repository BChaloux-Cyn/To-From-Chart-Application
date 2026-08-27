from tests.conftest import run


def add(wb, connector_id="DTM-04P"):
    return run(wb, "modConnectors.AddConnectorInstance",
               connector_id, "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)


def test_remove_deletes_the_connector_row(wb):
    add(wb)  # J1
    add(wb)  # J2

    ok = run(wb, "modConnectors.RemoveConnectorInstance", "J1")
    assert ok is True

    conn_sheet = wb.Worksheets("Connectors")
    assert run(wb, "modConnectors.PinCountFor", "J1") == 0
    assert run(wb, "modConnectors.PinCountFor", "J2") == 4  # untouched


def test_remove_clears_only_the_referencing_endpoint(wb):
    add(wb)  # J1
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(7, 2).Value = 1
    sheet.Cells(7, 4).Value = "+12V Batt"  # Signal - must survive

    run(wb, "modConnectors.RemoveConnectorInstance", "J1")

    assert sheet.Cells(7, 1).Value is None
    assert sheet.Cells(7, 2).Value is None
    assert sheet.Cells(7, 4).Value == "+12V Batt"


def test_remove_clears_a_to_endpoint_without_touching_the_from_endpoint(wb):
    add(wb)  # J1
    add(wb)  # J2
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J2"
    sheet.Cells(7, 9).Value = "J1"

    run(wb, "modConnectors.RemoveConnectorInstance", "J1")

    assert sheet.Cells(7, 1).Value == "J2"
    assert sheet.Cells(7, 9).Value is None


def test_remove_unknown_ref_des_returns_false(wb):
    assert run(wb, "modConnectors.RemoveConnectorInstance", "J99") is False


def test_remove_does_not_trip_the_sheet_s_rename_detection(wb):
    # shConnectors.evt caches whatever ref des was selected before an edit,
    # to detect a rename on the next Worksheet_Change. RemoveConnectorInstance
    # writes column A directly (compacting J2's row into J1's slot) - with
    # events left enabled, that write looks exactly like a J1 -> J2 rename to
    # a row that still (briefly) has J2 sitting at the old row too, so the
    # handler used to see a collision and revert it mid-removal.
    add(wb)  # J1
    add(wb)  # J2
    conn_sheet = wb.Worksheets("Connectors")
    conn_sheet.Activate()
    conn_sheet.Cells(2, 1).Select()  # caches "J1" as shConnectors.evt's prior value

    ok = run(wb, "modConnectors.RemoveConnectorInstance", "J1")
    assert ok is True

    assert conn_sheet.Cells(2, 1).Value == "J2"
    assert conn_sheet.Cells(3, 1).Value is None
