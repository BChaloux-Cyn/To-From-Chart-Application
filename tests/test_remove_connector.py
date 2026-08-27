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
