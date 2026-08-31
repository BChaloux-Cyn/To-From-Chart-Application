from tests.conftest import run, run_action


def add(wb, connector_id="DTM-04P"):
    return run(wb, "modConnectors.AddConnectorInstance",
               connector_id, "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)


def test_instance_index_is_empty_with_no_connectors_placed(wb):
    assert run(wb, "modConnectors.InstanceIndex") is None


def test_instance_index_lists_ref_des_and_name(wb):
    add(wb)  # J1
    add(wb, "DTM-08P")  # J2

    index = run(wb, "modConnectors.InstanceIndex")
    assert [row[0] for row in index] == ["J1 - Deutsch DTM 4-way", "J2 - Deutsch DTM 4-way"]
    assert [row[1] for row in index] == ["J1", "J2"]


def test_remove_instance_action_removes_and_reports_the_ref_des(wb):
    add(wb)  # J1
    result = run_action(wb, "modConnectorActions.RemoveInstance", "J1")
    assert (result.ok, result.outcome, result.payload) == (True, "INSTANCE_REMOVED", "J1")
    assert run(wb, "modConnectors.PinCountFor", "J1") == 0


def test_remove_instance_action_reports_an_unknown_ref_des(wb):
    result = run_action(wb, "modConnectorActions.RemoveInstance", "J99")
    assert (result.ok, result.outcome, result.payload) == (False, "INSTANCE_NOT_FOUND", "J99")


def test_remove_instances_of_type_removes_every_matching_ref_des(wb):
    add(wb)  # J1, DTM-04P
    add(wb, "DTM-08P")  # J2
    add(wb)  # J3, DTM-04P

    removed = run(wb, "modConnectors.RemoveInstancesOfConnectorType", "DTM-04P")
    assert list(removed) == ["J1", "J3"]
    assert run(wb, "modConnectors.PinCountFor", "J1") == 0
    assert run(wb, "modConnectors.PinCountFor", "J2") == 4  # untouched
    assert run(wb, "modConnectors.PinCountFor", "J3") == 0


def test_remove_instances_of_type_is_empty_when_none_are_placed(wb):
    add(wb, "DTM-08P")  # J1, a different library connector
    assert run(wb, "modConnectors.RemoveInstancesOfConnectorType", "DTM-04P") is None


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
