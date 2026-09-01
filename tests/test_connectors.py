import pytest

from tests.conftest import run

HEADERS = ["Ref Des", "ConnectorID", "Name", "Part Number", "Type", "Pin Count"]


@pytest.mark.parametrize("index,header", list(enumerate(HEADERS, start=1)))
def test_connectors_headers(wb, index, header):
    assert wb.Worksheets("Connectors").Cells(1, index).Value == header


@pytest.mark.parametrize(
    "connector_type,prefix",
    [
        ("Connector", "J"),
        ("Stud", "ST"),
        ("Splice", "SP"),
        ("Tail", "TL"),
        ("connector", "J"),
        ("  Tail  ", "TL"),
    ],
)
def test_prefix_for_type(wb, connector_type, prefix):
    assert run(wb, "modConnectors.PrefixForType", connector_type) == prefix


def test_unknown_type_has_no_prefix(wb):
    assert run(wb, "modConnectors.PrefixForType", "Widget") == ""


def test_first_ref_des_of_each_prefix_is_one(wb):
    assert run(wb, "modConnectors.NextRefDes", "J") == "J1"
    assert run(wb, "modConnectors.NextRefDes", "ST") == "ST1"


def test_ref_des_numbering_increments_per_prefix(wb):
    assert run(wb, "modConnectors.AddConnectorInstance",
               "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4) == "J1"
    assert run(wb, "modConnectors.AddConnectorInstance",
               "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4) == "J2"
    # A different prefix numbers independently.
    assert run(wb, "modConnectors.AddConnectorInstance",
               "GND-STUD", "Chassis ground stud", "", "Stud", 1) == "ST1"


def test_same_part_can_appear_twice_as_distinct_instances(wb):
    first = run(wb, "modConnectors.AddConnectorInstance",
                "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    second = run(wb, "modConnectors.AddConnectorInstance",
                 "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    assert first != second
    sheet = wb.Worksheets("Connectors")
    assert sheet.Cells(2, 2).Value == sheet.Cells(3, 2).Value == "DTM-04P"


def test_instance_row_is_written_in_full(wb):
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    sheet = wb.Worksheets("Connectors")
    row = [sheet.Cells(2, c).Value for c in range(1, 7)]
    assert row == ["J1", "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4]


def test_unknown_type_is_rejected(wb):
    assert run(wb, "modConnectors.AddConnectorInstance",
               "X", "X", "", "Widget", 4) == ""
    assert wb.Worksheets("Connectors").Cells(2, 1).Value is None


def test_pin_count_below_one_is_rejected(wb):
    assert run(wb, "modConnectors.AddConnectorInstance",
               "X", "X", "", "Connector", 0) == ""


def test_adding_a_connector_marks_the_workbook_dirty(wb):
    run(wb, "modState.ClearDirty")
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    assert run(wb, "modState.IsDirty") is True


def test_pin_count_lookup(wb):
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-12P", "Deutsch DTM 12-way", "DTM06-12S", "Connector", 12)
    assert run(wb, "modConnectors.PinCountFor", "J1") == 12
    assert run(wb, "modConnectors.PinCountFor", "j1") == 12
    assert run(wb, "modConnectors.PinCountFor", "J99") == 0


def test_ref_des_dropdown_sees_added_connectors(wb):
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    assert wb.Names("ListRefDes").RefersToRange.Rows.Count == 1
    run(wb, "modConnectors.AddConnectorInstance",
        "GND-STUD", "Chassis ground stud", "", "Stud", 1)
    assert wb.Names("ListRefDes").RefersToRange.Rows.Count == 2


def test_all_instances_returns_every_placed_connector(wb):
    ws = wb.Worksheets("Connectors")
    ws.Cells(2, 1).Value = "J1"
    ws.Cells(2, 2).Value = "DTM-04P"
    ws.Cells(2, 3).Value = "Deutsch DTM 4-way"
    ws.Cells(2, 4).Value = "DTM06-4S"
    ws.Cells(2, 5).Value = "Connector"
    ws.Cells(2, 6).Value = 4
    ws.Cells(3, 1).Value = "ST1"
    ws.Cells(3, 2).Value = "GND-STUD"
    ws.Cells(3, 3).Value = "Chassis Ground"
    ws.Cells(3, 4).Value = ""
    ws.Cells(3, 5).Value = "Stud"
    ws.Cells(3, 6).Value = 1

    result = run(wb, "modConnectors.AllInstances")
    assert [row[0] for row in result] == ["J1", "ST1"]
    assert result[0][1] == "DTM-04P"
    assert int(result[1][5]) == 1


def test_all_instances_returns_empty_when_nothing_is_placed(wb):
    result = run(wb, "modConnectors.AllInstances")
    assert result is None
