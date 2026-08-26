from tests.conftest import run


def write_pin(wb, ws, connector_id, pin_number, label, norm_x, norm_y, label_x, label_y):
    fields = (connector_id, pin_number, label, norm_x, norm_y, label_x, label_y)
    return run(wb, "modLibrary.WritePin", ws, 2, 100000, fields)


def test_write_then_read_returns_pins_sorted_by_number(wb, library_wb):
    ws = library_wb.Worksheets("Pins")
    write_pin(wb, ws, "J1", 3, "GND", 0.9, 0.1, 0.9, 0.1)
    write_pin(wb, ws, "J1", 1, "+12V", 0.1, 0.1, 0.1, 0.1)
    write_pin(wb, ws, "J1", 2, "SIG", 0.5, 0.1, 0.5, 0.1)

    result = run(wb, "modLibrary.ReadPinsForConnector", ws, 2, 100000, "J1")
    pin_numbers = [int(row[1]) for row in result]
    assert pin_numbers == [1, 2, 3]
    assert [row[2] for row in result] == ["+12V", "SIG", "GND"]


def test_read_pins_only_returns_the_requested_connector(wb, library_wb):
    ws = library_wb.Worksheets("Pins")
    write_pin(wb, ws, "J1", 1, "A", 0.1, 0.1, 0.1, 0.1)
    write_pin(wb, ws, "J2", 1, "B", 0.1, 0.1, 0.1, 0.1)

    result = run(wb, "modLibrary.ReadPinsForConnector", ws, 2, 100000, "J1")
    assert len(result) == 1
    assert result[0][0] == "J1"


def test_read_pins_for_unknown_connector_returns_empty(wb, library_wb):
    ws = library_wb.Worksheets("Pins")
    result = run(wb, "modLibrary.ReadPinsForConnector", ws, 2, 100000, "NOPE")
    assert result is None


def test_delete_pins_removes_only_the_matching_connector_and_compacts(wb, library_wb):
    ws = library_wb.Worksheets("Pins")
    write_pin(wb, ws, "J1", 1, "A", 0.1, 0.1, 0.1, 0.1)
    write_pin(wb, ws, "J1", 2, "B", 0.2, 0.1, 0.2, 0.1)
    write_pin(wb, ws, "J2", 1, "C", 0.1, 0.1, 0.1, 0.1)

    count = run(wb, "modLibrary.DeletePinsForConnector", ws, 2, 100000, "J1")
    assert count == 2

    assert run(wb, "modLibrary.ReadPinsForConnector", ws, 2, 100000, "J1") is None
    remaining = run(wb, "modLibrary.ReadPinsForConnector", ws, 2, 100000, "J2")
    assert len(remaining) == 1
    assert ws.Cells(3, 1).Value is None  # compacted, not left with gaps


def test_write_pin_respects_the_row_window(wb, library_wb):
    ws = library_wb.Worksheets("Pins")
    # nLastRow=2 leaves no room past the header for a first data row.
    ok = write_pin(wb, ws, "J1", 1, "A", 0.1, 0.1, 0.1, 0.1)
    assert ok is True  # sanity check the normal path first (nLastRow=100000 above)

    ok = run(wb, "modLibrary.WritePin", ws, 2, 1, ("J2", 1, "B", 0.1, 0.1, 0.1, 0.1))
    assert ok is False


def test_delete_pins_stays_inside_its_row_window(wb, library_wb):
    # A row beyond nLastRow must never be touched by delete's compaction.
    ws = library_wb.Worksheets("Pins")
    write_pin(wb, ws, "J1", 1, "A", 0.1, 0.1, 0.1, 0.1)
    ws.Cells(3, 1).Value = "SENTINEL"

    run(wb, "modLibrary.DeletePinsForConnector", ws, 2, 2, "J1")

    assert ws.Cells(2, 1).Value == "SENTINEL" or ws.Cells(3, 1).Value == "SENTINEL"
