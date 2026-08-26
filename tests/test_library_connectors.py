from tests.conftest import run

FIELDS = (
    "DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
    4, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
)

# Excel's COM Value property always reads numbers back as float (4 -> 4.0)
# and an empty string written to a cell reads back as None (Excel treats ""
# as clearing the cell to blank) - this is what a round trip actually
# yields, independent of anything modLibrary does.
FIELDS_ROUND_TRIPPED = FIELDS[:5] + (4.0, None, None) + FIELDS[8:]


def test_write_then_read_round_trips_every_field(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    ok = run(wb, "modLibrary.WriteConnector", ws, 2, 100000, FIELDS)
    assert ok is True

    result = run(wb, "modLibrary.ReadConnector", ws, 2, 100000, "DTM-04P")
    assert tuple(result) == FIELDS_ROUND_TRIPPED


def test_read_missing_connector_returns_empty(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    result = run(wb, "modLibrary.ReadConnector", ws, 2, 100000, "NO-SUCH-ID")
    assert result is None


def test_write_upserts_an_existing_connector(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, FIELDS)

    updated = FIELDS[:5] + (8,) + FIELDS[6:]
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, updated)

    assert ws.Cells(2, 1).Value == "DTM-04P"
    assert ws.Cells(3, 1).Value is None  # no second row was added
    result = run(wb, "modLibrary.ReadConnector", ws, 2, 100000, "DTM-04P")
    assert int(result[5]) == 8


def test_write_respects_the_row_window(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    # nLastRow=2 leaves no room past the header for a first data row.
    ok = run(wb, "modLibrary.WriteConnector", ws, 2, 1, FIELDS)
    assert ok is False


def test_delete_removes_a_connector_and_compacts(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, FIELDS)
    other = ("GND-STUD",) + FIELDS[1:]
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, other)

    ok = run(wb, "modLibrary.DeleteConnector", ws, 2, 100000, "DTM-04P")
    assert ok is True
    assert run(wb, "modLibrary.ReadConnector", ws, 2, 100000, "DTM-04P") is None
    assert run(wb, "modLibrary.ReadConnector", ws, 2, 100000, "GND-STUD") is not None
    assert ws.Cells(3, 1).Value is None  # compacted, not left as a gap


def test_delete_stays_inside_its_row_window(wb, library_wb):
    # A row beyond nLastRow must never be touched by delete's compaction.
    ws = library_wb.Worksheets("Connectors")
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, FIELDS)
    ws.Cells(3, 1).Value = "SENTINEL"

    run(wb, "modLibrary.DeleteConnector", ws, 2, 2, "DTM-04P")

    assert ws.Cells(2, 1).Value == "SENTINEL" or ws.Cells(3, 1).Value == "SENTINEL"
