from tests.conftest import run


def test_place_pin_sets_anchor_and_marker_identical(wb):
    sheet = wb.Worksheets("_Edit")
    ok = run(wb, "modPinEditor.PlacePin", sheet, "J1", 1, "+12V", 0.25, 0.75)
    assert ok is True

    pins = run(wb, "modLibrary.ReadPinsForConnector", sheet, 2, 2000, "J1")
    assert len(pins) == 1
    connector_id, pin_number, label, norm_x, norm_y, label_x, label_y = pins[0]
    assert (pin_number, label) == (1, "+12V")
    assert (norm_x, norm_y) == (label_x, label_y) == (0.25, 0.75)


def test_re_placing_the_same_pin_number_replaces_it(wb):
    sheet = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.PlacePin", sheet, "J1", 1, "A", 0.1, 0.1)
    run(wb, "modPinEditor.PlacePin", sheet, "J1", 1, "B", 0.2, 0.2)

    pins = run(wb, "modLibrary.ReadPinsForConnector", sheet, 2, 2000, "J1")
    assert len(pins) == 1
    assert pins[0][2] == "B"


def test_remove_pin_leaves_others_intact(wb):
    sheet = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.PlacePin", sheet, "J1", 1, "A", 0.1, 0.1)
    run(wb, "modPinEditor.PlacePin", sheet, "J1", 2, "B", 0.2, 0.2)

    ok = run(wb, "modPinEditor.RemovePin", sheet, "J1", 1)
    assert ok is True

    pins = run(wb, "modLibrary.ReadPinsForConnector", sheet, 2, 2000, "J1")
    assert len(pins) == 1
    assert pins[0][1] == 2


def test_remove_unknown_pin_returns_false(wb):
    sheet = wb.Worksheets("_Edit")
    assert run(wb, "modPinEditor.RemovePin", sheet, "J1", 99) is False


def test_clear_scratch_pins_empties_the_sheet(wb):
    sheet = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.PlacePin", sheet, "J1", 1, "A", 0.1, 0.1)

    run(wb, "modPinEditor.ClearScratchPins", sheet)

    assert sheet.Cells(2, 1).Value is None


def test_load_scratch_pins_copies_from_the_library(wb, library_wb):
    ws_lib_pins = library_wb.Worksheets("Pins")
    for pin_number, label in [(1, "A"), (2, "B")]:
        fields = ("J1", pin_number, label, 0.1, 0.1, 0.1, 0.1)
        run(wb, "modLibrary.WritePin", ws_lib_pins, 2, 100000, fields)

    sheet = wb.Worksheets("_Edit")
    count = run(wb, "modPinEditor.LoadScratchPins", sheet, ws_lib_pins, "J1")

    assert count == 2
    pins = run(wb, "modLibrary.ReadPinsForConnector", sheet, 2, 2000, "J1")
    assert len(pins) == 2


def test_load_scratch_pins_clears_any_prior_session_first(wb, library_wb):
    sheet = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.PlacePin", sheet, "STALE", 1, "old", 0.5, 0.5)

    ws_lib_pins = library_wb.Worksheets("Pins")
    run(wb, "modPinEditor.LoadScratchPins", sheet, ws_lib_pins, "J1")

    assert run(wb, "modLibrary.ReadPinsForConnector", sheet, 2, 2000, "STALE") is None
