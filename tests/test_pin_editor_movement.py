from tests.conftest import run


def place(wb, sheet, pin_number=1, x=0.5, y=0.5):
    run(wb, "modPinEditor.PlacePin", sheet, "J1", pin_number, "A", x, y)


def read_pin(wb, sheet):
    pins = run(wb, "modLibrary.ReadPinsForConnector", sheet, 2, 2000, "J1")
    return pins[0]  # connector_id, pin_number, label, norm_x, norm_y, label_x, label_y


def test_move_marker_only_changes_the_label_position(wb):
    sheet = wb.Worksheets("_Edit")
    place(wb, sheet, x=0.2, y=0.2)

    ok = run(wb, "modPinEditor.MoveMarker", sheet, "J1", 1, 0.8, 0.9)
    assert ok is True

    _, _, _, norm_x, norm_y, label_x, label_y = read_pin(wb, sheet)
    assert (norm_x, norm_y) == (0.2, 0.2)
    assert (label_x, label_y) == (0.8, 0.9)


def test_move_anchor_carries_a_marker_that_was_still_on_it(wb):
    sheet = wb.Worksheets("_Edit")
    place(wb, sheet, x=0.2, y=0.2)  # marker starts on the anchor

    ok = run(wb, "modPinEditor.MoveAnchor", sheet, "J1", 1, 0.6, 0.7)
    assert ok is True

    _, _, _, norm_x, norm_y, label_x, label_y = read_pin(wb, sheet)
    assert (norm_x, norm_y) == (0.6, 0.7)
    assert (label_x, label_y) == (0.6, 0.7)  # traveled with the anchor


def test_move_anchor_leaves_a_pulled_away_marker_in_place(wb):
    sheet = wb.Worksheets("_Edit")
    place(wb, sheet, x=0.2, y=0.2)
    run(wb, "modPinEditor.MoveMarker", sheet, "J1", 1, 0.9, 0.9)  # pull it away first

    run(wb, "modPinEditor.MoveAnchor", sheet, "J1", 1, 0.3, 0.3)

    _, _, _, norm_x, norm_y, label_x, label_y = read_pin(wb, sheet)
    assert (norm_x, norm_y) == (0.3, 0.3)
    assert (label_x, label_y) == (0.9, 0.9)  # stayed put, leader re-aims


def test_snap_label_to_pin_returns_the_marker_to_its_anchor(wb):
    sheet = wb.Worksheets("_Edit")
    place(wb, sheet, x=0.2, y=0.2)
    run(wb, "modPinEditor.MoveMarker", sheet, "J1", 1, 0.9, 0.9)

    ok = run(wb, "modPinEditor.SnapLabelToPin", sheet, "J1", 1)
    assert ok is True

    _, _, _, norm_x, norm_y, label_x, label_y = read_pin(wb, sheet)
    assert (label_x, label_y) == (norm_x, norm_y) == (0.2, 0.2)


def test_marker_sits_on_anchor_within_threshold(wb):
    assert run(wb, "modPinEditor.MarkerSitsOnAnchor", 0.5, 0.5, 0.505, 0.505) is True
    assert run(wb, "modPinEditor.MarkerSitsOnAnchor", 0.5, 0.5, 0.6, 0.6) is False


def test_needs_leader_line_reflects_marker_state(wb):
    sheet = wb.Worksheets("_Edit")
    place(wb, sheet, x=0.2, y=0.2)
    assert run(wb, "modPinEditor.NeedsLeaderLine", sheet, "J1", 1) is False

    run(wb, "modPinEditor.MoveMarker", sheet, "J1", 1, 0.9, 0.9)
    assert run(wb, "modPinEditor.NeedsLeaderLine", sheet, "J1", 1) is True


def test_move_unknown_pin_returns_false(wb):
    sheet = wb.Worksheets("_Edit")
    assert run(wb, "modPinEditor.MoveAnchor", sheet, "J1", 99, 0.1, 0.1) is False
    assert run(wb, "modPinEditor.MoveMarker", sheet, "J1", 99, 0.1, 0.1) is False
    assert run(wb, "modPinEditor.SnapLabelToPin", sheet, "J1", 99) is False


def test_pin_geometry_returns_anchor_and_label_coordinates(wb):
    # Added alongside NeedsLeaderLine for a leader-line visual indicator
    # (phase-2-manual-verification.md, 2b) - deferred (Forms.Line.1 isn't
    # available on this machine), but the read helper is correct and kept
    # for whenever the visual is added back.
    sheet = wb.Worksheets("_Edit")
    place(wb, sheet, x=0.2, y=0.2)
    run(wb, "modPinEditor.MoveMarker", sheet, "J1", 1, 0.9, 0.8)

    result = run(wb, "modPinEditor.PinGeometry", sheet, "J1", 1)

    assert tuple(result) == (0.2, 0.2, 0.9, 0.8)


def test_pin_geometry_for_unknown_pin_is_empty(wb):
    sheet = wb.Worksheets("_Edit")
    result = run(wb, "modPinEditor.PinGeometry", sheet, "J1", 99)
    assert result is None
