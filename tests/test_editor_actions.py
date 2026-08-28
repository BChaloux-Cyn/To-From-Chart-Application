import pytest

from tests.conftest import run

CAP = 100000
SCRATCH_FIRST, SCRATCH_LAST = 2, 2000


def write_scratch_pin(wb, ws, connector_id, pin_number, label, nx, ny, lx, ly):
    return run(wb, "modLibrary.WritePin", ws, SCRATCH_FIRST, SCRATCH_LAST,
               (connector_id, pin_number, label, nx, ny, lx, ly))


def test_photo_file_filter_offers_jpg_only(wb):
    # LoadPicture raises error 481 on valid PNGs on some Office builds, so
    # the picker must not offer a format it cannot open.
    filter_string = run(wb, "modEditorActions.PhotoFileFilter")
    assert "*.jpg" in filter_string and "*.jpeg" in filter_string
    assert "*.png" not in filter_string.lower()


def test_marker_control_name_is_stable(wb):
    assert run(wb, "modEditorActions.MarkerControlName", 7) == "lblMarker7"


def test_cache_refresh_target_is_the_cache_path_for_a_new_photo(wb, tmp_path):
    target = run(wb, "modEditorActions.PhotoCacheRefreshTarget",
                 str(tmp_path), "DTM-04P", str(tmp_path / "chosen.jpg"))
    assert target.endswith("Photos\\DTM-04P.jpg")


def test_cache_refresh_target_is_blank_when_the_source_is_already_the_cache(wb, tmp_path):
    cache = run(wb, "modLibrary.CachePhotoPath", str(tmp_path), "DTM-04P", "jpg")
    assert run(wb, "modEditorActions.PhotoCacheRefreshTarget",
               str(tmp_path), "DTM-04P", cache) == ""


def test_cache_refresh_target_is_blank_when_no_photo_was_chosen(wb, tmp_path):
    assert run(wb, "modEditorActions.PhotoCacheRefreshTarget",
               str(tmp_path), "DTM-04P", "") == ""


def test_type_list_items_reads_this_workbooks_lists_sheet(wb):
    # The old RowSource resolved against ActiveWorkbook, which left the
    # combo empty during the Edit flow.
    items = run(wb, "modEditorActions.TypeListItems", wb.Worksheets("_Lists"))
    assert "Connector" in [row for row in items]


def test_pin_list_items_renders_display_strings_and_pin_numbers(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    write_scratch_pin(wb, ws, "J1", 1, "Pin 1", 0.1, 0.1, 0.1, 0.1)
    write_scratch_pin(wb, ws, "J1", 2, "Pin 2", 0.2, 0.2, 0.3, 0.4)

    items = run(wb, "modEditorActions.PinListItems", ws, "J1")
    assert [row[0] for row in items] == ["Pin 1", "Pin 2"]
    assert [int(row[1]) for row in items] == [1, 2]
    assert items[1][2] == pytest.approx(0.3)
    assert items[1][3] == pytest.approx(0.4)


def test_pin_list_items_survives_a_deleted_middle_pin(wb):
    # This is what retires mListPinNumbers: list position and pin number
    # diverge, and the pin number must still be recoverable by position.
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    for n in (1, 2, 3):
        write_scratch_pin(wb, ws, "J1", n, f"Pin {n}", 0.1, 0.1, 0.1, 0.1)
    run(wb, "modPinEditor.RemovePin", ws, "J1", 2)

    items = run(wb, "modEditorActions.PinListItems", ws, "J1")
    assert [int(row[1]) for row in items] == [1, 3]


def test_pin_list_items_of_an_unplaced_connector_is_empty(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    assert run(wb, "modEditorActions.PinListItems", ws, "NOPE") is None


def test_next_pin_number_is_one_past_the_highest(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    assert run(wb, "modEditorActions.NextPinNumber", ws, "J1") == 1

    write_scratch_pin(wb, ws, "J1", 1, "Pin 1", 0.1, 0.1, 0.1, 0.1)
    write_scratch_pin(wb, ws, "J1", 5, "Pin 5", 0.1, 0.1, 0.1, 0.1)
    assert run(wb, "modEditorActions.NextPinNumber", ws, "J1") == 6


from tests.conftest import run_action


@pytest.mark.parametrize("name,part,expected", [
    ("DTM 4-way", "DTM06-4S", "OK"),
    ("", "DTM06-4S", "MISSING_NAME_OR_PART"),
    ("DTM 4-way", "", "MISSING_NAME_OR_PART"),
    ("", "", "MISSING_NAME_OR_PART"),
    ("   ", "DTM06-4S", "MISSING_NAME_OR_PART"),
])
def test_can_load_photo_requires_both_fields(wb, name, part, expected):
    result = run_action(wb, "modEditorActions.CanLoadPhoto", name, part)
    assert result.outcome == expected
    assert result.ok is (expected == "OK")


def test_photo_source_needs_backfill_when_no_cache_file_exists(wb, tmp_path):
    result = run_action(wb, "modEditorActions.PhotoSourceForEdit", str(tmp_path), "DTM-04P")
    assert result.outcome == "NEEDS_BACKFILL"
    assert result.payload.endswith("Photos\\DTM-04P.jpg")


def test_photo_source_is_cache_ready_once_the_file_exists(wb, tmp_path):
    cache = run(wb, "modLibrary.CachePhotoPath", str(tmp_path), "DTM-04P", "jpg")
    from pathlib import Path
    Path(cache).write_bytes(b"not a real jpeg, but the file exists")

    result = run_action(wb, "modEditorActions.PhotoSourceForEdit", str(tmp_path), "DTM-04P")
    assert result.outcome == "CACHE_READY"
    assert result.payload == cache


def test_delete_pin_removes_it_from_the_scratch_sheet(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    write_scratch_pin(wb, ws, "J1", 1, "Pin 1", 0.1, 0.1, 0.1, 0.1)
    write_scratch_pin(wb, ws, "J1", 2, "Pin 2", 0.2, 0.2, 0.2, 0.2)

    result = run_action(wb, "modEditorActions.DeletePinRequest", ws, "J1", 2)
    assert (result.ok, result.outcome, result.payload) == (True, "PIN_DELETED", 2)
    assert [int(r[1]) for r in run(wb, "modEditorActions.PinListItems", ws, "J1")] == [1]


def test_delete_pin_reports_a_missing_pin(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    result = run_action(wb, "modEditorActions.DeletePinRequest", ws, "J1", 9)
    assert (result.ok, result.outcome, result.payload) == (False, "PIN_NOT_FOUND", 9)


def test_click_in_place_mode_places_the_next_pin(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)

    result = run_action(wb, "modEditorActions.PhotoClickAction",
                        ws, "J1", True, 0, "4", 0.25, 0.75)
    assert (result.ok, result.outcome, result.payload) == (True, "PLACED", 1)

    items = run(wb, "modEditorActions.PinListItems", ws, "J1")
    assert items[0][2] == pytest.approx(0.25)
    assert items[0][3] == pytest.approx(0.75)


def test_placing_is_capped_at_the_entered_pin_count(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    for _ in range(2):
        run_action(wb, "modEditorActions.PhotoClickAction", ws, "J1", True, 0, "2", 0.1, 0.1)

    result = run_action(wb, "modEditorActions.PhotoClickAction",
                        ws, "J1", True, 0, "2", 0.5, 0.5)
    assert (result.ok, result.outcome, result.payload) == (False, "PIN_LIMIT_REACHED", 2)


@pytest.mark.parametrize("pin_count_text", ["", "   ", "abc", "0", "-3"])
def test_a_non_positive_pin_count_is_rejected(wb, pin_count_text):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    result = run_action(wb, "modEditorActions.PhotoClickAction",
                        ws, "J1", True, 0, pin_count_text, 0.5, 0.5)
    assert (result.ok, result.outcome) == (False, "BAD_PIN_COUNT")


def test_a_click_with_no_connector_id_is_a_no_op(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    result = run_action(wb, "modEditorActions.PhotoClickAction",
                        ws, "", True, 0, "4", 0.5, 0.5)
    assert (result.ok, result.outcome) == (False, "NO_OP")


def test_click_out_of_place_mode_moves_the_selected_pins_anchor(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    run_action(wb, "modEditorActions.PhotoClickAction", ws, "J1", True, 0, "4", 0.1, 0.1)

    result = run_action(wb, "modEditorActions.PhotoClickAction",
                        ws, "J1", False, 1, "4", 0.8, 0.9)
    assert (result.ok, result.outcome, result.payload) == (True, "MOVED_ANCHOR", 1)

    geometry = run(wb, "modPinEditor.PinGeometry", ws, "J1", 1)
    assert geometry[0] == pytest.approx(0.8)
    assert geometry[1] == pytest.approx(0.9)


def test_click_out_of_place_mode_with_nothing_selected_is_a_no_op(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    result = run_action(wb, "modEditorActions.PhotoClickAction",
                        ws, "J1", False, 0, "4", 0.5, 0.5)
    assert (result.ok, result.outcome) == (False, "NO_OP")


def test_pin_numbers_are_not_reused_after_a_deletion(wb):
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    for _ in range(3):
        run_action(wb, "modEditorActions.PhotoClickAction", ws, "J1", True, 0, "9", 0.1, 0.1)
    run_action(wb, "modEditorActions.DeletePinRequest", ws, "J1", 2)

    result = run_action(wb, "modEditorActions.PhotoClickAction",
                        ws, "J1", True, 0, "9", 0.4, 0.4)
    assert result.payload == 4
