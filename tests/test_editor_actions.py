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
