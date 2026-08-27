import pytest

VBEXT_CT_CLASSMODULE = 2


def module_source(wb, component_name):
    module = wb.VBProject.VBComponents(component_name).CodeModule
    return module.Lines(1, module.CountOfLines)


def test_pin_marker_class_exists(wb):
    names = [wb.VBProject.VBComponents(i + 1).Name for i in range(wb.VBProject.VBComponents.Count)]
    assert "clsPinMarker" in names


def test_pin_marker_is_a_class_module(wb):
    assert wb.VBProject.VBComponents("clsPinMarker").Type == VBEXT_CT_CLASSMODULE


def test_pin_marker_handles_mouse_drag_events(wb):
    source = module_source(wb, "clsPinMarker")
    for handler in ("mLabel_MouseDown", "mLabel_MouseMove", "mLabel_MouseUp"):
        assert handler in source


def test_pin_marker_calls_move_marker_on_drop(wb):
    assert "modPinEditor.MoveMarker" in module_source(wb, "clsPinMarker")


def test_pin_marker_label_control_is_readable(wb):
    # LabelControl must be readable, not just settable: frmConnectorEditor
    # reads it back (cmdClearPins_Click, RepositionMarkerControl) to find
    # the on-screen Label tied to a given pin.
    assert "Property Get LabelControl" in module_source(wb, "clsPinMarker")


@pytest.mark.parametrize(
    "handler",
    [
        "cmdLoadPhoto_Click", "imgPhoto_MouseUp", "cmdDeletePin_Click",
        "cmdClearPins_Click", "cmdSnapLabel_Click", "cmdSave_Click", "cmdCancel_Click",
    ],
)
def test_form_wires_every_command(wb, handler):
    assert handler in module_source(wb, "frmConnectorEditor")


def test_form_calls_into_the_tested_logic_module_for_every_pin_action(wb):
    source = module_source(wb, "frmConnectorEditor")
    for call in (
        "modPinEditor.PlacePin", "modPinEditor.RemovePin", "modPinEditor.ClearScratchPins",
        "modPinEditor.SnapLabelToPin", "modPinEditor.SaveConnector", "modPinEditor.FitAspectRatio",
        "modPinEditor.MoveAnchor",
    ):
        assert call in source


def test_form_tracks_pin_numbers_independently_of_list_position(wb):
    # lstPins position and PinNumber diverge once anything is deleted -
    # cmdDeletePin_Click/cmdSnapLabel_Click must resolve the pin number
    # through mListPinNumbers rather than assuming ListIndex + 1 == PinNumber.
    source = module_source(wb, "frmConnectorEditor")
    assert "mListPinNumbers" in source
    assert "RemovePin ThisWorkbook.Worksheets(\"_Edit\"), mConnectorID, lstPins.ListIndex + 1" not in source
    assert "SnapLabelToPin ThisWorkbook.Worksheets(\"_Edit\"), mConnectorID, lstPins.ListIndex + 1" not in source


def test_load_photo_filter_excludes_png(wb):
    # Manual verification (phase-2-manual-verification.md, 2b) found VBA's
    # LoadPicture throws error 481 "Invalid picture" on valid PNGs on some
    # Windows/Office configurations, even though Shapes.AddPicture handles
    # the same file fine. The picker must only offer formats LoadPicture
    # actually loads here (JPG), not silently let a PNG through to fail.
    source = module_source(wb, "frmConnectorEditor")
    filter_line = next(line for line in source.splitlines() if "GetOpenFilename" in line
                        or ("*.jpg" in line and "*.jpeg" in line))
    assert "*.png" not in filter_line
    assert "*.jpg" in filter_line and "*.jpeg" in filter_line


def test_type_combo_is_populated_from_this_workbooks_lists_sheet(wb):
    # Manual verification (phase-2-manual-verification.md, 2c) found the
    # Type dropdown had no options at all during frmManageLibrary's Edit
    # flow: cboType's old RowSource ("_Lists!D2:D5", unqualified) resolved
    # against ActiveWorkbook, which by then was ConnectorLibrary.xlsx, not
    # this workbook. UserForm_Initialize must populate it explicitly against
    # ThisWorkbook instead, and this must happen before LoadForEdit's
    # cboType.Text assignment can rely on the list being populated.
    source = module_source(wb, "frmConnectorEditor")
    init_sub = source[source.index("Private Sub UserForm_Initialize"):]
    init_body = init_sub.split("End Sub", 1)[0]
    assert "ThisWorkbook.Worksheets(\"_Lists\")" in init_body
    assert "cboType.AddItem" in init_body
    assert "If cboType.ListCount > 0 Then cboType.ListIndex = 0" in init_body


def test_load_existing_photo_prefers_the_disk_cache_over_the_embedded_shape(wb):
    # Manual verification (phase-2-manual-verification.md, 2c) found the
    # Edit-mode photo preview permanently blank: re-exporting the embedded
    # library Shape via Shape.Copy/Chart.Paste (ExportShapeToFile) is
    # unreliable for VBA-triggered clipboard operations on this machine.
    # LoadExistingPhoto must check the on-disk cache first - no clipboard -
    # falling back to ExportShapeToFile only as a one-time backfill attempt
    # when no cache exists yet (e.g. a connector saved before this fix).
    source = module_source(wb, "frmConnectorEditor")
    load_existing = source[source.index("Private Sub LoadExistingPhoto"):]
    body = load_existing.split("End Sub", 1)[0]
    cache_check = body.index("modLibrary.CachePhotoPath(ThisWorkbook.Path, sConnectorID, \"jpg\")")
    fallback = body.index("ExportShapeToFile")
    assert cache_check < fallback  # cache is checked before the fallback runs


def test_save_refreshes_the_photo_cache_via_plain_file_copy(wb):
    # The cache LoadExistingPhoto reads from (above) must be kept current -
    # a plain FileCopy after a successful save, not another clipboard trip.
    source = module_source(wb, "frmConnectorEditor")
    save_click = source[source.index("Private Sub cmdSave_Click"):]
    body = save_click.split("End Sub", 1)[0]
    assert "modLibrary.CachePhotoPath(ThisWorkbook.Path, mConnectorID, \"jpg\")" in body
    assert "FileCopy mPhotoPath, sCachePath" in body


def test_load_photo_requires_name_and_part_number_first(wb):
    # Manual verification (phase-2-manual-verification.md, 2b) found
    # mConnectorID is computed from Name/Part Number at photo-load time and
    # never recomputed - placing pins before both are filled in silently
    # does nothing. cmdLoadPhoto_Click must reject a blank Name or Part
    # Number before opening the file picker, not let the photo load anyway.
    source = module_source(wb, "frmConnectorEditor")
    click_start = source.index("Private Sub cmdLoadPhoto_Click")
    click_end = source.index("GetOpenFilename", click_start)
    guard = source[click_start:click_end]
    assert "txtName.Text" in guard and "txtPartNumber.Text" in guard


def test_pin_marker_uses_a_small_fixed_badge_size(wb):
    # Manual verification (phase-2-manual-verification.md, 2b) found pin
    # marker labels used the Label control's default (body-text-sized)
    # dimensions, dwarfing the photo. AddMarkerControl must set an explicit
    # small size rather than leaving the control at its default.
    source = module_source(wb, "frmConnectorEditor")
    assert "PIN_MARKER_SIZE" in source
    add_marker = source[source.index("Private Sub AddMarkerControl"):]
    assert "lbl.Width = PIN_MARKER_SIZE" in add_marker
    assert "lbl.Height = PIN_MARKER_SIZE" in add_marker


def test_place_pins_is_capped_at_pin_count(wb):
    # Manual verification (phase-2-manual-verification.md, 2b) found nothing
    # stopped placing more pins than the entered Pin Count. imgPhoto_MouseUp
    # must refuse once the placed count reaches txtPinCount's value.
    source = module_source(wb, "frmConnectorEditor")
    handler = source[source.index("Private Sub imgPhoto_MouseUp"):]
    assert "txtPinCount.Text" in handler
    assert "lstPins.ListCount >= nPinCount" in handler


def test_save_rejects_a_part_number_that_collides_with_another_connector(wb):
    # Manual verification (phase-2-manual-verification.md, 2b) flagged that
    # Part Number had no uniqueness check - saving a new connector with an
    # existing Part Number silently overwrote that connector's row (same
    # derived ID). cmdSave_Click must check for a collision with a
    # DIFFERENT connector (mOriginalConnectorID excludes re-saving self).
    source = module_source(wb, "frmConnectorEditor")
    assert "mOriginalConnectorID" in source
    save_click = source[source.index("Private Sub cmdSave_Click"):]
    assert "modLibrary.FindConnectorRow(lib.Worksheets(\"Connectors\")" in save_click
    assert "StrComp(mConnectorID, mOriginalConnectorID" in save_click


def test_photo_fit_box_is_a_fixed_constant(wb):
    # Manual verification (phase-2-manual-verification.md, 2b) found that
    # fitting against imgPhoto.Width/Height - which the same code then
    # overwrites with the fitted result - shrinks the preview further on
    # every subsequent Load Photo. The box must come from a fixed constant.
    source = module_source(wb, "frmConnectorEditor")
    assert "PHOTO_BOX_WIDTH" in source and "PHOTO_BOX_HEIGHT" in source
    assert "FitAspectRatio(CDbl(pic.Width), CDbl(pic.Height), imgPhoto.Width, imgPhoto.Height)" not in source
    for call in source.splitlines():
        if "modPinEditor.FitAspectRatio(" in call:
            assert "PHOTO_BOX_WIDTH" in call and "PHOTO_BOX_HEIGHT" in call
