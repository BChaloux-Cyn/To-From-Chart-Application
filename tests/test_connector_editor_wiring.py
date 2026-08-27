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
