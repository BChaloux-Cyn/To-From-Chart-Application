import pytest

FIELD_NAMES = [
    "txtName", "txtManufacturer", "txtPartNumber", "cboType",
    "txtPinCount", "txtNotes",
]

COMMAND_CONTROLS = [
    "cmdLoadPhoto", "tglPlacePins", "cmdDeletePin", "cmdClearPins",
    "cmdSnapLabel", "cmdSave", "cmdCancel",
]


def controls(wb):
    return wb.VBProject.VBComponents("frmConnectorEditor").Designer.Controls


def test_form_exists_with_the_expected_caption(wb):
    form = wb.VBProject.VBComponents("frmConnectorEditor").Designer
    assert form.Caption == "Connector Editor"


@pytest.mark.parametrize("name", FIELD_NAMES)
def test_field_control_exists(wb, name):
    assert controls(wb)(name).Name == name


@pytest.mark.parametrize("name", COMMAND_CONTROLS)
def test_command_control_exists(wb, name):
    assert controls(wb)(name).Name == name


def test_image_control_exists_for_the_photo(wb):
    assert controls(wb)("imgPhoto").Name == "imgPhoto"


def test_pin_list_control_exists(wb):
    assert controls(wb)("lstPins").Name == "lstPins"


def test_type_list_source_has_the_four_types(wb):
    # cboType is populated at runtime (frmConnectorEditor.UserForm_Initialize)
    # from _Lists!D2:D5, not via a design-time RowSource binding - an
    # unqualified RowSource string resolves against whichever workbook is
    # ActiveWorkbook when the control binds, which left the Type dropdown
    # empty during frmManageLibrary's Edit flow (manual verification,
    # phase-2-manual-verification.md, 2c). This checks the underlying data;
    # test_type_combo_is_populated_from_this_workbooks_lists_sheet (wiring
    # test) checks the runtime population code itself.
    lists_sheet = wb.Worksheets("_Lists")
    values = [lists_sheet.Cells(r, 4).Value for r in range(2, 6)]
    assert values == ["Connector", "Stud", "Splice", "Tail"]


def test_notes_field_is_multiline(wb):
    assert controls(wb)("txtNotes").MultiLine is True


def test_save_button_caption(wb):
    assert controls(wb)("cmdSave").Caption == "Save"


def test_cancel_button_caption(wb):
    assert controls(wb)("cmdCancel").Caption == "Cancel"
