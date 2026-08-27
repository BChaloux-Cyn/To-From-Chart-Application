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


def test_type_combo_is_seeded_with_the_four_types(wb):
    # cboType is RowSource-bound (build/form_layout.py) rather than seeded
    # via AddItem, so win32com surfaces .List as the bound range's raw
    # 2D array (row tuples), not an indexed List(i) accessor.
    combo = controls(wb)("cboType")
    values = [row[0] for row in combo.List]
    assert values == ["Connector", "Stud", "Splice", "Tail"]


def test_notes_field_is_multiline(wb):
    assert controls(wb)("txtNotes").MultiLine is True


def test_save_button_caption(wb):
    assert controls(wb)("cmdSave").Caption == "Save"


def test_cancel_button_caption(wb):
    assert controls(wb)("cmdCancel").Caption == "Cancel"
