"""Field and control layout for frmConnectorEditor, built via the VBIDE
Designer object model at build time - there is no static .frm source."""
from __future__ import annotations

FORM_NAME = "frmConnectorEditor"
FORM_CAPTION = "Connector Editor"
FORM_WIDTH = 520
FORM_HEIGHT = 420

TYPE_CHOICES = ["Connector", "Stud", "Splice", "Tail"]
TYPE_LIST_COLUMN = 4

# (control ProgID, name, left, top, width, height, extra properties)
FIELD_CONTROLS = [
    ("Forms.Label.1", "lblName", 12, 12, 80, 16, {"Caption": "Name"}),
    ("Forms.TextBox.1", "txtName", 100, 12, 200, 18, {}),
    ("Forms.Label.1", "lblManufacturer", 12, 36, 80, 16, {"Caption": "Manufacturer"}),
    ("Forms.TextBox.1", "txtManufacturer", 100, 36, 200, 18, {}),
    ("Forms.Label.1", "lblPartNumber", 12, 60, 80, 16, {"Caption": "Part Number"}),
    ("Forms.TextBox.1", "txtPartNumber", 100, 60, 200, 18, {}),
    ("Forms.Label.1", "lblType", 12, 84, 80, 16, {"Caption": "Type"}),
    ("Forms.ComboBox.1", "cboType", 100, 84, 120, 18, {"Style": 2}),  # fmStyleDropDownList
    ("Forms.Label.1", "lblPinCount", 12, 108, 80, 16, {"Caption": "Pin Count"}),
    ("Forms.TextBox.1", "txtPinCount", 100, 108, 60, 18, {}),
    ("Forms.Label.1", "lblNotes", 12, 132, 80, 16, {"Caption": "Notes"}),
    ("Forms.TextBox.1", "txtNotes", 100, 132, 200, 40, {"MultiLine": True}),
    ("Forms.CommandButton.1", "cmdLoadPhoto", 320, 12, 90, 20, {"Caption": "Load Photo"}),
    ("Forms.Image.1", "imgPhoto", 320, 40, 180, 180, {}),
    # IntegralHeight defaults to True, which silently shrinks the box's own
    # Height the first time an item is added (rounding down to a whole
    # number of visible rows) - here that shrank a 120pt box to ~117pt,
    # which reads as the list "moving" even though Top never changes.
    ("Forms.ListBox.1", "lstPins", 12, 184, 180, 120, {"IntegralHeight": False}),
    ("Forms.ToggleButton.1", "tglPlacePins", 200, 184, 100, 20, {"Caption": "Place Pins"}),
    ("Forms.CommandButton.1", "cmdDeletePin", 200, 210, 100, 20, {"Caption": "Delete Pin"}),
    ("Forms.CommandButton.1", "cmdClearPins", 200, 236, 100, 20, {"Caption": "Clear Pins"}),
    ("Forms.CommandButton.1", "cmdSnapLabel", 200, 262, 100, 20, {"Caption": "Snap Label to Pin"}),
    ("Forms.CommandButton.1", "cmdSave", 320, 340, 80, 24, {"Caption": "Save"}),
    ("Forms.CommandButton.1", "cmdCancel", 410, 340, 80, 24, {"Caption": "Cancel"}),
]


def build_connector_editor_form(wb, add_userform) -> None:
    designer = add_userform(wb, FORM_NAME)
    designer.Caption = FORM_CAPTION
    # Designer.Width/Height aren't settable through win32com's dynamic
    # dispatch (unlike Caption); the VBComponent's Properties collection
    # reaches the same underlying value, same technique excel_com.set_codename
    # uses for a sheet's _CodeName.
    component = wb.VBProject.VBComponents(FORM_NAME)
    component.Properties("Width").Value = FORM_WIDTH
    component.Properties("Height").Value = FORM_HEIGHT

    for progid, name, left, top, width, height, extra in FIELD_CONTROLS:
        control = designer.Controls.Add(progid)
        control.Name = name
        control.Left = left
        control.Top = top
        control.Width = width
        control.Height = height
        for prop, value in extra.items():
            setattr(control, prop, value)

    # _Lists!D2:D5 remains the source of truth for the Type choices, read
    # explicitly (ThisWorkbook-qualified) by frmConnectorEditor's
    # UserForm_Initialize at runtime - not via RowSource, which is an
    # unqualified "_Lists!D2:D5" string that resolves against whichever
    # workbook is ActiveWorkbook when the control binds, not necessarily
    # this one. That bit manual verification (phase-2-manual-verification.md,
    # 2c): frmManageLibrary's Edit flow has ConnectorLibrary.xlsx active by
    # the time frmConnectorEditor loads, so RowSource silently resolved to
    # nothing there.
    lists_sheet = wb.Worksheets("_Lists")
    lists_sheet.Cells(1, TYPE_LIST_COLUMN).Value = "Type"
    for offset, choice in enumerate(TYPE_CHOICES):
        lists_sheet.Cells(offset + 2, TYPE_LIST_COLUMN).Value = choice


PICKER_NAME = "frmConnectorPicker"
PICKER_CONTROLS = [
    ("Forms.ListBox.1", "lstConnectors", 12, 12, 300, 200, {}),
    ("Forms.CommandButton.1", "cmdAdd", 12, 220, 90, 24, {"Caption": "Add"}),
    ("Forms.CommandButton.1", "cmdNew", 110, 220, 90, 24, {"Caption": "New..."}),
    ("Forms.CommandButton.1", "cmdCancel", 222, 220, 90, 24, {"Caption": "Cancel"}),
]

MANAGE_LIBRARY_NAME = "frmManageLibrary"
# 6 buttons (Edit/Delete/Import/Export Connector/Export Library/Close) need
# more width than frmConnectorPicker's 3-button layout - the form itself
# must be widened to match, or the last two buttons render past its right
# edge.
MANAGE_LIBRARY_WIDTH = 550
MANAGE_LIBRARY_CONTROLS = [
    ("Forms.ListBox.1", "lstConnectors", 12, 12, 526, 200, {}),
    ("Forms.CommandButton.1", "cmdEdit", 12, 220, 80, 24, {"Caption": "Edit"}),
    ("Forms.CommandButton.1", "cmdDelete", 96, 220, 80, 24, {"Caption": "Delete"}),
    ("Forms.CommandButton.1", "cmdImport", 180, 220, 80, 24, {"Caption": "Import..."}),
    ("Forms.CommandButton.1", "cmdExport", 264, 220, 80, 24, {"Caption": "Export Connector"}),
    ("Forms.CommandButton.1", "cmdExportLibrary", 348, 220, 100, 24, {"Caption": "Export Library"}),
    ("Forms.CommandButton.1", "cmdClose", 452, 220, 80, 24, {"Caption": "Close"}),
]


def _build_form(wb, add_userform, name, caption, width, height, control_specs):
    designer = add_userform(wb, name)
    designer.Caption = caption
    # Designer.Width/Height aren't settable through win32com's dynamic
    # dispatch (unlike Caption); the VBComponent's Properties collection
    # reaches the same underlying value, same technique
    # build_connector_editor_form and excel_com.set_codename use.
    component = wb.VBProject.VBComponents(name)
    component.Properties("Width").Value = width
    component.Properties("Height").Value = height

    for progid, ctl_name, left, top, ctl_width, ctl_height, extra in control_specs:
        control = designer.Controls.Add(progid)
        control.Name = ctl_name
        control.Left = left
        control.Top = top
        control.Width = ctl_width
        control.Height = ctl_height
        for prop, value in extra.items():
            setattr(control, prop, value)


def build_connector_picker_form(wb, add_userform) -> None:
    # Same fix as MANAGE_LIBRARY: buttons bottom out at Top(220) + Height(24)
    # = 244 - 260 left only a 16px margin, which rendered as buttons cut off
    # at the bottom (phase-2-manual-verification.md, 2c manual verification).
    _build_form(wb, add_userform, PICKER_NAME, "Add Connector", 340, 300, PICKER_CONTROLS)


def build_manage_library_form(wb, add_userform) -> None:
    # cmdEdit/etc. bottom out at Top(220) + Height(24) = 244 - 260 left only
    # a 16px margin below them, which read as "not tall enough" once shown.
    _build_form(wb, add_userform, MANAGE_LIBRARY_NAME, "Manage Library", MANAGE_LIBRARY_WIDTH, 300, MANAGE_LIBRARY_CONTROLS)


REMOVE_CONNECTOR_NAME = "frmRemoveConnector"
REMOVE_CONNECTOR_CONTROLS = [
    ("Forms.ListBox.1", "lstConnectors", 12, 12, 300, 200, {}),
    ("Forms.CommandButton.1", "cmdRemove", 12, 220, 90, 24, {"Caption": "Remove"}),
    ("Forms.CommandButton.1", "cmdCancel", 110, 220, 90, 24, {"Caption": "Cancel"}),
]


def build_remove_connector_form(wb, add_userform) -> None:
    _build_form(wb, add_userform, REMOVE_CONNECTOR_NAME, "Remove Connector", 340, 300, REMOVE_CONNECTOR_CONTROLS)
