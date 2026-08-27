"""Field and control layout for frmConnectorEditor, built via the VBIDE
Designer object model at build time - there is no static .frm source."""
from __future__ import annotations

FORM_NAME = "frmConnectorEditor"
FORM_CAPTION = "Connector Editor"
FORM_WIDTH = 520
FORM_HEIGHT = 420

TYPE_CHOICES = ["Connector", "Stud", "Splice", "Tail"]
TYPE_LIST_COLUMN = 4
TYPE_LIST_COLUMN_LETTER = "D"

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
    ("Forms.ListBox.1", "lstPins", 12, 184, 180, 120, {}),
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

    # ComboBox.AddItem only populates the control's live runtime state - it
    # is not written into the .frx property bag, so the list is empty again
    # after SaveAs/reopen. RowSource bound to a worksheet range is what
    # actually persists, the same mechanism this workbook's other pick lists
    # (_Lists + Harness chart validation) already rely on.
    lists_sheet = wb.Worksheets("_Lists")
    lists_sheet.Cells(1, TYPE_LIST_COLUMN).Value = "Type"
    for offset, choice in enumerate(TYPE_CHOICES):
        lists_sheet.Cells(offset + 2, TYPE_LIST_COLUMN).Value = choice

    combo = designer.Controls("cboType")
    combo.RowSource = "_Lists!{0}2:{0}{1}".format(TYPE_LIST_COLUMN_LETTER, len(TYPE_CHOICES) + 1)
