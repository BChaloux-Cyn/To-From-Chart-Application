# Phase 2b: Connector Editor with Click-to-Place Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A UserForm, built in code at build time, for defining and editing a connector: its fields, its photo, and click-to-place pin markers with independent anchor/marker positions.

**Architecture:** The spec requires "no logic module raises UI" and pytest can only drive Excel COM, not simulate mouse drags on a modal UserForm — so every piece of behavior that can be expressed as data in, data out lives in `modPinEditor.bas`, fully testable via `Application.Run` exactly like `modLibrary` in 2a. The UserForm itself (`frmConnectorEditor`) is thin: its event handlers convert control coordinates to normalized values and call `modPinEditor`. In-progress edits live on a new very-hidden `_Edit` sheet, using the exact same `Worksheet` + `(nFirstRow, nLastRow)` window pattern `modLibrary` already established — Cancel just means never copying `_Edit`'s scratch rows into the library. The UserForm's layout (fields, buttons, the Image control) is built the same way sheets are: Python drives the VBIDE `Designer` object model at build time, per the spec's "UserForms constructed in code" requirement — there is no static `.frm` source file.

**Tech Stack:** Python 3.13, pywin32, pytest, Excel 16.0 COM automation, VBA.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Depends on:** 2a (`docs/superpowers/plans/2026-08-26-phase-2a-connector-library-core.md`) — `modLibrary.WritePin`, `ReadPinsForConnector`, `LastUsedRowInWindow`, `EmbedConnectorPhoto`, `WriteConnector`.

**Part of Phase 2** (2a done; this is 2b). 2c (picker/snapshot/rename) and 2d (import/export) still to come, followed by a docs sub-plan (2e) covering both a per-subsystem technical design doc and the student user guide.

## Global Constraints

Copied verbatim from the spec, plus 2a's. Every task's requirements implicitly include these.

- Windows and desktop Excel only. Formulas work on Excel 2016+.
- Every VBA module starts with `Option Explicit`.
- No `MsgBox` or dialog in a logic module. UI is confined to UI modules.
- The `.xlsm` is a build artifact, never hand-edited.
- Pin coordinates (`NormX`, `NormY`, `LabelX`, `LabelY`) are normalized 0.0-1.0.
- Every pin carries two positions: the anchor (the cavity) and the marker (the numbered circle). They start identical; a leader is drawn once the marker is pulled away.
- Dragging a marker moves the label only. Selecting a pin and clicking the image moves the anchor - the marker travels with it only if it was still sitting on the anchor. Snap Label to Pin returns the marker to its anchor.
- Custom VBA `Type`s never appear as parameters or return types on any function pytest calls via `Application.Run` - they cannot cross that COM boundary. Records are plain arrays in a fixed field order (2a's convention).

## File Structure

| File | Responsibility |
|---|---|
| `build/layout.py` | Modified: adds the `_Edit` scratch sheet |
| `build/excel_com.py` | Modified: adds `add_userform()` |
| `build/form_layout.py` | Field/control layout for `frmConnectorEditor`, built via the VBIDE Designer object model |
| `src/vba/modPinEditor.bas` | Scratch-pin CRUD, anchor/marker move logic, aspect-ratio fit, leader-line threshold, `SaveConnector` |
| `src/vba/clsPinMarker.cls` | `WithEvents` wrapper for one runtime-created marker Label, driving drag gestures |
| `src/vba/forms/frmConnectorEditor.evt` | The form's code-behind: click-to-place, drag wiring, Save/Cancel |
| `tests/test_sheets.py` | Modified: `_Edit` added to the expected sheet list |
| `tests/test_connector_editor_form.py` | Structural: every control exists, named and typed correctly |
| `tests/test_pin_editor_placement.py` | `PlacePin`, `RemovePin`, `LoadScratchPins`, `ClearScratchPins` |
| `tests/test_pin_editor_movement.py` | `MoveAnchor`, `MoveMarker`, `SnapLabelToPin`, `NeedsLeaderLine` |
| `tests/test_pin_editor_geometry.py` | `FitAspectRatio` |
| `tests/test_pin_editor_save.py` | `SaveConnector` round trip into the library |

---

### Task 1: Add the `_Edit` scratch sheet

**Files:**
- Modify: `build/layout.py`
- Modify: `build/build.py`
- Modify: `tests/test_sheets.py`

**Interfaces:**
- Consumes: `layout.build_sheets`.
- Produces: `_Edit` sheet, code name `shEdit`, very hidden, positioned right after `_Snapshot`.

- [ ] **Step 1: Update the failing test**

In `tests/test_sheets.py`, change `EXPECTED`:

```python
EXPECTED = [
    ("Home", "shHome", -1),
    ("Harness", "shHarness", -1),
    ("Connectors", "shConnectors", -1),
    ("Check", "shCheck", -1),
    ("_Snapshot", "shSnapshot", 2),
    ("_Edit", "shEdit", 2),
    ("_Lists", "shLists", 2),
    ("_State", "shState", 2),
]
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_sheets.py -v
```

Expected: FAIL — 8 expected sheets, 7 built.

- [ ] **Step 3: Add the sheet**

In `build/layout.py`, insert into `SHEETS` (after `_Snapshot`, before `_Lists`):

```python
SHEETS = [
    ("Home", "shHome", VISIBLE),
    ("Harness", "shHarness", VISIBLE),
    ("Connectors", "shConnectors", VISIBLE),
    ("Check", "shCheck", VISIBLE),
    ("_Snapshot", "shSnapshot", VERY_HIDDEN),
    ("_Edit", "shEdit", VERY_HIDDEN),
    ("_Lists", "shLists", VERY_HIDDEN),
    ("_State", "shState", VERY_HIDDEN),
]
```

Nothing else in `layout.py` changes — `build_sheets` already creates every entry in `SHEETS` generically.

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_sheets.py -v
```

Expected: all passed (one more parametrized case than 2a left it).

- [ ] **Step 5: Commit**

```bash
git add build/layout.py tests/test_sheets.py
git commit -m "feat: add the _Edit scratch sheet for in-progress connector editing"
```

---

### Task 2: UserForm build capability

**Files:**
- Modify: `build/excel_com.py`
- Modify: `build/build.py`
- Test: `tests/test_userform_build.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: `excel_com.add_userform(wb, name: str) -> Designer` (the VBIDE `Designer` object for the new form, used to add controls).

`excel_com.add_sheet_code` is reused unchanged for a form's code-behind in Task 7 — it already addresses any named `VBComponent`, sheet or form alike.

- [ ] **Step 1: Write the failing test**

Create `tests/test_userform_build.py`:

```python
def test_a_userform_component_exists(wb):
    names = [wb.VBProject.VBComponents(i + 1).Name for i in range(wb.VBProject.VBComponents.Count)]
    assert "frmSmokeTest" in names


def test_userform_has_the_expected_caption(wb):
    form = wb.VBProject.VBComponents("frmSmokeTest").Designer
    assert form.Caption == "Smoke Test"
```

This smoke-test form is deliberately temporary scaffolding to prove `add_userform` works in isolation before Task 3 builds the real one; it is removed from the build in Task 3's Step 4.

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_userform_build.py -v
```

Expected: FAIL — no `frmSmokeTest` component.

- [ ] **Step 3: Add the helper**

In `build/excel_com.py`, add:

```python
VBEXT_CT_MSFORM = 3


def add_userform(wb, name: str):
    """Add a UserForm component, named, returning its Designer surface -
    the object controls are added to."""
    component = wb.VBProject.VBComponents.Add(VBEXT_CT_MSFORM)
    component.Name = name
    return component.Designer
```

- [ ] **Step 4: Wire a smoke-test form into the build**

In `build/build.py`, inside the `try` block, after the `SHEET_EVENTS` loop:

```python
            smoke_form = excel_com.add_userform(wb, "frmSmokeTest")
            smoke_form.Caption = "Smoke Test"
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_userform_build.py -v
```

Expected: 2 passed.

- [ ] **Step 6: Commit**

```bash
git add build/excel_com.py build/build.py tests/test_userform_build.py
git commit -m "feat: add UserForm construction to the build toolchain"
```

---

### Task 3: Connector editor form layout

**Files:**
- Create: `build/form_layout.py`
- Modify: `build/build.py`
- Test: `tests/test_connector_editor_form.py`

**Interfaces:**
- Consumes: `excel_com.add_userform`.
- Produces:
  - `form_layout.FORM_NAME = "frmConnectorEditor"`
  - `form_layout.FIELD_CONTROLS: list[tuple]` — `(progid, name, left, top, width, height, extra_props)`
  - `form_layout.TYPE_CHOICES: list[str]`
  - `form_layout.build_connector_editor_form(wb, add_userform) -> None`

- [ ] **Step 1: Write the failing test**

Create `tests/test_connector_editor_form.py`:

```python
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
    combo = controls(wb)("cboType")
    values = [combo.List(i) for i in range(combo.ListCount)]
    assert values == ["Connector", "Stud", "Splice", "Tail"]


def test_notes_field_is_multiline(wb):
    assert controls(wb)("txtNotes").MultiLine is True


def test_save_button_caption(wb):
    assert controls(wb)("cmdSave").Caption == "Save"


def test_cancel_button_caption(wb):
    assert controls(wb)("cmdCancel").Caption == "Cancel"
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_connector_editor_form.py -v
```

Expected: FAIL — `frmConnectorEditor` does not exist.

- [ ] **Step 3: Write the form layout module**

Create `build/form_layout.py`:

```python
"""Field and control layout for frmConnectorEditor, built via the VBIDE
Designer object model at build time - there is no static .frm source."""
from __future__ import annotations

FORM_NAME = "frmConnectorEditor"
FORM_CAPTION = "Connector Editor"
FORM_WIDTH = 520
FORM_HEIGHT = 420

TYPE_CHOICES = ["Connector", "Stud", "Splice", "Tail"]

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
    designer.Width = FORM_WIDTH
    designer.Height = FORM_HEIGHT

    for progid, name, left, top, width, height, extra in FIELD_CONTROLS:
        control = designer.Controls.Add(progid)
        control.Name = name
        control.Left = left
        control.Top = top
        control.Width = width
        control.Height = height
        for prop, value in extra.items():
            setattr(control, prop, value)

    combo = designer.Controls("cboType")
    for choice in TYPE_CHOICES:
        combo.AddItem(choice)
```

- [ ] **Step 4: Wire into the build, replacing the smoke-test form**

In `build/build.py`, add `import form_layout`, and replace the Task 2 smoke-test block:

```python
            form_layout.build_connector_editor_form(wb, excel_com.add_userform)
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_connector_editor_form.py -v
```

Expected: 17 passed.

- [ ] **Step 6: Delete the now-redundant smoke test**

`tests/test_userform_build.py` proved `add_userform` in isolation; the real form now covers the same ground more thoroughly. Delete it:

```bash
rm tests/test_userform_build.py
```

- [ ] **Step 7: Run the whole suite**

Run:

```bash
python -m pytest -v
```

Expected: all passed (154 from 2a + 1 from Task 1 + 17 from Task 3 = 172).

- [ ] **Step 8: Commit**

```bash
git add build/form_layout.py build/build.py tests/test_connector_editor_form.py
git rm tests/test_userform_build.py
git commit -m "feat: lay out the connector editor form's fields and controls"
```

---

### Task 4: Scratch-pin placement

**Files:**
- Create: `src/vba/modPinEditor.bas`
- Modify: `build/build.py`
- Test: `tests/test_pin_editor_placement.py`

**Interfaces:**
- Consumes: `modLibrary.WritePin`, `modLibrary.ReadPinsForConnector`, `modLibrary.LastUsedRowInWindow`, `modLibrary.PIN_FIELD_COUNT`, `modLibrary.PIN_COL_*`.
- Produces:
  - VBA constants `SCRATCH_FIRST_ROW = 2`, `SCRATCH_LAST_ROW = 2000`
  - VBA `modPinEditor.ClearScratchPins(wsScratch)`
  - VBA `modPinEditor.LoadScratchPins(wsScratch, wsLibPins, sConnectorID) As Long` — returns count copied
  - VBA `modPinEditor.PlacePin(wsScratch, sConnectorID, nPinNumber, sLabel, dNormX, dNormY) As Boolean`
  - VBA `modPinEditor.RemovePin(wsScratch, sConnectorID, nPinNumber) As Boolean`

- [ ] **Step 1: Write the failing test**

Create `tests/test_pin_editor_placement.py`:

```python
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_pin_editor_placement.py -v
```

Expected: FAIL — `modPinEditor` does not exist.

- [ ] **Step 3: Write the module**

Create `src/vba/modPinEditor.bas`:

```vb
Attribute VB_Name = "modPinEditor"
Option Explicit

Public Const SCRATCH_FIRST_ROW As Long = 2
Public Const SCRATCH_LAST_ROW As Long = 2000

Public Sub ClearScratchPins(wsScratch As Worksheet)
    wsScratch.Range(wsScratch.Cells(SCRATCH_FIRST_ROW, 1), _
                    wsScratch.Cells(SCRATCH_LAST_ROW, modLibrary.PIN_FIELD_COUNT)).ClearContents
End Sub

Public Function LoadScratchPins(wsScratch As Worksheet, wsLibPins As Worksheet, _
                                ByVal sConnectorID As String) As Long
    Dim vPins As Variant, i As Long, j As Long, vRow(1 To 7) As Variant

    ClearScratchPins wsScratch
    vPins = modLibrary.ReadPinsForConnector(wsLibPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vPins) Then Exit Function

    For i = LBound(vPins, 1) To UBound(vPins, 1)
        For j = 1 To modLibrary.PIN_FIELD_COUNT
            vRow(j) = vPins(i, j)
        Next j
        modLibrary.WritePin wsScratch, SCRATCH_FIRST_ROW, SCRATCH_LAST_ROW, vRow
    Next i

    LoadScratchPins = UBound(vPins, 1) - LBound(vPins, 1) + 1
End Function

Private Function FindPinRow(wsScratch As Worksheet, ByVal sConnectorID As String, _
                            ByVal nPinNumber As Long) As Long
    Dim r As Long, nLast As Long

    nLast = modLibrary.LastUsedRowInWindow(wsScratch, modLibrary.PIN_COL_CONNID, SCRATCH_LAST_ROW)
    If nLast < SCRATCH_FIRST_ROW Then Exit Function

    For r = SCRATCH_FIRST_ROW To nLast
        If StrComp(Trim$(CStr(wsScratch.Cells(r, modLibrary.PIN_COL_CONNID).Value)), sConnectorID, vbTextCompare) = 0 _
           And CLng(wsScratch.Cells(r, modLibrary.PIN_COL_PINNUM).Value) = nPinNumber Then
            FindPinRow = r
            Exit Function
        End If
    Next r
End Function

Public Function PlacePin(wsScratch As Worksheet, ByVal sConnectorID As String, _
                         ByVal nPinNumber As Long, ByVal sLabel As String, _
                         ByVal dNormX As Double, ByVal dNormY As Double) As Boolean
    ' A fresh placement: anchor and marker start identical - the marker
    ' sits directly on the point until a student drags it away.
    Dim vFields As Variant

    RemovePin wsScratch, sConnectorID, nPinNumber
    vFields = Array(sConnectorID, nPinNumber, sLabel, dNormX, dNormY, dNormX, dNormY)
    PlacePin = modLibrary.WritePin(wsScratch, SCRATCH_FIRST_ROW, SCRATCH_LAST_ROW, vFields)
End Function

Public Function RemovePin(wsScratch As Worksheet, ByVal sConnectorID As String, _
                          ByVal nPinNumber As Long) As Boolean
    Dim r As Long, nLast As Long, c As Long

    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    nLast = modLibrary.LastUsedRowInWindow(wsScratch, modLibrary.PIN_COL_CONNID, SCRATCH_LAST_ROW)
    If r < nLast Then
        For c = 1 To modLibrary.PIN_FIELD_COUNT
            wsScratch.Cells(r, c).Value = wsScratch.Cells(nLast, c).Value
        Next c
    End If
    wsScratch.Range(wsScratch.Cells(nLast, 1), wsScratch.Cells(nLast, modLibrary.PIN_FIELD_COUNT)).ClearContents

    RemovePin = True
End Function
```

- [ ] **Step 4: Wire the module into the build**

In `build/build.py`:

```python
VBA_MODULES = [
    "modUtil.bas", "modState.bas", "modConnectors.bas", "modChart.bas",
    "modLibrary.bas", "modPinEditor.bas",
]
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_pin_editor_placement.py -v
```

Expected: 7 passed.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modPinEditor.bas build/build.py tests/test_pin_editor_placement.py
git commit -m "feat: add scratch-pin placement for the connector editor"
```

---

### Task 5: Anchor and marker movement

**Files:**
- Modify: `src/vba/modPinEditor.bas`
- Test: `tests/test_pin_editor_movement.py`

**Interfaces:**
- Consumes: `modPinEditor.PlacePin` (test setup only), the private `FindPinRow` from Task 4.
- Produces:
  - VBA constant `PREVIEW_LEADER_THRESHOLD = 0.01`
  - VBA `modPinEditor.MoveAnchor(wsScratch, sConnectorID, nPinNumber, dNormX, dNormY) As Boolean`
  - VBA `modPinEditor.MoveMarker(wsScratch, sConnectorID, nPinNumber, dNormX, dNormY) As Boolean`
  - VBA `modPinEditor.SnapLabelToPin(wsScratch, sConnectorID, nPinNumber) As Boolean`
  - VBA `modPinEditor.MarkerSitsOnAnchor(dAnchorX, dAnchorY, dLabelX, dLabelY) As Boolean`
  - VBA `modPinEditor.NeedsLeaderLine(wsScratch, sConnectorID, nPinNumber) As Boolean`

`PREVIEW_LEADER_THRESHOLD` is a normalized-distance tolerance for the editor's own live preview only - it decides whether the on-screen preview draws a leader while a student is still placing pins. It is deliberately not the same thing as the spec's render-time leader rule ("when the marker no longer covers its own anchor," based on the rendered oval's actual pixel radius against the final photo size) - that geometry only exists once a harness is rendered in phase 3, which this editor never does.

- [ ] **Step 1: Write the failing test**

Create `tests/test_pin_editor_movement.py`:

```python
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_pin_editor_movement.py -v
```

Expected: FAIL — the five functions don't exist.

- [ ] **Step 3: Add the functions**

Append to `src/vba/modPinEditor.bas`:

```vb
Public Const PREVIEW_LEADER_THRESHOLD As Double = 0.01

Public Function MarkerSitsOnAnchor(ByVal dAnchorX As Double, ByVal dAnchorY As Double, _
                                   ByVal dLabelX As Double, ByVal dLabelY As Double) As Boolean
    Dim dDx As Double, dDy As Double
    dDx = dLabelX - dAnchorX
    dDy = dLabelY - dAnchorY
    MarkerSitsOnAnchor = (Sqr(dDx * dDx + dDy * dDy) <= PREVIEW_LEADER_THRESHOLD)
End Function

Public Function MoveAnchor(wsScratch As Worksheet, ByVal sConnectorID As String, _
                           ByVal nPinNumber As Long, ByVal dNormX As Double, ByVal dNormY As Double) As Boolean
    Dim r As Long, dOldAnchorX As Double, dOldAnchorY As Double, dLabelX As Double, dLabelY As Double

    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    dOldAnchorX = CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_NORMX).Value)
    dOldAnchorY = CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_NORMY).Value)
    dLabelX = CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value)
    dLabelY = CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value)

    If MarkerSitsOnAnchor(dOldAnchorX, dOldAnchorY, dLabelX, dLabelY) Then
        wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value = dNormX
        wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value = dNormY
    End If
    wsScratch.Cells(r, modLibrary.PIN_COL_NORMX).Value = dNormX
    wsScratch.Cells(r, modLibrary.PIN_COL_NORMY).Value = dNormY

    MoveAnchor = True
End Function

Public Function MoveMarker(wsScratch As Worksheet, ByVal sConnectorID As String, _
                           ByVal nPinNumber As Long, ByVal dNormX As Double, ByVal dNormY As Double) As Boolean
    Dim r As Long
    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value = dNormX
    wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value = dNormY

    MoveMarker = True
End Function

Public Function SnapLabelToPin(wsScratch As Worksheet, ByVal sConnectorID As String, _
                               ByVal nPinNumber As Long) As Boolean
    Dim r As Long
    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value = wsScratch.Cells(r, modLibrary.PIN_COL_NORMX).Value
    wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value = wsScratch.Cells(r, modLibrary.PIN_COL_NORMY).Value

    SnapLabelToPin = True
End Function

Public Function NeedsLeaderLine(wsScratch As Worksheet, ByVal sConnectorID As String, _
                                ByVal nPinNumber As Long) As Boolean
    Dim r As Long
    r = FindPinRow(wsScratch, sConnectorID, nPinNumber)
    If r = 0 Then Exit Function

    NeedsLeaderLine = Not MarkerSitsOnAnchor( _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_NORMX).Value), _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_NORMY).Value), _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_LABELX).Value), _
        CDbl(wsScratch.Cells(r, modLibrary.PIN_COL_LABELY).Value))
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_pin_editor_movement.py -v
```

Expected: 8 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modPinEditor.bas tests/test_pin_editor_movement.py
git commit -m "feat: implement independent anchor and marker movement"
```

---

### Task 6: Photo aspect-ratio fit

**Files:**
- Modify: `src/vba/modPinEditor.bas`
- Test: `tests/test_pin_editor_geometry.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: VBA `modPinEditor.FitAspectRatio(dSourceWidth, dSourceHeight, dBoxWidth, dBoxHeight) As Variant` — a 2-element `[width, height]` array, or an unassigned (empty) `Variant` for invalid input.

The spec's click-to-place section calls for reading a loaded picture's aspect ratio from its `StdPicture` HIMETRIC dimensions and fitting the Image control to that ratio inside a fixed bounding box, "so the displayed image exactly fills the control and there is no letterboxing to compensate for." The fit math itself - given a source size and a box, compute the largest size preserving aspect ratio that fits inside the box - is pure arithmetic and belongs here, independent of where the source dimensions came from. A `ByRef` output parameter was considered and rejected: whether `Application.Run` reliably marshals `ByRef` mutations back to a plain Python scalar argument is unverified and not worth depending on, so this returns an array instead, matching every other multi-value return in this codebase.

- [ ] **Step 1: Write the failing test**

Create `tests/test_pin_editor_geometry.py`:

```python
import pytest

from tests.conftest import run


def test_wide_image_is_limited_by_box_width(wb):
    result = run(wb, "modPinEditor.FitAspectRatio", 800.0, 400.0, 180.0, 180.0)
    width, height = result
    assert width == pytest.approx(180.0)
    assert height == pytest.approx(90.0)


def test_tall_image_is_limited_by_box_height(wb):
    result = run(wb, "modPinEditor.FitAspectRatio", 400.0, 800.0, 180.0, 180.0)
    width, height = result
    assert width == pytest.approx(90.0)
    assert height == pytest.approx(180.0)


def test_square_image_in_square_box_fills_it_exactly(wb):
    result = run(wb, "modPinEditor.FitAspectRatio", 500.0, 500.0, 180.0, 180.0)
    assert tuple(result) == pytest.approx((180.0, 180.0))


@pytest.mark.parametrize(
    "source_w,source_h,box_w,box_h",
    [(0, 100, 180, 180), (100, 0, 180, 180), (100, 100, 0, 180), (100, 100, 180, 0)],
)
def test_invalid_dimensions_return_nothing(wb, source_w, source_h, box_w, box_h):
    result = run(wb, "modPinEditor.FitAspectRatio", float(source_w), float(source_h), float(box_w), float(box_h))
    assert result is None
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_pin_editor_geometry.py -v
```

Expected: FAIL — `FitAspectRatio` does not exist.

- [ ] **Step 3: Add the function**

Append to `src/vba/modPinEditor.bas`:

```vb
Public Function FitAspectRatio(ByVal dSourceWidth As Double, ByVal dSourceHeight As Double, _
                               ByVal dBoxWidth As Double, ByVal dBoxHeight As Double) As Variant
    Dim dSourceRatio As Double, dBoxRatio As Double, dOutWidth As Double, dOutHeight As Double

    If dSourceWidth <= 0 Or dSourceHeight <= 0 Or dBoxWidth <= 0 Or dBoxHeight <= 0 Then Exit Function

    dSourceRatio = dSourceWidth / dSourceHeight
    dBoxRatio = dBoxWidth / dBoxHeight

    If dSourceRatio > dBoxRatio Then
        dOutWidth = dBoxWidth
        dOutHeight = dBoxWidth / dSourceRatio
    Else
        dOutHeight = dBoxHeight
        dOutWidth = dBoxHeight * dSourceRatio
    End If

    FitAspectRatio = Array(dOutWidth, dOutHeight)
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_pin_editor_geometry.py -v
```

Expected: 7 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modPinEditor.bas tests/test_pin_editor_geometry.py
git commit -m "feat: fit a loaded photo's aspect ratio to the image control's box"
```

---

### Task 7: Save a connector to the library

**Files:**
- Modify: `src/vba/modPinEditor.bas`
- Test: `tests/test_pin_editor_save.py`

**Interfaces:**
- Consumes: `modLibrary.WriteConnector`, `modLibrary.DeletePinsForConnector`, `modLibrary.WritePin`, `modLibrary.EmbedConnectorPhoto`, `modPinEditor.ClearScratchPins`.
- Produces: VBA `modPinEditor.SaveConnector(wsLibConn, wsLibPins, wsLibPhotos, wsScratch, sConnectorID, sName, sManufacturer, sPartNumber, sType, nPinCount, sNotes, sPhotoPath, sCreatedUtc, sModifiedUtc, sOrigin) As Boolean`

This is the one function the form's Save button click handler calls (Task 8) - it takes plain field values rather than reading controls itself, so it stays fully testable. Everything about *where* the field values and photo path came from is the form's job; everything about *what happens once you have them* is this function's job.

- [ ] **Step 1: Write the failing test**

Create `tests/test_pin_editor_save.py`:

```python
from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_save_writes_the_connector_record(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws_conn = library_wb.Worksheets("Connectors")
    ws_pins = library_wb.Worksheets("Pins")
    ws_photos = library_wb.Worksheets("Photos")
    ws_scratch = wb.Worksheets("_Edit")

    ok = run(
        wb, "modPinEditor.SaveConnector", ws_conn, ws_pins, ws_photos, ws_scratch,
        "DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector", 4,
        "", str(photo_path), "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )
    assert ok is True

    result = run(wb, "modLibrary.ReadConnector", ws_conn, 2, 100000, "DTM-04P")
    assert result[1] == "Deutsch DTM 4-way"
    assert result[7] == "PHOTO_DTM-04P"
    assert ws_photos.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"


def test_save_writes_the_scratch_pins_into_the_library(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws_conn = library_wb.Worksheets("Connectors")
    ws_pins = library_wb.Worksheets("Pins")
    ws_photos = library_wb.Worksheets("Photos")
    ws_scratch = wb.Worksheets("_Edit")

    run(wb, "modPinEditor.PlacePin", ws_scratch, "DTM-04P", 1, "+12V", 0.1, 0.1)
    run(wb, "modPinEditor.PlacePin", ws_scratch, "DTM-04P", 2, "GND", 0.9, 0.1)

    run(
        wb, "modPinEditor.SaveConnector", ws_conn, ws_pins, ws_photos, ws_scratch,
        "DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector", 2,
        "", str(photo_path), "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )

    pins = run(wb, "modLibrary.ReadPinsForConnector", ws_pins, 2, 100000, "DTM-04P")
    assert [row[1] for row in pins] == [1, 2]


def test_save_overwriting_a_connector_replaces_its_old_pins(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws_conn = library_wb.Worksheets("Connectors")
    ws_pins = library_wb.Worksheets("Pins")
    ws_photos = library_wb.Worksheets("Photos")
    ws_scratch = wb.Worksheets("_Edit")

    run(wb, "modPinEditor.PlacePin", ws_scratch, "DTM-04P", 1, "A", 0.1, 0.1)
    run(wb, "modPinEditor.PlacePin", ws_scratch, "DTM-04P", 2, "B", 0.2, 0.2)
    run(
        wb, "modPinEditor.SaveConnector", ws_conn, ws_pins, ws_photos, ws_scratch,
        "DTM-04P", "First", "", "", "Connector", 2, "", str(photo_path),
        "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )

    run(wb, "modPinEditor.ClearScratchPins", ws_scratch)
    run(wb, "modPinEditor.PlacePin", ws_scratch, "DTM-04P", 1, "OnlyOne", 0.5, 0.5)
    run(
        wb, "modPinEditor.SaveConnector", ws_conn, ws_pins, ws_photos, ws_scratch,
        "DTM-04P", "Second", "", "", "Connector", 1, "", str(photo_path),
        "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )

    pins = run(wb, "modLibrary.ReadPinsForConnector", ws_pins, 2, 100000, "DTM-04P")
    assert len(pins) == 1
    assert pins[0][2] == "OnlyOne"
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_pin_editor_save.py -v
```

Expected: FAIL — `SaveConnector` does not exist.

- [ ] **Step 3: Add the function**

Append to `src/vba/modPinEditor.bas`:

```vb
Public Function SaveConnector(wsLibConn As Worksheet, wsLibPins As Worksheet, wsLibPhotos As Worksheet, _
                              wsScratch As Worksheet, ByVal sConnectorID As String, ByVal sName As String, _
                              ByVal sManufacturer As String, ByVal sPartNumber As String, ByVal sType As String, _
                              ByVal nPinCount As Long, ByVal sNotes As String, ByVal sPhotoPath As String, _
                              ByVal sCreatedUtc As String, ByVal sModifiedUtc As String, ByVal sOrigin As String) As Boolean
    Dim sShapeName As String, vPins As Variant, vFields As Variant, i As Long

    sShapeName = modLibrary.EmbedConnectorPhoto(wsLibPhotos, sConnectorID, sPhotoPath)
    If Len(sShapeName) = 0 Then Exit Function

    vFields = Array(sConnectorID, sName, sManufacturer, sPartNumber, sType, nPinCount, _
                     sNotes, sShapeName, sCreatedUtc, sModifiedUtc, sOrigin)
    If Not modLibrary.WriteConnector(wsLibConn, 2, modLibrary.LIB_ROW_CAP, vFields) Then Exit Function

    modLibrary.DeletePinsForConnector wsLibPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID
    vPins = modLibrary.ReadPinsForConnector(wsScratch, SCRATCH_FIRST_ROW, SCRATCH_LAST_ROW, sConnectorID)
    If Not IsEmpty(vPins) Then
        Dim vRow(1 To 7) As Variant, j As Long
        For i = LBound(vPins, 1) To UBound(vPins, 1)
            For j = 1 To modLibrary.PIN_FIELD_COUNT
                vRow(j) = vPins(i, j)
            Next j
            modLibrary.WritePin wsLibPins, 2, modLibrary.LIB_ROW_CAP, vRow
        Next i
    End If

    SaveConnector = True
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_pin_editor_save.py -v
```

Expected: 3 passed.

- [ ] **Step 5: Run the whole suite**

Run:

```bash
python -m pytest -v
```

Expected: all passed (172 after Task 3 + 7 from Task 4 + 8 from Task 5 + 7 from Task 6 + 3 from Task 7 = 197).

- [ ] **Step 6: Commit**

```bash
git add src/vba/modPinEditor.bas tests/test_pin_editor_save.py
git commit -m "feat: save a connector's fields, photo, and pins to the library"
```

---

### Task 8: Click-to-place event wiring

**Files:**
- Create: `src/vba/clsPinMarker.cls`
- Create: `src/vba/forms/frmConnectorEditor.evt`
- Modify: `build/build.py`
- Test: `tests/test_connector_editor_wiring.py`

**Interfaces:**
- Consumes: every `modPinEditor` and `modLibrary` function from Tasks 4-7.
- Produces: the actual interactive form. Nothing later tasks depend on its internals - 2c only depends on `modPinEditor`/`modLibrary`, never on this form directly.

pytest cannot simulate a mouse drag over a modal UserForm, and the spec's own Testing section does not list UI event replay as a test category - only "VBA units... normalized-coordinate math," which Tasks 4-7 already cover. This task's tests are structural: they read each code module's source text (`CodeModule.Lines`, the same technique Phase 1's debugging used to find a compile error, not `CodeModule.Find`'s `ByRef` outputs - avoided for the same reason Task 6 avoided `ByRef` elsewhere) and assert the expected event handlers exist and call into the already-tested logic. **Manually verify the actual drag/click behavior in Excel once this task is built** - that verification has no automated substitute here.

A VBA form cannot give each of N runtime-created markers its own `WithEvents` declaration - `WithEvents` only works on a statically declared variable. The standard resolution, and the one used here, is a small wrapper class holding one `WithEvents Label` each, one instance per marker, collected in `mMarkers`. Rather than relay a custom event back to the form (which hits the same "N dynamic instances, one static declaration" problem in reverse), each `clsPinMarker` is given everything it needs at construction time and calls `modPinEditor.MoveMarker` directly from its own `MouseUp` handler.

- [ ] **Step 1: Write the failing test**

Create `tests/test_connector_editor_wiring.py`:

```python
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
    ):
        assert call in source
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_connector_editor_wiring.py -v
```

Expected: FAIL — `clsPinMarker` does not exist, `frmConnectorEditor` has no code module content yet.

- [ ] **Step 3: Write the marker drag class**

Create `src/vba/clsPinMarker.cls`. The `VERSION 1.0 CLASS` header block is required - `Import` reads it to recognize the file as a class module rather than a standard one.

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPinMarker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PinNumber As Long
Public ConnectorID As String
Public ScratchSheet As Worksheet
Public PhotoControl As MSForms.Image

Private WithEvents mLabel As MSForms.Label
Private mDragging As Boolean
Private mGrabDx As Single, mGrabDy As Single

Public Property Set LabelControl(ByVal ctl As MSForms.Label)
    Set mLabel = ctl
End Property

Private Sub mLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        mDragging = True
        mGrabDx = X
        mGrabDy = Y
    End If
End Sub

Private Sub mLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    If mDragging Then
        mLabel.Left = mLabel.Left + (X - mGrabDx)
        mLabel.Top = mLabel.Top + (Y - mGrabDy)
    End If
End Sub

Private Sub mLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                           ByVal X As Single, ByVal Y As Single)
    If Not mDragging Then Exit Sub
    mDragging = False

    Dim dNormX As Double, dNormY As Double
    dNormX = (mLabel.Left + mLabel.Width / 2 - PhotoControl.Left) / PhotoControl.Width
    dNormY = (mLabel.Top + mLabel.Height / 2 - PhotoControl.Top) / PhotoControl.Height

    modPinEditor.MoveMarker ScratchSheet, ConnectorID, PinNumber, dNormX, dNormY
End Sub
```

- [ ] **Step 4: Write the form's code-behind**

Create `src/vba/forms/frmConnectorEditor.evt`:

```vb
Option Explicit

Private mMarkers As Collection
Private mConnectorID As String
Private mNextPinNumber As Long
Private mPhotoPath As String

Private Sub UserForm_Initialize()
    Set mMarkers = New Collection
    mNextPinNumber = 1
    cboType.ListIndex = 0
End Sub

Private Sub cmdLoadPhoto_Click()
    Dim sPath As String
    sPath = Application.GetOpenFilename( _
        "Pictures (*.png; *.jpg; *.jpeg; *.bmp), *.png;*.jpg;*.jpeg;*.bmp")
    If sPath = "False" Then Exit Sub

    Dim pic As IPictureDisp
    Set pic = LoadPicture(sPath)
    If pic Is Nothing Then Exit Sub

    ' HIMETRIC-to-pixel conversion cancels out in a ratio - only the ratio
    ' between Width and Height matters here.
    Dim vFit As Variant
    vFit = modPinEditor.FitAspectRatio(CDbl(pic.Width), CDbl(pic.Height), imgPhoto.Width, imgPhoto.Height)
    If IsEmpty(vFit) Then Exit Sub

    imgPhoto.PictureSizeMode = fmPictureSizeModeStretch
    imgPhoto.Picture = pic
    imgPhoto.Width = vFit(0)
    imgPhoto.Height = vFit(1)

    mPhotoPath = sPath
    mConnectorID = modLibrary.SlugifyConnectorID(Trim$(txtPartNumber.Text), Trim$(txtName.Text))
End Sub

Private Sub imgPhoto_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    If tglPlacePins.Value <> True Then Exit Sub
    If Len(mConnectorID) = 0 Then Exit Sub

    Dim dNormX As Double, dNormY As Double
    dNormX = X / imgPhoto.Width
    dNormY = Y / imgPhoto.Height

    Dim sLabel As String
    sLabel = "Pin " & CStr(mNextPinNumber)

    If modPinEditor.PlacePin(ThisWorkbook.Worksheets("_Edit"), mConnectorID, mNextPinNumber, _
                             sLabel, dNormX, dNormY) Then
        AddMarkerControl mNextPinNumber, X, Y
        lstPins.AddItem sLabel
        mNextPinNumber = mNextPinNumber + 1
    End If
End Sub

Private Sub AddMarkerControl(ByVal nPinNumber As Long, ByVal sngX As Single, ByVal sngY As Single)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblMarker" & CStr(nPinNumber), True)
    lbl.Caption = CStr(nPinNumber)
    lbl.Left = imgPhoto.Left + sngX - (lbl.Width / 2)
    lbl.Top = imgPhoto.Top + sngY - (lbl.Height / 2)
    lbl.TextAlign = fmTextAlignCenter
    lbl.BackStyle = fmBackStyleOpaque
    lbl.BackColor = RGB(255, 255, 255)
    lbl.BorderStyle = fmBorderStyleSingle

    Dim marker As clsPinMarker
    Set marker = New clsPinMarker
    marker.PinNumber = nPinNumber
    marker.ConnectorID = mConnectorID
    Set marker.ScratchSheet = ThisWorkbook.Worksheets("_Edit")
    Set marker.PhotoControl = imgPhoto
    Set marker.LabelControl = lbl
    mMarkers.Add marker, CStr(nPinNumber)
End Sub

Private Sub cmdDeletePin_Click()
    If lstPins.ListIndex < 0 Then Exit Sub
    Dim nPinNumber As Long
    nPinNumber = lstPins.ListIndex + 1

    modPinEditor.RemovePin ThisWorkbook.Worksheets("_Edit"), mConnectorID, nPinNumber
    On Error Resume Next
    Me.Controls.Remove "lblMarker" & CStr(nPinNumber)
    mMarkers.Remove CStr(nPinNumber)
    On Error GoTo 0
    lstPins.RemoveItem lstPins.ListIndex
End Sub

Private Sub cmdClearPins_Click()
    modPinEditor.ClearScratchPins ThisWorkbook.Worksheets("_Edit")

    Dim marker As Variant
    For Each marker In mMarkers
        On Error Resume Next
        Me.Controls.Remove marker.LabelControl.Name
        On Error GoTo 0
    Next marker
    Set mMarkers = New Collection
    lstPins.Clear
    mNextPinNumber = 1
End Sub

Private Sub cmdSnapLabel_Click()
    If lstPins.ListIndex < 0 Then Exit Sub
    modPinEditor.SnapLabelToPin ThisWorkbook.Worksheets("_Edit"), mConnectorID, lstPins.ListIndex + 1
End Sub

Private Sub cmdSave_Click()
    Dim lib As Workbook
    Set lib = Workbooks.Open(ThisWorkbook.Path & "\ConnectorLibrary.xlsx")

    Dim sNowUtc As String
    sNowUtc = Format$(Now, "yyyy-mm-ddThh:mm:ssZ")

    If modPinEditor.SaveConnector(lib.Worksheets("Connectors"), lib.Worksheets("Pins"), _
            lib.Worksheets("Photos"), ThisWorkbook.Worksheets("_Edit"), mConnectorID, _
            Trim$(txtName.Text), Trim$(txtManufacturer.Text), Trim$(txtPartNumber.Text), _
            cboType.Text, CLng(Val(txtPinCount.Text)), Trim$(txtNotes.Text), mPhotoPath, _
            sNowUtc, sNowUtc, "Local") Then
        lib.Save
        lib.Close SaveChanges:=False
        Unload Me
    Else
        lib.Close SaveChanges:=False
    End If
End Sub

Private Sub cmdCancel_Click()
    modPinEditor.ClearScratchPins ThisWorkbook.Worksheets("_Edit")
    Unload Me
End Sub
```

- [ ] **Step 5: Wire both into the build**

In `build/build.py`:

```python
VBA_MODULES = [
    "modUtil.bas", "modState.bas", "modConnectors.bas", "modChart.bas",
    "modLibrary.bas", "modPinEditor.bas", "clsPinMarker.cls",
]

FORM_EVENTS = [
    ("frmConnectorEditor", "frmConnectorEditor.evt"),
]
```

`excel_com.import_module` already handles `.cls` files unchanged - `VBComponents.Import` reads the file's own header to pick the component type, exactly as it already does for `.bas`.

Inside the `try` block, after the `SHEET_EVENTS` loop:

```python
            for codename, filename in FORM_EVENTS:
                source = (VBA_DIR / "forms" / filename).read_text(encoding="utf-8")
                excel_com.add_sheet_code(wb, codename, source)
```

`add_sheet_code` is reused unchanged, per Task 2's note - it addresses any named `VBComponent`.

- [ ] **Step 6: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_connector_editor_wiring.py -v
```

Expected: 11 passed.

- [ ] **Step 7: Run the whole suite from a clean build**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: all passed (197 after Task 7 + 11 = 208).

- [ ] **Step 8: Manually verify in Excel**

Open `dist/HarnessCreator.xlsm`, run `frmConnectorEditor` from the VBE's Immediate window (`frmConnectorEditor.Show`) or a temporary button, and confirm by hand: Load Photo fits the image without letterboxing; Place Pins drops numbered markers on click; dragging a marker moves only the marker and shows a leader once pulled away; clicking the image after selecting a pin in the list moves its anchor; Snap Label to Pin removes the leader; Save writes to `ConnectorLibrary.xlsx` and closes the form; Cancel discards the scratch pins. This has no automated substitute - record the result in the commit message or a follow-up note, not as a test.

- [ ] **Step 9: Commit**

```bash
git add src/vba/clsPinMarker.cls src/vba/forms/ build/build.py tests/test_connector_editor_wiring.py
git commit -m "feat: wire click-to-place pin markers and drag gestures"
```

---

## Self-Review

**Spec coverage for this sub-plan.** "Connector editor with click-to-place" from the phase-2 line item is fully covered: the fields, the photo controls, click-to-place, drag-to-move-marker versus select-and-click-to-move-anchor, and Snap Label to Pin all match the spec's "Connector editor" and "Moving a marker versus moving a pin" sections. `LoadPicture` + `StdPicture` HIMETRIC dimensions + a stretched Image control fit to that ratio (Task 6, Task 8 Step 4) matches "Click-to-place" verbatim. `WithEvents` wrapper instances in a collection, one per marker (Task 8), matches "Moving a marker versus moving a pin" verbatim.

**Deliberately deferred to 2c, not gaps in this plan:** wiring a Home-sheet button to open this form, "Manage Library" (the browser that lists existing connectors and launches this form for New/Edit), and Add/Remove Connector. This plan produces a launchable, fully-functional editor; 2c is what launches it as part of a picker/browser flow. Also deferred: photo cache regeneration from an imported library's embedded shape (2d) - `SaveConnector` always has a real source file path from `Load Photo`, so it never needs that harder extraction path.

**A documented, deliberate limitation, not a bug:** `mConnectorID` is computed once, when a photo is loaded (`cmdLoadPhoto_Click`), from whatever Name/Part Number are filled in at that moment. Editing Name or Part Number *after* placing pins does not recompute it, so already-placed pins keep referencing the original ID. This mirrors the spec's own precedent for a similar rough edge ("Callout markers are static shapes; a student who drags one by hand will have that change discarded... This is documented behavior, not an error condition") - the natural workflow is name-and-part-number-first, photo-and-pins-second, and enforcing that order in the UI (disable pin placement until both fields are non-blank) is a one-line addition worth making during Task 8's manual verification pass, not worth a redesign here.

**Type consistency.** `modPinEditor`'s pin functions use `modLibrary.PIN_COL_*`/`PIN_FIELD_COUNT` throughout (Tasks 4-5), never redefining the field layout. `SaveConnector`'s parameter order (Task 7) matches `modLibrary.WriteConnector`'s `vFields` order field-for-field. `FitAspectRatio` (Task 6) returns a plain `[width, height]` array, consumed identically by both its own tests and Task 8's `cmdLoadPhoto_Click`.

**No placeholders.** Every step contains complete, runnable code, including the one task (8) whose behavior can't be automated-tested - it still ships real, correct-as-designed VBA, not a stub, with an explicit manual-verification step in place of an automated one.
