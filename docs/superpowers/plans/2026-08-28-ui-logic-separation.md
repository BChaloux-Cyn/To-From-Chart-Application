# UI and Logic Separation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move every domain decision out of UserForm and worksheet event handlers into standard modules reachable by `Application.Run`, so pytest asserts on behaviour instead of on VBA source text.

**Architecture:** Three layers. Layer 0 primitives keep taking `Worksheet`/`Range` parameters. A new layer 1 of action modules composes them into user-intent transactions and returns a validated three-element result envelope. Layer 2 (`.evt` files) shrinks to reading controls, making one layer 1 call, and writing controls.

**Tech Stack:** VBA (Excel 2016+ object model), Python 3.12, pywin32, pytest. Build is `python build/build.py`, which regenerates `dist/HarnessCreator.xlsm` from `src/vba/` via COM injection.

**Spec:** `docs/superpowers/specs/2026-08-28-ui-logic-separation-design.md`

## Global Constraints

- Every VBA module starts with `Attribute VB_Name = "<name>"` then `Option Explicit`. No `Option Base` directive may be added anywhere - it would break the zero-based `Array()` property the result envelope depends on.
- Layer 1 modules (`modContract`, `modMessages`, `modEditorActions`, `modPickerActions`, `modManageActions`) may not reference `MSForms.*`, `MsgBox`, `InputBox`, `GetOpenFilename`, `GetSaveAsFilename`, `.Show`, `Unload`, `Workbooks.Open`, or `DoEvents`.
- Layer 2 (`.evt` files) may not reference a layer 0 module (`modLibrary`, `modPinEditor`, `modChart`, `modConnectors`, `modSnapshot`, `modLibraryTransfer`, `modState`, `modUtil`) directly.
- Layer 1 gets no blanket `On Error Resume Next`. Expected failure is an outcome code; unexpected failure propagates.
- In every `.evt` handler: read all control values and form-level variables into locals **before** the layer 1 call; `Unload Me` is the **last statement in its branch**; no line after `Unload Me` references a control or a form-level variable.
- Every new module name must be appended to `VBA_MODULES` in `build/build.py:21`, or the build silently omits it.
- Tests run `python -m pytest -v` from the repo root. The `artifact` fixture rebuilds the workbook automatically; never hand-edit `dist/`.
- Existing test helper: `from tests.conftest import run` calls `wb.Application.Run(f"'{wb.Name}'!{macro}", *args)`.
- Normalized coordinates are compared with `pytest.approx`, never `==`. `imgPhoto` arithmetic is `Single` widened to `Double`.

---

### Task 1: modContract - the result envelope

**Files:**
- Create: `src/vba/modContract.bas`
- Modify: `build/build.py:21-25`
- Test: `tests/test_contract.py`

**Interfaces:**
- Consumes: nothing.
- Produces: `modContract.Success(sOutcome, [vPayload]) As Variant`, `modContract.Failure(sOutcome, [vPayload]) As Variant`, `modContract.Ok(vResult) As Boolean`, `modContract.Outcome(vResult) As String`, `modContract.Payload(vResult) As Variant`, `modContract.TableRowCount(vPayload) As Long`, `modContract.PayloadKind(sOutcome) As String`, `modContract.OutcomeCodes() As Variant`. Every later task's action functions return `Success`/`Failure`.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_contract.py`:

```python
import pytest
import pywintypes

from tests.conftest import run


def test_success_builds_a_three_element_result(wb):
    assert run(wb, "modContract.Success", "SAVED", "DTM-04P") == (True, "SAVED", "DTM-04P")


def test_failure_builds_a_three_element_result(wb):
    assert run(wb, "modContract.Failure", "ID_COLLISION", "DTM-04P") == (False, "ID_COLLISION", "DTM-04P")


def test_a_none_kind_outcome_carries_an_empty_payload(wb):
    assert run(wb, "modContract.Success", "OK") == (True, "OK", None)


def test_accessors_read_the_slots_by_name(wb):
    result = run(wb, "modContract.Success", "PLACED", 3)
    assert run(wb, "modContract.Ok", result) is True
    assert run(wb, "modContract.Outcome", result) == "PLACED"
    assert run(wb, "modContract.Payload", result) == 3


def test_an_unknown_outcome_code_raises(wb):
    with pytest.raises(pywintypes.com_error):
        run(wb, "modContract.Success", "NOT_A_REAL_CODE", "x")


def test_a_payload_of_the_wrong_kind_raises(wb):
    # SAVED declares STRING; 42 arrives as a numeric Variant.
    with pytest.raises(pywintypes.com_error):
        run(wb, "modContract.Success", "SAVED", 42)


def test_every_declared_code_has_a_payload_kind(wb):
    for code in run(wb, "modContract.OutcomeCodes"):
        assert run(wb, "modContract.PayloadKind", code) != ""


def test_table_row_count_handles_empty_and_populated(wb):
    assert run(wb, "modContract.TableRowCount", None) == 0
    assert run(wb, "modContract.TableRowCount", ((1, 2), (3, 4))) == 2
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_contract.py -v`
Expected: every test FAILS - `Application.Run` raises because `modContract` does not exist.

- [ ] **Step 3: Create the module**

Create `src/vba/modContract.bas`:

```vba
Attribute VB_Name = "modContract"
Option Explicit

Public Const KIND_NONE As String = "NONE"
Public Const KIND_STRING As String = "STRING"
Public Const KIND_LONG As String = "LONG"
Public Const KIND_DOUBLE As String = "DOUBLE"
Public Const KIND_TABLE As String = "TABLE"

' Every outcome code any layer 1 action may return. Success/Failure reject
' anything absent from this list, so a typo fails at construction rather
' than reaching an adapter that has no case for it.
Public Function OutcomeCodes() As Variant
    OutcomeCodes = Array( _
        "PLACED", "MOVED_ANCHOR", "BAD_PIN_COUNT", "PIN_LIMIT_REACHED", "NO_OP", _
        "SAVED", "ID_COLLISION", "SAVE_FAILED", _
        "CACHE_READY", "NEEDS_BACKFILL", _
        "OK", "MISSING_NAME_OR_PART", _
        "PIN_DELETED", "PIN_NOT_FOUND", _
        "ADDED", "ADD_FAILED", "CONNECTOR_NOT_FOUND", "CONNECTOR_DELETED", _
        "EXPORTED", "EXPORT_FAILED", "IMPORTED", _
        "PHOTO_ATTACHED", "PHOTO_FAILED", _
        "RENAMED", "RENAME_REJECTED", "NO_RENAME", _
        "BULK_REBUILT", "CELLS_REBUILT", "UNITS_SET")
End Function

' One code maps to exactly one payload kind. PIN_DELETED and
' CONNECTOR_DELETED are deliberately distinct codes rather than one
' DELETED: their payloads are a pin number and a connector ID.
Public Function PayloadKind(ByVal sOutcome As String) As String
    Select Case sOutcome
        Case "PLACED", "MOVED_ANCHOR", "PIN_LIMIT_REACHED", "PIN_DELETED", "PIN_NOT_FOUND"
            PayloadKind = KIND_LONG
        Case "BULK_REBUILT", "CELLS_REBUILT"
            PayloadKind = KIND_LONG
        Case "SAVED", "ID_COLLISION", "SAVE_FAILED", "CACHE_READY", "NEEDS_BACKFILL"
            PayloadKind = KIND_STRING
        Case "ADDED", "ADD_FAILED", "CONNECTOR_NOT_FOUND", "CONNECTOR_DELETED"
            PayloadKind = KIND_STRING
        Case "EXPORTED", "EXPORT_FAILED", "PHOTO_ATTACHED", "PHOTO_FAILED"
            PayloadKind = KIND_STRING
        Case "RENAMED", "RENAME_REJECTED", "UNITS_SET"
            PayloadKind = KIND_STRING
        Case "IMPORTED"
            PayloadKind = KIND_TABLE
        Case "BAD_PIN_COUNT", "NO_OP", "OK", "MISSING_NAME_OR_PART", "NO_RENAME"
            PayloadKind = KIND_NONE
        Case Else
            PayloadKind = ""
    End Select
End Function

Public Function Success(ByVal sOutcome As String, Optional ByVal vPayload As Variant) As Variant
    If IsMissing(vPayload) Then
        Success = Build(True, sOutcome, Empty)
    Else
        Success = Build(True, sOutcome, vPayload)
    End If
End Function

Public Function Failure(ByVal sOutcome As String, Optional ByVal vPayload As Variant) As Variant
    If IsMissing(vPayload) Then
        Failure = Build(False, sOutcome, Empty)
    Else
        Failure = Build(False, sOutcome, vPayload)
    End If
End Function

' Array() is zero based for an in-process VBA caller and for a COM caller
' alike, so vResult(0) here and result[0] in pytest are the same element.
Private Function Build(ByVal bOk As Boolean, ByVal sOutcome As String, _
                       ByVal vPayload As Variant) As Variant
    Dim sKind As String
    sKind = PayloadKind(sOutcome)
    If Len(sKind) = 0 Then
        Err.Raise vbObjectError + 1, "modContract", "Unknown outcome code: " & sOutcome
    End If
    If Not PayloadMatches(sKind, vPayload) Then
        Err.Raise vbObjectError + 2, "modContract", _
            "Payload for " & sOutcome & " must be " & sKind
    End If
    Build = Array(bOk, sOutcome, vPayload)
End Function

Private Function PayloadMatches(ByVal sKind As String, ByVal vPayload As Variant) As Boolean
    Select Case sKind
        Case KIND_NONE:   PayloadMatches = IsEmpty(vPayload)
        Case KIND_STRING: PayloadMatches = (VarType(vPayload) = vbString)
        Case KIND_LONG:   PayloadMatches = (Not IsArray(vPayload)) And IsNumeric(vPayload)
        Case KIND_DOUBLE: PayloadMatches = (Not IsArray(vPayload)) And IsNumeric(vPayload)
        Case KIND_TABLE:  PayloadMatches = IsEmpty(vPayload) Or IsArray(vPayload)
    End Select
End Function

Public Function Ok(vResult As Variant) As Boolean
    Ok = CBool(vResult(0))
End Function

Public Function Outcome(vResult As Variant) As String
    Outcome = CStr(vResult(1))
End Function

Public Function Payload(vResult As Variant) As Variant
    If IsObject(vResult(2)) Then
        Set Payload = vResult(2)
    Else
        Payload = vResult(2)
    End If
End Function

' Adapters call this rather than LBound/UBound, so a zero row table cannot
' produce a subscript error in a form.
Public Function TableRowCount(vPayload As Variant) As Long
    If IsEmpty(vPayload) Then Exit Function
    If Not IsArray(vPayload) Then Exit Function
    TableRowCount = UBound(vPayload, 1) - LBound(vPayload, 1) + 1
End Function
```

- [ ] **Step 4: Register the module with the build**

In `build/build.py`, change the `VBA_MODULES` list to append `"modContract.bas"`:

```python
VBA_MODULES = [
    "modUtil.bas", "modState.bas", "modConnectors.bas", "modChart.bas",
    "modLibrary.bas", "modPinEditor.bas", "clsPinMarker.cls", "modSnapshot.bas",
    "modConnectorUI.bas", "modLibraryTransfer.bas", "modContract.bas",
]
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m pytest tests/test_contract.py -v`
Expected: all 8 PASS.

- [ ] **Step 6: Run the full suite to confirm nothing regressed**

Run: `python -m pytest -v`
Expected: all pre-existing tests still PASS.

- [ ] **Step 7: Commit**

```bash
git add src/vba/modContract.bas build/build.py tests/test_contract.py
git commit -m "feat: add the layer 1 result envelope and its outcome registry"
```

---

### Task 2: conftest run_action helper

**Files:**
- Modify: `tests/conftest.py`
- Test: `tests/test_contract.py` (extend)

**Interfaces:**
- Consumes: `modContract.OutcomeCodes` from Task 1.
- Produces: `tests.conftest.Result` (a frozen dataclass with `.ok`, `.outcome`, `.payload`) and `tests.conftest.run_action(wb, macro, *args) -> Result`. Every later task's action tests call `run_action`; query tests keep calling `run`.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_contract.py`:

```python
def test_run_action_validates_and_unpacks(wb):
    from tests.conftest import run_action

    result = run_action(wb, "modContract.Success", "PLACED", 3)
    assert result.ok is True
    assert result.outcome == "PLACED"
    assert result.payload == 3


def test_run_action_rejects_a_non_envelope_return(wb):
    from tests.conftest import run_action

    # PayloadKind is a query returning a bare string, not an envelope.
    with pytest.raises(AssertionError):
        run_action(wb, "modContract.PayloadKind", "SAVED")
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_contract.py -k run_action -v`
Expected: FAIL with `ImportError: cannot import name 'run_action'`.

- [ ] **Step 3: Add the helper**

Append to `tests/conftest.py`:

```python
from dataclasses import dataclass


@dataclass(frozen=True)
class Result:
    """A layer 1 action's three-element envelope, unpacked."""

    ok: bool
    outcome: str
    payload: object


_known_outcomes: set[str] | None = None


def known_outcomes(wb) -> set[str]:
    """The outcome registry, read once from VBA so it cannot drift from a
    Python mirror."""
    global _known_outcomes
    if _known_outcomes is None:
        _known_outcomes = set(run(wb, "modContract.OutcomeCodes"))
    return _known_outcomes


def run_action(wb, macro: str, *args) -> Result:
    """Call a layer 1 action and validate its envelope before returning it.

    Queries return bare values and must use `run` instead.
    """
    raw = run(wb, macro, *args)
    assert isinstance(raw, (tuple, list)) and len(raw) == 3, (
        f"{macro} returned {raw!r}, not a three-element result"
    )
    ok, outcome, payload = raw
    assert isinstance(ok, bool), f"{macro} returned a non-boolean ok: {ok!r}"
    assert isinstance(outcome, str) and outcome, f"{macro} returned a blank outcome"
    assert outcome in known_outcomes(wb), f"{macro} returned unregistered outcome {outcome!r}"
    return Result(ok, outcome, payload)
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_contract.py -v`
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add tests/conftest.py tests/test_contract.py
git commit -m "test: add run_action, which validates every action envelope"
```

---

### Task 3: modMessages - all user-visible text

**Files:**
- Create: `src/vba/modMessages.bas`
- Modify: `build/build.py:21-25`
- Test: `tests/test_messages.py`

**Interfaces:**
- Consumes: `modContract.Outcome`, `modContract.Payload`, `modContract.Ok`, `modContract.TableRowCount` from Task 1.
- Produces: `modMessages.MessageFor(vResult) As String`, `modMessages.MessageStyleFor(vResult) As Long`. Every `.evt` handler that shows a message calls these two.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_messages.py`:

```python
from tests.conftest import run


def message_for(wb, ok, outcome, payload=None):
    builder = "modContract.Success" if ok else "modContract.Failure"
    result = run(wb, builder, outcome, payload) if payload is not None else run(wb, builder, outcome)
    return run(wb, "modMessages.MessageFor", result)


def test_id_collision_names_the_offending_id(wb):
    assert message_for(wb, False, "ID_COLLISION", "DTM-04P") == (
        "Part Number already exists in the library (DTM-04P). "
        "Choose a different Part Number."
    )


def test_save_failed_tells_the_student_what_to_do(wb):
    # The current build fails this save silently. The message is the fix.
    assert message_for(wb, False, "SAVE_FAILED", "DTM-04P") == (
        "Could not save DTM-04P. Load a photo before saving."
    )


def test_pin_limit_reached_names_the_cap(wb):
    assert message_for(wb, False, "PIN_LIMIT_REACHED", 4) == "All 4 pins have been placed."


def test_bad_pin_count_needs_no_payload(wb):
    assert message_for(wb, False, "BAD_PIN_COUNT") == (
        "Enter a valid Pin Count before placing pins."
    )


def test_missing_name_or_part(wb):
    assert message_for(wb, False, "MISSING_NAME_OR_PART") == (
        "Enter Name and Part Number before loading a photo."
    )


def test_export_success_and_failure(wb):
    assert message_for(wb, True, "EXPORTED", "DTM-04P") == "Exported DTM-04P."
    assert message_for(wb, False, "EXPORT_FAILED", "DTM-04P") == "Could not export DTM-04P."


def test_silent_outcomes_produce_no_message(wb):
    for outcome in ("PLACED", "MOVED_ANCHOR", "NO_OP", "OK", "NO_RENAME"):
        payload = 1 if outcome in ("PLACED", "MOVED_ANCHOR") else None
        assert message_for(wb, True, outcome, payload) == ""


def test_style_is_information_on_success_and_exclamation_on_failure(wb):
    ok = run(wb, "modContract.Success", "EXPORTED", "DTM-04P")
    bad = run(wb, "modContract.Failure", "EXPORT_FAILED", "DTM-04P")
    assert run(wb, "modMessages.MessageStyleFor", ok) == 64      # vbInformation
    assert run(wb, "modMessages.MessageStyleFor", bad) == 48     # vbExclamation
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_messages.py -v`
Expected: every test FAILS - `modMessages` does not exist.

- [ ] **Step 3: Create the module**

Create `src/vba/modMessages.bas`:

```vba
Attribute VB_Name = "modMessages"
Option Explicit

' Every string a student reads lives here rather than in an .evt file, so
' the exact rendered text is asserted by a test. An outcome with no case
' below is silent by design - the adapter shows nothing for it.
Public Function MessageFor(vResult As Variant) As String
    Dim sOutcome As String, vPayload As Variant
    sOutcome = modContract.Outcome(vResult)
    vPayload = modContract.Payload(vResult)

    Select Case sOutcome
        Case "ID_COLLISION"
            MessageFor = "Part Number already exists in the library (" & _
                CStr(vPayload) & "). Choose a different Part Number."
        Case "SAVE_FAILED"
            MessageFor = "Could not save " & CStr(vPayload) & ". Load a photo before saving."
        Case "BAD_PIN_COUNT"
            MessageFor = "Enter a valid Pin Count before placing pins."
        Case "PIN_LIMIT_REACHED"
            MessageFor = "All " & CStr(vPayload) & " pins have been placed."
        Case "MISSING_NAME_OR_PART"
            MessageFor = "Enter Name and Part Number before loading a photo."
        Case "CONNECTOR_NOT_FOUND"
            MessageFor = "No connector '" & CStr(vPayload) & "' found in the library."
        Case "ADD_FAILED"
            MessageFor = "Could not add an instance of " & CStr(vPayload) & "."
        Case "CONNECTOR_DELETED"
            MessageFor = "Deleted " & CStr(vPayload) & "."
        Case "EXPORTED"
            MessageFor = "Exported " & CStr(vPayload) & "."
        Case "EXPORT_FAILED"
            MessageFor = "Could not export " & CStr(vPayload) & "."
        Case "IMPORTED"
            MessageFor = "Import complete. " & _
                CStr(modContract.TableRowCount(vPayload)) & " connector(s) imported."
        Case "PHOTO_FAILED"
            MessageFor = "Could not attach a photo for " & CStr(vPayload) & "."
        Case Else
            MessageFor = ""
    End Select
End Function

Public Function MessageStyleFor(vResult As Variant) As Long
    If modContract.Ok(vResult) Then
        MessageStyleFor = vbInformation
    Else
        MessageStyleFor = vbExclamation
    End If
End Function
```

- [ ] **Step 4: Register the module with the build**

In `build/build.py`, append `"modMessages.bas"` to `VBA_MODULES` after `"modContract.bas"`.

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m pytest tests/test_messages.py -v`
Expected: all 8 PASS.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modMessages.bas build/build.py tests/test_messages.py
git commit -m "feat: move every user-visible message into a tested string table"
```

---

### Task 4: modPinEditor marker geometry

**Files:**
- Modify: `src/vba/modPinEditor.bas` (append two functions after `FitAspectRatio`)
- Test: `tests/test_marker_geometry.py`

**Interfaces:**
- Consumes: nothing.
- Produces: `modPinEditor.MarkerTopLeft(dNormX, dNormY, dPhotoLeft, dPhotoTop, dPhotoW, dPhotoH, dMarkerW, dMarkerH) As Variant` returning `Array(left, top)`, and `modPinEditor.NormFromMarker(dLeft, dTop, dMarkerW, dMarkerH, dPhotoLeft, dPhotoTop, dPhotoW, dPhotoH) As Variant` returning `Array(normX, normY)`. Task 9's `AddMarkerControl` and `clsPinMarker` call them. These are layer 0 primitives - bare `Array` returns, not envelopes.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_marker_geometry.py`:

```python
import pytest

from tests.conftest import run

PHOTO_LEFT, PHOTO_TOP, PHOTO_W, PHOTO_H = 12.0, 30.0, 180.0, 120.0
MARKER = 16.0


def test_marker_top_left_centres_the_badge_on_the_point(wb):
    left, top = run(wb, "modPinEditor.MarkerTopLeft",
                    0.5, 0.5, PHOTO_LEFT, PHOTO_TOP, PHOTO_W, PHOTO_H, MARKER, MARKER)
    # Centre of the photo is (12+90, 30+60); the badge is offset by half its size.
    assert left == pytest.approx(102.0 - 8.0)
    assert top == pytest.approx(90.0 - 8.0)


def test_norm_from_marker_recovers_the_normalized_point(wb):
    norm_x, norm_y = run(wb, "modPinEditor.NormFromMarker",
                         94.0, 82.0, MARKER, MARKER,
                         PHOTO_LEFT, PHOTO_TOP, PHOTO_W, PHOTO_H)
    assert norm_x == pytest.approx(0.5)
    assert norm_y == pytest.approx(0.5)


@pytest.mark.parametrize("norm_x,norm_y", [(0.0, 0.0), (1.0, 1.0), (0.25, 0.75), (0.6111, 0.3333)])
def test_the_two_conversions_round_trip(wb, norm_x, norm_y):
    left, top = run(wb, "modPinEditor.MarkerTopLeft",
                    norm_x, norm_y, PHOTO_LEFT, PHOTO_TOP, PHOTO_W, PHOTO_H, MARKER, MARKER)
    back_x, back_y = run(wb, "modPinEditor.NormFromMarker",
                         left, top, MARKER, MARKER,
                         PHOTO_LEFT, PHOTO_TOP, PHOTO_W, PHOTO_H)
    assert back_x == pytest.approx(norm_x)
    assert back_y == pytest.approx(norm_y)


def test_a_zero_sized_photo_returns_empty_rather_than_dividing_by_zero(wb):
    assert run(wb, "modPinEditor.NormFromMarker",
               94.0, 82.0, MARKER, MARKER, PHOTO_LEFT, PHOTO_TOP, 0.0, PHOTO_H) is None
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_marker_geometry.py -v`
Expected: every test FAILS - the functions do not exist.

- [ ] **Step 3: Append the two conversions**

Add to `src/vba/modPinEditor.bas`, immediately after `FitAspectRatio`:

```vba
' The on-screen top-left for a marker badge whose centre sits on the
' normalized point (dNormX, dNormY) within the photo control's box.
' Exact inverse of NormFromMarker.
Public Function MarkerTopLeft(ByVal dNormX As Double, ByVal dNormY As Double, _
                              ByVal dPhotoLeft As Double, ByVal dPhotoTop As Double, _
                              ByVal dPhotoW As Double, ByVal dPhotoH As Double, _
                              ByVal dMarkerW As Double, ByVal dMarkerH As Double) As Variant
    MarkerTopLeft = Array( _
        dPhotoLeft + dNormX * dPhotoW - (dMarkerW / 2), _
        dPhotoTop + dNormY * dPhotoH - (dMarkerH / 2))
End Function

' The normalized point under a marker badge's centre. Exact inverse of
' MarkerTopLeft. Empty when the photo box has no area to normalize against.
Public Function NormFromMarker(ByVal dLeft As Double, ByVal dTop As Double, _
                               ByVal dMarkerW As Double, ByVal dMarkerH As Double, _
                               ByVal dPhotoLeft As Double, ByVal dPhotoTop As Double, _
                               ByVal dPhotoW As Double, ByVal dPhotoH As Double) As Variant
    If dPhotoW <= 0 Or dPhotoH <= 0 Then Exit Function

    NormFromMarker = Array( _
        (dLeft + dMarkerW / 2 - dPhotoLeft) / dPhotoW, _
        (dTop + dMarkerH / 2 - dPhotoTop) / dPhotoH)
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_marker_geometry.py -v`
Expected: all 7 PASS (4 parametrized round trips plus 3).

- [ ] **Step 5: Commit**

```bash
git add src/vba/modPinEditor.bas tests/test_marker_geometry.py
git commit -m "feat: extract marker pixel and normalized coordinate conversions"
```

---

### Task 5: modLibrary.ConnectorIndex

**Files:**
- Modify: `src/vba/modLibrary.bas` (append after `ReadConnector`)
- Test: `tests/test_library_connectors.py` (extend)

**Interfaces:**
- Consumes: `modLibrary.LIB_COL_ID`, `modLibrary.LIB_COL_NAME`, `modLibrary.LastUsedRowInWindow` (all existing).
- Produces: `modLibrary.ConnectorIndex(wsConn) As Variant` returning `Empty` or a 2D array `(1 To n, 1 To 2)` of display string and ConnectorID. Tasks 10 and 12 use it to fill list boxes; it replaces `modConnectorUI.RefreshConnectorList`.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_library_connectors.py`:

```python
def test_connector_index_renders_display_strings_and_ids(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    fields_a = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
                4, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    fields_b = ("AMP-02", "AMP 2-way", "TE", "AMP-2", "Connector",
                2, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, fields_a)
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, fields_b)

    index = run(wb, "modLibrary.ConnectorIndex", ws)
    assert [row[0] for row in index] == ["DTM-04P - Deutsch DTM 4-way", "AMP-02 - AMP 2-way"]
    assert [row[1] for row in index] == ["DTM-04P", "AMP-02"]


def test_connector_index_of_an_empty_sheet_is_empty(wb, library_wb):
    assert run(wb, "modLibrary.ConnectorIndex", library_wb.Worksheets("Connectors")) is None
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_library_connectors.py -k connector_index -v`
Expected: FAIL - `ConnectorIndex` does not exist.

- [ ] **Step 3: Append the function**

Add to `src/vba/modLibrary.bas`, immediately after `ReadConnector`:

```vba
' The library's browsable index: one row per connector, holding the display
' string a list box shows and the ConnectorID that string resolves to.
' Rendering happens here, not in a form, so the exact text is testable.
' Empty when the sheet holds no connectors - never a zero length array.
Public Function ConnectorIndex(wsConn As Worksheet) As Variant
    Dim nLast As Long, r As Long, n As Long
    Dim vRows() As Variant

    nLast = wsConn.Cells(wsConn.Rows.Count, LIB_COL_ID).End(xlUp).Row
    If nLast < 2 Then Exit Function

    ReDim vRows(1 To nLast - 1, 1 To 2)
    For r = 2 To nLast
        If Len(Trim$(CStr(wsConn.Cells(r, LIB_COL_ID).Value))) > 0 Then
            n = n + 1
            vRows(n, 1) = Trim$(CStr(wsConn.Cells(r, LIB_COL_ID).Value)) & " - " & _
                          CStr(wsConn.Cells(r, LIB_COL_NAME).Value)
            vRows(n, 2) = Trim$(CStr(wsConn.Cells(r, LIB_COL_ID).Value))
        End If
    Next r

    If n = 0 Then Exit Function

    Dim vResult() As Variant, i As Long
    ReDim vResult(1 To n, 1 To 2)
    For i = 1 To n
        vResult(i, 1) = vRows(i, 1)
        vResult(i, 2) = vRows(i, 2)
    Next i
    ConnectorIndex = vResult
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_library_connectors.py -v`
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modLibrary.bas tests/test_library_connectors.py
git commit -m "feat: add ConnectorIndex, the testable replacement for RefreshConnectorList"
```

---

### Task 6: modEditorActions - the query functions

**Files:**
- Create: `src/vba/modEditorActions.bas`
- Modify: `build/build.py:21-25`
- Test: `tests/test_editor_actions.py`

**Interfaces:**
- Consumes: `modLibrary.CachePhotoPath`, `modLibrary.ReadPinsForConnector`, `modPinEditor.SCRATCH_FIRST_ROW`, `modPinEditor.SCRATCH_LAST_ROW`, `modLibrary.PIN_COL_CONNID`, `modLibrary.PIN_COL_PINNUM` (all existing).
- Produces (all **queries** - bare returns, tested with `run` not `run_action`): `PhotoFileFilter() As String`, `MarkerControlName(nPinNumber) As String`, `PhotoCacheRefreshTarget(sWorkbookPath, sConnectorID, sPhotoPath) As String`, `TypeListItems(wsLists) As Variant`, `PinListItems(wsScratch, sConnectorID) As Variant` (2D `(1 To n, 1 To 4)`: display, pin number, labelX, labelY), `NextPinNumber(wsScratch, sConnectorID) As Long`.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_editor_actions.py`:

```python
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
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_editor_actions.py -v`
Expected: every test FAILS - `modEditorActions` does not exist.

- [ ] **Step 3: Create the module with its query functions**

Create `src/vba/modEditorActions.bas`:

```vba
Attribute VB_Name = "modEditorActions"
Option Explicit

' Restricted to JPG: LoadPicture's legacy OLE loader rejects valid PNGs
' with error 481 on some Windows/Office configurations, even though
' Shapes.AddPicture handles the same file. Offering only what LoadPicture
' opens is the fix, and keeping the string here makes it assertable.
Public Function PhotoFileFilter() As String
    PhotoFileFilter = "Pictures (*.jpg; *.jpeg), *.jpg;*.jpeg"
End Function

Public Function MarkerControlName(ByVal nPinNumber As Long) As String
    MarkerControlName = "lblMarker" & CStr(nPinNumber)
End Function

' Where a just-saved photo should be copied so the editor's preview can
' read it back without a clipboard round trip. Empty when nothing needs
' copying: no photo was chosen, or the chosen file already is the cache
' (FileCopy onto itself raises).
Public Function PhotoCacheRefreshTarget(ByVal sWorkbookPath As String, _
                                        ByVal sConnectorID As String, _
                                        ByVal sPhotoPath As String) As String
    Dim sCachePath As String

    If Len(Trim$(sPhotoPath)) = 0 Then Exit Function

    sCachePath = modLibrary.CachePhotoPath(sWorkbookPath, sConnectorID, "jpg")
    If StrComp(sPhotoPath, sCachePath, vbTextCompare) = 0 Then Exit Function

    PhotoCacheRefreshTarget = sCachePath
End Function

' The connector Type options, read from the sheet passed in rather than
' through an unqualified RowSource, which resolved against ActiveWorkbook
' and left the combo empty during frmManageLibrary's Edit flow.
Public Function TypeListItems(wsLists As Worksheet) As Variant
    Dim vItems() As String, r As Long, n As Long

    r = 2
    Do While Len(Trim$(CStr(wsLists.Cells(r, 4).Value))) > 0
        n = n + 1
        ReDim Preserve vItems(1 To n)
        vItems(n) = CStr(wsLists.Cells(r, 4).Value)
        r = r + 1
    Loop

    If n = 0 Then Exit Function
    TypeListItems = vItems
End Function

' One row per placed pin: the display string the list box shows, the pin
' number that row resolves to, and the marker's normalized position.
' Derived from the scratch sheet on every call, which is what lets the
' form drop the mListPinNumbers collection that used to desync from it.
Public Function PinListItems(wsScratch As Worksheet, ByVal sConnectorID As String) As Variant
    Dim vPins As Variant, i As Long, n As Long
    Dim vRows() As Variant

    vPins = modLibrary.ReadPinsForConnector(wsScratch, modPinEditor.SCRATCH_FIRST_ROW, _
                                            modPinEditor.SCRATCH_LAST_ROW, sConnectorID)
    If IsEmpty(vPins) Then Exit Function

    n = UBound(vPins, 1) - LBound(vPins, 1) + 1
    ReDim vRows(1 To n, 1 To 4)
    For i = 1 To n
        vRows(i, 1) = "Pin " & CStr(CLng(vPins(LBound(vPins, 1) + i - 1, 2)))
        vRows(i, 2) = CLng(vPins(LBound(vPins, 1) + i - 1, 2))
        vRows(i, 3) = CDbl(vPins(LBound(vPins, 1) + i - 1, 6))
        vRows(i, 4) = CDbl(vPins(LBound(vPins, 1) + i - 1, 7))
    Next i

    PinListItems = vRows
End Function

' One past the highest placed pin number, so a deletion never causes a
' reused number. Replaces the form's mNextPinNumber counter.
Public Function NextPinNumber(wsScratch As Worksheet, ByVal sConnectorID As String) As Long
    Dim vItems As Variant, i As Long, nMax As Long

    vItems = PinListItems(wsScratch, sConnectorID)
    If IsEmpty(vItems) Then
        NextPinNumber = 1
        Exit Function
    End If

    For i = LBound(vItems, 1) To UBound(vItems, 1)
        If CLng(vItems(i, 2)) > nMax Then nMax = CLng(vItems(i, 2))
    Next i

    NextPinNumber = nMax + 1
End Function
```

- [ ] **Step 4: Register the module with the build**

In `build/build.py`, append `"modEditorActions.bas"` to `VBA_MODULES`.

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m pytest tests/test_editor_actions.py -v`
Expected: all 10 PASS.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modEditorActions.bas build/build.py tests/test_editor_actions.py
git commit -m "feat: add the connector editor's query functions"
```

---

### Task 7: modEditorActions - CanLoadPhoto, PhotoSourceForEdit, DeletePinRequest

**Files:**
- Modify: `src/vba/modEditorActions.bas`
- Test: `tests/test_editor_actions.py` (extend)

**Interfaces:**
- Consumes: `modContract.Success`/`Failure` (Task 1), `modLibrary.CachePhotoPath`, `modPinEditor.RemovePin` (existing).
- Produces (all **actions** - envelopes, tested with `run_action`): `CanLoadPhoto(sName, sPartNumber) As Variant` → `OK` | `MISSING_NAME_OR_PART`; `PhotoSourceForEdit(sWorkbookPath, sConnectorID) As Variant` → `CACHE_READY` | `NEEDS_BACKFILL`, payload the cache path; `DeletePinRequest(wsScratch, sConnectorID, nPinNumber) As Variant` → `PIN_DELETED` | `PIN_NOT_FOUND`, payload the pin number.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_editor_actions.py`:

```python
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
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_editor_actions.py -k "can_load or photo_source or delete_pin" -v`
Expected: FAIL - the three functions do not exist.

- [ ] **Step 3: Append the three actions**

Add to `src/vba/modEditorActions.bas`:

```vba
' mConnectorID is derived from Name and Part Number when the photo loads
' and never recomputed, so placing pins before both are filled in used to
' do nothing at all, silently. The guard runs before the file picker opens.
Public Function CanLoadPhoto(ByVal sName As String, ByVal sPartNumber As String) As Variant
    If Len(Trim$(sName)) = 0 Or Len(Trim$(sPartNumber)) = 0 Then
        CanLoadPhoto = modContract.Failure("MISSING_NAME_OR_PART")
        Exit Function
    End If
    CanLoadPhoto = modContract.Success("OK")
End Function

' Where the editor's photo preview should read from, and whether a one-time
' backfill from the embedded Shape is needed first. The on-disk cache is
' preferred because re-exporting the Shape goes through the clipboard,
' which is unreliable for VBA-triggered operations on this machine.
Public Function PhotoSourceForEdit(ByVal sWorkbookPath As String, _
                                   ByVal sConnectorID As String) As Variant
    Dim sCachePath As String
    sCachePath = modLibrary.CachePhotoPath(sWorkbookPath, sConnectorID, "jpg")

    If Len(Dir$(sCachePath)) = 0 Then
        PhotoSourceForEdit = modContract.Failure("NEEDS_BACKFILL", sCachePath)
        Exit Function
    End If

    PhotoSourceForEdit = modContract.Success("CACHE_READY", sCachePath)
End Function

Public Function DeletePinRequest(wsScratch As Worksheet, ByVal sConnectorID As String, _
                                 ByVal nPinNumber As Long) As Variant
    If modPinEditor.RemovePin(wsScratch, sConnectorID, nPinNumber) Then
        DeletePinRequest = modContract.Success("PIN_DELETED", nPinNumber)
    Else
        DeletePinRequest = modContract.Failure("PIN_NOT_FOUND", nPinNumber)
    End If
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_editor_actions.py -v`
Expected: all PASS (10 from Task 6 plus 9 here).

- [ ] **Step 5: Commit**

```bash
git add src/vba/modEditorActions.bas tests/test_editor_actions.py
git commit -m "feat: add the editor's photo guard, cache source, and pin delete actions"
```

---

### Task 8: modEditorActions.PhotoClickAction

**Files:**
- Modify: `src/vba/modEditorActions.bas`
- Test: `tests/test_editor_actions.py` (extend)

**Interfaces:**
- Consumes: `NextPinNumber`, `PinListItems` (Task 6), `modContract.Success`/`Failure` (Task 1), `modPinEditor.PlacePin`, `modPinEditor.MoveAnchor` (existing).
- Produces: `PhotoClickAction(wsScratch, sConnectorID, bPlaceMode, nSelectedPin, sPinCountText, dNormX, dNormY) As Variant` → `PLACED` (payload pin number) | `MOVED_ANCHOR` (payload pin number) | `BAD_PIN_COUNT` | `PIN_LIMIT_REACHED` (payload the cap) | `NO_OP`. **`nSelectedPin` is a pin number, not a list index**; `0` means nothing selected. Task 9's `imgPhoto_MouseUp` resolves the list index through `PinListItems` before calling.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_editor_actions.py`:

```python
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
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_editor_actions.py -k click -v`
Expected: FAIL - `PhotoClickAction` does not exist.

- [ ] **Step 3: Append the action**

Add to `src/vba/modEditorActions.bas`:

```vba
' Everything a click on the photo can mean. bPlaceMode is the Place Pins
' toggle; nSelectedPin is a PIN NUMBER (0 when nothing is selected), not a
' list index - the caller resolves that through PinListItems. The placed
' count and next pin number are derived from wsScratch rather than passed
' in, so no counter in the form can drift out of sync with the sheet.
Public Function PhotoClickAction(wsScratch As Worksheet, ByVal sConnectorID As String, _
                                 ByVal bPlaceMode As Boolean, ByVal nSelectedPin As Long, _
                                 ByVal sPinCountText As String, _
                                 ByVal dNormX As Double, ByVal dNormY As Double) As Variant
    If Len(Trim$(sConnectorID)) = 0 Then
        PhotoClickAction = modContract.Failure("NO_OP")
        Exit Function
    End If

    If bPlaceMode Then
        Dim nPinCount As Long, nPlaced As Long, nNext As Long

        If Not IsNumeric(Trim$(sPinCountText)) Then
            PhotoClickAction = modContract.Failure("BAD_PIN_COUNT")
            Exit Function
        End If
        nPinCount = CLng(Val(sPinCountText))
        If nPinCount <= 0 Then
            PhotoClickAction = modContract.Failure("BAD_PIN_COUNT")
            Exit Function
        End If

        nPlaced = modContract.TableRowCount(PinListItems(wsScratch, sConnectorID))
        If nPlaced >= nPinCount Then
            PhotoClickAction = modContract.Failure("PIN_LIMIT_REACHED", nPinCount)
            Exit Function
        End If

        nNext = NextPinNumber(wsScratch, sConnectorID)
        If modPinEditor.PlacePin(wsScratch, sConnectorID, nNext, _
                                 "Pin " & CStr(nNext), dNormX, dNormY) Then
            PhotoClickAction = modContract.Success("PLACED", nNext)
        Else
            PhotoClickAction = modContract.Failure("NO_OP")
        End If
        Exit Function
    End If

    ' Place Pins is off and a pin is selected: the click moves that pin's
    ' anchor. modPinEditor.MoveAnchor decides whether the marker travels
    ' with it (it does only if it was still sitting on the anchor).
    If nSelectedPin > 0 Then
        If modPinEditor.MoveAnchor(wsScratch, sConnectorID, nSelectedPin, dNormX, dNormY) Then
            PhotoClickAction = modContract.Success("MOVED_ANCHOR", nSelectedPin)
            Exit Function
        End If
    End If

    PhotoClickAction = modContract.Failure("NO_OP")
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_editor_actions.py -v`
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modEditorActions.bas tests/test_editor_actions.py
git commit -m "feat: make every photo click outcome testable"
```

---

### Task 9: modEditorActions.SaveFromEditor

**Files:**
- Modify: `src/vba/modEditorActions.bas`
- Test: `tests/test_editor_actions.py` (extend)

**Interfaces:**
- Consumes: `modContract.Success`/`Failure` (Task 1), `modLibrary.FindConnectorRow`, `modLibrary.LIB_ROW_CAP`, `modPinEditor.SaveConnector` (existing).
- Produces: `SaveFromEditor(wsLibConn, wsLibPins, wsLibPhotos, wsScratch, sConnectorID, sOriginalID, vFields, sPhotoPath, sNowUtc) As Variant` → `SAVED` | `ID_COLLISION` | `SAVE_FAILED`, payload the connector ID in all three cases. `vFields` is a six-element zero-based array: name, manufacturer, part number, type, **pin count as a String**, notes.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_editor_actions.py`:

```python
from tests.fixtures.sample_photo import write_sample_photo

FIELDS = ("Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector", "4", "some notes")
NOW = "2026-08-28T09:00:00Z"


def save(wb, library_wb, connector_id, original_id, photo_path, fields=FIELDS):
    return run_action(
        wb, "modEditorActions.SaveFromEditor",
        library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
        library_wb.Worksheets("Photos"), wb.Worksheets("_Edit"),
        connector_id, original_id, fields, photo_path, NOW,
    )


def test_save_writes_the_connector_row_and_its_pins(wb, library_wb, tmp_path):
    photo = write_sample_photo(tmp_path / "photo.jpg")
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    write_scratch_pin(wb, ws, "DTM-04P", 1, "Pin 1", 0.1, 0.2, 0.3, 0.4)

    result = save(wb, library_wb, "DTM-04P", "", str(photo))
    assert (result.ok, result.outcome, result.payload) == (True, "SAVED", "DTM-04P")

    row = run(wb, "modLibrary.ReadConnector",
              library_wb.Worksheets("Connectors"), 2, 100000, "DTM-04P")
    assert row[1] == "Deutsch DTM 4-way"      # Name
    assert row[5] == 4                         # PinCount, coerced from "4"
    assert row[8] == NOW                       # CreatedUtc, passed in, not read from Now
    pins = run(wb, "modLibrary.ReadPinsForConnector",
               library_wb.Worksheets("Pins"), 2, 100000, "DTM-04P")
    assert len(pins) == 1


def test_save_rejects_an_id_that_collides_with_a_different_connector(wb, library_wb, tmp_path):
    photo = write_sample_photo(tmp_path / "photo.jpg")
    save(wb, library_wb, "DTM-04P", "", str(photo))

    # A different connector (original id AMP-02) now derives the same id.
    result = save(wb, library_wb, "DTM-04P", "AMP-02", str(photo))
    assert (result.ok, result.outcome, result.payload) == (False, "ID_COLLISION", "DTM-04P")


def test_resaving_the_same_connector_is_not_a_collision(wb, library_wb, tmp_path):
    photo = write_sample_photo(tmp_path / "photo.jpg")
    save(wb, library_wb, "DTM-04P", "", str(photo))

    result = save(wb, library_wb, "DTM-04P", "DTM-04P", str(photo))
    assert (result.ok, result.outcome) == (True, "SAVED")


def test_saving_with_no_photo_fails_loudly_rather_than_silently(wb, library_wb):
    # The pre-refactor build closed the library and did nothing at all here,
    # leaving the student with no indication why Save had no effect.
    result = save(wb, library_wb, "NOPHOTO-01", "", "")
    assert (result.ok, result.outcome, result.payload) == (False, "SAVE_FAILED", "NOPHOTO-01")


def test_save_replaces_the_previous_pin_set_rather_than_appending(wb, library_wb, tmp_path):
    photo = write_sample_photo(tmp_path / "photo.jpg")
    ws = wb.Worksheets("_Edit")
    run(wb, "modPinEditor.ClearScratchPins", ws)
    write_scratch_pin(wb, ws, "DTM-04P", 1, "Pin 1", 0.1, 0.1, 0.1, 0.1)
    write_scratch_pin(wb, ws, "DTM-04P", 2, "Pin 2", 0.2, 0.2, 0.2, 0.2)
    save(wb, library_wb, "DTM-04P", "", str(photo))

    run(wb, "modPinEditor.ClearScratchPins", ws)
    write_scratch_pin(wb, ws, "DTM-04P", 1, "Pin 1", 0.9, 0.9, 0.9, 0.9)
    save(wb, library_wb, "DTM-04P", "DTM-04P", str(photo))

    pins = run(wb, "modLibrary.ReadPinsForConnector",
               library_wb.Worksheets("Pins"), 2, 100000, "DTM-04P")
    assert len(pins) == 1
    assert pins[0][3] == pytest.approx(0.9)
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_editor_actions.py -k save -v`
Expected: FAIL - `SaveFromEditor` does not exist.

- [ ] **Step 3: Append the action**

Add to `src/vba/modEditorActions.bas`:

```vba
' The whole Save transaction bar the workbook open, close, and photo copy,
' which stay in the adapter. vFields is zero based: name, manufacturer,
' part number, type, pin count AS TEXT, notes - the pin count arrives as
' the text box supplies it so the coercion is tested here.
' sNowUtc is passed in rather than read from Now, which keeps this
' deterministic and lets a test assert the exact timestamp written.
Public Function SaveFromEditor(wsLibConn As Worksheet, wsLibPins As Worksheet, _
                               wsLibPhotos As Worksheet, wsScratch As Worksheet, _
                               ByVal sConnectorID As String, ByVal sOriginalID As String, _
                               ByVal vFields As Variant, ByVal sPhotoPath As String, _
                               ByVal sNowUtc As String) As Variant
    Dim nExistingRow As Long

    ' A collision only matters when the row it would overwrite belongs to a
    ' DIFFERENT connector than the one this session opened for editing -
    ' re-saving the connector you are editing must not flag itself.
    nExistingRow = modLibrary.FindConnectorRow(wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If nExistingRow > 0 And StrComp(sConnectorID, sOriginalID, vbTextCompare) <> 0 Then
        SaveFromEditor = modContract.Failure("ID_COLLISION", sConnectorID)
        Exit Function
    End If

    If modPinEditor.SaveConnector(wsLibConn, wsLibPins, wsLibPhotos, wsScratch, _
            sConnectorID, _
            CStr(vFields(LBound(vFields))), _
            CStr(vFields(LBound(vFields) + 1)), _
            CStr(vFields(LBound(vFields) + 2)), _
            CStr(vFields(LBound(vFields) + 3)), _
            CLng(Val(CStr(vFields(LBound(vFields) + 4)))), _
            CStr(vFields(LBound(vFields) + 5)), _
            sPhotoPath, sNowUtc, sNowUtc, "Local") Then
        SaveFromEditor = modContract.Success("SAVED", sConnectorID)
    Else
        SaveFromEditor = modContract.Failure("SAVE_FAILED", sConnectorID)
    End If
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_editor_actions.py -v`
Expected: all PASS.

- [ ] **Step 5: Run the full suite**

Run: `python -m pytest -v`
Expected: everything still PASS.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modEditorActions.bas tests/test_editor_actions.py
git commit -m "feat: make the editor Save transaction testable, and fail loudly with no photo"
```

---

### Task 10: Rewire frmConnectorEditor and clsPinMarker

**Files:**
- Modify: `src/vba/forms/frmConnectorEditor.evt` (rewrite)
- Modify: `src/vba/clsPinMarker.cls:46-56`
- Modify: `tests/test_connector_editor_wiring.py` (delete retired assertions)

**Interfaces:**
- Consumes: every function from Tasks 1, 3, 4, 6, 7, 8, 9.
- Produces: no new callable surface. `mListPinNumbers` and `mNextPinNumber` are deleted from the form; `mConnectorID`, `mOriginalConnectorID`, `mPhotoPath` and `mMarkers` remain.

- [ ] **Step 1: Rewrite the form's event module**

Replace `src/vba/forms/frmConnectorEditor.evt` entirely:

```vba
Option Explicit

' imgPhoto's design-time size (form_layout.py). Fitting must target this
' fixed box, not imgPhoto.Width/Height - those get overwritten with each
' fitted result, so using them as the box would shrink the photo further
' on every subsequent load.
Private Const PHOTO_BOX_WIDTH As Double = 180
Private Const PHOTO_BOX_HEIGHT As Double = 180

' Pin marker badge size - the Label control's default is sized for body
' text, not a small numbered dot on a photo.
Private Const PIN_MARKER_SIZE As Double = 16

Private mMarkers As Collection
Private mConnectorID As String
Private mOriginalConnectorID As String ' "" for a new connector
Private mPhotoPath As String

Private Sub UserForm_Initialize()
    Set mMarkers = New Collection

    Dim vTypes As Variant, i As Long
    vTypes = modEditorActions.TypeListItems(ThisWorkbook.Worksheets("_Lists"))
    If Not IsEmpty(vTypes) Then
        For i = LBound(vTypes) To UBound(vTypes)
            cboType.AddItem vTypes(i)
        Next i
    End If
    If cboType.ListCount > 0 Then cboType.ListIndex = 0
End Sub

' Called (after Load, before Show) by frmManageLibrary's cmdEdit_Click,
' which has already seeded _Edit's scratch pins. vFields is the 1-based
' array modLibrary.ReadConnector returns to an in-process caller.
Public Sub LoadForEdit(ByVal sConnectorID As String, ByVal vFields As Variant, _
                       wsLibPhotos As Worksheet)
    mConnectorID = sConnectorID
    mOriginalConnectorID = sConnectorID

    txtName.Text = CStr(vFields(2))
    txtManufacturer.Text = CStr(vFields(3))
    txtPartNumber.Text = CStr(vFields(4))
    cboType.Text = CStr(vFields(5))
    txtPinCount.Text = CStr(vFields(6))
    txtNotes.Text = CStr(vFields(7))

    LoadExistingPhoto sConnectorID, wsLibPhotos
    RebuildPinList
End Sub

Private Sub LoadExistingPhoto(ByVal sConnectorID As String, wsLibPhotos As Worksheet)
    Dim vResult As Variant, sCachePath As String
    vResult = modEditorActions.PhotoSourceForEdit(ThisWorkbook.Path, sConnectorID)
    sCachePath = CStr(modContract.Payload(vResult))

    If Not modContract.Ok(vResult) Then
        ' No cache yet. One-time backfill attempt from the embedded Shape:
        ' the clipboard route is unreliable here, but it gives an older
        ' connector a chance to get a cache going forward.
        Dim shp As Shape
        On Error Resume Next
        Set shp = wsLibPhotos.Shapes("PHOTO_" & sConnectorID)
        On Error GoTo 0
        If Not shp Is Nothing Then modLibraryShim_ExportShape shp, sCachePath
        If Len(Dir$(sCachePath)) = 0 Then Exit Sub
    End If

    ShowPhoto sCachePath
End Sub

' The one place this form still reaches a layer 0 routine, isolated so the
' layering lint can allow it by name: Shape export needs a live Shape
' object, which cannot cross Application.Run usefully.
Private Sub modLibraryShim_ExportShape(shp As Shape, ByVal sPath As String)
    modLibrary.ExportShapeToFile shp, sPath, "JPG"
End Sub

Private Sub ShowPhoto(ByVal sPath As String)
    Dim pic As IPictureDisp, vFit As Variant
    Set pic = LoadPicture(sPath)
    If pic Is Nothing Then Exit Sub

    vFit = modPinEditor.FitAspectRatio(CDbl(pic.Width), CDbl(pic.Height), _
                                       PHOTO_BOX_WIDTH, PHOTO_BOX_HEIGHT)
    If IsEmpty(vFit) Then Exit Sub

    imgPhoto.PictureSizeMode = fmPictureSizeModeStretch
    imgPhoto.Picture = pic
    imgPhoto.Width = vFit(0)
    imgPhoto.Height = vFit(1)
    mPhotoPath = sPath
End Sub

' The list box and the on-screen markers are both rebuilt from the scratch
' sheet, which is the only place a pin exists. No parallel collection.
Private Sub RebuildPinList()
    Dim vItems As Variant, i As Long

    lstPins.Clear
    ClearMarkerControls

    vItems = modEditorActions.PinListItems(ThisWorkbook.Worksheets("_Edit"), mConnectorID)
    If IsEmpty(vItems) Then Exit Sub

    For i = LBound(vItems, 1) To UBound(vItems, 1)
        lstPins.AddItem CStr(vItems(i, 1))
        AddMarkerControl CLng(vItems(i, 2)), CDbl(vItems(i, 3)), CDbl(vItems(i, 4))
    Next i
End Sub

Private Function SelectedPinNumber() As Long
    Dim vItems As Variant
    If lstPins.ListIndex < 0 Then Exit Function

    vItems = modEditorActions.PinListItems(ThisWorkbook.Worksheets("_Edit"), mConnectorID)
    If IsEmpty(vItems) Then Exit Function
    If lstPins.ListIndex + 1 > modContract.TableRowCount(vItems) Then Exit Function

    SelectedPinNumber = CLng(vItems(LBound(vItems, 1) + lstPins.ListIndex, 2))
End Function

Private Sub cmdLoadPhoto_Click()
    Dim sName As String, sPart As String, vResult As Variant
    sName = Trim$(txtName.Text)
    sPart = Trim$(txtPartNumber.Text)

    vResult = modEditorActions.CanLoadPhoto(sName, sPart)
    If Not modContract.Ok(vResult) Then
        MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
        Exit Sub
    End If

    Dim sPath As String
    sPath = Application.GetOpenFilename(modEditorActions.PhotoFileFilter())
    If sPath = "False" Then Exit Sub

    ShowPhoto sPath
    mConnectorID = modLibrary.SlugifyConnectorID(sPart, sName)
End Sub

Private Sub imgPhoto_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim dNormX As Double, dNormY As Double
    Dim bPlaceMode As Boolean, nSelected As Long, sPinCountText As String
    Dim vResult As Variant

    If imgPhoto.Width <= 0 Or imgPhoto.Height <= 0 Then Exit Sub

    ' Capture every control value before the action call.
    dNormX = X / imgPhoto.Width
    dNormY = Y / imgPhoto.Height
    bPlaceMode = (tglPlacePins.Value = True)
    nSelected = SelectedPinNumber()
    sPinCountText = txtPinCount.Text

    vResult = modEditorActions.PhotoClickAction(ThisWorkbook.Worksheets("_Edit"), _
        mConnectorID, bPlaceMode, nSelected, sPinCountText, dNormX, dNormY)

    Select Case modContract.Outcome(vResult)
        Case "PLACED"
            RebuildPinList
        Case "MOVED_ANCHOR"
            RebuildPinList
        Case "NO_OP"
            ' Nothing to say: no connector yet, or nothing selected.
        Case Else
            MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
    End Select
End Sub

Private Sub AddMarkerControl(ByVal nPinNumber As Long, ByVal dNormX As Double, _
                             ByVal dNormY As Double)
    Dim lbl As MSForms.Label, vPos As Variant

    Set lbl = Me.Controls.Add("Forms.Label.1", _
        modEditorActions.MarkerControlName(nPinNumber), True)
    lbl.Caption = CStr(nPinNumber)
    lbl.Width = PIN_MARKER_SIZE
    lbl.Height = PIN_MARKER_SIZE
    lbl.Font.Size = 8
    lbl.TextAlign = fmTextAlignCenter
    lbl.BackStyle = fmBackStyleOpaque
    lbl.BackColor = RGB(255, 255, 255)
    lbl.BorderStyle = fmBorderStyleSingle

    vPos = modPinEditor.MarkerTopLeft(dNormX, dNormY, _
        CDbl(imgPhoto.Left), CDbl(imgPhoto.Top), _
        CDbl(imgPhoto.Width), CDbl(imgPhoto.Height), _
        CDbl(lbl.Width), CDbl(lbl.Height))
    lbl.Left = vPos(0)
    lbl.Top = vPos(1)

    Dim marker As clsPinMarker
    Set marker = New clsPinMarker
    marker.PinNumber = nPinNumber
    marker.ConnectorID = mConnectorID
    Set marker.ScratchSheet = ThisWorkbook.Worksheets("_Edit")
    Set marker.PhotoControl = imgPhoto
    Set marker.LabelControl = lbl
    mMarkers.Add marker, CStr(nPinNumber)
End Sub

Private Sub ClearMarkerControls()
    Dim marker As Variant
    For Each marker In mMarkers
        On Error Resume Next
        Me.Controls.Remove marker.LabelControl.Name
        On Error GoTo 0
    Next marker
    Set mMarkers = New Collection
End Sub

Private Sub cmdDeletePin_Click()
    Dim nPinNumber As Long
    nPinNumber = SelectedPinNumber()
    If nPinNumber = 0 Then Exit Sub

    modEditorActions.DeletePinRequest ThisWorkbook.Worksheets("_Edit"), mConnectorID, nPinNumber
    RebuildPinList
End Sub

Private Sub cmdClearPins_Click()
    modPinEditor.ClearScratchPins ThisWorkbook.Worksheets("_Edit")
    RebuildPinList
End Sub

Private Sub cmdSnapLabel_Click()
    Dim nPinNumber As Long
    nPinNumber = SelectedPinNumber()
    If nPinNumber = 0 Then Exit Sub

    If modPinEditor.SnapLabelToPin(ThisWorkbook.Worksheets("_Edit"), mConnectorID, nPinNumber) Then
        RebuildPinList
    End If
End Sub

Private Sub cmdSave_Click()
    Dim sID As String, sOriginal As String, sPhoto As String
    Dim vFields As Variant, vResult As Variant

    ' Capture every control value and form variable before the action call.
    sID = mConnectorID
    sOriginal = mOriginalConnectorID
    sPhoto = mPhotoPath
    vFields = Array(Trim$(txtName.Text), Trim$(txtManufacturer.Text), _
                    Trim$(txtPartNumber.Text), cboType.Text, _
                    txtPinCount.Text, Trim$(txtNotes.Text))

    Dim lib As Workbook
    Set lib = Workbooks.Open(ThisWorkbook.Path & "\ConnectorLibrary.xlsx")

    vResult = modEditorActions.SaveFromEditor( _
        lib.Worksheets("Connectors"), lib.Worksheets("Pins"), lib.Worksheets("Photos"), _
        ThisWorkbook.Worksheets("_Edit"), sID, sOriginal, vFields, sPhoto, _
        Format$(Now, "yyyy-mm-ddThh:mm:ssZ"))

    If modContract.Ok(vResult) Then lib.Save
    lib.Close SaveChanges:=False

    If Not modContract.Ok(vResult) Then
        MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
        Exit Sub
    End If

    Dim sCopyTo As String
    sCopyTo = modEditorActions.PhotoCacheRefreshTarget(ThisWorkbook.Path, sID, sPhoto)
    If Len(sCopyTo) > 0 Then FileCopy sPhoto, sCopyTo

    ' A standard module's variable, not a property on this form: this form
    ' unloads itself right below, and re-reading a property off a
    ' just-unloaded predeclared instance risks re-triggering
    ' UserForm_Initialize instead of returning the value just set.
    modConnectorUI.LastSavedConnectorID = sID
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    modPinEditor.ClearScratchPins ThisWorkbook.Worksheets("_Edit")
    Unload Me
End Sub
```

- [ ] **Step 2: Rewire clsPinMarker's drop handler**

In `src/vba/clsPinMarker.cls`, replace `mLabel_MouseUp` (lines 46-56) with:

```vba
Private Sub mLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                           ByVal X As Single, ByVal Y As Single)
    If Not mDragging Then Exit Sub
    mDragging = False

    Dim vNorm As Variant
    vNorm = modPinEditor.NormFromMarker( _
        CDbl(mLabel.Left), CDbl(mLabel.Top), CDbl(mLabel.Width), CDbl(mLabel.Height), _
        CDbl(PhotoControl.Left), CDbl(PhotoControl.Top), _
        CDbl(PhotoControl.Width), CDbl(PhotoControl.Height))
    If IsEmpty(vNorm) Then Exit Sub

    modPinEditor.MoveMarker ScratchSheet, ConnectorID, PinNumber, vNorm(0), vNorm(1)
End Sub
```

- [ ] **Step 3: Delete the retired grep assertions**

From `tests/test_connector_editor_wiring.py`, delete these test functions entirely:
`test_pin_marker_handles_mouse_drag_events`, `test_pin_marker_calls_move_marker_on_drop`,
`test_pin_marker_label_control_is_readable`, `test_form_wires_every_command`,
`test_form_calls_into_the_tested_logic_module_for_every_pin_action`,
`test_form_tracks_pin_numbers_independently_of_list_position`,
`test_load_photo_filter_excludes_png`,
`test_type_combo_is_populated_from_this_workbooks_lists_sheet`,
`test_load_existing_photo_prefers_the_disk_cache_over_the_embedded_shape`,
`test_save_refreshes_the_photo_cache_via_plain_file_copy`,
`test_load_photo_requires_name_and_part_number_first`,
`test_place_pins_is_capped_at_pin_count`,
`test_save_rejects_a_part_number_that_collides_with_another_connector`.

Keep `test_pin_marker_class_exists`, `test_pin_marker_is_a_class_module`,
`test_pin_marker_uses_a_small_fixed_badge_size`, and
`test_photo_fit_box_is_a_fixed_constant`. Remove the now-unused `import pytest`
only if no remaining test uses it.

- [ ] **Step 4: Run the full suite**

Run: `python -m pytest -v`
Expected: all PASS. The form must still compile - a VBA compile error surfaces as a build failure in the `artifact` fixture, failing every test at once.

- [ ] **Step 5: Commit**

```bash
git add src/vba/forms/frmConnectorEditor.evt src/vba/clsPinMarker.cls tests/test_connector_editor_wiring.py
git commit -m "refactor: reduce the connector editor form to a UI adapter"
```

---

### Task 11: modPickerActions and frmConnectorPicker

**Files:**
- Create: `src/vba/modPickerActions.bas`
- Modify: `build/build.py:21-25`, `src/vba/forms/frmConnectorPicker.evt`
- Test: `tests/test_picker_actions.py`

**Interfaces:**
- Consumes: `modContract` (Task 1), `modLibrary.ConnectorIndex` (Task 5), `modLibrary.ReadConnector`, `modConnectors.AddConnectorInstance`, `modSnapshot.SnapshotConnector` (existing).
- Produces: `modPickerActions.AddFromLibrary(wsSnapshot, wsLibConn, wsLibPins, wsLibPhotos, sConnectorID) As Variant` → `ADDED` (payload the ref des) | `CONNECTOR_NOT_FOUND` | `ADD_FAILED`.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_picker_actions.py`:

```python
from tests.conftest import run, run_action
from tests.fixtures.sample_photo import write_sample_photo


def seed(wb, library_wb, tmp_path, connector_id="DTM-04P", pin_count=2):
    photo = write_sample_photo(tmp_path / "photo.png")
    ws_conn = library_wb.Worksheets("Connectors")
    fields = (connector_id, "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
              pin_count, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields)
    for n in range(1, pin_count + 1):
        run(wb, "modLibrary.WritePin", library_wb.Worksheets("Pins"), 2, 100000,
            (connector_id, n, f"Pin {n}", 0.1, 0.1, 0.1, 0.1))
    shape = run(wb, "modLibrary.EmbedConnectorPhoto",
                library_wb.Worksheets("Photos"), connector_id, str(photo))
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000,
        fields[:7] + (shape,) + fields[8:])


def add(wb, library_wb, connector_id):
    return run_action(
        wb, "modPickerActions.AddFromLibrary",
        wb.Worksheets("_Snapshot"), library_wb.Worksheets("Connectors"),
        library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"), connector_id,
    )


def test_add_from_library_creates_an_instance_and_snapshots_it(wb, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)

    result = add(wb, library_wb, "DTM-04P")
    assert result.ok is True
    assert result.outcome == "ADDED"
    assert result.payload == "J1"

    snapshot = run(wb, "modLibrary.ReadConnector", wb.Worksheets("_Snapshot"), 2, 201, "DTM-04P")
    assert snapshot[1] == "Deutsch DTM 4-way"


def test_adding_twice_allocates_sequential_ref_designators(wb, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)
    assert add(wb, library_wb, "DTM-04P").payload == "J1"
    assert add(wb, library_wb, "DTM-04P").payload == "J2"


def test_adding_an_unknown_connector_reports_not_found(wb, library_wb):
    result = add(wb, library_wb, "NOPE")
    assert (result.ok, result.outcome, result.payload) == (False, "CONNECTOR_NOT_FOUND", "NOPE")
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_picker_actions.py -v`
Expected: FAIL - `modPickerActions` does not exist.

- [ ] **Step 3: Create the module**

Create `src/vba/modPickerActions.bas`:

```vba
Attribute VB_Name = "modPickerActions"
Option Explicit

' Adding a library connector to this harness: allocate a ref des, write the
' instance row, and freeze the definition into _Snapshot. One transaction,
' called from both the picker's Add button and its New-then-save chain,
' which previously duplicated it.
Public Function AddFromLibrary(wsSnapshot As Worksheet, wsLibConn As Worksheet, _
                               wsLibPins As Worksheet, wsLibPhotos As Worksheet, _
                               ByVal sConnectorID As String) As Variant
    Dim vFields As Variant, sRefDes As String

    vFields = modLibrary.ReadConnector(wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vFields) Then
        AddFromLibrary = modContract.Failure("CONNECTOR_NOT_FOUND", sConnectorID)
        Exit Function
    End If

    sRefDes = modConnectors.AddConnectorInstance(CStr(vFields(1)), CStr(vFields(2)), _
        CStr(vFields(4)), CStr(vFields(5)), CLng(vFields(6)))
    If Len(sRefDes) = 0 Then
        AddFromLibrary = modContract.Failure("ADD_FAILED", sConnectorID)
        Exit Function
    End If

    modSnapshot.SnapshotConnector wsSnapshot, wsLibConn, wsLibPins, wsLibPhotos, sConnectorID

    AddFromLibrary = modContract.Success("ADDED", sRefDes)
End Function
```

- [ ] **Step 4: Register the module and rewire the form**

Append `"modPickerActions.bas"` to `VBA_MODULES` in `build/build.py`.

Replace `src/vba/forms/frmConnectorPicker.evt` entirely:

```vba
Option Explicit

Private mLibrary As Workbook
Private mConnectorIDs() As String

Private Sub UserForm_Initialize()
    Set mLibrary = Workbooks.Open(ThisWorkbook.Path & "\ConnectorLibrary.xlsx")
    RefreshList
End Sub

' QueryClose fires for every unload path - a button's Unload Me as well as
' the form's own system Close button - so this is the one place that needs
' to close mLibrary.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    mLibrary.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub RefreshList()
    Dim vIndex As Variant, i As Long, n As Long

    lstConnectors.Clear
    Erase mConnectorIDs

    vIndex = modLibrary.ConnectorIndex(mLibrary.Worksheets("Connectors"))
    n = modContract.TableRowCount(vIndex)
    If n = 0 Then Exit Sub

    ReDim mConnectorIDs(1 To n)
    For i = 1 To n
        lstConnectors.AddItem CStr(vIndex(i, 1))
        mConnectorIDs(i) = CStr(vIndex(i, 2))
    Next i
End Sub

Private Sub cmdAdd_Click()
    If lstConnectors.ListIndex < 0 Then Exit Sub

    Dim sConnectorID As String, vResult As Variant
    sConnectorID = mConnectorIDs(lstConnectors.ListIndex + 1)

    vResult = modPickerActions.AddFromLibrary(ThisWorkbook.Worksheets("_Snapshot"), _
        mLibrary.Worksheets("Connectors"), mLibrary.Worksheets("Pins"), _
        mLibrary.Worksheets("Photos"), sConnectorID)

    If Not modContract.Ok(vResult) Then
        MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
        Exit Sub
    End If

    Unload Me
End Sub

Private Sub cmdNew_Click()
    ' Cleared first so a stale value from an earlier save can never cause
    ' this to wrongly re-add an old connector after a Cancel.
    modConnectorUI.LastSavedConnectorID = ""
    frmConnectorEditor.Show

    Dim sConnectorID As String
    sConnectorID = modConnectorUI.LastSavedConnectorID
    If Len(sConnectorID) > 0 Then
        ' frmConnectorEditor's Save closed mLibrary (Workbooks.Open on an
        ' already-open path returns the same object) - reopen to read what
        ' it just wrote.
        Set mLibrary = Workbooks.Open(ThisWorkbook.Path & "\ConnectorLibrary.xlsx")

        Dim vResult As Variant
        vResult = modPickerActions.AddFromLibrary(ThisWorkbook.Worksheets("_Snapshot"), _
            mLibrary.Worksheets("Connectors"), mLibrary.Worksheets("Pins"), _
            mLibrary.Worksheets("Photos"), sConnectorID)
        If Not modContract.Ok(vResult) Then
            MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
        End If
    End If

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
```

- [ ] **Step 5: Run the full suite**

Run: `python -m pytest -v`
Expected: all PASS, including the existing `tests/test_picker_form.py`.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modPickerActions.bas src/vba/forms/frmConnectorPicker.evt build/build.py tests/test_picker_actions.py
git commit -m "feat: extract the add-from-library transaction the picker duplicated"
```

---

### Task 12: modManageActions - delete and export

**Files:**
- Create: `src/vba/modManageActions.bas`
- Modify: `build/build.py:21-25`
- Test: `tests/test_manage_actions.py`

**Interfaces:**
- Consumes: `modContract` (Task 1), `modLibrary.DeleteConnector`, `DeletePinsForConnector`, `RemoveConnectorPhoto`, `CachePhotoPath`, `modLibraryTransfer.BuildExportSheets`, `ExportConnector` (all existing).
- Produces: `DeleteFromLibrary(wsLibConn, wsLibPins, wsLibPhotos, sWorkbookPath, sConnectorID) As Variant` → `CONNECTOR_DELETED`; `ExportToWorkbook(wsLibConn, wsLibPins, wsLibPhotos, destWb, sConnectorID) As Variant` → `EXPORTED` | `EXPORT_FAILED`.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_manage_actions.py`:

```python
from pathlib import Path

from tests.conftest import run, run_action
from tests.fixtures.sample_photo import write_sample_photo


def seed(wb, library_wb, tmp_path, connector_id="DTM-04P"):
    photo = write_sample_photo(tmp_path / f"{connector_id}.png")
    ws_conn = library_wb.Worksheets("Connectors")
    fields = (connector_id, "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
              2, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields)
    run(wb, "modLibrary.WritePin", library_wb.Worksheets("Pins"), 2, 100000,
        (connector_id, 1, "Pin 1", 0.1, 0.1, 0.1, 0.1))
    shape = run(wb, "modLibrary.EmbedConnectorPhoto",
                library_wb.Worksheets("Photos"), connector_id, str(photo))
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000,
        fields[:7] + (shape,) + fields[8:])


def test_delete_removes_the_row_the_pins_and_the_cache_file(wb, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)
    cache = run(wb, "modLibrary.CachePhotoPath", str(tmp_path), "DTM-04P", "jpg")
    Path(cache).write_bytes(b"cached preview")

    result = run_action(
        wb, "modManageActions.DeleteFromLibrary",
        library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
        library_wb.Worksheets("Photos"), str(tmp_path), "DTM-04P",
    )
    assert (result.ok, result.outcome, result.payload) == (True, "CONNECTOR_DELETED", "DTM-04P")

    assert run(wb, "modLibrary.ReadConnector",
               library_wb.Worksheets("Connectors"), 2, 100000, "DTM-04P") is None
    assert run(wb, "modLibrary.ReadPinsForConnector",
               library_wb.Worksheets("Pins"), 2, 100000, "DTM-04P") is None
    assert not Path(cache).exists()


def test_delete_succeeds_when_no_cache_file_was_ever_written(wb, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)
    result = run_action(
        wb, "modManageActions.DeleteFromLibrary",
        library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
        library_wb.Worksheets("Photos"), str(tmp_path), "DTM-04P",
    )
    assert result.outcome == "CONNECTOR_DELETED"


def test_export_builds_the_sheets_and_copies_the_record(wb, app, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path)
    dest = app.Workbooks.Add()
    try:
        result = run_action(
            wb, "modManageActions.ExportToWorkbook",
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"), dest, "DTM-04P",
        )
        assert (result.ok, result.outcome, result.payload) == (True, "EXPORTED", "DTM-04P")

        row = run(wb, "modLibrary.ReadConnector", dest.Worksheets("Connectors"), 2, 100000, "DTM-04P")
        assert row[1] == "Deutsch DTM 4-way"
    finally:
        dest.Close(SaveChanges=False)


def test_exporting_an_unknown_connector_fails(wb, app, library_wb):
    dest = app.Workbooks.Add()
    try:
        result = run_action(
            wb, "modManageActions.ExportToWorkbook",
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"), dest, "NOPE",
        )
        assert (result.ok, result.outcome, result.payload) == (False, "EXPORT_FAILED", "NOPE")
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_manage_actions.py -v`
Expected: FAIL - `modManageActions` does not exist.

- [ ] **Step 3: Create the module**

Create `src/vba/modManageActions.bas`:

```vba
Attribute VB_Name = "modManageActions"
Option Explicit

' Removing a connector everywhere it exists: the three library sheets plus
' the editor's on-disk preview cache, which would otherwise be orphaned -
' nothing ever reads a cache file whose connector is gone.
Public Function DeleteFromLibrary(wsLibConn As Worksheet, wsLibPins As Worksheet, _
                                  wsLibPhotos As Worksheet, ByVal sWorkbookPath As String, _
                                  ByVal sConnectorID As String) As Variant
    Dim sCachePath As String

    modLibrary.DeleteConnector wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID
    modLibrary.DeletePinsForConnector wsLibPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID
    modLibrary.RemoveConnectorPhoto wsLibPhotos, sConnectorID

    sCachePath = modLibrary.CachePhotoPath(sWorkbookPath, sConnectorID, "jpg")
    If Len(Dir$(sCachePath)) > 0 Then Kill sCachePath

    DeleteFromLibrary = modContract.Success("CONNECTOR_DELETED", sConnectorID)
End Function

' destWb is created by the adapter (Workbooks.Add) and saved by it
' afterwards; this shapes the sheets and copies the record into them.
Public Function ExportToWorkbook(wsLibConn As Worksheet, wsLibPins As Worksheet, _
                                 wsLibPhotos As Worksheet, destWb As Workbook, _
                                 ByVal sConnectorID As String) As Variant
    modLibraryTransfer.BuildExportSheets destWb

    If modLibraryTransfer.ExportConnector(wsLibConn, wsLibPins, wsLibPhotos, _
            destWb.Worksheets("Connectors"), destWb.Worksheets("Pins"), _
            destWb.Worksheets("Photos"), sConnectorID) Then
        ExportToWorkbook = modContract.Success("EXPORTED", sConnectorID)
    Else
        ExportToWorkbook = modContract.Failure("EXPORT_FAILED", sConnectorID)
    End If
End Function
```

- [ ] **Step 4: Register the module with the build**

Append `"modManageActions.bas"` to `VBA_MODULES` in `build/build.py`.

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m pytest tests/test_manage_actions.py -v`
Expected: all 4 PASS.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modManageActions.bas build/build.py tests/test_manage_actions.py
git commit -m "feat: extract the library delete and export transactions"
```

---

### Task 13: modManageActions - import, and rewire frmManageLibrary

**Files:**
- Modify: `src/vba/modManageActions.bas`, `src/vba/forms/frmManageLibrary.evt`
- Modify: `tests/test_manage_actions.py` (extend), delete `tests/test_manage_library_transfer_wiring.py`

**Interfaces:**
- Consumes: Task 12's module, `modLibraryTransfer.ImportConnector`, `modLibrary.EmbedConnectorPhoto`, `modLibrary.LIB_COL_ID` (existing).
- Produces: `ImportAllFromWorkbook(srcWb, wsLibConn, wsLibPins, wsLibPhotos) As Variant` → `IMPORTED`, payload a 2D `(1 To n, 1 To 2)` table of destination ID and a photo-ok Boolean; `AttachReplacementPhoto(wsLibPhotos, sDestID, sPath) As Variant` → `PHOTO_ATTACHED` | `PHOTO_FAILED`.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_manage_actions.py`:

```python
def test_import_copies_every_connector_and_reports_photo_status(wb, app, library_wb, tmp_path):
    seed(wb, library_wb, tmp_path, "DTM-04P")
    export_path = tmp_path / "export.xlsx"
    dest = app.Workbooks.Add()
    try:
        run_action(wb, "modManageActions.ExportToWorkbook",
                   library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
                   library_wb.Worksheets("Photos"), dest, "DTM-04P")
        dest.SaveAs(Filename=str(export_path), FileFormat=51)
    finally:
        dest.Close(SaveChanges=False)

    src = app.Workbooks.Open(str(export_path))
    try:
        # Import into a library that does not yet hold this connector.
        run(wb, "modLibrary.DeleteConnector", library_wb.Worksheets("Connectors"),
            2, 100000, "DTM-04P")

        result = run_action(
            wb, "modManageActions.ImportAllFromWorkbook", src,
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"),
        )
        assert (result.ok, result.outcome) == (True, "IMPORTED")
        assert [row[0] for row in result.payload] == ["DTM-04P"]
        assert isinstance(result.payload[0][1], bool)
    finally:
        src.Close(SaveChanges=False)


def test_import_of_an_empty_workbook_reports_an_empty_table(wb, app, library_wb, tmp_path):
    empty_path = tmp_path / "empty.xlsx"
    dest = app.Workbooks.Add()
    try:
        run(wb, "modLibraryTransfer.BuildExportSheets", dest)
        dest.SaveAs(Filename=str(empty_path), FileFormat=51)
    finally:
        dest.Close(SaveChanges=False)

    src = app.Workbooks.Open(str(empty_path))
    try:
        result = run_action(
            wb, "modManageActions.ImportAllFromWorkbook", src,
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"),
            library_wb.Worksheets("Photos"),
        )
        assert result.outcome == "IMPORTED"
        assert run(wb, "modContract.TableRowCount", result.payload) == 0
    finally:
        src.Close(SaveChanges=False)


def test_attach_replacement_photo_embeds_the_chosen_file(wb, library_wb, tmp_path):
    photo = write_sample_photo(tmp_path / "replacement.png")
    result = run_action(wb, "modManageActions.AttachReplacementPhoto",
                        library_wb.Worksheets("Photos"), "DTM-04P", str(photo))
    assert (result.ok, result.outcome, result.payload) == (True, "PHOTO_ATTACHED", "DTM-04P")


def test_attach_replacement_photo_reports_a_missing_file(wb, library_wb, tmp_path):
    result = run_action(wb, "modManageActions.AttachReplacementPhoto",
                        library_wb.Worksheets("Photos"), "DTM-04P",
                        str(tmp_path / "does-not-exist.png"))
    assert (result.ok, result.outcome, result.payload) == (False, "PHOTO_FAILED", "DTM-04P")
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_manage_actions.py -k "import or replacement" -v`
Expected: FAIL - the two functions do not exist.

- [ ] **Step 3: Append the two functions**

Add to `src/vba/modManageActions.bas`:

```vba
' Every connector in a shared export file, copied into this library.
' modLibraryTransfer.ImportConnector renames on an ID collision and
' attempts the photo copy itself; the photo copy goes through the
' clipboard and is not reliable, so rather than redoing it here (which
' could give a different answer than the one that actually happened) this
' checks whether the shape it should have produced exists, and reports
' per connector. The adapter prompts for a replacement where it did not.
Public Function ImportAllFromWorkbook(srcWb As Workbook, wsLibConn As Worksheet, _
                                      wsLibPins As Worksheet, wsLibPhotos As Worksheet) As Variant
    Dim wsSrcConn As Worksheet, nLast As Long, r As Long, n As Long
    Dim sConnectorID As String, sDestID As String, sOriginName As String
    Dim vRows() As Variant, bPhotoOk As Boolean

    Set wsSrcConn = srcWb.Worksheets("Connectors")
    sOriginName = srcWb.Name
    nLast = wsSrcConn.Cells(wsSrcConn.Rows.Count, modLibrary.LIB_COL_ID).End(xlUp).Row
    If nLast < 2 Then
        ImportAllFromWorkbook = modContract.Success("IMPORTED", Empty)
        Exit Function
    End If

    ReDim vRows(1 To nLast - 1, 1 To 2)
    For r = 2 To nLast
        sConnectorID = Trim$(CStr(wsSrcConn.Cells(r, modLibrary.LIB_COL_ID).Value))
        If Len(sConnectorID) > 0 Then
            sDestID = modLibraryTransfer.ImportConnector(wsSrcConn, _
                srcWb.Worksheets("Pins"), srcWb.Worksheets("Photos"), _
                wsLibConn, wsLibPins, wsLibPhotos, sConnectorID, sOriginName)

            If Len(sDestID) > 0 Then
                bPhotoOk = False
                On Error Resume Next
                bPhotoOk = Not (wsLibPhotos.Shapes("PHOTO_" & sDestID) Is Nothing)
                On Error GoTo 0

                n = n + 1
                vRows(n, 1) = sDestID
                vRows(n, 2) = bPhotoOk
            End If
        End If
    Next r

    If n = 0 Then
        ImportAllFromWorkbook = modContract.Success("IMPORTED", Empty)
        Exit Function
    End If

    Dim vResult() As Variant, i As Long
    ReDim vResult(1 To n, 1 To 2)
    For i = 1 To n
        vResult(i, 1) = vRows(i, 1)
        vResult(i, 2) = vRows(i, 2)
    Next i
    ImportAllFromWorkbook = modContract.Success("IMPORTED", vResult)
End Function

Public Function AttachReplacementPhoto(wsLibPhotos As Worksheet, ByVal sDestID As String, _
                                       ByVal sPath As String) As Variant
    If Len(modLibrary.EmbedConnectorPhoto(wsLibPhotos, sDestID, sPath)) > 0 Then
        AttachReplacementPhoto = modContract.Success("PHOTO_ATTACHED", sDestID)
    Else
        AttachReplacementPhoto = modContract.Failure("PHOTO_FAILED", sDestID)
    End If
End Function
```

- [ ] **Step 4: Rewire the form**

Replace `src/vba/forms/frmManageLibrary.evt` entirely:

```vba
Option Explicit

Private mLibrary As Workbook
Private mConnectorIDs() As String

Private Sub UserForm_Initialize()
    Set mLibrary = Workbooks.Open(ThisWorkbook.Path & "\ConnectorLibrary.xlsx")
    RefreshList
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    mLibrary.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub RefreshList()
    Dim vIndex As Variant, i As Long, n As Long

    lstConnectors.Clear
    Erase mConnectorIDs

    vIndex = modLibrary.ConnectorIndex(mLibrary.Worksheets("Connectors"))
    n = modContract.TableRowCount(vIndex)
    If n = 0 Then Exit Sub

    ReDim mConnectorIDs(1 To n)
    For i = 1 To n
        lstConnectors.AddItem CStr(vIndex(i, 1))
        mConnectorIDs(i) = CStr(vIndex(i, 2))
    Next i
End Sub

Private Function SelectedConnectorID() As String
    If lstConnectors.ListIndex < 0 Then Exit Function
    SelectedConnectorID = mConnectorIDs(lstConnectors.ListIndex + 1)
End Function

Private Sub cmdEdit_Click()
    Dim sConnectorID As String, vFields As Variant
    sConnectorID = SelectedConnectorID()
    If Len(sConnectorID) = 0 Then Exit Sub

    vFields = modLibrary.ReadConnector(mLibrary.Worksheets("Connectors"), 2, _
                                       modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vFields) Then Exit Sub

    modPinEditor.LoadScratchPins ThisWorkbook.Worksheets("_Edit"), _
        mLibrary.Worksheets("Pins"), sConnectorID

    ' Load (not Show) runs UserForm_Initialize so LoadForEdit can populate
    ' the form before it is shown. Unloading this form first was observed
    ' to discard everything LoadForEdit had just written, so Show (modal,
    ' blocking) comes first and Unload Me stays last.
    Load frmConnectorEditor
    frmConnectorEditor.LoadForEdit sConnectorID, vFields, mLibrary.Worksheets("Photos")
    frmConnectorEditor.Show

    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim sConnectorID As String, vResult As Variant
    sConnectorID = SelectedConnectorID()
    If Len(sConnectorID) = 0 Then Exit Sub

    If MsgBox("Delete " & sConnectorID & " from the library? This cannot be undone.", _
              vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    vResult = modManageActions.DeleteFromLibrary(mLibrary.Worksheets("Connectors"), _
        mLibrary.Worksheets("Pins"), mLibrary.Worksheets("Photos"), _
        ThisWorkbook.Path, sConnectorID)
    mLibrary.Save

    RefreshList
End Sub

Private Sub cmdExport_Click()
    Dim sConnectorID As String, sPath As Variant, vResult As Variant
    sConnectorID = SelectedConnectorID()
    If Len(sConnectorID) = 0 Then Exit Sub

    sPath = Application.GetSaveAsFilename( _
        InitialFileName:=sConnectorID & "_export.xlsx", _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If sPath = False Then Exit Sub

    Dim destWb As Workbook
    Set destWb = Workbooks.Add

    vResult = modManageActions.ExportToWorkbook(mLibrary.Worksheets("Connectors"), _
        mLibrary.Worksheets("Pins"), mLibrary.Worksheets("Photos"), destWb, sConnectorID)

    If modContract.Ok(vResult) Then destWb.SaveAs Filename:=CStr(sPath), FileFormat:=51
    destWb.Close SaveChanges:=False

    MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
End Sub

Private Sub cmdImport_Click()
    Dim sPath As Variant
    sPath = Application.GetOpenFilename(FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If sPath = False Then Exit Sub

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(CStr(sPath))

    Dim vResult As Variant
    vResult = modManageActions.ImportAllFromWorkbook(srcWb, _
        mLibrary.Worksheets("Connectors"), mLibrary.Worksheets("Pins"), _
        mLibrary.Worksheets("Photos"))
    srcWb.Close SaveChanges:=False

    ' Prompt only for the connectors whose photo could not be carried over.
    Dim vRows As Variant, i As Long, sDestID As String, sReplacement As Variant
    vRows = modContract.Payload(vResult)
    For i = 1 To modContract.TableRowCount(vRows)
        If Not CBool(vRows(i, 2)) Then
            sDestID = CStr(vRows(i, 1))
            sReplacement = Application.GetOpenFilename( _
                "Pictures (*.png; *.jpg; *.jpeg; *.bmp), *.png;*.jpg;*.jpeg;*.bmp", _
                , "Photo for " & sDestID & " could not be extracted - choose a replacement")
            If sReplacement <> False Then
                modManageActions.AttachReplacementPhoto mLibrary.Worksheets("Photos"), _
                    sDestID, CStr(sReplacement)
            End If
        End If
    Next i

    mLibrary.Save
    RefreshList
    MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
```

- [ ] **Step 5: Delete the retired wiring tests**

```bash
git rm tests/test_manage_library_transfer_wiring.py
```

- [ ] **Step 6: Run the full suite**

Run: `python -m pytest -v`
Expected: all PASS.

- [ ] **Step 7: Commit**

```bash
git add src/vba/modManageActions.bas src/vba/forms/frmManageLibrary.evt tests/test_manage_actions.py
git commit -m "refactor: reduce Manage Library to a UI adapter over tested transactions"
```

---

### Task 14: Sheet event logic

**Files:**
- Modify: `src/vba/modChart.bas`, `src/vba/modConnectors.bas`
- Modify: `src/vba/sheets/shHarness.evt`, `src/vba/sheets/shConnectors.evt`
- Test: `tests/test_sheet_event_logic.py`

**Interfaces:**
- Consumes: `modContract` (Task 1), `modChart.RebuildPinValidation`, `modChart.SetLengthUnits`, `modChart.CHART_FIRST_ROW`, `CHART_LAST_ROW`, `COL_FROM_CONN`, `COL_TO_CONN`, `modConnectors.RenameRefDes`, `modChart.RefreshChartRowsForConnector` (all existing).
- Produces: `modChart.ApplyHarnessEdit(wsHarness, rTarget) As Variant` → `BULK_REBUILT` (payload row count) | `CELLS_REBUILT` (payload cell count) | `NO_OP`; `modConnectors.ApplyConnectorEdit(rTarget, sPriorRefDes, nPriorRow) As Variant` → `RENAMED` (payload the new ref des) | `RENAME_REJECTED` (payload the value to restore) | `NO_RENAME`.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_sheet_event_logic.py`:

```python
from tests.conftest import run, run_action

# modChart.bas declarations: the chart occupies rows 7 to 1006, with the
# From-connector column at 1 and the To-connector column at 9.
CHART_FIRST_ROW, CHART_LAST_ROW = 7, 1006
COL_FROM_CONN, COL_TO_CONN = 1, 9


def test_a_bulk_edit_rebuilds_validation_for_the_clamped_row_range(wb):
    sheet = wb.Worksheets("Harness")
    # 3000 cells spanning rows 1 to 1000, so it exceeds the bulk threshold
    # and starts above the chart. Only rows 7 to 1000 should be rebuilt.
    target = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1000, 3))

    result = run_action(wb, "modChart.ApplyHarnessEdit", sheet, target)
    assert result.outcome == "BULK_REBUILT"
    assert result.payload == 1000 - CHART_FIRST_ROW + 1


def test_a_single_connector_cell_edit_rebuilds_that_row(wb):
    sheet = wb.Worksheets("Harness")
    target = sheet.Cells(CHART_FIRST_ROW, COL_FROM_CONN)

    result = run_action(wb, "modChart.ApplyHarnessEdit", sheet, target)
    assert (result.ok, result.outcome, result.payload) == (True, "CELLS_REBUILT", 1)


def test_an_edit_outside_the_connector_columns_rebuilds_nothing(wb):
    sheet = wb.Worksheets("Harness")
    target = sheet.Cells(CHART_FIRST_ROW, COL_TO_CONN - 1)

    result = run_action(wb, "modChart.ApplyHarnessEdit", sheet, target)
    assert (result.ok, result.outcome) == (False, "NO_OP")


def test_a_rename_is_applied_and_reported(wb):
    sheet = wb.Worksheets("Connectors")
    run(wb, "modConnectors.AddConnectorInstance", "DTM-04P", "DTM 4-way",
        "DTM06-4S", "Connector", 4)
    row = 2
    sheet.Cells(row, 1).Value = "J9"

    result = run_action(wb, "modConnectors.ApplyConnectorEdit",
                        sheet.Cells(row, 1), "J1", row)
    assert (result.ok, result.outcome, result.payload) == (True, "RENAMED", "J9")


def test_a_rename_onto_an_existing_ref_des_is_rejected_with_the_value_to_restore(wb):
    sheet = wb.Worksheets("Connectors")
    run(wb, "modConnectors.AddConnectorInstance", "DTM-04P", "DTM 4-way",
        "DTM06-4S", "Connector", 4)
    run(wb, "modConnectors.AddConnectorInstance", "DTM-04P", "DTM 4-way",
        "DTM06-4S", "Connector", 4)
    sheet.Cells(2, 1).Value = "J2"

    result = run_action(wb, "modConnectors.ApplyConnectorEdit",
                        sheet.Cells(2, 1), "J1", 2)
    assert (result.ok, result.outcome, result.payload) == (False, "RENAME_REJECTED", "J1")


def test_an_edit_that_is_not_a_rename_reports_no_rename(wb):
    sheet = wb.Worksheets("Connectors")
    run(wb, "modConnectors.AddConnectorInstance", "DTM-04P", "DTM 4-way",
        "DTM06-4S", "Connector", 4)

    result = run_action(wb, "modConnectors.ApplyConnectorEdit",
                        sheet.Cells(2, 3), "J1", 2)
    assert (result.ok, result.outcome) == (False, "NO_RENAME")
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_sheet_event_logic.py -v`
Expected: FAIL - neither function exists.

- [ ] **Step 3: Append ApplyHarnessEdit to modChart**

First add the threshold constant to the **declarations section** of
`src/vba/modChart.bas`, immediately after `Private Const MAX_FORMULA1 As Long = 255`
(around line 17). VBA does not recognise a module-level `Const` declared between
procedures - `modChart.bas:19-20` already carries a comment saying so, and appending
it at the end of the file is a compile error:

```vba
' Above this threshold an edit is a bulk clear or paste, with no per-cell
' edit worth reacting to.
Public Const BULK_EDIT_THRESHOLD As Long = 500
```

Then add the function at the end of the file:

```vba
' Everything shHarness's Worksheet_Change decides. A bulk clear or paste
' has no per-cell edit worth reacting to, but the pin dropdowns for the
' affected rows are still stale and must be rebuilt - without clearing pin
' values the same paste may have just set.
Public Function ApplyHarnessEdit(wsHarness As Worksheet, rTarget As Range) As Variant
    Dim cel As Range, nFirst As Long, nLast As Long, r As Long, n As Long

    If rTarget.Cells.Count > BULK_EDIT_THRESHOLD Then
        nFirst = rTarget.Row
        If nFirst < CHART_FIRST_ROW Then nFirst = CHART_FIRST_ROW
        nLast = rTarget.Row + rTarget.Rows.Count - 1
        If nLast > CHART_LAST_ROW Then nLast = CHART_LAST_ROW

        For r = nFirst To nLast
            RebuildPinValidation r, COL_FROM_CONN, False
            RebuildPinValidation r, COL_TO_CONN, False
            n = n + 1
        Next r
        modState.MarkDirty
        ApplyHarnessEdit = modContract.Success("BULK_REBUILT", n)
        Exit Function
    End If

    For Each cel In rTarget.Cells
        If cel.Row >= CHART_FIRST_ROW And cel.Row <= CHART_LAST_ROW Then
            If cel.Column = COL_FROM_CONN Or cel.Column = COL_TO_CONN Then
                RebuildPinValidation cel.Row, cel.Column
                n = n + 1
            End If
            modState.MarkDirty
        ElseIf cel.Row < CHART_HEADER_ROW Then
            If Not Application.Intersect(cel, ThisWorkbook.Names("TB_Units").RefersToRange) _
               Is Nothing Then
                SetLengthUnits CStr(cel.Value)
            End If
            modState.MarkDirty
        End If
    Next cel

    If n = 0 Then
        ApplyHarnessEdit = modContract.Failure("NO_OP")
    Else
        ApplyHarnessEdit = modContract.Success("CELLS_REBUILT", n)
    End If
End Function
```

- [ ] **Step 4: Append ApplyConnectorEdit to modConnectors**

Add to `src/vba/modConnectors.bas`:

```vba
' Everything shConnectors's Worksheet_Change decides. sPriorRefDes and
' nPriorRow come from the sheet module's SelectionChange bookkeeping, which
' is genuine event-lifecycle state and stays there. On rejection the
' payload is the value the caller must write back into the cell.
Public Function ApplyConnectorEdit(rTarget As Range, ByVal sPriorRefDes As String, _
                                   ByVal nPriorRow As Long) As Variant
    Dim rw As Range, sRef As String, sNewRefDes As String

    modState.MarkDirty

    For Each rw In rTarget.Rows
        If rw.Row >= CONN_FIRST_ROW Then
            sRef = Trim$(CStr(rTarget.Worksheet.Cells(rw.Row, 1).Value))
            modChart.RefreshChartRowsForConnector sRef
        End If
    Next rw

    If Not (rTarget.Cells.Count = 1 And rTarget.Column = 1 _
            And rTarget.Row = nPriorRow And Len(sPriorRefDes) > 0) Then
        ApplyConnectorEdit = modContract.Failure("NO_RENAME")
        Exit Function
    End If

    sNewRefDes = Trim$(CStr(rTarget.Value))
    If StrComp(sNewRefDes, sPriorRefDes, vbTextCompare) = 0 Then
        ApplyConnectorEdit = modContract.Failure("NO_RENAME")
        Exit Function
    End If

    If RenameRefDes(sPriorRefDes, sNewRefDes) Then
        ApplyConnectorEdit = modContract.Success("RENAMED", sNewRefDes)
    Else
        ApplyConnectorEdit = modContract.Failure("RENAME_REJECTED", sPriorRefDes)
    End If
End Function
```

- [ ] **Step 5: Reduce both sheet modules to adapters**

Replace `src/vba/sheets/shHarness.evt`:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim bEvents As Boolean

    If Not Application.EnableEvents Then Exit Sub

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    modChart.ApplyHarnessEdit Me, Target

CleanUp:
    Application.EnableEvents = bEvents
End Sub
```

Replace `src/vba/sheets/shConnectors.evt`:

```vba
Private mLastRefDes As String
Private mLastRefDesRow As Long

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count = 1 And Target.Column = 1 _
       And Target.Row >= modConnectors.CONN_FIRST_ROW Then
        mLastRefDesRow = Target.Row
        mLastRefDes = Trim$(CStr(Target.Value))
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim bEvents As Boolean, vResult As Variant

    If Not Application.EnableEvents Then Exit Sub

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    vResult = modConnectors.ApplyConnectorEdit(Target, mLastRefDes, mLastRefDesRow)

    Select Case modContract.Outcome(vResult)
        Case "RENAMED"
            mLastRefDes = CStr(modContract.Payload(vResult))
        Case "RENAME_REJECTED"
            Target.Value = CStr(modContract.Payload(vResult))
    End Select

CleanUp:
    Application.EnableEvents = bEvents
End Sub
```

- [ ] **Step 6: Run the full suite**

Run: `python -m pytest -v`
Expected: all PASS, including `tests/test_ref_des_rename.py` and `tests/test_pin_dropdown.py`, which exercise these paths through the sheet events.

- [ ] **Step 7: Commit**

```bash
git add src/vba/modChart.bas src/vba/modConnectors.bas src/vba/sheets/ tests/test_sheet_event_logic.py
git commit -m "refactor: extract worksheet change logic out of the sheet event modules"
```

---

### Task 15: Layering lints and dead code removal

**Files:**
- Create: `tests/test_layering.py`
- Modify: `src/vba/modConnectorUI.bas`, `src/vba/modState.bas`, `README.md`

**Interfaces:**
- Consumes: every module from Tasks 1-14.
- Produces: no callable surface. Deletes `modConnectorUI.RefreshConnectorList` and `modState.IsTestMode`.

- [ ] **Step 1: Write the failing lint tests**

Create `tests/test_layering.py`:

```python
import re

import pytest

LAYER0 = [
    "modUtil", "modState", "modLibrary", "modChart", "modConnectors",
    "modSnapshot", "modLibraryTransfer", "modPinEditor",
]
LAYER1 = [
    "modContract", "modMessages", "modEditorActions", "modPickerActions",
    "modManageActions",
]
ADAPTERS = ["frmConnectorEditor", "frmConnectorPicker", "frmManageLibrary",
            "clsPinMarker", "shHarness", "shConnectors"]
FORBIDDEN_IN_LAYER1 = [
    "MsgBox", "InputBox", "GetOpenFilename", "GetSaveAsFilename",
    ".Show", "Unload", "Workbooks.Open", "DoEvents", "MSForms",
]

# frmConnectorEditor exports a Shape to disk, which needs a live Shape object
# that cannot usefully cross Application.Run. Isolated in one named wrapper.
ALLOWED_LAYER0_IN_ADAPTERS = {
    "frmConnectorEditor": {
        "modPinEditor.FitAspectRatio", "modPinEditor.MarkerTopLeft",
        "modPinEditor.ClearScratchPins", "modPinEditor.SnapLabelToPin",
        "modLibrary.SlugifyConnectorID", "modLibrary.ExportShapeToFile",
    },
    "frmConnectorPicker": {"modLibrary.ConnectorIndex"},
    "frmManageLibrary": {
        "modLibrary.ConnectorIndex", "modLibrary.ReadConnector",
        "modLibrary.LIB_ROW_CAP", "modPinEditor.LoadScratchPins",
    },
    # The drag handler converts pixels to a normalized point and stores it.
    # Both are layer 0 primitives with their own tests; there is no
    # transaction here to lift into an action module.
    "clsPinMarker": {"modPinEditor.NormFromMarker", "modPinEditor.MoveMarker"},
    "shHarness": {"modChart.ApplyHarnessEdit"},
    "shConnectors": {"modConnectors.CONN_FIRST_ROW", "modConnectors.ApplyConnectorEdit"},
}

# Handlers that legitimately do no domain work: they only unload, or they
# seed scratch state and hand off to another form.
NON_DELEGATING_HANDLERS = {"cmdCancel_Click", "cmdClose_Click", "cmdEdit_Click"}


def module_source(wb, component_name):
    module = wb.VBProject.VBComponents(component_name).CodeModule
    return module.Lines(1, module.CountOfLines)


def all_sources(wb):
    return {name: module_source(wb, name) for name in LAYER0 + LAYER1 + ADAPTERS}


@pytest.mark.parametrize("module", LAYER1)
def test_action_modules_open_no_dialogs(wb, module):
    source = module_source(wb, module)
    for token in FORBIDDEN_IN_LAYER1:
        assert token not in source, f"{module} references {token}, which belongs in an adapter"


@pytest.mark.parametrize("adapter", ADAPTERS)
def test_adapters_call_only_permitted_layer0_members(wb, adapter):
    source = module_source(wb, adapter)
    allowed = ALLOWED_LAYER0_IN_ADAPTERS.get(adapter, set())
    for module in LAYER0:
        for match in re.finditer(rf"\b{module}\.(\w+)", source):
            reference = f"{module}.{match.group(1)}"
            assert reference in allowed, (
                f"{adapter} calls {reference} directly; route it through an action module "
                f"or add it to ALLOWED_LAYER0_IN_ADAPTERS with a reason"
            )


def test_no_doevents_anywhere(wb):
    for name, source in all_sources(wb).items():
        assert "DoEvents" not in source, f"{name} calls DoEvents, which lets a form unload mid-action"


@pytest.mark.parametrize("adapter", ADAPTERS)
def test_nothing_follows_unload_me(wb, adapter):
    # After Unload Me the controls are gone and form-level variables are
    # reset; reading either afterwards silently loses data.
    source = module_source(wb, adapter)
    for sub in re.split(r"\n(?=(?:Private|Public)\s+Sub\s)", source):
        lines = sub.splitlines()
        indexes = [i for i, line in enumerate(lines) if line.strip() == "Unload Me"]
        if not indexes:
            continue
        for line in lines[indexes[0] + 1:]:
            body = line.split("'")[0].strip()
            if not body or body in ("End Sub", "End If", "Else", "Exit Sub"):
                continue
            assert not re.search(r"\bm[A-Z]\w*", body), (
                f"{adapter}: '{body}' reads form state after Unload Me"
            )
            assert not re.search(r"\b(txt|lst|cbo|cmd|img|tgl)\w*\.", body), (
                f"{adapter}: '{body}' reads a control after Unload Me"
            )


@pytest.mark.parametrize("adapter", ["frmConnectorEditor", "frmConnectorPicker", "frmManageLibrary"])
def test_every_click_handler_delegates(wb, adapter):
    source = module_source(wb, adapter)
    for sub in re.split(r"\n(?=(?:Private|Public)\s+Sub\s)", source):
        header = sub.splitlines()[0] if sub.splitlines() else ""
        match = re.search(r"Sub\s+(cmd\w+_Click)", header)
        if not match:
            continue
        if match.group(1) in NON_DELEGATING_HANDLERS:
            continue
        assert re.search(r"\bmod(Editor|Picker|Manage)Actions\.", sub), \
            f"{adapter}: {match.group(1)} does no work through an action module"


def test_no_option_base_directive(wb):
    # Option Base 1 would make Array() one based, silently shifting every
    # result envelope index apart from a COM caller's view of it.
    for name, source in all_sources(wb).items():
        assert "Option Base" not in source, f"{name} declares Option Base"
```

- [ ] **Step 2: Run the lints to verify they fail or reveal violations**

Run: `python -m pytest tests/test_layering.py -v`
Expected: `test_adapters_call_only_permitted_layer0_members` FAILS for
`frmConnectorPicker` and `frmManageLibrary` while `modConnectorUI.RefreshConnectorList`
still exists and is referenced. Fix violations by adding the reference to
`ALLOWED_LAYER0_IN_ADAPTERS` only when it is genuinely irreducible; otherwise route it
through an action module.

- [ ] **Step 3: Delete the dead code**

In `src/vba/modConnectorUI.bas`, delete `RefreshConnectorList` entirely (lines 31-51 of the
original file) and the `MSForms` reference it carried. The module keeps only
`LastSavedConnectorID`, `ShowAddConnector`, `ShowManageLibrary` and `ShowRemoveConnector`.

In `src/vba/modState.bas`, delete `IsTestMode` (lines 56-58 of the original file). Leave
the `TestMode` row in `build/layout.py:225` alone - it is a state sheet key, not code, and
removing it would change the built sheet layout that `tests/test_state.py` asserts.

- [ ] **Step 4: Run the full suite**

Run: `python -m pytest -v`
Expected: all PASS.

- [ ] **Step 5: Update the README layout table**

In `README.md`, under `## Layout`, change the `src/vba/` row to describe the layering:

```markdown
| `src/vba/mod{Contract,Messages}.bas` | Result envelope and user-visible text |
| `src/vba/mod*Actions.bas` | User-intent transactions; the only thing forms may call |
| `src/vba/` | Layer 0 primitives; `sheets/*.evt` and `forms/*.evt` are UI adapters |
```

- [ ] **Step 6: Commit**

```bash
git add tests/test_layering.py src/vba/modConnectorUI.bas src/vba/modState.bas README.md
git commit -m "test: enforce the layering and handler lifecycle rules; drop dead code"
```

---

### Task 16: Manual verification pass

**Files:**
- Create: `docs/superpowers/plans/2026-08-28-ui-separation-manual-verification.md`

**Interfaces:**
- Consumes: the built `dist/HarnessCreator.xlsm` from Task 15.
- Produces: a completed checklist recording what was observed.

- [ ] **Step 1: Rebuild and open the workbook**

```bash
python build/build.py
```

Open `dist/HarnessCreator.xlsm` in Excel and enable macros.

- [ ] **Step 2: Write the checklist file**

Create `docs/superpowers/plans/2026-08-28-ui-separation-manual-verification.md` with these
items, and fill in the outcome column as each is exercised. Automated tests now cover every
rule; this pass only confirms that controls are wired to those rules and that the form
renders correctly.

```markdown
# UI Separation - Manual Verification

Date run:
Build:

| # | Check | Expected | Outcome |
|---|---|---|---|
| 1 | Home > Add Connector | Picker opens, list shows "ID - Name" for every library connector | |
| 2 | Pick one, Add | Connector row appears on Connectors, _Snapshot gains the definition | |
| 3 | Add the same one again | Second ref des allocated (J1 then J2) | |
| 4 | Picker > New | Editor opens with the Type dropdown populated | |
| 5 | Editor: fill Name and Part Number, Load Photo | File picker offers JPG only; photo appears fitted, not stretched | |
| 6 | Load Photo again | Photo does not shrink further than the first load | |
| 7 | Editor: Load Photo with Name blank | "Enter Name and Part Number before loading a photo." | |
| 8 | Place Pins on, click the photo four times with Pin Count 4 | Four small numbered badges appear where clicked | |
| 9 | Click a fifth time | "All 4 pins have been placed." | |
| 10 | Clear Pin Count, click the photo | "Enter a valid Pin Count before placing pins." | |
| 11 | Drag a badge away from its pin | Badge moves and stays where dropped | |
| 12 | Select a pin, Place Pins off, click elsewhere | Anchor moves; badge follows only if it was still on the anchor | |
| 13 | Select the middle pin, Delete Pin | That badge and list row disappear; remaining numbers unchanged | |
| 14 | Place another pin after that deletion | New pin takes the next unused number, not the deleted one | |
| 15 | Snap Label with a pin selected | Badge jumps back onto its pin | |
| 16 | Clear Pins | All badges and list rows disappear | |
| 17 | Save with all fields and a photo | Editor closes; connector appears in Manage Library | |
| 18 | Save a second connector using an existing Part Number | "Part Number already exists in the library (...)" | |
| 19 | **Save without ever loading a photo** | "Could not save ... Load a photo before saving." (was silent before) | |
| 20 | Manage Library > Edit an existing connector | Fields populate; photo preview appears; pins show as badges | |
| 21 | Manage Library > Delete | Confirmation, then the row disappears and the list refreshes | |
| 22 | Manage Library > Export | Save dialog, then "Exported <ID>." | |
| 23 | Manage Library > Import that file into another library | "Import complete. 1 connector(s) imported." | |
| 24 | Import a file whose photo cannot be extracted | Replacement-photo prompt appears for that connector only | |
| 25 | Rename a ref des on Connectors to an unused value | Rename sticks; chart references follow | |
| 26 | Rename a ref des onto an existing one | Cell reverts to the previous value | |
| 27 | Paste a large block over the chart | Pin dropdowns rebuild; pasted values are not cleared | |
| 28 | Change the units field in the title block | Length units update on the chart | |
```

- [ ] **Step 3: Record the outcomes and commit**

```bash
git add docs/superpowers/plans/2026-08-28-ui-separation-manual-verification.md
git commit -m "docs: record the UI separation manual verification pass"
```

---

## Notes for the executor

- **A VBA compile error fails every test at once**, because the `artifact` fixture rebuilds the workbook for the whole session. If the entire suite goes red after a form rewrite, open `dist/HarnessCreator.xlsm`, press Alt+F11, and use Debug > Compile VBAProject to find the line.
- **Excel processes can leak** if a test run is interrupted. If builds start failing oddly, check Task Manager for orphaned `EXCEL.EXE` processes and end them.
- **`run` vs `run_action`:** actions return the envelope and use `run_action`; queries return bare values and use `run`. Getting this wrong produces a confusing `AssertionError` about a three-element result rather than a real failure.
- **Never edit `dist/`.** Every change goes in `src/vba/` or `build/`, then rebuild.
