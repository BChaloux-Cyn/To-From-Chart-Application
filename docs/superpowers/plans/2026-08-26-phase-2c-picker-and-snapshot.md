# Phase 2c: Connector Picker, Snapshot Embedding, and Ref Des Rename Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Wire the library (2a) and the connector editor (2b) into the Creator's actual workflow: Add Connector picks a library part and freezes its definition into `_Snapshot`; Manage Library browses, edits, and deletes library entries; Remove Connector drops an instance from the current harness; and ref des rename - required by the spec but never built in Phase 1 - rewrites every chart reference and rejects collisions.

**Architecture:** `_Snapshot` becomes real in this plan: one very-hidden sheet, three fixed row/column regions mirroring 2a's Connectors/Pins schema plus a photo-shapes area, populated by `modLibrary`'s already-tested bounded-window functions - no new storage abstraction, just new bounds. Two new UserForms (`frmConnectorPicker`, `frmManageLibrary`) are UI wiring only, following 2b's discipline: every consequential action is a call into a plain, `Application.Run`-testable module (`modSnapshot`, extensions to `modConnectors`), never logic embedded in a form event handler.

**Tech Stack:** Python 3.13, pywin32, pytest, Excel 16.0 COM automation, VBA.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Depends on:** 2a (`modLibrary`), 2b (`frmConnectorEditor`, `modPinEditor`), and Phase 1 (`modConnectors`, `modChart`, `modState`).

**Part of Phase 2** (2a, 2b done; this is 2c). 2d (library import/export) still to come, followed by 2e (design docs + student user guide).

## Global Constraints

Copied verbatim from the spec, plus 2a/2b's. Every task's requirements implicitly include these.

- Every VBA module starts with `Option Explicit`. No `MsgBox`/dialog in a logic module.
- The `.xlsm` is a build artifact, never hand-edited.
- A harness's connector snapshot is frozen at the moment the connector is added; later library edits never propagate into an existing harness.
- Ref designators must be unique within a harness. Renaming one rewrites every reference to it in the chart; a rename that collides is rejected.
- ConnectorID names a library part; Ref Des names one physical instance in one harness. A harness can use the same library part more than once, each with its own ref des resolving to one library definition.
- Custom VBA `Type`s never cross the `Application.Run` boundary (2a's convention, carried forward).

## File Structure

| File | Responsibility |
|---|---|
| `build/layout.py` | Modified: lays out `_Snapshot`'s three regions |
| `src/vba/modSnapshot.bas` | Copies a library connector's full definition into `_Snapshot`, idempotently |
| `src/vba/modConnectors.bas` | Modified: adds `RenameRefDes`, `RemoveConnectorInstance` |
| `src/vba/sheets/shConnectors.evt` | Modified: detects a ref des rename via cached prior selection, reverts on collision |
| `build/form_layout.py` | Modified: adds `frmConnectorPicker` and `frmManageLibrary` |
| `src/vba/forms/frmConnectorPicker.evt` | Add Connector: pick existing or launch the editor for a new one |
| `src/vba/forms/frmManageLibrary.evt` | Browse, edit, delete library connectors; Import/Export buttons present, unwired until 2d |
| `build/layout.py` | Modified: Home gets Add Connector / Remove Connector / Manage Library buttons |
| `tests/test_snapshot.py` | `_Snapshot` layout and `SnapshotConnector` |
| `tests/test_ref_des_rename.py` | `RenameRefDes` |
| `tests/test_remove_connector.py` | `RemoveConnectorInstance` |
| `tests/test_picker_form.py` | Structural: picker/browser controls and wiring |

---

### Task 1: `_Snapshot` layout

**Files:**
- Modify: `build/layout.py`
- Test: `tests/test_snapshot.py`

**Interfaces:**
- Consumes: `modLibrary`'s schema constants (referenced by row window only, not redefined).
- Produces:
  - `layout.SNAP_CONN_FIRST_ROW = 2`, `layout.SNAP_CONN_LAST_ROW = 201`
  - `layout.SNAP_PINS_HEADER_ROW = 210`, `layout.SNAP_PINS_FIRST_ROW = 211`, `layout.SNAP_PINS_LAST_ROW = 2210`
  - `layout.build_snapshot(sheets) -> None`

200 distinct connector definitions and 2000 pin rows is generous headroom for what one harness actually uses - a handful of connectors, not hundreds - while staying a small, fixed cost on a very-hidden sheet nobody scrolls through by hand. The photo-shapes area shares the same sheet at arbitrary pixel positions (via `modLibrary.EmbedConnectorPhoto`'s existing grid logic) and cosmetically overlaps the data rows when viewed - harmless, since the sheet is very hidden and shapes and cells are independent object models.

- [ ] **Step 1: Write the failing test**

Create `tests/test_snapshot.py`:

```python
def test_snapshot_connectors_header_matches_the_library_schema(wb):
    sheet = wb.Worksheets("_Snapshot")
    conn_headers = [
        "ConnectorID", "Name", "Manufacturer", "PartNumber", "Type",
        "PinCount", "Notes", "PhotoShapeName", "CreatedUtc", "ModifiedUtc", "Origin",
    ]
    for index, header in enumerate(conn_headers, start=1):
        assert sheet.Cells(1, index).Value == header


def test_snapshot_pins_header_matches_the_library_schema(wb):
    sheet = wb.Worksheets("_Snapshot")
    pin_headers = ["ConnectorID", "PinNumber", "PinLabel", "NormX", "NormY", "LabelX", "LabelY"]
    for index, header in enumerate(pin_headers, start=1):
        assert sheet.Cells(210, index).Value == header
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_snapshot.py -v
```

Expected: FAIL — `_Snapshot` is empty.

- [ ] **Step 3: Add the layout**

Append to `build/layout.py`:

```python
SNAP_CONN_FIRST_ROW = 2
SNAP_CONN_LAST_ROW = 201
SNAP_PINS_HEADER_ROW = 210
SNAP_PINS_FIRST_ROW = 211
SNAP_PINS_LAST_ROW = 2210


def build_snapshot(sheets) -> None:
    sheet = sheets["_Snapshot"]
    for index, header in enumerate(library_layout.CONN_HEADERS, start=1):
        sheet.Cells(1, index).Value = header
    for index, header in enumerate(library_layout.PIN_HEADERS, start=1):
        sheet.Cells(SNAP_PINS_HEADER_ROW, index).Value = header
```

Add `import library_layout` alongside `layout.py`'s other imports.

- [ ] **Step 4: Wire into the build**

In `build/build.py`, after `layout.build_home(sheets)`:

```python
            layout.build_snapshot(sheets)
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_snapshot.py -v
```

Expected: 2 passed.

- [ ] **Step 6: Commit**

```bash
git add build/layout.py build/build.py tests/test_snapshot.py
git commit -m "feat: lay out the _Snapshot sheet's connector and pin regions"
```

---

### Task 2: Snapshot embedding

**Files:**
- Create: `src/vba/modSnapshot.bas`
- Modify: `build/build.py`
- Test: `tests/test_snapshot.py` (append)

**Interfaces:**
- Consumes: `modLibrary.FindConnectorRow`, `ReadConnector`, `WriteConnector`, `ReadPinsForConnector`, `WritePin`, `EmbedConnectorPhoto`, `CachePhotoPath`, `LIB_ROW_CAP`.
- Produces:
  - VBA `modSnapshot.SnapshotConnector(wsSnap, wsLibConn, wsLibPins, wsLibPhotos, sConnectorID) As Boolean` - idempotent
  - VBA `modSnapshot.LibraryFolder() As String`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_snapshot.py`:

```python
from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def seed_library_connector(wb, library_wb, tmp_path, connector_id="DTM-04P", pins=((1, "A"), (2, "B"))):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws_conn = library_wb.Worksheets("Connectors")
    ws_pins = library_wb.Worksheets("Pins")
    ws_photos = library_wb.Worksheets("Photos")

    fields = (connector_id, "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
              len(pins), "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields)
    for pin_number, label in pins:
        run(wb, "modLibrary.WritePin", ws_pins, 2, 100000, (connector_id, pin_number, label, 0.1, 0.1, 0.1, 0.1))
    shape_name = run(wb, "modLibrary.EmbedConnectorPhoto", ws_photos, connector_id, str(photo_path))
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields[:7] + (shape_name,) + fields[8:])
    return ws_conn, ws_pins, ws_photos


def test_snapshot_connector_copies_the_full_definition(wb, library_wb, tmp_path):
    ws_conn, ws_pins, ws_photos = seed_library_connector(wb, library_wb, tmp_path)
    wsnap = wb.Worksheets("_Snapshot")

    ok = run(wb, "modSnapshot.SnapshotConnector", wsnap, ws_conn, ws_pins, ws_photos, "DTM-04P")
    assert ok is True

    result = run(wb, "modLibrary.ReadConnector", wsnap, 2, 201, "DTM-04P")
    assert result[1] == "Deutsch DTM 4-way"

    pins = run(wb, "modLibrary.ReadPinsForConnector", wsnap, 211, 2210, "DTM-04P")
    assert [row[1] for row in pins] == [1, 2]

    assert wsnap.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"


def test_snapshot_connector_is_idempotent(wb, library_wb, tmp_path):
    ws_conn, ws_pins, ws_photos = seed_library_connector(wb, library_wb, tmp_path)
    wsnap = wb.Worksheets("_Snapshot")

    run(wb, "modSnapshot.SnapshotConnector", wsnap, ws_conn, ws_pins, ws_photos, "DTM-04P")
    run(wb, "modSnapshot.SnapshotConnector", wsnap, ws_conn, ws_pins, ws_photos, "DTM-04P")

    assert wsnap.Cells(3, 1).Value is None  # not duplicated on a second call


def test_snapshot_connector_for_unknown_id_returns_false(wb, library_wb):
    wsnap = wb.Worksheets("_Snapshot")
    ws_conn = library_wb.Worksheets("Connectors")
    ws_pins = library_wb.Worksheets("Pins")
    ws_photos = library_wb.Worksheets("Photos")

    assert run(wb, "modSnapshot.SnapshotConnector", wsnap, ws_conn, ws_pins, ws_photos, "NO-SUCH-ID") is False
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_snapshot.py -v
```

Expected: FAIL — `modSnapshot` does not exist.

- [ ] **Step 3: Write the module**

Create `src/vba/modSnapshot.bas`:

```vb
Attribute VB_Name = "modSnapshot"
Option Explicit

Public Const SNAP_CONN_FIRST_ROW As Long = 2
Public Const SNAP_CONN_LAST_ROW As Long = 201
Public Const SNAP_PINS_FIRST_ROW As Long = 211
Public Const SNAP_PINS_LAST_ROW As Long = 2210

Public Function LibraryFolder() As String
    LibraryFolder = ThisWorkbook.Path
End Function

Public Function SnapshotConnector(wsSnap As Worksheet, wsLibConn As Worksheet, wsLibPins As Worksheet, _
                                  wsLibPhotos As Worksheet, ByVal sConnectorID As String) As Boolean
    ' Frozen once per distinct ConnectorID - a second instance of the same
    ' library part (J1 and J2 from one DTM-04P) never duplicates the
    ' definition it shares.
    If modLibrary.FindConnectorRow(wsSnap, SNAP_CONN_FIRST_ROW, SNAP_CONN_LAST_ROW, sConnectorID) > 0 Then
        SnapshotConnector = True
        Exit Function
    End If

    Dim vFields As Variant
    vFields = modLibrary.ReadConnector(wsLibConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vFields) Then Exit Function
    If Not modLibrary.WriteConnector(wsSnap, SNAP_CONN_FIRST_ROW, SNAP_CONN_LAST_ROW, vFields) Then Exit Function

    Dim vPins As Variant, i As Long, j As Long, vRow(1 To 7) As Variant
    vPins = modLibrary.ReadPinsForConnector(wsLibPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If Not IsEmpty(vPins) Then
        For i = LBound(vPins, 1) To UBound(vPins, 1)
            For j = 1 To modLibrary.PIN_FIELD_COUNT
                vRow(j) = vPins(i, j)
            Next j
            modLibrary.WritePin wsSnap, SNAP_PINS_FIRST_ROW, SNAP_PINS_LAST_ROW, vRow
        Next i
    End If

    Dim sCachePath As String
    sCachePath = modLibrary.CachePhotoPath(LibraryFolder(), sConnectorID)
    If Len(Dir$(sCachePath)) > 0 Then
        modLibrary.EmbedConnectorPhoto wsSnap, sConnectorID, sCachePath
    Else
        ' No local cache file yet (e.g. the sample fixture in this plan's
        ' own tests never wrote one) - copy the shape straight off the
        ' library's own Photos sheet instead of failing the snapshot.
        modLibrary.RemoveConnectorPhoto wsSnap, sConnectorID
        On Error Resume Next
        wsLibPhotos.Shapes("PHOTO_" & sConnectorID).Copy
        wsSnap.Paste
        wsSnap.Shapes(wsSnap.Shapes.Count).Name = "PHOTO_" & sConnectorID
        On Error GoTo 0
    End If

    SnapshotConnector = True
End Function
```

- [ ] **Step 4: Wire the module into the build**

In `build/build.py`:

```python
VBA_MODULES = [
    "modUtil.bas", "modState.bas", "modConnectors.bas", "modChart.bas",
    "modLibrary.bas", "modPinEditor.bas", "clsPinMarker.cls", "modSnapshot.bas",
]
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_snapshot.py -v
```

Expected: 5 passed.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modSnapshot.bas build/build.py tests/test_snapshot.py
git commit -m "feat: freeze a library connector's definition into _Snapshot"
```

---

### Task 3: Ref des rename

**Files:**
- Modify: `src/vba/modConnectors.bas`
- Modify: `src/vba/sheets/shConnectors.evt`
- Test: `tests/test_ref_des_rename.py`

**Interfaces:**
- Consumes: `modChart.CHART_SHEET`, `CHART_FIRST_ROW`, `CHART_LAST_ROW`, `COL_FROM_CONN`, `COL_TO_CONN`.
- Produces: VBA `modConnectors.RenameRefDes(sOldRefDes, sNewRefDes) As Boolean`

`Worksheet_Change` only ever sees a cell's *new* value - there is no built-in "previous value" for a plain cell edit. The standard, and only reliable, way to know what a ref des changed *from* is to cache it on `Worksheet_SelectionChange`, immediately before the cell is edited, then compare against that cache when `Worksheet_Change` fires.

- [ ] **Step 1: Write the failing test**

Create `tests/test_ref_des_rename.py`:

```python
from tests.conftest import run


def add(wb, connector_id="DTM-04P"):
    return run(wb, "modConnectors.AddConnectorInstance",
               connector_id, "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)


def test_rename_rewrites_every_chart_reference(wb):
    add(wb)  # J1
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(8, 9).Value = "J1"

    conn_sheet = wb.Worksheets("Connectors")
    conn_sheet.Cells(2, 1).Value = "J99"  # simulates the cell already having been edited

    ok = run(wb, "modConnectors.RenameRefDes", "J1", "J99")
    assert ok is True
    assert sheet.Cells(7, 1).Value == "J99"
    assert sheet.Cells(8, 9).Value == "J99"


def test_rename_leaves_unrelated_rows_alone(wb):
    add(wb)  # J1
    add(wb)  # J2
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(8, 1).Value = "J2"

    conn_sheet = wb.Worksheets("Connectors")
    conn_sheet.Cells(2, 1).Value = "J99"

    run(wb, "modConnectors.RenameRefDes", "J1", "J99")
    assert sheet.Cells(8, 1).Value == "J2"


def test_rename_colliding_with_an_existing_ref_des_is_rejected(wb):
    add(wb)  # J1
    add(wb)  # J2

    conn_sheet = wb.Worksheets("Connectors")
    conn_sheet.Cells(2, 1).Value = "J2"  # simulates renaming J1 to the already-used J2

    ok = run(wb, "modConnectors.RenameRefDes", "J1", "J2")
    assert ok is False


def test_rename_to_the_same_value_is_a_no_op(wb):
    add(wb)  # J1
    assert run(wb, "modConnectors.RenameRefDes", "J1", "J1") is False


def test_renaming_via_the_sheet_reverts_the_cell_on_collision(wb):
    add(wb)  # J1
    add(wb)  # J2
    conn_sheet = wb.Worksheets("Connectors")

    conn_sheet.Cells(2, 1).Select()  # caches "J1" as the prior value - () matters: a bare `.Select`
                                      # is a Python attribute access on the COM method, never calls it
    conn_sheet.Cells(2, 1).Value = "J2"  # collides - shConnectors.evt must revert this

    assert conn_sheet.Cells(2, 1).Value == "J1"
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_ref_des_rename.py -v
```

Expected: `test_renaming_via_the_sheet_reverts_the_cell_on_collision` FAILs (nothing reverts it yet); the four direct `RenameRefDes` calls FAIL with "function not defined."

- [ ] **Step 3: Add `RenameRefDes`**

Append to `src/vba/modConnectors.bas`:

```vb
Public Function RenameRefDes(ByVal sOldRefDes As String, ByVal sNewRefDes As String) As Boolean
    Dim ws As Worksheet, wsChart As Worksheet
    Dim r As Long, nLast As Long, nMatches As Long

    sOldRefDes = Trim$(sOldRefDes)
    sNewRefDes = Trim$(sNewRefDes)
    If Len(sOldRefDes) = 0 Or Len(sNewRefDes) = 0 Then Exit Function
    If StrComp(sOldRefDes, sNewRefDes, vbTextCompare) = 0 Then Exit Function

    ' The renamed row already carries sNewRefDes by the time this runs (the
    ' sheet edit happens before Worksheet_Change fires), so exactly one
    ' match is the non-colliding case; more than one means a different row
    ' already used that ref des.
    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = CONN_FIRST_ROW To nLast
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), sNewRefDes, vbTextCompare) = 0 Then
            nMatches = nMatches + 1
        End If
    Next r
    If nMatches <> 1 Then Exit Function

    Set wsChart = ThisWorkbook.Worksheets(modChart.CHART_SHEET)
    For r = modChart.CHART_FIRST_ROW To modChart.CHART_LAST_ROW
        If StrComp(Trim$(CStr(wsChart.Cells(r, modChart.COL_FROM_CONN).Value)), sOldRefDes, vbTextCompare) = 0 Then
            wsChart.Cells(r, modChart.COL_FROM_CONN).Value = sNewRefDes
        End If
        If StrComp(Trim$(CStr(wsChart.Cells(r, modChart.COL_TO_CONN).Value)), sOldRefDes, vbTextCompare) = 0 Then
            wsChart.Cells(r, modChart.COL_TO_CONN).Value = sNewRefDes
        End If
    Next r

    RenameRefDes = True
End Function
```

- [ ] **Step 4: Wire the cache-and-revert into the Connectors sheet handler**

In `src/vba/sheets/shConnectors.evt`, replace the whole file:

```vb
Private mLastRefDes As String
Private mLastRefDesRow As Long

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count = 1 And Target.Column = 1 And Target.Row >= modConnectors.CONN_FIRST_ROW Then
        mLastRefDesRow = Target.Row
        mLastRefDes = Trim$(CStr(Target.Value))
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim bEvents As Boolean
    Dim rw As Range
    Dim sRef As String

    If Not Application.EnableEvents Then Exit Sub

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    modState.MarkDirty

    If Target.Cells.Count = 1 And Target.Column = 1 And Target.Row = mLastRefDesRow _
       And Len(mLastRefDes) > 0 Then
        Dim sNewRefDes As String
        sNewRefDes = Trim$(CStr(Target.Value))
        If StrComp(sNewRefDes, mLastRefDes, vbTextCompare) <> 0 Then
            If Not modConnectors.RenameRefDes(mLastRefDes, sNewRefDes) Then
                Target.Value = mLastRefDes
            Else
                mLastRefDes = sNewRefDes
            End If
        End If
    End If

    ' A pin-count (or other) edit to an existing connector row must refresh
    ' any chart rows already referencing it, so their pin dropdowns stay
    ' in sync with the connector's current pin count.
    For Each rw In Target.Rows
        If rw.Row >= modConnectors.CONN_FIRST_ROW Then
            sRef = Trim$(CStr(Me.Cells(rw.Row, 1).Value))
            modChart.RefreshChartRowsForConnector sRef
        End If
    Next rw

CleanUp:
    Application.EnableEvents = bEvents
End Sub
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_ref_des_rename.py -v
```

Expected: 5 passed.

- [ ] **Step 6: Run the whole suite**

Run:

```bash
python -m pytest -v
```

Expected: all passed (208 from 2b + 2 from Task 1 + 5 from Task 2 + 5 from Task 3 = 220).

- [ ] **Step 7: Commit**

```bash
git add src/vba/modConnectors.bas src/vba/sheets/shConnectors.evt tests/test_ref_des_rename.py
git commit -m "feat: rewrite chart references on ref des rename, reject collisions"
```

---

### Task 4: Remove Connector

**Files:**
- Modify: `src/vba/modConnectors.bas`
- Test: `tests/test_remove_connector.py`

**Interfaces:**
- Consumes: `modChart.CHART_FIRST_ROW`, `CHART_LAST_ROW`, `COL_FROM_CONN`, `COL_FROM_PIN`, `COL_TO_CONN`, `COL_TO_PIN`, `RebuildPinValidation`.
- Produces: VBA `modConnectors.RemoveConnectorInstance(sRefDes) As Boolean`

Removing an instance clears only the endpoint cells that referenced it (From Conn/From Pin, or To Conn/To Pin), never the whole wire row - the rest of that row's data (Signal, Color, AWG, Length, Notes) is not the student's mistake to lose because a connector went away.

- [ ] **Step 1: Write the failing test**

Create `tests/test_remove_connector.py`:

```python
from tests.conftest import run


def add(wb, connector_id="DTM-04P"):
    return run(wb, "modConnectors.AddConnectorInstance",
               connector_id, "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)


def test_remove_deletes_the_connector_row(wb):
    add(wb)  # J1
    add(wb)  # J2

    ok = run(wb, "modConnectors.RemoveConnectorInstance", "J1")
    assert ok is True

    conn_sheet = wb.Worksheets("Connectors")
    assert run(wb, "modConnectors.PinCountFor", "J1") == 0
    assert run(wb, "modConnectors.PinCountFor", "J2") == 4  # untouched


def test_remove_clears_only_the_referencing_endpoint(wb):
    add(wb)  # J1
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(7, 2).Value = 1
    sheet.Cells(7, 4).Value = "+12V Batt"  # Signal - must survive

    run(wb, "modConnectors.RemoveConnectorInstance", "J1")

    assert sheet.Cells(7, 1).Value is None
    assert sheet.Cells(7, 2).Value is None
    assert sheet.Cells(7, 4).Value == "+12V Batt"


def test_remove_clears_a_to_endpoint_without_touching_the_from_endpoint(wb):
    add(wb)  # J1
    add(wb)  # J2
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J2"
    sheet.Cells(7, 9).Value = "J1"

    run(wb, "modConnectors.RemoveConnectorInstance", "J1")

    assert sheet.Cells(7, 1).Value == "J2"
    assert sheet.Cells(7, 9).Value is None


def test_remove_unknown_ref_des_returns_false(wb):
    assert run(wb, "modConnectors.RemoveConnectorInstance", "J99") is False
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_remove_connector.py -v
```

Expected: FAIL — `RemoveConnectorInstance` does not exist.

- [ ] **Step 3: Add the function**

Append to `src/vba/modConnectors.bas`:

```vb
Public Function RemoveConnectorInstance(ByVal sRefDes As String) As Boolean
    Dim ws As Worksheet, wsChart As Worksheet
    Dim r As Long, nLast As Long, c As Long

    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    r = 0
    Dim i As Long
    For i = CONN_FIRST_ROW To nLast
        If StrComp(Trim$(CStr(ws.Cells(i, 1).Value)), sRefDes, vbTextCompare) = 0 Then
            r = i
            Exit For
        End If
    Next i
    If r = 0 Then Exit Function

    If r < nLast Then
        For c = 1 To 6
            ws.Cells(r, c).Value = ws.Cells(nLast, c).Value
        Next c
    End If
    ws.Range(ws.Cells(nLast, 1), ws.Cells(nLast, 6)).ClearContents

    Set wsChart = ThisWorkbook.Worksheets(modChart.CHART_SHEET)
    For i = modChart.CHART_FIRST_ROW To modChart.CHART_LAST_ROW
        If StrComp(Trim$(CStr(wsChart.Cells(i, modChart.COL_FROM_CONN).Value)), sRefDes, vbTextCompare) = 0 Then
            wsChart.Cells(i, modChart.COL_FROM_CONN).ClearContents
            wsChart.Cells(i, modChart.COL_FROM_PIN).Validation.Delete
            wsChart.Cells(i, modChart.COL_FROM_PIN).ClearContents
        End If
        If StrComp(Trim$(CStr(wsChart.Cells(i, modChart.COL_TO_CONN).Value)), sRefDes, vbTextCompare) = 0 Then
            wsChart.Cells(i, modChart.COL_TO_CONN).ClearContents
            wsChart.Cells(i, modChart.COL_TO_PIN).Validation.Delete
            wsChart.Cells(i, modChart.COL_TO_PIN).ClearContents
        End If
    Next i

    RemoveConnectorInstance = True
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_remove_connector.py -v
```

Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modConnectors.bas tests/test_remove_connector.py
git commit -m "feat: remove a connector instance and clear its chart references"
```

---

### Task 5: Connector picker and library browser forms

**Files:**
- Modify: `build/form_layout.py`
- Modify: `build/build.py`
- Modify: `build/layout.py` (Home buttons)
- Create: `src/vba/forms/frmConnectorPicker.evt`
- Create: `src/vba/forms/frmManageLibrary.evt`
- Test: `tests/test_picker_form.py`

**Interfaces:**
- Consumes: `modLibrary`, `modConnectors.AddConnectorInstance`, `modSnapshot.SnapshotConnector`, `frmConnectorEditor`.
- Produces: `frmConnectorPicker`, `frmManageLibrary` UserForms; Home buttons "Add Connector", "Remove Connector", "Manage Library".

Both forms follow 2b's rule: every button click that changes state calls a single already-tested function. `frmManageLibrary` gets Import/Export buttons now, for layout stability, but no `cmdImport_Click`/`cmdExport_Click` handlers - an unwired button click does nothing, which is correct until 2d implements them. This is the same reasoning 2b used to leave a manual-only verification step for mouse-driven UI: pytest cannot drive either form's actual dialog interaction, so tests here are structural, matching 2b Task 8's pattern.

- [ ] **Step 1: Write the failing test**

Create `tests/test_picker_form.py`:

```python
import pytest


def module_source(wb, component_name):
    module = wb.VBProject.VBComponents(component_name).CodeModule
    return module.Lines(1, module.CountOfLines)


def controls(wb, form_name):
    return wb.VBProject.VBComponents(form_name).Designer.Controls


@pytest.mark.parametrize("name", ["lstConnectors", "cmdAdd", "cmdNew", "cmdCancel"])
def test_picker_has_its_controls(wb, name):
    assert controls(wb, "frmConnectorPicker")(name).Name == name


@pytest.mark.parametrize("name", ["lstConnectors", "cmdEdit", "cmdDelete", "cmdImport", "cmdExport", "cmdClose"])
def test_manage_library_has_its_controls(wb, name):
    assert controls(wb, "frmManageLibrary")(name).Name == name


def test_picker_add_calls_add_connector_instance_and_snapshot(wb):
    source = module_source(wb, "frmConnectorPicker")
    assert "modConnectors.AddConnectorInstance" in source
    assert "modSnapshot.SnapshotConnector" in source


def test_picker_new_launches_the_connector_editor(wb):
    assert "frmConnectorEditor" in module_source(wb, "frmConnectorPicker")


def test_manage_library_edit_launches_the_connector_editor(wb):
    source = module_source(wb, "frmManageLibrary")
    assert "frmConnectorEditor" in source
    assert "LoadScratchPins" in source


def test_manage_library_delete_calls_library_delete_functions(wb):
    source = module_source(wb, "frmManageLibrary")
    assert "modLibrary.DeleteConnector" in source
    assert "modLibrary.DeletePinsForConnector" in source


def test_manage_library_import_export_are_unwired_for_now(wb):
    source = module_source(wb, "frmManageLibrary")
    assert "cmdImport_Click" not in source
    assert "cmdExport_Click" not in source


def test_home_has_the_three_new_buttons(wb):
    shapes = wb.Worksheets("Home").Shapes
    actions = [shapes(i + 1).OnAction for i in range(shapes.Count)]
    assert "modConnectorUI.ShowAddConnector" in actions
    assert "modConnectorUI.ShowManageLibrary" in actions
    assert "modConnectorUI.ShowRemoveConnector" in actions
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_picker_form.py -v
```

Expected: FAIL — neither form nor the Home buttons exist.

- [ ] **Step 3: Add the form layouts**

Append to `build/form_layout.py`:

```python
PICKER_NAME = "frmConnectorPicker"
PICKER_CONTROLS = [
    ("Forms.ListBox.1", "lstConnectors", 12, 12, 300, 200, {}),
    ("Forms.CommandButton.1", "cmdAdd", 12, 220, 90, 24, {"Caption": "Add"}),
    ("Forms.CommandButton.1", "cmdNew", 110, 220, 90, 24, {"Caption": "New..."}),
    ("Forms.CommandButton.1", "cmdCancel", 222, 220, 90, 24, {"Caption": "Cancel"}),
]

MANAGE_LIBRARY_NAME = "frmManageLibrary"
MANAGE_LIBRARY_CONTROLS = [
    ("Forms.ListBox.1", "lstConnectors", 12, 12, 300, 200, {}),
    ("Forms.CommandButton.1", "cmdEdit", 12, 220, 80, 24, {"Caption": "Edit"}),
    ("Forms.CommandButton.1", "cmdDelete", 96, 220, 80, 24, {"Caption": "Delete"}),
    ("Forms.CommandButton.1", "cmdImport", 180, 220, 80, 24, {"Caption": "Import..."}),
    ("Forms.CommandButton.1", "cmdExport", 264, 220, 80, 24, {"Caption": "Export..."}),
    ("Forms.CommandButton.1", "cmdClose", 348, 220, 80, 24, {"Caption": "Close"}),
]


def _build_form(wb, add_userform, name, caption, width, height, control_specs):
    designer = add_userform(wb, name)
    designer.Caption = caption
    designer.Width = width
    designer.Height = height
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
    _build_form(wb, add_userform, PICKER_NAME, "Add Connector", 340, 260, PICKER_CONTROLS)


def build_manage_library_form(wb, add_userform) -> None:
    _build_form(wb, add_userform, MANAGE_LIBRARY_NAME, "Manage Library", 340, 260, MANAGE_LIBRARY_CONTROLS)
```

- [ ] **Step 4: Write the picker's code-behind**

Create `src/vba/forms/frmConnectorPicker.evt`:

```vb
Option Explicit

Private mLibrary As Workbook
Private mConnectorIDs() As String

Private Sub UserForm_Initialize()
    Set mLibrary = Workbooks.Open(ThisWorkbook.Path & "\ConnectorLibrary.xlsx")
    RefreshList
End Sub

Private Sub RefreshList()
    Dim ws As Worksheet, nLast As Long, r As Long, n As Long

    lstConnectors.Clear
    Set ws = mLibrary.Worksheets("Connectors")
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If nLast < 2 Then Exit Sub

    ReDim mConnectorIDs(1 To nLast - 1)
    For r = 2 To nLast
        n = n + 1
        mConnectorIDs(n) = Trim$(CStr(ws.Cells(r, modLibrary.LIB_COL_ID).Value))
        lstConnectors.AddItem mConnectorIDs(n) & " - " & CStr(ws.Cells(r, modLibrary.LIB_COL_NAME).Value)
    Next r
End Sub

Private Sub cmdAdd_Click()
    If lstConnectors.ListIndex < 0 Then Exit Sub
    Dim sConnectorID As String
    sConnectorID = mConnectorIDs(lstConnectors.ListIndex + 1)

    Dim vFields As Variant
    vFields = modLibrary.ReadConnector(mLibrary.Worksheets("Connectors"), 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vFields) Then Exit Sub

    Dim sRefDes As String
    sRefDes = modConnectors.AddConnectorInstance(vFields(0), vFields(1), vFields(3), vFields(4), CLng(vFields(5)))
    If Len(sRefDes) = 0 Then Exit Sub

    modSnapshot.SnapshotConnector ThisWorkbook.Worksheets("_Snapshot"), mLibrary.Worksheets("Connectors"), _
        mLibrary.Worksheets("Pins"), mLibrary.Worksheets("Photos"), sConnectorID

    mLibrary.Close SaveChanges:=False
    Unload Me
End Sub

Private Sub cmdNew_Click()
    mLibrary.Close SaveChanges:=False
    Unload Me
    frmConnectorEditor.Show
End Sub

Private Sub cmdCancel_Click()
    mLibrary.Close SaveChanges:=False
    Unload Me
End Sub
```

- [ ] **Step 5: Write the library browser's code-behind**

Create `src/vba/forms/frmManageLibrary.evt`:

```vb
Option Explicit

Private mLibrary As Workbook
Private mConnectorIDs() As String

Private Sub UserForm_Initialize()
    Set mLibrary = Workbooks.Open(ThisWorkbook.Path & "\ConnectorLibrary.xlsx")
    RefreshList
End Sub

Private Sub RefreshList()
    Dim ws As Worksheet, nLast As Long, r As Long, n As Long

    lstConnectors.Clear
    Set ws = mLibrary.Worksheets("Connectors")
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If nLast < 2 Then Exit Sub

    ReDim mConnectorIDs(1 To nLast - 1)
    For r = 2 To nLast
        n = n + 1
        mConnectorIDs(n) = Trim$(CStr(ws.Cells(r, modLibrary.LIB_COL_ID).Value))
        lstConnectors.AddItem mConnectorIDs(n) & " - " & CStr(ws.Cells(r, modLibrary.LIB_COL_NAME).Value)
    Next r
End Sub

Private Sub cmdEdit_Click()
    If lstConnectors.ListIndex < 0 Then Exit Sub
    Dim sConnectorID As String
    sConnectorID = mConnectorIDs(lstConnectors.ListIndex + 1)

    modPinEditor.LoadScratchPins ThisWorkbook.Worksheets("_Edit"), mLibrary.Worksheets("Pins"), sConnectorID
    mLibrary.Close SaveChanges:=False
    Unload Me
    frmConnectorEditor.Show
End Sub

Private Sub cmdDelete_Click()
    If lstConnectors.ListIndex < 0 Then Exit Sub
    Dim sConnectorID As String
    sConnectorID = mConnectorIDs(lstConnectors.ListIndex + 1)

    If MsgBox("Delete " & sConnectorID & " from the library? This cannot be undone.", _
              vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    modLibrary.DeleteConnector mLibrary.Worksheets("Connectors"), 2, modLibrary.LIB_ROW_CAP, sConnectorID
    modLibrary.DeletePinsForConnector mLibrary.Worksheets("Pins"), 2, modLibrary.LIB_ROW_CAP, sConnectorID
    modLibrary.RemoveConnectorPhoto mLibrary.Worksheets("Photos"), sConnectorID
    mLibrary.Save

    RefreshList
End Sub

Private Sub cmdClose_Click()
    mLibrary.Close SaveChanges:=False
    Unload Me
End Sub
```

`MsgBox` on `cmdDelete_Click` is a confirmation dialog inside the UI module (the form's own code-behind), not a logic module - it does not violate "no `MsgBox` in a logic module." A `TestMode`-suppressed path is not added here because no automated test drives this click at all (2b Task 8's precedent); it is exercised only in the manual verification step below.

- [ ] **Step 6: Wire both forms and the Home buttons into the build**

In `build/build.py`, extend `FORM_EVENTS`:

```python
FORM_EVENTS = [
    ("frmConnectorEditor", "frmConnectorEditor.evt"),
    ("frmConnectorPicker", "frmConnectorPicker.evt"),
    ("frmManageLibrary", "frmManageLibrary.evt"),
]
```

and after `form_layout.build_connector_editor_form(...)`:

```python
            form_layout.build_connector_picker_form(wb, excel_com.add_userform)
            form_layout.build_manage_library_form(wb, excel_com.add_userform)
```

Add `"modConnectorUI.bas"` to `VBA_MODULES`. Create `src/vba/modConnectorUI.bas` - the thin `Show`-calling glue the Home buttons' `OnAction` points at, kept separate from the forms themselves so `OnAction` (a plain module-qualified string) never has to name a form directly:

```vb
Attribute VB_Name = "modConnectorUI"
Option Explicit

Public Sub ShowAddConnector()
    frmConnectorPicker.Show
End Sub

Public Sub ShowManageLibrary()
    frmManageLibrary.Show
End Sub

Public Sub ShowRemoveConnector()
    Dim sRefDes As String
    sRefDes = InputBox("Ref des to remove:", "Remove Connector")
    If Len(Trim$(sRefDes)) = 0 Then Exit Sub
    If Not modConnectors.RemoveConnectorInstance(Trim$(sRefDes)) Then
        MsgBox "No connector instance '" & sRefDes & "' found.", vbExclamation
    End If
End Sub
```

In `build/layout.py`, extend `HOME_BUTTONS`:

```python
HOME_BUTTONS = [
    ("New Harness", "modChart.NewHarness", 20, 220, 120, 32),
    ("Add Connector", "modConnectorUI.ShowAddConnector", 20, 260, 120, 32),
    ("Remove Connector", "modConnectorUI.ShowRemoveConnector", 150, 260, 130, 32),
    ("Manage Library", "modConnectorUI.ShowManageLibrary", 20, 300, 120, 32),
]
```

- [ ] **Step 7: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_picker_form.py -v
```

Expected: 15 passed.

- [ ] **Step 8: Run the whole suite from a clean build**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: all passed (220 after Task 3 + 4 from Task 4 + 15 from Task 5 = 239).

- [ ] **Step 9: Manually verify in Excel**

Open `dist/HarnessCreator.xlsm`. Add Connector should list the library, add an instance, and populate `_Snapshot`. Manage Library's Edit should reopen the editor pre-loaded with existing pins; Delete should prompt and then remove the entry. Remove Connector should prompt for a ref des and clear its chart references. Record the result in the commit message, not as a test - matching 2b Task 8's precedent.

- [ ] **Step 10: Commit**

```bash
git add build/form_layout.py build/build.py build/layout.py src/vba/forms/ src/vba/modConnectorUI.bas tests/test_picker_form.py
git commit -m "feat: add the connector picker, library browser, and Home commands"
```

---

## Self-Review

**Spec coverage for this sub-plan.** "Connector picker" and "snapshot embedding" from the phase-2 line item are both covered (Tasks 1, 2, 5). Ref des rename and its collision rejection - required by the spec's "Connector instances and reference designators" section and its Testing section, but missing from Phase 1 - is covered in full (Task 3), including the reverting-the-cell behavior the Testing section implies ("a colliding rename is rejected") but the bare data function alone cannot guarantee without the sheet-level wiring. Remove Connector (Task 4) and Manage Library (Task 5) were confirmed in scope by your "Full scope" answer.

**Deliberately deferred to 2d, not gaps in this plan:** the actual Import/Export logic behind `frmManageLibrary`'s two buttons - the buttons exist now so 2d only adds code-behind, not layout.

**A scope decision worth confirming:** `modSnapshot.SnapshotConnector`'s photo step tries `CachePhotoPath` first (matching the spec's "the cache is what the editor loads") and falls back to copying the shape directly off the library's own `Photos` sheet when no cache file exists yet - covering the case where a connector was defined in this same session and the cache write from 2b's `SaveConnector` hasn't happened yet, or a test seeds the library directly without going through `SaveConnector`. This is a reasonable interpretation, not a spec quote - flag it if a stricter cache-only rule was intended.

**Type consistency.** `RenameRefDes`'s collision check and `RemoveConnectorInstance`'s row lookup both use case-insensitive `StrComp(..., vbTextCompare)`, matching every other ref-des comparison since Phase 1's `modConnectors.PinCountFor`. `SnapshotConnector`'s pin copy loop matches 2b's `LoadScratchPins`/`SaveConnector` field-by-field pattern exactly - the same 7-element array shape moves between all three without reinterpretation.

**No placeholders.** Every step contains complete, runnable code. The one deliberately-unwired piece (Import/Export) is a documented, tested absence (`test_manage_library_import_export_are_unwired_for_now`), not a stub standing in for real logic.
