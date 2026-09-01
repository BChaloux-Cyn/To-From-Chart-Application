# Phase 3e: Harness Load Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Open Harness reads a chosen saved `.xlsx` and reconstructs the Creator's entire working state from it: `_Snapshot`, the `Connectors` instance list (derived from the file's `CONN_<RefDes>` sheets, since no such list is saved directly - see 3a/3b), the title block, and every chart row, with pin dependent-dropdown validation rebuilt across the board. No validation runs (Check Drawing does not exist until Phase 4 - confirmed during this plan's discussion); Load only clears `Check`.

**Architecture:** A new layer 0 module, `modHarnessLoad.bas`, holds the copy-in-the-opposite-direction primitives 3a's `modHarnessBuild.bas` does not provide (that module's functions all assume `ThisWorkbook` is the source and a freshly built workbook is the destination; here it is the reverse, and the destination is the Creator's own persistent, already-built sheets, which carry live `Worksheet_Change` handlers that must not fire mid-load). `modHarnessActions.LoadHarness` (layer 1) orchestrates them and validates the chosen file is actually a saved harness before touching any Creator state. `modHarnessUI.OpenHarness` (layer 2) owns the `GetOpenFilename` dialog and `Workbooks.Open`/`.Close`.

**Tech Stack:** VBA (Excel 16.0 COM automation), Python 3.13/pywin32/pytest.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Depends on:** 3a, 3b, 3c, and 3d, implemented and merged - this plan reads the exact sheet shapes (join-key columns, `CONN_<RefDes>` pages, the metadata cell) those four plans produce.

**Part of Phase 3** (see 3a's header for the full breakdown). This is 3e, the last implementation sub-plan - `phase-3-manual-verification.md` and 3f (docs) follow it.

## Global Constraints

- **Decision (reaffirmed from 3a): Load runs no validation.** It clears `Check` (via `modChart.NewHarness`, called first) and stops there. Check Drawing is Phase 4.
- **Decision: no unsaved-changes prompt before Open Harness.** The spec assigns "if there are unsaved changes, prompts to save first" to Export PDF, not to Load - this plan does not invent an equivalent guard for Open Harness. Revisit alongside Export PDF in Phase 4 if wanted; out of scope here.
- **Every bulk write into a live Creator sheet (`Harness`, `Connectors`) must guard `Application.EnableEvents`**, exactly as `modChart.NewHarness` and `modConnectors.RemoveConnectorInstance` already do - both sheets have `Worksheet_Change` handlers (`shHarness.evt`, `shConnectors.evt`) that must not fire once per cell during a bulk load. `_Snapshot` needs no such guard; it is very hidden and has no code-behind.
- **`ReadConnector`'s payload array indexing must use `LBound(vFields) + <column constant> - 1`, never a hardcoded offset.** `docs/superpowers/plans/2026-08-28-ui-logic-separation-design.md` documents that this exact array re-bases from 1 to 0 when it crosses `Application.Run` (a COM caller) but stays 1-based for an in-process VBA-to-VBA call (which is what this plan does, calling `modLibrary.ReadConnector` directly from `modHarnessLoad`, not through `Application.Run`) - `modLibrary.WriteConnector` already uses the `LBound`-relative form for exactly this reason, and this plan follows the same idiom rather than assuming either fixed base.
- Every VBA module starts with `Option Explicit`. `modHarnessLoad.bas` is layer 0; `modHarnessActions.LoadHarness` is layer 1 (no dialogs, no `Workbooks.Open`/`.Close`); `modHarnessUI.OpenHarness` is layer 2.

## File Structure

| File | Responsibility |
|---|---|
| `src/vba/modHarnessLoad.bas` | Layer 0. `CopySnapshotInto`, `CopyTitleBlockValues`, `CopyChartValues`, `RebuildConnectorInstances` - all copy Creator-ward, the reverse of `modHarnessBuild`'s Save-ward functions. |
| `src/vba/modHarnessActions.bas` | Modified: adds `LoadHarness(srcWb As Workbook) As Variant`. |
| `src/vba/modHarnessUI.bas` | Modified: adds `OpenHarness()`. |
| `src/vba/modContract.bas` | Modified: adds `HARNESS_LOADED` (payload `LONG`) and `HARNESS_LOAD_FAILED` (payload `STRING`). |
| `src/vba/modMessages.bas` | Modified: message text for the two new outcomes. |
| `build/build.py` | Modified: adds `modHarnessLoad.bas` to `VBA_MODULES`; adds the Open Harness Home button. |
| `build/layout.py` | Modified: adds the Open Harness Home button. |
| `tests/test_layering.py` | Modified: adds `modHarnessLoad` to `LAYER0`. |
| `tests/test_harness_load.py` | `modHarnessLoad`'s functions and `modHarnessActions.LoadHarness`. |
| `tests/test_harness_save_integration.py` | Modified: the round trip now loads the saved file back into a second, independent Creator instance and asserts the reconstructed state. |

---

### Task 1: Register the new outcome codes

**Files:**
- Modify: `src/vba/modContract.bas`
- Modify: `src/vba/modMessages.bas`
- Test: `tests/test_contract.py`, `tests/test_messages.py` (additions)

**Interfaces:**
- Produces: `HARNESS_LOADED` (payload `LONG`, used-row count), `HARNESS_LOAD_FAILED` (payload `STRING`, reason).

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_contract.py`:

```python
def test_harness_loaded_declares_a_long_payload(wb):
    assert run(wb, "modContract.PayloadKind", "HARNESS_LOADED") == "LONG"


def test_harness_load_failed_declares_a_string_payload(wb):
    assert run(wb, "modContract.PayloadKind", "HARNESS_LOAD_FAILED") == "STRING"
```

Append to `tests/test_messages.py`:

```python
def test_message_for_harness_loaded(wb):
    result = run(wb, "modContract.Success", "HARNESS_LOADED", 5)
    assert run(wb, "modMessages.MessageFor", result) == "Loaded. 5 wire(s) read."


def test_message_for_harness_load_failed(wb):
    result = run(wb, "modContract.Failure", "HARNESS_LOAD_FAILED", "not a harness file")
    assert run(wb, "modMessages.MessageFor", result) == \
        "Could not load the harness: not a harness file."
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
python -m pytest tests/test_contract.py tests/test_messages.py -v -k "harness_load or harness_loaded"
```

Expected: FAIL.

- [ ] **Step 3: Register the codes**

In `src/vba/modContract.bas`, extend `OutcomeCodes`'s final line:

```vb
        "HARNESS_SAVED", "HARNESS_SAVE_FAILED", "HARNESS_LOADED", "HARNESS_LOAD_FAILED")
```

Extend `PayloadKind`:

```vb
        Case "HARNESS_LOADED"
            PayloadKind = KIND_LONG
        Case "HARNESS_LOAD_FAILED"
            PayloadKind = KIND_STRING
```

- [ ] **Step 4: Add the message text**

In `src/vba/modMessages.bas`, add before `Case Else`:

```vb
        Case "HARNESS_LOADED"
            MessageFor = "Loaded. " & CStr(vPayload) & " wire(s) read."
        Case "HARNESS_LOAD_FAILED"
            MessageFor = "Could not load the harness: " & CStr(vPayload) & "."
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_contract.py tests/test_messages.py -v -k "harness_load or harness_loaded"
```

Expected: 4 passed.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modContract.bas src/vba/modMessages.bas tests/test_contract.py tests/test_messages.py
git commit -m "feat: register harness load outcome codes"
```

---

### Task 2: Copy the snapshot and title block into the Creator

**Files:**
- Create: `src/vba/modHarnessLoad.bas`
- Test: `tests/test_harness_load.py`

**Interfaces:**
- Consumes: `modSnapshot.SNAP_CONN_FIRST_ROW`/`LAST_ROW`, `SNAP_PINS_FIRST_ROW`/`LAST_ROW`, `modLibrary.LIB_FIELD_COUNT`/`PIN_FIELD_COUNT` (existing).
- Produces:
  - VBA `modHarnessLoad.CopySnapshotInto(wsSrcSnapshot As Worksheet, wsDestSnapshot As Worksheet)` - copies both fixed-row blocks and every photo Shape; clears whatever shapes `wsDestSnapshot` already had first, so a photo from a previously loaded harness does not linger.
  - VBA `modHarnessLoad.CopyTitleBlockValues(wsSrcHarness As Worksheet, wsDestHarness As Worksheet)` - copies the seven plain title-block values, then calls `modChart.SetLengthUnits` with the source's units value so the Length column header and `TB_Units` named range update correctly (a plain value copy would leave the header text stale).

- [ ] **Step 1: Write the failing test**

Create `tests/test_harness_load.py`:

```python
from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_copy_snapshot_into_replaces_prior_contents(wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    srcWb = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", srcWb)
        wsSrcSnap = srcWb.Worksheets("_Snapshot")
        fields = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
                  4, "", "PHOTO_DTM-04P", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
        run(wb, "modLibrary.WriteConnector", wsSrcSnap, 2, 201, fields)
        run(wb, "modLibrary.WritePin", wsSrcSnap, 211, 2210, ("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1))
        run(wb, "modLibrary.EmbedConnectorPhoto", wsSrcSnap, "DTM-04P", str(photo_path))

        wsDestSnap = wb.Worksheets("_Snapshot")
        # Stale data from a previous session, which the copy must overwrite/replace.
        run(wb, "modLibrary.WriteConnector", wsDestSnap, 2, 201,
            ("OLD-ID", "Old Connector", "", "", "Connector", 2, "", "", "", "", "Local"))
        run(wb, "modLibrary.EmbedConnectorPhoto", wsDestSnap, "OLD-ID", str(photo_path))

        run(wb, "modHarnessLoad.CopySnapshotInto", wsSrcSnap, wsDestSnap)

        result = run(wb, "modLibrary.ReadConnector", wsDestSnap, 2, 201, "DTM-04P")
        assert tuple(result) == fields
        assert wsDestSnap.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"
        assert wsDestSnap.Shapes.Count == 1  # the stale OLD-ID photo is gone
    finally:
        srcWb.Close(SaveChanges=False)


def test_copy_title_block_values_updates_length_units_header(wb, app):
    srcWb = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", srcWb)
        wsSrc = srcWb.Worksheets("Harness")
        wsSrc.Range("B2").Value = "Loaded Harness"
        wsSrc.Range("E2").Value = "HN-200"
        wsSrc.Range("H4").Value = "mm"

        wsDest = wb.Worksheets("Harness")
        run(wb, "modHarnessLoad.CopyTitleBlockValues", wsSrc, wsDest)

        assert wsDest.Range("B2").Value == "Loaded Harness"
        assert wsDest.Range("E2").Value == "HN-200"
        assert wsDest.Cells(6, 7).Value == "Length (mm)"
    finally:
        srcWb.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_harness_load.py -v
```

Expected: FAIL - `modHarnessLoad` does not exist.

- [ ] **Step 3: Write the module**

Create `src/vba/modHarnessLoad.bas`:

```vb
Attribute VB_Name = "modHarnessLoad"
Option Explicit

Public Sub CopySnapshotInto(wsSrcSnapshot As Worksheet, wsDestSnapshot As Worksheet)
    Dim r As Long, c As Long, shp As Shape

    For r = modSnapshot.SNAP_CONN_FIRST_ROW To modSnapshot.SNAP_CONN_LAST_ROW
        For c = 1 To modLibrary.LIB_FIELD_COUNT
            wsDestSnapshot.Cells(r, c).Value = wsSrcSnapshot.Cells(r, c).Value
        Next c
    Next r

    For r = modSnapshot.SNAP_PINS_FIRST_ROW To modSnapshot.SNAP_PINS_LAST_ROW
        For c = 1 To modLibrary.PIN_FIELD_COUNT
            wsDestSnapshot.Cells(r, c).Value = wsSrcSnapshot.Cells(r, c).Value
        Next c
    Next r

    Do While wsDestSnapshot.Shapes.Count > 0
        wsDestSnapshot.Shapes(1).Delete
    Loop

    For Each shp In wsSrcSnapshot.Shapes
        shp.Copy
        wsDestSnapshot.Paste
        wsDestSnapshot.Shapes(wsDestSnapshot.Shapes.Count).Name = shp.Name
    Next shp
End Sub

Private Const TB_VALUE_CELLS_NO_UNITS As String = "B2,E2,H2,B3,E3,H3,B4"

Public Sub CopyTitleBlockValues(wsSrcHarness As Worksheet, wsDestHarness As Worksheet)
    Dim vCells As Variant, i As Long, sCell As String
    Dim bEvents As Boolean

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    vCells = Split(TB_VALUE_CELLS_NO_UNITS, ",")
    For i = LBound(vCells) To UBound(vCells)
        sCell = CStr(vCells(i))
        wsDestHarness.Range(sCell).Value = wsSrcHarness.Range(sCell).Value
    Next i

CleanUp:
    Application.EnableEvents = bEvents
    modChart.SetLengthUnits CStr(wsSrcHarness.Range("H4").Value)
End Sub
```

`SetLengthUnits` runs after the `EnableEvents` guard is released (it manages its own guard internally, per `modChart.bas`) - calling it inside the guarded region would be harmless but redundant; calling it after is what lets its own `TB_Units` named-range write and header rewrite behave exactly as it does everywhere else it is called.

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_harness_load.py -v
```

Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modHarnessLoad.bas tests/test_harness_load.py
git commit -m "feat: copy a loaded harness's snapshot and title block into the Creator"
```

---

### Task 3: Copy the chart and rebuild connector instances

**Files:**
- Modify: `src/vba/modHarnessLoad.bas`
- Test: `tests/test_harness_load.py` (additions)

**Interfaces:**
- Consumes: `modChart.CHART_FIRST_ROW`/`LAST_ROW` (existing), `modConnectors.CONN_FIRST_ROW` (existing), `modConnectorPage.CONN_META_COL` (3b), `modLibrary.LIB_COL_NAME`/`PARTNUM`/`TYPE`/`PINCOUNT` (existing).
- Produces:
  - VBA `modHarnessLoad.CopyChartValues(wsSrcHarness As Worksheet, wsDestHarness As Worksheet) As Long` - copies columns A-K, rows 7-1006, returns the used-row count (same convention as 3a's `CopyChartRows`).
  - VBA `modHarnessLoad.RebuildConnectorInstances(srcWb As Workbook, wsSnapshot As Worksheet, wsConnectors As Worksheet) As Long` - one row per `CONN_<RefDes>` sheet found in `srcWb`, resolved against the just-copied `_Snapshot`. Returns the count written.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_harness_load.py`:

```python
def test_copy_chart_values_round_trips_and_counts_used_rows(wb, app):
    srcWb = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", srcWb)
        wsSrc = srcWb.Worksheets("Harness")
        wsSrc.Cells(7, 1).Value = "J1"
        wsSrc.Cells(7, 9).Value = "J2"
        wsSrc.Cells(10, 1).Value = "J1"
        wsSrc.Cells(10, 9).Value = "J2"

        wsDest = wb.Worksheets("Harness")
        n = run(wb, "modHarnessLoad.CopyChartValues", wsSrc, wsDest)

        assert n == 2
        assert wsDest.Cells(7, 1).Value == "J1"
        assert wsDest.Cells(10, 9).Value == "J2"
    finally:
        srcWb.Close(SaveChanges=False)


def test_rebuild_connector_instances_reads_ref_des_from_sheet_names(wb, app, tmp_path):
    srcWb = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", srcWb)
        wsSnap = srcWb.Worksheets("_Snapshot")
        fields = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
                  4, "", "PHOTO_DTM-04P", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
        run(wb, "modLibrary.WriteConnector", wsSnap, 2, 201, fields)

        page = srcWb.Worksheets.Add(After=srcWb.Worksheets(srcWb.Worksheets.Count))
        page.Name = "CONN_J1"
        run(wb, "modConnectorPage.WriteMetadata", page, "DTM-04P")

        wsDestSnap = wb.Worksheets("_Snapshot")
        run(wb, "modHarnessLoad.CopySnapshotInto", wsSnap, wsDestSnap)

        wsDestConn = wb.Worksheets("Connectors")
        n = run(wb, "modHarnessLoad.RebuildConnectorInstances", srcWb, wsDestSnap, wsDestConn)

        assert n == 1
        assert wsDestConn.Cells(2, 1).Value == "J1"
        assert wsDestConn.Cells(2, 2).Value == "DTM-04P"
        assert wsDestConn.Cells(2, 3).Value == "Deutsch DTM 4-way"
        assert int(wsDestConn.Cells(2, 6).Value) == 4
    finally:
        srcWb.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
python -m pytest tests/test_harness_load.py -v -k "copy_chart_values or rebuild_connector_instances"
```

Expected: FAIL - neither function exists.

- [ ] **Step 3: Write the functions**

Append to `src/vba/modHarnessLoad.bas`:

```vb
Public Function CopyChartValues(wsSrcHarness As Worksheet, wsDestHarness As Worksheet) As Long
    Dim r As Long, c As Long, n As Long, sFrom As String, sTo As String
    Dim bEvents As Boolean

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    For r = modChart.CHART_FIRST_ROW To modChart.CHART_LAST_ROW
        For c = 1 To 11
            wsDestHarness.Cells(r, c).Value = wsSrcHarness.Cells(r, c).Value
        Next c
        sFrom = Trim$(CStr(wsSrcHarness.Cells(r, 1).Value))
        sTo = Trim$(CStr(wsSrcHarness.Cells(r, 9).Value))
        If Len(sFrom) > 0 Or Len(sTo) > 0 Then n = n + 1
    Next r

CleanUp:
    Application.EnableEvents = bEvents
    CopyChartValues = n
End Function

Public Function RebuildConnectorInstances(srcWb As Workbook, wsSnapshot As Worksheet, _
                                          wsConnectors As Worksheet) As Long
    Dim sh As Worksheet, sRefDes As String, sConnectorID As String, vFields As Variant
    Dim r As Long, n As Long, bEvents As Boolean

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    r = modConnectors.CONN_FIRST_ROW
    For Each sh In srcWb.Worksheets
        If Left$(sh.Name, 5) = "CONN_" Then
            sRefDes = Mid$(sh.Name, 6)
            sConnectorID = Trim$(CStr(sh.Cells(1, modConnectorPage.CONN_META_COL).Value))

            vFields = modLibrary.ReadConnector(wsSnapshot, modSnapshot.SNAP_CONN_FIRST_ROW, _
                modSnapshot.SNAP_CONN_LAST_ROW, sConnectorID)
            If Not IsEmpty(vFields) Then
                wsConnectors.Cells(r, 1).Value = sRefDes
                wsConnectors.Cells(r, 2).Value = sConnectorID
                wsConnectors.Cells(r, 3).Value = vFields(LBound(vFields) + modLibrary.LIB_COL_NAME - 1)
                wsConnectors.Cells(r, 4).Value = vFields(LBound(vFields) + modLibrary.LIB_COL_PARTNUM - 1)
                wsConnectors.Cells(r, 5).Value = vFields(LBound(vFields) + modLibrary.LIB_COL_TYPE - 1)
                wsConnectors.Cells(r, 6).Value = vFields(LBound(vFields) + modLibrary.LIB_COL_PINCOUNT - 1)
                r = r + 1
                n = n + 1
            End If
        End If
    Next sh

CleanUp:
    Application.EnableEvents = bEvents
    RebuildConnectorInstances = n
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_harness_load.py -v -k "copy_chart_values or rebuild_connector_instances"
```

Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modHarnessLoad.bas tests/test_harness_load.py
git commit -m "feat: copy a loaded harness's chart and rebuild its connector instances"
```

---

### Task 4: The LoadHarness action

**Files:**
- Modify: `src/vba/modHarnessActions.bas`
- Test: `tests/test_harness_load.py` (additions)

**Interfaces:**
- Consumes: `modHarnessLoad.CopySnapshotInto`/`CopyTitleBlockValues`/`CopyChartValues`/`RebuildConnectorInstances` (Tasks 2-3), `modChart.NewHarness`/`RebuildPinValidation`/`CHART_FIRST_ROW`/`LAST_ROW`/`COL_FROM_CONN`/`COL_TO_CONN` (existing), `modState.SetState`/`ClearDirty` (existing).
- Produces: VBA `modHarnessActions.LoadHarness(srcWb As Workbook) As Variant` - outcomes `HARNESS_LOADED` (payload `Long`, used-row count) and `HARNESS_LOAD_FAILED` (payload `String`, reason).

- [ ] **Step 1: Write the failing test**

Append to `tests/test_harness_load.py`:

```python
from tests.conftest import run_action


def test_load_harness_rejects_a_file_with_no_harness_sheet(wb, app):
    srcWb = app.Workbooks.Add()
    try:
        result = run_action(wb, "modHarnessActions.LoadHarness", srcWb)
        assert result.ok is False
        assert result.outcome == "HARNESS_LOAD_FAILED"
    finally:
        srcWb.Close(SaveChanges=False)


def test_load_harness_reconstructs_creator_state(wb, app, tmp_path):
    srcWb = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", srcWb)
        wsSrcHarness = srcWb.Worksheets("Harness")
        wsSrcHarness.Range("B2").Value = "Loaded Harness"
        wsSrcHarness.Range("E2").Value = "HN-300"
        wsSrcHarness.Cells(7, 1).Value = "J1"
        wsSrcHarness.Cells(7, 2).Value = 1
        wsSrcHarness.Cells(7, 9).Value = "J2"
        wsSrcHarness.Cells(7, 10).Value = 1

        wsSrcSnap = srcWb.Worksheets("_Snapshot")
        fields = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
                  4, "", "PHOTO_DTM-04P", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
        run(wb, "modLibrary.WriteConnector", wsSrcSnap, 2, 201, fields)

        page = srcWb.Worksheets.Add(After=srcWb.Worksheets(srcWb.Worksheets.Count))
        page.Name = "CONN_J1"
        run(wb, "modConnectorPage.WriteMetadata", page, "DTM-04P")

        result = run_action(wb, "modHarnessActions.LoadHarness", srcWb)
        assert result.ok is True
        assert result.outcome == "HARNESS_LOADED"
        assert result.payload == 1

        wsDestHarness = wb.Worksheets("Harness")
        assert wsDestHarness.Range("B2").Value == "Loaded Harness"
        assert wsDestHarness.Cells(7, 1).Value == "J1"

        wsDestConn = wb.Worksheets("Connectors")
        assert wsDestConn.Cells(2, 1).Value == "J1"
        assert wsDestConn.Cells(2, 2).Value == "DTM-04P"

        # Pin dropdown validation rebuilt against the reconstructed Connectors sheet.
        validation_formula = wsDestHarness.Cells(7, 2).Validation.Formula1
        assert validation_formula == "1,2,3,4"

        assert run(wb, "modState.GetState", "Dirty") == "FALSE"
    finally:
        srcWb.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_harness_load.py -v -k load_harness
```

Expected: FAIL - `LoadHarness` does not exist.

- [ ] **Step 3: Write the function**

Append to `src/vba/modHarnessActions.bas`:

```vb
Private Function SheetExists(wb As Workbook, ByVal sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Function LoadHarness(srcWb As Workbook) As Variant
    If Not SheetExists(srcWb, "Harness") Or Not SheetExists(srcWb, "_Snapshot") Then
        LoadHarness = modContract.Failure("HARNESS_LOAD_FAILED", "not a harness file")
        Exit Function
    End If

    modChart.NewHarness

    Dim wsSrcHarness As Worksheet, wsSrcSnapshot As Worksheet
    Set wsSrcHarness = srcWb.Worksheets("Harness")
    Set wsSrcSnapshot = srcWb.Worksheets("_Snapshot")

    Dim wsDestHarness As Worksheet, wsDestSnapshot As Worksheet, wsDestConnectors As Worksheet
    Set wsDestHarness = ThisWorkbook.Worksheets(modChart.CHART_SHEET)
    Set wsDestSnapshot = ThisWorkbook.Worksheets("_Snapshot")
    Set wsDestConnectors = ThisWorkbook.Worksheets(modConnectors.CONN_SHEET)

    modHarnessLoad.CopySnapshotInto wsSrcSnapshot, wsDestSnapshot
    modHarnessLoad.CopyTitleBlockValues wsSrcHarness, wsDestHarness
    modHarnessLoad.RebuildConnectorInstances srcWb, wsDestSnapshot, wsDestConnectors

    Dim nUsedRows As Long
    nUsedRows = modHarnessLoad.CopyChartValues(wsSrcHarness, wsDestHarness)

    Dim r As Long
    For r = modChart.CHART_FIRST_ROW To modChart.CHART_LAST_ROW
        modChart.RebuildPinValidation r, modChart.COL_FROM_CONN, False
        modChart.RebuildPinValidation r, modChart.COL_TO_CONN, False
    Next r

    modState.SetState "HarnessPath", srcWb.FullName
    modState.ClearDirty

    LoadHarness = modContract.Success("HARNESS_LOADED", nUsedRows)
End Function
```

`RebuildConnectorInstances` runs before `CopyChartValues` deliberately: pin dropdown validation (the loop right after) reads `modConnectors.PinCountFor`, which scans the `Connectors` sheet - it must already hold the reconstructed instances before any row's dropdown is rebuilt, even though the chart values themselves do not depend on that order.

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_harness_load.py -v -k load_harness
```

Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modHarnessActions.bas tests/test_harness_load.py
git commit -m "feat: add the LoadHarness action"
```

---

### Task 5: Wire Open Harness and register the new module

**Files:**
- Modify: `src/vba/modHarnessUI.bas`
- Modify: `build/build.py`
- Modify: `build/layout.py`
- Modify: `tests/test_layering.py`

**Interfaces:**
- Consumes: `modHarnessActions.LoadHarness` (Task 4).
- Produces: `modHarnessUI.OpenHarness()`, wired as a Home button.

- [ ] **Step 1: Write the adapter**

Append to `src/vba/modHarnessUI.bas`:

```vb
Public Sub OpenHarness()
    Dim vPath As Variant
    vPath = Application.GetOpenFilename(FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If vPath = False Then Exit Sub

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(CStr(vPath))

    Dim vResult As Variant
    vResult = modHarnessActions.LoadHarness(srcWb)

    srcWb.Close SaveChanges:=False

    MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
End Sub
```

- [ ] **Step 2: Wire the Home button**

In `build/layout.py`, extend `HOME_BUTTONS`:

```python
HOME_BUTTONS = [
    ("New Harness", "modChart.NewHarness", 20, 220, 120, 32),
    ("Save Harness", "modHarnessUI.SaveHarness", 150, 220, 120, 32),
    ("Save Harness As", "modHarnessUI.SaveHarnessAs", 280, 220, 120, 32),
    ("Open Harness", "modHarnessUI.OpenHarness", 410, 220, 120, 32),
    ("Add Connector", "modConnectorUI.ShowAddConnector", 20, 260, 120, 32),
    ("Remove Connector", "modConnectorUI.ShowRemoveConnector", 150, 260, 130, 32),
    ("Manage Library", "modConnectorUI.ShowManageLibrary", 20, 300, 120, 32),
]
```

- [ ] **Step 3: Register the new module**

In `build/build.py`, add `"modHarnessLoad.bas"` to `VBA_MODULES` (after `"modHarnessBuild.bas"`).

In `tests/test_layering.py`, add `"modHarnessLoad"` to `LAYER0`.

- [ ] **Step 4: Run the full suite**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: everything passes.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modHarnessUI.bas build/build.py build/layout.py tests/test_layering.py
git commit -m "feat: wire Open Harness to the Home sheet"
```

---

### Task 6: Full save-then-load round trip

**Files:**
- Modify: `tests/test_harness_save_integration.py`

**Interfaces:**
- Consumes: everything from Tasks 1-5, plus 3a-3d's existing `test_full_harness_round_trips_through_a_saved_file` fixture and save path.

This closes the loop the spec's own testing section describes only in halves ("Library round trip", "Render") by proving Save's output is exactly what Load can reconstruct from, using a second, independent Creator workbook so the test cannot pass by coincidentally reusing in-memory state from the Save half.

- [ ] **Step 1: Add the load half to the existing test**

Extend `test_full_harness_round_trips_through_a_saved_file` (which already builds a harness, saves it, and reopens it directly to assert file contents) with a final phase: open a **second, independent** copy of the Creator artifact, load the saved file into it, and assert its state matches what was originally built.

```python
def test_full_harness_round_trips_through_a_saved_file(wb, app, artifact, tmp_path):
    ...  # existing build-and-save body, unchanged through the `reopened` block

    # Load half: a second, independent Creator instance proves the save/load
    # pair actually reconstructs state, not just that the file looks right.
    wb2 = app.Workbooks.Open(str(artifact))
    try:
        srcWb = app.Workbooks.Open(str(saved_path))
        try:
            load_result = run_action(wb2, "modHarnessActions.LoadHarness", srcWb)
            assert load_result.ok is True
            assert load_result.outcome == "HARNESS_LOADED"
        finally:
            srcWb.Close(SaveChanges=False)

        loaded_harness = wb2.Worksheets("Harness")
        assert loaded_harness.Range("B2").Value == "Dash Harness"
        assert loaded_harness.Cells(7, 1).Value == "J1"
        assert loaded_harness.Cells(7, 4).Value == "12V_SW"

        loaded_conn = wb2.Worksheets("Connectors")
        assert loaded_conn.Cells(2, 1).Value == "J1"
        assert loaded_conn.Cells(2, 2).Value == "DTM-04P"

        loaded_snap = wb2.Worksheets("_Snapshot")
        loaded_fields = run(wb2, "modLibrary.ReadConnector", loaded_snap, 2, 201, "DTM-04P")
        assert tuple(loaded_fields) == fields

        assert run(wb2, "modState.GetState", "HarnessPath") == str(saved_path)
        assert run(wb2, "modState.GetState", "Dirty") == "FALSE"
    finally:
        wb2.Close(SaveChanges=False)
```

(`artifact`, `fields`, and `saved_path` are already in scope from the test's existing body - `artifact` is the session-scoped fixture pointing at `dist/HarnessCreator.xlsm`, reused here for a second, independent workbook handle rather than the `wb` fixture, which this test must not mutate further after its own save/reopen phase.)

- [ ] **Step 2: Run the test**

Run:

```bash
python -m pytest tests/test_harness_save_integration.py -v
```

Expected: 1 passed.

- [ ] **Step 3: Run the whole suite from a clean build**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: everything passes. This is the last automated gate before Phase 3's manual-verification pass (`phase-3-manual-verification.md`) and 3f (docs).

- [ ] **Step 4: Commit**

```bash
git add tests/test_harness_save_integration.py
git commit -m "test: prove a saved harness loads back into an independent Creator instance"
```

---

## Self-Review

**Spec coverage for this sub-plan.** "Load opens the chosen .xlsx, copies _Snapshot into the Creator, reads the title block and every used chart row, closes the file" is implemented exactly, with "closes the file" left to the layer-2 adapter (`modHarnessUI.OpenHarness`) per the same split 3a used for Save. "Then runs validation and reports findings on Check" is deliberately not implemented - restated from Global Constraints, this is Phase 4's addition once Check Drawing exists; `modChart.NewHarness`'s existing `Check`-clearing behavior is what Load relies on instead.

**Why reconstructing `Connectors` from `CONN_<RefDes>` sheet names is necessary at all, and why it is not a gap in 3a/3b.** The spec's own list of what a saved harness contains has no separate Ref-Des-to-ConnectorID table - flagged explicitly in 3a's and 3b's self-reviews as an obligation deferred to this plan. Without it, a loaded harness's chart would have From/To Conn dropdowns with no valid list to validate against and `PinCountFor` would return 0 for every ref des, breaking every dependent Pin dropdown - this plan's `RebuildConnectorInstances` plus the pin-validation rebuild loop in `LoadHarness` is what prevents that.

**Why `ReadConnector`'s payload is indexed `LBound(vFields) + <constant> - 1` instead of a fixed `-1` or `0` offset.** Restated from Global Constraints because it is easy to get backwards: this plan calls `modLibrary.ReadConnector` directly, in-process, not through `Application.Run`, so the array stays 1-based as declared (`ReDim vResult(1 To 11)`) - a hardcoded `vFields(constant - 1)` would be correct here but silently wrong if this code were ever refactored to cross a COM boundary. Matching `WriteConnector`'s own `LBound`-relative form keeps it correct either way, which is exactly why that module already does it that way.

**Why `Application.EnableEvents` is guarded in every function that writes to `Harness` or `Connectors`, but not in `CopySnapshotInto`.** `_Snapshot` is very hidden and carries no `Worksheet_Change` handler; `Harness` and `Connectors` do (`shHarness.evt`, `shConnectors.evt`), and a bulk load firing one change event per cell would be needlessly slow and could interleave with `modChart.RebuildPinValidation`'s own per-row work in ways never exercised elsewhere in this codebase. This mirrors `modChart.NewHarness`'s and `modConnectors.RemoveConnectorInstance`'s existing guards exactly, including the save/restore-prior-value pattern (`bEvents = Application.EnableEvents` rather than assuming `True`), which is what lets `LoadHarness`'s call to `modChart.NewHarness` and this plan's own guarded functions nest safely without either one clobbering the other's restore.

**Why stale `_Snapshot` connector/pin rows from a previous session are left in place, while stale photo shapes are explicitly cleared.** A previously loaded harness with more connectors than the newly loaded one leaves old rows sitting past the new data's extent - harmless, because nothing in the newly reconstructed `Connectors` sheet references those old `ConnectorID`s by ref des, and `ReadConnector`/`ReadPinsForConnector` only ever look up an ID something currently references. Photo shapes are different: they are named artifacts (`PHOTO_<ConnectorID>`) that would otherwise accumulate across every Load in a session with no corresponding cleanup, which is why `CopySnapshotInto` clears them first. This is a deliberate, asymmetric choice, not an oversight - documented here so a future session does not "fix" the row leftovers into a change that adds complexity for no observable bug.

**Type consistency.** `LoadHarness`'s `HARNESS_LOADED` payload is a `Long` (used-row count from `CopyChartValues`), matching `HARNESS_SAVED`'s payload convention from 3a exactly - a caller reading either outcome's payload never has to check which one it got before treating it as a count.

**No placeholders.** Every step contains complete VBA and Python, including the exact `EnableEvents` guard pattern reused from existing code and the exact assertions the round-trip test makes against the reconstructed state - no "reconstruct the remaining fields" or "similar to Save" stand-ins.
