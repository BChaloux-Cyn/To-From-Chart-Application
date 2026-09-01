# Phase 3a: Harness Workbook Shell and Save Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce a `SaveHarness`/`SaveHarnessAs` Home command that writes the Creator's current state to a new macro-free `<name>.xlsx` containing a `Harness` sheet (title block, chart values, and the two hidden join-key columns the rest of Phase 3 renders against) and a very-hidden `_Snapshot` sheet copied verbatim from the Creator's own. Connector pages (`CONN_<RefDes>`), live pin-table formulas, and page setup are explicitly out of scope here - see 3b, 3c, 3d.

**Architecture:** Follows the layer 0/1/2 split from `docs/superpowers/specs/2026-08-28-ui-logic-separation-design.md` exactly. `modHarnessBuild.bas` (layer 0) is the sheet-population primitive, taking already-created `Worksheet`/`Workbook` objects and writing values/formulas into them - it never calls `Workbooks.Add`, `.SaveAs`, or `.Close`. `modHarnessActions.SaveHarness` (layer 1) orchestrates the primitives into one transaction and returns the `modContract` envelope. `modHarnessUI.bas` (layer 2, a plain standard module playing the same role `modConnectorUI.bas` already does for forms) owns `Workbooks.Add`, the `GetSaveAsFilename` dialog, `.SaveAs`, `.Close`, and the `MsgBox`. This mirrors `modManageActions.ExportToWorkbook` / `frmManageLibrary.cmdExport_Click` precisely - that pair is the closest existing precedent for "build a fresh workbook's sheets, then let the caller save and close it."

**Tech Stack:** VBA (Excel 16.0 COM automation), Python 3.13/pywin32/pytest for the build and test harness.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Depends on:** Phase 1, Phase 2 (2a-2d), and the UI/logic-separation plan - all merged on `master`.

**Part of Phase 3**, split into five sub-plans plus a manual-verification pass and a docs pass (mirroring Phase 2's 2a-2e structure):
- **3a (this plan): Harness workbook shell and Save** - `Harness` sheet, hidden join-key columns, `_Snapshot` copy, Save/Save As.
- 3b: Connector page rendering - `CONN_<RefDes>` sheets, photo placement, oval callouts, leader lines (depends on 3a).
- 3c: Live pin-table formulas against 3a's join-key columns (depends on 3a, 3b).
- 3d: Page setup - print area, titles, scaling, footer, baked in at save time (depends on 3a, 3b, 3c).
- 3e: Harness load - Open Harness reconstructs the Creator's state from a saved file (depends on 3a-3d).
- `phase-3-manual-verification.md`: consolidated manual-only checklist, run after 3e, before 3f.
- 3f: retrospective design docs and student user guide update (after manual verification).

## Global Constraints

Copied verbatim from the spec, plus decisions fixed during this plan's discussion. Every task's requirements implicitly include these.

- Harness files are macro-free `.xlsx`. `SaveAs ... FileFormat:=51` on a workbook that never had a VBA module imported into it stays macro-free, matching `modManageActions`'s existing export path - never import a `.bas` module into a harness workbook.
- The visible to-from chart is the canonical wire data; the Creator reads it back verbatim on load (implemented in 3e, not here).
- Windows and desktop Excel only. Formulas must work on Excel 2016+ (INDEX/MATCH, never XLOOKUP) - relevant to 3c, noted here since the join-key columns this plan adds exist to serve those formulas.
- Every VBA module starts with `Option Explicit`.
- No `MsgBox`, `InputBox`, `GetOpenFilename`, `GetSaveAsFilename`, `.Show`, `Workbooks.Open`, or `DoEvents` in a layer 1 module (`modHarnessActions.bas`). These are permitted only in `modHarnessUI.bas` (layer 2).
- `Array()` builds every `modContract` result; no `Option Base` directive anywhere.
- **Decision (this plan's discussion): the saved `Harness` sheet's chart occupies the same fixed row range as the Creator's own** (rows 7-1006, mirroring `modChart.CHART_FIRST_ROW`/`CHART_LAST_ROW`) rather than being trimmed to the last used row. This gives a student room to hand-add wire rows in the saved file and have 3c's live pin-table formulas (which read the same range) pick them up, and it lets page setup (3d) compute the printed area from actual content rather than this plan needing to guess a trim point.
- **Decision: the two hidden join-key columns are `L` (`FromConn|FromPin`) and `M` (`ToConn|ToPin`)**, one formula per chart row, `=IF(A7="","",A7&"|"&B7)` and `=IF(I7="","",I7&"|"&J7)` shifted per row - never static text, so a student's hand edit to the chart after saving keeps the join keys (and therefore 3c's pin tables) correct with no macro involved.
- **Decision: Load (3e) does not run validation.** Check Drawing does not exist until Phase 4. 3a's Save does not touch the Creator's `Check` sheet at all - there is nothing to run yet.
- **Decision: page setup (print area, titles, scaling, footer) is entirely out of scope for this plan** - see 3d. The sheets this plan produces are otherwise plain (Excel's default print settings).
- **Decision: no connector-instance table is duplicated into the saved file in this plan.** 3e's Load will reconstruct the Creator's `Connectors` sheet from the `CONN_<RefDes>` sheet names plus a per-page metadata cell that 3b adds - not from anything 3a writes. This plan's `Harness` sheet and `_Snapshot` copy are exactly the two pieces the spec's "Harness rendering" section lists as saved (`CONN_` pages are 3b's addition to that same list).

## File Structure

| File | Responsibility |
|---|---|
| `src/vba/modHarnessBuild.bas` | Layer 0. Shapes a fresh `Workbook`'s sheets into `Harness` + `_Snapshot`; copies the Creator's title block, chart values, join-key formulas, and `_Snapshot` contents into them. |
| `src/vba/modHarnessActions.bas` | Layer 1. `SaveHarness(destWb)` - orchestrates `modHarnessBuild` into one transaction, returns the `modContract` envelope. |
| `src/vba/modHarnessUI.bas` | Layer 2. `SaveHarness()`/`SaveHarnessAs()` Home button macros - `Workbooks.Add`, the Save As dialog, `.SaveAs`, `.Close`, the confirmation `MsgBox`. Also owns `modState` bookkeeping (`HarnessPath`, dirty-clearing) after a successful save. |
| `src/vba/modContract.bas` | Modified: adds `HARNESS_SAVED` (payload `LONG`, the count of used wire rows) and `HARNESS_SAVE_FAILED` (payload `STRING`) to the outcome registry. |
| `src/vba/modMessages.bas` | Modified: adds message text for the two new outcomes. |
| `build/build.py` | Modified: adds `modHarnessBuild.bas`, `modHarnessActions.bas`, `modHarnessUI.bas` to `VBA_MODULES`. |
| `build/layout.py` | Modified: adds the Save Harness / Save Harness As Home buttons. |
| `tests/test_layering.py` | Modified: adds the three new modules to `LAYER0`/`LAYER1` (`modHarnessUI` is a launcher module like `modConnectorUI`, not added to `ADAPTERS` - see Task 6). |
| `tests/test_contract.py` | Modified: the new codes are covered by the existing generic tests (`test_every_declared_code_has_a_payload_kind` etc.) with no changes needed beyond the registry itself. |
| `tests/test_harness_build.py` | `modHarnessBuild`'s sheet-shaping and title-block/chart/snapshot copy functions. |
| `tests/test_harness_actions.py` | `modHarnessActions.SaveHarness`'s envelope. |
| `tests/test_harness_save_integration.py` | Full round trip: populate the Creator, save, reopen the saved file directly, assert its contents. |

---

### Task 1: Register the new outcome codes

**Files:**
- Modify: `src/vba/modContract.bas`
- Modify: `src/vba/modMessages.bas`
- Test: `tests/test_contract.py` (additions), `tests/test_messages.py` (additions)

**Interfaces:**
- Consumes: nothing new.
- Produces: outcome codes `HARNESS_SAVED` (payload `LONG`), `HARNESS_SAVE_FAILED` (payload `STRING`), registered in `modContract.OutcomeCodes`/`PayloadKind`, with message text in `modMessages.MessageFor`.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_contract.py`:

```python
def test_harness_saved_declares_a_long_payload(wb):
    assert run(wb, "modContract.PayloadKind", "HARNESS_SAVED") == "LONG"


def test_harness_save_failed_declares_a_string_payload(wb):
    assert run(wb, "modContract.PayloadKind", "HARNESS_SAVE_FAILED") == "STRING"
```

Append to `tests/test_messages.py`:

```python
def test_message_for_harness_saved(wb):
    result = run(wb, "modContract.Success", "HARNESS_SAVED", 12)
    assert run(wb, "modMessages.MessageFor", result) == "Saved. 12 wire(s) written."


def test_message_for_harness_save_failed(wb):
    result = run(wb, "modContract.Failure", "HARNESS_SAVE_FAILED", "destination workbook is not fresh")
    assert run(wb, "modMessages.MessageFor", result) == \
        "Could not save the harness: destination workbook is not fresh."
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
python -m pytest tests/test_contract.py tests/test_messages.py -v -k harness
```

Expected: FAIL - `PayloadKind` raises (unknown code) or returns wrong text; `MessageFor` returns `""`.

- [ ] **Step 3: Register the codes**

In `src/vba/modContract.bas`, extend `OutcomeCodes`:

```vb
Public Function OutcomeCodes() As Variant
    OutcomeCodes = Array( _
        "PLACED", "MOVED_ANCHOR", "BAD_PIN_COUNT", "PIN_LIMIT_REACHED", "NO_OP", _
        "SAVED", "ID_COLLISION", "SAVE_FAILED", _
        "CACHE_READY", "NEEDS_BACKFILL", _
        "OK", "MISSING_NAME_OR_PART", _
        "PIN_DELETED", "PIN_NOT_FOUND", _
        "ADDED", "ADD_FAILED", "CONNECTOR_NOT_FOUND", "CONNECTOR_DELETED", _
        "CONNECTOR_DELETED_CASCADED", _
        "CONNECTOR_IMPORTED", "CONNECTOR_IMPORTED_CASCADED", _
        "EXPORTED", "EXPORT_FAILED", "LIBRARY_EXPORTED", _
        "PHOTO_ATTACHED", "PHOTO_FAILED", _
        "RENAMED", "RENAME_REJECTED", "NO_RENAME", _
        "BULK_REBUILT", "CELLS_REBUILT", "UNITS_SET", _
        "INSTANCE_REMOVED", "INSTANCE_NOT_FOUND", _
        "HARNESS_SAVED", "HARNESS_SAVE_FAILED")
End Function
```

Extend `PayloadKind`'s `Select Case` with one more arm:

```vb
        Case "HARNESS_SAVED"
            PayloadKind = KIND_LONG
        Case "HARNESS_SAVE_FAILED"
            PayloadKind = KIND_STRING
```

- [ ] **Step 4: Add the message text**

In `src/vba/modMessages.bas`, add two more `Case` arms to `MessageFor` (before `Case Else`):

```vb
        Case "HARNESS_SAVED"
            MessageFor = "Saved. " & CStr(vPayload) & " wire(s) written."
        Case "HARNESS_SAVE_FAILED"
            MessageFor = "Could not save the harness: " & CStr(vPayload) & "."
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_contract.py tests/test_messages.py -v -k harness
```

Expected: 4 passed.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modContract.bas src/vba/modMessages.bas tests/test_contract.py tests/test_messages.py
git commit -m "feat: register harness save outcome codes"
```

---

### Task 2: Shape a fresh workbook into Harness + _Snapshot

**Files:**
- Create: `src/vba/modHarnessBuild.bas`
- Test: `tests/test_harness_build.py`

**Interfaces:**
- Consumes: nothing new.
- Produces:
  - VBA constants `SAVED_CHART_HEADER_ROW = 6`, `SAVED_CHART_FIRST_ROW = 7`, `SAVED_CHART_LAST_ROW = 1006` (mirroring `modChart`'s), `SAVED_COL_JOIN_FROM = 12`, `SAVED_COL_JOIN_TO = 13`
  - VBA `modHarnessBuild.BuildHarnessSheets(destWb As Workbook) As Boolean` - adds `Harness` and `_Snapshot`, deletes `destWb`'s original default sheet, sets `_Snapshot` very hidden. Returns `False` (and touches nothing) if `destWb.Worksheets.Count <> 1` on entry - the same "caller must pass a fresh `Workbooks.Add` result" contract `modLibraryTransfer.BuildExportSheets` relies on implicitly, made explicit and checkable here since Save's failure path (Task 1) needs a real condition to test against.

- [ ] **Step 1: Write the failing test**

Create `tests/test_harness_build.py`:

```python
from tests.conftest import run


def test_build_harness_sheets_creates_harness_and_snapshot(wb, app):
    dest = app.Workbooks.Add()
    try:
        ok = run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        assert ok is True
        names = [dest.Worksheets(i + 1).Name for i in range(dest.Worksheets.Count)]
        assert names == ["Harness", "_Snapshot"]
    finally:
        dest.Close(SaveChanges=False)


def test_snapshot_sheet_is_very_hidden(wb, app):
    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        assert dest.Worksheets("_Snapshot").Visible == 2  # xlSheetVeryHidden
    finally:
        dest.Close(SaveChanges=False)


def test_build_harness_sheets_rejects_a_non_fresh_workbook(wb, app):
    dest = app.Workbooks.Add()
    try:
        dest.Worksheets.Add()  # now has 2 sheets, no longer "fresh"
        ok = run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        assert ok is False
        assert dest.Worksheets.Count == 2  # untouched
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_harness_build.py -v
```

Expected: FAIL - `modHarnessBuild` does not exist.

- [ ] **Step 3: Write the module**

Create `src/vba/modHarnessBuild.bas`:

```vb
Attribute VB_Name = "modHarnessBuild"
Option Explicit

Public Const SAVED_CHART_HEADER_ROW As Long = 6
Public Const SAVED_CHART_FIRST_ROW As Long = 7
Public Const SAVED_CHART_LAST_ROW As Long = 1006
Public Const SAVED_COL_JOIN_FROM As Long = 12
Public Const SAVED_COL_JOIN_TO As Long = 13

XL_SHEET_VERY_HIDDEN = 2

Public Function BuildHarnessSheets(destWb As Workbook) As Boolean
    If destWb.Worksheets.Count <> 1 Then Exit Function

    Dim original As Worksheet, wsHarness As Worksheet, wsSnapshot As Worksheet
    Set original = destWb.Worksheets(1)

    Set wsHarness = destWb.Worksheets.Add(After:=original)
    wsHarness.Name = "Harness"
    Set wsSnapshot = destWb.Worksheets.Add(After:=wsHarness)
    wsSnapshot.Name = "_Snapshot"

    Dim bPriorAlerts As Boolean
    bPriorAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    original.Delete
    Application.DisplayAlerts = bPriorAlerts

    wsSnapshot.Visible = XL_SHEET_VERY_HIDDEN

    BuildHarnessSheets = True
End Function
```

`XL_SHEET_VERY_HIDDEN` is declared without `Const` deliberately corrected below - VBA requires `Public Const` for a module-level constant; fix in the same step:

```vb
Public Const XL_SHEET_VERY_HIDDEN As Long = 2
```

(Replace the bare `XL_SHEET_VERY_HIDDEN = 2` line above with this declaration, placed with the other `Public Const` declarations at the top of the module.)

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_harness_build.py -v
```

Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modHarnessBuild.bas tests/test_harness_build.py
git commit -m "feat: shape a fresh workbook into Harness and _Snapshot sheets"
```

---

### Task 3: Copy the title block and chart, with join-key formulas

**Files:**
- Modify: `src/vba/modHarnessBuild.bas`
- Test: `tests/test_harness_build.py` (additions)

**Interfaces:**
- Consumes: `SAVED_CHART_HEADER_ROW`/`FIRST_ROW`/`LAST_ROW`, `SAVED_COL_JOIN_FROM`/`TO` from Task 2. `modChart.CHART_HEADER_ROW`/`FIRST_ROW`/`LAST_ROW` (existing, on `ThisWorkbook`).
- Produces:
  - VBA `modHarnessBuild.CopyTitleBlock(wsDestHarness As Worksheet)` - copies the Creator's title text and eight title-block values verbatim, and re-renders the visible layout (title text formatting, label cells, chart header row) since the destination workbook has none of the Creator's build-time formatting.
  - VBA `modHarnessBuild.CopyChartRows(wsDestHarness As Worksheet) As Long` - copies chart rows 7-1006 verbatim from the Creator, writes join-key formulas in columns L/M for every row in that range, hides columns L/M, and returns the count of rows with a non-blank From Conn or To Conn (the "used" row count Task 1's `HARNESS_SAVED` payload reports).

`CopyTitleBlock`'s value-cell addresses (`B2, E2, H2, B3, E3, H3, B4, H4`) are the same eight addresses `build/layout.py`'s `TB_NAMES` uses in the Creator - hardcoded here rather than read via `ThisWorkbook.Names(...)`, because a freshly created `destWb` has no Defined Names of its own and this plan writes plain values into it, not names.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_harness_build.py`:

```python
def _fill_title_block(wb):
    ws = wb.Worksheets("Harness")
    ws.Range("B2").Value = "Test Harness"
    ws.Range("E2").Value = "HN-001"
    ws.Range("H2").Value = "A"
    ws.Range("B3").Value = "A Student"
    ws.Range("E3").Value = "Shop 1"
    ws.Range("H3").Value = "2026-09-01"
    ws.Range("B4").Value = "A test harness"
    ws.Range("H4").Value = "in"


def test_copy_title_block_round_trips_every_field(wb, app):
    _fill_title_block(wb)
    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        run(wb, "modHarnessBuild.CopyTitleBlock", dest.Worksheets("Harness"))
        ws = dest.Worksheets("Harness")
        assert ws.Range("B2").Value == "Test Harness"
        assert ws.Range("E2").Value == "HN-001"
        assert ws.Range("H2").Value == "A"
        assert ws.Range("B3").Value == "A Student"
        assert ws.Range("E3").Value == "Shop 1"
        assert ws.Range("H4").Value == "in"
        assert ws.Range("A1").Value == "WIRE HARNESS TO-FROM CHART"
        assert ws.Cells(6, 1).Value == "From Conn"  # chart header row rendered
    finally:
        dest.Close(SaveChanges=False)


def test_copy_chart_rows_round_trips_values_and_counts_used_rows(wb, app):
    _fill_title_block(wb)
    wsSrc = wb.Worksheets("Harness")
    wsSrc.Cells(7, 1).Value = "J1"
    wsSrc.Cells(7, 2).Value = 1
    wsSrc.Cells(7, 9).Value = "J2"
    wsSrc.Cells(7, 10).Value = 1
    wsSrc.Cells(8, 1).Value = "J1"
    wsSrc.Cells(8, 2).Value = 2
    wsSrc.Cells(8, 9).Value = "J2"
    wsSrc.Cells(8, 10).Value = 2

    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        n = run(wb, "modHarnessBuild.CopyChartRows", dest.Worksheets("Harness"))
        assert n == 2

        ws = dest.Worksheets("Harness")
        assert ws.Cells(7, 1).Value == "J1"
        assert ws.Cells(7, 10).Value == 2
        assert ws.Cells(9, 1).Value is None  # untouched beyond the used rows
    finally:
        dest.Close(SaveChanges=False)


def test_copy_chart_rows_writes_live_join_key_formulas(wb, app):
    _fill_title_block(wb)
    wsSrc = wb.Worksheets("Harness")
    wsSrc.Cells(7, 1).Value = "J1"
    wsSrc.Cells(7, 2).Value = 3
    wsSrc.Cells(7, 9).Value = "J2"
    wsSrc.Cells(7, 10).Value = 4

    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        run(wb, "modHarnessBuild.CopyChartRows", dest.Worksheets("Harness"))
        ws = dest.Worksheets("Harness")
        assert ws.Cells(7, 12).Value == "J1|3"
        assert ws.Cells(7, 13).Value == "J2|4"
        assert ws.Cells(8, 12).Value == ""  # blank row: formula present, resolves empty
        assert ws.Columns(12).Hidden is True
        assert ws.Columns(13).Hidden is True

        # A hand edit to the chart after saving keeps the join key correct -
        # this is what makes 3c's pin tables react to it with no macro.
        ws.Cells(7, 2).Value = 9
        assert ws.Cells(7, 12).Value == "J1|9"
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
python -m pytest tests/test_harness_build.py -v -k "title_block or chart_rows"
```

Expected: FAIL - `CopyTitleBlock`/`CopyChartRows` don't exist.

- [ ] **Step 3: Write the functions**

Append to `src/vba/modHarnessBuild.bas`:

```vb
Private Const TB_VALUE_CELLS As String = "B2,E2,H2,B3,E3,H3,B4,H4"

Public Sub CopyTitleBlock(wsDestHarness As Worksheet)
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(modChart.CHART_SHEET)

    wsDestHarness.Range("A1").Value = "WIRE HARNESS TO-FROM CHART"
    wsDestHarness.Range("A1").Font.Size = 16
    wsDestHarness.Range("A1").Font.Bold = True

    Dim vCells As Variant, i As Long, sCell As String
    vCells = Split(TB_VALUE_CELLS, ",")
    For i = LBound(vCells) To UBound(vCells)
        sCell = CStr(vCells(i))
        wsDestHarness.Range(sCell).Value = wsSrc.Range(sCell).Value
        wsDestHarness.Range(sCell).Interior.Color = 0xF2F2F2
    Next i

    Dim vHeaders As Variant, nUnitsIndex As Long
    vHeaders = Array("From Conn", "From Pin", "From Term", "Signal", "Color", "AWG", _
                      wsSrc.Cells(SAVED_CHART_HEADER_ROW, 7).Value, "To Term", "To Conn", "To Pin", "Notes")
    For i = LBound(vHeaders) To UBound(vHeaders)
        Dim cel As Range
        Set cel = wsDestHarness.Cells(SAVED_CHART_HEADER_ROW, i + 1)
        cel.Value = vHeaders(i)
        cel.Font.Bold = True
        cel.Interior.Color = 0xD9D9D9
    Next i
End Sub

Public Function CopyChartRows(wsDestHarness As Worksheet) As Long
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets(modChart.CHART_SHEET)

    Dim r As Long, c As Long, n As Long
    Dim sFrom As String, sTo As String

    For r = SAVED_CHART_FIRST_ROW To SAVED_CHART_LAST_ROW
        For c = 1 To 11
            wsDestHarness.Cells(r, c).Value = wsSrc.Cells(r, c).Value
        Next c

        sFrom = Trim$(CStr(wsSrc.Cells(r, 1).Value))
        sTo = Trim$(CStr(wsSrc.Cells(r, 9).Value))
        If Len(sFrom) > 0 Or Len(sTo) > 0 Then n = n + 1

        wsDestHarness.Cells(r, SAVED_COL_JOIN_FROM).Formula = _
            "=IF(A" & r & "=""""," """", A" & r & "&""|""&B" & r & ")"
        wsDestHarness.Cells(r, SAVED_COL_JOIN_TO).Formula = _
            "=IF(I" & r & "=""""," """", I" & r & "&""|""&J" & r & ")"
    Next r

    wsDestHarness.Columns(SAVED_COL_JOIN_FROM).Hidden = True
    wsDestHarness.Columns(SAVED_COL_JOIN_TO).Hidden = True

    CopyChartRows = n
End Function
```

`.Formula` (not `.Value`) is what keeps the join key live against a later hand edit to columns A/B or I/J - the round trip in the third test above is the behavior this exists for.

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_harness_build.py -v -k "title_block or chart_rows"
```

Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modHarnessBuild.bas tests/test_harness_build.py
git commit -m "feat: copy the title block and chart into a saved harness, with live join keys"
```

---

### Task 4: Copy the connector snapshot

**Files:**
- Modify: `src/vba/modHarnessBuild.bas`
- Test: `tests/test_harness_build.py` (additions)

**Interfaces:**
- Consumes: `modSnapshot.SNAP_CONN_FIRST_ROW`/`LAST_ROW`, `SNAP_PINS_FIRST_ROW`/`LAST_ROW` (existing), `modLibrary.LIB_ROW_CAP`, `library_layout.CONN_HEADERS`/`PIN_HEADERS` field order (existing, mirrored by `modLibrary`'s `LIB_COL_*`/`PIN_COL_*` constants).
- Produces: VBA `modHarnessBuild.CopySnapshot(wsDestSnapshot As Worksheet)` - copies every used row of the Creator's `_Snapshot` Connectors and Pins blocks, plus every embedded photo Shape, into the destination `_Snapshot` sheet at the same fixed row addresses.

The destination `_Snapshot` sheet reuses the Creator's own row layout (`SNAP_CONN_FIRST_ROW = 2`/`SNAP_CONN_LAST_ROW = 201`, `SNAP_PINS_FIRST_ROW = 211`/`SNAP_PINS_LAST_ROW = 2210`) so that 3e's Load can read it back with the exact same `modLibrary.ReadConnector`/`ReadPinsForConnector` calls it already uses elsewhere, unchanged.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_harness_build.py`:

```python
from tests.fixtures.sample_photo import write_sample_photo


def test_copy_snapshot_round_trips_connectors_pins_and_photos(wb, app, tmp_path):
    wsSnap = wb.Worksheets("_Snapshot")
    fields = (
        "DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
        4, "", "PHOTO_DTM-04P", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )
    run(wb, "modLibrary.WriteConnector", wsSnap, 2, 201, fields)
    run(wb, "modLibrary.WritePin", wsSnap, 211, 2210, ("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1))
    photo_path = write_sample_photo(tmp_path / "photo.png")
    run(wb, "modLibrary.EmbedConnectorPhoto", wsSnap, "DTM-04P", str(photo_path))

    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        run(wb, "modHarnessBuild.CopySnapshot", dest.Worksheets("_Snapshot"))

        wsDest = dest.Worksheets("_Snapshot")
        result = run(wb, "modLibrary.ReadConnector", wsDest, 2, 201, "DTM-04P")
        assert tuple(result) == fields

        pins = run(wb, "modLibrary.ReadPinsForConnector", wsDest, 211, 2210, "DTM-04P")
        assert [int(row[1]) for row in pins] == [1]

        assert wsDest.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_harness_build.py -v -k copy_snapshot
```

Expected: FAIL - `CopySnapshot` does not exist.

- [ ] **Step 3: Write the function**

Append to `src/vba/modHarnessBuild.bas`:

```vb
Public Sub CopySnapshot(wsDestSnapshot As Worksheet)
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets("_Snapshot")

    Dim r As Long, c As Long

    For r = modSnapshot.SNAP_CONN_FIRST_ROW To modSnapshot.SNAP_CONN_LAST_ROW
        For c = 1 To modLibrary.LIB_FIELD_COUNT
            wsDestSnapshot.Cells(r, c).Value = wsSrc.Cells(r, c).Value
        Next c
    Next r

    For r = modSnapshot.SNAP_PINS_FIRST_ROW To modSnapshot.SNAP_PINS_LAST_ROW
        For c = 1 To modLibrary.PIN_FIELD_COUNT
            wsDestSnapshot.Cells(r, c).Value = wsSrc.Cells(r, c).Value
        Next c
    Next r

    Dim shp As Shape
    For Each shp In wsSrc.Shapes
        shp.Copy
        wsDestSnapshot.Paste
        wsDestSnapshot.Shapes(wsDestSnapshot.Shapes.Count).Name = shp.Name
    Next shp
End Sub
```

This reuses the same `Shape.Copy`/`Paste` mechanism `modLibraryTransfer.CopyConnectorPhoto` already uses for the same purpose (copying an embedded photo between sheets) - `docs/superpowers/plans/phase-2-manual-verification.md`'s 2d section already flags this mechanism as the one clipboard-dependent, occasionally-flaky operation in the codebase. Note this explicitly in the manual-verification checklist (Task 7) rather than treating a rare flake here as a regression.

- [ ] **Step 4: Run the test to verify it passes**

Run:

```bash
python -m pytest tests/test_harness_build.py -v -k copy_snapshot
```

Expected: 1 passed. If it fails in a way that looks clipboard-related (a shape silently missing rather than a clear error), re-run once before treating it as a real bug - see the note above.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modHarnessBuild.bas tests/test_harness_build.py
git commit -m "feat: copy the connector snapshot into a saved harness"
```

---

### Task 5: The SaveHarness action

**Files:**
- Create: `src/vba/modHarnessActions.bas`
- Test: `tests/test_harness_actions.py`

**Interfaces:**
- Consumes: `modHarnessBuild.BuildHarnessSheets`/`CopyTitleBlock`/`CopyChartRows`/`CopySnapshot` (Tasks 2-4), `modContract.Success`/`Failure` (Task 1's new codes).
- Produces: VBA `modHarnessActions.SaveHarness(destWb As Workbook) As Variant` - outcomes `HARNESS_SAVED` (payload: `Long`, used-row count) and `HARNESS_SAVE_FAILED` (payload: `String`, reason).

- [ ] **Step 1: Write the failing test**

Create `tests/test_harness_actions.py`:

```python
from tests.conftest import run_action


def test_save_harness_succeeds_against_a_fresh_workbook(wb, app):
    wsSrc = wb.Worksheets("Harness")
    wsSrc.Range("B2").Value = "Test Harness"
    wsSrc.Cells(7, 1).Value = "J1"
    wsSrc.Cells(7, 2).Value = 1
    wsSrc.Cells(7, 9).Value = "J2"
    wsSrc.Cells(7, 10).Value = 1

    dest = app.Workbooks.Add()
    try:
        result = run_action(wb, "modHarnessActions.SaveHarness", dest)
        assert result.ok is True
        assert result.outcome == "HARNESS_SAVED"
        assert result.payload == 1
        assert dest.Worksheets("Harness").Range("B2").Value == "Test Harness"
    finally:
        dest.Close(SaveChanges=False)


def test_save_harness_rejects_a_non_fresh_workbook(wb, app):
    dest = app.Workbooks.Add()
    try:
        dest.Worksheets.Add()
        result = run_action(wb, "modHarnessActions.SaveHarness", dest)
        assert result.ok is False
        assert result.outcome == "HARNESS_SAVE_FAILED"
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_harness_actions.py -v
```

Expected: FAIL - `modHarnessActions` does not exist.

- [ ] **Step 3: Write the module**

Create `src/vba/modHarnessActions.bas`:

```vb
Attribute VB_Name = "modHarnessActions"
Option Explicit

Public Function SaveHarness(destWb As Workbook) As Variant
    If Not modHarnessBuild.BuildHarnessSheets(destWb) Then
        SaveHarness = modContract.Failure("HARNESS_SAVE_FAILED", "destination workbook is not fresh")
        Exit Function
    End If

    Dim wsHarness As Worksheet, wsSnapshot As Worksheet
    Set wsHarness = destWb.Worksheets("Harness")
    Set wsSnapshot = destWb.Worksheets("_Snapshot")

    modHarnessBuild.CopyTitleBlock wsHarness
    Dim nUsedRows As Long
    nUsedRows = modHarnessBuild.CopyChartRows(wsHarness)
    modHarnessBuild.CopySnapshot wsSnapshot

    SaveHarness = modContract.Success("HARNESS_SAVED", nUsedRows)
End Function
```

- [ ] **Step 4: Run the test to verify it passes**

Run:

```bash
python -m pytest tests/test_harness_actions.py -v
```

Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modHarnessActions.bas tests/test_harness_actions.py
git commit -m "feat: add the SaveHarness action"
```

---

### Task 6: Wire the Home buttons and register the new modules

**Files:**
- Create: `src/vba/modHarnessUI.bas`
- Modify: `build/build.py`
- Modify: `build/layout.py`
- Modify: `tests/test_layering.py`
- Modify: `src/vba/modState.bas` (no code change - confirms existing `HarnessPath`/`Dirty` keys are reused, not duplicated)

**Interfaces:**
- Consumes: `modHarnessActions.SaveHarness` (Task 5), `modState.SetState`/`ClearDirty` (existing), `modMessages.MessageFor`/`MessageStyleFor` (existing plus Task 1's additions).
- Produces: `modHarnessUI.SaveHarness()` and `modHarnessUI.SaveHarnessAs()` Public Subs, wired as Home button macros `modHarnessUI.SaveHarness` / `modHarnessUI.SaveHarnessAs`.

`modHarnessUI.bas` is layer 2 but, like `modConnectorUI.bas`, is a plain standard module rather than a form/sheet code-behind - `tests/test_layering.py`'s `ADAPTERS` list only covers form/sheet modules (`modConnectorUI` itself is not in it either), so this plan does not add `modHarnessUI` there; it is still built to the same rules by convention; adding it to that enforcement is out of scope for this plan.

- [ ] **Step 1: Write the module**

Create `src/vba/modHarnessUI.bas`:

```vb
Attribute VB_Name = "modHarnessUI"
Option Explicit

Public Sub SaveHarness()
    Dim sPath As String
    sPath = modState.GetState("HarnessPath")
    If Len(sPath) = 0 Then
        SaveHarnessAs
        Exit Sub
    End If

    SaveToPath sPath
End Sub

Public Sub SaveHarnessAs()
    Dim vPath As Variant
    vPath = Application.GetSaveAsFilename( _
        InitialFileName:=DefaultFileName(), _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If vPath = False Then Exit Sub

    SaveToPath CStr(vPath)
End Sub

Private Sub SaveToPath(ByVal sPath As String)
    Dim destWb As Workbook
    Set destWb = Workbooks.Add

    Dim vResult As Variant
    vResult = modHarnessActions.SaveHarness(destWb)

    If modContract.Ok(vResult) Then
        destWb.SaveAs Filename:=sPath, FileFormat:=51
        modState.SetState "HarnessPath", sPath
        modState.ClearDirty
    End If
    destWb.Close SaveChanges:=False

    MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
End Sub

Private Function DefaultFileName() As String
    Dim sNumber As String
    sNumber = Trim$(CStr(ThisWorkbook.Worksheets(modChart.CHART_SHEET).Range("E2").Value))
    If Len(sNumber) = 0 Then
        DefaultFileName = "harness.xlsx"
    Else
        DefaultFileName = sNumber & ".xlsx"
    End If
End Function
```

- [ ] **Step 2: Wire the Home buttons**

In `build/layout.py`, extend `HOME_BUTTONS`:

```python
HOME_BUTTONS = [
    # (caption, macro, left, top, width, height)
    ("New Harness", "modChart.NewHarness", 20, 220, 120, 32),
    ("Save Harness", "modHarnessUI.SaveHarness", 150, 220, 120, 32),
    ("Save Harness As", "modHarnessUI.SaveHarnessAs", 280, 220, 120, 32),
    ("Add Connector", "modConnectorUI.ShowAddConnector", 20, 260, 120, 32),
    ("Remove Connector", "modConnectorUI.ShowRemoveConnector", 150, 260, 130, 32),
    ("Manage Library", "modConnectorUI.ShowManageLibrary", 20, 300, 120, 32),
]
```

- [ ] **Step 3: Register the new modules in the build**

In `build/build.py`, extend `VBA_MODULES`:

```python
VBA_MODULES = [
    "modUtil.bas", "modState.bas", "modConnectors.bas", "modChart.bas",
    "modLibrary.bas", "modPinEditor.bas", "clsPinMarker.cls", "modSnapshot.bas",
    "modConnectorUI.bas", "modLibraryTransfer.bas", "modContract.bas",
    "modMessages.bas", "modEditorActions.bas", "modPickerActions.bas",
    "modManageActions.bas", "modConnectorActions.bas",
    "modHarnessBuild.bas", "modHarnessActions.bas", "modHarnessUI.bas",
]
```

- [ ] **Step 4: Register the new modules in the layering test**

In `tests/test_layering.py`:

```python
LAYER0 = [
    "modUtil", "modState", "modLibrary", "modChart", "modConnectors",
    "modSnapshot", "modLibraryTransfer", "modPinEditor", "modHarnessBuild",
]
LAYER1 = [
    "modContract", "modMessages", "modEditorActions", "modPickerActions",
    "modManageActions", "modConnectorActions", "modHarnessActions",
]
```

`test_action_modules_open_no_dialogs` (parametrized over `LAYER1`) now covers `modHarnessActions` for free - it will fail immediately if a later sub-plan accidentally moves a dialog or `Workbooks.Open`/`.Add` into it instead of `modHarnessUI`.

- [ ] **Step 5: Run the full suite**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: everything passes from a clean build, including the new `test_action_modules_open_no_dialogs[modHarnessActions]` case.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modHarnessUI.bas build/build.py build/layout.py tests/test_layering.py
git commit -m "feat: wire Save Harness and Save Harness As to the Home sheet"
```

---

### Task 7: Full round-trip integration test

**Files:**
- Test: `tests/test_harness_save_integration.py`

**Interfaces:**
- Consumes: every function from Tasks 2-6.
- Produces: nothing later tasks depend on - a capstone test proving the pieces compose against the real built artifact, matching the precedent set by `tests/test_library_integration.py` in Phase 2a.

- [ ] **Step 1: Write the test**

Create `tests/test_harness_save_integration.py`:

```python
from tests.conftest import run, run_action
from tests.fixtures.sample_photo import write_sample_photo


def test_full_harness_round_trips_through_a_saved_file(wb, app, tmp_path):
    wsHarness = wb.Worksheets("Harness")
    wsHarness.Range("B2").Value = "Dash Harness"
    wsHarness.Range("E2").Value = "HN-100"
    wsHarness.Range("H2").Value = "A"
    wsHarness.Cells(7, 1).Value = "J1"
    wsHarness.Cells(7, 2).Value = 1
    wsHarness.Cells(7, 4).Value = "12V_SW"
    wsHarness.Cells(7, 5).Value = "Red"
    wsHarness.Cells(7, 6).Value = "18"
    wsHarness.Cells(7, 7).Value = 24
    wsHarness.Cells(7, 9).Value = "J2"
    wsHarness.Cells(7, 10).Value = 1

    wsSnap = wb.Worksheets("_Snapshot")
    fields = (
        "DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
        4, "", "PHOTO_DTM-04P", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
    )
    run(wb, "modLibrary.WriteConnector", wsSnap, 2, 201, fields)
    run(wb, "modLibrary.WritePin", wsSnap, 211, 2210, ("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1))
    photo_path = write_sample_photo(tmp_path / "photo.png")
    run(wb, "modLibrary.EmbedConnectorPhoto", wsSnap, "DTM-04P", str(photo_path))

    dest = app.Workbooks.Add()
    saved_path = tmp_path / "HN-100.xlsx"
    try:
        result = run_action(wb, "modHarnessActions.SaveHarness", dest)
        assert result.ok is True
        dest.SaveAs(Filename=str(saved_path), FileFormat=51)
    finally:
        dest.Close(SaveChanges=False)

    # Reopen as a wholly separate file handle.
    reopened = app.Workbooks.Open(str(saved_path))
    try:
        ws = reopened.Worksheets("Harness")
        assert ws.Range("B2").Value == "Dash Harness"
        assert ws.Cells(7, 1).Value == "J1"
        assert ws.Cells(7, 4).Value == "12V_SW"
        assert ws.Cells(7, 12).Value == "J1|1"
        assert ws.Cells(7, 13).Value == "J2|1"

        wsSnapDest = reopened.Worksheets("_Snapshot")
        assert wsSnapDest.Visible == 2
        result2 = run(wb, "modLibrary.ReadConnector", wsSnapDest, 2, 201, "DTM-04P")
        assert tuple(result2) == fields
        assert wsSnapDest.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"

        assert reopened.VBProject.VBComponents.Count == 0  # macro-free
    finally:
        reopened.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test**

Run:

```bash
python -m pytest tests/test_harness_save_integration.py -v
```

Expected: 1 passed. If it fails, the bug is in how the individually-tested pieces compose (or a flaky clipboard-driven photo copy - re-run once first per Task 4's note), not in any single function.

- [ ] **Step 3: Run the whole suite from a clean build**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: everything passes.

- [ ] **Step 4: Commit**

```bash
git add tests/test_harness_save_integration.py
git commit -m "test: prove a harness round-trips through a saved file"
```

---

## Self-Review

**Spec coverage for this sub-plan.** Covers the "Save" half of the spec's "Save, load, and export" section for the `Harness` and `_Snapshot` pieces of "Harness rendering" - explicitly not the `CONN_<RefDes>` pages (3b), pin-table formulas beyond the join keys themselves (3c), or page setup (3d), all named in the spec's own Phasing list as Phase 3 but split across these five sub-plans per this plan's header.

**Why `.Formula`, not `.Value`, for the join-key columns.** The spec says a saved harness's chart stays hand-editable and that edits "survive a round trip" - Task 3's third test asserts editing `B7` after save immediately changes `L7`'s resolved value, which only a live formula gives. A static text join key would silently go stale the moment a student corrected a pin number by hand, breaking 3c's pin tables without any visible error.

**Why the saved chart keeps the Creator's full 1000-row range instead of trimming to content.** Decided explicitly in Global Constraints, not an oversight: trimming would either block a student from hand-adding a wire past the saved boundary, or require 3c's formulas to somehow extend themselves at print time - reusing the fixed range sidesteps both, at the cost of 3d needing to compute the printed area from content rather than the sheet's dimensions (which it needs to do regardless, since Excel's used range would otherwise include all 1000 rows).

**Why no connector-instance (Ref Des -> ConnectorID) table is saved in this plan.** Traced against the spec's own list of what a saved harness contains (`Harness`, `CONN_<RefDes>` pages, `_Snapshot`) - there is no fourth sheet for it. This plan defers reconstructing that mapping to 3e, sourced from 3b's `CONN_<RefDes>` sheet names and a metadata cell 3b adds - flagged explicitly so 3b's plan does not lose track of that obligation.

**Type consistency.** `SaveHarness`'s `HARNESS_SAVED` payload is a `Long` (used-row count) everywhere it is produced (Task 5) and consumed (Task 1's message test, Task 7's integration test) - never a path string, which Task 1's `HARNESS_SAVE_FAILED` (a `String` reason) is instead. `CopyChartRows`'s return type (`Long`) matches what `SaveHarness` forwards as the payload directly, with no intermediate reinterpretation.

**No placeholders.** Every step contains complete VBA/Python, including the join-key formula strings and the exact cell addresses copied - no "add the remaining fields" or "similar to Task N" stand-ins.
