# Phase 3d: Page Setup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Bake print area, orientation, fit-to-width scaling, and a footer (harness number, revision, page numbering) into every sheet of a saved harness at Save time, plus repeating header-row print titles on the `Harness` sheet only - so `Ctrl+P` on the saved, macro-free file produces a correct PDF without the Creator's involvement.

**Architecture:** A new layer 0 module, `modPageSetup.bas`, applies `Worksheet.PageSetup` properties given a worksheet and the two title-block values (harness number, revision) every footer needs. It computes the printed row range itself rather than trusting the sheet's full fixed allotment (1006 chart rows, or a connector page's minimum print rows) - printing 1000 blank rows would violate "print-ready" on its own. Two entry points exist because the two sheet kinds differ in exactly the ways confirmed during this plan's discussion: the `Harness` sheet is landscape with repeating title rows; `CONN_<RefDes>` pages are portrait with no repeating titles (there is nothing above their own row 1 to repeat).

**Tech Stack:** VBA (Excel 16.0 COM automation), Python 3.13/pywin32/pytest.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Depends on:** 3a, 3b, and 3c, implemented and merged - this plan applies page setup to sheets those three plans already build and populate.

**Part of Phase 3** (see 3a's header for the full breakdown). This is 3d.

## Global Constraints

- **Decision (confirmed during this plan's discussion): differentiated page setup.** `Harness`: landscape, `PrintTitleRows` repeating the chart header row (`modHarnessBuild.SAVED_CHART_HEADER_ROW`). `CONN_<RefDes>` pages: portrait, no `PrintTitleRows`. Both: print area computed from actual content, `FitToPagesWide = 1`, and a footer with harness number, revision, and `Page &P of &N`.
- "Page breaks" (spec wording) are not set manually anywhere in this plan - assigning `PrintArea` and `FitToPagesWide` causes Excel to compute automatic page breaks for whatever content overflows one printed page, which is what the spec's "baked in at save time" list is describing; no `.HPageBreaks.Add`/`.VPageBreaks.Add` call is needed or made.
- Every VBA module starts with `Option Explicit`. `modPageSetup.bas` is layer 0: it may reference `Worksheet`/`Range`/scalars and other layer 0 modules, never dialogs or workbook lifecycle calls.
- Harness number and revision are read from the same two title-block cells 3a's `CopyTitleBlock` already writes (`E2`, `H2`) - this plan does not re-derive them from anywhere else.

## File Structure

| File | Responsibility |
|---|---|
| `src/vba/modPageSetup.bas` | Layer 0. `ApplyHarnessPageSetup`, `ApplyConnectorPageSetup`, the shared footer and used-row computations. |
| `src/vba/modHarnessActions.bas` | Modified: `SaveHarness` calls `ApplyHarnessPageSetup` after `CopyChartRows`. |
| `src/vba/modHarnessBuild.bas` | Modified: `BuildConnectorPages` calls `ApplyConnectorPageSetup` for each page. |
| `build/build.py` | Modified: adds `modPageSetup.bas` to `VBA_MODULES`. |
| `tests/test_layering.py` | Modified: adds `modPageSetup` to `LAYER0`. |
| `tests/test_page_setup.py` | `modPageSetup`'s functions. |
| `tests/test_harness_save_integration.py` | Modified: asserts page setup properties on the saved, reopened file. |

---

### Task 1: Harness sheet page setup

**Files:**
- Create: `src/vba/modPageSetup.bas`
- Test: `tests/test_page_setup.py`

**Interfaces:**
- Consumes: `modHarnessBuild.SAVED_CHART_HEADER_ROW`/`FIRST_ROW`/`LAST_ROW` (3a).
- Produces:
  - VBA `modPageSetup.LastUsedChartRow(wsHarness As Worksheet) As Long` - the last row (7-1006) with any non-blank cell in columns A-K, or the header row (6) if the chart is entirely empty.
  - VBA `modPageSetup.ApplyHarnessPageSetup(wsHarness As Worksheet, ByVal sHarnessNumber As String, ByVal sRevision As String)`.

- [ ] **Step 1: Write the failing test**

Create `tests/test_page_setup.py`:

```python
from tests.conftest import run

XL_LANDSCAPE = 2
XL_PORTRAIT = 1


def test_last_used_chart_row_finds_the_last_populated_row(wb, app):
    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        ws = dest.Worksheets("Harness")
        ws.Cells(7, 1).Value = "J1"
        ws.Cells(10, 9).Value = "J2"  # gap between row 7 and row 10 is realistic

        n = run(wb, "modPageSetup.LastUsedChartRow", ws)
        assert n == 10
    finally:
        dest.Close(SaveChanges=False)


def test_last_used_chart_row_falls_back_to_the_header_row_when_empty(wb, app):
    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        ws = dest.Worksheets("Harness")
        n = run(wb, "modPageSetup.LastUsedChartRow", ws)
        assert n == 6
    finally:
        dest.Close(SaveChanges=False)


def test_apply_harness_page_setup(wb, app):
    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        ws = dest.Worksheets("Harness")
        ws.Cells(7, 1).Value = "J1"

        run(wb, "modPageSetup.ApplyHarnessPageSetup", ws, "HN-100", "A")

        ps = ws.PageSetup
        assert ps.Orientation == XL_LANDSCAPE
        assert ps.FitToPagesWide == 1
        assert ps.PrintArea == "$A$1:$K$7"
        assert ps.PrintTitleRows == "$6:$6"
        assert "HN-100" in ps.CenterFooter
        assert "A" in ps.CenterFooter
        assert "&P" in ps.CenterFooter and "&N" in ps.CenterFooter
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_page_setup.py -v
```

Expected: FAIL - `modPageSetup` does not exist.

- [ ] **Step 3: Write the module**

Create `src/vba/modPageSetup.bas`:

```vb
Attribute VB_Name = "modPageSetup"
Option Explicit

Public Function LastUsedChartRow(wsHarness As Worksheet) As Long
    Dim r As Long

    For r = modHarnessBuild.SAVED_CHART_LAST_ROW To modHarnessBuild.SAVED_CHART_FIRST_ROW Step -1
        If Application.WorksheetFunction.CountA( _
                wsHarness.Range(wsHarness.Cells(r, 1), wsHarness.Cells(r, 11))) > 0 Then
            LastUsedChartRow = r
            Exit Function
        End If
    Next r

    LastUsedChartRow = modHarnessBuild.SAVED_CHART_HEADER_ROW
End Function

Private Function FooterText(ByVal sHarnessNumber As String, ByVal sRevision As String) As String
    FooterText = Trim$(sHarnessNumber & " Rev " & sRevision) & " - Page &P of &N"
End Function

Public Sub ApplyHarnessPageSetup(wsHarness As Worksheet, ByVal sHarnessNumber As String, ByVal sRevision As String)
    Dim nLastRow As Long
    nLastRow = LastUsedChartRow(wsHarness)

    With wsHarness.PageSetup
        .PrintArea = "$A$1:$K$" & CStr(nLastRow)
        .PrintTitleRows = "$" & modHarnessBuild.SAVED_CHART_HEADER_ROW & ":$" & modHarnessBuild.SAVED_CHART_HEADER_ROW
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .CenterFooter = FooterText(sHarnessNumber, sRevision)
    End With
End Sub
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_page_setup.py -v
```

Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modPageSetup.bas tests/test_page_setup.py
git commit -m "feat: bake page setup into the saved harness sheet"
```

---

### Task 2: Connector page setup

**Files:**
- Modify: `src/vba/modPageSetup.bas`
- Test: `tests/test_page_setup.py` (additions)

**Interfaces:**
- Consumes: `modConnectorPage.CONN_TABLE_FIRST_COL`/`HEADER_ROW` (3b).
- Produces:
  - VBA constant `CONN_PAGE_MIN_PRINT_ROW = 30` - a floor on the printed row range so a page with only one or two pins (a short table) still prints its whole photo rather than clipping it; the photo's fixed maximum height (`CONN_PHOTO_MAX_HEIGHT = 300` points, from 3b) comfortably fits within 30 default-height rows.
  - VBA `modPageSetup.ApplyConnectorPageSetup(wsPage As Worksheet, ByVal sHarnessNumber As String, ByVal sRevision As String)`.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_page_setup.py`:

```python
def test_apply_connector_page_setup(wb, app):
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        pins = (("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1),)
        run(wb, "modConnectorPage.WriteTableSkeleton", ws, pins)

        run(wb, "modPageSetup.ApplyConnectorPageSetup", ws, "HN-100", "A")

        ps = ws.PageSetup
        assert ps.Orientation == XL_PORTRAIT
        assert ps.FitToPagesWide == 1
        assert ps.FitToPagesTall == 1
        assert ps.PrintArea == "$A$1:$Q$30"
        assert ps.PrintTitleRows == ""
        assert "HN-100" in ps.CenterFooter
    finally:
        dest.Close(SaveChanges=False)


def test_apply_connector_page_setup_grows_past_the_minimum_for_a_large_pin_count(wb, app):
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        pins = tuple(("DTM-64P", n, "", 0.1, 0.1, 0.1, 0.1) for n in range(1, 41))
        run(wb, "modConnectorPage.WriteTableSkeleton", ws, pins)

        run(wb, "modPageSetup.ApplyConnectorPageSetup", ws, "HN-100", "A")

        assert ws.PageSetup.PrintArea == "$A$1:$Q$41"  # header + 40 pins
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
python -m pytest tests/test_page_setup.py -v -k connector_page_setup
```

Expected: FAIL - `ApplyConnectorPageSetup` does not exist.

- [ ] **Step 3: Write the function**

Append to `src/vba/modPageSetup.bas`:

```vb
Public Const CONN_PAGE_MIN_PRINT_ROW As Long = 30

Private Function LastUsedTableRow(wsPage As Worksheet) As Long
    Dim nTableLast As Long
    nTableLast = wsPage.Cells(wsPage.Rows.Count, modConnectorPage.CONN_TABLE_FIRST_COL).End(xlUp).Row
    If nTableLast < CONN_PAGE_MIN_PRINT_ROW Then nTableLast = CONN_PAGE_MIN_PRINT_ROW
    LastUsedTableRow = nTableLast
End Function

Public Sub ApplyConnectorPageSetup(wsPage As Worksheet, ByVal sHarnessNumber As String, ByVal sRevision As String)
    With wsPage.PageSetup
        .PrintArea = "$A$1:$Q$" & CStr(LastUsedTableRow(wsPage))
        .PrintTitleRows = ""
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Zoom = False
        .CenterFooter = FooterText(sHarnessNumber, sRevision)
    End With
End Sub
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_page_setup.py -v -k connector_page_setup
```

Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modPageSetup.bas tests/test_page_setup.py
git commit -m "feat: bake page setup into every connector page"
```

---

### Task 3: Wire page setup into Save

**Files:**
- Modify: `src/vba/modHarnessActions.bas`
- Modify: `src/vba/modHarnessBuild.bas`
- Modify: `build/build.py`
- Modify: `tests/test_layering.py`

**Interfaces:**
- Consumes: `ApplyHarnessPageSetup`/`ApplyConnectorPageSetup` (Tasks 1-2).
- Produces: every sheet in a saved harness gets page setup applied as part of the same `SaveHarness` transaction.

- [ ] **Step 1: Apply harness page setup in SaveHarness**

In `src/vba/modHarnessActions.bas`, `SaveHarness` needs the harness number and revision, already sitting in `wsHarness` by the time page setup runs (`CopyTitleBlock` wrote them). Add after the existing `CopyChartRows`/`BuildConnectorPages` calls:

```vb
    Dim sHarnessNumber As String, sRevision As String
    sHarnessNumber = CStr(wsHarness.Range("E2").Value)
    sRevision = CStr(wsHarness.Range("H2").Value)

    modPageSetup.ApplyHarnessPageSetup wsHarness, sHarnessNumber, sRevision
```

(Placed after `modHarnessBuild.BuildConnectorPages destWb, wsSnapshot` and before the final `SaveHarness = modContract.Success(...)` line.)

- [ ] **Step 2: Apply connector page setup per page**

In `src/vba/modHarnessBuild.bas`'s `BuildConnectorPages`, add the harness number/revision as parameters and call page setup for each page:

```vb
Public Sub BuildConnectorPages(destWb As Workbook, wsSnapshot As Worksheet, _
                               ByVal sHarnessNumber As String, ByVal sRevision As String)
```

(Signature change - add the two new parameters at the end.) Add one line inside the existing `For i = ...` loop, after `modConnectorPage.WriteMetadata wsPage, sConnectorID`:

```vb
        modConnectorPage.WriteMetadata wsPage, sConnectorID
        modPageSetup.ApplyConnectorPageSetup wsPage, sHarnessNumber, sRevision
```

Update `modHarnessActions.SaveHarness`'s call site to pass the two values (computed once, before `BuildConnectorPages` is called - reorder so the harness number/revision are read immediately after `CopyTitleBlock`, ahead of both `BuildConnectorPages` and `ApplyHarnessPageSetup`):

```vb
    modHarnessBuild.CopyTitleBlock wsHarness
    Dim sHarnessNumber As String, sRevision As String
    sHarnessNumber = CStr(wsHarness.Range("E2").Value)
    sRevision = CStr(wsHarness.Range("H2").Value)

    Dim nUsedRows As Long
    nUsedRows = modHarnessBuild.CopyChartRows(wsHarness)
    modHarnessBuild.CopySnapshot wsSnapshot
    modHarnessBuild.BuildConnectorPages destWb, wsSnapshot, sHarnessNumber, sRevision
    modPageSetup.ApplyHarnessPageSetup wsHarness, sHarnessNumber, sRevision
```

(This replaces the corresponding lines added by 3a/3b - the net effect is that `sHarnessNumber`/`sRevision` are read once and reused for every page's footer.)

- [ ] **Step 3: Update the existing connector-page tests' call sites**

`tests/test_harness_build.py`'s `test_build_connector_pages_creates_one_sheet_per_instance` (3b, Task 6) calls `modHarnessBuild.BuildConnectorPages` directly with two arguments; update it to pass the two new ones:

```python
        run(wb, "modHarnessBuild.BuildConnectorPages", dest, dest.Worksheets("_Snapshot"), "HN-100", "A")
```

- [ ] **Step 4: Register the new module**

In `build/build.py`, add `"modPageSetup.bas"` to `VBA_MODULES` (after `"modConnectorPage.bas"`).

In `tests/test_layering.py`, add `"modPageSetup"` to `LAYER0`.

- [ ] **Step 5: Run the full suite**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: everything passes.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modHarnessActions.bas src/vba/modHarnessBuild.bas build/build.py tests/test_layering.py tests/test_harness_build.py
git commit -m "feat: apply page setup to every sheet during Save"
```

---

### Task 4: Extend the round-trip integration test

**Files:**
- Modify: `tests/test_harness_save_integration.py`

**Interfaces:**
- Consumes: everything from Tasks 1-3.

- [ ] **Step 1: Add page setup assertions**

Append inside the `reopened` block of `test_full_harness_round_trips_through_a_saved_file`:

```python
        harness_ps = ws.PageSetup
        assert harness_ps.Orientation == 2  # xlLandscape
        assert harness_ps.PrintTitleRows == "$6:$6"
        assert "HN-100" in harness_ps.CenterFooter

        page_ps = j1.PageSetup
        assert page_ps.Orientation == 1  # xlPortrait
        assert page_ps.PrintTitleRows == ""
        assert "HN-100" in page_ps.CenterFooter
```

(`j1` already exists in this test from 3b/3c's extensions, referencing `reopened.Worksheets("CONN_J1")`.)

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

Expected: everything passes.

- [ ] **Step 4: Commit**

```bash
git add tests/test_harness_save_integration.py
git commit -m "test: assert page setup on the saved harness round trip"
```

---

## Self-Review

**Spec coverage for this sub-plan.** "Print area, page breaks, print titles repeating the chart header row, fit-to-width scaling, landscape orientation for the chart, and a footer carrying harness number, revision, and page numbering" - print area, fit-to-width, and the footer are applied to every sheet; print titles and landscape are applied to `Harness` only, per this plan's confirmed differentiated-page-setup decision; page breaks are Excel's automatic consequence of the other two, not a separate call, as stated in Global Constraints.

**Why `LastUsedChartRow` scans for content instead of trusting `CopyChartRows`'s row count.** 3a's `CopyChartRows` returns a *count* of used rows, not the *last row index* - a chart with a gap (row 7 filled, rows 8-9 blank, row 10 filled, matching a student who deleted a middle wire by hand without compacting the sheet) has `nUsedRows = 2` but a last-row index of 10. Printing only through row 8 would silently drop row 10's wire from the PDF. The first test in Task 1 asserts exactly this gapped case.

**Why `CONN_PAGE_MIN_PRINT_ROW` exists and why it is a floor, not a fixed value.** A connector with one or two pins would otherwise get a print area only a few rows tall - short enough to clip the photo, which is anchored well below where a two-row table ends (`CONN_PHOTO_TOP + CONN_PHOTO_MAX_HEIGHT` in points, from 3b, comfortably corresponds to about 24 default-height rows; 30 leaves margin). Task 2's second test proves the floor is a minimum, not a cap - a 40-pin connector's print area grows past it rather than clipping the table instead.

**Type consistency.** `ApplyHarnessPageSetup` and `ApplyConnectorPageSetup` both take `(worksheet, sHarnessNumber As String, sRevision As String)` in the same order and both delegate to the same private `FooterText` - a footer's wording cannot drift between the two sheet kinds since there is exactly one function that renders it.

**No placeholders.** Every step contains the exact `PageSetup` properties set and their exact values or formulas (`PrintArea`, `PrintTitleRows`, `CenterFooter`'s literal text) - no "configure the appropriate print settings" left unspecified.
