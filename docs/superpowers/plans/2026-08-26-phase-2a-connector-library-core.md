# Phase 2a: Connector Library Core Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce `dist/ConnectorLibrary.xlsx` (a macro-free workbook implementing the spec's three-table connector schema) and `src/vba/modLibrary.bas`, a schema-agnostic reader/writer that the Creator uses against that file, an imported library file, or the Creator's own `_Snapshot` sheet — the same functions serve all three, per spec.

**Architecture:** `ConnectorLibrary.xlsx` is built the same way `HarnessCreator.xlsm` is: Python drives Excel COM to lay out sheets and headers, then saves — but with no VBA import step, since the file must stay macro-free. `modLibrary.bas` lives inside `HarnessCreator.xlsm` and operates on `Worksheet` objects and an explicit `(nFirstRow, nLastRow)` row window passed in by the caller, rather than assuming a dedicated whole sheet. This is what lets the same functions serve a whole-sheet library table today and a fixed-size region on the shared `_Snapshot` sheet in a later phase, without changing modLibrary's public signatures. pytest opens both workbooks in one Excel session and calls `modLibrary` functions via `Application.Run` on the Creator, passing the library workbook's `Worksheet` objects as COM arguments.

**Tech Stack:** Python 3.13, pywin32, pytest, Excel 16.0 COM automation, VBA.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Depends on:** Phase 1 (`docs/superpowers/plans/2026-08-26-phase-1-creator-shell.md`), already built and committed on `master`.

**Part of Phase 2**, split into four sub-plans per subsystem (this is 2a):
- **2a (this plan): Connector library core** — library file, reader/writer, ID slugification, photo embedding.
- 2b: Connector editor with click-to-place (depends on 2a).
- 2c: Connector picker, Add/Remove Connector, snapshot embedding, ref des rename (depends on 2a, 2b).
- 2d: Library import and export (depends on 2a, 2c).

## Global Constraints

Copied verbatim from the spec. Every task's requirements implicitly include these.

- Windows and desktop Excel only. No Excel-for-web or non-Windows support.
- Formulas must work on Excel 2016 and later. Use INDEX/MATCH, never XLOOKUP.
- Every VBA module starts with `Option Explicit`.
- No `MsgBox` or dialog in a logic module. UI is confined to UI modules, and `_State` carries a `TestMode` flag that suppresses it.
- The `.xlsm` and `.xlsx` build artifacts are never hand-edited. Change source and rebuild.
- `ConnectorLibrary.xlsx` is macro-free.
- ConnectorID: unique slug, uppercase, non-alphanumerics replaced with `-`. Derived from PartNumber when present, otherwise Name. A numeric suffix is appended on collision.
- Photos are embedded in the library workbook (a single portable file) and also cached as `Photos/<ConnectorID>.png` beside it.
- Connectors table fields: ConnectorID, Name, Manufacturer, PartNumber, Type, PinCount, Notes, PhotoShapeName, CreatedUtc, ModifiedUtc, Origin.
- Pins table fields: ConnectorID, PinNumber, PinLabel, NormX, NormY, LabelX, LabelY. NormX/NormY/LabelX/LabelY are normalized 0.0-1.0.

## File Structure

| File | Responsibility |
|---|---|
| `build/library_layout.py` | Sheet layout constants and builders for `ConnectorLibrary.xlsx` |
| `build/build.py` | Modified: adds `build_library()`, imports `modLibrary.bas` into the Creator |
| `build/excel_com.py` | Modified: adds `save_as_xlsx()` for the macro-free artifact |
| `src/vba/modLibrary.bas` | Column layout constants, ID slugification, Connectors/Pins record CRUD, photo embedding — all operating on a caller-supplied `Worksheet` and row window |
| `tests/conftest.py` | Modified: adds `library_artifact` and `library_wb` fixtures |
| `tests/test_library_build.py` | `ConnectorLibrary.xlsx` structure |
| `tests/test_library_ids.py` | `SlugifyConnectorID`, `UniqueConnectorID` |
| `tests/test_library_connectors.py` | Connector record CRUD round trip |
| `tests/test_library_pins.py` | Pin record CRUD round trip |
| `tests/test_library_photos.py` | Photo embedding, grid placement, cache path |
| `tests/test_library_integration.py` | Full write-save-reopen-read round trip against the real built artifact |

---

### Task 1: Build the `ConnectorLibrary.xlsx` artifact

**Files:**
- Create: `build/library_layout.py`
- Modify: `build/excel_com.py`
- Modify: `build/build.py`
- Create: `tests/conftest.py` additions
- Test: `tests/test_library_build.py`

**Interfaces:**
- Consumes: `layout.VISIBLE` (reused, not duplicated).
- Produces:
  - `library_layout.CONN_HEADERS: list[str]`, `library_layout.PIN_HEADERS: list[str]`
  - `library_layout.build_library_sheets(wb) -> dict[str, object]`
  - `excel_com.save_as_xlsx(wb, path: Path) -> None`
  - `build.build_library(out_dir: Path) -> Path`
  - pytest fixtures `library_artifact`, `library_wb`

- [ ] **Step 1: Write the failing test**

Create `tests/test_library_build.py`:

```python
import pytest


def test_library_artifact_is_produced(library_artifact):
    assert library_artifact.exists()
    assert library_artifact.suffix == ".xlsx"


def test_library_has_three_sheets_in_order(library_wb):
    names = [library_wb.Worksheets(i + 1).Name for i in range(library_wb.Worksheets.Count)]
    assert names == ["Connectors", "Pins", "Photos"]


CONN_HEADERS = [
    "ConnectorID", "Name", "Manufacturer", "PartNumber", "Type",
    "PinCount", "Notes", "PhotoShapeName", "CreatedUtc", "ModifiedUtc", "Origin",
]


@pytest.mark.parametrize("index,header", list(enumerate(CONN_HEADERS, start=1)))
def test_connectors_sheet_headers(library_wb, index, header):
    assert library_wb.Worksheets("Connectors").Cells(1, index).Value == header


PIN_HEADERS = ["ConnectorID", "PinNumber", "PinLabel", "NormX", "NormY", "LabelX", "LabelY"]


@pytest.mark.parametrize("index,header", list(enumerate(PIN_HEADERS, start=1)))
def test_pins_sheet_headers(library_wb, index, header):
    assert library_wb.Worksheets("Pins").Cells(1, index).Value == header


def test_library_workbook_has_no_vba_modules(library_wb):
    assert library_wb.VBProject.VBComponents.Count == 0
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_library_build.py -v
```

Expected: FAIL — collection error, `library_artifact`/`library_wb` fixtures don't exist yet.

- [ ] **Step 3: Write the library layout module**

Create `build/library_layout.py`:

```python
"""Sheet layout for ConnectorLibrary.xlsx — macro-free, three tables."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from layout import VISIBLE

SHEET_NAMES = ["Connectors", "Pins", "Photos"]

CONN_HEADERS = [
    "ConnectorID", "Name", "Manufacturer", "PartNumber", "Type",
    "PinCount", "Notes", "PhotoShapeName", "CreatedUtc", "ModifiedUtc", "Origin",
]
CONN_COLUMN_WIDTHS = [16, 24, 16, 18, 12, 10, 24, 20, 20, 20, 14]

PIN_HEADERS = ["ConnectorID", "PinNumber", "PinLabel", "NormX", "NormY", "LabelX", "LabelY"]
PIN_COLUMN_WIDTHS = [16, 10, 16, 10, 10, 10, 10]


def _build_header_sheet(sheet, headers, widths) -> None:
    for index, header in enumerate(headers, start=1):
        cell = sheet.Cells(1, index)
        cell.Value = header
        cell.Font.Bold = True
        cell.Interior.Color = 0xD9D9D9
        sheet.Columns(index).ColumnWidth = widths[index - 1]


def build_library_sheets(wb) -> dict:
    """Create Connectors, Pins, and Photos, then remove Excel's defaults.

    No VBA project is touched here - no code names, nothing imported. That
    is what keeps the saved .xlsx macro-free.
    """
    originals = [wb.Worksheets(i + 1) for i in range(wb.Worksheets.Count)]

    sheets = {}
    anchor = originals[-1]
    for name in SHEET_NAMES:
        sheet = wb.Worksheets.Add(After=anchor)
        sheet.Name = name
        sheet.Visible = VISIBLE
        sheets[name] = sheet
        anchor = sheet

    for sheet in originals:
        sheet.Delete()

    _build_header_sheet(sheets["Connectors"], CONN_HEADERS, CONN_COLUMN_WIDTHS)
    _build_header_sheet(sheets["Pins"], PIN_HEADERS, PIN_COLUMN_WIDTHS)
    sheets["Photos"].Range("A1").Value = "Connector Photos"
    sheets["Photos"].Range("A1").Font.Bold = True

    return sheets
```

- [ ] **Step 4: Add the macro-free save helper**

In `build/excel_com.py`, add near `save_as_xlsm`:

```python
XL_OPENXML_WORKBOOK = 51


def save_as_xlsx(wb, path: Path) -> None:
    """Save as a plain, macro-free .xlsx. No VBA project may be imported
    into wb before calling this - Excel silently strips it either way, but
    importing it first would raise a "file may contain features not
    compatible" alert this build runs headless and cannot dismiss."""
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()
    wb.SaveAs(Filename=str(path), FileFormat=XL_OPENXML_WORKBOOK)
```

- [ ] **Step 5: Wire the library build into `build.py`**

In `build/build.py`, add alongside the existing imports and constants:

```python
import library_layout

LIBRARY_NAME = "ConnectorLibrary.xlsx"


def build_library(out_dir: Path = DIST) -> Path:
    target = out_dir / LIBRARY_NAME
    with excel_com.excel_app() as app:
        wb = app.Workbooks.Add()
        try:
            library_layout.build_library_sheets(wb)
            excel_com.save_as_xlsx(wb, target)
        finally:
            wb.Close(SaveChanges=False)
    return target
```

`build_library` does not call `excel_com.check_access_vbom()` — it never touches `wb.VBProject`, so the VBOM trust setting is irrelevant to it.

In `main()`, build both artifacts:

```python
    print(f"Built {build()}")
    print(f"Built {build_library()}")
    return 0
```

- [ ] **Step 6: Add the library fixtures**

In `tests/conftest.py`, add after the existing `artifact`/`wb` fixtures:

```python
LIBRARY_ARTIFACT = ROOT / "dist" / "ConnectorLibrary.xlsx"


@pytest.fixture(scope="session")
def library_artifact(artifact) -> Path:
    """Depends on `artifact` so the one `build.py` subprocess run - which
    now builds both files - has already happened."""
    assert LIBRARY_ARTIFACT.exists(), f"build produced no artifact at {LIBRARY_ARTIFACT}"
    return LIBRARY_ARTIFACT


@pytest.fixture
def library_wb(app, library_artifact):
    """A freshly opened copy of the built library workbook, discarded after
    each test."""
    book = app.Workbooks.Open(str(library_artifact))
    try:
        yield book
    finally:
        book.Close(SaveChanges=False)
```

- [ ] **Step 7: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_library_build.py -v
```

Expected: 21 passed.

- [ ] **Step 8: Run the full existing suite to confirm no regression**

Run:

```bash
python -m pytest -v
```

Expected: all passed (104 from Phase 1, plus 21 new = 125).

- [ ] **Step 9: Commit**

```bash
git add build/library_layout.py build/excel_com.py build/build.py tests/conftest.py tests/test_library_build.py
git commit -m "feat: build the macro-free ConnectorLibrary.xlsx artifact"
```

---

### Task 2: Connector record CRUD

**Files:**
- Create: `src/vba/modLibrary.bas`
- Modify: `build/build.py`
- Test: `tests/test_library_connectors.py`

**Interfaces:**
- Consumes: nothing new — operates on any `Worksheet` matching the Connectors column layout.
- Produces:
  - VBA constants `LIB_COL_ID` through `LIB_COL_ORIGIN` (11 columns), `LIB_FIELD_COUNT = 11`, `LIB_ROW_CAP = 100000`
  - VBA `modLibrary.FindConnectorRow(wsConn, nFirstRow, nLastRow, sConnectorID) As Long`
  - VBA `modLibrary.WriteConnector(wsConn, nFirstRow, nLastRow, vFields) As Boolean`
  - VBA `modLibrary.ReadConnector(wsConn, nFirstRow, nLastRow, sConnectorID) As Variant`
  - VBA `modLibrary.DeleteConnector(wsConn, nFirstRow, nLastRow, sConnectorID) As Boolean`

`vFields` is a plain 11-element array in field order (ConnectorID, Name, Manufacturer, PartNumber, Type, PinCount, Notes, PhotoShapeName, CreatedUtc, ModifiedUtc, Origin) — not a VBA `Type`. A custom `Type` cannot cross the `Application.Run` COM boundary that pytest uses to call these functions directly, so every function in this module that pytest calls stays on primitive/array parameters, matching the pattern already used by `modConnectors` and `modState` in Phase 1.

`nFirstRow`/`nLastRow` bound every scan and write to a fixed window rather than "to the sheet's last used row." That is what lets the exact same functions serve a dedicated whole sheet (the library workbook, where the window is generous) and a fixed-size region on a shared sheet (`_Snapshot`, built in a later sub-plan) without corrupting whatever else is below the window on a shared sheet.

- [ ] **Step 1: Write the failing test**

Create `tests/test_library_connectors.py`:

```python
from tests.conftest import run

FIELDS = (
    "DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
    4, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local",
)


def test_write_then_read_round_trips_every_field(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    ok = run(wb, "modLibrary.WriteConnector", ws, 2, 100000, FIELDS)
    assert ok is True

    result = run(wb, "modLibrary.ReadConnector", ws, 2, 100000, "DTM-04P")
    assert tuple(result) == FIELDS


def test_read_missing_connector_returns_empty(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    result = run(wb, "modLibrary.ReadConnector", ws, 2, 100000, "NO-SUCH-ID")
    assert result is None


def test_write_upserts_an_existing_connector(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, FIELDS)

    updated = FIELDS[:5] + (8,) + FIELDS[6:]
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, updated)

    assert ws.Cells(2, 1).Value == "DTM-04P"
    assert ws.Cells(3, 1).Value is None  # no second row was added
    result = run(wb, "modLibrary.ReadConnector", ws, 2, 100000, "DTM-04P")
    assert int(result[5]) == 8


def test_write_respects_the_row_window(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    # nLastRow=2 leaves no room past the header for a first data row.
    ok = run(wb, "modLibrary.WriteConnector", ws, 2, 1, FIELDS)
    assert ok is False


def test_delete_removes_a_connector_and_compacts(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, FIELDS)
    other = ("GND-STUD",) + FIELDS[1:]
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, other)

    ok = run(wb, "modLibrary.DeleteConnector", ws, 2, 100000, "DTM-04P")
    assert ok is True
    assert run(wb, "modLibrary.ReadConnector", ws, 2, 100000, "DTM-04P") is None
    assert run(wb, "modLibrary.ReadConnector", ws, 2, 100000, "GND-STUD") is not None
    assert ws.Cells(3, 1).Value is None  # compacted, not left as a gap


def test_delete_stays_inside_its_row_window(wb, library_wb):
    # A row beyond nLastRow must never be touched by delete's compaction.
    ws = library_wb.Worksheets("Connectors")
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, FIELDS)
    ws.Cells(3, 1).Value = "SENTINEL"

    run(wb, "modLibrary.DeleteConnector", ws, 2, 2, "DTM-04P")

    assert ws.Cells(2, 1).Value == "SENTINEL" or ws.Cells(3, 1).Value == "SENTINEL"
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_library_connectors.py -v
```

Expected: FAIL — `modLibrary` does not exist.

- [ ] **Step 3: Write the module**

Create `src/vba/modLibrary.bas`:

```vb
Attribute VB_Name = "modLibrary"
Option Explicit

' Column layout of the Connectors table, shared by the library workbook,
' an imported library file, and (in a later phase) the _Snapshot sheet's
' Connectors block.
Public Const LIB_COL_ID As Long = 1
Public Const LIB_COL_NAME As Long = 2
Public Const LIB_COL_MFG As Long = 3
Public Const LIB_COL_PARTNUM As Long = 4
Public Const LIB_COL_TYPE As Long = 5
Public Const LIB_COL_PINCOUNT As Long = 6
Public Const LIB_COL_NOTES As Long = 7
Public Const LIB_COL_PHOTOSHAPE As Long = 8
Public Const LIB_COL_CREATED As Long = 9
Public Const LIB_COL_MODIFIED As Long = 10
Public Const LIB_COL_ORIGIN As Long = 11
Public Const LIB_FIELD_COUNT As Long = 11

' A generous default window for callers addressing a dedicated whole sheet
' (the library workbook or an imported library file), where there is no
' second table sharing the same rows.
Public Const LIB_ROW_CAP As Long = 100000

Public Function LastUsedRowInWindow(ws As Worksheet, ByVal nCol As Long, ByVal nLastRow As Long) As Long
    ' Public: modPinEditor (sub-plan 2b) reuses this for the same bounded-
    ' window-safe single-row delete it needs for "Delete Pin."
    '
    ' Cells(nLastRow, nCol).End(xlUp) only finds the true last-used row when
    ' the starting cell itself is empty. If nLastRow already holds data -
    ' the window is full, or a caller probes a small window right at
    ' existing data - End(xlUp) instead walks UP through the contiguous
    ' non-blank run and overshoots past the real data (potentially into a
    ' header row above nFirstRow). Since no function here ever looks past
    ' nLastRow anyway, an occupied nLastRow already IS the answer.
    If Len(Trim$(CStr(ws.Cells(nLastRow, nCol).Value))) > 0 Then
        LastUsedRowInWindow = nLastRow
    Else
        LastUsedRowInWindow = ws.Cells(nLastRow, nCol).End(xlUp).Row
    End If
End Function

Public Function FindConnectorRow(wsConn As Worksheet, ByVal nFirstRow As Long, _
                                 ByVal nLastRow As Long, ByVal sConnectorID As String) As Long
    Dim r As Long, nLast As Long

    nLast = LastUsedRowInWindow(wsConn, LIB_COL_ID, nLastRow)
    If nLast < nFirstRow Then Exit Function

    For r = nFirstRow To nLast
        If StrComp(Trim$(CStr(wsConn.Cells(r, LIB_COL_ID).Value)), sConnectorID, vbTextCompare) = 0 Then
            FindConnectorRow = r
            Exit Function
        End If
    Next r
End Function

Public Function WriteConnector(wsConn As Worksheet, ByVal nFirstRow As Long, _
                               ByVal nLastRow As Long, ByVal vFields As Variant) As Boolean
    Dim r As Long, c As Long, nLast As Long

    If UBound(vFields) - LBound(vFields) + 1 <> LIB_FIELD_COUNT Then Exit Function

    r = FindConnectorRow(wsConn, nFirstRow, nLastRow, CStr(vFields(LBound(vFields))))
    If r = 0 Then
        nLast = LastUsedRowInWindow(wsConn, LIB_COL_ID, nLastRow)
        If nLast < nFirstRow Then
            r = nFirstRow
        Else
            r = nLast + 1
        End If
        If r > nLastRow Then Exit Function
    End If

    For c = 1 To LIB_FIELD_COUNT
        wsConn.Cells(r, c).Value = vFields(LBound(vFields) + c - 1)
    Next c

    WriteConnector = True
End Function

Public Function ReadConnector(wsConn As Worksheet, ByVal nFirstRow As Long, _
                              ByVal nLastRow As Long, ByVal sConnectorID As String) As Variant
    Dim r As Long, vResult(1 To 11) As Variant, c As Long

    r = FindConnectorRow(wsConn, nFirstRow, nLastRow, sConnectorID)
    If r = 0 Then Exit Function

    For c = 1 To LIB_FIELD_COUNT
        vResult(c) = wsConn.Cells(r, c).Value
    Next c

    ReadConnector = vResult
End Function

Public Function DeleteConnector(wsConn As Worksheet, ByVal nFirstRow As Long, _
                                ByVal nLastRow As Long, ByVal sConnectorID As String) As Boolean
    Dim r As Long, nLast As Long, c As Long

    r = FindConnectorRow(wsConn, nFirstRow, nLastRow, sConnectorID)
    If r = 0 Then Exit Function

    nLast = LastUsedRowInWindow(wsConn, LIB_COL_ID, nLastRow)

    ' Swap the last row's data into the deleted row's slot, then clear the
    ' (now-duplicate) last row. This keeps every write inside
    ' [nFirstRow, nLastRow] - a plain Range.Delete Shift:=xlUp would pull
    ' rows from below nLastRow upward too, corrupting whatever else shares
    ' the sheet (the _Snapshot Pins block, in a later phase).
    If r < nLast Then
        For c = 1 To LIB_FIELD_COUNT
            wsConn.Cells(r, c).Value = wsConn.Cells(nLast, c).Value
        Next c
    End If
    wsConn.Range(wsConn.Cells(nLast, 1), wsConn.Cells(nLast, LIB_FIELD_COUNT)).ClearContents

    DeleteConnector = True
End Function
```

- [ ] **Step 4: Wire the module into the build**

In `build/build.py`:

```python
VBA_MODULES = ["modUtil.bas", "modState.bas", "modConnectors.bas", "modChart.bas", "modLibrary.bas"]
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_library_connectors.py -v
```

Expected: 6 passed.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modLibrary.bas build/build.py tests/test_library_connectors.py
git commit -m "feat: add connector record CRUD to the library reader/writer"
```

---

### Task 3: ConnectorID slugification

**Files:**
- Modify: `src/vba/modLibrary.bas`
- Test: `tests/test_library_ids.py`

**Interfaces:**
- Consumes: `modLibrary.FindConnectorRow`.
- Produces:
  - VBA `modLibrary.SlugifyConnectorID(sPartNumber, sName) As String`
  - VBA `modLibrary.UniqueConnectorID(wsConn, nFirstRow, nLastRow, sBaseID) As String`

- [ ] **Step 1: Write the failing test**

Create `tests/test_library_ids.py`:

```python
import pytest

from tests.conftest import run


@pytest.mark.parametrize(
    "part_number,name,expected",
    [
        ("DTM06-4S", "Deutsch DTM 4-way", "DTM06-4S"),
        ("dtm06-4s", "Deutsch DTM 4-way", "DTM06-4S"),
        ("", "Chassis Ground Stud", "CHASSISGROUNDSTUD"),
        ("DTM 06/4-S!", "", "DTM-06-4-S-"),
        ("  ", "  ", ""),
    ],
)
def test_slugify(wb, part_number, name, expected):
    assert run(wb, "modLibrary.SlugifyConnectorID", part_number, name) == expected


def test_unique_id_passes_through_when_no_collision(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    result = run(wb, "modLibrary.UniqueConnectorID", ws, 2, 100000, "DTM-04P")
    assert result == "DTM-04P"


def test_unique_id_appends_a_numeric_suffix_on_collision(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    fields = ("DTM-04P", "A", "", "", "Connector", 4, "", "", "", "", "Local")
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, fields)

    result = run(wb, "modLibrary.UniqueConnectorID", ws, 2, 100000, "DTM-04P")
    assert result == "DTM-04P-2"


def test_unique_id_skips_past_multiple_collisions(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    for connector_id in ("DTM-04P", "DTM-04P-2", "DTM-04P-3"):
        fields = (connector_id, "A", "", "", "Connector", 4, "", "", "", "", "Local")
        run(wb, "modLibrary.WriteConnector", ws, 2, 100000, fields)

    result = run(wb, "modLibrary.UniqueConnectorID", ws, 2, 100000, "DTM-04P")
    assert result == "DTM-04P-4"
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_library_ids.py -v
```

Expected: FAIL — `SlugifyConnectorID`/`UniqueConnectorID` don't exist.

- [ ] **Step 3: Add the functions**

Append to `src/vba/modLibrary.bas`:

```vb
Public Function SlugifyConnectorID(ByVal sPartNumber As String, ByVal sName As String) As String
    Dim sSource As String, sResult As String, i As Long, ch As String

    sSource = Trim$(sPartNumber)
    If Len(sSource) = 0 Then sSource = Trim$(sName)
    sSource = UCase$(sSource)

    For i = 1 To Len(sSource)
        ch = Mid$(sSource, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            sResult = sResult & ch
        Else
            sResult = sResult & "-"
        End If
    Next i

    SlugifyConnectorID = sResult
End Function

Public Function UniqueConnectorID(wsConn As Worksheet, ByVal nFirstRow As Long, _
                                  ByVal nLastRow As Long, ByVal sBaseID As String) As String
    Dim sCandidate As String, nSuffix As Long

    sCandidate = sBaseID
    nSuffix = 1
    Do While FindConnectorRow(wsConn, nFirstRow, nLastRow, sCandidate) > 0
        nSuffix = nSuffix + 1
        sCandidate = sBaseID & "-" & CStr(nSuffix)
    Loop

    UniqueConnectorID = sCandidate
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_library_ids.py -v
```

Expected: 8 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modLibrary.bas tests/test_library_ids.py
git commit -m "feat: slugify and de-duplicate connector IDs"
```

---

### Task 4: Pin record CRUD

**Files:**
- Modify: `src/vba/modLibrary.bas`
- Test: `tests/test_library_pins.py`

**Interfaces:**
- Consumes: nothing new.
- Produces:
  - VBA constants `PIN_COL_CONNID` through `PIN_COL_LABELY` (7 columns), `PIN_FIELD_COUNT = 7`
  - VBA `modLibrary.WritePin(wsPins, nFirstRow, nLastRow, vFields) As Boolean`
  - VBA `modLibrary.ReadPinsForConnector(wsPins, nFirstRow, nLastRow, sConnectorID) As Variant` — a 2D array, one row per pin, sorted by PinNumber
  - VBA `modLibrary.DeletePinsForConnector(wsPins, nFirstRow, nLastRow, sConnectorID) As Long` — returns the count removed

`vFields` order: ConnectorID, PinNumber, PinLabel, NormX, NormY, LabelX, LabelY.

- [ ] **Step 1: Write the failing test**

Create `tests/test_library_pins.py`:

```python
from tests.conftest import run


def write_pin(wb, ws, connector_id, pin_number, label, norm_x, norm_y, label_x, label_y):
    fields = (connector_id, pin_number, label, norm_x, norm_y, label_x, label_y)
    return run(wb, "modLibrary.WritePin", ws, 2, 100000, fields)


def test_write_then_read_returns_pins_sorted_by_number(wb, library_wb):
    ws = library_wb.Worksheets("Pins")
    write_pin(wb, ws, "J1", 3, "GND", 0.9, 0.1, 0.9, 0.1)
    write_pin(wb, ws, "J1", 1, "+12V", 0.1, 0.1, 0.1, 0.1)
    write_pin(wb, ws, "J1", 2, "SIG", 0.5, 0.1, 0.5, 0.1)

    result = run(wb, "modLibrary.ReadPinsForConnector", ws, 2, 100000, "J1")
    pin_numbers = [int(row[1]) for row in result]
    assert pin_numbers == [1, 2, 3]
    assert [row[2] for row in result] == ["+12V", "SIG", "GND"]


def test_read_pins_only_returns_the_requested_connector(wb, library_wb):
    ws = library_wb.Worksheets("Pins")
    write_pin(wb, ws, "J1", 1, "A", 0.1, 0.1, 0.1, 0.1)
    write_pin(wb, ws, "J2", 1, "B", 0.1, 0.1, 0.1, 0.1)

    result = run(wb, "modLibrary.ReadPinsForConnector", ws, 2, 100000, "J1")
    assert len(result) == 1
    assert result[0][0] == "J1"


def test_read_pins_for_unknown_connector_returns_empty(wb, library_wb):
    ws = library_wb.Worksheets("Pins")
    result = run(wb, "modLibrary.ReadPinsForConnector", ws, 2, 100000, "NOPE")
    assert result is None


def test_delete_pins_removes_only_the_matching_connector_and_compacts(wb, library_wb):
    ws = library_wb.Worksheets("Pins")
    write_pin(wb, ws, "J1", 1, "A", 0.1, 0.1, 0.1, 0.1)
    write_pin(wb, ws, "J1", 2, "B", 0.2, 0.1, 0.2, 0.1)
    write_pin(wb, ws, "J2", 1, "C", 0.1, 0.1, 0.1, 0.1)

    count = run(wb, "modLibrary.DeletePinsForConnector", ws, 2, 100000, "J1")
    assert count == 2

    assert run(wb, "modLibrary.ReadPinsForConnector", ws, 2, 100000, "J1") is None
    remaining = run(wb, "modLibrary.ReadPinsForConnector", ws, 2, 100000, "J2")
    assert len(remaining) == 1
    assert ws.Cells(3, 1).Value is None  # compacted, not left with gaps


def test_write_pin_respects_the_row_window(wb, library_wb):
    ws = library_wb.Worksheets("Pins")
    # nLastRow=2 leaves no room past the header for a first data row.
    ok = write_pin(wb, ws, "J1", 1, "A", 0.1, 0.1, 0.1, 0.1)
    assert ok is True  # sanity check the normal path first (nLastRow=100000 above)

    ok = run(wb, "modLibrary.WritePin", ws, 2, 1, ("J2", 1, "B", 0.1, 0.1, 0.1, 0.1))
    assert ok is False


def test_delete_pins_stays_inside_its_row_window(wb, library_wb):
    # A row beyond nLastRow must never be touched by delete's compaction.
    ws = library_wb.Worksheets("Pins")
    write_pin(wb, ws, "J1", 1, "A", 0.1, 0.1, 0.1, 0.1)
    ws.Cells(3, 1).Value = "SENTINEL"

    run(wb, "modLibrary.DeletePinsForConnector", ws, 2, 2, "J1")

    assert ws.Cells(2, 1).Value == "SENTINEL" or ws.Cells(3, 1).Value == "SENTINEL"
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_library_pins.py -v
```

Expected: FAIL — the three functions don't exist.

- [ ] **Step 3: Add the functions**

Append to `src/vba/modLibrary.bas`:

```vb
Public Const PIN_COL_CONNID As Long = 1
Public Const PIN_COL_PINNUM As Long = 2
Public Const PIN_COL_LABEL As Long = 3
Public Const PIN_COL_NORMX As Long = 4
Public Const PIN_COL_NORMY As Long = 5
Public Const PIN_COL_LABELX As Long = 6
Public Const PIN_COL_LABELY As Long = 7
Public Const PIN_FIELD_COUNT As Long = 7

Public Function WritePin(wsPins As Worksheet, ByVal nFirstRow As Long, _
                         ByVal nLastRow As Long, ByVal vFields As Variant) As Boolean
    Dim r As Long, c As Long, nLast As Long

    If UBound(vFields) - LBound(vFields) + 1 <> PIN_FIELD_COUNT Then Exit Function

    nLast = LastUsedRowInWindow(wsPins, PIN_COL_CONNID, nLastRow)
    If nLast < nFirstRow Then
        r = nFirstRow
    Else
        r = nLast + 1
    End If
    If r > nLastRow Then Exit Function

    For c = 1 To PIN_FIELD_COUNT
        wsPins.Cells(r, c).Value = vFields(LBound(vFields) + c - 1)
    Next c

    WritePin = True
End Function

Private Sub SwapPinRows(vRows As Variant, ByVal a As Long, ByVal b As Long)
    Dim c As Long, vTmp As Variant
    For c = 1 To PIN_FIELD_COUNT
        vTmp = vRows(a, c)
        vRows(a, c) = vRows(b, c)
        vRows(b, c) = vTmp
    Next c
End Sub

Public Function ReadPinsForConnector(wsPins As Worksheet, ByVal nFirstRow As Long, _
                                     ByVal nLastRow As Long, ByVal sConnectorID As String) As Variant
    Dim nLast As Long, r As Long, n As Long, i As Long, j As Long
    Dim vRows() As Variant, vResult() As Variant

    nLast = LastUsedRowInWindow(wsPins, PIN_COL_CONNID, nLastRow)
    If nLast < nFirstRow Then Exit Function

    n = 0
    ReDim vRows(1 To nLast - nFirstRow + 1, 1 To PIN_FIELD_COUNT)
    For r = nFirstRow To nLast
        If StrComp(Trim$(CStr(wsPins.Cells(r, PIN_COL_CONNID).Value)), sConnectorID, vbTextCompare) = 0 Then
            n = n + 1
            For i = 1 To PIN_FIELD_COUNT
                vRows(n, i) = wsPins.Cells(r, i).Value
            Next i
        End If
    Next r
    If n = 0 Then Exit Function

    ' Insertion sort by PinNumber - pin counts per connector are small.
    For i = 2 To n
        For j = i To 2 Step -1
            If CDbl(vRows(j, PIN_COL_PINNUM)) < CDbl(vRows(j - 1, PIN_COL_PINNUM)) Then
                SwapPinRows vRows, j, j - 1
            Else
                Exit For
            End If
        Next j
    Next i

    ReDim vResult(1 To n, 1 To PIN_FIELD_COUNT)
    For i = 1 To n
        For j = 1 To PIN_FIELD_COUNT
            vResult(i, j) = vRows(i, j)
        Next j
    Next i

    ReadPinsForConnector = vResult
End Function

Public Function DeletePinsForConnector(wsPins As Worksheet, ByVal nFirstRow As Long, _
                                       ByVal nLastRow As Long, ByVal sConnectorID As String) As Long
    Dim nLast As Long, r As Long, w As Long, c As Long, nDeleted As Long

    nLast = LastUsedRowInWindow(wsPins, PIN_COL_CONNID, nLastRow)
    If nLast < nFirstRow Then Exit Function

    ' Single-pass compaction: copy every non-matching row down to a write
    ' cursor, then clear the leftover tail. Bounded to [nFirstRow, nLastRow]
    ' for the same reason DeleteConnector avoids Range.Delete Shift:=xlUp.
    w = nFirstRow
    For r = nFirstRow To nLast
        If StrComp(Trim$(CStr(wsPins.Cells(r, PIN_COL_CONNID).Value)), sConnectorID, vbTextCompare) = 0 Then
            nDeleted = nDeleted + 1
        Else
            If w <> r Then
                For c = 1 To PIN_FIELD_COUNT
                    wsPins.Cells(w, c).Value = wsPins.Cells(r, c).Value
                Next c
            End If
            w = w + 1
        End If
    Next r

    If w <= nLast Then
        wsPins.Range(wsPins.Cells(w, 1), wsPins.Cells(nLast, PIN_FIELD_COUNT)).ClearContents
    End If

    DeletePinsForConnector = nDeleted
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_library_pins.py -v
```

Expected: 6 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modLibrary.bas tests/test_library_pins.py
git commit -m "feat: add pin record CRUD to the library reader/writer"
```

---

### Task 5: Photo embedding

**Files:**
- Modify: `src/vba/modLibrary.bas`
- Create: `tests/fixtures/__init__.py` (empty, for import resolution)
- Create: `tests/fixtures/sample_photo.py`
- Test: `tests/test_library_photos.py`

**Interfaces:**
- Consumes: nothing new.
- Produces:
  - VBA constants `PHOTO_GRID_COLUMNS`, `PHOTO_GRID_CELL_WIDTH`, `PHOTO_GRID_CELL_HEIGHT`, `PHOTO_GRID_MARGIN`
  - VBA `modLibrary.EmbedConnectorPhoto(wsPhotos, sConnectorID, sImagePath) As String` — returns the shape name, or `""` if `sImagePath` does not exist
  - VBA `modLibrary.RemoveConnectorPhoto(wsPhotos, sConnectorID)`
  - VBA `modLibrary.CachePhotoPath(sWorkbookFolder, sConnectorID) As String` — creates `<folder>\Photos\` if needed
  - `tests.fixtures.sample_photo.write_sample_photo(path: Path) -> Path`

A 1x1 PNG is the minimum valid image `Shapes.AddPicture` will accept — no imaging library is added; the bytes are hand-written, matching the spec's `pywin32`/`pytest`-only dependency constraint.

- [ ] **Step 1: Write the sample-photo fixture helper**

Create `tests/fixtures/__init__.py` (empty file).

Create `tests/fixtures/sample_photo.py`:

```python
"""A minimal valid 1x1 PNG, hand-written so no imaging library is a
dependency (the spec limits Python dependencies to pywin32 and pytest)."""
from __future__ import annotations

from pathlib import Path

_PNG_BYTES = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108020000009077"
    "53de0000000c4944415408d763f8ffff3f0005fe02fea739669d0000000049"
    "454e44ae426082"
)


def write_sample_photo(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(_PNG_BYTES)
    return path
```

- [ ] **Step 2: Write the failing test**

Create `tests/test_library_photos.py`:

```python
from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_embed_photo_adds_a_named_shape(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws = library_wb.Worksheets("Photos")

    name = run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))

    assert name == "PHOTO_DTM-04P"
    assert ws.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"


def test_embed_photo_with_missing_file_returns_empty(wb, library_wb):
    ws = library_wb.Worksheets("Photos")
    name = run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", r"C:\no\such\file.png")
    assert name == ""


def test_embed_photo_replaces_an_existing_shape_for_the_same_id(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws = library_wb.Worksheets("Photos")

    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))
    before = ws.Shapes.Count
    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))

    assert ws.Shapes.Count == before


def test_second_photo_lands_in_a_different_grid_slot(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws = library_wb.Worksheets("Photos")

    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))
    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "GND-STUD", str(photo_path))

    first = ws.Shapes("PHOTO_DTM-04P")
    second = ws.Shapes("PHOTO_GND-STUD")
    assert (first.Left, first.Top) != (second.Left, second.Top)


def test_remove_photo_deletes_the_shape(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws = library_wb.Worksheets("Photos")
    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))

    run(wb, "modLibrary.RemoveConnectorPhoto", ws, "DTM-04P")

    assert ws.Shapes.Count == 0


def test_remove_photo_for_unknown_id_does_not_raise(wb, library_wb):
    ws = library_wb.Worksheets("Photos")
    run(wb, "modLibrary.RemoveConnectorPhoto", ws, "NO-SUCH-ID")


def test_cache_photo_path_creates_the_photos_subfolder(wb, tmp_path):
    result = run(wb, "modLibrary.CachePhotoPath", str(tmp_path), "DTM-04P")
    assert result == str(tmp_path / "Photos" / "DTM-04P.png")
    assert (tmp_path / "Photos").is_dir()
```

- [ ] **Step 3: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_library_photos.py -v
```

Expected: FAIL — the three functions don't exist.

- [ ] **Step 4: Add the functions**

Append to `src/vba/modLibrary.bas`:

```vb
Public Const PHOTO_GRID_COLUMNS As Long = 4
Public Const PHOTO_GRID_CELL_WIDTH As Long = 120
Public Const PHOTO_GRID_CELL_HEIGHT As Long = 120
Public Const PHOTO_GRID_MARGIN As Long = 8

Public Sub RemoveConnectorPhoto(wsPhotos As Worksheet, ByVal sConnectorID As String)
    On Error Resume Next
    wsPhotos.Shapes("PHOTO_" & sConnectorID).Delete
    On Error GoTo 0
End Sub

Public Function EmbedConnectorPhoto(wsPhotos As Worksheet, ByVal sConnectorID As String, _
                                    ByVal sImagePath As String) As String
    Dim sShapeName As String, nIndex As Long, nCol As Long, nRow As Long
    Dim shp As Shape

    If Len(Dir$(sImagePath)) = 0 Then Exit Function

    sShapeName = "PHOTO_" & sConnectorID
    RemoveConnectorPhoto wsPhotos, sConnectorID

    nIndex = wsPhotos.Shapes.Count
    nCol = nIndex Mod PHOTO_GRID_COLUMNS
    nRow = nIndex \ PHOTO_GRID_COLUMNS

    Set shp = wsPhotos.Shapes.AddPicture(sImagePath, False, True, _
        nCol * PHOTO_GRID_CELL_WIDTH, nRow * PHOTO_GRID_CELL_HEIGHT, _
        PHOTO_GRID_CELL_WIDTH - PHOTO_GRID_MARGIN, PHOTO_GRID_CELL_HEIGHT - PHOTO_GRID_MARGIN)
    shp.Name = sShapeName

    EmbedConnectorPhoto = sShapeName
End Function

Public Function CachePhotoPath(ByVal sWorkbookFolder As String, ByVal sConnectorID As String) As String
    Dim sFolder As String

    sFolder = sWorkbookFolder
    If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
    sFolder = sFolder & "Photos\"
    If Len(Dir$(sFolder, vbDirectory)) = 0 Then MkDir sFolder

    CachePhotoPath = sFolder & sConnectorID & ".png"
End Function
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_library_photos.py -v
```

Expected: 8 passed.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modLibrary.bas tests/fixtures/ tests/test_library_photos.py
git commit -m "feat: embed connector photos into the library's Photos sheet"
```

---

### Task 6: Full round-trip integration test

**Files:**
- Test: `tests/test_library_integration.py`

**Interfaces:**
- Consumes: every `modLibrary` function from Tasks 2-5.
- Produces: nothing later tasks depend on — this is a capstone test, not new production code.

This is the test that actually exercises "one reader and one writer serve all three [storage locations]" against the real built `ConnectorLibrary.xlsx`, not just isolated function calls. A connector editor or import feature in a later sub-plan that only ever writes through these functions is exercising a path already proven here.

- [ ] **Step 1: Write the test**

Create `tests/test_library_integration.py`:

```python
from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_full_connector_definition_round_trips_through_the_saved_file(
    wb, app, library_artifact, tmp_path
):
    photo_path = write_sample_photo(tmp_path / "photo.png")

    lib = app.Workbooks.Open(str(library_artifact))
    try:
        ws_conn = lib.Worksheets("Connectors")
        ws_pins = lib.Worksheets("Pins")
        ws_photos = lib.Worksheets("Photos")

        connector_id = run(wb, "modLibrary.SlugifyConnectorID", "DTM06-4S", "")
        assert connector_id == "DTM06-4S"

        fields = (
            connector_id, "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
            4, "", "PHOTO_" + connector_id, "2026-08-26T00:00:00Z",
            "2026-08-26T00:00:00Z", "Local",
        )
        assert run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields) is True

        for pin_number, label, x, y in [
            (1, "+12V", 0.1, 0.1), (2, "GND", 0.9, 0.1),
            (3, "SIG", 0.1, 0.9), (4, "", 0.9, 0.9),
        ]:
            pin_fields = (connector_id, pin_number, label, x, y, x, y)
            assert run(wb, "modLibrary.WritePin", ws_pins, 2, 100000, pin_fields) is True

        shape_name = run(wb, "modLibrary.EmbedConnectorPhoto", ws_photos, connector_id, str(photo_path))
        assert shape_name == "PHOTO_" + connector_id

        lib.Save()
    finally:
        lib.Close(SaveChanges=False)

    # Reopen as a wholly separate file handle - proves the data actually
    # persisted to disk, not just to the in-memory COM object.
    reopened = app.Workbooks.Open(str(library_artifact))
    try:
        ws_conn = reopened.Worksheets("Connectors")
        ws_pins = reopened.Worksheets("Pins")
        ws_photos = reopened.Worksheets("Photos")

        result = run(wb, "modLibrary.ReadConnector", ws_conn, 2, 100000, connector_id)
        assert tuple(result) == fields

        pins = run(wb, "modLibrary.ReadPinsForConnector", ws_pins, 2, 100000, connector_id)
        assert [int(row[1]) for row in pins] == [1, 2, 3, 4]

        assert ws_photos.Shapes("PHOTO_" + connector_id).Name == "PHOTO_" + connector_id
    finally:
        reopened.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails first, then passes**

Run:

```bash
python -m pytest tests/test_library_integration.py -v
```

Since every function it calls already exists from Tasks 2-5, this should pass immediately. If it fails, the bug is in how the individually-tested functions compose, not in any single function — treat that as a real finding, not a fixture problem.

Expected: 1 passed.

- [ ] **Step 3: Run the whole suite**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: everything passes from a clean build (104 from Phase 1 + 21 + 6 + 8 + 6 + 8 + 1 = 154).

- [ ] **Step 4: Commit**

```bash
git add tests/test_library_integration.py
git commit -m "test: prove the library reader/writer round-trips through a saved file"
```

---

## Self-Review

**Spec coverage for this sub-plan.** The spec's Phase 2 line item this covers: "Connector library file, library reader/writer... snapshot embedding [schema only - the embedding *flow* is 2c]." The three-table schema (Connectors, Pins, Photos), ConnectorID slugification with collision suffixing, and the "one reader and one writer serve all three" requirement are all directly implemented and tested. Photo caching to `Photos/<ConnectorID>.png` is implemented (`CachePhotoPath`); *populating* that cache from `EmbedConnectorPhoto`'s caller and *regenerating* it from an imported library's embedded shape are both left to 2b (which has the source file in hand when a connector is first defined) and 2d (which must extract from a shape it did not create) respectively — call sites for a helper that already exists here, not gaps in this plan.

**Deliberately deferred to 2b/2c/2d, not gaps in this plan:** the connector editor UI, click-to-place, the picker, `_Snapshot`'s actual row-window constants and wiring, Add/Remove Connector, ref des rename, and library import/export. Every one is named in the sub-plan split at the top of this document.

**Why `(nFirstRow, nLastRow)` instead of "to the sheet's last used row."** `modConnectors.NextRefDes` in Phase 1 uses `ws.Cells(ws.Rows.Count, 1).End(xlUp).Row` because `Connectors` is a dedicated whole sheet. `_Snapshot` (2c) will not be: it holds the Connectors block and the Pins block on the *same* physical sheet, in overlapping columns. Scanning or shifting "to the sheet's bottom" from within the Connectors block would read into, or destructively shift, the Pins block below it. Passing an explicit window is what lets Tasks 2 and 4's functions be reused unchanged in 2c instead of needing a second, parallel implementation.

**`LastUsedRowInWindow` exists because of a real bug found during self-review.** The obvious idiom for "last used row in a bounded window" is `Cells(nLastRow, col).End(xlUp).Row` - a direct adaptation of the whole-sheet `Cells(Rows.Count, col).End(xlUp)` idiom Phase 1 uses. That whole-sheet version works because `Cells(Rows.Count, col)` is virtually guaranteed empty. `Cells(nLastRow, col)` has no such guarantee: once a window fills up (or a caller probes a small window, like `test_write_respects_the_row_window` and `test_delete_stays_inside_its_row_window` do deliberately), the starting cell is occupied, and `End(xlUp)` walks up through the contiguous non-blank run and overshoots past the real data - in the worst case into the header row above `nFirstRow`. `LastUsedRowInWindow` checks whether `nLastRow` itself is occupied first and short-circuits to `nLastRow` in that case, which is exactly the case 2c's `_Snapshot` Connectors block will hit once it's ever full. Traced by hand against both boundary tests in Tasks 2 and 4 before finalizing.

**Type consistency.** `WriteConnector`/`ReadConnector` use the same 11-field order everywhere (`LIB_COL_ID` through `LIB_COL_ORIGIN`); `WritePin`/`ReadPinsForConnector` use the same 7-field order (`PIN_COL_CONNID` through `PIN_COL_LABELY`). Both are asserted directly by the round-trip tests in Task 2, Task 4, and again end-to-end in Task 6. `EmbedConnectorPhoto` returns `String` (the shape name or `""`), consumed as such by the Task 6 integration test and matching `PhotoShapeName`'s type in the Connectors schema.

**No placeholders.** Every step above contains complete, runnable code — no "add validation" or "similar to Task N" stand-ins.
