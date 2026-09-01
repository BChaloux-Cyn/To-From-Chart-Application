# Phase 3b: Connector Page Rendering Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extend `SaveHarness` so it also writes one `CONN_<RefDes>` sheet per connector instance: the connector's photo at a fixed anchor, one numbered oval callout per pin at its stored marker position, and a leader line wherever a marker has been pulled off its anchor point. Live pin-table formulas (the numeric columns beside each callout) are 3c's addition, not this plan's - this plan reserves their header row and the two static columns (Pin, Label) only.

**Architecture:** A new layer 0 module, `modConnectorPage.bas`, does the per-page drawing given a `Worksheet` already created for it plus the pin geometry read out of `_Snapshot`. It reuses `modPinEditor.FitAspectRatio`, `MarkerTopLeft`, and `MarkerSitsOnAnchor` unchanged - all three take plain `Double`s, not a worksheet, so they compose against `_Snapshot`'s pin rows exactly as they already do against the connector editor's `_Edit` scratch sheet, with no new geometry math. `modHarnessBuild.bas` gains the connector-instance-list read and the orchestration loop that creates one page per instance; `modHarnessActions.SaveHarness` calls it as one more step in the same transaction Task pipeline 3a already built.

**Tech Stack:** VBA (Excel 16.0 COM automation), Python 3.13/pywin32/pytest.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Depends on:** 3a (`docs/superpowers/plans/2026-09-01-phase-3a-harness-save-shell.md`), implemented and merged - this plan extends `modHarnessBuild.bas`, `modHarnessActions.SaveHarness`, and `tests/test_harness_save_integration.py`'s fixture data directly.

**Part of Phase 3** (see 3a's header for the full five-sub-plan breakdown). This is 3b.

## Global Constraints

- Windows/desktop Excel only, Excel 2016+ formulas (not exercised by this plan directly, but the metadata cell and reserved columns below exist for 3c, which needs INDEX/MATCH).
- `modConnectorPage.bas` is layer 0: it may reference `Worksheet`/`Range`/`Shape`/scalars and other layer 0 modules (`modPinEditor`, `modLibrary`), never `MsgBox`, dialogs, or workbook lifecycle calls.
- Every VBA module starts with `Option Explicit`.
- **Decision (carried over from 3a): no connector-instance table is duplicated into the saved file.** This plan is what 3e's Load will actually read to reconstruct the Creator's `Connectors` sheet - each `CONN_<RefDes>` sheet gets a hidden metadata cell holding its `ConnectorID`, and 3e enumerates `CONN_*` sheet names for the Ref Des half of the mapping. This plan must not skip writing that cell.
- **Decision: geometry constants are fixed points, not derived from photo size.** `CONN_PHOTO_LEFT/TOP` and `CONN_PHOTO_MAX_WIDTH/HEIGHT` are constants in `modConnectorPage.bas`; every page uses the same anchor and box, matching the spec's "the photo is placed at a fixed anchor, scaled to a fixed maximum width."
- **Decision: the photo source is the on-disk cache, not a clipboard copy.** `modSnapshot`'s existing `.jpg`-then-`.png` cache lookup (`modLibrary.CachePhotoPath`) is reused via `Shapes.AddPicture`, avoiding a second `Shape.Copy`/`Paste` hop beyond the one 3a's `CopySnapshot` already does - `docs/superpowers/plans/phase-2-manual-verification.md` documents that mechanism as occasionally flaky, so this plan does not add a second use of it where a file-based alternative already exists.

## File Structure

| File | Responsibility |
|---|---|
| `src/vba/modConnectorPage.bas` | Layer 0. Per-page drawing: photo placement, oval callouts, leader lines, the pin-table skeleton, the metadata cell. |
| `src/vba/modConnectors.bas` | Modified: adds `AllInstances()`, a full-row reader of the Creator's `Connectors` sheet (Ref Des, ConnectorID, Name, Part Number, Type, Pin Count per row). |
| `src/vba/modHarnessBuild.bas` | Modified: adds `BuildConnectorPages(destWb, wsSnapshot)`, which reads `modConnectors.AllInstances()` and calls `modConnectorPage` once per instance. |
| `src/vba/modHarnessActions.bas` | Modified: `SaveHarness` calls `BuildConnectorPages` after `CopySnapshot`. |
| `build/build.py` | Modified: adds `modConnectorPage.bas` to `VBA_MODULES`. |
| `tests/test_layering.py` | Modified: adds `modConnectorPage` to `LAYER0`. |
| `tests/test_connector_page.py` | `modConnectorPage`'s photo, callout, leader-line, and metadata-cell functions. |
| `tests/test_connectors.py` | Additions: `modConnectors.AllInstances`. |
| `tests/test_harness_save_integration.py` | Modified: the existing round-trip test gains connector-page assertions (oval count, positions, leader presence). |

---

### Task 1: Read all connector instances

**Files:**
- Modify: `src/vba/modConnectors.bas`
- Test: `tests/test_connectors.py` (additions)

**Interfaces:**
- Consumes: nothing new.
- Produces: VBA `modConnectors.AllInstances() As Variant` - a 2D array, one row per placed instance, columns (RefDes, ConnectorID, Name, PartNumber, Type, PinCount), `LBound` 1 on both dimensions, or `Empty` if none are placed.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_connectors.py`:

```python
def test_all_instances_returns_every_placed_connector(wb):
    ws = wb.Worksheets("Connectors")
    ws.Cells(2, 1).Value = "J1"
    ws.Cells(2, 2).Value = "DTM-04P"
    ws.Cells(2, 3).Value = "Deutsch DTM 4-way"
    ws.Cells(2, 4).Value = "DTM06-4S"
    ws.Cells(2, 5).Value = "Connector"
    ws.Cells(2, 6).Value = 4
    ws.Cells(3, 1).Value = "ST1"
    ws.Cells(3, 2).Value = "GND-STUD"
    ws.Cells(3, 3).Value = "Chassis Ground"
    ws.Cells(3, 4).Value = ""
    ws.Cells(3, 5).Value = "Stud"
    ws.Cells(3, 6).Value = 1

    result = run(wb, "modConnectors.AllInstances")
    assert [row[0] for row in result] == ["J1", "ST1"]
    assert result[0][1] == "DTM-04P"
    assert int(result[1][5]) == 1


def test_all_instances_returns_empty_when_nothing_is_placed(wb):
    result = run(wb, "modConnectors.AllInstances")
    assert result is None
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_connectors.py -v -k all_instances
```

Expected: FAIL - `AllInstances` does not exist.

- [ ] **Step 3: Write the function**

Append to `src/vba/modConnectors.bas`:

```vb
Public Function AllInstances() As Variant
    Dim ws As Worksheet
    Dim r As Long, nLast As Long, n As Long, c As Long
    Dim vResult() As Variant

    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If nLast < CONN_FIRST_ROW Then Exit Function

    n = nLast - CONN_FIRST_ROW + 1
    ReDim vResult(1 To n, 1 To 6)
    For r = CONN_FIRST_ROW To nLast
        For c = 1 To 6
            vResult(r - CONN_FIRST_ROW + 1, c) = ws.Cells(r, c).Value
        Next c
    Next r

    AllInstances = vResult
End Function
```

- [ ] **Step 4: Run the test to verify it passes**

Run:

```bash
python -m pytest tests/test_connectors.py -v -k all_instances
```

Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modConnectors.bas tests/test_connectors.py
git commit -m "feat: read every placed connector instance"
```

---

### Task 2: Photo placement

**Files:**
- Create: `src/vba/modConnectorPage.bas`
- Test: `tests/test_connector_page.py`

**Interfaces:**
- Consumes: `modPinEditor.FitAspectRatio` (existing), `modLibrary.CachePhotoPath` (existing).
- Produces:
  - VBA constants `CONN_PHOTO_LEFT = 20`, `CONN_PHOTO_TOP = 60`, `CONN_PHOTO_MAX_WIDTH = 300`, `CONN_PHOTO_MAX_HEIGHT = 300` (points), `CONN_META_COL = 27` (column AA), `CONN_TABLE_FIRST_COL = 10` (column J), `CONN_TABLE_HEADER_ROW = 1`, `CONN_TABLE_FIRST_ROW = 2`.
  - VBA `modConnectorPage.PagePhotoPath(ByVal sLibraryFolder As String, ByVal sConnectorID As String) As String` - the `.jpg` cache path if it exists, else the `.png` cache path if that exists, else `""`.
  - VBA `modConnectorPage.PlacePhoto(wsPage As Worksheet, ByVal sPhotoPath As String) As Boolean` - places the picture at the fixed anchor, sized by `FitAspectRatio` into the fixed box, named `"PAGE_PHOTO"`. Returns `False` (touches nothing) if `sPhotoPath` is blank or the file does not exist.

- [ ] **Step 1: Write the failing test**

Create `tests/test_connector_page.py`:

```python
from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_page_photo_path_prefers_jpg_over_png(wb, tmp_path):
    (tmp_path / "Photos").mkdir()
    write_sample_photo(tmp_path / "Photos" / "DTM-04P.png")
    (tmp_path / "Photos" / "DTM-04P.jpg").write_bytes((tmp_path / "Photos" / "DTM-04P.png").read_bytes())

    result = run(wb, "modConnectorPage.PagePhotoPath", str(tmp_path), "DTM-04P")
    assert result.endswith("DTM-04P.jpg")


def test_page_photo_path_falls_back_to_png(wb, tmp_path):
    write_sample_photo(tmp_path / "Photos" / "DTM-04P.png")
    result = run(wb, "modConnectorPage.PagePhotoPath", str(tmp_path), "DTM-04P")
    assert result.endswith("DTM-04P.png")


def test_page_photo_path_returns_empty_when_no_cache_exists(wb, tmp_path):
    result = run(wb, "modConnectorPage.PagePhotoPath", str(tmp_path), "NOPE")
    assert result == ""


def test_place_photo_adds_a_named_shape_at_the_fixed_anchor(wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        ok = run(wb, "modConnectorPage.PlacePhoto", ws, str(photo_path))
        assert ok is True
        shp = ws.Shapes("PAGE_PHOTO")
        assert shp.Left == 20
        assert shp.Top == 60
        assert shp.Width <= 300 and shp.Height <= 300
    finally:
        dest.Close(SaveChanges=False)


def test_place_photo_returns_false_for_a_missing_file(wb, app):
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        ok = run(wb, "modConnectorPage.PlacePhoto", ws, r"C:\no\such\file.png")
        assert ok is False
        assert ws.Shapes.Count == 0
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_connector_page.py -v
```

Expected: FAIL - `modConnectorPage` does not exist.

- [ ] **Step 3: Write the module**

Create `src/vba/modConnectorPage.bas`:

```vb
Attribute VB_Name = "modConnectorPage"
Option Explicit

Public Const CONN_PHOTO_LEFT As Double = 20
Public Const CONN_PHOTO_TOP As Double = 60
Public Const CONN_PHOTO_MAX_WIDTH As Double = 300
Public Const CONN_PHOTO_MAX_HEIGHT As Double = 300
Public Const CONN_META_COL As Long = 27
Public Const CONN_TABLE_FIRST_COL As Long = 10
Public Const CONN_TABLE_HEADER_ROW As Long = 1
Public Const CONN_TABLE_FIRST_ROW As Long = 2

Public Function PagePhotoPath(ByVal sLibraryFolder As String, ByVal sConnectorID As String) As String
    Dim sJpg As String, sPng As String
    sJpg = modLibrary.CachePhotoPath(sLibraryFolder, sConnectorID, "jpg")
    If Len(Dir$(sJpg)) > 0 Then
        PagePhotoPath = sJpg
        Exit Function
    End If

    sPng = modLibrary.CachePhotoPath(sLibraryFolder, sConnectorID)
    If Len(Dir$(sPng)) > 0 Then PagePhotoPath = sPng
End Function

Public Function PlacePhoto(wsPage As Worksheet, ByVal sPhotoPath As String) As Boolean
    If Len(sPhotoPath) = 0 Then Exit Function
    If Len(Dir$(sPhotoPath)) = 0 Then Exit Function

    Dim shpProbe As Shape
    Set shpProbe = wsPage.Shapes.AddPicture(sPhotoPath, False, True, 0, 0, -1, -1)
    Dim vFit As Variant
    vFit = modPinEditor.FitAspectRatio(shpProbe.Width, shpProbe.Height, CONN_PHOTO_MAX_WIDTH, CONN_PHOTO_MAX_HEIGHT)

    shpProbe.Left = CONN_PHOTO_LEFT
    shpProbe.Top = CONN_PHOTO_TOP
    shpProbe.LockAspectRatio = False
    shpProbe.Width = vFit(0)
    shpProbe.Height = vFit(1)
    shpProbe.Name = "PAGE_PHOTO"

    PlacePhoto = True
End Function
```

`AddPicture`'s `-1, -1` width/height loads the picture at its natural size first, purely so `.Width`/`.Height` can be read as the source dimensions for `FitAspectRatio` - the same "load first, then measure, then resize" order `frmConnectorEditor`'s photo preview already uses (`LoadPicture`'s `StdPicture` HIMETRIC dimensions), adapted to a `Shape` instead of an `Image` control.

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_connector_page.py -v
```

Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modConnectorPage.bas tests/test_connector_page.py
git commit -m "feat: place a connector page's photo at a fixed anchor"
```

---

### Task 3: Oval callouts

**Files:**
- Modify: `src/vba/modConnectorPage.bas`
- Test: `tests/test_connector_page.py` (additions)

**Interfaces:**
- Consumes: `modPinEditor.MarkerTopLeft` (existing), `modLibrary.ReadPinsForConnector`/`PIN_COL_*` (existing).
- Produces: VBA constants `CONN_OVAL_DIAMETER = 14` (points). VBA `modConnectorPage.PlaceCallouts(wsPage As Worksheet, shpPhoto As Shape, vPins As Variant) As Long` - one oval per row of `vPins` (the array `ReadPinsForConnector` returns), named `"PIN_<n>"`, white fill, black border, centered on `MarkerTopLeft`'s result for that pin's `LabelX`/`LabelY` against the photo's actual placed geometry. Returns the oval count.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_connector_page.py`:

```python
def test_place_callouts_draws_one_oval_per_pin(wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        run(wb, "modConnectorPage.PlacePhoto", ws, str(photo_path))
        shp = ws.Shapes("PAGE_PHOTO")

        pins = (
            ("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1),
            ("DTM-04P", 2, "GND", 0.9, 0.1, 0.9, 0.1),
        )
        n = run(wb, "modConnectorPage.PlaceCallouts", ws, shp, pins)
        assert n == 2
        assert ws.Shapes("PIN_1").Name == "PIN_1"
        assert ws.Shapes("PIN_2").Name == "PIN_2"
        assert ws.Shapes("PIN_1").TextFrame2.TextRange.Text == "1"
    finally:
        dest.Close(SaveChanges=False)


def test_place_callouts_centers_the_oval_on_the_label_position(wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        run(wb, "modConnectorPage.PlacePhoto", ws, str(photo_path))
        shp = ws.Shapes("PAGE_PHOTO")

        pins = (("DTM-04P", 1, "", 0.5, 0.5, 0.5, 0.5),)
        run(wb, "modConnectorPage.PlaceCallouts", ws, shp, pins)

        oval = ws.Shapes("PIN_1")
        expected_center_x = shp.Left + 0.5 * shp.Width
        expected_center_y = shp.Top + 0.5 * shp.Height
        assert abs((oval.Left + oval.Width / 2) - expected_center_x) < 0.5
        assert abs((oval.Top + oval.Height / 2) - expected_center_y) < 0.5
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_connector_page.py -v -k callouts
```

Expected: FAIL - `PlaceCallouts` does not exist.

- [ ] **Step 3: Write the function**

Append to `src/vba/modConnectorPage.bas`:

```vb
Public Const CONN_OVAL_DIAMETER As Double = 14

MSO_SHAPE_OVAL = 9

Public Function PlaceCallouts(wsPage As Worksheet, shpPhoto As Shape, vPins As Variant) As Long
    Dim i As Long, nPinNumber As Long, dLabelX As Double, dLabelY As Double
    Dim vTopLeft As Variant, shp As Shape, n As Long

    If IsEmpty(vPins) Then Exit Function

    For i = LBound(vPins, 1) To UBound(vPins, 1)
        nPinNumber = CLng(vPins(i, modLibrary.PIN_COL_PINNUM))
        dLabelX = CDbl(vPins(i, modLibrary.PIN_COL_LABELX))
        dLabelY = CDbl(vPins(i, modLibrary.PIN_COL_LABELY))

        vTopLeft = modPinEditor.MarkerTopLeft(dLabelX, dLabelY, _
            shpPhoto.Left, shpPhoto.Top, shpPhoto.Width, shpPhoto.Height, _
            CONN_OVAL_DIAMETER, CONN_OVAL_DIAMETER)

        Set shp = wsPage.Shapes.AddShape(MSO_SHAPE_OVAL, vTopLeft(0), vTopLeft(1), _
            CONN_OVAL_DIAMETER, CONN_OVAL_DIAMETER)
        shp.Name = "PIN_" & CStr(nPinNumber)
        shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.Line.ForeColor.RGB = RGB(0, 0, 0)
        shp.TextFrame2.TextRange.Text = CStr(nPinNumber)
        shp.TextFrame2.TextRange.Font.Size = 8
        shp.TextFrame2.WordWrap = False

        n = n + 1
    Next i

    PlaceCallouts = n
End Function
```

`MSO_SHAPE_OVAL = 9` must be declared as `Public Const MSO_SHAPE_OVAL As Long = 9` alongside the module's other constants at the top (VBA requires module-level `Const` declarations before the first procedure) - fix the placement in the same step rather than leaving it mid-module.

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_connector_page.py -v -k callouts
```

Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modConnectorPage.bas tests/test_connector_page.py
git commit -m "feat: draw one numbered oval callout per pin"
```

---

### Task 4: Leader lines

**Files:**
- Modify: `src/vba/modConnectorPage.bas`
- Test: `tests/test_connector_page.py` (additions)

**Interfaces:**
- Consumes: `modPinEditor.MarkerSitsOnAnchor`, `MarkerTopLeft` (existing).
- Produces: VBA `modConnectorPage.PlaceLeaderLines(wsPage As Worksheet, shpPhoto As Shape, vPins As Variant)` - for every pin whose marker (`LabelX`/`LabelY`) does not sit on its anchor (`NormX`/`NormY`), draws a thin line named `"LEADER_<n>"` from the oval's edge (nearest the anchor) to the anchor point. Draws nothing for a pin whose marker sits on its anchor - this is the leader line `docs/superpowers/plans/phase-2-manual-verification.md`'s 2b section recorded as deferred in the editor's live preview; this plan renders it for the first time, on the printed page rather than the on-screen preview.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_connector_page.py`:

```python
def test_leader_line_drawn_only_when_marker_is_pulled_off_anchor(wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        run(wb, "modConnectorPage.PlacePhoto", ws, str(photo_path))
        shp = ws.Shapes("PAGE_PHOTO")

        pins = (
            ("DTM-04P", 1, "", 0.1, 0.1, 0.1, 0.1),   # marker on anchor: no leader
            ("DTM-04P", 2, "", 0.9, 0.1, 0.3, 0.6),   # marker pulled away: leader
        )
        run(wb, "modConnectorPage.PlaceCallouts", ws, shp, pins)
        run(wb, "modConnectorPage.PlaceLeaderLines", ws, shp, pins)

        names = [ws.Shapes(i + 1).Name for i in range(ws.Shapes.Count)]
        assert "LEADER_1" not in names
        assert "LEADER_2" in names
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_connector_page.py -v -k leader
```

Expected: FAIL - `PlaceLeaderLines` does not exist.

- [ ] **Step 3: Write the function**

Append to `src/vba/modConnectorPage.bas`:

```vb
Public Sub PlaceLeaderLines(wsPage As Worksheet, shpPhoto As Shape, vPins As Variant)
    Dim i As Long, nPinNumber As Long
    Dim dAnchorX As Double, dAnchorY As Double, dLabelX As Double, dLabelY As Double
    Dim vAnchorPt As Variant, vMarkerPt As Variant
    Dim dMarkerCx As Double, dMarkerCy As Double, dAnchorCx As Double, dAnchorCy As Double
    Dim dDx As Double, dDy As Double, dDist As Double, dStartX As Double, dStartY As Double
    Dim ln As Shape

    If IsEmpty(vPins) Then Exit Sub

    For i = LBound(vPins, 1) To UBound(vPins, 1)
        nPinNumber = CLng(vPins(i, modLibrary.PIN_COL_PINNUM))
        dAnchorX = CDbl(vPins(i, modLibrary.PIN_COL_NORMX))
        dAnchorY = CDbl(vPins(i, modLibrary.PIN_COL_NORMY))
        dLabelX = CDbl(vPins(i, modLibrary.PIN_COL_LABELX))
        dLabelY = CDbl(vPins(i, modLibrary.PIN_COL_LABELY))

        If modPinEditor.MarkerSitsOnAnchor(dAnchorX, dAnchorY, dLabelX, dLabelY) Then GoTo NextPin

        vAnchorPt = modPinEditor.MarkerTopLeft(dAnchorX, dAnchorY, _
            shpPhoto.Left, shpPhoto.Top, shpPhoto.Width, shpPhoto.Height, 0, 0)
        vMarkerPt = modPinEditor.MarkerTopLeft(dLabelX, dLabelY, _
            shpPhoto.Left, shpPhoto.Top, shpPhoto.Width, shpPhoto.Height, 0, 0)

        dAnchorCx = vAnchorPt(0): dAnchorCy = vAnchorPt(1)
        dMarkerCx = vMarkerPt(0): dMarkerCy = vMarkerPt(1)

        dDx = dAnchorCx - dMarkerCx
        dDy = dAnchorCy - dMarkerCy
        dDist = Sqr(dDx * dDx + dDy * dDy)
        If dDist > 0 Then
            dStartX = dMarkerCx + (dDx / dDist) * (CONN_OVAL_DIAMETER / 2)
            dStartY = dMarkerCy + (dDy / dDist) * (CONN_OVAL_DIAMETER / 2)
        Else
            dStartX = dMarkerCx
            dStartY = dMarkerCy
        End If

        Set ln = wsPage.Shapes.AddLine(dStartX, dStartY, dAnchorCx, dAnchorCy)
        ln.Name = "LEADER_" & CStr(nPinNumber)
        ln.Line.Weight = 0.75

NextPin:
    Next i
End Sub
```

`MarkerTopLeft(..., 0, 0)` with a zero-size box returns the exact center point (a box of width/height 0 offset by half of 0 is the point itself) - reused here to get a plain pixel point from a normalized one, rather than duplicating that arithmetic.

- [ ] **Step 4: Run the test to verify it passes**

Run:

```bash
python -m pytest tests/test_connector_page.py -v -k leader
```

Expected: 1 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modConnectorPage.bas tests/test_connector_page.py
git commit -m "feat: draw a leader line for any pin whose marker was pulled off its anchor"
```

---

### Task 5: Pin-table skeleton and the load-time metadata cell

**Files:**
- Modify: `src/vba/modConnectorPage.bas`
- Test: `tests/test_connector_page.py` (additions)

**Interfaces:**
- Consumes: nothing new.
- Produces:
  - VBA `modConnectorPage.WriteTableSkeleton(wsPage As Worksheet, vPins As Variant)` - writes the 8-column header row (`Pin, Label, Wire To, Signal, Color, AWG, Termination, Length`) at `CONN_TABLE_HEADER_ROW`, then one row per pin starting at `CONN_TABLE_FIRST_ROW` with the static `Pin` and `Label` values filled in (columns `CONN_TABLE_FIRST_COL`/`+1`) and the remaining six columns left blank for 3c.
  - VBA `modConnectorPage.WriteMetadata(wsPage As Worksheet, ByVal sConnectorID As String)` - writes `sConnectorID` to `(1, CONN_META_COL)` and hides that column, so it never prints regardless of 3d's print-area choice and is invisible on screen without being very-hidden (a plain hidden column, matching how the harness sheet's join-key columns are hidden rather than very-hidden, since both exist to be read back programmatically, not to be secret).

- [ ] **Step 1: Write the failing test**

Append to `tests/test_connector_page.py`:

```python
def test_write_table_skeleton_writes_headers_and_static_columns(wb, app):
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        pins = (
            ("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1),
            ("DTM-04P", 2, "GND", 0.9, 0.1, 0.9, 0.1),
        )
        run(wb, "modConnectorPage.WriteTableSkeleton", ws, pins)

        assert ws.Cells(1, 10).Value == "Pin"
        assert ws.Cells(1, 17).Value == "Length"
        assert ws.Cells(2, 10).Value == 1
        assert ws.Cells(2, 11).Value == "+12V"
        assert ws.Cells(3, 10).Value == 2
        assert ws.Cells(3, 11).Value == "GND"
        assert ws.Cells(2, 12).Value is None  # Wire To left for 3c
    finally:
        dest.Close(SaveChanges=False)


def test_write_metadata_hides_the_connector_id_column(wb, app):
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        run(wb, "modConnectorPage.WriteMetadata", ws, "DTM-04P")
        assert ws.Cells(1, 27).Value == "DTM-04P"
        assert ws.Columns(27).Hidden is True
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
python -m pytest tests/test_connector_page.py -v -k "skeleton or metadata"
```

Expected: FAIL - neither function exists.

- [ ] **Step 3: Write the functions**

Append to `src/vba/modConnectorPage.bas`:

```vb
Private Const TABLE_HEADERS As String = "Pin,Label,Wire To,Signal,Color,AWG,Termination,Length"

Public Sub WriteTableSkeleton(wsPage As Worksheet, vPins As Variant)
    Dim vHeaders As Variant, i As Long, r As Long
    Dim cel As Range

    vHeaders = Split(TABLE_HEADERS, ",")
    For i = LBound(vHeaders) To UBound(vHeaders)
        Set cel = wsPage.Cells(CONN_TABLE_HEADER_ROW, CONN_TABLE_FIRST_COL + i)
        cel.Value = vHeaders(i)
        cel.Font.Bold = True
        cel.Interior.Color = 0xD9D9D9
    Next i

    If IsEmpty(vPins) Then Exit Sub

    For i = LBound(vPins, 1) To UBound(vPins, 1)
        r = CONN_TABLE_FIRST_ROW + (i - LBound(vPins, 1))
        wsPage.Cells(r, CONN_TABLE_FIRST_COL).Value = CLng(vPins(i, modLibrary.PIN_COL_PINNUM))
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 1).Value = CStr(vPins(i, modLibrary.PIN_COL_LABEL))
    Next i
End Sub

Public Sub WriteMetadata(wsPage As Worksheet, ByVal sConnectorID As String)
    wsPage.Cells(1, CONN_META_COL).Value = sConnectorID
    wsPage.Columns(CONN_META_COL).Hidden = True
End Sub
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_connector_page.py -v -k "skeleton or metadata"
```

Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modConnectorPage.bas tests/test_connector_page.py
git commit -m "feat: add the pin-table skeleton and the connector-page metadata cell"
```

---

### Task 6: Orchestrate one page per instance inside Save

**Files:**
- Modify: `src/vba/modHarnessBuild.bas`
- Modify: `src/vba/modHarnessActions.bas`
- Modify: `build/build.py`
- Modify: `tests/test_layering.py`
- Test: `tests/test_harness_build.py` (additions)

**Interfaces:**
- Consumes: `modConnectors.AllInstances` (Task 1), `modConnectorPage.PagePhotoPath`/`PlacePhoto`/`PlaceCallouts`/`PlaceLeaderLines`/`WriteTableSkeleton`/`WriteMetadata` (Tasks 2-5), `modLibrary.ReadPinsForConnector` (existing), `modSnapshot.LibraryFolder`/`SNAP_PINS_FIRST_ROW`/`LAST_ROW` (existing).
- Produces: VBA `modHarnessBuild.BuildConnectorPages(destWb As Workbook, wsSnapshot As Worksheet)` - one `CONN_<RefDes>` sheet per row of `modConnectors.AllInstances()`.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_harness_build.py`:

```python
def test_build_connector_pages_creates_one_sheet_per_instance(wb, app, tmp_path):
    conn_ws = wb.Worksheets("Connectors")
    conn_ws.Cells(2, 1).Value = "J1"
    conn_ws.Cells(2, 2).Value = "DTM-04P"
    conn_ws.Cells(2, 3).Value = "Deutsch DTM 4-way"
    conn_ws.Cells(2, 5).Value = "Connector"
    conn_ws.Cells(2, 6).Value = 2

    wsSnap = wb.Worksheets("_Snapshot")
    fields = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
              2, "", "PHOTO_DTM-04P", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", wsSnap, 2, 201, fields)
    run(wb, "modLibrary.WritePin", wsSnap, 211, 2210, ("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1))
    run(wb, "modLibrary.WritePin", wsSnap, 211, 2210, ("DTM-04P", 2, "GND", 0.9, 0.1, 0.9, 0.1))
    photo_path = write_sample_photo(tmp_path / "photo.png")
    run(wb, "modLibrary.EmbedConnectorPhoto", wsSnap, "DTM-04P", str(photo_path))

    dest = app.Workbooks.Add()
    try:
        run(wb, "modHarnessBuild.BuildHarnessSheets", dest)
        run(wb, "modHarnessBuild.CopySnapshot", dest.Worksheets("_Snapshot"))
        run(wb, "modHarnessBuild.BuildConnectorPages", dest, dest.Worksheets("_Snapshot"))

        page = dest.Worksheets("CONN_J1")
        assert page.Cells(1, 27).Value == "DTM-04P"
        assert page.Cells(1, 10).Value == "Pin"
        assert page.Shapes.Count >= 2  # at least the two ovals (photo depends on cache existing)
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_harness_build.py -v -k connector_pages
```

Expected: FAIL - `BuildConnectorPages` does not exist.

- [ ] **Step 3: Write the function**

Append to `src/vba/modHarnessBuild.bas`:

```vb
Public Sub BuildConnectorPages(destWb As Workbook, wsSnapshot As Worksheet)
    Dim vInstances As Variant, i As Long
    Dim sRefDes As String, sConnectorID As String
    Dim wsPage As Worksheet, vPins As Variant, shpPhoto As Shape
    Dim sPhotoPath As String

    vInstances = modConnectors.AllInstances()
    If IsEmpty(vInstances) Then Exit Sub

    For i = LBound(vInstances, 1) To UBound(vInstances, 1)
        sRefDes = CStr(vInstances(i, 1))
        sConnectorID = CStr(vInstances(i, 2))

        Set wsPage = destWb.Worksheets.Add(After:=destWb.Worksheets(destWb.Worksheets.Count))
        wsPage.Name = "CONN_" & sRefDes

        vPins = modLibrary.ReadPinsForConnector(wsSnapshot, modSnapshot.SNAP_PINS_FIRST_ROW, _
            modSnapshot.SNAP_PINS_LAST_ROW, sConnectorID)

        sPhotoPath = modConnectorPage.PagePhotoPath(modSnapshot.LibraryFolder(), sConnectorID)
        If modConnectorPage.PlacePhoto(wsPage, sPhotoPath) Then
            Set shpPhoto = wsPage.Shapes("PAGE_PHOTO")
            modConnectorPage.PlaceCallouts wsPage, shpPhoto, vPins
            modConnectorPage.PlaceLeaderLines wsPage, shpPhoto, vPins
        End If

        modConnectorPage.WriteTableSkeleton wsPage, vPins
        modConnectorPage.WriteMetadata wsPage, sConnectorID
    Next i
End Sub
```

A connector whose photo cache is missing (an elevated-risk case `docs/superpowers/plans/phase-2-manual-verification.md`'s 2c section already flags) still gets its `CONN_<RefDes>` sheet, pin table, and metadata cell - it just has no photo or callouts, rather than failing the whole save. This matches the spec's validation philosophy elsewhere ("the tool being wrong must never stop [a student] from turning in work") even though Check Drawing itself is Phase 4 - a missing photo here is silently incomplete, not fatal, consistent with that spirit.

- [ ] **Step 4: Wire it into SaveHarness**

In `src/vba/modHarnessActions.bas`, add one line after `modHarnessBuild.CopySnapshot wsSnapshot`:

```vb
    modHarnessBuild.CopySnapshot wsSnapshot
    modHarnessBuild.BuildConnectorPages destWb, wsSnapshot
```

- [ ] **Step 5: Register the new module**

In `build/build.py`, add `"modConnectorPage.bas"` to `VBA_MODULES` (after `"modHarnessBuild.bas"`).

In `tests/test_layering.py`, add `"modConnectorPage"` to `LAYER0`.

- [ ] **Step 6: Run the tests to verify they pass**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: everything passes, including the new `test_build_connector_pages_creates_one_sheet_per_instance`.

- [ ] **Step 7: Commit**

```bash
git add src/vba/modHarnessBuild.bas src/vba/modHarnessActions.bas build/build.py tests/test_layering.py tests/test_harness_build.py
git commit -m "feat: render one connector page per instance during Save"
```

---

### Task 7: Extend the round-trip integration test

**Files:**
- Modify: `tests/test_harness_save_integration.py`

**Interfaces:**
- Consumes: everything from Tasks 1-6, plus the fixture data `test_full_harness_round_trips_through_a_saved_file` (3a, Task 7) already sets up (one `J1`/`J2` wire against a `DTM-04P` `_Snapshot` entry with one pin placed).

- [ ] **Step 1: Add connector-page assertions to the existing test**

In `tests/test_harness_save_integration.py`, extend `test_full_harness_round_trips_through_a_saved_file`'s `Connectors` sheet setup (so a `CONN_J1` page actually gets built) and its post-reopen assertions:

```python
    wsConn = wb.Worksheets("Connectors")
    wsConn.Cells(2, 1).Value = "J1"
    wsConn.Cells(2, 2).Value = "DTM-04P"
    wsConn.Cells(2, 3).Value = "Deutsch DTM 4-way"
    wsConn.Cells(2, 5).Value = "Connector"
    wsConn.Cells(2, 6).Value = 4
```

(Insert this block after the existing `wsSnap`/photo setup, before `dest = app.Workbooks.Add()`.)

Add after the existing `_Snapshot` assertions, still inside the `reopened` block:

```python
        page = reopened.Worksheets("CONN_J1")
        assert page.Cells(1, 27).Value == "DTM-04P"  # metadata cell, for 3e
        assert page.Columns(27).Hidden is True
        assert page.Shapes("PIN_1").Name == "PIN_1"
        assert page.Cells(1, 10).Value == "Pin"
        assert page.Cells(2, 10).Value == 1
        assert page.Cells(2, 11).Value == "+12V"
```

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
git commit -m "test: extend the harness round-trip test to cover connector pages"
```

---

## Self-Review

**Spec coverage for this sub-plan.** Covers the spec's "Connector page layout" section in full except the pin table's live formulas (explicitly 3c) and page setup (explicitly 3d): fixed photo anchor scaled by aspect ratio, one oval per pin at its stored marker position, and a leader line exactly when the marker no longer covers its anchor - all three traced to specific functions above and to `modPinEditor`'s existing, already-tested geometry (`FitAspectRatio`, `MarkerTopLeft`, `MarkerSitsOnAnchor`) rather than new math duplicating it.

**Why this plan reuses `modPinEditor`'s geometry functions but not its worksheet-level functions (`PinGeometry`, `NeedsLeaderLine`, `FindPinRow`).** Traced directly against the source: `FindPinRow` (and everything built on it) hardcodes the `_Edit` scratch sheet's row window (`SCRATCH_FIRST_ROW`/`LAST_ROW`, 2-2000), not `_Snapshot`'s (`SNAP_PINS_FIRST_ROW`/`LAST_ROW`, 211-2210) - calling them against `_Snapshot` would silently scan the wrong rows. `modLibrary.ReadPinsForConnector` already reads the correct window and returns every field a caller needs, so this plan reads pins that way and passes the plain `Double`s into `MarkerTopLeft`/`MarkerSitsOnAnchor`, which take no worksheet at all and are window-agnostic by construction.

**Why the metadata cell exists and why it is a plain hidden column, not very-hidden.** Flagged as a load-bearing decision in Global Constraints: without it, 3e's Load has no way to recover which library `ConnectorID` each `CONN_<RefDes>` sheet renders, since the spec's own list of what a saved harness contains has no separate instance table. A plain hidden column (rather than very-hidden, which would require unhiding the whole sheet to reach it) is enough - nothing about it needs to be hard to find, only excluded from the printed page, which `Columns(...).Hidden = True` already guarantees regardless of whatever print area 3d sets.

**A connector with no cached photo does not fail the whole save.** Traced directly to `BuildConnectorPages`: `PlacePhoto` returning `False` skips only the photo/callout/leader steps for that one instance, not the page's creation or its pin table and metadata cell - consistent with the "elevated risk for a connector with no cache" case `phase-2-manual-verification.md` already documents as a real, encountered condition, and with the spec's broader stance that an imperfect artifact must never block a student from having something to turn in.

**Type consistency.** `PlaceCallouts`/`PlaceLeaderLines`/`WriteTableSkeleton` all take the same `vPins` shape - the 2D, 1-based array `modLibrary.ReadPinsForConnector` returns, columns addressed via `modLibrary.PIN_COL_*` everywhere, never a re-derived index. `BuildConnectorPages` passes `vPins` to all three unchanged, computed once per instance rather than re-read per function.

**No placeholders.** Every step contains complete VBA - the exact oval/leader geometry, the exact header list, the exact metadata cell address - no "compute the position" or "similar to the callout above" stand-ins.
