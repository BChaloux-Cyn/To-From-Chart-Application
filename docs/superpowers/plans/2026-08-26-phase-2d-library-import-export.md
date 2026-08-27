# Phase 2d: Library Import and Export Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move one connector's full definition - record, pins, and photo - between two library-schema workbooks, so a connector defined by one student can be merged into another's library, with a collision-safe rename and a documented fallback when photo extraction fails.

**Architecture:** `modLibraryTransfer.bas` composes 2a's already-tested `modLibrary` functions - it never touches a cell directly. The one genuinely new capability is copying an *already-embedded* picture shape between two workbooks, which VBA has no direct `Shape.Export`-to-another-sheet method for; the standard technique is `Shape.Copy` + `Worksheet.Paste`, which goes through the Windows clipboard. That is exactly the operation the spec's own "Photo cache" section anticipates can fail ("If extraction fails, the editor prompts for the image file rather than failing outright") - so it is designed from the start to return a clear success/failure signal rather than raising, with the UI layer (not a logic module) responsible for the fallback prompt.

**Tech Stack:** Python 3.13, pywin32, pytest, Excel 16.0 COM automation, VBA.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Depends on:** 2a (`modLibrary`), 2c (`frmManageLibrary`).

**Part of Phase 2** (2a, 2b, 2c done; this is 2d, the last implementation sub-plan). Followed by 2e (design docs + student user guide).

## Global Constraints

Copied verbatim from the spec, plus 2a/2b/2c's. Every task's requirements implicitly include these.

- Every VBA module starts with `Option Explicit`. No `MsgBox`/dialog in a logic module.
- ConnectorID: unique slug, numeric suffix appended on collision (2a's `UniqueConnectorID`, reused unchanged).
- Origin field: `Local`, or the filename of the library it was imported from.
- "If extraction fails, the editor prompts for the image file rather than failing outright" - a logic-module function reports failure; only a UI module may prompt.
- Custom VBA `Type`s never cross the `Application.Run` boundary.

## File Structure

| File | Responsibility |
|---|---|
| `src/vba/modLibraryTransfer.bas` | `ExportConnector`, `ImportConnector`, `CopyConnectorPhoto` |
| `src/vba/forms/frmManageLibrary.evt` | Modified: `cmdImport_Click`, `cmdExport_Click` added |
| `build/build.py` | Modified: adds `modLibraryTransfer.bas` |
| `tests/conftest.py` | Modified: `build/` added to `sys.path` so tests can build a throwaway second library workbook in-process |
| `tests/test_library_transfer_export.py` | `ExportConnector` |
| `tests/test_library_transfer_import.py` | `ImportConnector`, collision renaming |
| `tests/test_library_transfer_photo.py` | `CopyConnectorPhoto` |
| `tests/test_manage_library_transfer_wiring.py` | Structural: Import/Export click handlers exist and call the tested functions |

---

### Task 1: Export a connector

**Files:**
- Create: `src/vba/modLibraryTransfer.bas`
- Modify: `build/build.py`
- Modify: `tests/conftest.py`
- Test: `tests/test_library_transfer_export.py`

**Interfaces:**
- Consumes: `modLibrary.ReadConnector`, `WriteConnector`, `ReadPinsForConnector`, `WritePin`, `LIB_ROW_CAP`, `PIN_FIELD_COUNT`, `RemoveConnectorPhoto`.
- Produces:
  - VBA `modLibraryTransfer.CopyConnectorPhoto(wsSrcPhotos, wsDestPhotos, sSrcConnectorID, sDestConnectorID) As Boolean`
  - VBA `modLibraryTransfer.ExportConnector(wsSrcConn, wsSrcPins, wsSrcPhotos, wsDestConn, wsDestPins, wsDestPhotos, sConnectorID) As Boolean`

- [ ] **Step 1: Add `build/` to the test path**

In `tests/conftest.py`, near the top (after the existing `ROOT`/`ARTIFACT` constants):

```python
sys.path.insert(0, str(ROOT / "build"))
```

This lets test files `import library_layout` directly to build a throwaway second library workbook in-process, the same way `build.py` itself imports its sibling modules - no test needs the built `dist/ConnectorLibrary.xlsx` artifact to exist twice on disk.

- [ ] **Step 2: Write the failing test**

Create `tests/test_library_transfer_export.py`:

```python
import library_layout

from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_export_connector_copies_record_pins_and_photo(wb, library_wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws_src_conn = library_wb.Worksheets("Connectors")
    ws_src_pins = library_wb.Worksheets("Pins")
    ws_src_photos = library_wb.Worksheets("Photos")

    fields = ("DTM-04P", "Deutsch DTM 4-way", "Deutsch", "DTM06-4S", "Connector",
              2, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws_src_conn, 2, 100000, fields)
    run(wb, "modLibrary.WritePin", ws_src_pins, 2, 100000, ("DTM-04P", 1, "A", 0.1, 0.1, 0.1, 0.1))
    run(wb, "modLibrary.EmbedConnectorPhoto", ws_src_photos, "DTM-04P", str(photo_path))

    dest_wb = app.Workbooks.Add()
    try:
        library_layout.build_library_sheets(dest_wb)
        ws_dest_conn = dest_wb.Worksheets("Connectors")
        ws_dest_pins = dest_wb.Worksheets("Pins")
        ws_dest_photos = dest_wb.Worksheets("Photos")

        ok = run(wb, "modLibraryTransfer.ExportConnector", ws_src_conn, ws_src_pins, ws_src_photos,
                 ws_dest_conn, ws_dest_pins, ws_dest_photos, "DTM-04P")
        assert ok is True

        result = run(wb, "modLibrary.ReadConnector", ws_dest_conn, 2, 100000, "DTM-04P")
        assert result[1] == "Deutsch DTM 4-way"
        pins = run(wb, "modLibrary.ReadPinsForConnector", ws_dest_pins, 2, 100000, "DTM-04P")
        assert len(pins) == 1
        assert ws_dest_photos.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"
    finally:
        dest_wb.Close(SaveChanges=False)


def test_export_unknown_connector_returns_false(wb, library_wb, app):
    dest_wb = app.Workbooks.Add()
    try:
        library_layout.build_library_sheets(dest_wb)
        ok = run(
            wb, "modLibraryTransfer.ExportConnector",
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"),
            dest_wb.Worksheets("Connectors"), dest_wb.Worksheets("Pins"), dest_wb.Worksheets("Photos"),
            "NO-SUCH-ID",
        )
        assert ok is False
    finally:
        dest_wb.Close(SaveChanges=False)
```

- [ ] **Step 3: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_library_transfer_export.py -v
```

Expected: FAIL — `modLibraryTransfer` does not exist.

- [ ] **Step 4: Write the module**

Create `src/vba/modLibraryTransfer.bas`:

```vb
Attribute VB_Name = "modLibraryTransfer"
Option Explicit

Public Function CopyConnectorPhoto(wsSrcPhotos As Worksheet, wsDestPhotos As Worksheet, _
                                   ByVal sSrcConnectorID As String, ByVal sDestConnectorID As String) As Boolean
    On Error GoTo Failed
    modLibrary.RemoveConnectorPhoto wsDestPhotos, sDestConnectorID
    wsSrcPhotos.Shapes("PHOTO_" & sSrcConnectorID).Copy
    wsDestPhotos.Paste
    wsDestPhotos.Shapes(wsDestPhotos.Shapes.Count).Name = "PHOTO_" & sDestConnectorID
    CopyConnectorPhoto = True
    Exit Function
Failed:
    CopyConnectorPhoto = False
End Function

Public Function ExportConnector(wsSrcConn As Worksheet, wsSrcPins As Worksheet, wsSrcPhotos As Worksheet, _
                                wsDestConn As Worksheet, wsDestPins As Worksheet, wsDestPhotos As Worksheet, _
                                ByVal sConnectorID As String) As Boolean
    Dim vFields As Variant
    vFields = modLibrary.ReadConnector(wsSrcConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vFields) Then Exit Function
    If Not modLibrary.WriteConnector(wsDestConn, 2, modLibrary.LIB_ROW_CAP, vFields) Then Exit Function

    Dim vPins As Variant, i As Long, j As Long, vRow(1 To 7) As Variant
    vPins = modLibrary.ReadPinsForConnector(wsSrcPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If Not IsEmpty(vPins) Then
        For i = LBound(vPins, 1) To UBound(vPins, 1)
            For j = 1 To modLibrary.PIN_FIELD_COUNT
                vRow(j) = vPins(i, j)
            Next j
            modLibrary.WritePin wsDestPins, 2, modLibrary.LIB_ROW_CAP, vRow
        Next i
    End If

    ExportConnector = CopyConnectorPhoto(wsSrcPhotos, wsDestPhotos, sConnectorID, sConnectorID)
End Function
```

- [ ] **Step 5: Wire the module into the build**

In `build/build.py`:

```python
VBA_MODULES = [
    "modUtil.bas", "modState.bas", "modConnectors.bas", "modChart.bas",
    "modLibrary.bas", "modPinEditor.bas", "clsPinMarker.cls", "modSnapshot.bas",
    "modLibraryTransfer.bas",
]
```

- [ ] **Step 6: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_library_transfer_export.py -v
```

Expected: 2 passed.

**If `CopyConnectorPhoto` fails in this environment** (a `Shape.Copy`/`Paste` clipboard operation, run headless, is the one piece of this whole codebase not proven reliable end to end): do not force it through. Confirm by running the same two lines interactively in the VBE Immediate window against a visible Excel instance. If it fails there too, this is a real environment limitation the spec already anticipated - note it plainly in the commit message and let Task 3's fallback prompt (which exists for exactly this reason) carry the feature, rather than spending further effort trying to make headless clipboard automation reliable.

- [ ] **Step 7: Commit**

```bash
git add tests/conftest.py src/vba/modLibraryTransfer.bas build/build.py tests/test_library_transfer_export.py
git commit -m "feat: export a connector's record, pins, and photo to another library"
```

---

### Task 2: Import a connector with collision-safe renaming

**Files:**
- Modify: `src/vba/modLibraryTransfer.bas`
- Test: `tests/test_library_transfer_import.py`

**Interfaces:**
- Consumes: `modLibrary.UniqueConnectorID`, `modLibraryTransfer.CopyConnectorPhoto` (Task 1).
- Produces: VBA `modLibraryTransfer.ImportConnector(wsSrcConn, wsSrcPins, wsSrcPhotos, wsDestConn, wsDestPins, wsDestPhotos, sConnectorID, sOriginFileName) As String` — returns the ConnectorID actually used in the destination (renamed on collision), or `""` on failure.

The photo copy's own success or failure is intentionally not folded into this function's return value - the caller (Task 3's UI) checks `CopyConnectorPhoto`'s result separately, since a successful record import with a failed photo copy is a real, distinct outcome the spec's fallback prompt exists to handle.

- [ ] **Step 1: Write the failing test**

Create `tests/test_library_transfer_import.py`:

```python
import library_layout

from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def seed_source_connector(wb, app, tmp_path, connector_id="DTM-04P", name="Deutsch DTM 4-way"):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    src_wb = app.Workbooks.Add()
    library_layout.build_library_sheets(src_wb)
    ws_conn = src_wb.Worksheets("Connectors")
    ws_pins = src_wb.Worksheets("Pins")
    ws_photos = src_wb.Worksheets("Photos")

    fields = (connector_id, name, "Deutsch", "DTM06-4S", "Connector",
              1, "", "", "2026-08-26T00:00:00Z", "2026-08-26T00:00:00Z", "Local")
    run(wb, "modLibrary.WriteConnector", ws_conn, 2, 100000, fields)
    run(wb, "modLibrary.WritePin", ws_pins, 2, 100000, (connector_id, 1, "A", 0.1, 0.1, 0.1, 0.1))
    run(wb, "modLibrary.EmbedConnectorPhoto", ws_photos, connector_id, str(photo_path))
    return src_wb, ws_conn, ws_pins, ws_photos


def test_import_with_no_collision_keeps_its_id(wb, library_wb, app, tmp_path):
    src_wb, ws_src_conn, ws_src_pins, ws_src_photos = seed_source_connector(wb, app, tmp_path)
    try:
        dest_id = run(
            wb, "modLibraryTransfer.ImportConnector", ws_src_conn, ws_src_pins, ws_src_photos,
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"),
            "DTM-04P", "other_library.xlsx",
        )
        assert dest_id == "DTM-04P"

        result = run(wb, "modLibrary.ReadConnector", library_wb.Worksheets("Connectors"), 2, 100000, "DTM-04P")
        assert result[10] == "other_library.xlsx"
    finally:
        src_wb.Close(SaveChanges=False)


def test_import_colliding_with_a_local_id_is_renamed_and_the_original_kept(wb, library_wb, app, tmp_path):
    ws_dest_conn = library_wb.Worksheets("Connectors")
    existing = ("DTM-04P", "A different local part", "", "", "Connector", 1, "", "", "", "", "Local")
    run(wb, "modLibrary.WriteConnector", ws_dest_conn, 2, 100000, existing)

    src_wb, ws_src_conn, ws_src_pins, ws_src_photos = seed_source_connector(wb, app, tmp_path)
    try:
        dest_id = run(
            wb, "modLibraryTransfer.ImportConnector", ws_src_conn, ws_src_pins, ws_src_photos,
            ws_dest_conn, library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"),
            "DTM-04P", "other_library.xlsx",
        )

        assert dest_id == "DTM-04P-2"
        original = run(wb, "modLibrary.ReadConnector", ws_dest_conn, 2, 100000, "DTM-04P")
        assert original[1] == "A different local part"
        imported = run(wb, "modLibrary.ReadConnector", ws_dest_conn, 2, 100000, "DTM-04P-2")
        assert imported[1] == "Deutsch DTM 4-way"
    finally:
        src_wb.Close(SaveChanges=False)


def test_import_pins_are_rewritten_under_the_renamed_id(wb, library_wb, app, tmp_path):
    ws_dest_conn = library_wb.Worksheets("Connectors")
    existing = ("DTM-04P", "A different local part", "", "", "Connector", 1, "", "", "", "", "Local")
    run(wb, "modLibrary.WriteConnector", ws_dest_conn, 2, 100000, existing)

    src_wb, ws_src_conn, ws_src_pins, ws_src_photos = seed_source_connector(wb, app, tmp_path)
    try:
        dest_id = run(
            wb, "modLibraryTransfer.ImportConnector", ws_src_conn, ws_src_pins, ws_src_photos,
            ws_dest_conn, library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"),
            "DTM-04P", "other_library.xlsx",
        )

        pins = run(wb, "modLibrary.ReadPinsForConnector", library_wb.Worksheets("Pins"), 2, 100000, dest_id)
        assert len(pins) == 1
        assert pins[0][0] == dest_id
    finally:
        src_wb.Close(SaveChanges=False)


def test_import_unknown_source_id_returns_empty_string(wb, library_wb, app):
    src_wb = app.Workbooks.Add()
    try:
        library_layout.build_library_sheets(src_wb)
        result = run(
            wb, "modLibraryTransfer.ImportConnector",
            src_wb.Worksheets("Connectors"), src_wb.Worksheets("Pins"), src_wb.Worksheets("Photos"),
            library_wb.Worksheets("Connectors"), library_wb.Worksheets("Pins"), library_wb.Worksheets("Photos"),
            "NO-SUCH-ID", "other.xlsx",
        )
        assert result == ""
    finally:
        src_wb.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_library_transfer_import.py -v
```

Expected: FAIL — `ImportConnector` does not exist.

- [ ] **Step 3: Add the function**

Append to `src/vba/modLibraryTransfer.bas`:

```vb
Public Function ImportConnector(wsSrcConn As Worksheet, wsSrcPins As Worksheet, wsSrcPhotos As Worksheet, _
                                wsDestConn As Worksheet, wsDestPins As Worksheet, wsDestPhotos As Worksheet, _
                                ByVal sConnectorID As String, ByVal sOriginFileName As String) As String
    Dim vFields As Variant
    vFields = modLibrary.ReadConnector(wsSrcConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If IsEmpty(vFields) Then Exit Function

    Dim sDestID As String
    sDestID = modLibrary.UniqueConnectorID(wsDestConn, 2, modLibrary.LIB_ROW_CAP, sConnectorID)

    vFields(0) = sDestID
    vFields(7) = "PHOTO_" & sDestID
    vFields(10) = sOriginFileName
    If Not modLibrary.WriteConnector(wsDestConn, 2, modLibrary.LIB_ROW_CAP, vFields) Then Exit Function

    Dim vPins As Variant, i As Long, j As Long, vRow(1 To 7) As Variant
    vPins = modLibrary.ReadPinsForConnector(wsSrcPins, 2, modLibrary.LIB_ROW_CAP, sConnectorID)
    If Not IsEmpty(vPins) Then
        For i = LBound(vPins, 1) To UBound(vPins, 1)
            For j = 1 To modLibrary.PIN_FIELD_COUNT
                vRow(j) = vPins(i, j)
            Next j
            vRow(1) = sDestID
            modLibrary.WritePin wsDestPins, 2, modLibrary.LIB_ROW_CAP, vRow
        Next i
    End If

    CopyConnectorPhoto wsSrcPhotos, wsDestPhotos, sConnectorID, sDestID

    ImportConnector = sDestID
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_library_transfer_import.py -v
```

Expected: 4 passed.

- [ ] **Step 5: Run the whole suite**

Run:

```bash
python -m pytest -v
```

Expected: all passed (239 from 2c + 2 from Task 1 + 4 from Task 2 = 245).

- [ ] **Step 6: Commit**

```bash
git add src/vba/modLibraryTransfer.bas tests/test_library_transfer_import.py
git commit -m "feat: import a connector, renaming on ID collision"
```

---

### Task 3: Wire Import/Export into the library browser

**Files:**
- Modify: `src/vba/forms/frmManageLibrary.evt`
- Test: `tests/test_manage_library_transfer_wiring.py`

**Interfaces:**
- Consumes: `modLibraryTransfer.ExportConnector`, `ImportConnector`.
- Produces: `cmdImport_Click`, `cmdExport_Click` on `frmManageLibrary`.

This is the second (and last) piece of this codebase pytest cannot exercise end to end - `Application.GetOpenFilename`/`GetSaveAsFilename` block on real user interaction. Tests here stay structural, matching 2b Task 8 and 2c Task 5's precedent.

- [ ] **Step 1: Write the failing test**

Create `tests/test_manage_library_transfer_wiring.py`:

```python
def module_source(wb, component_name):
    module = wb.VBProject.VBComponents(component_name).CodeModule
    return module.Lines(1, module.CountOfLines)


def test_export_click_calls_export_connector(wb):
    assert "modLibraryTransfer.ExportConnector" in module_source(wb, "frmManageLibrary")


def test_import_click_calls_import_connector(wb):
    assert "modLibraryTransfer.ImportConnector" in module_source(wb, "frmManageLibrary")


def test_import_click_prompts_for_a_replacement_image_on_photo_failure(wb):
    source = module_source(wb, "frmManageLibrary")
    assert "CopyConnectorPhoto" in source
    assert "GetOpenFilename" in source
    assert "modLibrary.EmbedConnectorPhoto" in source
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_manage_library_transfer_wiring.py -v
```

Expected: FAIL — neither handler exists yet.

- [ ] **Step 3: Add the handlers**

In `src/vba/forms/frmManageLibrary.evt`, add (the file already has `mLibrary`/`mConnectorIDs`/`RefreshList` from 2c):

```vb
Private Sub cmdExport_Click()
    If lstConnectors.ListIndex < 0 Then Exit Sub
    Dim sConnectorID As String
    sConnectorID = mConnectorIDs(lstConnectors.ListIndex + 1)

    Dim sPath As Variant
    sPath = Application.GetSaveAsFilename( _
        InitialFileName:=sConnectorID & "_export.xlsx", _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If sPath = False Then Exit Sub

    Dim destWb As Workbook
    Set destWb = Workbooks.Add
    modLibraryTransfer.BuildExportSheets destWb

    Dim ok As Boolean
    ok = modLibraryTransfer.ExportConnector(mLibrary.Worksheets("Connectors"), mLibrary.Worksheets("Pins"), _
        mLibrary.Worksheets("Photos"), destWb.Worksheets("Connectors"), destWb.Worksheets("Pins"), _
        destWb.Worksheets("Photos"), sConnectorID)

    If ok Then
        destWb.SaveAs Filename:=CStr(sPath), FileFormat:=51
        MsgBox "Exported " & sConnectorID & ".", vbInformation
    Else
        MsgBox "Could not export " & sConnectorID & ".", vbExclamation
    End If
    destWb.Close SaveChanges:=False
End Sub

Private Sub cmdImport_Click()
    Dim sPath As Variant
    sPath = Application.GetOpenFilename(FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If sPath = False Then Exit Sub

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(CStr(sPath))

    Dim sOriginName As String
    sOriginName = srcWb.Name

    Dim nLast As Long, r As Long, sConnectorID As String, sDestID As String
    Dim wsSrcConn As Worksheet
    Set wsSrcConn = srcWb.Worksheets("Connectors")
    nLast = wsSrcConn.Cells(wsSrcConn.Rows.Count, 1).End(xlUp).Row

    For r = 2 To nLast
        sConnectorID = Trim$(CStr(wsSrcConn.Cells(r, modLibrary.LIB_COL_ID).Value))
        If Len(sConnectorID) > 0 Then
            sDestID = modLibraryTransfer.ImportConnector(wsSrcConn, srcWb.Worksheets("Pins"), _
                srcWb.Worksheets("Photos"), mLibrary.Worksheets("Connectors"), mLibrary.Worksheets("Pins"), _
                mLibrary.Worksheets("Photos"), sConnectorID, sOriginName)

            If Len(sDestID) > 0 Then
                ' ImportConnector already attempted the photo copy internally -
                ' calling CopyConnectorPhoto again here would redo a clipboard
                ' operation that may not even be deterministic, and could give
                ' a different answer than the one that actually happened. Check
                ' whether the shape it should have produced exists instead.
                Dim bPhotoExists As Boolean
                bPhotoExists = False
                On Error Resume Next
                bPhotoExists = Not (mLibrary.Worksheets("Photos").Shapes("PHOTO_" & sDestID) Is Nothing)
                On Error GoTo 0

                If Not bPhotoExists Then
                    Dim sReplacement As Variant
                    sReplacement = Application.GetOpenFilename( _
                        "Pictures (*.png; *.jpg; *.jpeg; *.bmp), *.png;*.jpg;*.jpeg;*.bmp", _
                        , "Photo for " & sDestID & " could not be extracted - choose a replacement")
                    If sReplacement <> False Then
                        modLibrary.EmbedConnectorPhoto mLibrary.Worksheets("Photos"), sDestID, CStr(sReplacement)
                    End If
                End If
            End If
        End If
    Next r

    srcWb.Close SaveChanges:=False
    mLibrary.Save
    RefreshList
    MsgBox "Import complete.", vbInformation
End Sub
```

`cmdExport_Click` needs a fresh library-schema workbook to export into, but the form's code-behind cannot `import library_layout` - that is a Python-side, build-time-only module, unreachable from the running `.xlsm`. Add a small VBA-side equivalent to `modLibraryTransfer.bas`:

```vb
Public Sub BuildExportSheets(wb As Workbook)
    Dim names As Variant, i As Long, sheet As Worksheet, original As Worksheet
    names = Array("Connectors", "Pins", "Photos")

    Set original = wb.Worksheets(1)
    For i = LBound(names) To UBound(names)
        Set sheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        sheet.Name = CStr(names(i))
    Next i
    original.Delete

    Dim connHeaders As Variant, pinHeaders As Variant, c As Long
    connHeaders = Array("ConnectorID", "Name", "Manufacturer", "PartNumber", "Type", _
                         "PinCount", "Notes", "PhotoShapeName", "CreatedUtc", "ModifiedUtc", "Origin")
    pinHeaders = Array("ConnectorID", "PinNumber", "PinLabel", "NormX", "NormY", "LabelX", "LabelY")

    For c = LBound(connHeaders) To UBound(connHeaders)
        wb.Worksheets("Connectors").Cells(1, c + 1).Value = connHeaders(c)
    Next c
    For c = LBound(pinHeaders) To UBound(pinHeaders)
        wb.Worksheets("Pins").Cells(1, c + 1).Value = pinHeaders(c)
    Next c
End Sub
```

and use `modLibraryTransfer.BuildExportSheets destWb` in place of `LibraryLayoutBuilder destWb` above. This intentionally duplicates 2a's `library_layout.build_library_sheets` header list rather than sharing it across the Python/VBA boundary - the two are already kept in lockstep by 2a's own structural tests (`test_connectors_sheet_headers`, `test_pins_sheet_headers`) failing if the schema ever drifts, so a VBA-side copy used only for this one runtime export path is a small, contained duplication rather than a maintenance hazard.

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_manage_library_transfer_wiring.py -v
```

Expected: 3 passed.

- [ ] **Step 5: Run the whole suite from a clean build**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: all passed (245 from Task 2 + 3 = 248).

- [ ] **Step 6: Manually verify in Excel**

Open `dist/HarnessCreator.xlsm`, `Manage Library`. Export a connector to a new file and confirm the saved file opens with the correct schema and data. Import that same file back in and confirm it round-trips (or, since it is now present, renames to `-2` and the original is untouched). If a test connector's photo shape is deliberately removed before import to force the extraction-failure path, confirm the replacement-image prompt appears and the substituted photo lands correctly. Record the result in the commit message - matching 2b Task 8 and 2c Task 5's precedent.

Tracked in the consolidated `docs/superpowers/plans/phase-2-manual-verification.md`, along with Task 1's separate clipboard-reliability check above - run both as part of that batch after this sub-plan finishes, before 2e.

- [ ] **Step 7: Commit**

```bash
git add src/vba/forms/frmManageLibrary.evt src/vba/modLibraryTransfer.bas tests/test_manage_library_transfer_wiring.py
git commit -m "feat: wire library import and export into Manage Library"
```

---

## Self-Review

**Spec coverage for this sub-plan.** "Library import and export" from the phase-2 line item is fully covered: `ImportConnector`'s collision-safe renaming and Origin-field stamping match "Libraries can be exported and imported so a connector defined by one student can be merged into another's library" and the Connectors schema's `Origin` field description exactly. The extraction-failure fallback (Task 3) matches "If extraction fails, the editor prompts for the image file rather than failing outright" verbatim - this was the one line in the whole spec that explicitly named a failure mode and its required behavior, and both are implemented, not just the happy path.

**The one deliberately named risk in this plan.** `CopyConnectorPhoto`'s `Shape.Copy`/`Paste` goes through the Windows clipboard - the only clipboard-dependent operation anywhere in this codebase, and clipboard access from a headless, backgrounded Excel COM automation process is not something the rest of this project has needed to prove out. Task 1 Step 6 says explicitly what to do if it turns out unreliable in this environment: verify by hand in a visible Excel instance before concluding it is broken, and if it genuinely is, let it stand as a known limitation the spec already designed around, rather than forcing a workaround.

**Deliberately out of scope, not a gap:** multi-select export (exporting more than one connector in a single operation). The spec's own phrasing is singular - "a connector defined by one student" - and `frmManageLibrary`'s list is single-selection by construction (2c Task 5 never gave it `MultiSelect`). Extending to bulk export/import is a reasonable future enhancement, not a requirement this plan silently dropped.

**Type consistency.** `ImportConnector`'s destination `vFields` array is built by mutating the exact array `ReadConnector` returned (`vFields(0)`, `vFields(7)`, `vFields(10)`) rather than reconstructing it positionally by hand, so it cannot drift from `modLibrary`'s 11-field order even if that order is ever reordered upstream in a way `LIB_COL_*` constants alone wouldn't catch in VBA. `ExportConnector` and `ImportConnector` both reuse `CopyConnectorPhoto`'s two-ID signature (`sSrcConnectorID`, `sDestConnectorID`) rather than assuming the two are always equal, which is exactly what makes the collision-rename case correct without a second, parallel photo-copy implementation.

**No placeholders.** Every step contains complete, runnable code, including the one piece with an honestly-uncertain outcome (Task 1's clipboard operation) - it ships real code and a real test, with instructions for what "this doesn't work here" looks like and what to do about it, not a stub.
