# Phase 1: Build System and Creator Shell Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce a buildable, tested `HarnessCreator.xlsm` containing the Creator's sheet structure and a working to-from chart with dependent pin dropdowns.

**Architecture:** The `.xlsm` is a build artifact, never hand-edited. Python drives Excel COM to construct the workbook and import VBA source from `src/vba/*.bas`, so all real source stays as reviewable text. pytest drives the same COM interface against the built artifact, calling VBA entry points through `Application.Run` and asserting on the result.

**Tech Stack:** Python 3.13, pywin32, pytest, Excel 16.0 COM automation, VBA.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

## Global Constraints

Copied verbatim from the spec. Every task's requirements implicitly include these.

- Windows and desktop Excel only. No Excel-for-web or non-Windows support.
- Formulas must work on Excel 2016 and later. Use INDEX/MATCH, never XLOOKUP.
- Every VBA module starts with `Option Explicit`.
- No `MsgBox` or dialog in a logic module. UI is confined to UI modules, and `_State` carries a `TestMode` flag that suppresses it.
- Validation is advisory. Nothing the tool reports may block saving or exporting.
- The `.xlsm` is a build artifact. Never edit it by hand; change source and rebuild.
- To-from chart column order is fixed: From Conn, From Pin, From Term, Signal, Color, AWG, Length, To Term, To Conn, To Pin, Notes. There is no wire ID column.
- Reference designator prefixes by type: `Connector`→`J`, `Stud`→`ST`, `Splice`→`SP`, `Tail`→`TL`.
- Pin coordinates are normalized 0.0–1.0. (Phase 2; listed here because it is project-wide.)
- Chart data occupies rows 7 through 1006, a practical cap of 1000 wires.

## Prerequisite

Excel's **Trust access to the VBA project object model** must be enabled before any task runs: Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings. Task 1 builds a check for this that prints instructions when it is off.

## File Structure

| File | Responsibility |
|---|---|
| `requirements.txt` | Python dependencies |
| `build/excel_com.py` | COM lifecycle, VBA import, code-name assignment, prerequisite check |
| `build/layout.py` | All sheet layout constants and sheet builders — the single source of truth for cell addresses |
| `build/build.py` | Orchestration and CLI |
| `src/vba/modUtil.bas` | Build stamp, join-key helper |
| `src/vba/modState.bas` | `_State` key/value access, dirty flag |
| `src/vba/modConnectors.bas` | Ref des allocation, connector instance rows |
| `src/vba/modChart.bas` | Pin validation rebuild, length units, New Harness |
| `src/vba/sheets/shHarness.evt` | Harness `Worksheet_Change` handler |
| `src/vba/sheets/shConnectors.evt` | Connectors `Worksheet_Change` handler |
| `tests/conftest.py` | Build fixture, Excel app fixture, workbook fixture, `run()` helper |
| `tests/test_build.py` | Toolchain proof |
| `tests/test_sheets.py` | Sheet presence, visibility, code names, headers |
| `tests/test_lists.py` | Pick lists and named ranges |
| `tests/test_titleblock.py` | Title block names and length units |
| `tests/test_validation.py` | Column data validation |
| `tests/test_connectors.py` | Ref des allocation and instance rows |
| `tests/test_pin_dropdown.py` | Dependent pin dropdowns |
| `tests/test_state.py` | Dirty tracking |
| `tests/test_new_harness.py` | New Harness command |

---

### Task 1: Build toolchain producing an .xlsm with imported VBA

Proves the whole toolchain end to end: Python creates a workbook, injects VBA from text, saves as `.xlsm`, and a test calls into that VBA.

**Files:**
- Create: `requirements.txt`
- Create: `build/excel_com.py`
- Create: `build/build.py`
- Create: `src/vba/modUtil.bas`
- Create: `tests/conftest.py`
- Test: `tests/test_build.py`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `excel_com.check_access_vbom() -> bool`
  - `excel_com.excel_app()` — context manager yielding an Excel `Application`
  - `excel_com.import_module(wb, path: Path) -> None`
  - `excel_com.set_codename(wb, sheet, codename: str) -> None`
  - `excel_com.add_sheet_code(wb, codename: str, code: str) -> None`
  - `build.build(out_dir: Path) -> Path` — returns the written `.xlsm` path
  - VBA `modUtil.BuildStamp() As String` returns `"0.1.0"`
  - VBA `modUtil.JoinKey(sConn, vPin) As String` returns `"J1|3"` form
  - pytest fixtures `artifact`, `app`, `wb`, and helper `run(wb, macro, *args)`

- [ ] **Step 1: Install dependencies**

Create `requirements.txt`:

```
pywin32>=306
pytest>=8.0
```

Run:

```bash
python -m pip install -r requirements.txt
```

- [ ] **Step 2: Write the failing test**

Create `tests/test_build.py`:

```python
from tests.conftest import run


def test_artifact_is_produced(artifact):
    assert artifact.exists()
    assert artifact.suffix == ".xlsm"


def test_vba_module_is_present_and_callable(wb):
    assert run(wb, "modUtil.BuildStamp") == "0.1.0"


def test_join_key_normalizes_case_and_whitespace(wb):
    assert run(wb, "modUtil.JoinKey", " j1 ", 3) == "J1|3"
```

- [ ] **Step 3: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_build.py -v
```

Expected: FAIL — `tests/conftest.py` does not exist, collection error.

- [ ] **Step 4: Write the COM helpers**

Create `build/excel_com.py`:

```python
"""Excel COM lifecycle and VBA injection helpers."""
from __future__ import annotations

import contextlib
import winreg
from pathlib import Path

import win32com.client as win32

XL_OPENXML_MACRO_ENABLED = 52
MSO_AUTOMATION_SECURITY_LOW = 1

VBOM_KEY = r"Software\Microsoft\Office\16.0\Excel\Security"


def check_access_vbom() -> bool:
    """True when Excel permits programmatic access to the VBA project."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, VBOM_KEY) as key:
            value, _ = winreg.QueryValueEx(key, "AccessVBOM")
            return int(value) == 1
    except OSError:
        return False


VBOM_INSTRUCTIONS = (
    "Excel is blocking programmatic access to the VBA project.\n"
    "Enable it once: Excel > File > Options > Trust Center > "
    "Trust Center Settings > Macro Settings >\n"
    "  check 'Trust access to the VBA project object model'.\n"
    "This is a per-user, reversible setting and is not required on "
    "student machines."
)


@contextlib.contextmanager
def excel_app():
    """Yield a hidden Excel Application and guarantee it is closed."""
    app = win32.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
    try:
        yield app
    finally:
        app.Quit()


def import_module(wb, path: Path) -> None:
    """Import a .bas file as a standard module."""
    wb.VBProject.VBComponents.Import(str(path))


def set_codename(wb, sheet, codename: str) -> None:
    """Rename a worksheet's VBA code name so its module can be addressed."""
    component = wb.VBProject.VBComponents(sheet.CodeName)
    component.Properties("_CodeName").Value = codename


def add_sheet_code(wb, codename: str, code: str) -> None:
    """Append source to a worksheet's document module."""
    wb.VBProject.VBComponents(codename).CodeModule.AddFromString(code)


def save_as_xlsm(wb, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()
    wb.SaveAs(Filename=str(path), FileFormat=XL_OPENXML_MACRO_ENABLED)
```

- [ ] **Step 5: Write the VBA utility module**

Create `src/vba/modUtil.bas`. The `Attribute VB_Name` line is required — `Import` uses it to name the module.

```vb
Attribute VB_Name = "modUtil"
Option Explicit

Public Const BUILD_VERSION As String = "0.1.0"

Public Function BuildStamp() As String
    BuildStamp = BUILD_VERSION
End Function

Public Function JoinKey(ByVal sConn As String, ByVal vPin As Variant) As String
    JoinKey = UCase$(Trim$(sConn)) & "|" & Trim$(CStr(vPin))
End Function
```

- [ ] **Step 6: Write the build orchestrator**

Create `build/build.py`:

```python
"""Build HarnessCreator.xlsm from source."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import excel_com

ROOT = Path(__file__).resolve().parents[1]
VBA_DIR = ROOT / "src" / "vba"
DIST = ROOT / "dist"
CREATOR_NAME = "HarnessCreator.xlsm"

VBA_MODULES = ["modUtil.bas"]


def build(out_dir: Path = DIST) -> Path:
    if not excel_com.check_access_vbom():
        raise SystemExit(excel_com.VBOM_INSTRUCTIONS)

    target = out_dir / CREATOR_NAME
    with excel_com.excel_app() as app:
        wb = app.Workbooks.Add()
        try:
            for name in VBA_MODULES:
                excel_com.import_module(wb, VBA_DIR / name)
            excel_com.save_as_xlsm(wb, target)
        finally:
            wb.Close(SaveChanges=False)
    return target


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the Harness Creator workbook.")
    parser.add_argument(
        "--check",
        action="store_true",
        help="Verify build prerequisites and exit.",
    )
    args = parser.parse_args()

    if args.check:
        if excel_com.check_access_vbom():
            print("Prerequisites OK.")
            return 0
        print(excel_com.VBOM_INSTRUCTIONS)
        return 1

    print(f"Built {build()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 7: Write the test fixtures**

Create `tests/conftest.py`:

```python
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest
import win32com.client as win32

ROOT = Path(__file__).resolve().parents[1]
ARTIFACT = ROOT / "dist" / "HarnessCreator.xlsm"

MSO_AUTOMATION_SECURITY_LOW = 1


def run(wb, macro: str, *args):
    """Call a VBA entry point in the given workbook."""
    return wb.Application.Run(f"'{wb.Name}'!{macro}", *args)


@pytest.fixture(scope="session")
def artifact() -> Path:
    subprocess.run(
        [sys.executable, str(ROOT / "build" / "build.py")],
        check=True,
        cwd=ROOT,
    )
    assert ARTIFACT.exists(), f"build produced no artifact at {ARTIFACT}"
    return ARTIFACT


@pytest.fixture(scope="session")
def app():
    application = win32.Dispatch("Excel.Application")
    application.Visible = False
    application.DisplayAlerts = False
    application.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
    try:
        yield application
    finally:
        application.Quit()


@pytest.fixture
def wb(app, artifact):
    """A freshly opened copy of the built workbook, discarded after each test."""
    book = app.Workbooks.Open(str(artifact))
    try:
        yield book
    finally:
        book.Close(SaveChanges=False)
```

Create an empty `tests/__init__.py` so `from tests.conftest import run` resolves:

```bash
touch tests/__init__.py
```

- [ ] **Step 8: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_build.py -v
```

Expected: 3 passed. If the run stops with the Trust Center message, enable the setting and re-run.

- [ ] **Step 9: Add .gitignore entries and commit**

`dist/` is already ignored. Verify `__pycache__/` and `.pytest_cache/` are too, then commit.

```bash
git add requirements.txt build/ src/ tests/
git commit -m "feat: build toolchain producing an xlsm with imported VBA"
```

---

### Task 2: Creator sheet skeleton

**Files:**
- Create: `build/layout.py`
- Modify: `build/build.py`
- Test: `tests/test_sheets.py`

**Interfaces:**
- Consumes: `excel_com.set_codename`, `excel_com.save_as_xlsm`, `build.build`.
- Produces:
  - `layout.SHEETS: list[tuple[str, str, int]]` — `(tab_name, codename, visibility)`
  - `layout.VISIBLE = -1`, `layout.VERY_HIDDEN = 2`
  - `layout.build_sheets(wb) -> dict[str, object]` — creates every sheet, sets code names and visibility, deletes Excel's default sheets, returns a `{tab_name: worksheet}` map

- [ ] **Step 1: Write the failing test**

Create `tests/test_sheets.py`:

```python
import pytest

EXPECTED = [
    ("Home", "shHome", -1),
    ("Harness", "shHarness", -1),
    ("Connectors", "shConnectors", -1),
    ("Check", "shCheck", -1),
    ("_Snapshot", "shSnapshot", 2),
    ("_Lists", "shLists", 2),
    ("_State", "shState", 2),
]


def test_sheet_count(wb):
    assert wb.Worksheets.Count == len(EXPECTED)


@pytest.mark.parametrize("tab,codename,visibility", EXPECTED)
def test_sheet_exists_with_codename_and_visibility(wb, tab, codename, visibility):
    sheet = wb.Worksheets(tab)
    assert sheet.CodeName == codename
    assert sheet.Visible == visibility


def test_sheet_order_matches_spec(wb):
    actual = [wb.Worksheets(i + 1).Name for i in range(wb.Worksheets.Count)]
    assert actual == [tab for tab, _, _ in EXPECTED]
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_sheets.py -v
```

Expected: FAIL — the workbook has one default sheet named `Sheet1`.

- [ ] **Step 3: Write the layout module**

Create `build/layout.py`:

```python
"""Sheet layout constants and builders — the single source of truth for addresses."""
from __future__ import annotations

VISIBLE = -1
VERY_HIDDEN = 2

# (tab name, VBA code name, visibility)
SHEETS = [
    ("Home", "shHome", VISIBLE),
    ("Harness", "shHarness", VISIBLE),
    ("Connectors", "shConnectors", VISIBLE),
    ("Check", "shCheck", VISIBLE),
    ("_Snapshot", "shSnapshot", VERY_HIDDEN),
    ("_Lists", "shLists", VERY_HIDDEN),
    ("_State", "shState", VERY_HIDDEN),
]


def build_sheets(wb, set_codename) -> dict:
    """Create every sheet in order, then remove Excel's defaults.

    Visibility is applied last: a sheet cannot be hidden while it is the only
    visible sheet, and Excel refuses to delete the last visible sheet.
    """
    originals = [wb.Worksheets(i + 1) for i in range(wb.Worksheets.Count)]

    sheets = {}
    anchor = originals[-1]
    for tab, codename, _ in SHEETS:
        sheet = wb.Worksheets.Add(After=anchor)
        sheet.Name = tab
        set_codename(wb, sheet, codename)
        sheets[tab] = sheet
        anchor = sheet

    for sheet in originals:
        sheet.Delete()

    for tab, _, visibility in SHEETS:
        sheets[tab].Visible = visibility

    return sheets
```

- [ ] **Step 4: Wire the layout into the build**

In `build/build.py`, add `import layout` beside `import excel_com`, and replace the body of the `with excel_com.excel_app()` block's `try` so sheets are built before the save:

```python
        try:
            layout.build_sheets(wb, excel_com.set_codename)
            for name in VBA_MODULES:
                excel_com.import_module(wb, VBA_DIR / name)
            excel_com.save_as_xlsm(wb, target)
        finally:
            wb.Close(SaveChanges=False)
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_sheets.py tests/test_build.py -v
```

Expected: all passed.

- [ ] **Step 6: Commit**

```bash
git add build/layout.py build/build.py tests/test_sheets.py
git commit -m "feat: create Creator sheet skeleton with code names and visibility"
```

---

### Task 3: Pick lists and named ranges

**Files:**
- Modify: `build/layout.py`
- Modify: `build/build.py`
- Test: `tests/test_lists.py`

**Interfaces:**
- Consumes: `layout.build_sheets`.
- Produces:
  - `layout.COLORS`, `layout.AWGS`, `layout.TERMINATIONS` — `list[str]`
  - `layout.build_lists(sheets) -> None`
  - `layout.build_names(wb) -> None` defining workbook names `ListColor`, `ListAWG`, `ListTermination`, `ListRefDes`

Each list name is a dynamic `OFFSET`/`COUNTA` range. `MAX(1, ...)` guards the zero-row case: an `OFFSET` of height 0 is an error, and data validation bound to an erroring name fails outright.

- [ ] **Step 1: Write the failing test**

Create `tests/test_lists.py`:

```python
import pytest

COLORS = [
    "Black", "White", "Red", "Green", "Blue", "Yellow", "Orange",
    "Brown", "Violet", "Gray", "Pink", "Tan", "Light Blue",
    "Light Green", "Other",
]
AWGS = ["24", "22", "20", "18", "16", "14", "12", "10", "8"]
TERMINATIONS = [
    "Crimp Pin", "Crimp Socket", "Ring Terminal", "Spade Terminal",
    "Butt Splice", "Ferrule", "Solder Cup", "Quick Disconnect",
    "Bare Tinned", "None",
]


def column_values(wb, column: int, count: int):
    sheet = wb.Worksheets("_Lists")
    return [str(sheet.Cells(row + 2, column).Value) for row in range(count)]


def test_color_list_seeded(wb):
    assert column_values(wb, 1, len(COLORS)) == COLORS


def test_awg_list_seeded(wb):
    assert column_values(wb, 2, len(AWGS)) == AWGS


def test_termination_list_seeded(wb):
    assert column_values(wb, 3, len(TERMINATIONS)) == TERMINATIONS


@pytest.mark.parametrize(
    "name,expected_count",
    [
        ("ListColor", len(COLORS)),
        ("ListAWG", len(AWGS)),
        ("ListTermination", len(TERMINATIONS)),
    ],
)
def test_list_name_resolves_to_the_right_height(wb, name, expected_count):
    assert wb.Names(name).RefersToRange.Rows.Count == expected_count


def test_refdes_name_survives_an_empty_connector_sheet(wb):
    # No connectors defined yet; the name must still resolve rather than error.
    assert wb.Names("ListRefDes").RefersToRange.Rows.Count == 1
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_lists.py -v
```

Expected: FAIL — `_Lists` is empty and the names do not exist.

- [ ] **Step 3: Add list data and names to the layout**

Append to `build/layout.py`:

```python
COLORS = [
    "Black", "White", "Red", "Green", "Blue", "Yellow", "Orange",
    "Brown", "Violet", "Gray", "Pink", "Tan", "Light Blue",
    "Light Green", "Other",
]

AWGS = ["24", "22", "20", "18", "16", "14", "12", "10", "8"]

TERMINATIONS = [
    "Crimp Pin", "Crimp Socket", "Ring Terminal", "Spade Terminal",
    "Butt Splice", "Ferrule", "Solder Cup", "Quick Disconnect",
    "Bare Tinned", "None",
]

LIST_COLUMNS = [
    ("Color", 1, COLORS),
    ("AWG", 2, AWGS),
    ("Termination", 3, TERMINATIONS),
]


def _dynamic_range(sheet: str, column_letter: str) -> str:
    """An OFFSET range that grows with its column and never has height 0."""
    return (
        f"=OFFSET('{sheet}'!${column_letter}$2,0,0,"
        f"MAX(1,COUNTA('{sheet}'!${column_letter}:${column_letter})-1),1)"
    )


LIST_NAMES = {
    "ListColor": _dynamic_range("_Lists", "A"),
    "ListAWG": _dynamic_range("_Lists", "B"),
    "ListTermination": _dynamic_range("_Lists", "C"),
    "ListRefDes": _dynamic_range("Connectors", "A"),
}


def build_lists(sheets) -> None:
    sheet = sheets["_Lists"]
    for header, column, values in LIST_COLUMNS:
        sheet.Cells(1, column).Value = header
        for offset, value in enumerate(values):
            # Text format keeps AWG sizes as strings rather than numbers.
            cell = sheet.Cells(offset + 2, column)
            cell.NumberFormat = "@"
            cell.Value = value


def build_names(wb) -> None:
    for name, refers_to in LIST_NAMES.items():
        wb.Names.Add(Name=name, RefersTo=refers_to)
```

`Connectors` column A is empty at build time, which is exactly the case `MAX(1, ...)` exists to survive.

- [ ] **Step 4: Wire into the build**

In `build/build.py`, inside the `try` block, after `layout.build_sheets(...)`:

```python
            sheets = layout.build_sheets(wb, excel_com.set_codename)
            layout.build_lists(sheets)
            layout.build_names(wb)
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_lists.py -v
```

Expected: 7 passed.

- [ ] **Step 6: Commit**

```bash
git add build/layout.py build/build.py tests/test_lists.py
git commit -m "feat: seed pick lists and dynamic named ranges"
```

---

### Task 4: Harness title block and chart headers

**Files:**
- Modify: `build/layout.py`
- Modify: `build/build.py`
- Test: `tests/test_titleblock.py`

**Interfaces:**
- Consumes: `layout.build_sheets`, `layout.build_names`.
- Produces:
  - `layout.CHART_HEADER_ROW = 6`, `layout.CHART_FIRST_ROW = 7`, `layout.CHART_LAST_ROW = 1006`
  - `layout.CHART_HEADERS: list[str]` — 11 headers in spec order
  - `layout.TITLE_BLOCK: list[tuple[str, str, str]]` — `(label, label_cell, value_cell)`
  - `layout.TB_NAMES: dict[str, str]` — workbook name to value cell
  - `layout.build_harness(sheets) -> None`

Title block names exist so VBA and tests address fields by meaning rather than by cell address, which keeps later phases from breaking when the layout shifts.

- [ ] **Step 1: Write the failing test**

Create `tests/test_titleblock.py`:

```python
import pytest

CHART_HEADERS = [
    "From Conn", "From Pin", "From Term", "Signal", "Color", "AWG",
    "Length (in)", "To Term", "To Conn", "To Pin", "Notes",
]

TB_NAMES = [
    "TB_Name", "TB_Number", "TB_Rev", "TB_Student",
    "TB_Class", "TB_Date", "TB_Desc", "TB_Units",
]


@pytest.mark.parametrize("index,header", list(enumerate(CHART_HEADERS, start=1)))
def test_chart_header_text_and_order(wb, index, header):
    assert wb.Worksheets("Harness").Cells(6, index).Value == header


def test_chart_has_exactly_eleven_columns(wb):
    assert wb.Worksheets("Harness").Cells(6, 12).Value is None


@pytest.mark.parametrize("name", TB_NAMES)
def test_title_block_name_resolves_to_the_harness_sheet(wb, name):
    target = wb.Names(name).RefersToRange
    assert target.Worksheet.Name == "Harness"
    assert target.Cells.Count == 1


def test_units_default_to_inches(wb):
    assert wb.Names("TB_Units").RefersToRange.Value == "in"


def test_units_cell_offers_both_options(wb):
    assert wb.Names("TB_Units").RefersToRange.Validation.Formula1 == "in,mm"
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_titleblock.py -v
```

Expected: FAIL — the Harness sheet is empty and the names do not exist.

- [ ] **Step 3: Add the harness layout**

Append to `build/layout.py`:

```python
CHART_HEADER_ROW = 6
CHART_FIRST_ROW = 7
CHART_LAST_ROW = 1006

CHART_HEADERS = [
    "From Conn", "From Pin", "From Term", "Signal", "Color", "AWG",
    "Length (in)", "To Term", "To Conn", "To Pin", "Notes",
]

CHART_COLUMN_WIDTHS = [11, 9, 15, 18, 13, 7, 12, 15, 11, 9, 30]

# (label text, label cell, value cell)
TITLE_BLOCK = [
    ("Harness Name", "A2", "B2"),
    ("Harness Number", "D2", "E2"),
    ("Revision", "G2", "H2"),
    ("Student", "A3", "B3"),
    ("Class / Project", "D3", "E3"),
    ("Date", "G3", "H3"),
    ("Description", "A4", "B4"),
    ("Length Units", "G4", "H4"),
]

TB_NAMES = {
    "TB_Name": "B2",
    "TB_Number": "E2",
    "TB_Rev": "H2",
    "TB_Student": "B3",
    "TB_Class": "E3",
    "TB_Date": "H3",
    "TB_Desc": "B4",
    "TB_Units": "H4",
}

XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1


def build_harness(sheets) -> None:
    sheet = sheets["Harness"]

    sheet.Range("A1").Value = "WIRE HARNESS TO-FROM CHART"
    sheet.Range("A1").Font.Size = 16
    sheet.Range("A1").Font.Bold = True

    for label, label_cell, value_cell in TITLE_BLOCK:
        sheet.Range(label_cell).Value = label
        sheet.Range(label_cell).Font.Bold = True
        sheet.Range(value_cell).Interior.Color = 0xF2F2F2

    sheet.Range(TB_NAMES["TB_Units"]).Value = "in"
    units = sheet.Range(TB_NAMES["TB_Units"]).Validation
    units.Delete()
    units.Add(
        Type=XL_VALIDATE_LIST,
        AlertStyle=XL_VALID_ALERT_STOP,
        Operator=XL_BETWEEN,
        Formula1="in,mm",
    )
    units.IgnoreBlank = True
    units.InCellDropdown = True

    for index, header in enumerate(CHART_HEADERS, start=1):
        cell = sheet.Cells(CHART_HEADER_ROW, index)
        cell.Value = header
        cell.Font.Bold = True
        cell.Interior.Color = 0xD9D9D9
        sheet.Columns(index).ColumnWidth = CHART_COLUMN_WIDTHS[index - 1]


def build_title_block_names(wb, sheets) -> None:
    for name, cell in TB_NAMES.items():
        wb.Names.Add(Name=name, RefersTo=f"='Harness'!${cell[0]}${cell[1:]}")
```

- [ ] **Step 4: Wire into the build**

In `build/build.py`, inside the `try` block after `layout.build_names(wb)`:

```python
            layout.build_harness(sheets)
            layout.build_title_block_names(wb, sheets)
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_titleblock.py -v
```

Expected: 22 passed.

- [ ] **Step 6: Commit**

```bash
git add build/layout.py build/build.py tests/test_titleblock.py
git commit -m "feat: lay out harness title block and to-from chart headers"
```

---

### Task 5: Chart column data validation

**Files:**
- Modify: `build/layout.py`
- Modify: `build/build.py`
- Test: `tests/test_validation.py`

**Interfaces:**
- Consumes: `layout.LIST_NAMES`, `layout.CHART_FIRST_ROW`, `layout.CHART_LAST_ROW`.
- Produces: `layout.build_chart_validation(sheets) -> None`

Pin columns (B and J) get no validation at build time — they are populated per row by VBA in Task 7, because the allowed values depend on which connector the row references.

- [ ] **Step 1: Write the failing test**

Create `tests/test_validation.py`:

```python
import pytest

XL_VALIDATE_LIST = 3
XL_VALIDATE_DECIMAL = 2

# (column index, validation type, Formula1)
EXPECTED = [
    (1, XL_VALIDATE_LIST, "=ListRefDes"),
    (3, XL_VALIDATE_LIST, "=ListTermination"),
    (5, XL_VALIDATE_LIST, "=ListColor"),
    (6, XL_VALIDATE_LIST, "=ListAWG"),
    (7, XL_VALIDATE_DECIMAL, "0"),
    (8, XL_VALIDATE_LIST, "=ListTermination"),
    (9, XL_VALIDATE_LIST, "=ListRefDes"),
]


@pytest.mark.parametrize("column,vtype,formula", EXPECTED)
def test_first_data_row_validation(wb, column, vtype, formula):
    cell = wb.Worksheets("Harness").Cells(7, column)
    assert cell.Validation.Type == vtype
    assert cell.Validation.Formula1 == formula


@pytest.mark.parametrize("column,vtype,formula", EXPECTED)
def test_last_data_row_validation(wb, column, vtype, formula):
    cell = wb.Worksheets("Harness").Cells(1006, column)
    assert cell.Validation.Type == vtype
    assert cell.Validation.Formula1 == formula


@pytest.mark.parametrize("column", [2, 10])
def test_pin_columns_start_without_validation(wb, column):
    cell = wb.Worksheets("Harness").Cells(7, column)
    with pytest.raises(Exception):
        _ = cell.Validation.Type


@pytest.mark.parametrize("column", [4, 11])
def test_free_text_columns_have_no_validation(wb, column):
    cell = wb.Worksheets("Harness").Cells(7, column)
    with pytest.raises(Exception):
        _ = cell.Validation.Type
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_validation.py -v
```

Expected: the `test_first_data_row_validation` and `test_last_data_row_validation` cases FAIL; the two "no validation" cases already pass.

- [ ] **Step 3: Add validation to the layout**

Append to `build/layout.py`:

```python
XL_VALIDATE_DECIMAL = 2
XL_GREATER = 5

# (column index, validation type, Formula1)
CHART_VALIDATION = [
    (1, XL_VALIDATE_LIST, "=ListRefDes"),
    (3, XL_VALIDATE_LIST, "=ListTermination"),
    (5, XL_VALIDATE_LIST, "=ListColor"),
    (6, XL_VALIDATE_LIST, "=ListAWG"),
    (7, XL_VALIDATE_DECIMAL, "0"),
    (8, XL_VALIDATE_LIST, "=ListTermination"),
    (9, XL_VALIDATE_LIST, "=ListRefDes"),
]


def build_chart_validation(sheets) -> None:
    sheet = sheets["Harness"]
    for column, vtype, formula in CHART_VALIDATION:
        target = sheet.Range(
            sheet.Cells(CHART_FIRST_ROW, column),
            sheet.Cells(CHART_LAST_ROW, column),
        )
        target.Validation.Delete()
        if vtype == XL_VALIDATE_DECIMAL:
            target.Validation.Add(
                Type=vtype,
                AlertStyle=XL_VALID_ALERT_STOP,
                Operator=XL_GREATER,
                Formula1=formula,
            )
        else:
            target.Validation.Add(
                Type=vtype,
                AlertStyle=XL_VALID_ALERT_STOP,
                Operator=XL_BETWEEN,
                Formula1=formula,
            )
            target.Validation.InCellDropdown = True
        target.Validation.IgnoreBlank = True
```

- [ ] **Step 4: Wire into the build**

In `build/build.py`, after `layout.build_title_block_names(wb, sheets)`:

```python
            layout.build_chart_validation(sheets)
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_validation.py -v
```

Expected: 18 passed.

- [ ] **Step 6: Commit**

```bash
git add build/layout.py build/build.py tests/test_validation.py
git commit -m "feat: apply pick-list and length validation to chart columns"
```

---

### Task 6: Connectors sheet and reference designator allocation

**Files:**
- Modify: `build/layout.py`
- Modify: `build/build.py`
- Create: `src/vba/modState.bas`
- Create: `src/vba/modConnectors.bas`
- Test: `tests/test_connectors.py`
- Test: `tests/test_state.py`

**Interfaces:**
- Consumes: `layout.build_sheets`, `modUtil`.
- Produces:
  - `layout.CONN_HEADERS`, `layout.CONN_FIRST_ROW = 2`, `layout.CHECK_HEADERS`, `layout.STATE_ROWS`
  - `layout.build_connectors(sheets)`, `layout.build_check(sheets)`, `layout.build_state(sheets, version)`
  - VBA `modState.GetState(sKey) As String`, `SetState(sKey, sValue)`, `MarkDirty()`, `ClearDirty()`, `IsDirty() As Boolean`
  - VBA `modConnectors.PrefixForType(sType) As String`
  - VBA `modConnectors.NextRefDes(sPrefix) As String`
  - VBA `modConnectors.AddConnectorInstance(sConnectorID, sName, sPartNumber, sType, nPinCount) As String` — returns the assigned ref des, or `""` when the type is unknown or the pin count is below 1
  - VBA `modConnectors.PinCountFor(sRefDes) As Long` — `0` when not found

- [ ] **Step 1: Write the failing tests**

Create `tests/test_state.py`:

```python
from tests.conftest import run


def test_seeded_keys_are_readable(wb):
    assert run(wb, "modState.GetState", "LengthUnits") == "in"
    assert run(wb, "modState.GetState", "HarnessPath") == ""


def test_unknown_key_returns_empty(wb):
    assert run(wb, "modState.GetState", "NoSuchKey") == ""


def test_set_state_round_trips(wb):
    run(wb, "modState.SetState", "HarnessPath", r"C:\temp\x.xlsx")
    assert run(wb, "modState.GetState", "HarnessPath") == r"C:\temp\x.xlsx"


def test_set_state_creates_a_missing_key(wb):
    run(wb, "modState.SetState", "BrandNewKey", "value")
    assert run(wb, "modState.GetState", "BrandNewKey") == "value"


def test_dirty_flag_starts_clear_and_toggles(wb):
    assert run(wb, "modState.IsDirty") is False
    run(wb, "modState.MarkDirty")
    assert run(wb, "modState.IsDirty") is True
    run(wb, "modState.ClearDirty")
    assert run(wb, "modState.IsDirty") is False
```

Create `tests/test_connectors.py`:

```python
import pytest

from tests.conftest import run

HEADERS = ["Ref Des", "ConnectorID", "Name", "Part Number", "Type", "Pin Count"]


@pytest.mark.parametrize("index,header", list(enumerate(HEADERS, start=1)))
def test_connectors_headers(wb, index, header):
    assert wb.Worksheets("Connectors").Cells(1, index).Value == header


@pytest.mark.parametrize(
    "connector_type,prefix",
    [
        ("Connector", "J"),
        ("Stud", "ST"),
        ("Splice", "SP"),
        ("Tail", "TL"),
        ("connector", "J"),
        ("  Tail  ", "TL"),
    ],
)
def test_prefix_for_type(wb, connector_type, prefix):
    assert run(wb, "modConnectors.PrefixForType", connector_type) == prefix


def test_unknown_type_has_no_prefix(wb):
    assert run(wb, "modConnectors.PrefixForType", "Widget") == ""


def test_first_ref_des_of_each_prefix_is_one(wb):
    assert run(wb, "modConnectors.NextRefDes", "J") == "J1"
    assert run(wb, "modConnectors.NextRefDes", "ST") == "ST1"


def test_ref_des_numbering_increments_per_prefix(wb):
    assert run(wb, "modConnectors.AddConnectorInstance",
               "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4) == "J1"
    assert run(wb, "modConnectors.AddConnectorInstance",
               "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4) == "J2"
    # A different prefix numbers independently.
    assert run(wb, "modConnectors.AddConnectorInstance",
               "GND-STUD", "Chassis ground stud", "", "Stud", 1) == "ST1"


def test_same_part_can_appear_twice_as_distinct_instances(wb):
    first = run(wb, "modConnectors.AddConnectorInstance",
                "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    second = run(wb, "modConnectors.AddConnectorInstance",
                 "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    assert first != second
    sheet = wb.Worksheets("Connectors")
    assert sheet.Cells(2, 2).Value == sheet.Cells(3, 2).Value == "DTM-04P"


def test_instance_row_is_written_in_full(wb):
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    sheet = wb.Worksheets("Connectors")
    row = [sheet.Cells(2, c).Value for c in range(1, 7)]
    assert row == ["J1", "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4]


def test_unknown_type_is_rejected(wb):
    assert run(wb, "modConnectors.AddConnectorInstance",
               "X", "X", "", "Widget", 4) == ""
    assert wb.Worksheets("Connectors").Cells(2, 1).Value is None


def test_pin_count_below_one_is_rejected(wb):
    assert run(wb, "modConnectors.AddConnectorInstance",
               "X", "X", "", "Connector", 0) == ""


def test_adding_a_connector_marks_the_workbook_dirty(wb):
    run(wb, "modState.ClearDirty")
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    assert run(wb, "modState.IsDirty") is True


def test_pin_count_lookup(wb):
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-12P", "Deutsch DTM 12-way", "DTM06-12S", "Connector", 12)
    assert run(wb, "modConnectors.PinCountFor", "J1") == 12
    assert run(wb, "modConnectors.PinCountFor", "j1") == 12
    assert run(wb, "modConnectors.PinCountFor", "J99") == 0


def test_ref_des_dropdown_sees_added_connectors(wb):
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    assert wb.Names("ListRefDes").RefersToRange.Rows.Count == 1
    run(wb, "modConnectors.AddConnectorInstance",
        "GND-STUD", "Chassis ground stud", "", "Stud", 1)
    assert wb.Names("ListRefDes").RefersToRange.Rows.Count == 2
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
python -m pytest tests/test_state.py tests/test_connectors.py -v
```

Expected: FAIL — the headers are absent and neither VBA module exists.

- [ ] **Step 3: Write the state module**

Create `src/vba/modState.bas`:

```vb
Attribute VB_Name = "modState"
Option Explicit

Public Const STATE_SHEET As String = "_State"
Private Const STATE_FIRST_ROW As Long = 2

Private Function KeyRow(ws As Worksheet, ByVal sKey As String) As Long
    Dim r As Long, nLast As Long
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = STATE_FIRST_ROW To nLast
        If StrComp(CStr(ws.Cells(r, 1).Value), sKey, vbTextCompare) = 0 Then
            KeyRow = r
            Exit Function
        End If
    Next r
    KeyRow = 0
End Function

Public Function GetState(ByVal sKey As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(STATE_SHEET)
    r = KeyRow(ws, sKey)
    If r = 0 Then
        GetState = ""
    Else
        GetState = CStr(ws.Cells(r, 2).Value)
    End If
End Function

Public Sub SetState(ByVal sKey As String, ByVal sValue As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(STATE_SHEET)
    r = KeyRow(ws, sKey)
    If r = 0 Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < STATE_FIRST_ROW Then r = STATE_FIRST_ROW
        ws.Cells(r, 1).Value = sKey
    End If
    ' Text format stops Excel coercing "TRUE" into a Boolean.
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = sValue
End Sub

Public Sub MarkDirty()
    SetState "Dirty", "TRUE"
End Sub

Public Sub ClearDirty()
    SetState "Dirty", "FALSE"
End Sub

Public Function IsDirty() As Boolean
    IsDirty = (UCase$(GetState("Dirty")) = "TRUE")
End Function

Public Function IsTestMode() As Boolean
    IsTestMode = (UCase$(GetState("TestMode")) = "TRUE")
End Function
```

- [ ] **Step 4: Write the connectors module**

Create `src/vba/modConnectors.bas`:

```vb
Attribute VB_Name = "modConnectors"
Option Explicit

Public Const CONN_SHEET As String = "Connectors"
Public Const CONN_FIRST_ROW As Long = 2

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

Public Function PrefixForType(ByVal sType As String) As String
    Select Case LCase$(Trim$(sType))
        Case "connector": PrefixForType = "J"
        Case "stud":      PrefixForType = "ST"
        Case "splice":    PrefixForType = "SP"
        Case "tail":      PrefixForType = "TL"
        Case Else:        PrefixForType = ""
    End Select
End Function

Public Function NextRefDes(ByVal sPrefix As String) As String
    Dim ws As Worksheet
    Dim r As Long, nLast As Long, nMax As Long, nNum As Long
    Dim sVal As String, sTail As String, sUpper As String

    If Len(sPrefix) = 0 Then Exit Function

    sUpper = UCase$(sPrefix)
    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = CONN_FIRST_ROW To nLast
        sVal = UCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
        If Left$(sVal, Len(sUpper)) = sUpper Then
            sTail = Mid$(sVal, Len(sUpper) + 1)
            If IsAllDigits(sTail) Then
                nNum = CLng(sTail)
                If nNum > nMax Then nMax = nNum
            End If
        End If
    Next r

    NextRefDes = sUpper & CStr(nMax + 1)
End Function

Public Function AddConnectorInstance(ByVal sConnectorID As String, _
                                     ByVal sName As String, _
                                     ByVal sPartNumber As String, _
                                     ByVal sType As String, _
                                     ByVal nPinCount As Long) As String
    Dim ws As Worksheet
    Dim sPrefix As String, sRef As String
    Dim r As Long

    sPrefix = PrefixForType(sType)
    If Len(sPrefix) = 0 Then Exit Function
    If nPinCount < 1 Then Exit Function

    sRef = NextRefDes(sPrefix)
    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < CONN_FIRST_ROW Then r = CONN_FIRST_ROW

    ws.Cells(r, 1).Value = sRef
    ws.Cells(r, 2).Value = sConnectorID
    ws.Cells(r, 3).Value = sName
    ws.Cells(r, 4).Value = sPartNumber
    ws.Cells(r, 5).Value = Trim$(sType)
    ws.Cells(r, 6).Value = nPinCount

    modState.MarkDirty
    AddConnectorInstance = sRef
End Function

Public Function PinCountFor(ByVal sRefDes As String) As Long
    Dim ws As Worksheet
    Dim r As Long, nLast As Long

    Set ws = ThisWorkbook.Worksheets(CONN_SHEET)
    nLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = CONN_FIRST_ROW To nLast
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), Trim$(sRefDes), vbTextCompare) = 0 Then
            PinCountFor = CLng(Val(ws.Cells(r, 6).Value))
            Exit Function
        End If
    Next r

    PinCountFor = 0
End Function
```

- [ ] **Step 5: Add the remaining sheet layouts**

Append to `build/layout.py`:

```python
CONN_HEADERS = ["Ref Des", "ConnectorID", "Name", "Part Number", "Type", "Pin Count"]
CONN_COLUMN_WIDTHS = [10, 16, 28, 18, 12, 10]
CONN_FIRST_ROW = 2

CHECK_HEADERS = ["Row", "Severity", "Message"]
CHECK_COLUMN_WIDTHS = [8, 12, 80]

STATE_KEYS = ["BuildVersion", "HarnessPath", "Dirty", "LengthUnits", "TestMode"]
STATE_DEFAULTS = {
    "HarnessPath": "",
    "Dirty": "FALSE",
    "LengthUnits": "in",
    "TestMode": "FALSE",
}


def _build_header_sheet(sheet, headers, widths) -> None:
    for index, header in enumerate(headers, start=1):
        cell = sheet.Cells(1, index)
        cell.Value = header
        cell.Font.Bold = True
        cell.Interior.Color = 0xD9D9D9
        sheet.Columns(index).ColumnWidth = widths[index - 1]


def build_connectors(sheets) -> None:
    _build_header_sheet(sheets["Connectors"], CONN_HEADERS, CONN_COLUMN_WIDTHS)


def build_check(sheets) -> None:
    _build_header_sheet(sheets["Check"], CHECK_HEADERS, CHECK_COLUMN_WIDTHS)


def build_state(sheets, version: str) -> None:
    sheet = sheets["_State"]
    sheet.Cells(1, 1).Value = "Key"
    sheet.Cells(1, 2).Value = "Value"
    values = dict(STATE_DEFAULTS, BuildVersion=version)
    for offset, key in enumerate(STATE_KEYS):
        row = offset + 2
        sheet.Cells(row, 1).Value = key
        sheet.Cells(row, 2).NumberFormat = "@"
        sheet.Cells(row, 2).Value = values[key]
```

- [ ] **Step 6: Wire into the build**

In `build/build.py`, extend the module list and the build body:

```python
VBA_MODULES = ["modUtil.bas", "modState.bas", "modConnectors.bas"]
BUILD_VERSION = "0.1.0"
```

and after `layout.build_chart_validation(sheets)`:

```python
            layout.build_connectors(sheets)
            layout.build_check(sheets)
            layout.build_state(sheets, BUILD_VERSION)
```

- [ ] **Step 7: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_state.py tests/test_connectors.py -v
```

Expected: all passed.

- [ ] **Step 8: Commit**

```bash
git add build/layout.py build/build.py src/vba/modState.bas src/vba/modConnectors.bas tests/test_state.py tests/test_connectors.py
git commit -m "feat: add connector instances with type-based ref des allocation"
```

---

### Task 7: Dependent pin dropdowns

**Files:**
- Create: `src/vba/modChart.bas`
- Create: `src/vba/sheets/shHarness.evt`
- Create: `src/vba/sheets/shConnectors.evt`
- Modify: `build/excel_com.py`
- Modify: `build/build.py`
- Test: `tests/test_pin_dropdown.py`

**Interfaces:**
- Consumes: `modConnectors.PinCountFor`, `modState.MarkDirty`, `excel_com.add_sheet_code`.
- Produces:
  - VBA `modChart.RebuildPinValidation(nRow As Long, nConnCol As Long)`
  - VBA `modChart.SetLengthUnits(sUnit As String)`
  - VBA constants `CHART_SHEET`, `CHART_HEADER_ROW`, `CHART_FIRST_ROW`, `CHART_LAST_ROW`, `COL_FROM_CONN = 1`, `COL_FROM_PIN = 2`, `COL_LENGTH = 7`, `COL_TO_CONN = 9`, `COL_TO_PIN = 10`, `COL_NOTES = 11`
  - `build.SHEET_EVENTS: list[tuple[str, str]]` — `(codename, source filename)`

Two behaviours matter and are easy to get wrong. First, the handler must disable events before it writes, because clearing a pin cell would otherwise re-enter the handler. Second, a pin list longer than 255 characters exceeds Excel's `Formula1` limit, so large connectors fall back to whole-number validation.

- [ ] **Step 1: Write the failing test**

Create `tests/test_pin_dropdown.py`:

```python
import pytest

from tests.conftest import run


def add_four_way(wb):
    return run(wb, "modConnectors.AddConnectorInstance",
               "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)


def test_typing_a_connector_builds_its_pin_list(wb):
    add_four_way(wb)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    assert sheet.Cells(7, 2).Validation.Formula1 == "1,2,3,4"


def test_to_connector_drives_the_to_pin_column(wb):
    add_four_way(wb)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 9).Value = "J1"
    assert sheet.Cells(7, 10).Validation.Formula1 == "1,2,3,4"


def test_changing_the_connector_clears_a_stale_pin(wb):
    add_four_way(wb)
    run(wb, "modConnectors.AddConnectorInstance",
        "GND-STUD", "Chassis ground stud", "", "Stud", 1)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(7, 2).Value = 4
    sheet.Cells(7, 1).Value = "ST1"
    assert sheet.Cells(7, 2).Value is None
    assert sheet.Cells(7, 2).Validation.Formula1 == "1"


def test_unknown_connector_leaves_no_pin_validation(wb):
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J99"
    with pytest.raises(Exception):
        _ = sheet.Cells(7, 2).Validation.Type


def test_large_connector_falls_back_to_a_numeric_range(wb):
    # 1,2,...,120 exceeds Excel's 255-character Formula1 limit.
    run(wb, "modConnectors.AddConnectorInstance",
        "BIG-120", "120 way", "", "Connector", 120)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    assert sheet.Cells(7, 2).Validation.Type == 1  # xlValidateWholeNumber
    assert sheet.Cells(7, 2).Validation.Formula1 == "1"
    assert sheet.Cells(7, 2).Validation.Formula2 == "120"


def test_each_row_gets_its_own_pin_list(wb):
    add_four_way(wb)
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-12P", "12 way", "", "Connector", 12)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(8, 1).Value = "J2"
    assert sheet.Cells(7, 2).Validation.Formula1 == "1,2,3,4"
    assert sheet.Cells(8, 2).Validation.Formula1 == "1,2,3,4,5,6,7,8,9,10,11,12"


def test_editing_the_chart_marks_the_workbook_dirty(wb):
    run(wb, "modState.ClearDirty")
    wb.Worksheets("Harness").Cells(7, 4).Value = "+12V Batt"
    assert run(wb, "modState.IsDirty") is True


def test_switching_units_rewrites_the_length_header(wb):
    sheet = wb.Worksheets("Harness")
    wb.Names("TB_Units").RefersToRange.Value = "mm"
    assert sheet.Cells(6, 7).Value == "Length (mm)"
    assert run(wb, "modState.GetState", "LengthUnits") == "mm"
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
python -m pytest tests/test_pin_dropdown.py -v
```

Expected: FAIL — no event handler exists, so nothing responds to the writes.

- [ ] **Step 3: Write the chart module**

Create `src/vba/modChart.bas`:

```vb
Attribute VB_Name = "modChart"
Option Explicit

Public Const CHART_SHEET As String = "Harness"
Public Const CHART_HEADER_ROW As Long = 6
Public Const CHART_FIRST_ROW As Long = 7
Public Const CHART_LAST_ROW As Long = 1006

Public Const COL_FROM_CONN As Long = 1
Public Const COL_FROM_PIN As Long = 2
Public Const COL_LENGTH As Long = 7
Public Const COL_TO_CONN As Long = 9
Public Const COL_TO_PIN As Long = 10
Public Const COL_NOTES As Long = 11

Private Const MAX_FORMULA1 As Long = 255

Public Sub RebuildPinValidation(ByVal nRow As Long, ByVal nConnCol As Long)
    Dim ws As Worksheet, cel As Range
    Dim nPinCol As Long, nPins As Long, i As Long
    Dim sRef As String, sList As String

    Select Case nConnCol
        Case COL_FROM_CONN: nPinCol = COL_FROM_PIN
        Case COL_TO_CONN:   nPinCol = COL_TO_PIN
        Case Else:          Exit Sub
    End Select

    Set ws = ThisWorkbook.Worksheets(CHART_SHEET)
    Set cel = ws.Cells(nRow, nPinCol)

    sRef = Trim$(CStr(ws.Cells(nRow, nConnCol).Value))
    nPins = modConnectors.PinCountFor(sRef)

    cel.Validation.Delete
    cel.ClearContents
    If nPins < 1 Then Exit Sub

    For i = 1 To nPins
        If Len(sList) > 0 Then sList = sList & ","
        sList = sList & CStr(i)
    Next i

    If Len(sList) <= MAX_FORMULA1 Then
        cel.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                           Operator:=xlBetween, Formula1:=sList
        cel.Validation.InCellDropdown = True
    Else
        cel.Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                           Operator:=xlBetween, Formula1:="1", Formula2:=CStr(nPins)
    End If
    cel.Validation.IgnoreBlank = True
End Sub

Public Sub SetLengthUnits(ByVal sUnit As String)
    Dim ws As Worksheet
    Dim s As String
    Dim bEvents As Boolean

    s = LCase$(Trim$(sUnit))
    If s <> "in" And s <> "mm" Then Exit Sub

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    Set ws = ThisWorkbook.Worksheets(CHART_SHEET)
    ws.Cells(CHART_HEADER_ROW, COL_LENGTH).Value = "Length (" & s & ")"
    ThisWorkbook.Names("TB_Units").RefersToRange.Value = s
    modState.SetState "LengthUnits", s

CleanUp:
    Application.EnableEvents = bEvents
End Sub
```

- [ ] **Step 4: Write the sheet event handlers**

Create `src/vba/sheets/shHarness.evt`:

```vb
Private Const BULK_EDIT_THRESHOLD As Long = 500

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cel As Range
    Dim bEvents As Boolean

    If Not Application.EnableEvents Then Exit Sub

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    ' A bulk clear or paste has nothing per-cell worth doing.
    If Target.Cells.Count > BULK_EDIT_THRESHOLD Then
        modState.MarkDirty
        GoTo CleanUp
    End If

    For Each cel In Target.Cells
        If cel.Row >= modChart.CHART_FIRST_ROW And cel.Row <= modChart.CHART_LAST_ROW Then
            If cel.Column = modChart.COL_FROM_CONN Or cel.Column = modChart.COL_TO_CONN Then
                modChart.RebuildPinValidation cel.Row, cel.Column
            End If
            modState.MarkDirty
        ElseIf cel.Row < modChart.CHART_HEADER_ROW Then
            If Not Application.Intersect(cel, ThisWorkbook.Names("TB_Units").RefersToRange) Is Nothing Then
                modChart.SetLengthUnits CStr(cel.Value)
            End If
            modState.MarkDirty
        End If
    Next cel

CleanUp:
    Application.EnableEvents = bEvents
End Sub
```

Create `src/vba/sheets/shConnectors.evt`:

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Application.EnableEvents Then Exit Sub
    modState.MarkDirty
End Sub
```

- [ ] **Step 5: Teach the build to inject sheet code**

In `build/build.py`, add the event list and inject after the module import:

```python
VBA_MODULES = ["modUtil.bas", "modState.bas", "modConnectors.bas", "modChart.bas"]

SHEET_EVENTS = [
    ("shHarness", "shHarness.evt"),
    ("shConnectors", "shConnectors.evt"),
]
```

and inside the `try` block, after the `import_module` loop:

```python
            for codename, filename in SHEET_EVENTS:
                source = (VBA_DIR / "sheets" / filename).read_text(encoding="utf-8")
                excel_com.add_sheet_code(wb, codename, source)
```

- [ ] **Step 6: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_pin_dropdown.py -v
```

Expected: 8 passed.

- [ ] **Step 7: Run the whole suite**

Run:

```bash
python -m pytest -v
```

Expected: everything passes. A failure here means the event handler is firing during another test's setup — check the `EnableEvents` guard.

- [ ] **Step 8: Commit**

```bash
git add src/vba/modChart.bas src/vba/sheets/ build/build.py tests/test_pin_dropdown.py
git commit -m "feat: rebuild pin dropdowns from the referenced connector"
```

---

### Task 8: New Harness command and Home sheet

**Files:**
- Modify: `src/vba/modChart.bas`
- Modify: `build/layout.py`
- Modify: `build/build.py`
- Test: `tests/test_new_harness.py`

**Interfaces:**
- Consumes: `modChart.SetLengthUnits`, `modState`, `layout.TB_NAMES`.
- Produces:
  - VBA `modChart.NewHarness()`
  - `layout.HOME_TEXT: list[tuple[str, str]]` — `(cell, text)`
  - `layout.build_home(sheets) -> None` placing the instructions and a New Harness button

- [ ] **Step 1: Write the failing test**

Create `tests/test_new_harness.py`:

```python
import pytest

from tests.conftest import run

TB_NAMES = [
    "TB_Name", "TB_Number", "TB_Rev", "TB_Student",
    "TB_Class", "TB_Date", "TB_Desc",
]


def populate(wb):
    run(wb, "modConnectors.AddConnectorInstance",
        "DTM-04P", "Deutsch DTM 4-way", "DTM06-4S", "Connector", 4)
    sheet = wb.Worksheets("Harness")
    sheet.Cells(7, 1).Value = "J1"
    sheet.Cells(7, 2).Value = 1
    sheet.Cells(7, 4).Value = "+12V Batt"
    for name in TB_NAMES:
        wb.Names(name).RefersToRange.Value = "seeded"
    wb.Worksheets("Check").Cells(2, 3).Value = "stale finding"


def test_new_harness_clears_the_chart(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    sheet = wb.Worksheets("Harness")
    assert sheet.Cells(7, 1).Value is None
    assert sheet.Cells(7, 4).Value is None


def test_new_harness_clears_the_connector_list(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    assert wb.Worksheets("Connectors").Cells(2, 1).Value is None


def test_new_harness_clears_the_title_block(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    for name in TB_NAMES:
        assert wb.Names(name).RefersToRange.Value is None


def test_new_harness_clears_stale_check_results(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    assert wb.Worksheets("Check").Cells(2, 3).Value is None


def test_new_harness_resets_units_and_path(wb):
    populate(wb)
    wb.Names("TB_Units").RefersToRange.Value = "mm"
    run(wb, "modState.SetState", "HarnessPath", r"C:\temp\x.xlsx")
    run(wb, "modChart.NewHarness")
    assert wb.Worksheets("Harness").Cells(6, 7).Value == "Length (in)"
    assert run(wb, "modState.GetState", "HarnessPath") == ""


def test_new_harness_leaves_the_workbook_clean(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    assert run(wb, "modState.IsDirty") is False


def test_new_harness_drops_stale_pin_validation(wb):
    populate(wb)
    run(wb, "modChart.NewHarness")
    with pytest.raises(Exception):
        _ = wb.Worksheets("Harness").Cells(7, 2).Validation.Type


def test_home_sheet_has_a_new_harness_button(wb):
    shapes = wb.Worksheets("Home").Shapes
    actions = [shapes(i + 1).OnAction for i in range(shapes.Count)]
    assert "modChart.NewHarness" in actions
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
python -m pytest tests/test_new_harness.py -v
```

Expected: FAIL — `NewHarness` does not exist and Home has no shapes.

- [ ] **Step 3: Implement New Harness**

Append to `src/vba/modChart.bas`:

```vb
Private Const TB_CLEAR_NAMES As String = _
    "TB_Name,TB_Number,TB_Rev,TB_Student,TB_Class,TB_Date,TB_Desc"

Public Sub NewHarness()
    Dim wsHarness As Worksheet, wsConn As Worksheet, wsCheck As Worksheet
    Dim vNames As Variant, i As Long
    Dim bEvents As Boolean

    bEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    Set wsHarness = ThisWorkbook.Worksheets(CHART_SHEET)
    Set wsConn = ThisWorkbook.Worksheets(modConnectors.CONN_SHEET)
    Set wsCheck = ThisWorkbook.Worksheets("Check")

    wsHarness.Range(wsHarness.Cells(CHART_FIRST_ROW, COL_FROM_CONN), _
                    wsHarness.Cells(CHART_LAST_ROW, COL_NOTES)).ClearContents

    ' Pin validation is built per row, so it must be torn down per row too.
    wsHarness.Range(wsHarness.Cells(CHART_FIRST_ROW, COL_FROM_PIN), _
                    wsHarness.Cells(CHART_LAST_ROW, COL_FROM_PIN)).Validation.Delete
    wsHarness.Range(wsHarness.Cells(CHART_FIRST_ROW, COL_TO_PIN), _
                    wsHarness.Cells(CHART_LAST_ROW, COL_TO_PIN)).Validation.Delete

    vNames = Split(TB_CLEAR_NAMES, ",")
    For i = LBound(vNames) To UBound(vNames)
        ThisWorkbook.Names(vNames(i)).RefersToRange.ClearContents
    Next i

    wsConn.Range(wsConn.Cells(modConnectors.CONN_FIRST_ROW, 1), _
                 wsConn.Cells(wsConn.Rows.Count, 6)).ClearContents

    wsCheck.Range(wsCheck.Cells(2, 1), _
                  wsCheck.Cells(wsCheck.Rows.Count, 3)).ClearContents

    modState.SetState "HarnessPath", ""
    SetLengthUnits "in"
    modState.ClearDirty

CleanUp:
    Application.EnableEvents = bEvents
End Sub
```

`SetLengthUnits` toggles `EnableEvents` itself and restores it to whatever it found, so calling it from inside this guarded block is safe.

- [ ] **Step 4: Build the Home sheet**

Append to `build/layout.py`:

```python
MSO_SHAPE_ROUNDED_RECTANGLE = 5

HOME_TEXT = [
    ("A1", "WIRE HARNESS CREATOR"),
    ("A3", "This workbook is the editor. It is not a drawing."),
    ("A4", "Use it to build harness files, which are saved separately as .xlsx."),
    ("A6", "1. Add the connectors your harness uses on the Connectors sheet."),
    ("A7", "2. Fill in the to-from chart on the Harness sheet, one row per wire."),
    ("A8", "3. Pick From Conn first - the From Pin list is built from that connector."),
    ("A10", "New Harness clears everything and starts over."),
]

HOME_BUTTONS = [
    # (caption, macro, left, top, width, height)
    ("New Harness", "modChart.NewHarness", 20, 220, 120, 32),
]


def build_home(sheets) -> None:
    sheet = sheets["Home"]
    for cell, text in HOME_TEXT:
        sheet.Range(cell).Value = text
    sheet.Range("A1").Font.Size = 16
    sheet.Range("A1").Font.Bold = True
    sheet.Columns(1).ColumnWidth = 90

    for caption, macro, left, top, width, height in HOME_BUTTONS:
        shape = sheet.Shapes.AddShape(
            MSO_SHAPE_ROUNDED_RECTANGLE, left, top, width, height
        )
        shape.TextFrame2.TextRange.Text = caption
        shape.OnAction = macro
```

- [ ] **Step 5: Wire into the build**

In `build/build.py`, after `layout.build_state(sheets, BUILD_VERSION)`:

```python
            layout.build_home(sheets)
```

- [ ] **Step 6: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_new_harness.py -v
```

Expected: 8 passed.

- [ ] **Step 7: Commit**

```bash
git add src/vba/modChart.bas build/layout.py build/build.py tests/test_new_harness.py
git commit -m "feat: add New Harness command and Home sheet"
```

---

### Task 9: Full-suite green and developer documentation

**Files:**
- Create: `README.md`
- Test: entire suite

**Interfaces:**
- Consumes: everything above.
- Produces: nothing later tasks depend on.

- [ ] **Step 1: Run the whole suite from a clean build**

```bash
rm -rf dist
python -m pytest -v
```

Expected: every test passes and `dist/HarnessCreator.xlsm` is regenerated. If Excel processes are left running, `excel_app()` is not reaching its `finally` — fix that before continuing.

- [ ] **Step 2: Verify the prerequisite check reports correctly**

```bash
python build/build.py --check
```

Expected: `Prerequisites OK.` and exit code 0.

- [ ] **Step 3: Write the README**

Create `README.md`:

```markdown
# Wire Harness Documentation Creator

An Excel-based tool for students learning electrical documentation. Fill in a
to-from chart, attach photographs of the connectors, number the positions on
those photos, and get a print-ready harness drawing.

## Artifacts

| File | What it is |
|---|---|
| `dist/HarnessCreator.xlsm` | The editor. Build artifact - never edit by hand. |
| `dist/ConnectorLibrary.xlsx` | The reusable connector library. (Phase 2.) |
| `<harness>.xlsx` | A saved drawing. Macro-free and print-ready. (Phase 3.) |

## Developer setup

Requires Windows and desktop Excel.

    python -m pip install -r requirements.txt

Enable Excel's programmatic VBA access once, or the build cannot inject code:
File > Options > Trust Center > Trust Center Settings > Macro Settings >
"Trust access to the VBA project object model". Per-user and reversible; it is
not required on student machines.

Verify:

    python build/build.py --check

## Build and test

    python build/build.py
    python -m pytest -v

The `.xlsm` is generated from `src/vba/*.bas` and `build/layout.py`. To change
behaviour, change the source and rebuild. Edits made directly in the workbook
are lost on the next build.

## Layout

| Path | Responsibility |
|---|---|
| `build/excel_com.py` | COM lifecycle, VBA injection, prerequisite check |
| `build/layout.py` | Sheet layout constants and builders |
| `build/build.py` | Build orchestration and CLI |
| `src/vba/` | VBA modules; `sheets/*.evt` are worksheet event handlers |
| `tests/` | pytest suites driving Excel COM |
| `docs/superpowers/specs/` | Design spec |
| `docs/superpowers/plans/` | Implementation plans |

## Status

Phase 1 complete: build system, Creator shell, to-from chart with dependent
pin dropdowns. Phases 2 to 4 (connector library and editor, harness save and
load with rendered connector pages, validation and export) are specified but
not yet built. See the spec for the full design.
```

- [ ] **Step 4: Commit**

```bash
git add README.md
git commit -m "docs: add developer README for the build and test loop"
```

---

## Self-Review

**Spec coverage for phase 1.** The spec's phase 1 is "build system, test harness, Creator shell, `Harness` sheet with the working to-from chart, pick lists, and dependent pin dropdowns." Build system: Task 1. Test harness: Task 1. Creator shell sheets: Tasks 2, 6, 8. Chart with spec column order: Task 4. Pick lists: Task 3. Dependent pin dropdowns: Task 7. Ref des allocation, which the chart dropdowns depend on, is Task 6.

**Deliberately deferred to later phases,** and not gaps in this plan: the connector library file and its schema, the connector editor and click-to-place, snapshot embedding, harness save and load, connector page rendering with callouts and leaders, page setup, Check Drawing, PDF export, and archive copy. Every one is named in the spec's phasing table under phases 2 to 4.

**Known limitation carried forward.** The `_Snapshot` sheet is created in Task 2 but stays empty until phase 2. That is intentional — creating it now keeps the sheet set and its code name stable so phase 2 does not have to restructure the workbook.

**Type consistency.** `PinCountFor` returns `Long` and is called only by `RebuildPinValidation`. `AddConnectorInstance` returns `String` and is asserted as such in tests. `GetState` returns `String` in every use, and `IsDirty` wraps it to return `Boolean`. Column constants are defined once in `modChart` and once in `layout.py`; the plan's tests assert the two agree by checking header text and validation position against literal column indices.
