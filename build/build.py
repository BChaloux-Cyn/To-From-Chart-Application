"""Build HarnessCreator.xlsm from source."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import excel_com
import form_layout
import layout
import library_layout

ROOT = Path(__file__).resolve().parents[1]
VBA_DIR = ROOT / "src" / "vba"
DIST = ROOT / "dist"
CREATOR_NAME = "HarnessCreator.xlsm"
LIBRARY_NAME = "ConnectorLibrary.xlsx"

VBA_MODULES = [
    "modUtil.bas", "modState.bas", "modConnectors.bas", "modChart.bas",
    "modLibrary.bas", "modPinEditor.bas", "clsPinMarker.cls", "modSnapshot.bas",
    "modConnectorUI.bas", "modLibraryTransfer.bas", "modContract.bas",
    "modMessages.bas", "modEditorActions.bas", "modPickerActions.bas",
    "modManageActions.bas", "modConnectorActions.bas",
    "modHarnessBuild.bas",
]
BUILD_VERSION = "0.1.0"

SHEET_EVENTS = [
    ("shHarness", "shHarness.evt"),
    ("shConnectors", "shConnectors.evt"),
]

FORM_EVENTS = [
    ("frmConnectorEditor", "frmConnectorEditor.evt"),
    ("frmConnectorPicker", "frmConnectorPicker.evt"),
    ("frmManageLibrary", "frmManageLibrary.evt"),
    ("frmRemoveConnector", "frmRemoveConnector.evt"),
]


def build(out_dir: Path = DIST) -> Path:
    if not excel_com.check_access_vbom():
        raise SystemExit(excel_com.VBOM_INSTRUCTIONS)

    target = out_dir / CREATOR_NAME
    with excel_com.excel_app() as app:
        wb = app.Workbooks.Add()
        try:
            sheets = layout.build_sheets(wb, excel_com.set_codename)
            layout.build_lists(sheets)
            layout.build_names(wb)
            layout.build_harness(sheets)
            layout.build_title_block_names(wb, sheets)
            layout.build_chart_validation(sheets)
            layout.build_connectors(sheets)
            layout.build_check(sheets)
            layout.build_state(sheets, BUILD_VERSION)
            layout.build_home(sheets)
            layout.build_snapshot(sheets)
            for name in VBA_MODULES:
                excel_com.import_module(wb, VBA_DIR / name)
            for codename, filename in SHEET_EVENTS:
                source = (VBA_DIR / "sheets" / filename).read_text(encoding="utf-8")
                excel_com.add_sheet_code(wb, codename, source)
            form_layout.build_connector_editor_form(wb, excel_com.add_userform)
            form_layout.build_connector_picker_form(wb, excel_com.add_userform)
            form_layout.build_manage_library_form(wb, excel_com.add_userform)
            form_layout.build_remove_connector_form(wb, excel_com.add_userform)
            for codename, filename in FORM_EVENTS:
                source = (VBA_DIR / "forms" / filename).read_text(encoding="utf-8")
                excel_com.add_sheet_code(wb, codename, source)
            excel_com.save_as_xlsm(wb, target)
        finally:
            wb.Close(SaveChanges=False)
    return target


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
    print(f"Built {build_library()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
