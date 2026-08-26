"""Build HarnessCreator.xlsm from source."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import excel_com
import layout

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
            sheets = layout.build_sheets(wb, excel_com.set_codename)
            layout.build_lists(sheets)
            layout.build_names(wb)
            layout.build_harness(sheets)
            layout.build_title_block_names(wb, sheets)
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
