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
