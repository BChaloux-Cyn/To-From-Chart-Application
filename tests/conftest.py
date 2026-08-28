from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest
import win32com.client as win32

ROOT = Path(__file__).resolve().parents[1]
ARTIFACT = ROOT / "dist" / "HarnessCreator.xlsm"

sys.path.insert(0, str(ROOT / "build"))

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
    # EnsureDispatch (early-bound), not Dispatch: see excel_com.excel_app().
    application = win32.gencache.EnsureDispatch("Excel.Application")
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


from dataclasses import dataclass


@dataclass(frozen=True)
class Result:
    """A layer 1 action's three-element envelope, unpacked."""

    ok: bool
    outcome: str
    payload: object


_known_outcomes: set[str] | None = None


def known_outcomes(wb) -> set[str]:
    """The outcome registry, read once from VBA so it cannot drift from a
    Python mirror."""
    global _known_outcomes
    if _known_outcomes is None:
        _known_outcomes = set(run(wb, "modContract.OutcomeCodes"))
    return _known_outcomes


def run_action(wb, macro: str, *args) -> Result:
    """Call a layer 1 action and validate its envelope before returning it.

    Queries return bare values and must use `run` instead.
    """
    raw = run(wb, macro, *args)
    assert isinstance(raw, (tuple, list)) and len(raw) == 3, (
        f"{macro} returned {raw!r}, not a three-element result"
    )
    ok, outcome, payload = raw
    assert isinstance(ok, bool), f"{macro} returned a non-boolean ok: {ok!r}"
    assert isinstance(outcome, str) and outcome, f"{macro} returned a blank outcome"
    assert outcome in known_outcomes(wb), f"{macro} returned unregistered outcome {outcome!r}"
    return Result(ok, outcome, payload)


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
