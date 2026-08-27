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
    # EnsureDispatch (early-bound) is required, not Dispatch: late-bound
    # dispatch fails to resolve named arguments like Worksheets.Add(After=...),
    # silently inserting at the front instead.
    app = win32.gencache.EnsureDispatch("Excel.Application")
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


VBEXT_CT_MSFORM = 3


def add_userform(wb, name: str):
    """Add a UserForm component, named, returning its Designer surface -
    the object controls are added to."""
    component = wb.VBProject.VBComponents.Add(VBEXT_CT_MSFORM)
    component.Name = name
    return component.Designer


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
