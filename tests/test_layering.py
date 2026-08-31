import re

import pytest

LAYER0 = [
    "modUtil", "modState", "modLibrary", "modChart", "modConnectors",
    "modSnapshot", "modLibraryTransfer", "modPinEditor",
]
LAYER1 = [
    "modContract", "modMessages", "modEditorActions", "modPickerActions",
    "modManageActions", "modConnectorActions",
]
ADAPTERS = ["frmConnectorEditor", "frmConnectorPicker", "frmManageLibrary",
            "frmRemoveConnector", "clsPinMarker", "shHarness", "shConnectors"]
FORBIDDEN_IN_LAYER1 = [
    "MsgBox", "InputBox", "GetOpenFilename", "GetSaveAsFilename",
    ".Show", "Unload", "Workbooks.Open", "DoEvents", "MSForms",
]

# frmConnectorEditor exports a Shape to disk, which needs a live Shape object
# that cannot usefully cross Application.Run. Isolated in one named wrapper.
ALLOWED_LAYER0_IN_ADAPTERS = {
    "frmConnectorEditor": {
        "modPinEditor.FitAspectRatio", "modPinEditor.MarkerTopLeft",
        "modPinEditor.ClearScratchPins",
        "modLibrary.SlugifyConnectorID", "modLibrary.ExportShapeToFile",
    },
    "frmConnectorPicker": {"modLibrary.ConnectorIndex"},
    "frmRemoveConnector": {"modConnectors.InstanceIndex"},
    "frmManageLibrary": {
        "modLibrary.ConnectorIndex", "modLibrary.ReadConnector",
        "modLibrary.LIB_ROW_CAP", "modPinEditor.LoadScratchPins",
        "modLibrary.FindConnectorRow",
    },
    # The drag handler converts pixels to a normalized point and stores it.
    # Both are layer 0 primitives with their own tests; there is no
    # transaction here to lift into an action module.
    "clsPinMarker": {"modPinEditor.NormFromMarker", "modPinEditor.MoveMarker"},
    "shHarness": {"modChart.ApplyHarnessEdit"},
    "shConnectors": {"modConnectors.CONN_FIRST_ROW", "modConnectors.ApplyConnectorEdit"},
}

# Handlers that legitimately do no domain work: they only unload, or they
# seed scratch state and hand off to another form.
NON_DELEGATING_HANDLERS = {"cmdCancel_Click", "cmdClose_Click", "cmdEdit_Click"}


def module_source(wb, component_name):
    module = wb.VBProject.VBComponents(component_name).CodeModule
    return module.Lines(1, module.CountOfLines)


def all_sources(wb):
    return {name: module_source(wb, name) for name in LAYER0 + LAYER1 + ADAPTERS}


@pytest.mark.parametrize("module", LAYER1)
def test_action_modules_open_no_dialogs(wb, module):
    source = module_source(wb, module)
    for token in FORBIDDEN_IN_LAYER1:
        assert token not in source, f"{module} references {token}, which belongs in an adapter"


@pytest.mark.parametrize("adapter", ADAPTERS)
def test_adapters_call_only_permitted_layer0_members(wb, adapter):
    source = module_source(wb, adapter)
    allowed = ALLOWED_LAYER0_IN_ADAPTERS.get(adapter, set())
    for module in LAYER0:
        for match in re.finditer(rf"\b{module}\.(\w+)", source):
            reference = f"{module}.{match.group(1)}"
            assert reference in allowed, (
                f"{adapter} calls {reference} directly; route it through an action module "
                f"or add it to ALLOWED_LAYER0_IN_ADAPTERS with a reason"
            )


def test_no_doevents_anywhere(wb):
    for name, source in all_sources(wb).items():
        assert "DoEvents" not in source, f"{name} calls DoEvents, which lets a form unload mid-action"


@pytest.mark.parametrize("adapter", ADAPTERS)
def test_nothing_follows_unload_me(wb, adapter):
    # After Unload Me the controls are gone and form-level variables are
    # reset; reading either afterwards silently loses data.
    source = module_source(wb, adapter)
    for sub in re.split(r"\n(?=(?:Private|Public)\s+Sub\s)", source):
        lines = sub.splitlines()
        indexes = [i for i, line in enumerate(lines) if line.strip() == "Unload Me"]
        if not indexes:
            continue
        for line in lines[indexes[0] + 1:]:
            body = line.split("'")[0].strip()
            if not body or body in ("End Sub", "End If", "Else", "Exit Sub"):
                continue
            assert not re.search(r"\bm[A-Z]\w*", body), (
                f"{adapter}: '{body}' reads form state after Unload Me"
            )
            assert not re.search(r"\b(txt|lst|cbo|cmd|img|tgl)\w*\.", body), (
                f"{adapter}: '{body}' reads a control after Unload Me"
            )


@pytest.mark.parametrize("adapter", ["frmConnectorEditor", "frmConnectorPicker",
                                      "frmManageLibrary", "frmRemoveConnector"])
def test_every_click_handler_delegates(wb, adapter):
    source = module_source(wb, adapter)
    for sub in re.split(r"\n(?=(?:Private|Public)\s+Sub\s)", source):
        header = sub.splitlines()[0] if sub.splitlines() else ""
        match = re.search(r"Sub\s+(cmd\w+_Click)", header)
        if not match:
            continue
        if match.group(1) in NON_DELEGATING_HANDLERS:
            continue
        assert re.search(r"\bmod(Editor|Picker|Manage|Connector)Actions\.", sub), \
            f"{adapter}: {match.group(1)} does no work through an action module"


def test_no_option_base_directive(wb):
    # Option Base 1 would make Array() one based, silently shifting every
    # result envelope index apart from a COM caller's view of it.
    for name, source in all_sources(wb).items():
        assert "Option Base" not in source, f"{name} declares Option Base"
