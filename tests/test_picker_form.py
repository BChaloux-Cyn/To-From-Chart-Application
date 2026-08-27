import pytest


def module_source(wb, component_name):
    module = wb.VBProject.VBComponents(component_name).CodeModule
    return module.Lines(1, module.CountOfLines)


def controls(wb, form_name):
    return wb.VBProject.VBComponents(form_name).Designer.Controls


@pytest.mark.parametrize("name", ["lstConnectors", "cmdAdd", "cmdNew", "cmdCancel"])
def test_picker_has_its_controls(wb, name):
    assert controls(wb, "frmConnectorPicker")(name).Name == name


@pytest.mark.parametrize("name", ["lstConnectors", "cmdEdit", "cmdDelete", "cmdImport", "cmdExport", "cmdClose"])
def test_manage_library_has_its_controls(wb, name):
    assert controls(wb, "frmManageLibrary")(name).Name == name


def test_picker_add_calls_add_connector_instance_and_snapshot(wb):
    source = module_source(wb, "frmConnectorPicker")
    assert "modConnectors.AddConnectorInstance" in source
    assert "modSnapshot.SnapshotConnector" in source


def test_picker_new_launches_the_connector_editor(wb):
    assert "frmConnectorEditor" in module_source(wb, "frmConnectorPicker")


def test_manage_library_edit_launches_the_connector_editor(wb):
    source = module_source(wb, "frmManageLibrary")
    assert "frmConnectorEditor" in source
    assert "LoadScratchPins" in source


def test_manage_library_edit_passes_connector_fields_before_closing_the_library(wb):
    # cmdEdit_Click must read the connector's fields and hand them to
    # frmConnectorEditor.LoadForEdit while mLibrary is still open - the
    # library workbook is closed right after, so LoadForEdit has no second
    # chance to read from it.
    assert "LoadForEdit" in module_source(wb, "frmManageLibrary")


def test_connector_editor_supports_loading_an_existing_connector(wb):
    source = module_source(wb, "frmConnectorEditor")
    assert "Public Sub LoadForEdit" in source
    assert "RebuildPinListFromScratch" in source


def test_manage_library_delete_calls_library_delete_functions(wb):
    source = module_source(wb, "frmManageLibrary")
    assert "modLibrary.DeleteConnector" in source
    assert "modLibrary.DeletePinsForConnector" in source


def test_manage_library_import_export_are_unwired_for_now(wb):
    source = module_source(wb, "frmManageLibrary")
    assert "cmdImport_Click" not in source
    assert "cmdExport_Click" not in source


def test_home_has_the_three_new_buttons(wb):
    shapes = wb.Worksheets("Home").Shapes
    actions = [shapes(i + 1).OnAction for i in range(shapes.Count)]
    assert "modConnectorUI.ShowAddConnector" in actions
    assert "modConnectorUI.ShowManageLibrary" in actions
    assert "modConnectorUI.ShowRemoveConnector" in actions
