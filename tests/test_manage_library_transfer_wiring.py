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
