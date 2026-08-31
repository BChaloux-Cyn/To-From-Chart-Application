import pytest


def module_source(wb, component_name):
    module = wb.VBProject.VBComponents(component_name).CodeModule
    return module.Lines(1, module.CountOfLines)


def controls(wb, form_name):
    return wb.VBProject.VBComponents(form_name).Designer.Controls


@pytest.mark.parametrize("name", ["lstConnectors", "cmdAdd", "cmdNew", "cmdCancel"])
def test_picker_has_its_controls(wb, name):
    assert controls(wb, "frmConnectorPicker")(name).Name == name


def test_pin_list_does_not_shrink_when_a_pin_is_placed(wb):
    # IntegralHeight defaults to True, which silently shrinks a ListBox's
    # own Height the first time an item is added (rounds down to a whole
    # number of visible rows) - live testing found this made lstPins look
    # like it had moved even though its Top never changed.
    assert controls(wb, "frmConnectorEditor")("lstPins").IntegralHeight is False


@pytest.mark.parametrize("name", ["lstConnectors", "cmdEdit", "cmdDelete", "cmdImport",
                                   "cmdExport", "cmdExportLibrary", "cmdClose"])
def test_manage_library_has_its_controls(wb, name):
    assert controls(wb, "frmManageLibrary")(name).Name == name


@pytest.mark.parametrize("name", ["lstConnectors", "cmdRemove", "cmdCancel"])
def test_remove_connector_has_its_controls(wb, name):
    assert controls(wb, "frmRemoveConnector")(name).Name == name


def test_remove_connector_delegates_through_the_action_module(wb):
    source = module_source(wb, "frmRemoveConnector")
    remove_click = source[source.index("Private Sub cmdRemove_Click"):]
    body = remove_click.split("End Sub", 1)[0]
    assert "modConnectorActions.RemoveInstance" in body


def test_picker_add_calls_add_connector_instance_and_snapshot(wb):
    # modPickerActions.AddFromLibrary (Task 11) is the one place that calls
    # AddConnectorInstance and SnapshotConnector - the picker delegates to it
    # rather than calling those layer 0 primitives directly.
    source = module_source(wb, "frmConnectorPicker")
    assert "modPickerActions.AddFromLibrary" in source


def test_picker_new_launches_the_connector_editor(wb):
    assert "frmConnectorEditor" in module_source(wb, "frmConnectorPicker")


def test_picker_new_chains_into_adding_an_instance_of_what_was_just_saved(wb):
    # Feature added during manual verification (phase-2-manual-verification.md,
    # 2c): creating a brand-new connector via "New..." from the Add
    # Connector dialog should immediately add an instance of it too, rather
    # than requiring a second trip back through the picker.
    source = module_source(wb, "frmConnectorPicker")
    new_click = source[source.index("Private Sub cmdNew_Click"):]
    body = new_click.split("End Sub", 1)[0]
    assert "modConnectorUI.LastSavedConnectorID = \"\"" in body
    assert "modPickerActions.AddFromLibrary" in body
    # Show must come before reading LastSavedConnectorID - it's only set
    # once frmConnectorEditor's Save actually runs.
    assert body.index("frmConnectorEditor.Show") < body.index("modConnectorUI.LastSavedConnectorID")


def test_connector_editor_save_records_the_saved_id_for_callers(wb):
    source = module_source(wb, "frmConnectorEditor")
    save_click = source[source.index("Private Sub cmdSave_Click"):]
    body = save_click.split("End Sub", 1)[0]
    # sID is mConnectorID captured into a local before the save action call,
    # per the layer 1/adapter split (Task 10): every control value and form
    # variable is read before the call, so nothing after Unload Me needs to
    # re-read form state.
    assert "modConnectorUI.LastSavedConnectorID = sID" in body


def test_manage_library_edit_launches_the_connector_editor(wb):
    source = module_source(wb, "frmManageLibrary")
    assert "frmConnectorEditor" in source
    assert "LoadScratchPins" in source


def test_manage_library_edit_passes_connector_fields_to_the_editor(wb):
    # cmdEdit_Click must read the connector's fields (vFields, via
    # ReadConnector) and the Photos sheet, and hand both to
    # frmConnectorEditor.LoadForEdit while mLibrary is still open. The photo
    # preview normally comes from an on-disk cache (modLibrary.CachePhotoPath),
    # but LoadExistingPhoto falls back to a live Shape read from Photos as a
    # one-time backfill when no cache exists yet.
    assert "LoadForEdit" in module_source(wb, "frmManageLibrary")


def test_manage_library_edit_shows_the_editor_before_unloading_itself(wb):
    # Manual verification (phase-2-manual-verification.md, 2c) found that
    # unloading frmManageLibrary (Unload Me) before showing frmConnectorEditor
    # - while still inside frmManageLibrary's own click handler - discarded
    # everything LoadForEdit had just populated (frmConnectorEditor came up
    # blank, UserForm_Initialize visibly re-firing). Show must come first.
    source = module_source(wb, "frmManageLibrary")
    click = source[source.index("Private Sub cmdEdit_Click"):]
    body = click.split("End Sub", 1)[0]
    assert body.index("frmConnectorEditor.Show") < body.index("Unload Me")


def test_connector_editor_supports_loading_an_existing_connector(wb):
    source = module_source(wb, "frmConnectorEditor")
    assert "Public Sub LoadForEdit" in source
    # RebuildPinList (Task 10) is the sole rebuild path - list box and
    # markers are both derived from the scratch sheet, not a parallel
    # in-memory collection.
    assert "RebuildPinList" in source


def test_manage_library_delete_calls_library_delete_functions(wb):
    # modManageActions.DeleteFromLibrary (Task 13) is the one place that
    # calls DeleteConnector/DeletePinsForConnector/RemoveConnectorPhoto and
    # removes the editor's on-disk preview cache - the form delegates to it
    # rather than calling those layer 0 primitives directly.
    source = module_source(wb, "frmManageLibrary")
    delete_click = source[source.index("Private Sub cmdDelete_Click"):]
    body = delete_click.split("End Sub", 1)[0]
    assert "modManageActions.DeleteFromLibrary" in body


@pytest.mark.parametrize("form_name", ["frmConnectorPicker", "frmManageLibrary", "frmRemoveConnector"])
def test_controls_fit_within_the_form_width(wb, form_name):
    # Manual verification (phase-2-manual-verification.md, 2c) found
    # frmManageLibrary's form too narrow to show all 5 buttons - Export and
    # Close rendered past the form's right edge (a regression from when
    # Import/Export were added without widening the form to match).
    component = wb.VBProject.VBComponents(form_name)
    form_width = component.Properties("Width").Value
    for ctl in component.Designer.Controls:
        assert ctl.Left + ctl.Width <= form_width, \
            f"{form_name}.{ctl.Name} extends past the form's right edge ({ctl.Left + ctl.Width} > {form_width})"


def test_home_has_the_three_new_buttons(wb):
    shapes = wb.Worksheets("Home").Shapes
    actions = [shapes(i + 1).OnAction for i in range(shapes.Count)]
    assert "modConnectorUI.ShowAddConnector" in actions
    assert "modConnectorUI.ShowManageLibrary" in actions
    assert "modConnectorUI.ShowRemoveConnector" in actions
