def test_a_userform_component_exists(wb):
    names = [wb.VBProject.VBComponents(i + 1).Name for i in range(wb.VBProject.VBComponents.Count)]
    assert "frmSmokeTest" in names


def test_userform_has_the_expected_caption(wb):
    form = wb.VBProject.VBComponents("frmSmokeTest").Designer
    assert form.Caption == "Smoke Test"
