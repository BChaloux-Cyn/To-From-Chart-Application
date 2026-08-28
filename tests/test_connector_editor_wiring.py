VBEXT_CT_CLASSMODULE = 2


def module_source(wb, component_name):
    module = wb.VBProject.VBComponents(component_name).CodeModule
    return module.Lines(1, module.CountOfLines)


def test_pin_marker_class_exists(wb):
    names = [wb.VBProject.VBComponents(i + 1).Name for i in range(wb.VBProject.VBComponents.Count)]
    assert "clsPinMarker" in names


def test_pin_marker_is_a_class_module(wb):
    assert wb.VBProject.VBComponents("clsPinMarker").Type == VBEXT_CT_CLASSMODULE


def test_pin_marker_uses_a_small_fixed_badge_size(wb):
    # Manual verification (phase-2-manual-verification.md, 2b) found pin
    # marker labels used the Label control's default (body-text-sized)
    # dimensions, dwarfing the photo. AddMarkerControl must set an explicit
    # small size rather than leaving the control at its default.
    source = module_source(wb, "frmConnectorEditor")
    assert "PIN_MARKER_SIZE" in source
    add_marker = source[source.index("Private Sub AddMarkerControl"):]
    assert "lbl.Width = PIN_MARKER_SIZE" in add_marker
    assert "lbl.Height = PIN_MARKER_SIZE" in add_marker


def test_photo_fit_box_is_a_fixed_constant(wb):
    # Manual verification (phase-2-manual-verification.md, 2b) found that
    # fitting against imgPhoto.Width/Height - which the same code then
    # overwrites with the fitted result - shrinks the preview further on
    # every subsequent Load Photo. The box must come from a fixed constant.
    source = module_source(wb, "frmConnectorEditor")
    assert "PHOTO_BOX_WIDTH" in source and "PHOTO_BOX_HEIGHT" in source
    assert "FitAspectRatio(CDbl(pic.Width), CDbl(pic.Height), imgPhoto.Width, imgPhoto.Height)" not in source
    for call in source.splitlines():
        if "modPinEditor.FitAspectRatio(" in call:
            assert "PHOTO_BOX_WIDTH" in call and "PHOTO_BOX_HEIGHT" in call
