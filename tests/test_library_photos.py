from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def module_source(wb, component_name):
    module = wb.VBProject.VBComponents(component_name).CodeModule
    return module.Lines(1, module.CountOfLines)


def test_embed_photo_adds_a_named_shape(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws = library_wb.Worksheets("Photos")

    name = run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))

    assert name == "PHOTO_DTM-04P"
    assert ws.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"


def test_embed_photo_with_missing_file_returns_empty(wb, library_wb):
    ws = library_wb.Worksheets("Photos")
    name = run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", r"C:\no\such\file.png")
    assert name == ""


def test_embed_photo_with_no_new_path_keeps_the_existing_one(wb, library_wb, tmp_path):
    # Manual verification (phase-2-manual-verification.md, 2c) found that
    # editing a connector's fields/pins without picking a new photo failed
    # the ENTIRE save silently: LoadExistingPhoto's re-export of the
    # existing photo is unreliable (ExportShapeToFile's Chart.Paste
    # mechanism), leaving mPhotoPath blank, which previously made
    # EmbedConnectorPhoto - and so the whole save - fail outright.
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws = library_wb.Worksheets("Photos")
    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))
    before_count = ws.Shapes.Count

    name = run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", "")

    assert name == "PHOTO_DTM-04P"
    assert ws.Shapes.Count == before_count
    assert ws.Shapes("PHOTO_DTM-04P").Name == "PHOTO_DTM-04P"


def test_embed_photo_replaces_an_existing_shape_for_the_same_id(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws = library_wb.Worksheets("Photos")

    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))
    before = ws.Shapes.Count
    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))

    assert ws.Shapes.Count == before


def test_second_photo_lands_in_a_different_grid_slot(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws = library_wb.Worksheets("Photos")

    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))
    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "GND-STUD", str(photo_path))

    first = ws.Shapes("PHOTO_DTM-04P")
    second = ws.Shapes("PHOTO_GND-STUD")
    assert (first.Left, first.Top) != (second.Left, second.Top)


def test_remove_photo_deletes_the_shape(wb, library_wb, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    ws = library_wb.Worksheets("Photos")
    run(wb, "modLibrary.EmbedConnectorPhoto", ws, "DTM-04P", str(photo_path))

    run(wb, "modLibrary.RemoveConnectorPhoto", ws, "DTM-04P")

    assert ws.Shapes.Count == 0


def test_remove_photo_for_unknown_id_does_not_raise(wb, library_wb):
    ws = library_wb.Worksheets("Photos")
    run(wb, "modLibrary.RemoveConnectorPhoto", ws, "NO-SUCH-ID")


def test_cache_photo_path_creates_the_photos_subfolder(wb, tmp_path):
    result = run(wb, "modLibrary.CachePhotoPath", str(tmp_path), "DTM-04P")
    assert result == str(tmp_path / "Photos" / "DTM-04P.png")
    assert (tmp_path / "Photos").is_dir()


def test_cache_photo_path_accepts_an_extension_override(wb, tmp_path):
    # frmConnectorEditor's editor-preview cache uses "jpg" (LoadPicture's
    # legacy OLE loader unreliably rejects valid PNGs - see 2b), distinct
    # from modSnapshot's "png" default for the harness-chart cache.
    result = run(wb, "modLibrary.CachePhotoPath", str(tmp_path), "DTM-04P", "jpg")
    assert result == str(tmp_path / "Photos" / "DTM-04P.jpg")


# ExportShapeToFile's Shape.Copy/Chart.Paste/Chart.Export mechanism is the
# other clipboard-dependent operation in this codebase (like
# CopyConnectorPhoto - see phase-2-manual-verification.md, 2d). Manual
# verification (2b/2c) found Chart.Paste can silently paste nothing here
# regardless of automation vs. a real interactive session, so - matching
# CopyConnectorPhoto's precedent - these are structural, not behavioral:
# a live Copy/Paste assertion would be no more reliable in CI than it was
# by hand.
def test_export_shape_to_file_verifies_the_paste_actually_landed(wb):
    # Chart.Paste can silently paste nothing - no error, no indication -
    # leaving a blank exported image. ExportShapeToFile must check
    # Chart.Shapes.Count before trusting the export, not just check that a
    # file landed on disk (Chart.Export writes a file either way).
    source = module_source(wb, "modLibrary")
    assert "cht.Chart.Shapes.Count" in source
    assert "bPasted" in source


def test_export_shape_to_file_activates_a_hidden_host_sheet(wb):
    # Chart.Paste's clipboard target must be the ActiveSheet, but a very
    # hidden sheet (_Edit, _Snapshot) can never become active - exporting a
    # shape that lives on one must temporarily unhide + activate its sheet.
    source = module_source(wb, "modLibrary")
    assert "xlSheetVisible" in source
    assert "wsHost.Activate" in source
