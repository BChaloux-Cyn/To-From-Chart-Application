from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


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
