from tests.conftest import run
from tests.fixtures.sample_photo import write_sample_photo


def test_page_photo_path_prefers_jpg_over_png(wb, tmp_path):
    (tmp_path / "Photos").mkdir()
    write_sample_photo(tmp_path / "Photos" / "DTM-04P.png")
    (tmp_path / "Photos" / "DTM-04P.jpg").write_bytes((tmp_path / "Photos" / "DTM-04P.png").read_bytes())

    result = run(wb, "modConnectorPage.PagePhotoPath", str(tmp_path), "DTM-04P")
    assert result.endswith("DTM-04P.jpg")


def test_page_photo_path_falls_back_to_png(wb, tmp_path):
    write_sample_photo(tmp_path / "Photos" / "DTM-04P.png")
    result = run(wb, "modConnectorPage.PagePhotoPath", str(tmp_path), "DTM-04P")
    assert result.endswith("DTM-04P.png")


def test_page_photo_path_returns_empty_when_no_cache_exists(wb, tmp_path):
    result = run(wb, "modConnectorPage.PagePhotoPath", str(tmp_path), "NOPE")
    assert result == ""


def test_place_photo_adds_a_named_shape_at_the_fixed_anchor(wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        ok = run(wb, "modConnectorPage.PlacePhoto", ws, str(photo_path))
        assert ok is True
        shp = ws.Shapes("PAGE_PHOTO")
        assert abs(shp.Left - 20) < 0.5
        assert abs(shp.Top - 60) < 0.5
        assert shp.Width <= 300 and shp.Height <= 300
    finally:
        dest.Close(SaveChanges=False)


def test_place_photo_returns_false_for_a_missing_file(wb, app):
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        ok = run(wb, "modConnectorPage.PlacePhoto", ws, r"C:\no\such\file.png")
        assert ok is False
        assert ws.Shapes.Count == 0
    finally:
        dest.Close(SaveChanges=False)
