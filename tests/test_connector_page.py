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


def test_place_callouts_draws_one_oval_per_pin(wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        run(wb, "modConnectorPage.PlacePhoto", ws, str(photo_path))
        shp = ws.Shapes("PAGE_PHOTO")

        pins = (
            ("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1),
            ("DTM-04P", 2, "GND", 0.9, 0.1, 0.9, 0.1),
        )
        n = run(wb, "modConnectorPage.PlaceCallouts", ws, shp, pins)
        assert n == 2
        assert ws.Shapes("PIN_1").Name == "PIN_1"
        assert ws.Shapes("PIN_2").Name == "PIN_2"
        assert ws.Shapes("PIN_1").TextFrame2.TextRange.Text == "1"
    finally:
        dest.Close(SaveChanges=False)


def test_place_callouts_centers_the_oval_on_the_label_position(wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        run(wb, "modConnectorPage.PlacePhoto", ws, str(photo_path))
        shp = ws.Shapes("PAGE_PHOTO")

        pins = (("DTM-04P", 1, "", 0.5, 0.5, 0.5, 0.5),)
        run(wb, "modConnectorPage.PlaceCallouts", ws, shp, pins)

        oval = ws.Shapes("PIN_1")
        expected_center_x = shp.Left + 0.5 * shp.Width
        expected_center_y = shp.Top + 0.5 * shp.Height
        assert abs((oval.Left + oval.Width / 2) - expected_center_x) < 0.5
        assert abs((oval.Top + oval.Height / 2) - expected_center_y) < 0.5
    finally:
        dest.Close(SaveChanges=False)


def test_leader_line_drawn_only_when_marker_is_pulled_off_anchor(wb, app, tmp_path):
    photo_path = write_sample_photo(tmp_path / "photo.png")
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        run(wb, "modConnectorPage.PlacePhoto", ws, str(photo_path))
        shp = ws.Shapes("PAGE_PHOTO")

        pins = (
            ("DTM-04P", 1, "", 0.1, 0.1, 0.1, 0.1),   # marker on anchor: no leader
            ("DTM-04P", 2, "", 0.9, 0.1, 0.3, 0.6),   # marker pulled away: leader
        )
        run(wb, "modConnectorPage.PlaceCallouts", ws, shp, pins)
        run(wb, "modConnectorPage.PlaceLeaderLines", ws, shp, pins)

        names = [ws.Shapes(i + 1).Name for i in range(ws.Shapes.Count)]
        assert "LEADER_1" not in names
        assert "LEADER_2" in names
    finally:
        dest.Close(SaveChanges=False)
