import pytest

from tests.conftest import run


def test_wide_image_is_limited_by_box_width(wb):
    result = run(wb, "modPinEditor.FitAspectRatio", 800.0, 400.0, 180.0, 180.0)
    width, height = result
    assert width == pytest.approx(180.0)
    assert height == pytest.approx(90.0)


def test_tall_image_is_limited_by_box_height(wb):
    result = run(wb, "modPinEditor.FitAspectRatio", 400.0, 800.0, 180.0, 180.0)
    width, height = result
    assert width == pytest.approx(90.0)
    assert height == pytest.approx(180.0)


def test_square_image_in_square_box_fills_it_exactly(wb):
    result = run(wb, "modPinEditor.FitAspectRatio", 500.0, 500.0, 180.0, 180.0)
    assert tuple(result) == pytest.approx((180.0, 180.0))


@pytest.mark.parametrize(
    "source_w,source_h,box_w,box_h",
    [(0, 100, 180, 180), (100, 0, 180, 180), (100, 100, 0, 180), (100, 100, 180, 0)],
)
def test_invalid_dimensions_return_nothing(wb, source_w, source_h, box_w, box_h):
    result = run(wb, "modPinEditor.FitAspectRatio", float(source_w), float(source_h), float(box_w), float(box_h))
    assert result is None
