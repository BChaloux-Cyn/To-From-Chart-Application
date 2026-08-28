import pytest

from tests.conftest import run

PHOTO_LEFT, PHOTO_TOP, PHOTO_W, PHOTO_H = 12.0, 30.0, 180.0, 120.0
MARKER = 16.0


def test_marker_top_left_centres_the_badge_on_the_point(wb):
    left, top = run(wb, "modPinEditor.MarkerTopLeft",
                    0.5, 0.5, PHOTO_LEFT, PHOTO_TOP, PHOTO_W, PHOTO_H, MARKER, MARKER)
    # Centre of the photo is (12+90, 30+60); the badge is offset by half its size.
    assert left == pytest.approx(102.0 - 8.0)
    assert top == pytest.approx(90.0 - 8.0)


def test_norm_from_marker_recovers_the_normalized_point(wb):
    norm_x, norm_y = run(wb, "modPinEditor.NormFromMarker",
                         94.0, 82.0, MARKER, MARKER,
                         PHOTO_LEFT, PHOTO_TOP, PHOTO_W, PHOTO_H)
    assert norm_x == pytest.approx(0.5)
    assert norm_y == pytest.approx(0.5)


@pytest.mark.parametrize("norm_x,norm_y", [(0.0, 0.0), (1.0, 1.0), (0.25, 0.75), (0.6111, 0.3333)])
def test_the_two_conversions_round_trip(wb, norm_x, norm_y):
    left, top = run(wb, "modPinEditor.MarkerTopLeft",
                    norm_x, norm_y, PHOTO_LEFT, PHOTO_TOP, PHOTO_W, PHOTO_H, MARKER, MARKER)
    back_x, back_y = run(wb, "modPinEditor.NormFromMarker",
                         left, top, MARKER, MARKER,
                         PHOTO_LEFT, PHOTO_TOP, PHOTO_W, PHOTO_H)
    assert back_x == pytest.approx(norm_x)
    assert back_y == pytest.approx(norm_y)


def test_a_zero_sized_photo_returns_empty_rather_than_dividing_by_zero(wb):
    assert run(wb, "modPinEditor.NormFromMarker",
               94.0, 82.0, MARKER, MARKER, PHOTO_LEFT, PHOTO_TOP, 0.0, PHOTO_H) is None
