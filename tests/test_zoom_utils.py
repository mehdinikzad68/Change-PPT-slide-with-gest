import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from zoom_utils import (
    ZoomController,
    is_pinch,
    pinch_center,
    pinch_distance,
    ZOOM_MAX,
    ZOOM_MIN,
)


class LM:
    def __init__(self, x, y):
        self.x = x
        self.y = y


def make_landmarks(thumb, index):
    landmarks = [LM(0, 0) for _ in range(21)]
    landmarks[4] = LM(*thumb)
    landmarks[8] = LM(*index)
    return landmarks


def test_pinch_distance_and_center():
    landmarks = make_landmarks((0.0, 0.0), (0.3, 0.4))
    assert abs(pinch_distance(landmarks) - 0.5) < 1e-9
    assert pinch_center(landmarks) == (0.15, 0.2)


def test_is_pinch_threshold():
    close = make_landmarks((0.50, 0.50), (0.53, 0.50))
    far = make_landmarks((0.30, 0.50), (0.50, 0.30))  # L-shape spotlight pose
    assert is_pinch(close) is True
    assert is_pinch(far) is False


def test_first_update_only_sets_anchor():
    zoom = ZoomController()
    assert zoom.update(0.02) is None
    assert zoom.level == 100


def test_spreading_pinch_zooms_in_by_steps():
    zoom = ZoomController()
    zoom.update(0.01)               # anchor
    assert zoom.update(0.035) == "zoom_in"
    assert zoom.level == 105
    assert zoom.update(0.06) == "zoom_in"
    assert zoom.level == 110


def test_closing_pinch_zooms_out():
    zoom = ZoomController(initial_level=120)
    zoom.update(0.04)               # anchor
    assert zoom.update(0.015) == "zoom_out"
    assert zoom.level == 115


def test_small_jitter_does_not_zoom():
    zoom = ZoomController()
    zoom.update(0.03)
    assert zoom.update(0.035) is None
    assert zoom.update(0.025) is None
    assert zoom.level == 100


def test_zoom_clamps_at_limits():
    zoom = ZoomController(initial_level=ZOOM_MAX)
    zoom.update(0.01)
    assert zoom.update(0.04) is None
    assert zoom.level == ZOOM_MAX

    zoom = ZoomController(initial_level=ZOOM_MIN)
    zoom.update(0.04)
    assert zoom.update(0.01) is None
    assert zoom.level == ZOOM_MIN


def test_reset_anchor_starts_fresh():
    zoom = ZoomController()
    zoom.update(0.01)
    zoom.reset_anchor()
    # A large distance after a reset must re-anchor, not fire a step.
    assert zoom.update(0.05) is None
    assert zoom.level == 100
