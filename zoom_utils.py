import math


# A pinch is active when the thumb tip (landmark 4) and index tip
# (landmark 8) are this close together, in normalized image coordinates.
# This is the same closeness threshold the legacy zoom script used to
# consider the thumb closed against the index finger. It also keeps the
# pinch distinct from the spotlight gesture (thumb + index in an open
# L-shape, where the two tips are far apart).
PINCH_DISTANCE_THRESHOLD = 0.05

# Each change of this much in pinch distance, relative to the anchor
# distance captured when the last step fired, triggers one zoom step.
ZOOM_STEP_DISTANCE = 0.02

# Slideshow zoom limits and step size, carried over from the legacy
# zoom script (View.Zoom is a percentage).
ZOOM_MIN = 10
ZOOM_MAX = 400
ZOOM_STEP_PERCENT = 5


def pinch_distance(landmarks):
    """Normalized distance between thumb tip (4) and index tip (8)."""
    return math.hypot(
        landmarks[4].x - landmarks[8].x,
        landmarks[4].y - landmarks[8].y,
    )


def pinch_center(landmarks):
    """Midpoint between thumb tip (4) and index tip (8)."""
    return (
        (landmarks[4].x + landmarks[8].x) / 2,
        (landmarks[4].y + landmarks[8].y) / 2,
    )


def is_pinch(landmarks):
    return pinch_distance(landmarks) < PINCH_DISTANCE_THRESHOLD


class ZoomController:
    """Stateful pinch-zoom logic, with no COM or UI dependencies.

    Feed it the current pinch distance once per frame while a pinch is
    active; it returns "zoom_in" or "zoom_out" when a step should be
    applied, otherwise None. Call reset_anchor() when the pinch ends so
    the next pinch starts from a fresh anchor.
    """

    def __init__(self, initial_level=100):
        self.level = max(ZOOM_MIN, min(ZOOM_MAX, initial_level))
        self._anchor_distance = None

    def reset_anchor(self):
        self._anchor_distance = None

    def update(self, distance):
        if self._anchor_distance is None:
            self._anchor_distance = distance
            return None

        delta = distance - self._anchor_distance
        if delta >= ZOOM_STEP_DISTANCE:
            self._anchor_distance = distance
            if self.level < ZOOM_MAX:
                self.level = min(ZOOM_MAX, self.level + ZOOM_STEP_PERCENT)
                return "zoom_in"
            return None

        if delta <= -ZOOM_STEP_DISTANCE:
            self._anchor_distance = distance
            if self.level > ZOOM_MIN:
                self.level = max(ZOOM_MIN, self.level - ZOOM_STEP_PERCENT)
                return "zoom_out"
            return None

        return None
