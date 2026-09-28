import time


class PowerPointProcessCache:
    def __init__(self, checker, ttl_seconds=1.5, now_func=None):
        self.checker = checker
        self.ttl_seconds = ttl_seconds
        self.now_func = now_func or time.monotonic
        self._last_checked_at = None
        self._last_result = None

    def is_running(self, force=False):
        now = self.now_func()
        if (
            force
            or self._last_checked_at is None
            or now - self._last_checked_at >= self.ttl_seconds
        ):
            self._last_result = bool(self.checker())
            self._last_checked_at = now
        return self._last_result


def resolve_active_presentation(powerpoint_app):
    if powerpoint_app is None or not hasattr(powerpoint_app, "Presentations"):
        return None

    presentations = powerpoint_app.Presentations
    if getattr(presentations, "Count", 0) <= 0:
        return None

    return powerpoint_app.ActivePresentation


def resolve_slideshow_view(powerpoint_app):
    if powerpoint_app is None or not hasattr(powerpoint_app, "SlideShowWindows"):
        return None

    slideshow_windows = powerpoint_app.SlideShowWindows
    if getattr(slideshow_windows, "Count", 0) <= 0:
        return None

    return slideshow_windows(1).View
