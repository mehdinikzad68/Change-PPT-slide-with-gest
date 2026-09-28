from powerpoint_runtime import PowerPointProcessCache, resolve_active_presentation, resolve_slideshow_view


class Clock:
    def __init__(self):
        self.current = 0.0

    def advance(self, seconds):
        self.current += seconds

    def now(self):
        return self.current


class FakePresentations:
    def __init__(self, count):
        self.Count = count


class FakeSlide:
    def __init__(self, slide_index):
        self.SlideIndex = slide_index


class FakeView:
    def __init__(self, slide_index=1):
        self.Slide = FakeSlide(slide_index)


class FakeSlideShowWindows:
    def __init__(self, count, view=None):
        self.Count = count
        self._view = view or FakeView()

    def __call__(self, index):
        assert index == 1
        return type("Window", (), {"View": self._view})()


class FakePowerPointApp:
    def __init__(self, presentation_count=1, slideshow_count=1):
        self.ActivePresentation = object()
        self.Presentations = FakePresentations(presentation_count)
        self.SlideShowWindows = FakeSlideShowWindows(slideshow_count)


def test_powerpoint_process_cache_reuses_recent_result():
    calls = []
    clock = Clock()

    def checker():
        calls.append(clock.now())
        return True

    cache = PowerPointProcessCache(checker=checker, ttl_seconds=1.5, now_func=clock.now)

    assert cache.is_running() is True
    assert cache.is_running() is True
    assert calls == [0.0]

    clock.advance(1.0)
    assert cache.is_running() is True
    assert calls == [0.0]

    clock.advance(0.6)
    assert cache.is_running() is True
    assert calls == [0.0, 1.6]


def test_resolve_active_presentation_returns_none_without_presentations():
    powerpoint = FakePowerPointApp(presentation_count=0)
    assert resolve_active_presentation(powerpoint) is None


def test_resolve_active_presentation_returns_active_presentation_when_available():
    powerpoint = FakePowerPointApp(presentation_count=1)
    assert resolve_active_presentation(powerpoint) is powerpoint.ActivePresentation


def test_resolve_slideshow_view_returns_none_without_slideshow_window():
    powerpoint = FakePowerPointApp(slideshow_count=0)
    assert resolve_slideshow_view(powerpoint) is None


def test_resolve_slideshow_view_returns_first_window_view():
    view = FakeView(slide_index=3)
    powerpoint = FakePowerPointApp(slideshow_count=1)
    powerpoint.SlideShowWindows = FakeSlideShowWindows(count=1, view=view)

    assert resolve_slideshow_view(powerpoint) is view
