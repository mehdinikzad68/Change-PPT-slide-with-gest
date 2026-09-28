import json
import threading


def compute_fps(current_capture_time, previous_capture_time):
    if previous_capture_time is None:
        return None

    delta = current_capture_time - previous_capture_time
    if delta <= 0:
        return None

    return 1.0 / delta


def build_telemetry_record(
    *,
    t_capture,
    t_infer_start=None,
    t_infer_end=None,
    t_decision=None,
    t_com_start=None,
    t_com_end=None,
    fps=None,
    gesture=None,
    command=None,
):
    record = {
        "t_capture": t_capture,
        "t_infer_start": t_infer_start,
        "t_infer_end": t_infer_end,
        "t_decision": t_decision,
        "t_com_start": t_com_start,
        "t_com_end": t_com_end,
        "fps": fps,
        "gesture": gesture,
        "command": command,
    }

    if t_capture is not None and t_decision is not None:
        record["frame_age"] = t_decision - t_capture

    if t_infer_start is not None and t_infer_end is not None:
        record["infer_duration"] = t_infer_end - t_infer_start

    if t_com_start is not None and t_com_end is not None:
        record["com_duration"] = t_com_end - t_com_start

    return {key: value for key, value in record.items() if value is not None}


class TelemetryWriter:
    def __init__(self, path=None):
        self.path = path
        self._lock = threading.Lock()
        self._file = open(path, "a", encoding="utf-8", buffering=1) if path else None

    @property
    def enabled(self):
        return self._file is not None

    def write(self, record):
        if not self.enabled:
            return

        with self._lock:
            self._file.write(json.dumps(record) + "\n")

    def close(self):
        if not self.enabled:
            return

        with self._lock:
            self._file.close()
            self._file = None
