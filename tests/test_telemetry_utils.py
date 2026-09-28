import json

import pytest

from telemetry_utils import TelemetryWriter, build_telemetry_record, compute_fps


def test_compute_fps_from_capture_times():
    assert compute_fps(10.5, 10.0) == pytest.approx(2.0)


def test_compute_fps_handles_missing_or_non_increasing_timestamps():
    assert compute_fps(10.0, None) is None
    assert compute_fps(10.0, 10.0) is None
    assert compute_fps(9.9, 10.0) is None


def test_build_telemetry_record_includes_frame_age_and_durations():
    record = build_telemetry_record(
        t_capture=10.0,
        t_infer_start=10.01,
        t_infer_end=10.04,
        t_decision=10.07,
        t_com_start=10.08,
        t_com_end=10.11,
        fps=20.0,
        gesture="palm_open",
        command="start_slideshow",
    )

    assert record["frame_age"] == pytest.approx(0.07)
    assert record["infer_duration"] == pytest.approx(0.03)
    assert record["com_duration"] == pytest.approx(0.03)
    assert record["fps"] == pytest.approx(20.0)
    assert record["gesture"] == "palm_open"
    assert record["command"] == "start_slideshow"


def test_telemetry_writer_writes_valid_jsonl(tmp_path):
    telemetry_path = tmp_path / "telemetry.jsonl"
    writer = TelemetryWriter(str(telemetry_path))

    writer.write(build_telemetry_record(
        t_capture=1.0,
        t_infer_start=1.01,
        t_infer_end=1.02,
        t_decision=1.03,
        fps=30.0,
        gesture="index_right",
        command="next_slide",
    ))
    writer.close()

    lines = telemetry_path.read_text(encoding="utf-8").splitlines()
    assert len(lines) == 1

    payload = json.loads(lines[0])
    assert payload["gesture"] == "index_right"
    assert payload["command"] == "next_slide"
    assert payload["frame_age"] == pytest.approx(0.03)
