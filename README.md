# Change PPT Slide with Gestures 🖐️📊

An AI-powered computer vision application that enables **touchless, real-time control of Microsoft PowerPoint presentations using hand gestures** via your webcam.

Built using **OpenCV**, **Google MediaPipe Hands**, and **Windows COM Automation (`pywin32`)**.

The main controller runs invisibly in the background — no camera preview window — with a transparent on-screen HUD, a system-tray icon, and AR tools (laser pointer and spotlight) drawn directly over your slides.

---

## 🌟 Features

- **Touchless Presentation Control**: Start/end the slideshow and move between slides without clickers or a keyboard.
- **Real-Time Hand Landmark Tracking**: Google MediaPipe's 21-landmark hand model (single hand, detection confidence 0.7).
- **Engagement “Clutch”**: The system stays asleep until you deliberately wake it, so casual hand movements never change a slide by accident.
- **Multithreaded Pipeline**: Separate camera, AI, and action threads keep gesture latency low; frame queues drop stale frames instead of lagging behind.
- **Landmark Smoothing**: Exponential Moving Average (EMA, α = 0.5) on hand landmarks reduces jitter.
- **Invisible HUD + System Tray**: A full-screen, transparent, click-through PyQt5 overlay shows short status popups (`AWAKE`, `NEXT SLIDE ->`, …); a `pystray` tray icon provides status and Quit.
- **AR Presentation Tools**: A red laser-pointer dot and a spotlight effect follow your index fingertip over the slides.
- **PowerPoint-Aware**: AI processing pauses while PowerPoint (`POWERPNT.EXE`) is not running; the process check is cached (1.5 s TTL) instead of scanning processes on every frame.
- **Thread-Safe COM Handling**: All PowerPoint COM calls live inside the ActionThread's own COM apartment (`CoInitialize` / `CoUninitialize`), and the active presentation / slideshow view is re-resolved for every command — so switching files or restarting the slideshow works without restarting the app.
- **Robust Gesture Recognition**: A gesture must be held for 5 consecutive frames before it fires, preventing accidental triggers.
- **Optional Telemetry**: With `--telemetry <path>`, every frame's timing (capture → inference → decision → COM), FPS, gesture, command, and errors are written as JSONL for latency analysis.
- **Pinch Zoom & Pan**: While a slideshow is running, pinch thumb and index together and spread them apart to zoom in, or close them to zoom out (5% steps, 10%–400%). Dragging a steady pinch while zoomed in pans the view.

---

## 🖐️ Gesture Reference Guide

### Main controller (`change_pptslide_with_gest.py`)

First, **wake the system**: hold an open palm until the HUD shows `AWAKE`. While the system is asleep, all other gestures are ignored.

| Gesture | Action | Description |
| :--- | :--- | :--- |
| ✋ **Open Palm (hold)** | **Wake / Start Slideshow** | Hold an open palm (all 5 fingers extended) for 15 consecutive frames to wake the system. Once awake, holding the open palm for 5 more frames starts the slideshow. |
| ✊ **Closed Fist** | **End Slideshow + Sleep** | Close all fingers into a fist for 5 frames: the system goes back to sleep (`ASLEEP`) and, if a slideshow is running, exits it. |
| ☝️ **Index Swipe Right** | **Next Slide** | Point the index finger and swipe horizontally to the right. |
| ✌️ **Two Fingers Swipe Left** | **Previous Slide** | Extend index + middle fingers and swipe to the left. |
| ☝️ **Index Finger Only** | **Laser Pointer** | While awake, hold up only the index finger: a red dot follows your fingertip (mirrored, like a mirror image) on the HUD overlay. |
| 🤏 **Thumb + Index (L-Shape)** | **Spotlight** | While awake, extend thumb + index: the screen darkens except for a circular spotlight around your fingertip. |
| 🤏 **Pinch (tips touching)** | **Zoom In / Out** | While a slideshow is running, pinch thumb and index tips together and spread/close them to zoom (see Zoom below). |

Slide navigation is bounded: “next” does nothing on the last slide and “previous” does nothing on the first slide.

### Zoom (main controller, slideshow running)

| Gesture | Action | Description |
| :--- | :--- | :--- |
| 🤏 **Pinch & Spread / Close** | **Zoom In / Out** | Touch thumb tip to index tip (a pinch), then spread them apart to zoom in or bring them back together to zoom out — 5% per step, between 10% and 400%. The HUD shows the current level (e.g. `ZOOM 150%`). |
| 🤏 **Pinch & Drag** | **Pan Zoomed View** | While zoomed in above 100%, drag a steady pinch to move the view around the slide. |

A pinch is recognised by the two fingertips *touching*, so it never conflicts with the spotlight gesture, where thumb and index form an open L-shape with the tips far apart.

---

## ⚙️ Prerequisites

- **Operating System**: Windows 10 / Windows 11 *(required for the Microsoft Office COM interface via `pywin32`)*
- **Software**: Microsoft PowerPoint installed
- **Hardware**: Standard USB webcam or built-in laptop camera
- **Python**: Python 3.8 or higher

> MediaPipe is pinned to **`mediapipe==0.10.14`** in `requirements.txt`, because the controllers use the classic `mp.solutions.hands` API.

---

## 🚀 Quick Start

### 1. Clone the Repository
```bash
git clone https://github.com/mehdinikzad68/Change-PPT-slide-with-gest.git
cd Change-PPT-slide-with-gest
```

### 2. Set Up a Virtual Environment (Optional but Recommended)
```powershell
python -m venv venv
.\venv\Scripts\activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

---

## 🖥️ How to Use

1. **Open your presentation**: Start Microsoft PowerPoint and open the `.pptx` file you want to present. (The main controller attaches to the running PowerPoint instance — it never creates a hidden one.)
2. **Launch the gesture controller**:
   - For the current controller (recommended):
     ```bash
     python change_pptslide_with_gest.py
     ```
   - With telemetry logging:
     ```bash
     python change_pptslide_with_gest.py --telemetry telemetry.jsonl
     ```
3. **Control your presentation** (main controller):
   - Hold an **open palm** until the HUD shows `AWAKE`.
   - Hold the **open palm** again briefly to start the slideshow.
   - Swipe **index finger right** for the next slide, **two fingers left** for the previous slide.
   - Use **index only** for the laser pointer, or **thumb + index** for the spotlight.
   - **Pinch** thumb and index tips together and spread them apart to zoom in; close them to zoom out; drag a steady pinch to pan while zoomed in.
   - Show a **fist** to end the slideshow and put the system back to sleep.
4. **Exit the program**:
   - Right-click the **system-tray icon** (“Invisible Presentation Assistant”) and choose **Quit**. There is no camera window. (On a headless/no-display run, press `Ctrl+C`.)

---

## 📊 Telemetry

Running the main controller with `--telemetry telemetry.jsonl` writes one JSON object per processed frame, including:

- `t_capture`, `t_infer_start`, `t_infer_end`, `t_decision`, `t_com_start`, `t_com_end`
- Derived timings: `frame_age`, `infer_duration`, `com_duration`
- `fps`, `gesture`, `command` (e.g. `start_slideshow`, `next_slide`, `previous_slide`, `end_slideshow`, `zoom_in`, `zoom_out`, `pan`)
- `engagement_active`, `slide_show_started`, and `error` when a COM call fails

---

## 📁 Project Structure

```
Change-PPT-slide-with-gest/
│
├── change_pptslide_with_gest.py           # Main controller: multithreaded, HUD + tray, clutch, AR tools, zoom, telemetry
├── powerpoint_runtime.py                  # PowerPoint helpers: cached process check + active presentation / slideshow resolvers
├── telemetry_utils.py                     # FPS calculation, telemetry records, thread-safe JSONL writer
├── zoom_utils.py                          # Pure pinch-zoom logic: pinch detection, anchor-based zoom steps, limits
├── tests/
│   ├── test_powerpoint_runtime.py         # Unit tests for the process cache and resolvers
│   ├── test_telemetry_utils.py            # Unit tests for FPS, frame age, and JSONL output
│   └── test_zoom_utils.py                 # Unit tests for pinch detection and zoom steps/clamps
├── requirements.txt                       # Project dependencies (mediapipe pinned to 0.10.14)
├── LICENSE                                # MIT License
├── .gitignore
└── README.md
```

> **Note:** pinch zoom originally lived in a separate standalone script (`change_pptslide_with_gest_with_zoom.py`). It was ported into the main controller (logic in `zoom_utils.py`) and the legacy script was removed; it remains available in the git history before that change.

---

## 🧪 Tests

Pure-logic unit tests (no camera, Windows, or PowerPoint required):

```bash
pytest tests/
```

---

## 🏗️ Architecture (main controller)

```
Webcam ──► CameraThread ──► frame queue ──► AIThread ──► action queue ──► ActionThread ──► PowerPoint (COM)
                                                │                                │
                                                │                                ├─► UI queue ──► PyQt5 HUD (status + AR overlay)
                                                │                                └─► TelemetryWriter (optional JSONL)
                                                └─ pauses while POWERPNT.EXE is not running (cached check, 1.5 s)
```

- **CameraThread** captures frames with capture timestamps; a full queue drops the oldest frame.
- **AIThread** runs MediaPipe Hands and forwards landmarks plus telemetry timestamps.
- **ActionThread** smooths landmarks (EMA), runs the engagement state machine and gesture counters, executes COM commands in its own COM apartment, and feeds the HUD and telemetry.

---

## 🛠️ Built With

- [OpenCV](https://opencv.org/) - Real-time computer vision and camera feed processing
- [Google MediaPipe](https://developers.google.com/mediapipe) - High-fidelity hand and finger tracking (pinned `0.10.14`)
- [pywin32](https://github.com/mhammond/pywin32) - Windows COM client integration with Microsoft Office PowerPoint
- [PyQt5](https://riverbankcomputing.com/software/pyqt/) - Transparent full-screen HUD overlay
- [pystray](https://github.com/moses-palmer/pystray) + Pillow - System-tray icon
- [psutil](https://github.com/giampaolo/psutil) - PowerPoint process detection

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).
