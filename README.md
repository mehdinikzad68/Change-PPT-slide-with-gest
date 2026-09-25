# Change PPT Slide with Gestures 🖐️📊

An AI-powered computer vision application that enables **touchless, real-time control of Microsoft PowerPoint presentations using hand gestures** via your webcam. 

Built using **OpenCV**, **Google MediaPipe Hands**, and **Windows COM Automation (`pywin32`)**.

---

## 🌟 Features

- **Touchless Presentation Control**: Advance or rewind slides without physical clickers or keyboards.
- **Real-Time Hand Landmark Tracking**: Utilizes Google MediaPipe's 21 3D hand landmark model for fast and accurate detection.
- **Robust Gesture Recognition**: Debounced consecutive frame validation prevents accidental slide triggers.
- **Two Controller Modes**:
  1. `change_pptslide_with_gest.py`: Lightweight controller for slide navigation and presentation lifecycle.
  2. `change_pptslide_with_gest_with_zoom.py`: Advanced controller including interactive slide zooming and pinching gestures.
- **Live Visual Feedback**: On-screen camera window overlaying tracked skeletal joints and landmarks.

---

## 🖐️ Gesture Reference Guide

| Gesture | Action | Compatible Scripts | Description |
| :--- | :--- | :--- | :--- |
| ✋ **Open Palm** | **Start Slideshow** | Both | Show an open hand with all 5 fingers extended |
| ✊ **Closed Fist** | **End Slideshow** | Both | Close all fingers into a tight fist |
| ☝️ **Index Swipe Right** | **Next Slide** | Both | Point index finger and swipe horizontally to the right |
| ✌️ **Two Fingers Swipe Left** | **Previous Slide** | Both | Extend index + middle fingers and swipe to the left |
| 🤏 **Pinch & Move** | **Zoom In / Out** | Zoom Version | Pinch thumb and index finger to scale/zoom active slide view |

---

## ⚙️ Prerequisites

- **Operating System**: Windows 10 / Windows 11 *(required for Microsoft Office COM interface via `pywin32`)*
- **Software**: Microsoft PowerPoint installed
- **Hardware**: Standard USB webcam or built-in laptop camera
- **Python**: Python 3.8 or higher

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

1. **Open your presentation**: Start Microsoft PowerPoint and open the `.pptx` file you want to present.
2. **Launch the gesture controller**:
   - For standard navigation:
     ```bash
     python change_pptslide_with_gest.py
     ```
   - For navigation + zoom features:
     ```bash
     python change_pptslide_with_gest_with_zoom.py
     ```
3. **Control your presentation**:
   - Show an **open palm** to initiate slideshow mode.
   - Swipe **index finger right** to move to the next slide.
   - Swipe **two fingers left** to return to the previous slide.
   - Show a **fist** to exit slideshow mode.
4. **Exit the program**: Focus on the camera preview window and press `ESC`.

---

## 📁 Project Structure

```
Change-PPT-slide-with-gest/
│
├── change_pptslide_with_gest.py           # Standard hand gesture presentation controller
├── change_pptslide_with_gest_with_zoom.py # Extended controller with slide zoom capabilities
├── requirements.txt                       # Project dependencies
├── .gitignore                             # Ignored files for Python & OS
└── README.md                              # Project documentation
```

---

## 🛠️ Built With

- [OpenCV](https://opencv.org/) - Real-time computer vision and camera feed processing
- [Google MediaPipe](https://developers.google.com/mediapipe) - High-fidelity hand and finger tracking
- [pywin32](https://github.com/mhammond/pywin32) - Windows COM client integration with Microsoft Office PowerPoint

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).
