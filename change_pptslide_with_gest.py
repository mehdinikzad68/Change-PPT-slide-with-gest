import argparse
import cv2
import logging
import mediapipe as mp
import os
import psutil
import queue
import sys
import threading
import time
from PIL import Image, ImageDraw
from PyQt5.QtCore import QObject, QPoint, Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QPainter, QPainterPath, QFont
from PyQt5.QtWidgets import QApplication, QLabel, QWidget

from powerpoint_runtime import PowerPointProcessCache, resolve_active_presentation, resolve_slideshow_view
from telemetry_utils import TelemetryWriter, build_telemetry_record, compute_fps

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

pystray_available = False
if not (os.name == 'posix' and not os.environ.get('DISPLAY')):
    try:
        import pystray
        pystray_available = True
    except Exception:
        logger.exception("Pystray import failed")
else:
    logger.info("Pystray not available (no DISPLAY in environment).")

try:
    import pythoncom
    import pywintypes
    import win32com.client

    COM_ERROR = pywintypes.com_error
    win32com_available = True
except ImportError:
    pythoncom = None
    pywintypes = None
    win32com = None
    COM_ERROR = RuntimeError
    win32com_available = False


def is_powerpoint_running():
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] and 'POWERPNT.EXE' in proc.info['name'].upper():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False


class DummyLandmark:
    def __init__(self, x, y, z):
        self.x = x
        self.y = y
        self.z = z


class CameraThread(threading.Thread):

    def __init__(self, frame_queue):
        super().__init__()
        self.frame_queue = frame_queue
        self.cap = cv2.VideoCapture(0)
        self.running = True

    def run(self):
        while self.running:
            success, image = self.cap.read()
            if success:
                t_capture = time.monotonic()
                if self.frame_queue.full():
                    try:
                        self.frame_queue.get_nowait()
                    except queue.Empty:
                        pass
                self.frame_queue.put((image, t_capture))
            else:
                time.sleep(0.01)
        self.cap.release()

    def stop(self):
        self.running = False


class AIThread(threading.Thread):
    def __init__(self, frame_queue, action_queue):
        super().__init__()
        self.frame_queue = frame_queue
        self.action_queue = action_queue
        self.mp_hands = mp.solutions.hands
        self.hands = self.mp_hands.Hands(static_image_mode=False, max_num_hands=1, min_detection_confidence=0.7)
        self.running = True
        self.previous_capture_time = None
        self.powerpoint_process_cache = (
            PowerPointProcessCache(checker=is_powerpoint_running, ttl_seconds=1.5)
            if win32com_available else None
        )

    def run(self):
        while self.running:
            try:
                # Pause AI processing if PPT is not running.
                # Process checks are cached to avoid scanning processes on every frame.
                if self.powerpoint_process_cache and not self.powerpoint_process_cache.is_running():
                    time.sleep(1)
                    # We still need to consume the frame queue so it doesn't block CameraThread indefinitely
                    try:
                        self.frame_queue.get_nowait()
                    except queue.Empty:
                        pass
                    continue

                image, t_capture = self.frame_queue.get(timeout=0.1)
                t_infer_start = time.monotonic()
                image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
                results = self.hands.process(image_rgb)
                t_infer_end = time.monotonic()

                landmarks = None
                if results.multi_hand_landmarks:
                    for hand_landmarks in results.multi_hand_landmarks:
                        landmarks = hand_landmarks.landmark
                        break

                if self.action_queue.full():
                    try:
                        self.action_queue.get_nowait()
                    except queue.Empty:
                        pass

                fps = compute_fps(t_capture, self.previous_capture_time)
                self.previous_capture_time = t_capture

                self.action_queue.put({
                    "landmarks": landmarks,
                    "image": image,
                    "multi_hand_landmarks": results.multi_hand_landmarks,
                    "telemetry": {
                        "t_capture": t_capture,
                        "t_infer_start": t_infer_start,
                        "t_infer_end": t_infer_end,
                        "fps": fps,
                    },
                })
            except queue.Empty:
                pass
        self.hands.close()

    def stop(self):
        self.running = False


class ActionThread(threading.Thread):
    def __init__(self, action_queue, ui_queue, telemetry_writer=None):
        super().__init__()
        self.action_queue = action_queue
        self.ui_queue = ui_queue
        self.telemetry_writer = telemetry_writer
        self.running = True
        self.mp_hands = mp.solutions.hands
        self.mp_drawing = mp.solutions.drawing_utils

        if win32com_available:
            logger.info("PowerPoint COM controls enabled.")
        else:
            logger.info("win32com not available. PowerPoint controls disabled.")

        self.MOVEMENT_THRESHOLD = 0.005
        self.GESTURE_CONSECUTIVE_FRAMES = 5

        self.engagement_active = False
        self.engagement_gesture_count = 0
        self.ENGAGEMENT_FRAMES_REQUIRED = 15

        self.previous_landmarks = None
        self.palm_open_count = 0
        self.fist_count = 0
        self.index_right_count = 0
        self.two_fingers_left_count = 0
        self.slide_show_started = False

        self.alpha = 0.5
        self.smoothed_landmarks = None

    def _get_powerpoint_application(self):
        if not win32com_available:
            return None
        return win32com.client.GetActiveObject("PowerPoint.Application")

    def _get_active_presentation(self):
        powerpoint = self._get_powerpoint_application()
        if powerpoint is None:
            return None
        return resolve_active_presentation(powerpoint)

    def _get_slideshow_view(self):
        powerpoint = self._get_powerpoint_application()
        if powerpoint is None:
            return None
        return resolve_slideshow_view(powerpoint)

    def apply_ema(self, current_landmarks):
        if self.smoothed_landmarks is None:
            self.smoothed_landmarks = [(lm.x, lm.y, lm.z) for lm in current_landmarks]
            return current_landmarks

        smoothed = []
        for i, lm in enumerate(current_landmarks):
            sx = self.alpha * lm.x + (1 - self.alpha) * self.smoothed_landmarks[i][0]
            sy = self.alpha * lm.y + (1 - self.alpha) * self.smoothed_landmarks[i][1]
            sz = self.alpha * lm.z + (1 - self.alpha) * self.smoothed_landmarks[i][2]
            self.smoothed_landmarks[i] = (sx, sy, sz)
            smoothed.append(DummyLandmark(sx, sy, sz))
        return smoothed

    def fingers_open(self, landmarks):
        fingers = []
        if landmarks[4].x > landmarks[3].x:
            fingers.append(1)
        else:
            fingers.append(0)
        for tip in [8, 12, 16, 20]:
            if landmarks[tip].y < landmarks[tip - 2].y:
                fingers.append(1)
            else:
                fingers.append(0)
        return fingers

    def is_index_right(self, landmarks, previous_landmarks):
        if previous_landmarks is None:
            return False
        current_index_tip = landmarks[8].x
        previous_index_tip = previous_landmarks[8].x
        if current_index_tip - previous_index_tip > self.MOVEMENT_THRESHOLD:
            return True
        return False

    def is_laser_pointer(self, fingers):
        # Only index finger is open
        return fingers[0] == 0 and fingers[1] == 1 and fingers[2] == 0 and fingers[3] == 0 and fingers[4] == 0

    def is_spotlight(self, fingers):
        # Thumb and index finger are open (L-shape)
        return fingers[0] == 1 and fingers[1] == 1 and fingers[2] == 0 and fingers[3] == 0 and fingers[4] == 0

    def are_two_fingers_left(self, landmarks, previous_landmarks):
        if previous_landmarks is None:
            return False
        current_index_tip = landmarks[8].x
        current_middle_tip = landmarks[12].x
        previous_index_tip = previous_landmarks[8].x
        previous_middle_tip = previous_landmarks[12].x
        if (current_index_tip - previous_index_tip < -self.MOVEMENT_THRESHOLD and
            current_middle_tip - previous_middle_tip < -self.MOVEMENT_THRESHOLD):
            return True
        return False

    def emit_telemetry(self, telemetry_context, gesture, command, t_decision, t_com_start=None, t_com_end=None, error=None):
        if not self.telemetry_writer:
            return

        record = build_telemetry_record(
            t_capture=telemetry_context.get("t_capture"),
            t_infer_start=telemetry_context.get("t_infer_start"),
            t_infer_end=telemetry_context.get("t_infer_end"),
            t_decision=t_decision,
            t_com_start=t_com_start,
            t_com_end=t_com_end,
            fps=telemetry_context.get("fps"),
            gesture=gesture,
            command=command,
        )
        record["engagement_active"] = self.engagement_active
        record["slide_show_started"] = self.slide_show_started
        if error is not None:
            record["error"] = error
        self.telemetry_writer.write(record)

    def run(self):
        com_initialized = False
        if win32com_available:
            pythoncom.CoInitialize()
            com_initialized = True
            logger.info("ActionThread initialized COM apartment.")

        try:
            while self.running:
                try:
                    data = self.action_queue.get(timeout=0.1)
                    landmarks = data["landmarks"]
                    image = data["image"]
                    multi_hand_landmarks = data["multi_hand_landmarks"]
                    telemetry_context = data["telemetry"]
                    t_decision = time.monotonic()
                    gesture = "no_hand"
                    command = None
                    t_com_start = None
                    t_com_end = None
                    telemetry_error = None

                    if landmarks and len(landmarks) == 21:
                        landmarks = self.apply_ema(landmarks)

                        fingers = self.fingers_open(landmarks)
                        palm_open = all(f == 1 for f in fingers)
                        fist = all(f == 0 for f in fingers)
                        laser_pointer = self.is_laser_pointer(fingers)
                        spotlight = self.is_spotlight(fingers)
                        index_right = False
                        two_fingers_left = False

                        if palm_open:
                            gesture = "palm_open"
                        elif fist:
                            gesture = "fist"
                        elif laser_pointer:
                            gesture = "laser_pointer"
                        elif spotlight:
                            gesture = "spotlight"
                        else:
                            gesture = "hand_detected"

                        # Always send AR updates for fluid movement, if active and engaged
                        if self.engagement_active:
                            if laser_pointer or spotlight:
                                # Index finger tip is landmark 8. We mirror the x coordinate so it feels like a mirror.
                                # Also sometimes coordinates go slightly out of 0-1 bounds.
                                x = 1.0 - landmarks[8].x
                                y = landmarks[8].y
                                self.ui_queue.put({
                                    "type": "ar_update",
                                    "pointer": laser_pointer,
                                    "spotlight": spotlight,
                                    "x": x,
                                    "y": y
                                })
                            else:
                                # Clear AR state
                                self.ui_queue.put({"type": "ar_update", "pointer": False, "spotlight": False})

                        if self.previous_landmarks:
                            index_right = self.is_index_right(landmarks, self.previous_landmarks)
                            two_fingers_left = self.are_two_fingers_left(landmarks, self.previous_landmarks)
                        self.previous_landmarks = landmarks

                        if index_right:
                            gesture = "index_right"
                        elif two_fingers_left:
                            gesture = "two_fingers_left"

                        if palm_open:
                            if not self.engagement_active:
                                self.engagement_gesture_count += 1
                                if self.engagement_gesture_count >= self.ENGAGEMENT_FRAMES_REQUIRED:
                                    self.engagement_active = True
                                    self.ui_queue.put({"type": "status", "msg": "AWAKE"})
                                    logger.info("Engagement State ACTIVE. System is listening for gestures...")
                                    self.engagement_gesture_count = 0
                                    time.sleep(0.5)
                            else:
                                self.palm_open_count += 1
                        else:
                            if not self.engagement_active:
                                self.engagement_gesture_count = 0

                        if not self.engagement_active:
                            self.palm_open_count = 0
                            self.fist_count = 0
                            self.index_right_count = 0
                            self.two_fingers_left_count = 0
                        else:
                            if palm_open:
                                self.fist_count = 0
                                self.index_right_count = 0
                                self.two_fingers_left_count = 0
                                if self.palm_open_count >= self.GESTURE_CONSECUTIVE_FRAMES:
                                    try:
                                        presentation = self._get_active_presentation()
                                        if presentation:
                                            t_com_start = time.monotonic()
                                            presentation.SlideShowSettings.Run()
                                            t_com_end = time.monotonic()
                                            self.slide_show_started = True
                                            command = "start_slideshow"
                                            self.ui_queue.put({"type": "action", "msg": "START SLIDESHOW"})
                                            logger.info("Starting slide show")
                                        else:
                                            logger.info("No active presentation available to start slideshow.")
                                    except COM_ERROR as e:
                                        telemetry_error = str(e)
                                        if t_com_start is not None:
                                            t_com_end = time.monotonic()
                                        logger.exception("PowerPoint COM error while starting slide show")
                                    self.palm_open_count = 0
                            elif fist:
                                self.fist_count += 1
                                self.palm_open_count = 0
                                self.index_right_count = 0
                                self.two_fingers_left_count = 0
                                if self.fist_count >= self.GESTURE_CONSECUTIVE_FRAMES:
                                    self.engagement_active = False
                                    self.ui_queue.put({"type": "status", "msg": "ASLEEP"})
                                    logger.info("Engagement State DEACTIVATED. Sleeping...")
                                    if self.slide_show_started:
                                        try:
                                            slideshow_view = self._get_slideshow_view()
                                            if slideshow_view:
                                                t_com_start = time.monotonic()
                                                slideshow_view.Exit()
                                                t_com_end = time.monotonic()
                                                self.slide_show_started = False
                                                command = "end_slideshow"
                                                self.ui_queue.put({"type": "action", "msg": "END SLIDESHOW"})
                                                logger.info("Ending slide show")
                                            else:
                                                logger.info("No slideshow window available to end slideshow.")
                                                self.slide_show_started = False
                                        except COM_ERROR as e:
                                            telemetry_error = str(e)
                                            if t_com_start is not None:
                                                t_com_end = time.monotonic()
                                            logger.exception("PowerPoint COM error while ending slide show")
                                    self.fist_count = 0
                            elif index_right:
                                self.index_right_count += 1
                                self.palm_open_count = 0
                                self.fist_count = 0
                                self.two_fingers_left_count = 0
                                if self.index_right_count >= self.GESTURE_CONSECUTIVE_FRAMES:
                                    try:
                                        slideshow_view = self._get_slideshow_view()
                                        if slideshow_view:
                                            current_slide = slideshow_view.Slide
                                            presentation = self._get_active_presentation()
                                            if presentation and current_slide.SlideIndex < presentation.Slides.Count:
                                                t_com_start = time.monotonic()
                                                slideshow_view.Next()
                                                t_com_end = time.monotonic()
                                                command = "next_slide"
                                                self.ui_queue.put({"type": "action", "msg": "NEXT SLIDE ->"})
                                                logger.info("Next slide")
                                                time.sleep(0.5)
                                        else:
                                            logger.info("No slideshow window available for next slide command.")
                                    except COM_ERROR as e:
                                        telemetry_error = str(e)
                                        if t_com_start is not None:
                                            t_com_end = time.monotonic()
                                        logger.exception("PowerPoint COM error while going to next slide")
                                    self.index_right_count = 0
                            elif two_fingers_left:
                                self.two_fingers_left_count += 1
                                self.palm_open_count = 0
                                self.fist_count = 0
                                self.index_right_count = 0
                                if self.two_fingers_left_count >= self.GESTURE_CONSECUTIVE_FRAMES:
                                    try:
                                        slideshow_view = self._get_slideshow_view()
                                        if slideshow_view:
                                            current_slide = slideshow_view.Slide
                                            if current_slide.SlideIndex > 1:
                                                t_com_start = time.monotonic()
                                                slideshow_view.Previous()
                                                t_com_end = time.monotonic()
                                                command = "previous_slide"
                                                self.ui_queue.put({"type": "action", "msg": "<- PREV SLIDE"})
                                                logger.info("Previous slide")
                                                time.sleep(0.5)
                                        else:
                                            logger.info("No slideshow window available for previous slide command.")
                                    except COM_ERROR as e:
                                        telemetry_error = str(e)
                                        if t_com_start is not None:
                                            t_com_end = time.monotonic()
                                        logger.exception("PowerPoint COM error while going to previous slide")
                                    self.two_fingers_left_count = 0
                            else:
                                self.palm_open_count = 0
                                self.fist_count = 0
                                self.index_right_count = 0
                                self.two_fingers_left_count = 0
                    else:
                        self.previous_landmarks = None
                        self.smoothed_landmarks = None
                        self.palm_open_count = 0
                        self.fist_count = 0
                        self.index_right_count = 0
                        self.two_fingers_left_count = 0
                        if not self.engagement_active:
                            self.engagement_gesture_count = 0

                    del image
                    del multi_hand_landmarks

                    self.emit_telemetry(
                        telemetry_context,
                        gesture,
                        command,
                        t_decision,
                        t_com_start=t_com_start,
                        t_com_end=t_com_end,
                        error=telemetry_error,
                    )
                except queue.Empty:
                    pass
        finally:
            if com_initialized:
                pythoncom.CoUninitialize()
                logger.info("ActionThread uninitialized COM apartment.")

    def stop(self):
        self.running = False


class UIUpdater(QObject):
    update_signal = pyqtSignal(dict)


class HUDWindow(QWidget):
    def __init__(self):
        super().__init__()

        # AR State
        self.pointer_active = False
        self.pointer_pos = QPoint(0, 0)
        self.spotlight_active = False

        self.initUI()

        self.fade_timer = QTimer(self)
        self.fade_timer.timeout.connect(self.hide_status)
        self.fade_timer.setSingleShot(True)

    def initUI(self):
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool | Qt.WindowTransparentForInput)
        self.setAttribute(Qt.WA_TranslucentBackground)

        # Always cover full screen for AR overlay
        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(screen)
        self.screen_width = screen.width()
        self.screen_height = screen.height()

        self.label = QLabel("", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setFont(QFont("Arial", 24, QFont.Bold))
        self.label.setStyleSheet("color: rgba(0, 255, 0, 255); background-color: rgba(0, 0, 0, 150); border-radius: 10px; padding: 10px;")

        # Position label in bottom right manually since we don't use a layout that fills the screen
        self.label.resize(300, 100)
        self.label.move(self.screen_width - 350, 50)
        self.label.hide()

        self.showFullScreen()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        if self.spotlight_active:
            # Draw semi-transparent dark background
            path = QPainterPath()
            path.addRect(0, 0, self.screen_width, self.screen_height)

            # Cut out a circle for the spotlight
            spot_path = QPainterPath()
            spot_path.addEllipse(self.pointer_pos, 150, 150)

            # Subtracted path
            final_path = path.subtracted(spot_path)

            painter.setBrush(QColor(0, 0, 0, 180))
            painter.setPen(Qt.NoPen)
            painter.drawPath(final_path)

        elif self.pointer_active:
            # Draw red laser dot
            painter.setBrush(QBrush(QColor(255, 0, 0, 255)))
            painter.setPen(Qt.NoPen)
            painter.drawEllipse(self.pointer_pos, 10, 10)

    def show_message(self, data):
        msg_type = data.get("type", "")

        if msg_type == "ar_update":
            self.pointer_active = data.get("pointer", False)
            self.spotlight_active = data.get("spotlight", False)

            if self.pointer_active or self.spotlight_active:
                # Map 0.0-1.0 coords to screen
                x = int(data.get("x", 0.5) * self.screen_width)
                y = int(data.get("y", 0.5) * self.screen_height)
                self.pointer_pos = QPoint(x, y)

            self.update()
            return

        msg = data.get("msg", "")
        if msg_type == "status":
            if msg == "AWAKE":
                self.label.setStyleSheet("color: rgba(0, 255, 0, 255); background-color: rgba(0, 0, 0, 150); border-radius: 10px; padding: 10px;")
            else:
                self.label.setStyleSheet("color: rgba(255, 0, 0, 255); background-color: rgba(0, 0, 0, 150); border-radius: 10px; padding: 10px;")
        else:
            self.label.setStyleSheet("color: rgba(0, 255, 255, 255); background-color: rgba(0, 0, 0, 150); border-radius: 10px; padding: 10px;")

        self.label.setText(msg)
        self.label.show()

        self.fade_timer.start(1500)

    def hide_status(self):
        self.label.hide()


def check_ui_queue(ui_queue, updater):
    try:
        while True:
            data = ui_queue.get_nowait()
            updater.update_signal.emit(data)
    except queue.Empty:
        pass


def create_image():
    image = Image.new('RGB', (64, 64), color=(50, 50, 50))
    dc = ImageDraw.Draw(image)
    dc.rectangle(
        [(16, 16), (48, 48)],
        fill=(0, 255, 0)
    )
    return image


def parse_args(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--telemetry", help="Path to a JSONL telemetry output file.")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    telemetry_writer = TelemetryWriter(args.telemetry) if args.telemetry else None
    if telemetry_writer:
        logger.info("Telemetry output enabled: %s", args.telemetry)

    frame_queue = queue.Queue(maxsize=2)
    action_queue = queue.Queue(maxsize=2)
    ui_queue = queue.Queue()

    camera_thread = CameraThread(frame_queue)
    ai_thread = AIThread(frame_queue, action_queue)
    action_thread = ActionThread(action_queue, ui_queue, telemetry_writer=telemetry_writer)

    camera_thread.start()
    ai_thread.start()
    action_thread.start()

    icon = None

    def on_quit(icon_arg, item):
        del item
        if icon_arg is not None:
            icon_arg.stop()
        camera_thread.stop()
        ai_thread.stop()
        action_thread.stop()
        if 'app' in globals() or 'app' in locals():
            app.quit()

    # PyQt setup
    # If no display, PyQt will crash, so we protect it
    has_display = True
    if os.name == 'posix' and not os.environ.get('DISPLAY'):
        has_display = False

    if has_display:
        app = QApplication(sys.argv)
        hud = HUDWindow()
        updater = UIUpdater()
        updater.update_signal.connect(hud.show_message)

        # Timer to poll the UI queue
        queue_timer = QTimer()
        queue_timer.timeout.connect(lambda: check_ui_queue(ui_queue, updater))
        queue_timer.start(100)

        if pystray_available:
            icon = pystray.Icon("InvisibleAssistant")
            icon.menu = pystray.Menu(pystray.MenuItem("Quit", on_quit))
            icon.icon = create_image()
            icon.title = "Invisible Presentation Assistant"

            # Pystray needs to run in a separate thread if PyQt is taking the main thread
            tray_thread = threading.Thread(target=icon.run)
            tray_thread.daemon = True
            tray_thread.start()
            logger.info("System Tray starting. Right click icon to quit.")

        logger.info("HUD starting.")
        try:
            app.exec_()
        except KeyboardInterrupt:
            pass
        finally:
            if icon is not None:
                icon.stop()
    else:
        logger.info("Running headless mode. Press Ctrl+C to exit.")
        try:
            while True:
                # Just drain the queue so it doesn't block
                try:
                    ui_queue.get_nowait()
                except queue.Empty:
                    pass
                time.sleep(0.1)
        except KeyboardInterrupt:
            pass

    logger.info("Shutting down...")
    camera_thread.stop()
    ai_thread.stop()
    action_thread.stop()

    camera_thread.join()
    ai_thread.join()
    action_thread.join()

    if telemetry_writer:
        telemetry_writer.close()

    logger.info("Shutdown complete.")


if __name__ == '__main__':
    main()
