import cv2
import mediapipe as mp
import time
import threading
import queue
import psutil
import os
import sys
from PIL import Image, ImageDraw
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QObject, QPoint
from PyQt5.QtGui import QFont, QColor, QPainter, QBrush, QPen, QPainterPath, QRegion

pystray_available = False
if not (os.name == 'posix' and not os.environ.get('DISPLAY')):
    try:
        import pystray
        pystray_available = True
    except Exception as e:
        print("Pystray exception:", e)
else:
    print("Pystray not available (no DISPLAY in environment).")

try:
    import win32com.client
    win32com_available = True
except ImportError:
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
                if self.frame_queue.full():
                    try:
                        self.frame_queue.get_nowait()
                    except queue.Empty:
                        pass
                self.frame_queue.put(image)
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

    def run(self):
        while self.running:
            try:
                # Pause AI processing if PPT is not running
                if win32com_available and not is_powerpoint_running():
                    time.sleep(1)
                    # We still need to consume the frame queue so it doesn't block CameraThread indefinitely
                    try:
                        self.frame_queue.get_nowait()
                    except queue.Empty:
                        pass
                    continue

                image = self.frame_queue.get(timeout=0.1)
                image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
                results = self.hands.process(image_rgb)

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

                self.action_queue.put((landmarks, image, results.multi_hand_landmarks))
            except queue.Empty:
                pass
        self.hands.close()

    def stop(self):
        self.running = False

class ActionThread(threading.Thread):
    def __init__(self, action_queue, ui_queue):
        super().__init__()
        self.action_queue = action_queue
        self.ui_queue = ui_queue
        self.running = True
        self.mp_hands = mp.solutions.hands
        self.mp_drawing = mp.solutions.drawing_utils

        self.powerpoint = None
        self.presentation = None
        if win32com_available:
            try:
                self.powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")

            if self.powerpoint and hasattr(self.powerpoint, "Presentations") and self.powerpoint.Presentations.Count > 0:
                self.presentation = self.powerpoint.ActivePresentation
                print("Active presentation title:", self.presentation.Name)
            else:
                print("No active presentation")
        else:
            print("win32com not available. PowerPoint controls disabled.")

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

    def run(self):
        while self.running:
            try:
                data = self.action_queue.get(timeout=0.1)
                landmarks, image, multi_hand_landmarks = data

                if landmarks and len(landmarks) == 21:
                    landmarks = self.apply_ema(landmarks)

                    fingers = self.fingers_open(landmarks)
                    palm_open = all(f == 1 for f in fingers)
                    fist = all(f == 0 for f in fingers)
                    laser_pointer = self.is_laser_pointer(fingers)
                    spotlight = self.is_spotlight(fingers)
                    index_right = False
                    two_fingers_left = False

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

                    if palm_open:
                        if not self.engagement_active:
                            self.engagement_gesture_count += 1
                            if self.engagement_gesture_count >= self.ENGAGEMENT_FRAMES_REQUIRED:
                                self.engagement_active = True
                                self.ui_queue.put({"type": "status", "msg": "AWAKE"})
                                print("Engagement State ACTIVE. System is listening for gestures...")
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
                                if not self.slide_show_started:
                                    try:
                                        if self.presentation:
                                            self.presentation.SlideShowSettings.Run()
                                            self.slide_show_started = True
                                            self.ui_queue.put({"type": "action", "msg": "START SLIDESHOW"})
                                            print("Starting slide show")
                                    except Exception as e:
                                        print(f"Error starting slide show: {e}")
                                self.palm_open_count = 0
                        elif fist:
                            self.fist_count += 1
                            self.palm_open_count = 0
                            self.index_right_count = 0
                            self.two_fingers_left_count = 0
                            if self.fist_count >= self.GESTURE_CONSECUTIVE_FRAMES:
                                self.engagement_active = False
                                self.ui_queue.put({"type": "status", "msg": "ASLEEP"})
                                print("Engagement State DEACTIVATED. Sleeping...")
                                if self.slide_show_started:
                                    try:
                                        if self.presentation and self.presentation.SlideShowWindow:
                                            self.presentation.SlideShowWindow.View.Exit()
                                            self.slide_show_started = False
                                            self.ui_queue.put({"type": "action", "msg": "END SLIDESHOW"})
                                            print("Ending slide show")
                                    except Exception as e:
                                        print(f"Error ending slide show: {e}")
                                self.fist_count = 0
                        elif index_right:
                            self.index_right_count += 1
                            self.palm_open_count = 0
                            self.fist_count = 0
                            self.two_fingers_left_count = 0
                            if self.index_right_count >= self.GESTURE_CONSECUTIVE_FRAMES:
                                try:
                                    if self.presentation and self.presentation.SlideShowWindow:
                                        current_slide = self.presentation.SlideShowWindow.View.Slide
                                        if current_slide.SlideIndex < self.presentation.Slides.Count:
                                            self.presentation.SlideShowWindow.View.Next()
                                            self.ui_queue.put({"type": "action", "msg": "NEXT SLIDE ->"})
                                            print("Next slide")
                                            time.sleep(0.5)
                                except Exception as e:
                                    print(f"Error going to next slide: {e}")
                                self.index_right_count = 0
                        elif two_fingers_left:
                            self.two_fingers_left_count += 1
                            self.palm_open_count = 0
                            self.fist_count = 0
                            self.index_right_count = 0
                            if self.two_fingers_left_count >= self.GESTURE_CONSECUTIVE_FRAMES:
                                try:
                                    if self.presentation and self.presentation.SlideShowWindow:
                                        current_slide = self.presentation.SlideShowWindow.View.Slide
                                        if current_slide.SlideIndex > 1:
                                            self.presentation.SlideShowWindow.View.Previous()
                                            self.ui_queue.put({"type": "action", "msg": "<- PREV SLIDE"})
                                            print("Previous slide")
                                            time.sleep(0.5)
                                except Exception as e:
                                    print(f"Error going to previous slide: {e}")
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

            except queue.Empty:
                pass

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

            painter.setBrush(QColor(0, 0, 0, 180)) # 70% opacity black
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

            self.update() # Trigger repaint
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

def main():
    frame_queue = queue.Queue(maxsize=2)
    action_queue = queue.Queue(maxsize=2)
    ui_queue = queue.Queue()

    camera_thread = CameraThread(frame_queue)
    ai_thread = AIThread(frame_queue, action_queue)
    action_thread = ActionThread(action_queue, ui_queue)

    camera_thread.start()
    ai_thread.start()
    action_thread.start()

    def on_quit(icon, item):
        icon.stop()
        camera_thread.stop()
        ai_thread.stop()
        action_thread.stop()
        if 'app' in globals() or 'app' in locals():
            app.quit()
        # Alternatively, force exit if PyQt hangs
        os._exit(0)


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
        queue_timer.start(100) # Poll every 100ms

        if pystray_available:
            icon = pystray.Icon("InvisibleAssistant")
            icon.menu = pystray.Menu(pystray.MenuItem("Quit", on_quit))
            icon.icon = create_image()
            icon.title = "Invisible Presentation Assistant"

            # Pystray needs to run in a separate thread if PyQt is taking the main thread
            tray_thread = threading.Thread(target=icon.run)
            tray_thread.daemon = True
            tray_thread.start()
            print("System Tray starting. Right click icon to quit.")

        print("HUD starting.")
        try:
            app.exec_()
        except KeyboardInterrupt:
            pass
        finally:
            if pystray_available:
                icon.stop()
    else:
        print("Running headless mode. Press Ctrl+C to exit.")
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



    print("Shutting down...")
    camera_thread.stop()
    ai_thread.stop()
    action_thread.stop()

    camera_thread.join()
    ai_thread.join()
    action_thread.join()
    print("Shutdown complete.")

if __name__ == '__main__':
    main()
