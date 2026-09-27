import cv2
import mediapipe as mp
import time
import threading
import queue
import psutil

def is_powerpoint_running():
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] and 'POWERPNT.EXE' in proc.info['name'].upper():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

try:
    import win32com.client
    win32com_available = True
except ImportError:
    win32com_available = False


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
    def __init__(self, action_queue):
        super().__init__()
        self.action_queue = action_queue
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

        # Engagement State variables
        self.engagement_active = False
        self.engagement_gesture_count = 0
        self.ENGAGEMENT_FRAMES_REQUIRED = 15  # ~1.5 seconds to wake up

        self.previous_landmarks = None
        self.palm_open_count = 0
        self.fist_count = 0
        self.index_right_count = 0
        self.two_fingers_left_count = 0
        self.slide_show_started = False

        # EMA Smoothing
        self.alpha = 0.5 # Smoothing factor
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

                if multi_hand_landmarks:
                    for hand_landmarks in multi_hand_landmarks:
                        self.mp_drawing.draw_landmarks(image, hand_landmarks, self.mp_hands.HAND_CONNECTIONS)

                if landmarks and len(landmarks) == 21:
                    # Apply smoothing filter
                    landmarks = self.apply_ema(landmarks)

                    fingers = self.fingers_open(landmarks)
                    palm_open = all(f == 1 for f in fingers)
                    fist = all(f == 0 for f in fingers)
                    index_right = False
                    two_fingers_left = False

                    if self.previous_landmarks:
                        index_right = self.is_index_right(landmarks, self.previous_landmarks)
                        two_fingers_left = self.are_two_fingers_left(landmarks, self.previous_landmarks)
                    self.previous_landmarks = landmarks

                    # Engagement Logic
                    if palm_open:
                        if not self.engagement_active:
                            self.engagement_gesture_count += 1
                            if self.engagement_gesture_count >= self.ENGAGEMENT_FRAMES_REQUIRED:
                                self.engagement_active = True
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
                        cv2.putText(image, "Status: ASLEEP (Hold Palm to Wake)", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
                    else:
                        cv2.putText(image, "Status: AWAKE", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)

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
                                print("Engagement State DEACTIVATED. Sleeping...")
                                if self.slide_show_started:
                                    try:
                                        if self.presentation and self.presentation.SlideShowWindow:
                                            self.presentation.SlideShowWindow.View.Exit()
                                            self.slide_show_started = False
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
                        cv2.putText(image, "Status: ASLEEP", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
                    else:
                        cv2.putText(image, "Status: AWAKE", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)

                cv2.imshow('PowerPoint Gesture Control', image)
                if cv2.waitKey(5) & 0xFF == 27:
                    import os
                    import signal
                    os.kill(os.getpid(), signal.SIGINT)

            except queue.Empty:
                # Still need to pump the OpenCV event loop so the window doesn't freeze
                try:
                    if cv2.getWindowProperty('PowerPoint Gesture Control', cv2.WND_PROP_VISIBLE) > 0:
                        if cv2.waitKey(5) & 0xFF == 27:
                            import os
                            import signal
                            os.kill(os.getpid(), signal.SIGINT)
                except cv2.error:
                    pass
                pass

    def stop(self):
        self.running = False
        cv2.destroyAllWindows()

def main():
    frame_queue = queue.Queue(maxsize=2)
    action_queue = queue.Queue(maxsize=2)

    camera_thread = CameraThread(frame_queue)
    ai_thread = AIThread(frame_queue, action_queue)
    action_thread = ActionThread(action_queue)

    camera_thread.start()
    ai_thread.start()
    action_thread.start()

    print("System started. Press Ctrl+C to exit.")
    try:
        while True:
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("Shutting down...")
        camera_thread.stop()
        ai_thread.stop()
        action_thread.stop()

        camera_thread.join()
        ai_thread.join()
        action_thread.join()
        print("Shutdown complete.")

if __name__ == "__main__":
    main()
