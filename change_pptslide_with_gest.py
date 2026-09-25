import cv2
import mediapipe as mp
import win32com.client
import time

# Initialize MediaPipe hands and drawing utilities
mp_hands = mp.solutions.hands
mp_drawing = mp.solutions.drawing_utils
hands = mp_hands.Hands(static_image_mode=False, max_num_hands=1, min_detection_confidence=0.7)

# Initialize OpenCV video capture
cap = cv2.VideoCapture(0)

# Connect to PowerPoint
try:
    powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
except:
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    # Optionally open a presentation
    # powerpoint.Presentations.Open('path_to_presentation.pptx')

presentation = None
if powerpoint.Presentations.Count > 0:
    presentation = powerpoint.ActivePresentation
    print("Active presentation title:", presentation.Name)
else:
    print("No active presentation")

# Function to check if fingers are open
def fingers_open(landmarks):
    fingers = []
    # Thumb
    if landmarks[4].x > landmarks[3].x:
        fingers.append(1)
    else:
        fingers.append(0)
    # Index, Middle, Ring, Pinky
    for tip in [8, 12, 16, 20]:
        if landmarks[tip].y < landmarks[tip - 2].y:
            fingers.append(1)
        else:
            fingers.append(0)
    return fingers

# Function to check if palm is fully open
def is_palm_open(fingers):
    return all(finger == 1 for finger in fingers)

# Function to check if fist is closed
def is_fist(fingers):
    return all(finger == 0 for finger in fingers)

# Function to check if index finger moved to the right
def is_index_right(landmarks, previous_landmarks):
    if previous_landmarks is None:
        return False
    current_index_tip = landmarks[8].x
    previous_index_tip = previous_landmarks[8].x
    print(f"Index finger moved: {current_index_tip - previous_index_tip}")
    if current_index_tip - previous_index_tip > MOVEMENT_THRESHOLD:
        return True
    return False

# Function to check if two fingers moved to the left
def are_two_fingers_left(landmarks, previous_landmarks):
    if previous_landmarks is None:
        return False
    current_index_tip = landmarks[8].x
    current_middle_tip = landmarks[12].x
    previous_index_tip = previous_landmarks[8].x
    previous_middle_tip = previous_landmarks[12].x
    print(f"Index finger moved: {current_index_tip - previous_index_tip}")
    print(f"Middle finger moved: {current_middle_tip - previous_middle_tip}")
    if (current_index_tip - previous_index_tip < -MOVEMENT_THRESHOLD and
        current_middle_tip - previous_middle_tip < -MOVEMENT_THRESHOLD):
        return True
    return False

# Define thresholds
MOVEMENT_THRESHOLD = 0.005
GESTURE_CONSECUTIVE_FRAMES = 5

# Initialize variables
previous_landmarks = None
palm_open_count = 0
fist_count = 0
index_right_count = 0
two_fingers_left_count = 0
slide_show_started = False

while cap.isOpened():
    success, image = cap.read()
    if not success:
        print("Ignoring empty frame.")
        continue

    # Convert the BGR image to RGB and process it
    image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    results = hands.process(image)
    image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)

    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            mp_drawing.draw_landmarks(image, hand_landmarks, mp_hands.HAND_CONNECTIONS)
            landmarks = hand_landmarks.landmark
            if len(landmarks) == 21:
                fingers = fingers_open(landmarks)
                palm_open = is_palm_open(fingers)
                fist = is_fist(fingers)
                index_right = False
                two_fingers_left = False
                if previous_landmarks:
                    index_right = is_index_right(landmarks, previous_landmarks)
                    two_fingers_left = are_two_fingers_left(landmarks, previous_landmarks)
                previous_landmarks = landmarks

                # Gesture detection logic
                if palm_open:
                    palm_open_count += 1
                    fist_count = 0
                    index_right_count = 0
                    two_fingers_left_count = 0
                    if palm_open_count >= GESTURE_CONSECUTIVE_FRAMES:
                        if not slide_show_started:
                            try:
                                if presentation:
                                    presentation.SlideShowSettings.Run()
                                    slide_show_started = True
                                    print("Starting slide show")
                            except Exception as e:
                                print(f"Error starting slide show: {e}")
                        palm_open_count = 0
                elif fist:
                    fist_count += 1
                    palm_open_count = 0
                    index_right_count = 0
                    two_fingers_left_count = 0
                    if fist_count >= GESTURE_CONSECUTIVE_FRAMES:
                        if slide_show_started:
                            try:
                                if presentation and presentation.SlideShowWindow:
                                    presentation.SlideShowWindow.View.Exit()
                                    slide_show_started = False
                                    print("Ending slide show")
                            except Exception as e:
                                print(f"Error ending slide show: {e}")
                        fist_count = 0
                elif index_right:
                    index_right_count += 1
                    palm_open_count = 0
                    fist_count = 0
                    two_fingers_left_count = 0
                    if index_right_count >= GESTURE_CONSECUTIVE_FRAMES:
                        try:
                            if presentation and presentation.SlideShowWindow:
                                current_slide = presentation.SlideShowWindow.View.Slide
                                if current_slide.SlideIndex < presentation.Slides.Count:
                                    presentation.SlideShowWindow.View.Next()
                                    print("Next slide")
                                    time.sleep(0.5)  # Added delay
                                else:
                                    print("Already on the last slide")
                        except Exception as e:
                            print(f"Error going to next slide: {e}")
                        index_right_count = 0
                elif two_fingers_left:
                    two_fingers_left_count += 1
                    palm_open_count = 0
                    fist_count = 0
                    index_right_count = 0
                    if two_fingers_left_count >= GESTURE_CONSECUTIVE_FRAMES:
                        try:
                            if presentation and presentation.SlideShowWindow:
                                current_slide = presentation.SlideShowWindow.View.Slide
                                if current_slide.SlideIndex > 1:
                                    presentation.SlideShowWindow.View.Previous()
                                    print("Previous slide")
                                    time.sleep(0.5)  # Added delay
                                else:
                                    print("Already on the first slide")
                        except Exception as e:
                            print(f"Error going to previous slide: {e}")
                        two_fingers_left_count = 0
                else:
                    palm_open_count = 0
                    fist_count = 0
                    index_right_count = 0
                    two_fingers_left_count = 0
    else:
        previous_landmarks = None
        palm_open_count = 0
        fist_count = 0
        index_right_count = 0
        two_fingers_left_count = 0

    cv2.imshow('PowerPoint Gesture Control', image)

    if cv2.waitKey(5) & 0xFF == 27:
        break

cap.release()
cv2.destroyAllWindows()

import cv2
print(cv2.__file__)