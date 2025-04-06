from tkinter import Tk
from tkinter.filedialog import askopenfilename
import cv2
import win32com.client
from cvzone.HandTrackingModule import HandDetector
import os

# =============================== #
#         CONFIGURATIONS         #
# =============================== #

GESTURE_THRESHOLD = 300
CAM_WIDTH, CAM_HEIGHT = 900, 720
DETECTION_CONFIDENCE = 0.8
MAX_HANDS = 1
GESTURE_DELAY = 30

# =============================== #
#         MAIN FUNCTION           #
# =============================== #

def main():
    # Open PowerPoint
    try:
        # File picker for PowerPoint
        Tk().withdraw()  # Hide root window
        ppt_path = askopenfilename(filetypes=[("PowerPoint Files", "*.pptx")], title="Select your PPT")

        if not ppt_path:
            print("No presentation selected. Exiting.")
            return
        app = win32com.client.Dispatch("PowerPoint.Application")
        presentation = app.Presentations.Open(ppt_path)
        presentation.SlideShowSettings.Run()
        print(f"Presentation Started: {presentation.Name}")
    except Exception as e:
        print("Failed to open PowerPoint:", e)
        return

    # Initialize camera
    cap = cv2.VideoCapture(0)
    cap.set(3, CAM_WIDTH)
    cap.set(4, CAM_HEIGHT)

    # Hand detector
    detector = HandDetector(detectionCon=DETECTION_CONFIDENCE, maxHands=MAX_HANDS)

    # Variables
    buttonPressed = False
    delayCounter = 0
    slideCount = 20  # Number of slides (dummy limit)
    annotations = [[]]
    annotationNumber = -1
    annotationStart = False

    while True:
        success, img = cap.read()
        if not success:
            break

        hands, img = detector.findHands(img)

        if hands and not buttonPressed:
            hand = hands[0]
            cx, cy = hand["center"]
            fingers = detector.fingersUp(hand)

            # Gesture zone
            if cy <= GESTURE_THRESHOLD:
                if fingers == [1, 1, 1, 1, 1]:
                    print("Next Slide")
                    presentation.SlideShowWindow.View.Next()
                    buttonPressed = True
                    slideCount = max(slideCount - 1, 0)
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False

                elif fingers == [1, 0, 0, 0, 0]:
                    print("Previous Slide")
                    presentation.SlideShowWindow.View.Previous()
                    buttonPressed = True
                    slideCount += 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False

        # Delay logic
        if buttonPressed:
            delayCounter += 1
            if delayCounter > GESTURE_DELAY:
                delayCounter = 0
                buttonPressed = False

        # Drawing annotations (placeholder, future feature)
        for annotation in annotations:
            for j in range(1, len(annotation)):
                cv2.line(img, annotation[j - 1], annotation[j], (0, 0, 255), 12)

        # Display image
        cv2.imshow("Gesture Controlled PPT", img)

        # Exit on 'q'
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    presentation.SlideShowWindow.View.Exit()
    print("Presentation ended.")

# =============================== #
#            EXECUTE              #
# =============================== #

if __name__ == "__main__":
    main()
