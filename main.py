import os
import cv2
from cvzone.HandTrackingModule import HandDetector

#variables
width, height=1280, 720
folderPath = "presentation"


#Camera Setup
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# Get the list of presentation images
pathImages = sorted(os.listdir(folderPath), key=len)
# print(pathImages)

# Variables
imgNumber = 0
hs, ws = int(120*1), int(213*1)
gestureThreshold = 300
buttonPressed = False
buttonCounter = 0
buttonDelay = 30

# Hand Detector
detector = HandDetector(detectionCon=0.8,maxHands=1)


while True:
    # Import images
    success, img = cap.read()
    img = cv2.flip(img, 1)
    pathFullImage = os.path.join(folderPath,pathImages[imgNumber])
    imgCurrent = cv2.imread(pathFullImage)

    hands, img = detector.findHands(img)
    cv2.line(img, (0, gestureThreshold), (width, gestureThreshold), (0, 255, 0), 10)

    if hands and buttonPressed is False:
        hand = hands[0]
        fingers = detector.fingersUp(hand)
        cx, cy = hand['center']

        if cy <=gestureThreshold: # if hand is the height of the face

            # Gesture 1 - Left
            if fingers == [1,0,0,0,0]:
                print("Left")
                if imgNumber > 0:
                    buttonPressed = True
                    imgNumber -= 1

            # Gesture 1 - Right
            if fingers == [0, 0, 0, 0, 1]:
                print("Right")
                if imgNumber < len(pathImages)-1:
                    buttonPressed = True
                    imgNumber += 1

    # Button Pressed iterations
    if buttonPressed:
        buttonCounter += 1
        if buttonCounter > buttonDelay:
            buttonCounter = 0
            buttonPressed = False


    # Adding webcam image on the slides
    imgSmall = cv2.resize(img, (ws, hs))
    h, w, _ = imgCurrent.shape
    imgCurrent[0:hs, w-ws:w] = imgSmall


    cv2.imshow("Image", img)
    cv2.imshow("Slides", imgCurrent)
    key = cv2.waitKey(1)
    if key == ord('q'):
        break
