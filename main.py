import os
import cv2
from cvzone.HandTrackingModule import HandDetector
import glob
import shutil
import tkinter as tk
from tkinter import simpledialog
import sys
import time

def cleanup_presentation_folder(destination_folder=None):
    if destination_folder:
        # Create the destination folder if it doesn't exist
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
        # Move all PNG files to the destination folder
        png_files = glob.glob(os.path.join(folderPath, "*.png"))
        for file in png_files:
            shutil.move(file, destination_folder)
    # Ensure the presentation folder is empty
    for file in os.listdir(folderPath):
        file_path = os.path.join(folderPath, file)
        if os.path.isfile(file_path):
            os.remove(file_path)

def get_folder_name():
    def on_confirm():
        global folder_name
        folder_name = entry.get()
        root.destroy()

    root = tk.Tk()
    root.title("Save Presentation")
    root.geometry("400x200")
    root.configure(bg="#f0f0f0")

    label = tk.Label(root, text="Enter the folder name:", font=("Helvetica", 14), bg="#f0f0f0")
    label.pack(pady=20, padx=50)

    entry = tk.Entry(root, font=("Helvetica", 14))
    entry.pack(pady=10)

    confirm_button = tk.Button(root, text="Confirm", command=on_confirm, font=("Helvetica", 14), bg="#4CAF50", fg="white", padx=0, pady=10)
    confirm_button.pack(pady=20)

    root.mainloop()
    return folder_name

# Print OpenCV build information for debugging
print(cv2.getBuildInformation())
print(cv2.__version__)
print(cv2.RETR_EXTERNAL)

#variables
width, height=1280, 720
folderPath = sys.argv[1] if len(sys.argv) > 1 else "presentation"

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
gestureThreshold = 700
buttonPressed = False
buttonCounter = 0
buttonDelay = 30
closeCounter = 0
closeThreshold = 3  # 

# Hand Detector
detector = HandDetector(detectionCon=0.8,maxHands=1)

try:
    while True:
        # Import images
        success, img = cap.read()
        img = cv2.flip(img, 1)
        pathFullImage = os.path.join(folderPath,pathImages[imgNumber])
        imgCurrent = cv2.imread(pathFullImage)

        hands, img = detector.findHands(img)
        cv2.line(img, (0, gestureThreshold), (width, gestureThreshold), (0, 255, 0), 10)

        if hands:
            hand = hands[0]
            fingers = detector.fingersUp(hand)
            cx, cy = hand['center']

            if cy <= gestureThreshold:  # if hand is the height of the face

                # Gesture 1 - Left
                if fingers == [1,0,0,0,0] and not buttonPressed:
                    print("Left")
                    if imgNumber > 0:
                        buttonPressed = True
                        imgNumber -= 1
                        

                # Gesture 1 - Right
                if fingers == [0, 0, 0, 0, 1] and not buttonPressed:
                    print("Right")
                    if imgNumber < len(pathImages)-1:
                        buttonPressed = True
                        imgNumber += 1
                    

                # Gesture to close the application
                if fingers == [1, 1, 1, 1, 1]:
                    closeCounter += 1
                    if closeCounter > closeThreshold * 30:  # Assuming 30 FPS
                        break
                else:
                    closeCounter = 0

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

        # cv2.imshow("Image", img)
        cv2.imshow("Slides", imgCurrent)
        key = cv2.waitKey(1)
        if key == ord('q'):
            break
finally:
    if folderPath == "presentation":
        folder_name = get_folder_name()
        if folder_name:
            destination_folder = os.path.join("saved", folder_name)
            cleanup_presentation_folder(destination_folder)
        else:
            cleanup_presentation_folder()
    else:
        cleanup_presentation_folder()
    cap.release()
    cv2.destroyAllWindows()
