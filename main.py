import aspose.slides as slides
import aspose.pydrawing as drawing
import os
import cv2
from cvzone.HandTrackingModule import HandDetector
import glob
import shutil
import tkinter as tk
from tkinter import Tk, Button, Label, filedialog, messagebox, Listbox, Scrollbar
import subprocess
import sys
import time

def show_loading_screen():
    global loading_screen
    loading_screen = Tk()
    loading_screen.title("Loading")
    loading_screen.geometry("300x100")
    Label(loading_screen, text="Loading. Please Wait.").pack(pady=20)
    loading_screen.update()

def hide_loading_screen():
    global loading_screen
    loading_screen.destroy()

def convert_pptx_to_png(pptx_file):
    show_loading_screen()
    # Load presentation
    pres = slides.Presentation(pptx_file)

    # Create presentation folder if it doesn't exist
    output_folder = "presentation"
    os.makedirs(output_folder, exist_ok=True)

    # Get slide size from the presentation
    slide_size = pres.slide_size.size

    # Loop through slides
    for index in range(pres.slides.length):
        # Get reference of slide
        slide = pres.slides[index]
    
        # Modifying size of the slide
        scale_x = 1280 / slide_size.width
        scale_y = 720 / slide_size.height
        thumbnail = slide.get_thumbnail(scale_x, scale_y)

        # Save as PNG in the presentation folder, starting from 1.png
        thumbnail.save(os.path.join(output_folder, "{i}.png".format(i=index + 1)), drawing.imaging.ImageFormat.png)
    hide_loading_screen()
    root.destroy()
    start_presentation()

def start_presentation():
    global root, folderPath
    # Print OpenCV build information for debugging
    print(cv2.getBuildInformation())
    print(cv2.__version__)
    print(cv2.RETR_EXTERNAL)

    #variables
    width, height=1280, 720

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
        cap.release()
        cv2.destroyAllWindows()

        # Ensure the presentation folder is empty
        clear_presentation_folder()

        # Redirect to home page
        subprocess.run([sys.executable, os.path.join(os.getcwd(), "main.py")])

def upload_pptx():
    global selected_pptx_file
    pptx_file = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if pptx_file:
        pptx_label.config(text=os.path.basename(pptx_file))
        selected_pptx_file = pptx_file

def go_back():
    root.destroy()
    subprocess.run([sys.executable, os.path.join(os.getcwd(), "main.py")])

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
    clear_presentation_folder()

def clear_presentation_folder():
    folderPath = "presentation"
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

def choose_saved_folder():
    def on_double_click(event):
        selected_folder = listbox.get(listbox.curselection())
        folder_path = os.path.join(saved_folder_path, selected_folder)
        if folder_path and os.path.commonpath([folder_path, saved_folder_path]) == saved_folder_path:
            # Clear the presentation folder
            clear_presentation_folder()
            
            # Copy PNG files from the chosen folder to the presentation folder
            for file_name in os.listdir(folder_path):
                if file_name.endswith(".png"):
                    shutil.copy(os.path.join(folder_path, file_name), "presentation")
            
            # Set folderPath to the selected folder
            global folderPath
            folderPath = "presentation"
            
            # Run the presentation
            folder_selection_window.destroy()
            root.destroy()
            start_presentation()
            
            # Ensure the presentation folder is empty
            clear_presentation_folder()
        else:
            messagebox.showerror("Error", "Please select a folder inside the 'saved' directory.")

    def on_close():
        folder_selection_window.destroy()
        upload_button.config(state=tk.NORMAL)
        choose_folder_button.config(state=tk.NORMAL)

    saved_folder_path = os.path.join(os.getcwd(), "saved")
    folders = [f for f in os.listdir(saved_folder_path) if os.path.isdir(os.path.join(saved_folder_path, f))]

    # Create a new window for folder selection
    folder_selection_window = tk.Toplevel(root)
    folder_selection_window.title("Select a Folder")
    folder_selection_window.geometry("400x300")
    folder_selection_window.configure(bg="#f0f0f0")
    folder_selection_window.protocol("WM_DELETE_WINDOW", on_close)

    # Disable main window buttons
    upload_button.config(state=tk.DISABLED)
    choose_folder_button.config(state=tk.DISABLED)

    # Add a title label
    title_label = tk.Label(folder_selection_window, text="Select a Folder", font=("Helvetica", 18), bg="#f0f0f0")
    title_label.pack(pady=20)

    # Add a listbox with a scrollbar
    scrollbar = Scrollbar(folder_selection_window)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = Listbox(folder_selection_window, font=("Helvetica", 14), yscrollcommand=scrollbar.set)
    for folder in folders:
        listbox.insert(tk.END, folder)
    listbox.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)

    listbox.bind('<Double-1>', on_double_click)
    scrollbar.config(command=listbox.yview)

if __name__ == "__main__":
    global root, selected_pptx_file, loading_screen, folderPath, folder_name
    selected_pptx_file = None
    folderPath = sys.argv[1] if len(sys.argv) > 1 else "presentation"

    # Ensure the presentation folder is empty at the start
    clear_presentation_folder()

    if len(sys.argv) > 1 and sys.argv[1] == "upload":
        root = Tk()
        root.title("Upload PPTX")
        root.geometry("1280x720")  # Set the resolution size

        # Set a background color
        root.configure(bg="#f0f0f0")

        # Add a back button
        back_button = Button(root, text="← Back", command=go_back, font=("Helvetica", 14), bg="#2196F3", fg="white", padx=20, pady=10)
        back_button.pack(pady=20, anchor='ne')

        # Add a title label
        title_label = Label(root, text="Upload PPTX", font=("Helvetica", 24), bg="#f0f0f0")
        title_label.pack(pady=40)

        upload_button = Button(root, text="Upload PPTX", command=upload_pptx, font=("Helvetica", 16), bg="#4CAF50", fg="white", padx=20, pady=10)
        upload_button.pack(pady=20)

        pptx_label = Label(root, text="No file selected", font=("Helvetica", 16), bg="#f0f0f0")
        pptx_label.pack(pady=20)

        submit_button = Button(root, text="Submit", command=lambda: convert_pptx_to_png(selected_pptx_file), font=("Helvetica", 16), bg="#2196F3", fg="white", padx=20, pady=10)
        submit_button.pack(pady=20)

        root.mainloop()
    else:
        root = tk.Tk()
        root.title("Home Page")
        root.geometry("1280x720")  # Set the resolution size

        # Set a background color
        root.configure(bg="#f0f0f0")

        # Add a title label
        title_label = tk.Label(root, text="Hand Gesture Presentation", font=("Helvetica", 24), bg="#f0f0f0")
        title_label.pack(pady=40)

        # Add buttons with improved styling
        upload_button = tk.Button(root, text="Upload PPT", command=lambda: [root.destroy(), subprocess.run([sys.executable, os.path.join(os.getcwd(), "main.py"), "upload"])], font=("Helvetica", 16), bg="#4CAF50", fg="white", padx=20, pady=10)
        upload_button.pack(pady=20)

        choose_folder_button = tk.Button(root, text="Choose Saved Folder", command=choose_saved_folder, font=("Helvetica", 16), bg="#2196F3", fg="white", padx=20, pady=10)
        choose_folder_button.pack(pady=20)

        root.mainloop()

        # Print OpenCV build information for debugging
        print(cv2.getBuildInformation())
        print(cv2.__version__)
        print(cv2.RETR_EXTERNAL)

        #variables
        width, height=1280, 720

        #Camera Setup
        cap = cv2.VideoCapture(0)
        cap.set(3, width)
        cap.set(4, height)

        # Get the list of presentation images
        pathImages = sorted(os.listdir(folderPath), key.len)
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

            # Ensure the presentation folder is empty
            clear_presentation_folder()

            # Redirect to home page
            subprocess.run([sys.executable, os.path.join(os.getcwd(), "main.py")])
