import tkinter as tk
from tkinter import filedialog, messagebox
import shutil
import os
import subprocess
import sys

def upload_ppt():
    root.destroy()
    subprocess.run([sys.executable, "upload_presentation.py"])

def choose_saved_folder():
    saved_folder_path = os.path.join(os.getcwd(), "saved")
    folder_path = filedialog.askdirectory(initialdir=saved_folder_path, title="Select a folder inside 'saved'")
    if folder_path and os.path.commonpath([folder_path, saved_folder_path]) == saved_folder_path:
        # Clear the presentation folder
        presentation_folder = os.path.join(os.getcwd(), "presentation")
        if os.path.exists(presentation_folder):
            shutil.rmtree(presentation_folder)
        os.makedirs(presentation_folder)
        
        # Copy PNG files from the chosen folder to the presentation folder
        for file_name in os.listdir(folder_path):
            if file_name.endswith(".png"):
                shutil.copy(os.path.join(folder_path, file_name), presentation_folder)
        
        # Run main.py with the presentation folder
        python_executable = sys.executable
        root.destroy()
        subprocess.run([python_executable, "main.py", presentation_folder], env=os.environ.copy())
        
        # Ensure the presentation folder is empty
        for file_name in os.listdir(presentation_folder):
            file_path = os.path.join(presentation_folder, file_name)
            if os.path.isfile(file_path):
                os.remove(file_path)
    else:
        messagebox.showerror("Error", "Please select a folder inside the 'saved' directory.")

def convert_ppt_to_images(ppt_path, output_folder):
    # Dummy function to represent PPT to images conversion
    # You need to implement this function
    pass

root = tk.Tk()
root.title("Home Page")

upload_button = tk.Button(root, text="Upload PPT", command=upload_ppt)
upload_button.pack(pady=20)

choose_folder_button = tk.Button(root, text="Choose Saved Folder", command=choose_saved_folder)
choose_folder_button.pack(pady=20)

root.mainloop()