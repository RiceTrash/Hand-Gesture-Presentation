import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar
import shutil
import os
import subprocess
import sys

def upload_ppt():
    root.destroy()
    subprocess.run([sys.executable, "upload_presentation.py"])

def choose_saved_folder():
    def on_double_click(event):
        selected_folder = listbox.get(listbox.curselection())
        folder_path = os.path.join(saved_folder_path, selected_folder)
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
            folder_selection_window.destroy()
            root.destroy()
            subprocess.run([python_executable, "main.py", presentation_folder], env=os.environ.copy())
            
            # Ensure the presentation folder is empty
            for file_name in os.listdir(presentation_folder):
                file_path = os.path.join(presentation_folder, file_name)
                if os.path.isfile(file_path):
                    os.remove(file_path)
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

def convert_ppt_to_images(ppt_path, output_folder):
    # Dummy function to represent PPT to images conversion
    # You need to implement this function
    pass

root = tk.Tk()
root.title("Home Page")
root.geometry("1280x720")  # Set the resolution size

# Set a background color
root.configure(bg="#f0f0f0")

# Add a title label
title_label = tk.Label(root, text="Hand Gesture Presentation", font=("Helvetica", 24), bg="#f0f0f0")
title_label.pack(pady=40)

# Add buttons with improved styling
upload_button = tk.Button(root, text="Upload PPT", command=upload_ppt, font=("Helvetica", 16), bg="#4CAF50", fg="white", padx=20, pady=10)
upload_button.pack(pady=20)

choose_folder_button = tk.Button(root, text="Choose Saved Folder", command=choose_saved_folder, font=("Helvetica", 16), bg="#2196F3", fg="white", padx=20, pady=10)
choose_folder_button.pack(pady=20)

root.mainloop()