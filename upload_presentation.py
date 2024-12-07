import aspose.slides as slides
import aspose.pydrawing as drawing
import os
from tkinter import Tk, Button, Label, filedialog, messagebox
import subprocess
import sys

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
    global root
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
    root.destroy()
    hide_loading_screen()
    subprocess.run([sys.executable, "main.py"])

def upload_pptx():
    global selected_pptx_file
    pptx_file = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if pptx_file:
        pptx_label.config(text=os.path.basename(pptx_file))
        selected_pptx_file = pptx_file
        convert_ppt_to_images(pptx_file, "presentation")
        messagebox.showinfo("Success", "PPT uploaded and converted to images successfully!")

def convert_ppt_to_images(ppt_path, output_folder):
    # Dummy function to represent PPT to images conversion
    # You need to implement this function
    pass

if __name__ == "__main__":
    global root, selected_pptx_file, loading_screen
    selected_pptx_file = None
    # Create the UI
    root = Tk()
    root.title("Upload PPTX")
    root.geometry("400x200")

    upload_button = Button(root, text="Upload PPTX", command=upload_pptx)
    upload_button.pack(pady=10)

    pptx_label = Label(root, text="No file selected")
    pptx_label.pack(pady=10)

    submit_button = Button(root, text="Submit", command=lambda: convert_pptx_to_png(selected_pptx_file))
    submit_button.pack(pady=10)

    root.mainloop()