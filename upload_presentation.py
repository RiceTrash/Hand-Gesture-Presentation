import aspose.slides as slides
import aspose.pydrawing as drawing
import os
from tkinter import Tk, Button, Label, filedialog
import subprocess
import sys

def convert_pptx_to_png(pptx_file):
    global root
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

        # Create a thumbnail with the specified size
        scale_x = 1280 / slide_size.width
        scale_y = 720 / slide_size.height
        thumbnail = slide.get_thumbnail(scale_x, scale_y)

        # Save as PNG in the presentation folder, starting from 1.png
        thumbnail.save(os.path.join(output_folder, "{i}.png".format(i=index + 1)), drawing.imaging.ImageFormat.png)
    root.destroy()
    subprocess.run([sys.executable, "main.py"])

def upload_pptx():
    pptx_file = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if pptx_file:
        pptx_label.config(text=os.path.basename(pptx_file))
        convert_pptx_to_png(pptx_file)

if __name__ == "__main__":
    global root
    # Create the UI
    root = Tk()
    root.title("Upload PPTX")
    root.geometry("400x200")

    upload_button = Button(root, text="Upload PPTX", command=upload_pptx)
    upload_button.pack(pady=10)

    pptx_label = Label(root, text="No file selected")
    pptx_label.pack(pady=10)

    submit_button = Button(root, text="Submit", command=lambda: convert_pptx_to_png(pptx_label.cget("text")))
    submit_button.pack(pady=10)

    root.mainloop()