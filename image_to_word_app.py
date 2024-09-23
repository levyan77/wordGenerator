import os
from tkinter import Tk, Label, Button, filedialog, messagebox
from docx import Document
from docx.shared import Inches
from PIL import Image as PILImage
from collections import defaultdict

class ImageToWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image to Word Document Generator")
        
        self.label = Label(root, text="Select a folder containing images:")
        self.label.pack(pady=10)

        self.button_select = Button(root, text="Select Folder", command=self.select_folder)
        self.button_select.pack(pady=5)

        self.button_generate = Button(root, text="Generate Document", command=self.generate_document, state="disabled")
        self.button_generate.pack(pady=5)

        self.image_folder = ""

    def select_folder(self):
        self.image_folder = filedialog.askdirectory()
        if self.image_folder:
            self.button_generate.config(state="normal")
            messagebox.showinfo("Folder Selected", f"Selected folder: {self.image_folder}")

    def generate_document(self):
        if not self.image_folder:
            messagebox.showwarning("No Folder Selected", "Please select a folder first.")
            return

        output_file = os.path.join(self.image_folder, "TC01.docx")
        doc = Document()
        images_dict = defaultdict(list)

        # Document dimensions in inches (A4 size)
        max_width = 6.5  # Adjust according to your needs
        max_height = 9.0  # Adjust according to your needs

                # Group images by common prefix (before the underscore)
        for filename in os.listdir(self.image_folder):
            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                # Extract the common prefix before the underscore
                title = filename.split('_')[0]  # Group by the part before the first underscore
                images_dict[title].append(os.path.join(self.image_folder, filename))



                # Open image to get dimensions
                with PILImage.open(self.image_folder) as img:
                    width, height = img.size
                    aspect_ratio = width / height

                    # Calculate new dimensions while maintaining aspect ratio
                    if aspect_ratio > 1:  # Landscape
                        new_width = min(max_width, width / 100)  # Convert pixels to inches
                        new_height = new_width / aspect_ratio
                    else:  # Portrait
                        new_height = min(max_height, height / 100)  # Convert pixels to inches
                        new_width = new_height * aspect_ratio

        doc.save(output_file)
        messagebox.showinfo("Success", f"Document '{output_file}' created successfully.")

if __name__ == "__main__":
    root = Tk()
    app = ImageToWordApp(root)
    root.mainloop()