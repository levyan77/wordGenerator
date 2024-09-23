import os
import logging
import threading
from tkinter import END, OptionMenu, Toplevel, messagebox, StringVar, Text, Label, Button, Entry, Frame, Canvas, Scrollbar, filedialog
from collections import defaultdict
from PIL import Image, ImageTk
from document_generator import create_document, add_image_to_doc
from help_feedback import open_help, open_feedback
from preferences import load_preferences, save_preferences

# Set up logging
logging.basicConfig(filename='app.log', level=logging.ERROR)

class ImageToWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image to Word Document Generator")
        self.root.geometry("600x400")
        self.image_folder = ""
        self.notes = {}
        self.layout_choices = {}
        self.image_types = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')
        self.max_image_width = 5  # Default maximum image width

        self.create_widgets()
        load_preferences()  # Load user preferences on start

    def create_widgets(self):
        Label(self.root, text="Select a folder containing images:", font=("Arial", 12)).pack(pady=10)
        Button(self.root, text="Select Folder", command=self.select_folder, font=("Arial", 12), bg="lightblue").pack(pady=5)
        self.button_generate = Button(self.root, text="Generate Document", command=self.generate_document, state="disabled", font=("Arial", 12), bg="lightgreen")
        self.button_generate.pack(pady=10)
        Button(self.root, text="Help", command=lambda: open_help(self.root), font=("Arial", 12), bg="lightgrey").pack(pady=5)
        Button(self.root, text="Feedback", command=lambda: open_feedback(self.root), font=("Arial", 12), bg="lightgrey").pack(pady=5)
        self.status_label = Label(self.root, text="", font=("Arial", 10))
        self.status_label.pack(pady=5)

    def update_status(self, message):
        self.status_label.config(text=message)

    def select_folder(self):
        threading.Thread(target=self._select_folder).start()

    def _select_folder(self):
        try:
            self.image_folder = filedialog.askdirectory()
            if self.image_folder:
                self.button_generate.config(state="normal")  # Enable document generation button
                self.update_status(f"Selected folder: {os.path.basename(self.image_folder)}")
                self.show_preview()
        except Exception as e:
            print(f"Error selecting folder: {e}")

    def show_preview(self):
        self.preview_window = Toplevel(self.root)
        self.preview_window.title("Image Preview")
        self.preview_window.geometry("700x500")
        self.preview_window.resizable(True, True)

        self.canvas = Canvas(self.preview_window)
        scrollbar = Scrollbar(self.preview_window, command=self.canvas.yview)
        self.scrollable_frame = Frame(self.canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        Label(self.scrollable_frame, text="Image Preview", font=("Arial", 16, "bold")).pack(pady=10)

        self.load_images(self.scrollable_frame)

        width_frame = Frame(self.scrollable_frame)
        width_frame.pack(pady=15)

        Label(width_frame, text="Max Image Width (inches):", font=("Arial", 12)).pack(side="left")
        self.max_width_entry = Entry(width_frame, width=5, font=("Arial", 12))
        self.max_width_entry.insert(0, str(self.max_image_width))
        self.max_width_entry.pack(side="left", padx=5)

        Button(self.preview_window, text="Save/Apply", command=self.on_save_apply, font=("Arial", 12), bg="lightgreen").pack(pady=10)

    def load_images(self, frame):
        images_dict = defaultdict(list)
        for filename in os.listdir(self.image_folder):
            if filename.lower().endswith(self.image_types):
                image_path = os.path.join(self.image_folder, filename)
                title = filename.split('_')[0]
                images_dict[title].append(image_path)

        for title, image_paths in images_dict.items():
            self.add_image_controls(frame, title, image_paths)

    def add_image_controls(self, frame, title, image_paths):
        title_frame = Frame(frame)
        title_frame.pack(pady=10)

        Label(title_frame, text=title, font=("Arial", 14, "bold")).pack(pady=5)

        layout_var = StringVar(value="Single Column")
        option_menu = OptionMenu(title_frame, layout_var, "Single Column", "Two Columns", command=lambda _: self.refresh_preview(title))
        option_menu.pack(side="left", padx=5)
        self.layout_choices[title] = layout_var

        note_entry = Text(title_frame, height=3, width=40, wrap='word', font=("Arial", 12))
        note_entry.pack(pady=5)
        self.notes[title] = note_entry

        self.preview_images(title_frame, title, image_paths)

    def preview_images(self, frame, title, image_paths):
        layout = self.layout_choices[title].get()
        columns = 2 if layout == "Two Columns" else 1

        preview_frame = Frame(frame)
        preview_frame.pack(pady=5)

        for i, image_path in enumerate(image_paths):
            img = Image.open(image_path)
            img.thumbnail((100, 100))
            img_tk = ImageTk.PhotoImage(img)

            label = Label(preview_frame, image=img_tk)
            label.image = img_tk
            label.grid(row=i // columns, column=i % columns, padx=5, pady=5)

        frame.pack(pady=10)

    def refresh_preview(self, title):
        for widget in self.scrollable_frame.winfo_children():
            if isinstance(widget, Frame) and widget.winfo_children()[0].cget("text") == title:
                for subwidget in widget.winfo_children():
                    if isinstance(subwidget, Frame):
                        subwidget.destroy()
                image_paths = [os.path.join(self.image_folder, filename) for filename in os.listdir(self.image_folder) if filename.startswith(title)]
                self.preview_images(widget, title, image_paths)
                break

    def on_save_apply(self):
        try:
            self.max_image_width = float(self.max_width_entry.get())
        except ValueError:
            messagebox.showwarning("Invalid Input", "Please enter a valid number for max image width.")
            return

        # Ensure we correctly save the note as a string while keeping the widget reference intact
        for title, note_widget in self.notes.items():
            if isinstance(note_widget, Text):
                self.notes[title] = note_widget.get("1.0", END).strip()

        save_preferences()
        messagebox.showinfo("Success", "Settings have been saved!")
        self.preview_window.destroy()



    def generate_document(self):
        self.update_status("Generating document...")
        threading.Thread(target=self.create_document).start()

    def create_document(self):
        if not self.image_folder:
            messagebox.showwarning("No Folder Selected", "Please select a folder first.")
            return

        folder_name = os.path.basename(self.image_folder)
        output_file = os.path.join(self.image_folder, f"{folder_name}.docx")

        if os.path.exists(output_file) and not messagebox.askyesno("File Exists", f"The file '{output_file}' already exists. Overwrite?"):
            return

        images_dict = self.compile_images()
        create_document(output_file, images_dict, self.max_image_width)

        messagebox.showinfo("Success", f"Document '{output_file}' created successfully.")
        self.update_status(f"Document '{output_file}' created successfully.")
        self.notes.clear()
        self.layout_choices.clear()

        os.startfile(os.path.dirname(output_file))

    def compile_images(self):
        images_dict = defaultdict(lambda: {'image_paths': [], 'note': '', 'layout': 'Single Column'})

        for title in self.layout_choices.keys():
            images_dict[title]['layout'] = self.layout_choices[title].get()
            images_dict[title]['note'] = self.notes[title]  # This should be a string now

            for filename in os.listdir(self.image_folder):
                if filename.lower().endswith(self.image_types):
                    image_path = os.path.join(self.image_folder, filename)
                    if filename.split('_')[0] == title:
                        images_dict[title]['image_paths'].append(image_path)
        return images_dict
