import glob
import os
import logging
import platform
import subprocess
import threading
from tkinter import END, Entry, OptionMenu, Toplevel, messagebox, StringVar, Text, Label, Button, Frame, Canvas, Scrollbar, filedialog, Listbox, MULTIPLE
from collections import defaultdict
from PIL import Image, ImageTk
from document_generator import create_document, add_image_to_doc
from help_feedback import open_help, open_feedback
from preferences import load_preferences, save_preferences

# Set up logging
logging.basicConfig(filename='app.log', level=logging.ERROR)

class DemoWindow:
    def __init__(self, parent):
        self.top = Toplevel(parent)
        self.top.title("Demo Instructions")
        self.top.geometry("400x300")
        
        self.instructions_label = Label(self.top, text="", wraplength=350, font=("Arial", 12))
        self.instructions_label.pack(pady=20, padx=10)

        self.close_button = Button(self.top, text="Close", command=self.close_demo_window)
        self.close_button.pack(pady=10)

    def update_instructions(self, instruction):
        self.instructions_label.config(text=instruction)

    def close_demo_window(self):
        # Close all opened .docx files
        self.close_all_word_documents()
        self.top.destroy()

    def close_all_word_documents(self):
        if platform.system() == "Windows":
            # Use taskkill to close Microsoft Word
            try:
                subprocess.call(["taskkill", "/F", "/IM", "WINWORD.EXE"])
            except Exception as e:
                print(f"Error closing Word documents: {e}")

class MultiFolderDialog:
    def __init__(self, parent, on_folders_selected):
        self.top = Toplevel(parent)
        self.top.title("Select Multiple Folders")
        self.top.geometry("400x300")

        # Set the dialog to be transient and always on top of the parent window
        self.top.transient(parent)
        self.top.grab_set()  # Ensures that all events are directed to this window until it's closed
        self.top.focus_set()  # Focus on the dialog window
        self.top.lift()  # Raise the dialog above other windows

        self.selected_folders = []
        self.folder_listbox = Listbox(self.top, selectmode=MULTIPLE)
        self.folder_listbox.pack(fill='both', expand=True, padx=10, pady=10)

        Button(self.top, text="Add Folder", command=self.add_folder).pack(pady=5)
        Button(self.top, text="Remove Selected", command=self.remove_selected).pack(pady=5)
        Button(self.top, text="OK", command=self.ok).pack(pady=5)

    def add_folder(self):
        folder_path = filedialog.askdirectory(mustexist=True)
        if folder_path and folder_path not in self.selected_folders:
            self.selected_folders.append(folder_path)
            self.folder_listbox.insert(END, folder_path)

    def remove_selected(self):
        selected_indices = self.folder_listbox.curselection()
        for index in reversed(selected_indices):
            del self.selected_folders[index]
            self.folder_listbox.delete(index)

    def ok(self):
        self.top.destroy()


class ImageToWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image to Word Document Generator")
        self.root.geometry("600x400")
        self.image_folders = []  # List to hold selected folders
        self.notes = {}
        self.layout_choices = {}
        self.image_types = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')
        self.max_image_width = 5  # Default maximum image width

        self.preview_window = None  # To keep track of the preview window
        self.create_widgets()
        load_preferences()  # Load user preferences on start

    def create_widgets(self):
        # Create a central frame for the content
        content_frame = Frame(self.root)
        content_frame.pack(expand=True)

        Label(content_frame, text="Select folders containing images:", font=("Arial", 12)).pack(pady=10, anchor="center")

        self.button_select_folders = Button(content_frame, text="Select Folders", command=self.select_folders, font=("Arial", 12), bg="lightblue")
        self.button_select_folders.pack(pady=5, anchor="center")

        self.button_generate = Button(content_frame, text="Generate Documents", command=self.generate_documents, state="disabled", font=("Arial", 12), bg="lightgreen")
        self.button_generate.pack(pady=10, anchor="center")

        Button(content_frame, text="Help", command=lambda: open_help(self.root), font=("Arial", 12), bg="lightgrey").pack(pady=5, anchor="center")
        Button(content_frame, text="Demo", command=self.start_demo, font=("Arial", 12), bg="red").pack(pady=5, anchor="center")
        Button(content_frame, text="Feedback", command=lambda: open_feedback(self.root), font=("Arial", 12), bg="lightgrey").pack(pady=5, anchor="center")

        self.status_label = Label(content_frame, text="", font=("Arial", 10))
        self.status_label.pack(pady=5, anchor="center")

    def update_status(self, message):
        self.status_label.config(text=message)

    def select_folders(self):
        # Get the current position of the root window
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        
        # Open the folder selection dialog next to the main window
        dialog = MultiFolderDialog(self.root,self.on_demo_folders_selected)
        
        # Adjust the position of the MultiFolderDialog window to be beside the root window
        dialog.top.geometry(f"+{root_x + root_width + 10}+{root_y}")
        
        self.root.wait_window(dialog.top)  # Wait for the dialog to close
        self.image_folders = dialog.selected_folders

        if self.image_folders:
            self.button_generate.config(state="normal")  # Enable document generation button
            self.update_status(f"Selected {len(self.image_folders)} folder(s).")
            self.show_preview()

    def show_preview(self):
        # Get the current position of the root window
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        
        # Create a new preview window and position it next to the root window
        self.preview_window = Toplevel(self.root)
        self.preview_window.title("Image Preview")
        self.preview_window.geometry(f"700x500+{root_x + root_width + 10}+{root_y}")
        self.preview_window.resizable(True, True)

        self.canvas = Canvas(self.preview_window)
        scrollbar = Scrollbar(self.preview_window, command=self.canvas.yview)
        self.scrollable_frame = Frame(self.canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        Label(self.scrollable_frame, text="Image Preview", font=("Arial", 16, "bold")).pack(pady=10, anchor="center")

        # Create an entry for max width input
        Label(self.scrollable_frame, text="Max Image Width:", font=("Arial", 12)).pack(pady=5, anchor="w")
        self.max_width_entry = Entry(self.scrollable_frame, width=10)  # Create the entry widget
        self.max_width_entry.pack(pady=5, anchor="w")
        self.max_width_entry.insert(0, str(self.max_image_width))  # Set default value

        # Load images from all selected folders
        self.load_images_with_folders()

        Button(self.preview_window, text="Save/Apply", command=self.on_save_apply, font=("Arial", 12), bg="lightgreen").pack(pady=10, anchor="center")



    def load_images_with_folders(self):
        # Clear previous notes and layout choices before loading new images
        self.notes.clear()
        self.layout_choices.clear()

        for folder in self.image_folders:  # Loop through each selected folder
            folder_name = os.path.basename(folder)
            Label(self.scrollable_frame, text=f"Folder: {folder_name}", font=("Arial", 14, "bold")).pack(pady=10, anchor="w")

            images_dict = defaultdict(list)
            try:
                for filename in os.listdir(folder):
                    if filename.lower().endswith(self.image_types):
                        image_path = os.path.join(folder, filename)
                        title = filename.split('_')[0]
                        images_dict[title].append(image_path)

                for title, image_paths in images_dict.items():
                    self.add_image_controls_for_folder(title, image_paths)

            except Exception as e:
                logging.error(f"Error loading images from folder {folder}: {e}")
                messagebox.showerror("Error", f"Could not load images from {folder}. Please check the folder.")

    def add_image_controls_for_folder(self, title, image_paths):
        title_frame = Frame(self.scrollable_frame)
        title_frame.pack(pady=10, anchor="w")

        Label(title_frame, text=title, font=("Arial", 12, "bold")).pack(pady=5, anchor="w")

        layout_var = StringVar(value="Single Column")
        option_menu = OptionMenu(title_frame, layout_var, "Single Column", "Two Columns", command=lambda _: self.refresh_preview(title))
        option_menu.pack(side="top", padx=5, pady=5, anchor="w")
        self.layout_choices[title] = layout_var

        note_entry = Text(title_frame, height=3, width=40, wrap='word', font=("Arial", 12))
        note_entry.pack(pady=5, anchor="w")
        self.notes[title] = note_entry

        self.preview_images_for_folder(title_frame, title, image_paths)

    def preview_images_for_folder(self, frame, title, image_paths):
        layout = self.layout_choices[title].get()
        columns = 2 if layout == "Two Columns" else 1

        preview_frame = Frame(frame)
        preview_frame.pack(pady=5, anchor="w")

        for i, image_path in enumerate(image_paths):
            try:
                img = Image.open(image_path)
                img.thumbnail((100, 100))  # Thumbnail for uniform image size
                img_tk = ImageTk.PhotoImage(img)

                label = Label(preview_frame, image=img_tk)
                label.image = img_tk  # Keep a reference
                label.grid(row=i // columns, column=i % columns, padx=5, pady=5)

            except Exception as e:
                logging.error(f"Error loading image {image_path}: {e}")
                messagebox.showerror("Error", f"Could not load image {image_path}. Please check the file.")

    def refresh_preview(self, title):
        for widget in self.scrollable_frame.winfo_children():
            if isinstance(widget, Frame) and widget.winfo_children()[0].cget("text") == title:
                for subwidget in widget.winfo_children():
                    if isinstance(subwidget, Frame):
                        subwidget.destroy()

                # Reload images based on the title and selected folders
                image_paths = [
                    os.path.join(folder, filename) 
                    for folder in self.image_folders 
                    for filename in os.listdir(folder) 
                    if filename.lower().endswith(self.image_types) and filename.startswith(title)
                ]
                self.preview_images_for_folder(widget, title, image_paths)
                break
            
    def on_save_apply(self):
        try:
            self.max_image_width = float(self.max_width_entry.get())
        except ValueError:
            messagebox.showwarning("Invalid Input", "Please enter a valid number for max image width.")
            return

        # Save the notes as actual string values instead of widget references
        for title, note_widget in self.notes.items():
            if isinstance(note_widget, Text):
                self.notes[title] = note_widget.get("1.0", END).strip()

        save_preferences()
        messagebox.showinfo("Success", "Settings have been saved!")
        self.preview_window.destroy()



    def generate_documents(self):
        self.update_status("Generating documents...")     
        threading.Thread(target=self.create_documents).start()


    def create_documents(self):
        for folder in self.image_folders:
            try:
                # Compile images and related data specific to this folder
                images_dict = self.compile_images(folder)

                # Create an output file path based on the folder name
                folder_name = os.path.basename(folder)
                output_file = f"{folder_name}.docx"  # Set the output file name and path
                output_file_path = os.path.join(folder, output_file)  # Full path to the output file

                # Call create_document with the correct order of arguments
                create_document(output_file_path, images_dict, self.max_image_width)

                self.update_status(f"Document created: {output_file_path}")

                # Open the document
                os.startfile(output_file_path)

            except Exception as e:
                logging.error(f"Error creating document for folder {folder}: {e}")
                messagebox.showerror("Error", f"Could not create document for {folder}. Please check the folder.")

    def compile_images(self, folder):
        images_dict = defaultdict(lambda: {'image_paths': [], 'note': '', 'layout': 'Single Column'})

        # Go through each image in the folder and match it to the corresponding title and note
        for filename in os.listdir(folder):
            if filename.lower().endswith(self.image_types):
                image_path = os.path.join(folder, filename)
                title = filename.split('_')[0]  # Extract title from image filename (before the first '_')

                # If the title exists in layout_choices, proceed
                if title in self.layout_choices:
                    images_dict[title]['image_paths'].append(image_path)
                    images_dict[title]['layout'] = self.layout_choices[title].get()

                    # Only add the note if it exists in the self.notes dictionary
                    if title in self.notes:
                        images_dict[title]['note'] = self.notes[title]

        return images_dict

    def start_demo(self):
        # Cleanup demo files first
        demo_folders = [
            "D:\\kerjaan\\iseng\\wordGenerator\\demofolder1",
            "D:\\kerjaan\\iseng\\wordGenerator\\demofolder2"
        ]
        self.cleanup_demo_files(demo_folders)

        # Create demo instructions window
        self.demo_window = DemoWindow(self.root)

        # Get the current position of the main window
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_height = self.root.winfo_height()

        # Position the demo window at the bottom of the main window
        self.demo_window.top.geometry(f"+{root_x}+{root_y + root_height + 10}")  # 10 pixels below the main window

        # Define demo steps with instructions
        steps = [
            (self.highlight_select_folders, "Highlighting the Select Folders button."),
            (self.demo_select_folders, "Simulating folder selection of images."),
            (self.highlight_generate_button, "Highlighting the Generate Documents button."),
            (self.demo_generate_documents, "Simulating the document generation process."),
            (self.close_demo_windows, "Closing all demo windows.")
        ]
        
        # Start running the demo steps
        self.run_demo_steps(steps)

    def run_demo_steps(self, steps, current_step=0):
        if current_step < len(steps):
            if self.demo_window.top.winfo_exists():  # Check if the demo window still exists
                if self.demo_window.instructions_label.winfo_exists():  # Check if the instructions label exists
                    highlight_function, instruction = steps[current_step]
                    highlight_function()  # Execute the highlight function
                    self.demo_window.update_instructions(instruction)  # Update instructions
                    self.update_status(f"Step {current_step + 1}: {instruction}")

                    # Adjusted timing for smoother flow
                    delay = 3000 if current_step != 1 else 2000  # Shorter delay for selecting folders
                    # Proceed to the next step after the adjusted delay
                    self.root.after(delay, lambda: self.run_demo_steps(steps, current_step + 1))
                else:
                    self.update_status("Demo label was closed or destroyed prematurely.")
            else:
                self.update_status("Demo window was closed or destroyed prematurely.")
        else:
            if self.demo_window.top.winfo_exists():  # Ensure the window exists for the final step
                self.update_status("Demo completed!")
                self.demo_window.update_instructions("Demo completed! You can close this window manually when you're done.")
                self.demo_window.close_button.config(state="normal")  # Enable the close button
                # Bring the demo window to the top of all windows and give it focus
                self.demo_window.top.lift()  # Brings the demo window to the front
                self.demo_window.top.focus_force()  # Ensures the demo window gets focus
                self.demo_window.top.attributes('-topmost', 1)  # Make sure it's always on top

    def close_demo_windows(self):
        if self.preview_window:
            self.preview_window.destroy()  # Close the preview window if open
        # if self.demo_window:
        #     self.demo_window.top.destroy()  # Close the demo window
        self.update_status("Reset all demo windows.")

    def highlight_select_folders(self):
        self.button_generate.config(bg="lightgrey")  # Reset Generate button highlight
        self.button_generate.config(state="normal")  # Simulate enabling the button
        # Highlight the Select Folders button
        self.blink_button(self.button_select_folders, "yellow", 2000)  # Blink for 2 seconds

    def highlight_generate_button(self):
        self.blink_button(self.button_generate, "yellow", 2000)  # Blink for 2 seconds


    def blink_button(self, button, color, duration):
        # This function will blink the button between its original and the new color
        original_color = button.cget("bg")  # Get the original background color
        blink_count = duration // 500  # Number of times to blink (every 500ms)
        
        def toggle_color(count):
            if count > 0:
                # Toggle between the original color and the new color
                current_color = button.cget("bg")
                new_color = color if current_color == original_color else original_color
                button.config(bg=new_color)
                self.root.after(500, lambda: toggle_color(count - 1))  # Repeat after 500ms
            else:
                button.config(bg=original_color)  # Reset to original color after blinking ends

        toggle_color(blink_count)  # Start the blinking process


    # Simulate folder selection in demo:
    def demo_select_folders(self):
        demo_folders = [
            "D:\\kerjaan\\iseng\\wordGenerator\\demofolder1",
            "D:\\kerjaan\\iseng\\wordGenerator\\demofolder2"
        ]

        # Simulate opening the multi-folder dialog
        self.dialog = MultiFolderDialog(self.root, self.on_demo_folders_selected)  # Save the reference
        
        # Position the dialog beside the demo window
        demo_window_x = self.demo_window.top.winfo_x()
        demo_window_y = self.demo_window.top.winfo_y()
        dialog_x = demo_window_x + self.demo_window.top.winfo_width() + 10  # Position beside the demo window
        dialog_y = demo_window_y  # Align vertically with the demo window

        self.dialog.top.geometry(f"+{dialog_x}+{dialog_y}")  # Set position of MultiFolderDialog
        
        for folder in demo_folders:
            self.dialog.selected_folders.append(folder)
            self.dialog.folder_listbox.insert(END, folder)

        # Call the selection callback with the simulated folders after a delay
        self.root.after(2000, lambda: self.on_demo_folders_selected(self.dialog.selected_folders))  # Use lambda to defer the call


    def on_demo_folders_selected(self, selected_folders):
        self.image_folders = selected_folders
        self.update_status(f"Selected {len(self.image_folders)} demo folder(s).")
        self.button_generate.config(state="normal")
        # Add a delay before closing the dialog
        self.root.after(2000, self.close_demo_dialog)

        # Add a delay before showing the preview window
        self.show_preview()

    def close_demo_dialog(self):
        if hasattr(self, 'dialog'):  # Check if dialog exists
            self.dialog.ok()  # Close the dialog
            del self.dialog  # Optionally delete the reference

    def demo_generate_documents(self):
        self.update_status("Simulating document generation...")
        threading.Thread(target=self.create_documents).start()

    def cleanup_demo_files(self, folders):
        # Loop through each folder provided
        for folder in folders:
            # Define the pattern to match .docx files
            pattern = os.path.join(folder, "*.docx")
            
            # Get a list of all .docx files in the current folder
            generated_files = glob.glob(pattern)
            
            # Check if any files are found
            if generated_files:
                for file_path in generated_files:
                    try:
                        os.remove(file_path)  # Remove the file
                        self.update_status(f"Deleted {file_path}.")
                    except Exception as e:
                        self.update_status(f"Error deleting {file_path}: {e}")
            else:
                self.update_status(f"No demo files found in {folder} to delete.")
