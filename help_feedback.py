from tkinter import Toplevel, messagebox, Label, Button, Text, Entry, Frame
import threading
import smtplib
import logging

def open_help(root):
    help_window = Toplevel(root)
    help_window.title("Help")
    help_window.geometry("600x400")
    
    content_frame = Frame(help_window)
    content_frame.pack(fill='both', expand=True, padx=10, pady=10)

    Label(content_frame, text="How to use the Image to Word Document Generator", font=("Arial", 14, "bold")).pack(pady=(10, 0))

    instructions = (
        "1. Click 'Select Folder' to choose a folder with your images.\n"
        "2. Preview your images and adjust the layout (single or two columns).\n"
        "3. Optionally add notes for each image group.\n"
        "4. Set the max image width if needed.\n"
        "5. Click 'Generate Document' to create a Word document with your images.\n"
        "6. After generation, the folder containing the document will open automatically."
    )

    instructions_label = Label(content_frame, text=instructions, font=("Arial", 12), justify="left")
    instructions_label.pack(fill='both', expand=True, padx=5, pady=(5, 10))

    close_button = Button(content_frame, text="Close", command=help_window.destroy, font=("Arial", 12))
    close_button.pack(pady=10)

def open_feedback(root):
    feedback_window = Toplevel(root)
    feedback_window.title("Feedback")
    feedback_window.geometry("600x600")

    input_frame = Frame(feedback_window)
    input_frame.pack(fill='both', expand=True, padx=10, pady=10)

    Label(input_frame, text="Feedback Form", font=("Arial", 14, "bold")).pack(pady=(10, 0))

    name_label = Label(input_frame, text="Your Name:", font=("Arial", 12))
    name_label.pack(pady=5)
    name_entry = Entry(input_frame, font=("Arial", 12))
    name_entry.pack(pady=5)

    email_label = Label(input_frame, text="Your Email:", font=("Arial", 12))
    email_label.pack(pady=5)
    email_entry = Entry(input_frame, font=("Arial", 12))
    email_entry.pack(pady=5)

    feedback_label = Label(input_frame, text="Your Feedback:", font=("Arial", 12))
    feedback_label.pack(pady=5)
    feedback_text = Text(input_frame, height=10, font=("Arial", 12))
    feedback_text.pack(pady=5, fill='both', expand=True)

    button_frame = Frame(input_frame)
    button_frame.pack(pady=(5, 10))

    submit_button = Button(button_frame, text="Submit", command=lambda: threading.Thread(target=send_feedback, args=(name_entry.get(), email_entry.get(), feedback_text.get("1.0", "end"))).start(), font=("Arial", 12), bg="lightgreen")
    submit_button.pack(side="left", padx=5)

    close_button = Button(button_frame, text="Close", command=feedback_window.destroy, font=("Arial", 12))
    close_button.pack(side="right", padx=5)

def send_feedback(name, email, feedback):
    if not name or not email or not feedback.strip():
        messagebox.showwarning("Incomplete Form", "Please fill out all fields before submitting.")
        return

    try:
        sender_email = "your_email@gmail.com"
        receiver_email = "your_receiver_email@gmail.com"
        password = "your_password"  # Use environment variables for security
        subject = "User Feedback"
        body = f"Name: {name}\nEmail: {email}\n\nFeedback:\n{feedback}"

        message = f"Subject: {subject}\n\n{body}"

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message)

        messagebox.showinfo("Feedback Sent", "Thank you for your feedback!")
    except Exception as e:
        logging.error(f"Failed to send feedback: {e}")
        messagebox.showerror("Error", "Could not send feedback. Please try again later.")