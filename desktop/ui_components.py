from tkinter import Label, Button, Entry

def create_label(parent, text, font=("Arial", 12), **kwargs):
    label = Label(parent, text=text, font=font, **kwargs)
    label.pack(pady=5)
    return label

def create_button(parent, text, command, state="normal"):
    button = Button(parent, text=text, command=command, state=state)
    button.pack(pady=5)
    return button

def create_entry(parent, width=20):
    entry = Entry(parent, width=width)
    entry.pack(pady=5)
    return entry
