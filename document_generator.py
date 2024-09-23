import os
import sys
from docx import Document
from docx.shared import Inches
from PIL import Image
import logging

def get_safe_max_image_width(max_image_width, is_two_column):
    page_width = 8.5  # inches
    if is_two_column:
        max_allowed_width = (page_width - 2) / 2  # Account for left and right margins
        return min(max_image_width, max_allowed_width) - 0.5  # Slightly reduce to prevent cropping
    return min(max_image_width, page_width - 2)  # For one column, subtract total margins

def create_document(output_file, images_dict, max_image_width, max_image_height=4):
    doc = Document()
    folder_name = os.path.basename(output_file).replace('.docx', '')
    doc.add_heading(folder_name, level=1)

    for title, content in images_dict.items():
        doc.add_heading(title, level=2)
        max_image_width_for_layout = get_safe_max_image_width(max_image_width, content['layout'] == "Two Columns")

        if content['layout'] == "Two Columns":
            add_images_to_doc_two_columns(doc, content, max_image_width_for_layout, max_image_height)
        else:
            add_images_to_doc_one_column(doc, content, max_image_width_for_layout, max_image_height)

        if isinstance(content['note'], str) and content['note']:
            doc.add_paragraph(content['note'])
        doc.add_paragraph()

    try:
        save_document(doc, output_file)
    except PermissionError:
        print(f"Permission denied: The file '{output_file}' is already open. Please close it and try again.")
        sys.exit(1)  # Exit the program with an error status

def save_document(doc, output_file):
    """Attempts to save the document, raising a PermissionError if the file is open."""
    doc.save(output_file)

def add_images_to_doc_two_columns(doc, content, max_image_width, max_image_height):
    table = doc.add_table(rows=0, cols=2)
    for i in range(0, len(content['image_paths']), 2):
        row_cells = table.add_row().cells
        for j in range(2):
            if i + j < len(content['image_paths']):
                add_image_to_cell(row_cells[j], content['image_paths'][i + j], max_image_width, max_image_height)

def add_images_to_doc_one_column(doc, content, max_image_width, max_image_height):
    for image_path in content['image_paths']:
        add_image_to_doc(doc, image_path, max_image_width, max_image_height)

def add_image_to_cell(cell, image_path, max_image_width, max_image_height):
    try:
        with Image.open(image_path) as img:
            width, height = img.size
            ratio = min(max_image_width / width, max_image_height / height)

            new_width = width * ratio
            new_height = height * ratio

            # Ensure the new dimensions are positive
            if new_width > 0 and new_height > 0:
                # Use the new width and height calculated from the aspect ratio
                cell.paragraphs[0].add_run().add_picture(image_path, width=Inches(new_width), height=Inches(new_height))
    except Exception as e:
        logging.error(f"Error adding image to cell: {image_path} - {e}")

def add_image_to_doc(doc, image_path, max_image_width, max_image_height):
    try:
        with Image.open(image_path) as img:
            width, height = img.size
            ratio = min(max_image_width / width, max_image_height / height)

            new_width = width * ratio
            new_height = height * ratio

            # Ensure the new dimensions are positive
            if new_width > 0 and new_height > 0:
                # Use the new width and height calculated from the aspect ratio
                doc.add_picture(image_path, width=Inches(new_width), height=Inches(new_height))
    except Exception as e:
        logging.error(f"Error adding image to document: {image_path} - {e}")