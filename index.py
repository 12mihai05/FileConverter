import os
import shutil
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from pdf2docx import Converter

def convert_files_in_folder(folder_path, target_format, output_folder):
    # Check if the folder exists
    if not os.path.exists(folder_path):
        print(f"Error: Folder {folder_path} not found.")
        return

    # Get the list of items (files and directories) in the folder
    items = os.listdir(folder_path)

    # Iterate over each item in the folder
    for item in items:
        item_path = os.path.join(folder_path, item)
        # If the item is a file, convert it to the target format
        if os.path.isfile(item_path):
            convert_file(item_path, target_format, output_folder)
        # If the item is a directory, create a corresponding folder in the output directory
        elif os.path.isdir(item_path):
            subdir_output_folder = os.path.join(output_folder, item)
            os.makedirs(subdir_output_folder, exist_ok=True)
            # Recursively call the function for the subdirectory
            convert_files_in_folder(item_path, target_format, subdir_output_folder)

def convert_file(file_path, target_format, output_folder):
    # Get the file name and extension
    file_name, file_extension = os.path.splitext(file_path)

    # Determine the target file path
    output_file_path = os.path.join(output_folder, os.path.basename(file_name) + '.' + target_format)

    print(f"Converting {file_path} to {target_format}. Output file: {output_file_path}")

    if target_format == 'pdf':
        if file_extension.lower() == '.docx':
            convert_docx_to_pdf(file_path, output_file_path)
        elif file_extension.lower() == '.pdf':
            # If it's already a PDF file, simply copy it to the output folder
            shutil.copyfile(file_path, output_file_path)
        else:
            print(f"Error: Unsupported file format for PDF conversion: {file_path}")
    elif target_format == 'docx':
        if file_extension.lower() == '.pdf':
            convert_pdf_to_docx(file_path, output_file_path)
        elif file_extension.lower() == '.docx':
            # If it's already a DOCX file, simply copy it to the output folder
            shutil.copyfile(file_path, output_file_path)
        else:
            print(f"Error: Unsupported file format for DOCX conversion: {file_path}")
    else:
        print(f"Error: Unsupported target format: {target_format}")

def convert_docx_to_pdf(docx_file_path, pdf_file_path):
    # Load the DOCX document
    doc = Document(docx_file_path)
    # Create a PDF canvas
    c = canvas.Canvas(pdf_file_path, pagesize=letter)
    y = 750  # Initial Y coordinate
    # Write each paragraph from the DOCX document to the PDF canvas
    for paragraph in doc.paragraphs:
        c.drawString(100, y, paragraph.text)
        y -= 15  # Move to the next line
    # Save the PDF file
    c.save()
    print(f"File {docx_file_path} converted to PDF and saved to {pdf_file_path}.")


def convert_pdf_to_docx(pdf_file_path, docx_file_path):
    # Convert PDF to DOCX
    cv = Converter(pdf_file_path)
    cv.convert(docx_file_path, start=0, end=None)
    cv.close()
    
    print(f"File {pdf_file_path} converted to DOCX and saved to {docx_file_path}.")

def clean_text(text):
    # Remove any characters that are not XML compatible
    cleaned_text = ''.join(char for char in text if 32 <= ord(char) <= 126 or char == '\n')
    return cleaned_text


def main():
    # Prompt the user for input
    folder_to_convert = r"C:\Users\pasar\Desktop\Converter\Files"
    target_format = input("Enter the target file format you want to convert to: ")
    output_folder = r"C:\Users\pasar\Desktop\Converter\Done"

    # Convert files in the folder
    convert_files_in_folder(folder_to_convert, target_format, output_folder)

if __name__ == "__main__":
    main()
