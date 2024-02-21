import os
import shutil
import comtypes.client
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
            subdir_output_folder = os.path.join(output_folder, item + '.' + target_format)
            os.makedirs(subdir_output_folder, exist_ok=True)
            # Recursively call the function for the subdirectory
            convert_files_in_folder(item_path, target_format, subdir_output_folder)

def convert_file(file_path, target_format, output_folder):
    # Get the file name and extension
    file_name, file_extension = os.path.splitext(file_path)
    
    # Construct the output file path
    output_file_path = os.path.join(output_folder, os.path.basename(file_name) + '.' + target_format)
    
    # Check if the target format is PDF or DOCX
    if target_format.lower() == 'pdf':
        if file_extension.lower() == '.docx':
            convert_docx_to_pdf(file_path, output_file_path)
        else:
            # Unsupported conversion, simply copy the file
            shutil.copyfile(file_path, output_file_path)
            print(f"Unsupported conversion for {file_path}. File copied to {output_file_path}.")
    elif target_format.lower() == 'docx':
        if file_extension.lower() == '.pdf':
            convert_pdf_to_docx(file_path, output_file_path)
        else:
            # Unsupported conversion, simply copy the file
            shutil.copyfile(file_path, output_file_path)
            print(f"Unsupported conversion for {file_path}. File copied to {output_file_path}.")
    else:
        # For other formats, simply copy the file
        shutil.copyfile(file_path, output_file_path)
        print(f"File {file_path} converted to {target_format} and saved to {output_file_path}.")

def convert_docx_to_pdf(docx_file_path, pdf_file_path):
    # Initialize COM
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False

    try:
        # Open the Word document
        docx_path = os.path.abspath(docx_file_path)
        pdf_path = os.path.abspath(pdf_file_path)
        in_file = word.Documents.Open(docx_path)

        # Save the Word document as PDF
        pdf_format = 17  # PDF file format code
        in_file.SaveAs(pdf_path, FileFormat=pdf_format)
        in_file.Close()
        print(f"File {docx_file_path} converted to PDF and saved to {pdf_file_path}.")
    except Exception as e:
        print(f"Error converting {docx_file_path} to PDF: {e}")
    finally:
        # Quit Microsoft Word
        word.Quit()

def convert_pdf_to_docx(pdf_file_path, docx_file_path):
    cv = Converter(pdf_file_path)
    cv.convert(docx_file_path, start=0, end=None)
    cv.close()
    
    print(f"File {pdf_file_path} converted to DOCX and saved to {docx_file_path}.")

def main():
    # Prompt the user for input
    folder_to_convert = r"C:\Users\pasar\Desktop\Converter\Files"
    target_format = input("Enter the target file format you want to convert to: ")
    output_folder = r"C:\Users\pasar\Desktop\Converter\Done"

    # Convert files in the folder
    convert_files_in_folder(folder_to_convert, target_format, output_folder)

if __name__ == "__main__":
    main()
