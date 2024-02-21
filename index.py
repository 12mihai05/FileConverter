import os
import shutil
import comtypes.client
from pdf2docx import Converter

def convert_files_in_folder(folder_path, target_format, output_folder):
    if not os.path.exists(folder_path):
        print(f"Error: Folder {folder_path} not found.")
        return

    items = os.listdir(folder_path)

    for item in items:
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path):
            convert_file(item_path, target_format, output_folder)
        elif os.path.isdir(item_path):
            subdir_output_folder = os.path.join(output_folder, item)
            os.makedirs(subdir_output_folder, exist_ok=True)
            convert_files_in_folder(item_path, target_format, subdir_output_folder)


def convert_file(file_path, target_format, output_folder):
    file_name, file_extension = os.path.splitext(file_path)
    output_file_path = os.path.join(output_folder, os.path.basename(file_name) + '.' + target_format)

    print(f"Converting {file_path} to {target_format}. Output file: {output_file_path}")

    if target_format == 'pdf':
        if file_extension.lower() == '.docx':
            convert_docx_to_pdf(file_path, output_file_path)
        elif file_extension.lower() == '.pdf':
            shutil.copyfile(file_path, output_file_path)
        else:
            print(f"Error: Unsupported file format for PDF conversion: {file_path}")
    elif target_format == 'docx':
        if file_extension.lower() == '.pdf':
            convert_pdf_to_docx(file_path, output_file_path)
        elif file_extension.lower() == '.docx':
            shutil.copyfile(file_path, output_file_path)
        else:
            print(f"Error: Unsupported file format for DOCX conversion: {file_path}")
    else:
        print(f"Error: Unsupported target format: {target_format}")

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

def clean_text(text):
    cleaned_text = ''.join(char for char in text if 32 <= ord(char) <= 126 or char == '\n')
    return cleaned_text

def main():
    folder_to_convert = r"C:\Users\pasar\Desktop\Converter\Files"
    target_format = input("Enter the target file format you want to convert to: ")
    output_folder = r"C:\Users\pasar\Desktop\Converter\Done"

    convert_files_in_folder(folder_to_convert, target_format, output_folder)

if __name__ == "__main__":
    main()
