import os
import shutil

def convert_files_in_folder(folder_path, target_format, output_folder):
    # Check if the folder exists
    if not os.path.isdir(folder_path):
        print(f"Error: Folder {folder_path} not found.")
        return

    # Iterate over all files and subdirectories in the folder
    for root, dirs, files in os.walk(folder_path):
        for item in dirs[:]:  # Make a copy of the list to iterate over
            subdir_path = os.path.join(root, item)
            # Create a corresponding folder in the output directory with the target format appended
            subdir_output_folder = os.path.join(output_folder, item + '.' + target_format)
            convert_files_in_folder(subdir_path, target_format, subdir_output_folder)

        for file in files:
            # Get the full path of the file
            file_path = os.path.join(root, file)
            # Convert and copy the file
            convert_file(file_path, target_format, folder_path, output_folder)

def convert_file(file_path, target_format, input_folder, output_folder):
    # Get the relative path of the file within the input folder
    relative_path = os.path.relpath(file_path, input_folder)
    # Construct the output file path
    output_file_path = os.path.join(output_folder, relative_path)
    # Create the output directory if it doesn't exist
    os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
    # Construct the path for the converted file
    converted_file_path = os.path.splitext(output_file_path)[0] + '.' + target_format
    # Convert and copy the file
    shutil.copyfile(file_path, converted_file_path)
    print(f"File {file_path} converted to {target_format} and saved to {converted_file_path}.")

def main():
    # Prompt the user for input
    folder_to_convert = r"C:\Users\pasar\Desktop\Converter\Files"
    target_format = input("Enter the target file format you want to convert to: ")
    output_folder = r"C:\Users\pasar\Desktop\Converter\Done"

    # Convert files in the folder
    convert_files_in_folder(folder_to_convert, target_format, output_folder)

if __name__ == "__main__":
    main()
