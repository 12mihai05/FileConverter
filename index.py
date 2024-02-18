import os
import shutil

def convert_files_in_folder(folder_path, target_format, output_folder):
    # Check if the folder exists
    if not os.path.isdir(folder_path):
        print(f"Error: Folder {folder_path} not found.")
        return

    # Get a list of all files in the folder
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    
    print("Files found in folder:")
    print(files)

    # Print the target format for debugging
    print(f"Target format: {target_format}")

    # Iterate over each file in the folder
    for file in files:
        # Get the file extension
        _, file_extension = os.path.splitext(file)

        print(f"Checking file: {file}, Extension: {file_extension}")

        # Check if the file extension matches any of the target formats
 # Add more formats as needed
        print(f"Converting file: {file}")
            # Construct the file paths
        file_path = os.path.join(folder_path, file)
        output_file_path = os.path.join(output_folder, os.path.basename(file))

            # Convert and copy the file
        convert_file(file_path, target_format, output_file_path)





def convert_file(file_path, target_format, output_file_path):
    # Check if the file exists
    if not os.path.isfile(file_path):
        print(f"Error: File {file_path} not found.")
        return

    # Print the file paths for debugging
    print("Converting file:")
    print(f"Source file: {file_path}")
    print(f"Output file: {output_file_path}")

    # Construct the path for the converted file
    converted_file_path = output_file_path.split('.')[0] + '.' + target_format

    # Convert the file and copy it to the output folder
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
