# FileConverter

FileConverter is a Python program that enables you to convert files from PDF, DOCX, and image formats to various other file formats.

## Usage

1. Clone this repository to your local machine.
2. Navigate to the directory where the program is located.
3. Run the program by executing the `index.py` file.
4. Follow the prompts to specify the target file format and input/output folders.
5. The program will convert the files in the specified input folder to the chosen format and save them to the output folder. If you choose to convert a folder, the program will create a corresponding folder in the output directory with the same name + `.` + `new format`.

## Supported File Formats

FileConverter is designed to convert files from a variety of formats to the target format of your choice. While it aims to support a wide range of file types, there are some considerations to keep in mind:

- **Supported Formats**: FileConverter supports conversion for various file formats, including but not limited to:
  - PDF (Portable Document Format)
  - DOCX (Microsoft Word Document)
  - Images (e.g., JPEG, PNG, GIF, etc.)
  - Text files (e.g., TXT)
  - and more.

- **Special Cases**: Certain file types, such as PDF and DOCX, may have limitations or special cases that could affect the conversion process. These formats often contain complex layouts, embedded objects, or specialized content that may not translate perfectly during conversion. While FileConverter includes specific functionality to handle PDF and DOCX conversions, users may encounter issues with these file types, especially when dealing with intricate formatting.

- **Unsupported Formats**: While FileConverter allows users to specify any target file format, not all formats are fully integrated or tested. Some formats may not be supported due to compatibility issues, limitations of the conversion libraries, or lack of testing. Users should exercise caution when converting to unsupported formats, as unexpected behavior or errors may occur.

- **Manual Verification**: It's recommended to review the converted files after conversion to ensure that the formatting and content are preserved as expected. In cases where the converted files do not meet your requirements, manual adjustments may be necessary.

- **Error Handling**: If encountering any issues or unexpected behavior during conversion, please check the program's output for error messages or warnings. Additionally, feel free to explore alternative conversion methods or tools that may better suit your needs.

## Dependencies

- Python 3.x
- `comtypes` library
- `pdf2docx` library

## Installation

1. Ensure you have Python 3.x installed on your system.
2. Install the required dependencies using pip:

```
pip install comtypes pdf2docx
```

3. Clone this repository:

```
git clone https://github.com/your_username/FileConverter.git
```
