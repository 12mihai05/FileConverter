import os
import shutil
import comtypes.client
from pdf2docx import Converter
import moviepy.editor as moviepy
import subprocess

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
            print(f"Processing file: {item_path}")
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
    
    movie_files = ['gif', 'mp4', 'avi', 'webm', 'mov']

    # Check if the target format is PDF, DOCX, or one of the movie files
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
    elif target_format.lower() in movie_files:
        if target_format.lower() == 'gif':
            convert_video_gif(file_path, output_file_path)
        elif file_extension.lower() == '.mov':
            convert_video_mov(file_path,output_file_path,target_format)
        elif target_format.lower() == 'mov':
            convert_mov_video(file_path, output_file_path, target_format)
            print("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")
        elif file_extension.lower() == '.gif':
            convert_gif_video(file_path ,output_file_path,target_format)
        elif target_format.lower() in movie_files:
            convert_webm_avi_mp4(file_path,output_file_path)
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

def convert_webm_avi_mp4(input_file_path, output_file_path):
    clip = moviepy.VideoFileClip(input_file_path)
    clip.write_videofile(output_file_path)


def convert_webm_avi_mp4(input_file_path, output_file_path):
    try:
        # Load the video clip
        clip = moviepy.VideoFileClip(input_file_path)

        # Get the frame rate of the video clip
        fps = clip.fps

        # Write the frames to a file based on the output file path extension
        if output_file_path.lower().endswith('.avi'):
            # For AVI format, use libx264 codec
            clip.write_videofile(output_file_path, codec='libx264', fps=fps)
        else:
            # For other formats, use default codec
            clip.write_videofile(output_file_path)
        
        print(f"Successfully converted to {output_file_path.split('.')[-1].upper()}: {output_file_path}")
    except Exception as e:
        print(f"Error converting video: {e}")


def convert_video_gif(input_file_path, output_file_path):
    try:
        # Load the video clip
        clip = moviepy.VideoFileClip(input_file_path)

        # Get the duration of the video clip
        duration = clip.duration

        # Get the frame rate of the video clip
        fps = clip.fps

        # Set the start and end time for the clip
        start_time = 0
        end_time = duration

        # Extract frames from the video clip
        clip = clip.subclip(start_time, end_time)

        # Write the frames to a GIF file
        clip.write_gif(output_file_path, fps=fps)
        print("Successfully converted to GIF:", output_file_path)
    except Exception as e:
        print("Error converting to GIF:", e)

def convert_mov_video(input_file_path, output_file_path, target_format):
    try:
        # Define the codec based on the target format
        if target_format.lower() == 'mp4':
            codec = 'libx264'  # H.264 codec for MP4
        elif target_format.lower() == 'webm':
            codec = 'libvpx'   # VP8 codec for WebM
        elif target_format.lower() == 'avi':
            codec = 'rawvideo' # Raw video codec for AVI
        elif target_format.lower() == 'mov':
            codec = 'copy'     # Copy codec for MOV (no transcoding)
        elif target_format.lower() == 'gif':
            codec = 'gif'      # GIF codec for GIF
        else:
            print(f"Unsupported target format: {target_format}")
            return

        # Specify ffmpeg command
        ffmpeg_command = [
            'ffmpeg',
            '-i', input_file_path,  # Input file
            '-c:v', codec,           # Codec for the target format
            '-pix_fmt', 'yuv420p',   # Pixel format for compatibility
            output_file_path        # Output file path
        ]

        # Run ffmpeg command
        subprocess.run(ffmpeg_command, check=True)

        print(f"Successfully converted video to {target_format.upper()}: {output_file_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error converting video to {target_format.upper()}: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")



#aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa



def convert_video_mov(input_file_path, output_file_path, target_format):
    try:
        # Define dictionary to map target formats to FFmpeg commands
        ffmpeg_commands = {
            'mp4': [
                'ffmpeg', '-i', input_file_path,
                '-vf', 'scale=trunc(iw/2)*2:trunc(ih/2)*2',
                '-c:v', 'libx264',
                '-pix_fmt', 'yuv420p',
                '-preset', 'slow',
                '-c:a', 'aac',
                '-b:a', '128k',
                '-movflags', 'faststart',
                output_file_path
            ],
            'gif': [
                'ffmpeg', '-i', input_file_path,
                '-vf', 'scale=trunc(iw/2)*2:trunc(ih/2)*2',
                '-c:v', 'libx264',  # MOV format with H.264 codec
                '-pix_fmt', 'yuv420p',
                '-c:a', 'aac',
                '-movflags', 'faststart',
                output_file_path
            ],
            'webm': [
            'ffmpeg', '-i', input_file_path,
            '-c:v', 'libx264',
            '-pix_fmt', 'yuv420p',
            '-c:a', 'aac',
            '-movflags', 'faststart',
            output_file_path
        ],
            'avi': [
                'ffmpeg', '-i', input_file_path,
                '-c:v', 'libx264',
                '-c:a', 'aac',
                output_file_path
            ]
        }

        # Get the corresponding FFmpeg command based on the target format
        ffmpeg_command = ffmpeg_commands.get(target_format)

        # Check if the target format is supported
        if not ffmpeg_command:
            print(f"Unsupported target format: {target_format}")
            return

        # Execute the FFmpeg command
        subprocess.run(ffmpeg_command, capture_output=True, check=True, text=True)

        print(f"Successfully converted {input_file_path} to MOV: {output_file_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error converting {input_file_path} to MOV: {e.stderr}")
    except Exception as e:
        print(f"An error occurred: {e}")



def convert_gif_video(input_file_path, output_file_path, target_format):
    try:
        # Define the codec based on the target format
        if target_format == 'mp4':
            codec = 'libx264'  # H.264 codec for MP4
        elif target_format == 'webm':
            codec = 'libvpx'   # VP8 codec for WebM
        elif target_format == 'avi':
            codec = 'rawvideo' # Raw video codec for AVI
        else:
            print(f"Unsupported target format: {target_format}")
            return

        # Specify ffmpeg command
        ffmpeg_command = [
            'ffmpeg',
            '-i', input_file_path,  # Input GIF file
            '-vf', 'scale=trunc(iw/2)*2:trunc(ih/2)*2',  # Ensure width and height are divisible by 2
            '-c:v', codec,  # Codec for the target format
            '-pix_fmt', 'yuv420p',  # Pixel format for compatibility
            output_file_path  # Output file path
        ]

        # Run ffmpeg command
        subprocess.run(ffmpeg_command, check=True)

        print(f"Successfully converted GIF to {target_format.upper()}: {output_file_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error converting GIF to {target_format.upper()}: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")


def main():
    # Prompt the user for input
    folder_to_convert = r"C:\Users\pasar\Desktop\Converter\Files"
    target_format = input("Enter the target file format you want to convert to: ")
    output_folder = r"C:\Users\pasar\Desktop\Converter\Done"

    # Convert files in the folder
    convert_files_in_folder(folder_to_convert, target_format, output_folder)

if __name__ == "__main__":
    main()
