import os
import shutil

def copy_excel_files_from_subfolders(src_folder, dest_folder, file_extension='.xlsx'):
    # Check if source folder exists
    if not os.path.exists(src_folder):
        raise FileNotFoundError(f"The source folder '{src_folder}' does not exist.")
    
    # Check if destination folder exists, create if not
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
    
    # Traverse through all subfolders in the source folder
    for subdir, _, files in os.walk(src_folder):
        for file in files:
            if file.endswith(file_extension):
                src_file_path = os.path.join(subdir, file)
                dest_file_path = os.path.join(dest_folder, file)

                shutil.copy2(src_file_path, dest_file_path)
                print(f"Copied '{src_file_path}' to '{dest_file_path}'.")

# Example usage
src_folder = "C:\\Users\\AShanbhag\\Documents\\New folder (4)"
dest_folder = "C:\\Users\\AShanbhag\\Documents\\offline conv"
copy_excel_files_from_subfolders(src_folder, dest_folder)