import os

root_folder = r"C:\Users\judep\Downloads\FORMS EDITING\drive-download-20251120T081400Z-1-001"
old_term = "BOM"
new_term = "VOM"

# Allowed file extensions
allowed_ext = {".docx", ".xlsx"}

# 1. Rename files first (top-down)
for current_path, folders, files in os.walk(root_folder, topdown=True):

    for filename in files:
        ext = os.path.splitext(filename)[1].lower()
        if ext not in allowed_ext:
            continue

        old_path = os.path.join(current_path, filename)
        new_filename = filename.replace(old_term, new_term)

        if new_filename != filename:
            new_path = os.path.join(current_path, new_filename)
            os.rename(old_path, new_path)
            print(f"File renamed: {filename}  →  {new_filename}")

# 2. Rename folders after contents (bottom-up)
for current_path, folders, files in os.walk(root_folder, topdown=False):

    for folder in folders:
        old_folder_path = os.path.join(current_path, folder)
        new_folder_name = folder.replace(old_term, new_term)

        if new_folder_name != folder:
            new_folder_path = os.path.join(current_path, new_folder_name)
            os.rename(old_folder_path, new_folder_path)
            print(f"Folder renamed: {folder}  →  {new_folder_name}")

print("Done!")
