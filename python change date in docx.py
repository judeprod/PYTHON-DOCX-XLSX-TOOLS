import os
import time
from datetime import datetime

# --- SETTINGS ---
root_folder = r"C:\Users\judep\Downloads\SMS FOR EDITING_VER 1"   # Change this
new_edit_date = "2025-10-22"              # Date only (YYYY-MM-DD)

# Convert date to timestamp
edit_timestamp = time.mktime(datetime.strptime(new_edit_date, "%Y-%m-%d").timetuple())

changed_files = []

for root, dirs, files in os.walk(root_folder):
    for file in files:
        # Only process real .docx files
        if file.lower().endswith(".docx"):
            full_path = os.path.join(root, file)
            try:
                os.utime(full_path, (edit_timestamp, edit_timestamp))
                changed_files.append(full_path)
                print("Updated date:", full_path)
            except Exception as e:
                print("Failed to update:", full_path, "| Error:", e)

print("\n=== DATE CHANGE COMPLETE ===")
print(f"Total .docx files updated: {len(changed_files)}\n")

for f in changed_files:
    print(" -", f)
