import os
import time
from datetime import datetime

# --- SETTINGS ---
root_folder = r"C:\Users\judep\Downloads\FORMS EDITING\1. Accounting Forms"  # Change this
new_edit_date = "2025-11-11"              # Date only (YYYY-MM-DD)

# Convert date to timestamp
edit_timestamp = time.mktime(datetime.strptime(new_edit_date, "%Y-%m-%d").timetuple())

changed_files = []

for root, dirs, files in os.walk(root_folder):
    for file in files:
        # Only process .xlsx files
        if file.lower().endswith(".xlsx"):
            full_path = os.path.join(root, file)
            try:
                os.utime(full_path, (edit_timestamp, edit_timestamp))
                changed_files.append(full_path)
                print("Updated date:", full_path)
            except Exception as e:
                print("Failed to update:", full_path, "| Error:", e)

print("\n=== DATE CHANGE COMPLETE ===")
print(f"Total .xlsx files updated: {len(changed_files)}\n")

for f in changed_files:
    print(" -", f)
