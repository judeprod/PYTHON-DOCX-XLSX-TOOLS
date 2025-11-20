import os
from openpyxl import load_workbook
from openpyxl.packaging.core import DocumentProperties

# === SETTINGS ===
root_folder = r"C:\Users\judep\Downloads\FORMS EDITING\1. Accounting Forms"  # 🔹 Folder with Excel files
new_author = "LMMS"  # 🔹 New author name

# === SCRIPT ===
count = 0
for root, _, files in os.walk(root_folder):
    for file in files:
        if file.lower().endswith(".xlsx"):
            path = os.path.join(root, file)
            try:
                wb = load_workbook(path)
                props = wb.properties

                # Optional: Print old author before changing
                print(f"📄 {file} — Old author: {props.creator or 'None'}")

                # Update author fields
                props.creator = new_author
                props.lastModifiedBy = new_author
                wb.properties = props

                wb.save(path)
                wb.close()
                count += 1
                print(f"✅ Updated author in: {file}")
            except Exception as e:
                print(f"⚠️ Could not update {file}: {e}")

print(f"\nDone! Updated author in {count} Excel files.")
