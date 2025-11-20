import os
from docx import Document

# === SETTINGS ===
root_folder = r"C:\Users\judep\Downloads\SMS FOR EDITING_VER 1"  # 🔹 Change to your folder path
new_author = "LMMS"  # 🔹 The author name you want to apply

# === SCRIPT ===
count = 0
for root, _, files in os.walk(root_folder):
    for file in files:
        if file.lower().endswith(".docx"):
            path = os.path.join(root, file)
            try:
                doc = Document(path)
                core_props = doc.core_properties
                core_props.author = new_author
                core_props.last_modified_by = new_author
                doc.save(path)
                count += 1
                print(f"✅ Updated author in: {file}")
            except Exception as e:
                print(f"⚠️ Could not update {file}: {e}")

print(f"\nDone! Updated author in {count} files.")
