import zipfile
import os
import shutil
import tempfile

# --- SETTINGS ---
input_folder = r"C:\Users\judep\Downloads\FORMS EDITING\5. VEM"  # Folder containing .docx files
output_folder = r"C:\Users\judep\Downloads\FORMS EDITING\UNLOCKED"  # Where to save unlocked files

# ------------------

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

count_unlocked = 0
count_failed = 0

# Walk through all subdirectories
for foldername, subfolders, filenames in os.walk(input_folder):
    for filename in filenames:
        if not filename.endswith(".docx"):
            continue
        
        input_file = os.path.join(foldername, filename)
        output_file = os.path.join(output_folder, filename)
        
        try:
            # Create temporary folder
            temp_dir = tempfile.mkdtemp()
            
            # Extract docx (ZIP)
            with zipfile.ZipFile(input_file, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            settings_path = os.path.join(temp_dir, "word", "settings.xml")
            
            # Check if settings.xml exists
            if not os.path.exists(settings_path):
                print(f"⚠️ No settings.xml in {filename}, skipping.")
                shutil.rmtree(temp_dir)
                continue
            
            # Remove protection tags
            with open(settings_path, "r", encoding="utf-8") as f:
                xml = f.read()
            
            # Tags that lock editing
            tags_to_remove = [
                "w:documentProtection",
                "w:writeProtection",
                "w:readOnlyRecommended",
                "w:enforcement"
            ]
            
            protection_found = False
            for tag in tags_to_remove:
                while tag in xml:
                    start = xml.find(f"<{tag}")
                    if start == -1:
                        break
                    end = xml.find("/>", start)
                    if end == -1:
                        break
                    xml = xml[:start] + xml[end+2:]
                    protection_found = True
            
            # Write cleaned XML
            with open(settings_path, "w", encoding="utf-8") as f:
                f.write(xml)
            
            # Repack into new unlocked file
            with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)
            
            # Clean up
            shutil.rmtree(temp_dir)
            
            status = "🔓 Unlocked" if protection_found else "✅ No protection found"
            print(f"{status}: {filename}")
            count_unlocked += 1
            
        except Exception as e:
            print(f"❌ Failed to process {filename}: {e}")
            count_failed += 1
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            continue

print(f"\n✅ Processed {count_unlocked} files.")
print(f"❌ Failed: {count_failed} files.")
print(f"📁 Unlocked files saved to: {output_folder}")

print("Unlocked file saved to:", output_file)
