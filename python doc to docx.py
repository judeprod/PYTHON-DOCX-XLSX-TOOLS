import os
import zipfile
import win32com.client as win32

def is_docx(file_path):
    """Check if a DOCX file is real (ZIP-based)."""
    if not file_path.lower().endswith(".docx"):
        return False
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            return "word/document.xml" in z.namelist()
    except:
        return False


def convert_doc_to_docx(input_path, converted_list):
    """Convert .doc or .dot files to .docx using Microsoft Word."""
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Open(input_path)
        output_path = input_path + "x"  # file.doc → file.docx
        doc.SaveAs(output_path, FileFormat=16)  # 16 = wdFormatXMLDocument
        doc.Close()

        converted_list.append((input_path, output_path))

    except Exception as e:
        print(f"[ERROR] Cannot convert {input_path}: {e}")

    finally:
        word.Quit()


def scan_and_convert(folder):
    docx_files = []
    old_files = []
    converted_files = []   # <-- THIS WILL STORE CONVERTED FILES

    for root, dirs, files in os.walk(folder):
        for filename in files:
            full_path = os.path.join(root, filename)
            ext = filename.lower().split(".")[-1]

            # Old formats (.doc, .dot)
            if ext in ["doc", "dot"]:
                old_files.append(full_path)
                convert_doc_to_docx(full_path, converted_files)

            # Modern .docx
            elif ext == "docx":
                if is_docx(full_path):
                    docx_files.append(full_path)
                else:
                    old_files.append(full_path + "  (INVALID DOCX)")
    
    return docx_files, old_files, converted_files


# ============================
# SET FOLDER HERE
# ============================
folder_to_scan = r"C:\Users\judep\Downloads\SMS FOR EDITING_VER 1"

docx_files, old_files, converted_files = scan_and_convert(folder_to_scan)

print("\n=== VALID DOCX FILES ===")
for f in docx_files:
    print(" -", f)

print("\n=== OLD / LEGACY FILES FOUND ===")
for f in old_files:
    print(" -", f)

print("\n==============================")
print(" FILES SUCCESSFULLY CONVERTED ")
print("==============================")
if converted_files:
    for old, new in converted_files:
        print(f"Converted: {old}  →  {new}")
else:
    print("No files were converted.")
