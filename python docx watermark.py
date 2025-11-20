import os
import shutil
import time
import win32com.client as win32
from pathlib import Path

# Configuration for watermark appearance
WATERMARK_TEXT = "CONFIDENTIAL"   # Text in watermark
WATERMARK_FONT = "Arial"          # Font used
WATERMARK_SIZE = 120              # Size of watermark text (increase for bigger)
WATERMARK_COLOR = 12632256        # Gray color RGB (light gray)
WATERMARK_ROTATION = -45          # Rotate watermark diagonally
WATERMARK_TRANSPARENCY = 0.8      # Transparency (0.0 opaque - 1.0 fully transparent)


# Starting directory (change to your folder)
START_DIR = r"C:\Users\judep\Downloads\FORMS EDITING\UNLOCKED"

def add_watermark_to_doc(doc_path):
    """Add a CONFIDENTIAL-style diagonal watermark to a DOCX file."""
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(doc_path)

        # Add watermark text effect as a shape
        shape = doc.Shapes.AddTextEffect(
            PresetTextEffect=1,  # msoTextEffect1 basic style
            Text=WATERMARK_TEXT,
            FontName=WATERMARK_FONT,
            FontSize=WATERMARK_SIZE,
            FontBold=False,
            FontItalic=False,
            Left=0,
            Top=0
        )

        # Set position relative to page
        shape.RelativeHorizontalPosition = 1  # wdRelativeHorizontalPositionPage
        shape.RelativeVerticalPosition = 1    # wdRelativeVerticalPositionPage
        
        # Center the watermark on the page
        shape.Left = -914400 / 2  # Center horizontally (in EMUs: -0.5 inches in points)
        shape.Top = -914400 / 2   # Center vertically
        
        # Alternative: Use LockAnchor and calculate based on page dimensions
        page_width = doc.PageSetup.PageWidth
        page_height = doc.PageSetup.PageHeight
        
        shape.Left = page_width / 2
        shape.Top = page_height / 2

        # Set style to appear behind text and transparency
        shape.Fill.ForeColor.RGB = WATERMARK_COLOR
        shape.Fill.Transparency = WATERMARK_TRANSPARENCY
        shape.Line.Transparency = WATERMARK_TRANSPARENCY
        shape.TextEffect.FontName = WATERMARK_FONT
        shape.TextEffect.FontSize = WATERMARK_SIZE
        shape.Rotation = WATERMARK_ROTATION
        shape.WrapFormat.Type = 3  # wdWrapBehind
        shape.LockAnchor = True
        shape.ZOrder(5)  # msoSendBehindText

        doc.Save()
        doc.Close()
        word.Quit()
        time.sleep(1)
        print(f"Watermark added to: {doc_path}")
    except Exception as e:
        print(f"Error processing {doc_path}: {str(e)}")

def process_directory(root_dir):
    """Recursively add watermark to all .docx files (excluding temp and backup files)."""
    root_path = Path(root_dir)
    if not root_path.exists():
        print(f"Directory does not exist: {root_dir}")
        return
    
    for file_path in root_path.rglob("*.docx"):
        if file_path.stem.startswith('~$'):
            print(f"Skipping temporary file: {file_path}")
            continue
        if ".backup" in file_path.name:
            print(f"Skipping backup file: {file_path}")
            continue
        
        # Optional: Create backup before watermarking, comment if not needed
        backup_path = file_path.with_suffix(".backup.docx")
        if not backup_path.exists():
            try:
                shutil.copy2(str(file_path), str(backup_path))
                print(f"Backup created: {backup_path}")
            except Exception as e:
                print(f"Warning: Could not create backup for {file_path}. Skipping. Error: {e}")
                continue
        
        add_watermark_to_doc(str(file_path))

if __name__ == "__main__":
    print("Starting watermark process...")
    process_directory(START_DIR)
    print("Process complete.")