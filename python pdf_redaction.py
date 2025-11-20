import fitz  # PyMuPDF
import re
import os
from datetime import datetime

# --- CONFIGURATION ---
input_folder = "PDF_Input"
output_folder = "PDF_Cleaned"
log_file = "Batch_Redaction_Log.txt"

# Create folders if not exist
os.makedirs(input_folder, exist_ok=True)
os.makedirs(output_folder, exist_ok=True)

# Custom confidential triggers
confidential_terms = [
    "Hydor.no",
    "Confidential",
    "Client Name",
    "Owner Representative",
    "Surveyor Name"
]

# Regex patterns for sensitive data
patterns = {
    "Email": r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
    "IMO number": r"\bIMO\s?\d{5,7}\b",
    "Phone": r"\+?\d{1,4}[\s-]?\d{3,}[\s-]?\d{3,}",
    "Date": r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",
}

# Redaction mode: "blackout" or "replace"
mode = "blackout"

# --- MAIN PROCESS ---
log_entries = []
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
log_entries.append(f"PDF REDACTION BATCH LOG – {timestamp}\n")

for file_name in os.listdir(input_folder):
    if not file_name.lower().endswith(".pdf"):
        continue

    input_pdf = os.path.join(input_folder, file_name)
    output_pdf = os.path.join(output_folder, f"Cleaned_{file_name}")
    print(f"🔍 Processing: {file_name}")

    try:
        doc = fitz.open(input_pdf)
        file_log = []

        for page_num, page in enumerate(doc, start=1):
            text_blocks = page.get_text("blocks")

            for block in text_blocks:
                block_text = block[4]
                trigger = False
                found_labels = []

                # Keyword triggers
                for term in confidential_terms:
                    if re.search(re.escape(term), block_text, re.IGNORECASE):
                        trigger = True
                        found_labels.append(term)

                # Pattern triggers
                for label, pattern in patterns.items():
                    if re.search(pattern, block_text):
                        trigger = True
                        found_labels.append(label)

                # Redact if triggered
                if trigger:
                    rect = fitz.Rect(block[0], block[1], block[2], block[3])
                    page.add_redact_annot(
                        rect,
                        text="[REDACTED PARAGRAPH]" if mode == "replace" else None,
                        fill=(0, 0, 0)
                    )
                    file_log.append(f"Page {page_num}: Redacted ({', '.join(found_labels)})")

            page.apply_redactions()

        doc.save(output_pdf)
        doc.close()

        log_entries.append(f"\n=== {file_name} ===")
        if file_log:
            log_entries.extend(file_log)
        else:
            log_entries.append("No redactions applied.")
        print(f"✅ Cleaned: {output_pdf}")

    except Exception as e:
        log_entries.append(f"\nERROR processing {file_name}: {str(e)}")
        print(f"⚠️ Error processing {file_name}: {e}")

# --- SAVE MASTER LOG ---
with open(log_file, "w", encoding="utf-8") as f:
    f.write("\n".join(log_entries))

print(f"\n📄 Batch completed. Log saved to: {log_file}")
