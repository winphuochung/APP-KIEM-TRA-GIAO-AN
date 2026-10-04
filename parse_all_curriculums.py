import pypdf
import os
import re
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

DATA_DIR = r"D:\APP-KIEM-TRA-GIAO-AN"

pl_mapping = {
    "KHTN 6": "PL17_ KHTN 6 - 2026-2027-NLS.pdf",
    "KHTN 7": "PL18_ KHTN 7 - 26-27_NLS.pdf",
    "KHTN 8": "PL19_ KHTN 8 năm học 2026-2027_NLS.pdf",
    "KHTN 9": "PL20_ môn KHTN 9_26-27_NLS.pdf",
    "CN 6": "PL21_ CN 6 26 -27_NLS.pdf",
    "CN 7": "PL22_ CN 7 26 - 27_NLS.pdf",
    "CN 8": "PL24_ Công nghệ 8 2026 2027_NLS.pdf",
    "CN 9": "PL23_ Công nghệ 9 2026 2027_NLS.pdf",
    "Toán 6": "PL25_ Toán 6 (2026-2027)_NLS.pdf",
    "Toán 7": "PL26_ Toan_7_2026-2027_NLS.pdf",
    "Toán 8": "PL27_ TOÁN 8 (2026 - 2027)_NLS.pdf",
    "Toán 9": "PL28_ Toán 9 (2026-2027)_NLS.pdf",
    "HĐTN 6": "PL29_ HĐTN-HN 6 (KNTTVCS) 26-27.pdf",
    "HĐTN 7": "PL30_ HĐTN K7.pdf",
    "HĐTN 8": "PL31_ HĐTN K8_26-27_NLS.pdf",
    "HĐTN 9": "PL32_ HĐTN K9_26-27_NLS.pdf",
}

curriculum_db = {}

for subj_key, fname in pl_mapping.items():
    fpath = os.path.join(DATA_DIR, fname)
    if not os.path.exists(fpath):
        prefix = fname.split()[0]
        for f in os.listdir(DATA_DIR):
            if f.startswith(prefix) and f.endswith(".pdf"):
                fpath = os.path.join(DATA_DIR, f)
                break
    
    if not os.path.exists(fpath):
        print(f"File not found: {fname}")
        continue
    
    reader = pypdf.PdfReader(fpath)
    full_text = ""
    for page in reader.pages:
        full_text += page.extract_text() + "\n"
    
    curriculum_db[subj_key] = {
        "file": os.path.basename(fpath),
        "total_pages": len(reader.pages),
        "text_length": len(full_text),
        "full_text": full_text
    }
    print(f"Loaded {subj_key}: {len(reader.pages)} pages, {len(full_text)} chars")

with open(os.path.join(DATA_DIR, "data", "curriculum_plans_db.json"), "w", encoding="utf-8") as f:
    # save metadata and snippets
    meta = {k: {"file": v["file"], "total_pages": v["total_pages"], "sample": v["full_text"][:500]} for k, v in curriculum_db.items()}
    json.dump(meta, f, ensure_ascii=False, indent=2)

print("\nSuccessfully parsed all 16 curriculum plans!")
