import pypdf
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

pl_files = {
    "KHTN 6": "PL17_ KHTN 6 - 2026-2027-NLS.pdf",
    "KHTN 7": "PL18_ KHTN 7 - 26-27_NLS.pdf",
    "Toán 8": "PL27_ TOÁN 8 (2026 - 2027)_NLS.pdf",
    "CN 6": "PL21_ CN 6 26 -27_NLS.pdf"
}

for name, fname in pl_files.items():
    fpath = os.path.join(r"D:\APP-KIEM-TRA-GIAO-AN", fname)
    if not os.path.exists(fpath):
        # find by prefix
        prefix = fname.split()[0]
        for f in os.listdir(r"D:\APP-KIEM-TRA-GIAO-AN"):
            if f.startswith(prefix) and f.endswith(".pdf"):
                fpath = os.path.join(r"D:\APP-KIEM-TRA-GIAO-AN", f)
                break
    print(f"\n=================== {name} ({os.path.basename(fpath)}) ===================")
    reader = pypdf.PdfReader(fpath)
    print(f"Total pages: {len(reader.pages)}")
    text_sample = ""
    for p in reader.pages[:3]:
        text_sample += p.extract_text() or ""
    lines = [line.strip() for line in text_sample.split("\n") if line.strip()]
    for l in lines[:25]:
        print("  ", l)
