import pypdf
import sys

sys.stdout.reconfigure(encoding='utf-8')

reader = pypdf.PdfReader(r"D:\APP-KIEM-TRA-GIAO-AN\PL17_ KHTN 6 - 2026-2027-NLS.pdf")
print("=== PL17 KHTN 6 Pages 1 to 4 ===")
for i in range(1, 4):
    print(f"\n--- PAGE {i} ---")
    print(reader.pages[i].extract_text()[:1200])

reader18 = pypdf.PdfReader(r"D:\APP-KIEM-TRA-GIAO-AN\PL18_ KHTN 7 - 26-27_NLS.pdf")
print("\n=== PL18 KHTN 7 Pages 1 to 3 ===")
for i in range(1, 3):
    print(f"\n--- PAGE {i} ---")
    print(reader18.pages[i].extract_text()[:1200])
