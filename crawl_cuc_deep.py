import urllib.request
import re
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')
from test_subfolder_contents import fetch_and_decode

with open("all_teachers_subfolders_extracted.json", "r", encoding="utf-8") as f:
    subs = json.load(f)

for sub in subs["cuc"]:
    cycle_name = sub["name"]
    subjects = fetch_and_decode(sub["id"])
    print(f"\n=== Cô Cúc [{cycle_name}] ===")
    total_files = 0
    for subj in subjects:
        if subj["type"] == "application/vnd.google-apps.folder":
            files = fetch_and_decode(subj["id"])
            doc_files = [f for f in files if f["type"] != "application/vnd.google-apps.folder"]
            sub_folders = [f for f in files if f["type"] == "application/vnd.google-apps.folder"]
            print(f"  Môn {subj['name']}: {len(doc_files)} tệp, {len(sub_folders)} thư mục con")
            for df in doc_files:
                print(f"    - {df['name']}")
            total_files += len(doc_files)
        else:
            print(f"  Tệp trực tiếp: {subj['name']}")
            total_files += 1
    print(f"  => Tổng cộng {cycle_name}: {total_files} tệp")
