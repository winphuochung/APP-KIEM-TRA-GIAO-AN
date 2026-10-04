import urllib.request
import re
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open("all_teachers_subfolders_extracted.json", "r", encoding="utf-8") as f:
    subs = json.load(f)

# check subfolders in cuc
from test_subfolder_contents import fetch_and_decode

for sub in subs["cuc"]:
    items = fetch_and_decode(sub["id"])
    print(f"Cô Cúc [{sub['name']}]: {len(items)} items")
    for it in items:
        print(f"   -> {it['name']} ({it['type']})")
