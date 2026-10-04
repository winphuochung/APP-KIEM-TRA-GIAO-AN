import re
import json

teacher_keys = ["cuc", "hang", "ke", "linh", "thuy", "trang", "thanh", "thang"]

all_teachers_subfolders = {}

for k in teacher_keys:
    with open(f"drive_{k}.html", "r", encoding="utf-8") as f:
        html = f.read()

    m = re.search(r"window\['_DRIVE_ivd'\]\s*=\s*'([^']+)'", html)
    if not m:
        continue
    
    raw = m.group(1)
    clean = re.sub(r'\\x([0-9a-fA-F]{2})', lambda m: chr(int(m.group(1), 16)), raw)
    clean = clean.replace(r'\/', '/')
    
    # fix mojibake if any: raw bytes were utf-8
    try:
        # the characters in clean represent raw bytes
        byte_arr = bytes([ord(c) for c in clean])
        text_utf8 = byte_arr.decode('utf-8')
        data = json.loads(text_utf8)
    except Exception as e:
        try:
            data = json.loads(clean)
        except Exception as e2:
            print(f"Error {k}: {e2}")
            continue

    folders = []
    # items are in data[0]
    if data and isinstance(data, list) and len(data) > 0 and isinstance(data[0], list):
        for item in data[0]:
            try:
                fid = item[0]
                fname = item[2]
                ftype = item[3]
                # timestamp in ms
                mod_time = item[9]
                folders.append({
                    "id": fid,
                    "name": fname,
                    "type": ftype,
                    "modified_timestamp": mod_time
                })
            except Exception as ex:
                pass

    all_teachers_subfolders[k] = folders
    print(f"Teacher {k}: found {len(folders)} subfolders/items")

with open("all_teachers_subfolders_extracted.json", "w", encoding="utf-8") as f:
    json.dump(all_teachers_subfolders, f, ensure_ascii=False, indent=2)

print("Saved all_teachers_subfolders_extracted.json")
