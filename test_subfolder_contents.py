import urllib.request
import re
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

def fetch_and_decode(folder_id):
    url = f"https://drive.google.com/drive/folders/{folder_id}?usp=sharing"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    try:
        html = urllib.request.urlopen(req, timeout=15).read().decode('utf-8', errors='ignore')
    except Exception as e:
        return {"error": str(e)}

    m = re.search(r"window\['_DRIVE_ivd'\]\s*=\s*'([^']+)'", html)
    if not m:
        return {"error": "no ivd", "length": len(html)}

    raw = m.group(1)
    clean = re.sub(r'\\x([0-9a-fA-F]{2})', lambda m: chr(int(m.group(1), 16)), raw)
    clean = clean.replace(r'\/', '/')

    try:
        byte_arr = bytes([ord(c) for c in clean])
        text_utf8 = byte_arr.decode('utf-8')
        data = json.loads(text_utf8)
    except Exception:
        try:
            data = json.loads(clean)
        except Exception as e:
            return {"error": str(e)}

    items = []
    if data and isinstance(data, list) and len(data) > 0 and isinstance(data[0], list):
        for it in data[0]:
            try:
                fid = it[0]
                fname = it[2]
                ftype = it[3]
                # timestamp
                mod_time = it[9]
                items.append({
                    "id": fid,
                    "name": fname,
                    "type": ftype,
                    "modified_timestamp": mod_time
                })
            except Exception:
                pass
    return items

# Test Cô Cúc Tuần 1-4
items_cuc_t1_4 = fetch_and_decode("1osWHYh3g7O47vyut81mjbOhGMGfhlpUU")
print("Cô Cúc Tuần 1-4 items:", len(items_cuc_t1_4))
for it in items_cuc_t1_4:
    print(" ", it)

# Test Cô Cúc Tuần 5-8
items_cuc_t5_8 = fetch_and_decode("1yaWWTjqy7-ytEEkoLEr5lb6UOnD6M-4j")
print("\nCô Cúc Tuần 5-8 items:", len(items_cuc_t5_8))
for it in items_cuc_t5_8:
    print(" ", it)

# Test Thầy Thắng Tuần 1-4 (let's get Thầy Thắng's subfolder ID from extracted)
with open("all_teachers_subfolders_extracted.json", "r", encoding="utf-8") as f:
    subs = json.load(f)

thang_t1_4_id = subs["thang"][0]["id"]
items_thang_t1_4 = fetch_and_decode(thang_t1_4_id)
print(f"\nThầy Thắng Tuần 1-4 ({subs['thang'][0]['name']}) items:", len(items_thang_t1_4))
for it in items_thang_t1_4:
    print(" ", it)
