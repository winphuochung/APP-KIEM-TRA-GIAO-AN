import re
import json

def parse_drive_ivd(html_path):
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    # Find window['_DRIVE_ivd'] = '...'
    m = re.search(r"window\['_DRIVE_ivd'\]\s*=\s*'([^']+)'", html)
    if not m:
        # Try double quotes
        m = re.search(r'window\[\'_DRIVE_ivd\'\]\s*=\s*"([^"]+)"', html)
    if not m:
        print(f"Could not find _DRIVE_ivd in {html_path}")
        return []

    raw = m.group(1)
    # decode \x22, \x5b, etc.
    decoded_str = raw.encode().decode('unicode-escape')
    
    try:
        data = json.loads(decoded_str)
        return data
    except Exception as e:
        print(f"JSON load error in {html_path}: {e}")
        # Try raw unescape
        clean = re.sub(r'\\x([0-9a-fA-F]{2})', lambda m: chr(int(m.group(1), 16)), raw)
        clean = clean.replace(r'\/', '/')
        try:
            return json.loads(clean)
        except Exception as e2:
            print(f"Second attempt failed: {e2}")
            return []

data_cuc = parse_drive_ivd("drive_cuc.html")
print("Parsed Cô Cúc items count:", len(data_cuc))

# Let's inspect the structure of items in Cô Cúc
with open("cuc_ivd_decoded.json", "w", encoding="utf-8") as f:
    json.dump(data_cuc, f, ensure_ascii=False, indent=2)

print("Saved cuc_ivd_decoded.json")
