import urllib.request
import re
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

teacher_folders = {
    "cuc": {"name": "Cô Cúc", "id": "1Y57MXplWyQitPENHEQO4McKd7hlLHBtV"},
    "hang": {"name": "Cô Hằng", "id": "1jfXZGd-nH9kgVRsJDg_ldIUFdbR4-MUI"},
    "ke": {"name": "Cô Kế", "id": "10aAFjxR95vKnmLzyvcU7OivbD3GTP88r"},
    "linh": {"name": "Cô Linh", "id": "1slzRqoYLd7pkSfeK3sxDBwzS_5fYA7vd"},
    "thuy": {"name": "Cô Thủy", "id": "1vRyDMAB4lLyxDyPgYsit2VdO7ZcPyuNi"},
    "trang": {"name": "Cô Trang", "id": "1OwIYYK7ENQ4f3aXwS6RrnuA8Mgl7_qKj"},
    "thanh": {"name": "Thầy Thành", "id": "1_hZDldFUOUvjJnkw9-Dm8I3kJbUCc5Hr"},
    "thang": {"name": "Thầy Thắng", "id": "1S1wipWDzCGRaEonbFDFClIzmui1yEXZJ"},
}

for key, info in teacher_folders.items():
    url = f"https://drive.google.com/drive/folders/{info['id']}?usp=sharing"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    try:
        html = urllib.request.urlopen(req, timeout=15).read().decode('utf-8', errors='ignore')
        with open(f"drive_{key}.html", "w", encoding="utf-8") as f:
            f.write(html)
        print(f"Downloaded {info['name']}: {len(html)} bytes")
    except Exception as e:
        print(f"Error {info['name']}: {e}")

