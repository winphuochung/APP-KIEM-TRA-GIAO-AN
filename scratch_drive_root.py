import urllib.request
import re
import json

root_url = "https://drive.google.com/drive/folders/1cdqOhxb05lt7r6cyu3YwPVecvBLcmHoR?usp=sharing"
req = urllib.request.Request(root_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
html = urllib.request.urlopen(req).read().decode('utf-8')

with open("scratch_drive_root.html", "w", encoding="utf-8") as f:
    f.write(html)

print("Saved root html, length:", len(html))
