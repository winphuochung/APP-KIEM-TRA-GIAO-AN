import re
import json

with open("scratch_drive_root.html", "r", encoding="utf-8") as f:
    html = f.read()

out_lines = []
out_lines.append("=== ALL ARIA LABELS ===")
labels = re.findall(r'aria-label="([^"]+)"', html)
for l in labels:
    out_lines.append(l)

out_lines.append("\n=== TEACHER MATCHES ===")
teachers = ["Cúc", "Hằng", "Kế", "Linh", "Thủy", "Trang", "Thành", "Thắng"]
for t in teachers:
    matches = [m.start() for m in re.finditer(t, html, re.IGNORECASE)]
    out_lines.append(f"Teacher {t}: {len(matches)} occurrences")
    for pos in matches[:3]:
        snippet = html[max(0, pos-100):min(len(html), pos+100)]
        out_lines.append(f"  Snippet: {snippet.strip()}")

# Search for any json arrays containing folder ids
folder_matches = re.findall(r'\["([a-zA-Z0-9_-]{25,})",\["([^"]+)"\]', html)
out_lines.append(f"\n=== FOLDER MATCHES ({len(folder_matches)}) ===")
for fid, fname in folder_matches:
    out_lines.append(f"{fid} -> {fname}")

with open("drive_parsed_utf8.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out_lines))

print("Done writing drive_parsed_utf8.txt")
