import re

with open("drive_cuc.html", "r", encoding="utf-8") as f:
    html = f.read()

out = []
matches = [m.start() for m in re.finditer(r"Tuần 1 đến tuần 4", html)]
out.append(f"Matches for Tuần 1 đến tuần 4: {len(matches)}")
for idx, pos in enumerate(matches):
    out.append(f"--- Match {idx} ---")
    out.append(html[max(0, pos-200):min(len(html), pos+300)])

matches5 = [m.start() for m in re.finditer(r"Tuần 5 đến tuần 8", html)]
out.append(f"\nMatches for Tuần 5 đến tuần 8: {len(matches5)}")
for idx, pos in enumerate(matches5):
    out.append(f"--- Match {idx} ---")
    out.append(html[max(0, pos-200):min(len(html), pos+300)])

# Also search for ssk attribute or data-id or drive id regex
# E.g. ssk='5:auSv138:1Y57MXplWyQitPENHEQO4McKd7hlLHBtV'
subfolder_ssks = re.findall(r'ssk=[\'"]5:auSv138:([^\'"]+)[\'"]', html)
out.append(f"\nSubfolder SSKs found ({len(subfolder_ssks)}):")
for s in subfolder_ssks:
    out.append(s)

with open("subfolder_ids_utf8.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))

print("Done writing subfolder_ids_utf8.txt")
