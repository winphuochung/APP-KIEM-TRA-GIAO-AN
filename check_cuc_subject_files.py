import urllib.request
import re
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

from test_subfolder_contents import fetch_and_decode

print("=== Checking Cô Cúc Tuần 1-4 Subject Folders ===")
khtn7_t1_4 = fetch_and_decode("10ypc7wavrO_CTWHjScXj63hPvmIrwYQF")
print(f"KHTN 7 Tuần 1-4 files: {len(khtn7_t1_4)}")
for f in khtn7_t1_4:
    print(f"  {f['name']} ({f['type']})")

sinh9_t1_4 = fetch_and_decode("1GKCjshu58RQ3-l1adaZzbprYLsqrHak2")
print(f"Sinh 9 Tuần 1-4 files: {len(sinh9_t1_4)}")
for f in sinh9_t1_4:
    print(f"  {f['name']} ({f['type']})")

print("\n=== Checking Cô Cúc Tuần 5-8 Subject Folders ===")
khtn7_t5_8 = fetch_and_decode("1n2Jf-gCS45cg84QQ2DEwz7T4s5zwPHWy")
print(f"KHTN 7 Tuần 5-8 files: {len(khtn7_t5_8)}")
for f in khtn7_t5_8:
    print(f"  {f['name']} ({f['type']})")

sinh9_t5_8 = fetch_and_decode("1RrT4O1oxGfcBBUTJSwAQwruxcrv6dVNg")
print(f"Sinh 9 Tuần 5-8 files: {len(sinh9_t5_8)}")
for f in sinh9_t5_8:
    print(f"  {f['name']} ({f['type']})")
