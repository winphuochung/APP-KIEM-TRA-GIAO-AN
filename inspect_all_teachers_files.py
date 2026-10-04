import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open(r"D:\APP-KIEM-TRA-GIAO-AN\data\drive_live_exact_report.json", "r", encoding="utf-8") as f:
    data = json.load(f)

for t_key, t_val in data.items():
    teacher_name = t_val["teacher"]["name"]
    cycles = t_val["cycles"]
    print(f"\n=======================================================")
    print(f"GIÁO VIÊN: {teacher_name}")
    print(f"=======================================================")
    for c_name, c_info in cycles.items():
        if c_info.get("has_files"):
            print(f"  [{c_name}] - {c_info['file_count']} tệp:")
            subjs = c_info.get("subjects_detail", {})
            for sname, sinfo in subjs.items():
                print(f"    * Môn {sname} ({sinfo['count']} tệp):")
                for fn in sinfo["files"]:
                    print(f"        - {fn}")
