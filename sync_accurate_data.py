import json
import shutil
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

# 1. Load exact live crawl
with open(r"D:\APP-KIEM-TRA-GIAO-AN\data\drive_live_exact_report.json", "r", encoding="utf-8") as f:
    live_data = json.load(f)

# Copy to drive_full_cycles_report.json
with open(r"D:\APP-KIEM-TRA-GIAO-AN\data\drive_full_cycles_report.json", "w", encoding="utf-8") as f:
    json.dump(live_data, f, ensure_ascii=False, indent=2)

print("Updated data/drive_full_cycles_report.json successfully!")

# 2. Build accurate drive_monitoring_log.json
teachers_summary = []
now_str = "03/10/2026 15:10"

for t_key, t_val in live_data.items():
    t_teacher = t_val["teacher"]
    cycles = t_val["cycles"]
    
    t1_4 = cycles.get("Tuần 1 đến tuần 4", {})
    t5_8 = cycles.get("Tuần 5 đến tuần 8", {})
    
    t1_4_files = t1_4.get("file_count", 0)
    t5_8_files = t5_8.get("file_count", 0)
    
    total_files = t_val["total_files"]
    last_updated = t_val["last_updated"]
    
    # Determine accurate status and notes
    if t1_4_files > 0 and t5_8_files > 0:
        status = "Hoàn thành Tuần 1-4 & Tuần 5-8"
        notes = f"Đã nộp Tuần 1-4 ({t1_4_files} tệp) và Tuần 5-8 ({t5_8_files} tệp)"
    elif t1_4_files > 0 and t5_8_files == 0:
        status = "Đã nộp Tuần 1-4 (Chưa nộp Tuần 5-8)"
        notes = f"Đã hoàn thành Tuần 1-4 ({t1_4_files} tệp); Thư mục Tuần 5-8 đang trống"
    else:
        status = "CHƯA NỘP (Thư mục trống)"
        notes = "Chưa cập nhật tệp giáo án nào từ Tuần 1 đến Tuần 35"

    teachers_summary.append({
        "id": t_key,
        "teacher_name": t_teacher["name"],
        "folder_name": t_teacher["folder_name"],
        "drive_id": t_teacher["drive_id"],
        "folder_url": f"https://drive.google.com/drive/folders/{t_teacher['drive_id']}?usp=sharing",
        "subjects": t_teacher["subjects"],
        "subfolders": list(cycles.keys()),
        "file_count": total_files,
        "last_updated": last_updated,
        "status": status,
        "notes": notes,
        "week1_4_count": t1_4_files,
        "week5_8_count": t5_8_files
    })

snapshot = {
    "scan_time": now_str,
    "drive_root_url": "https://drive.google.com/drive/folders/1cdqOhxb05lt7r6cyu3YwPVecvBLcmHoR?usp=sharing",
    "total_teachers": len(teachers_summary),
    "completed_count": sum(1 for m in teachers_summary if m["week1_4_count"] > 0),
    "teachers": teachers_summary
}

with open(r"D:\APP-KIEM-TRA-GIAO-AN\data\drive_monitoring_log.json", "w", encoding="utf-8") as f:
    json.dump(snapshot, f, ensure_ascii=False, indent=2)

print("Updated data/drive_monitoring_log.json successfully!")
