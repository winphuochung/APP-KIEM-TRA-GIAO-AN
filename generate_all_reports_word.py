import sys
from app.services.inspection_engine import export_teacher_inspection_report_word, TEACHERS_INFO

sys.stdout.reconfigure(encoding='utf-8')

for tid in TEACHERS_INFO.keys():
    # Tuần 1 đến tuần 4
    out1 = export_teacher_inspection_report_word(tid, "Tuần 1 đến tuần 4")
    print(f"[Tuần 1-4] {TEACHERS_INFO[tid]['name']} -> {out1}")
    
    # Tuần 5 đến tuần 8
    out2 = export_teacher_inspection_report_word(tid, "Tuần 5 đến tuần 8")
    print(f"[Tuần 5-8] {TEACHERS_INFO[tid]['name']} -> {out2}")

print("\nĐã xuất toàn bộ Biên bản kiểm tra gộp các môn của tất cả giáo viên!")
