import sys
import json
from app.services.inspection_engine import get_teacher_inspected_data, export_teacher_inspection_report_word

sys.stdout.reconfigure(encoding='utf-8')

for tid in ['hang', 'ke', 'thuy', 'linh', 'trang', 'thang']:
    d = get_teacher_inspected_data(tid, 'Tuần 1 đến tuần 4')
    out = export_teacher_inspection_report_word(tid, 'Tuần 1 đến tuần 4')
    print(f"{d['teacher']['name']}: {len(d['lessons'])} bài | Xếp loại: {d['overall_grade']} -> {out}")
