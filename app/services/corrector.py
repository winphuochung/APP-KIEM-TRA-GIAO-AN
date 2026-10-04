import os
import docx
from docx.shared import Mm, Pt

SPELLING_CORRECTIONS = {
    "BẢNG TUẦN TOÀN": "BẢNG TUẦN HOÀN",
    "Bảng tuần toàn": "Bảng tuần hoàn",
    "NGUYỀN TỐ": "NGUYÊN TỐ",
    "nguyền tố": "nguyên tố",
    "lí thuyêt": "lí thuyết",
    "gsk": "sgk",
    "Tuần 45": "Tuần 4-5",
    "Tiết 1418": "Tiết 14-18",
    "Tiết 15": "Tiết 1-5",
    "Tuần 23": "Tuần 2-3",
    "Tiết 36": "Tiết 6-9",
    "Tiết 1013": "Tiết 10-13",
    "Tuần 34": "Tuần 3-4"
}

def correct_document(input_path, output_path):
    doc = docx.Document(input_path)
    
    # Căn lề chuẩn Nghị định 30/2020/NĐ-CP: Top=20mm, Bottom=20mm, Left=30mm, Right=15mm
    for section in doc.sections:
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)
        section.page_width = Mm(210)
        section.page_height = Mm(297)

    fixed_count = 0
    # Chuẩn hóa phông chữ và sửa chính tả trong paragraphs
    for p in doc.paragraphs:
        for wrong, right in SPELLING_CORRECTIONS.items():
            if wrong in p.text:
                for run in p.runs:
                    if wrong in run.text:
                        run.text = run.text.replace(wrong, right)
                        fixed_count += 1
        for run in p.runs:
            run.font.name = "Times New Roman"
            if run.font.size and run.font.size.pt not in [11.0, 12.0, 13.0, 14.0, 15.0, 16.0]:
                run.font.size = Pt(13)

    # Chuẩn hóa trong bảng
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    for wrong, right in SPELLING_CORRECTIONS.items():
                        if wrong in p.text:
                            for run in p.runs:
                                if wrong in run.text:
                                    run.text = run.text.replace(wrong, right)
                                    fixed_count += 1
                    for run in p.runs:
                        run.font.name = "Times New Roman"
                        if run.font.size and run.font.size.pt not in [10.0, 11.0, 12.0, 13.0, 14.0]:
                            run.font.size = Pt(12)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return {
        "status": "success",
        "output_path": output_path,
        "margins": {"top": 20, "bottom": 20, "left": 30, "right": 15},
        "fixed_typos_count": fixed_count
    }
