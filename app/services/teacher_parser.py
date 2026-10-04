import os
import re
import docx
import pypdf
import openpyxl

def parse_teachers_from_file(file_path: str) -> list:
    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return parse_excel_teachers(file_path)
    elif ext == ".docx":
        return parse_docx_teachers(file_path)
    elif ext == ".pdf":
        return parse_pdf_teachers(file_path)
    elif ext in [".txt", ".csv"]:
        return parse_txt_csv_teachers(file_path)
    return []


def generate_teacher_id(name_str: str) -> str:
    """Tự động tạo id không dấu ngắn gọn từ Họ tên giáo viên"""
    if not name_str:
        return "gv_new"
    s = str(name_str).lower()
    # Loại bỏ dấu Tiếng Việt
    s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)
    s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)
    s = re.sub(r'[ìíịỉĩ]', 'i', s)
    s = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', s)
    s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)
    s = re.sub(r'[ỳýỵỷỹ]', 'y', s)
    s = re.sub(r'[đ]', 'd', s)
    parts = re.findall(r'[a-z0-9]+', s)
    if not parts:
        return "gv_new"
    if len(parts) == 1:
        return parts[0]
    return parts[-1]  # Lấy tên cuối

def clean_str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()

def parse_excel_teachers(file_path: str) -> list:
    teachers = []
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return teachers

    header_idx = -1
    col_map = {}

    for idx, row in enumerate(rows[:10]):
        row_str = [clean_str(c).lower() for c in row]
        for c_idx, cell in enumerate(row_str):
            if any(k in cell for k in ["họ tên", "họ và tên", "giáo viên", "giao vien", "ho ten"]):
                header_idx = idx
                break
        if header_idx != -1:
            break

    if header_idx != -1:
        header_row = [clean_str(c).lower() for c in rows[header_idx]]
        for c_idx, h in enumerate(header_row):
            if "họ" in h or "tên" in h or "giáo viên" in h:
                col_map["name"] = c_idx
            elif "chức vụ" in h or "nhiệm vụ" in h or "chuc vu" in h:
                col_map["role"] = c_idx
            elif "môn" in h or "bộ môn" in h or "phân công" in h:
                col_map["subjects"] = c_idx
            elif "khối" in h or "khoi" in h:
                col_map["grades"] = c_idx
            elif "lớp" in h or "dạy" in h or "lop" in h:
                col_map["classes"] = c_idx
            elif "chủ nhiệm" in h or "chu nhiem" in h:
                col_map["homeroom"] = c_idx
    else:
        # Giả định cột mặc định: 0: STT, 1: Họ tên, 2: Chức vụ, 3: Môn, 4: Lớp
        col_map = {"name": 1, "role": 2, "subjects": 3, "classes": 4}

    start_row = header_idx + 1 if header_idx != -1 else 0
    for row in rows[start_row:]:
        if not row:
            continue
        name_idx = col_map.get("name", 0 if len(row) > 0 else -1)
        if name_idx == -1 or name_idx >= len(row):
            continue
            
        name_val = clean_str(row[name_idx])
        if not name_val or name_val.lower() in ["stt", "họ tên", "họ và tên", "tổng cộng"]:
            continue
            
        role_val = clean_str(row[col_map["role"]]) if "role" in col_map and col_map["role"] < len(row) else "Giáo viên"
        subj_val = clean_str(row[col_map["subjects"]]) if "subjects" in col_map and col_map["subjects"] < len(row) else "Chuyên môn"
        grade_val = clean_str(row[col_map["grades"]]) if "grades" in col_map and col_map["grades"] < len(row) else ""
        class_val = clean_str(row[col_map["classes"]]) if "classes" in col_map and col_map["classes"] < len(row) else ""
        home_val = clean_str(row[col_map["homeroom"]]) if "homeroom" in col_map and col_map["homeroom"] < len(row) else ""

        classes_list = [c.strip() for c in re.split(r'[,;\s]+', class_val) if c.strip()]
        subjs_list = [s.strip() for s in re.split(r'[,;]+', subj_val) if s.strip()]

        teachers.append({
            "id": generate_teacher_id(name_val),
            "name": name_val,
            "role": role_val or "Giáo viên",
            "subjects": subjs_list or [subj_val],
            "subject_str": subj_val or "Chuyên môn",
            "grades": grade_val or "6, 7, 8, 9",
            "classes": classes_list or ["Lớp"],
            "homeroom": home_val,
            "completion_rate": 100,
            "progress_status": "Đúng tiến độ",
            "tii_score": 88,
            "car_score": 90,
            "radar": {"pedagogy": 9.0, "bloom": 8.8, "digital": 8.5, "format": 8.8, "stem": 8.5}
        })

    return teachers

def parse_docx_teachers(file_path: str) -> list:
    teachers = []
    doc = docx.Document(file_path)

    # Đọc bảng biểu nếu có
    for table in doc.tables:
        rows_data = []
        for row in table.rows:
            rows_data.append([c.text.strip() for c in row.cells])
        if len(rows_data) > 1:
            header = [c.lower() for c in rows_data[0]]
            name_col = -1
            role_col = -1
            subj_col = -1
            for idx, h in enumerate(header):
                if "họ tên" in h or "tên" in h or "giáo viên" in h:
                    name_col = idx
                elif "chức vụ" in h or "nhiệm vụ" in h:
                    role_col = idx
                elif "môn" in h or "dạy" in h:
                    subj_col = idx
            
            if name_col == -1 and len(rows_data[0]) > 1:
                name_col = 1 if len(rows_data[0]) > 1 else 0

            for r in rows_data[1:]:
                if name_col < len(r) and r[name_col]:
                    n_val = r[name_col]
                    if n_val.lower() in ["stt", "họ và tên", "họ tên"]:
                        continue
                    r_val = r[role_col] if role_col != -1 and role_col < len(r) else "Giáo viên"
                    s_val = r[subj_col] if subj_col != -1 and subj_col < len(r) else "Chuyên môn"
                    teachers.append({
                        "id": generate_teacher_id(n_val),
                        "name": n_val,
                        "role": r_val or "Giáo viên",
                        "subjects": [s_val] if s_val else ["Chuyên môn"],
                        "subject_str": s_val or "Chuyên môn",
                        "grades": "6, 7, 8, 9",
                        "classes": ["Lớp"],
                        "homeroom": "",
                        "completion_rate": 100,
                        "progress_status": "Đúng tiến độ",
                        "tii_score": 88,
                        "car_score": 90,
                        "radar": {"pedagogy": 9.0, "bloom": 8.8, "digital": 8.5, "format": 8.8, "stem": 8.5}
                    })

    # Nếu không tìm thấy bảng, trích xuất dòng văn bản
    if not teachers:
        for p in doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                continue
            # Tìm dòng dạng: 1. Nguyễn Văn A - Tổ trưởng - Môn Toán
            m = re.match(r'^(?:\d+[\.\s\-]+)?([A-ZÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ][a-zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ\s]+)(?:[\-\–:]+([^\-\–]+))?(?:[\-\–:]+(.+))?$', txt)
            if m:
                n_val = m.group(1).strip()
                r_val = m.group(2).strip() if m.group(2) else "Giáo viên"
                s_val = m.group(3).strip() if m.group(3) else "Chuyên môn"
                if len(n_val) >= 3 and not n_val.lower().startswith("trường"):
                    teachers.append({
                        "id": generate_teacher_id(n_val),
                        "name": n_val,
                        "role": r_val or "Giáo viên",
                        "subjects": [s_val],
                        "subject_str": s_val,
                        "grades": "6, 7, 8, 9",
                        "classes": ["Lớp"],
                        "homeroom": "",
                        "completion_rate": 100,
                        "progress_status": "Đúng tiến độ",
                        "tii_score": 88,
                        "car_score": 90,
                        "radar": {"pedagogy": 9.0, "bloom": 8.8, "digital": 8.5, "format": 8.8, "stem": 8.5}
                    })
    return teachers

def parse_pdf_teachers(file_path: str) -> list:
    teachers = []
    reader = pypdf.PdfReader(file_path)
    full_text = ""
    for page in reader.pages:
        full_text += page.extract_text() + "\n"

    lines = [l.strip() for l in full_text.splitlines() if l.strip()]
    for l in lines:
        m = re.match(r'^(?:\d+[\.\s\-]+)?([A-ZÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ][a-zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ\s]{3,30})(?:[\-\–:]+([^\-\–]+))?', l)
        if m:
            n_val = m.group(1).strip()
            r_val = m.group(2).strip() if m.group(2) else "Giáo viên"
            if len(n_val) >= 3 and not any(k in n_val.lower() for k in ["trường", "báo cáo", "kế hoạch", "cộng hòa"]):
                teachers.append({
                    "id": generate_teacher_id(n_val),
                    "name": n_val,
                    "role": r_val,
                    "subjects": ["Chuyên môn"],
                    "subject_str": "Chuyên môn",
                    "grades": "6, 7, 8, 9",
                    "classes": ["Lớp"],
                    "homeroom": "",
                    "completion_rate": 100,
                    "progress_status": "Đúng tiến độ",
                    "tii_score": 88,
                    "car_score": 90,
                    "radar": {"pedagogy": 9.0, "bloom": 8.8, "digital": 8.5, "format": 8.8, "stem": 8.5}
                })
    return teachers

def parse_txt_csv_teachers(file_path: str) -> list:
    teachers = []
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = [p.strip() for p in re.split(r'[,;\t|]+', line) if p.strip()]
            if parts:
                n_val = parts[0]
                if n_val.lower() in ["stt", "họ tên", "name"]:
                    continue
                r_val = parts[1] if len(parts) > 1 else "Giáo viên"
                s_val = parts[2] if len(parts) > 2 else "Chuyên môn"
                teachers.append({
                    "id": generate_teacher_id(n_val),
                    "name": n_val,
                    "role": r_val,
                    "subjects": [s_val],
                    "subject_str": s_val,
                    "grades": "6, 7, 8, 9",
                    "classes": ["Lớp"],
                    "homeroom": "",
                    "completion_rate": 100,
                    "progress_status": "Đúng tiến độ",
                    "tii_score": 88,
                    "car_score": 90,
                    "radar": {"pedagogy": 9.0, "bloom": 8.8, "digital": 8.5, "format": 8.8, "stem": 8.5}
                })
    return teachers
