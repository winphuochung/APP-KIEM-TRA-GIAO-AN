import os
import json
from app.config import DATA_DIR

TEACHERS_FILE = os.path.join(DATA_DIR, "teachers.json")

DEFAULT_TEACHERS_DATA = [
    {
        "id": "cuc",
        "name": "Phạm Thị Cúc",
        "title": "Cô Cúc",
        "role": "Tổ phó chuyên môn - GV KHTN",
        "subjects": ["KHTN 7", "Sinh 9"],
        "classes": ["7A2", "9A2"],
        "homeroom": "9A2",
        "completion_rate": 100,
        "progress_status": "Đúng tiến độ",
        "tii_score": 88,
        "car_score": 92,
        "radar": {"pedagogy": 9.2, "bloom": 9.4, "digital": 8.8, "format": 7.5, "stem": 8.0},
        "sample_files": [
            "Tuần 1-2 Tiết 1-5  BÀI 1 MỞ ĐẦU KHTN.docx",
            "Tuần 2-3  Tiết 6-9  Bài 2  Nguyên tử - KHTN7 - CTST.docx",
            "Tuần 3-4, Tiết 10-13  -CHỦ ĐỀ 1 - BÀI 3- NTHH.docx",
            "Tuần 4-5  -Tiết 14-18  Bài 4- SƠ LƯỢC BẢNG TUẦN TOÀN CÁC NTHH-KHTN 7-CTST-ST.docx",
            "Tuần 2 - Tiết 1-2   Bài 36-Khái quát về di truyền học-Sinh9- KNTT.docx",
            "Tuần 3,4,5,6 - Tiết 3,4,5,6  BÀI 37- khtn 9- kntt.docx"
        ]
    },
    {
        "id": "thang",
        "name": "Lê Văn Thắng",
        "title": "Thầy Thắng",
        "role": "Tổ trưởng chuyên môn - GV KHTN",
        "subjects": ["Hóa 9", "KHTN 7"],
        "classes": ["7A1", "9A1"],
        "homeroom": "9A1",
        "completion_rate": 100,
        "progress_status": "Đúng tiến độ",
        "tii_score": 94,
        "car_score": 96,
        "radar": {"pedagogy": 9.5, "bloom": 9.2, "digital": 9.6, "format": 9.5, "stem": 9.0},
        "sample_files": []
    },
    {
        "id": "linh",
        "name": "Phạm Thị Mỹ Linh",
        "title": "Cô Linh",
        "role": "Giáo viên Toán",
        "subjects": ["Toán 8", "Toán 9"],
        "classes": ["8A1", "9A1"],
        "homeroom": "",
        "completion_rate": 100,
        "progress_status": "Đúng tiến độ",
        "tii_score": 89,
        "car_score": 90,
        "radar": {"pedagogy": 9.0, "bloom": 9.0, "digital": 8.5, "format": 9.2, "stem": 8.4},
        "sample_files": []
    },
    {
        "id": "thuy",
        "name": "Lê Thị Thu Thủy",
        "title": "Cô Thủy",
        "role": "Giáo viên Toán",
        "subjects": ["Toán 6", "Toán 7", "Toán 8"],
        "classes": ["6A1", "7A1", "8A2"],
        "homeroom": "7A1",
        "completion_rate": 95,
        "progress_status": "Đúng tiến độ",
        "tii_score": 85,
        "car_score": 88,
        "radar": {"pedagogy": 8.8, "bloom": 8.6, "digital": 8.0, "format": 8.8, "stem": 7.9},
        "sample_files": []
    },
    {
        "id": "thanh",
        "name": "Nguyễn Chí Thành",
        "title": "Thầy Thành",
        "role": "Giáo viên KHTN",
        "subjects": ["HĐTN 8", "KHTN 6", "KHTN 8", "KHTN 9"],
        "classes": ["6A2", "8A1", "9A2"],
        "homeroom": "8A2",
        "completion_rate": 92,
        "progress_status": "Đúng tiến độ",
        "tii_score": 84,
        "car_score": 87,
        "radar": {"pedagogy": 8.7, "bloom": 8.5, "digital": 8.2, "format": 8.4, "stem": 8.2},
        "sample_files": []
    },
    {
        "id": "trang",
        "name": "Nguyễn Thị Thùy Trang",
        "title": "Cô Trang",
        "role": "Giáo viên Công nghệ",
        "subjects": ["CN 8", "CN 9"],
        "classes": ["8A1", "8A2", "9A1", "9A2"],
        "homeroom": "",
        "completion_rate": 94,
        "progress_status": "Đúng tiến độ",
        "tii_score": 86,
        "car_score": 89,
        "radar": {"pedagogy": 8.9, "bloom": 8.4, "digital": 8.5, "format": 8.7, "stem": 8.8},
        "sample_files": []
    },
    {
        "id": "hang",
        "name": "Lê Thị Thúy Hằng",
        "title": "Cô Hằng",
        "role": "Giáo viên Công nghệ & HĐTN",
        "subjects": ["CN 6", "CN 7", "HĐTN 6", "HĐTN 7", "HĐTN 9"],
        "classes": ["6A1", "6A2", "7A1", "7A2"],
        "homeroom": "6A1",
        "completion_rate": 90,
        "progress_status": "Đúng tiến độ",
        "tii_score": 82,
        "car_score": 85,
        "radar": {"pedagogy": 8.4, "bloom": 8.2, "digital": 8.0, "format": 8.3, "stem": 8.0},
        "sample_files": []
    },
    {
        "id": "ke",
        "name": "Hà Thị Kế",
        "title": "Cô Kế",
        "role": "Giáo viên KHTN & HĐTN",
        "subjects": ["HĐTN 8", "KHTN 6", "KHTN 8"],
        "classes": ["6A1", "8A1"],
        "homeroom": "6A1",
        "completion_rate": 90,
        "progress_status": "Đúng tiến độ",
        "tii_score": 83,
        "car_score": 86,
        "radar": {"pedagogy": 8.5, "bloom": 8.3, "digital": 7.8, "format": 8.5, "stem": 7.8},
        "sample_files": []
    }
]

def load_teachers_data():
    if os.path.exists(TEACHERS_FILE):
        try:
            with open(TEACHERS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list) and len(data) > 0:
                    return data
        except Exception:
            pass
    
    # Save default if not existing
    save_teachers_list_to_file(DEFAULT_TEACHERS_DATA)
    return [dict(t) for t in DEFAULT_TEACHERS_DATA]

def save_teachers_list_to_file(data_list):
    try:
        os.makedirs(os.path.dirname(TEACHERS_FILE), exist_ok=True)
        with open(TEACHERS_FILE, "w", encoding="utf-8") as f:
            json.dump(data_list, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

TEACHERS_DATA = load_teachers_data()

def save_teachers_data():
    global TEACHERS_DATA
    save_teachers_list_to_file(TEACHERS_DATA)

def get_heatmap_matrix():
    weeks = [f"T{i}" for i in range(1, 19)]
    matrix = []
    
    nb_path = os.path.join(DATA_DIR, "so_tay_to_truong.json")
    notebook_data = {}
    if os.path.exists(nb_path):
        try:
            with open(nb_path, "r", encoding="utf-8") as f:
                notebook_data = json.load(f)
        except Exception:
            pass

    for t in TEACHERS_DATA:
        t_id = t["id"]
        row = {"id": t_id, "name": t.get("title") or t.get("name"), "role": t.get("role", "Giáo viên"), "cells": []}
        
        for w_idx in range(1, 19):
            w_str = str(w_idx)
            rec = None
            if w_str in notebook_data:
                for r in notebook_data[w_str]:
                    if r.get("teacher_id") == t_id:
                        rec = r
                        break
            
            if rec:
                g_drive = str(rec.get("giao_an_drive", "")).lower()
                if "chưa" in g_drive or "trễ" in g_drive or "thiếu" in g_drive:
                    status = "late"
                elif w_idx <= 4:
                    status = "done"
                elif w_idx <= 6:
                    status = "early"
                else:
                    status = "done"
            else:
                if w_idx <= 4:
                    status = "done"
                elif w_idx <= 6:
                    status = "early" if t_id in ["cuc", "thang"] else "upcoming"
                else:
                    status = "pending"
                    
            row["cells"].append({"week": f"Tuần {w_idx}", "code": f"T{w_idx}", "status": status})
        matrix.append(row)
    return {"weeks": weeks, "matrix": matrix}

def get_bi_overview():
    avg_tii = round(sum(t.get("tii_score", 85) for t in TEACHERS_DATA) / max(len(TEACHERS_DATA), 1), 1)
    avg_car = round(sum(t.get("car_score", 88) for t in TEACHERS_DATA) / max(len(TEACHERS_DATA), 1), 1)
    
    sorted_teachers = sorted(TEACHERS_DATA, key=lambda x: x.get("tii_score", 0), reverse=True)
    top_str = f"{sorted_teachers[0].get('title', sorted_teachers[0].get('name'))} ({sorted_teachers[0].get('tii_score')})" if len(sorted_teachers) > 0 else "Chưa có"
    if len(sorted_teachers) > 1:
        top_str += f" & {sorted_teachers[1].get('title', sorted_teachers[1].get('name'))} ({sorted_teachers[1].get('tii_score')})"
    
    return {
        "kpi": {
            "total_teachers": len(TEACHERS_DATA),
            "inspected_rate": "100%",
            "average_tii": avg_tii,
            "average_car": avg_car,
            "top_performer": top_str
        },
        "error_distribution": {
            "labels": ["Căn lề chưa chuẩn NĐ 30", "Lỗi font chữ / size", "Lỗi chính tả / dính chữ", "Thiếu bước 4 (5512)", "Phần Vận dụng sơ sài"],
            "data": [45, 15, 20, 12, 8]
        },
        "grade_quality": {
            "labels": ["Khối 6", "Khối 7", "Khối 8", "Khối 9"],
            "scores": [8.4, 8.6, 8.8, 9.2]
        }
    }

def add_teacher_data(teacher_obj: dict):
    global TEACHERS_DATA
    from app.services.teacher_parser import generate_teacher_id
    t_id = teacher_obj.get("id") or generate_teacher_id(teacher_obj.get("name", ""))
    existing_ids = [t["id"] for t in TEACHERS_DATA]
    if t_id in existing_ids:
        t_id = f"{t_id}_{len(existing_ids)+1}"
    
    teacher_obj["id"] = t_id
    if "title" not in teacher_obj or not teacher_obj["title"]:
        teacher_obj["title"] = f"{teacher_obj.get('name', '')}"
    if "completion_rate" not in teacher_obj:
        teacher_obj["completion_rate"] = 100
    if "progress_status" not in teacher_obj:
        teacher_obj["progress_status"] = "Đúng tiến độ"
    if "tii_score" not in teacher_obj:
        teacher_obj["tii_score"] = 88
    if "car_score" not in teacher_obj:
        teacher_obj["car_score"] = 90
    if "radar" not in teacher_obj:
        teacher_obj["radar"] = {"pedagogy": 9.0, "bloom": 8.8, "digital": 8.5, "format": 8.8, "stem": 8.5}

    if "subjects" not in teacher_obj or not teacher_obj["subjects"]:
        teacher_obj["subjects"] = [teacher_obj.get("subject_str", "Chuyên môn")]
    if "classes" not in teacher_obj:
        teacher_obj["classes"] = []
    if "homeroom" not in teacher_obj:
        teacher_obj["homeroom"] = ""

    TEACHERS_DATA.append(teacher_obj)
    save_teachers_data()
    return teacher_obj

def update_teacher_data(teacher_id: str, teacher_obj: dict):
    global TEACHERS_DATA
    for idx, t in enumerate(TEACHERS_DATA):
        if t["id"] == teacher_id:
            if "name" in teacher_obj and teacher_obj["name"]:
                t["name"] = teacher_obj["name"]
                t["title"] = teacher_obj["name"]
            if "role" in teacher_obj:
                t["role"] = teacher_obj["role"]
            if "subjects" in teacher_obj:
                t["subjects"] = teacher_obj["subjects"] if isinstance(teacher_obj["subjects"], list) else [teacher_obj["subjects"]]
            if "classes" in teacher_obj:
                t["classes"] = teacher_obj["classes"] if isinstance(teacher_obj["classes"], list) else [teacher_obj["classes"]]
            if "homeroom" in teacher_obj:
                t["homeroom"] = teacher_obj["homeroom"] or ""
            
            TEACHERS_DATA[idx] = t
            save_teachers_data()
            return t
    return None

def delete_teacher_data(teacher_id: str):
    global TEACHERS_DATA
    initial_count = len(TEACHERS_DATA)
    TEACHERS_DATA[:] = [t for t in TEACHERS_DATA if t["id"] != teacher_id]
    if len(TEACHERS_DATA) < initial_count:
        save_teachers_data()
        return True
    return False
