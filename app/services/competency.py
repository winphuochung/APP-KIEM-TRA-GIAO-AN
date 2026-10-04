TEACHERS_DATA = [
    {
        "id": "cuc",
        "name": "Phạm Thị Cúc",
        "title": "Cô Cúc",
        "role": "Tổ phó chuyên môn - GV KHTN",
        "subjects": ["KHTN 7", "Sinh 9"],
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
        "completion_rate": 90,
        "progress_status": "Đúng tiến độ",
        "tii_score": 83,
        "car_score": 86,
        "radar": {"pedagogy": 8.5, "bloom": 8.3, "digital": 7.8, "format": 8.5, "stem": 7.8},
        "sample_files": []
    }
]

def get_heatmap_matrix():
    weeks = [f"T{i}" for i in range(1, 19)]
    matrix = []
    for t in TEACHERS_DATA:
        row = {"id": t["id"], "name": t["title"], "role": t["role"], "cells": []}
        for w_idx in range(1, 19):
            if w_idx <= 4:
                status = "done"  # Đã nộp và kiểm tra
            elif w_idx <= 6:
                status = "early" if t["id"] in ["cuc", "thang"] else "upcoming"  # Nộp đón đầu
            else:
                status = "pending"
            row["cells"].append({"week": f"Tuần {w_idx}", "code": f"T{w_idx}", "status": status})
        matrix.append(row)
    return {"weeks": weeks, "matrix": matrix}

def get_bi_overview():
    return {
        "kpi": {
            "total_teachers": len(TEACHERS_DATA),
            "inspected_rate": "100%",
            "average_tii": 87.0,
            "average_car": 89.5,
            "top_performer": "Thầy Thắng (94) & Cô Cúc (88)"
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
