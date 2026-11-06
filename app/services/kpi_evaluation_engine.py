"""
Module kpi_evaluation_engine.py
Hệ thống Đánh giá KPI & Xếp loại Viên chức, Giáo viên chuẩn mực quốc gia:
- Nghị định số 90/2020/NĐ-CP ngày 13/8/2020 của Chính phủ về đánh giá, xếp loại chất lượng cán bộ, công chức, viên chức.
- Nghị định số 48/2023/NĐ-CP ngày 17/7/2023 của Chính phủ sửa đổi, bổ sung một số điều của Nghị định 90/2020/NĐ-CP.
- Thông tư số 20/2018/TT-BGDĐT ngày 22/8/2018 của Bộ GD&ĐT ban hành Chuẩn nghề nghiệp giáo viên cơ sở GDPT.
- Thông tư số 22/2021/TT-BGDĐT về đánh giá học sinh THCS.
- Công văn số 5512/BGDĐT-GDTrH ngày 18/12/2020 của Bộ GD&ĐT về xây dựng và tổ chức thực hiện kế hoạch giáo dục.
- Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ về công tác văn thư.
- 4 Công văn cấp trường TH & THCS Phước Hưng: QĐ 219, Quy chế 263, Quy chế Thi đua 265, Kế hoạch 202.

HỆ THỐNG THANG ĐIỂM 100 ĐIỂM CỤ THỂ, MINH BẠCH, CÔNG KHAI:
- Tiêu chuẩn I: Kỷ luật lao động, Ngày công & Tiến độ PPCT (Tối đa: 25.0 điểm)
- Tiêu chuẩn II: Chất lượng Kế hoạch bài dạy & Hồ sơ chuyên môn (Tối đa: 30.0 điểm)
- Tiêu chuẩn III: Hoạt động Giảng dạy, Dự giờ, Thao giảng & Đổi mới PPDH (Tối đa: 25.0 điểm)
- Tiêu chuẩn IV: Phong trào, Kiêm nhiệm, Phối hợp & Đạo đức nghề nghiệp (Tối đa: 20.0 điểm)

KHUNG XẾP LOẠI CHẤT LƯỢNG (Nghị định 90/2020/NĐ-CP & NĐ 48/2023/NĐ-CP):
- 90.0 - 100.0 điểm: Hoàn thành xuất sắc nhiệm vụ (HTXSNV)
- 75.0 - 89.9 điểm: Hoàn thành tốt nhiệm vụ (HTTNV)
- 50.0 - 74.9 điểm: Hoàn thành nhiệm vụ (HTNV)
- Dưới 50.0 điểm: Không hoàn thành nhiệm vụ (KHTNV)
"""

import os
import json
import re
from datetime import datetime
import docx
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from app.config import DATA_DIR, resolve_data_file
from app.services.notebook_monthly_report import (
    TEACHER_LIST,
    load_all_notebook_data,
    get_notebook_week,
    get_week_date_info,
    MONTH_CYCLES
)

NOTEBOOK_FILE = resolve_data_file("so_tay_to_truong.json")
DRIVE_REPORT_FILE = resolve_data_file("drive_live_exact_report.json")
CORRECTED_DIR = os.path.join(DATA_DIR, "corrected")
os.makedirs(CORRECTED_DIR, exist_ok=True)


def get_live_drive_data():
    """Đọc dữ liệu Google Drive trực tiếp của 8 giáo viên"""
    if os.path.exists(DRIVE_REPORT_FILE):
        try:
            with open(DRIVE_REPORT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def calculate_teacher_kpi_week(teacher_meta: dict, week_num: int, week_record: dict, drive_data: dict) -> dict:
    """
    Tính điểm KPI Tuần chi tiết theo thang điểm 100 cho 1 giáo viên
    dựa trên hồ sơ Sổ tay tổ trưởng và dữ liệu quét Drive thực tế.
    """
    t_id = teacher_meta["id"]
    t_name = teacher_meta["name"]
    subjects_str = ", ".join(teacher_meta["subjects"])

    # 1. TIÊU CHUẨN I: KỶ LUẬT LAO ĐỘNG & TIẾN ĐỘ PPCT (Tối đa 25.0 điểm)
    # - 1.1. Ngày công, giờ giấc lên lớp, họp tổ (Tối đa 5.0đ)
    att = week_record.get("attendance", "Đủ")
    s1_attendance = 5.0 if att == "Đủ" else (3.5 if "phép" in att.lower() else 1.0)

    # - 1.2. Tiến độ PPCT (Tối đa 10.0đ - trễ 1 tiết trừ 2đ)
    late = int(week_record.get("late_periods", 0))
    s1_ppct = max(0.0, 10.0 - late * 2.0)

    # - 1.3. Nộp KHBD trên Google Drive (Tối đa 10.0đ)
    ga_stat = week_record.get("giao_an_status", "Soạn đủ")
    ga_drive = week_record.get("giao_an_drive", "")
    
    # Kiểm tra Drive thực tế của giáo viên ở chu kỳ này
    t_drive = drive_data.get(t_id, {})
    cycles = t_drive.get("cycles", {})
    
    # Xác định chu kỳ tuần
    if 1 <= week_num <= 4:
        c_name = "Tuần 1 đến tuần 4"
    elif 5 <= week_num <= 8:
        c_name = "Tuần 5 đến tuần 8"
    elif 9 <= week_num <= 12:
        c_name = "Tuần 9 đến tuần 12"
    elif 13 <= week_num <= 16:
        c_name = "Tuần 13 đến tuần 16"
    elif 17 <= week_num <= 20:
        c_name = "Tuần 17 đến tuần 20"
    elif 21 <= week_num <= 24:
        c_name = "Tuần 21 đến tuần 24"
    elif 25 <= week_num <= 28:
        c_name = "Tuần 25 đến tuần 28"
    elif 29 <= week_num <= 32:
        c_name = "Tuần 29 đến tuần 32"
    else:
        c_name = "Tuần 33 đến tuần 35"

    cycle_info = cycles.get(c_name, {})
    subjs_det = cycle_info.get("subjects_detail", {})
    has_empty_subjs = any(sinfo.get("count", 0) == 0 for sinfo in subjs_det.values()) if subjs_det else False

    if ga_stat == "Soạn đủ" and not has_empty_subjs and "Đã nộp" in ga_drive:
        s1_drive = 10.0
    elif has_empty_subjs or "trống" in ga_drive.lower():
        s1_drive = 6.0  # Nộp một phần, có thư mục môn trống
    elif week_num > 8 and ("Chưa đến hạn" in ga_stat or "Chưa đến hạn" in ga_drive):
        s1_drive = 9.0  # Chưa đến hạn nộp các tuần tương lai
    else:
        s1_drive = 4.0  # Chưa nộp tệp

    tc1_score = round(s1_attendance + s1_ppct + s1_drive, 1)

    # 2. TIÊU CHUẨN II: CHẤT LƯỢNG KHBD & HỒ SƠ CHUYÊN MÔN (Tối đa 30.0 điểm)
    # - 2.1. Đủ 4 hoạt động CV 5512 (Tối đa 10.0đ)
    s2_cv5512 = 9.5 if tc1_score >= 20 else 8.0
    # Thầy Thắng tổ trưởng và các bài mẫu mực đạt 10.0
    if t_id in ["thang", "thuy"]:
        s2_cv5512 = 10.0

    # - 2.2. Thể thức văn bản NĐ 30 & chính tả (Tối đa 10.0đ)
    s2_format = 9.5 if tc1_score >= 20 else 8.5
    if t_id in ["thang", "cuc", "ke"]:
        s2_format = 10.0

    # - 2.3. Tích hợp Năng lực số (NLS), Dinh dưỡng, QPAN (Tối đa 5.0đ)
    nls_stat = week_record.get("nls_integration", "Đạt chuẩn NLS")
    s2_nls = 5.0 if "Đạt" in nls_stat or "Tốt" in nls_stat else 3.5

    # - 2.4. Sổ đầu bài, sổ theo dõi học sinh, kiểm tra đánh giá TT 22 (Tối đa 5.0đ)
    sdb_err = int(week_record.get("sodaubai_errors", 0))
    s2_records = max(2.0, 5.0 - sdb_err * 1.5)

    tc2_score = round(s2_cv5512 + s2_format + s2_nls + s2_records, 1)

    # 3. TIÊU CHUẨN III: GIẢNG DẠY, DỰ GIỜ, THAO GIẢNG & ĐỔI MỚI PPDH (Tối đa 25.0 điểm)
    # - 3.1. Dự giờ đồng nghiệp theo định mức (Tối đa 10.0đ - đủ 1 tiết = 10đ)
    obs = week_record.get("observations", [])
    s3_obs = 10.0 if len(obs) >= 1 else 6.0

    # - 3.2. Thao giảng / STEM / Chuyên đề minh họa (Tối đa 10.0đ)
    stem = week_record.get("stem_or_demo", [])
    if len(stem) >= 1:
        s3_stem = 10.0
    elif t_id == "thang":  # Tổ trưởng chủ trì sinh hoạt chuyên môn NCBH
        s3_stem = 9.5
    else:
        s3_stem = 8.0  # Tích cực tham gia dự giờ phân tích bài học

    # - 3.3. Sử dụng ĐDDH & Ứng dụng CNTT, học liệu số (Tối đa 5.0đ)
    cntt = int(week_record.get("cntt_count", 0))
    dddh = int(week_record.get("dddh_count", 0))
    s3_tech = 5.0 if (cntt >= 2 or dddh >= 2) else (4.0 if (cntt >= 1 or dddh >= 1) else 3.0)

    tc3_score = round(s3_obs + s3_stem + s3_tech, 1)

    # 4. TIÊU CHUẨN IV: PHONG TRÀO, KIÊM NHIỆM & ĐẠO ĐỨC NGHỀ NGHIỆP (Tối đa 20.0 điểm)
    # - 4.1. Bồi dưỡng HSG / Phụ đạo HS / Hướng dẫn NCKH, TTN-NĐ (Tối đa 10.0đ)
    comp = week_record.get("competitions", [])
    if len(comp) >= 1 or t_id in ["thang", "cuc", "linh", "thanh"]:
        s4_comp = 10.0
    else:
        s4_comp = 8.5

    # - 4.2. Sinh hoạt tổ chuyên môn (thứ Tư hàng tuần) & tinh thần phối hợp (Tối đa 5.0đ)
    s4_teamwork = 5.0

    # - 4.3. Đạo đức nhà giáo, gương mẫu, chấp hành phân công (Tối đa 5.0đ)
    s4_ethics = 5.0

    tc4_score = round(s4_comp + s4_teamwork + s4_ethics, 1)

    # TỔNG ĐIỂM KPI (0 - 100)
    total_score = round(tc1_score + tc2_score + tc3_score + tc4_score, 1)
    if total_score > 100.0:
        total_score = 100.0

    # XẾP LOẠI CHẤT LƯỢNG (Theo Nghị định 90/2020/NĐ-CP & Nghị định 48/2023/NĐ-CP)
    if total_score >= 90.0 and late == 0 and not has_empty_subjs:
        classification = "Hoàn thành xuất sắc nhiệm vụ"
        rank_short = "HTXSNV"
        badge_color = "emerald"
    elif total_score >= 75.0:
        classification = "Hoàn thành tốt nhiệm vụ"
        rank_short = "HTTNV"
        badge_color = "blue"
    elif total_score >= 50.0:
        classification = "Hoàn thành nhiệm vụ"
        rank_short = "HTNV"
        badge_color = "amber"
    else:
        classification = "Không hoàn thành nhiệm vụ"
        rank_short = "KHTNV"
        badge_color = "rose"

    # Nhận xét cụ thể, công khai, minh bạch
    pros = []
    cons = []

    if s1_ppct == 10.0:
        pros.append("PPCT chuẩn xác 100%")
    else:
        cons.append(f"Còn trễ {late} tiết cần dạy bù")

    if s1_drive == 10.0:
        pros.append("Nộp KHBD Drive đầy đủ đúng hạn")
    elif has_empty_subjs:
        cons.append("Có thư mục môn trên Drive còn trống tệp")
    elif s1_drive < 8.0:
        cons.append("Chưa nộp kế hoạch bài dạy lên Drive")

    if tc2_score >= 28.0:
        pros.append("KHBD chuẩn CV 5512 & thể thức NĐ 30, tích hợp tốt NLS")
    if s3_obs == 10.0:
        pros.append("Dự giờ đồng nghiệp đủ định mức")
    if s3_stem == 10.0:
        pros.append("Thao giảng / nghiên cứu bài học chất lượng")

    pro_str = ", ".join(pros) if pros else "Thực hiện nhiệm vụ chuyên môn theo quy định"
    con_str = ", ".join(cons) if cons else "Không có khuyết điểm"

    notes = f"✔ Ưu điểm: {pro_str}. ❌ Tồn tại: {con_str}."

    return {
        "teacher_id": t_id,
        "teacher_name": t_name,
        "role": teacher_meta["role"],
        "subjects": subjects_str,
        "tc1_score": tc1_score,
        "tc2_score": tc2_score,
        "tc3_score": tc3_score,
        "tc4_score": tc4_score,
        "total_score": total_score,
        "classification": classification,
        "rank_short": rank_short,
        "badge_color": badge_color,
        "pros": pro_str,
        "cons": con_str,
        "notes": notes,
        "details": {
            "s1_attendance": s1_attendance,
            "s1_ppct": s1_ppct,
            "s1_drive": s1_drive,
            "s2_cv5512": s2_cv5512,
            "s2_format": s2_format,
            "s2_nls": s2_nls,
            "s2_records": s2_records,
            "s3_obs": s3_obs,
            "s3_stem": s3_stem,
            "s3_tech": s3_tech,
            "s4_comp": s4_comp,
            "s4_teamwork": s4_teamwork,
            "s4_ethics": s4_ethics
        }
    }



def _build_kpi_stats_dict(teachers_list: list, avg_score: float, rank_counts: dict) -> dict:
    """Tạo dict thống kê KPI chuẩn cho giao diện và báo cáo."""
    total = len(teachers_list) or 1
    cnt_exc = rank_counts.get("Hoàn thành xuất sắc nhiệm vụ", 0) + rank_counts.get("HTXSNV", 0)
    cnt_good = rank_counts.get("Hoàn thành tốt nhiệm vụ", 0) + rank_counts.get("HTTNV", 0)
    cnt_sat = rank_counts.get("Hoàn thành nhiệm vụ", 0) + rank_counts.get("HTNV", 0)
    cnt_unsat = rank_counts.get("Không hoàn thành nhiệm vụ", 0) + rank_counts.get("KHTNV", 0)
    return {
        "avg_score": avg_score,
        "count_excellent": cnt_exc,
        "pct_excellent": round((cnt_exc / total) * 100, 1),
        "count_good": cnt_good,
        "pct_good": round((cnt_good / total) * 100, 1),
        "count_satisfactory": cnt_sat,
        "pct_satisfactory": round((cnt_sat / total) * 100, 1),
        "count_unsatisfactory": cnt_unsat,
        "pct_unsatisfactory": round((cnt_unsat / total) * 100, 1)
    }

def calculate_all_teachers_kpi_week(week_num: int) -> dict:
    """
    Tính bảng điểm KPI tuần cho cả 8 giáo viên trong tổ chuyên môn.
    """
    all_data = load_all_notebook_data()
    week_records = all_data.get(str(week_num), [])
    if not week_records:
        week_records = get_notebook_week(week_num)

    drive_data = get_live_drive_data()
    w_info = get_week_date_info(week_num)

    teachers_kpi = []
    total_scores_sum = 0.0
    rank_counts = {"HTXSNV": 0, "HTTNV": 0, "HTNV": 0, "KHTNV": 0}

    # Map records by teacher_id or name
    r_map = {}
    for r in week_records:
        tid = r.get("teacher_id")
        if tid:
            r_map[tid] = r
        else:
            r_map[r.get("teacher_name", "").lower()] = r

    for t_meta in TEACHER_LIST:
        tid = t_meta["id"]
        t_rec = r_map.get(tid) or r_map.get(t_meta["name"].lower()) or {}
        kpi_res = calculate_teacher_kpi_week(t_meta, week_num, t_rec, drive_data)
        teachers_kpi.append(kpi_res)

        total_scores_sum += kpi_res["total_score"]
        rk = kpi_res["rank_short"]
        rank_counts[rk] = rank_counts.get(rk, 0) + 1

    # Sắp xếp theo tổng điểm giảm dần (công khai, minh bạch)
    teachers_kpi.sort(key=lambda x: x["total_score"], reverse=True)
    for idx, t in enumerate(teachers_kpi):
        t["stt"] = idx + 1

    avg_score = round(total_scores_sum / len(teachers_kpi), 1) if teachers_kpi else 0.0

    title = f"BẢNG ĐÁNH GIÁ KPI & XẾP LOẠI GIÁO VIÊN TUẦN {week_num}"
    subtitle = f"({w_info.get('range_str', '')} – {w_info.get('semester_str', '')} – Tháng {w_info.get('month_num', 9):02d})"
    stats = _build_kpi_stats_dict(teachers_kpi, avg_score, rank_counts)

    return {
        "period_type": "week",
        "period_val": week_num,
        "title": title,
        "subtitle": subtitle,
        "period_label": f"Tuần {week_num} ({w_info.get('range_str', '')}) - {w_info.get('semester_str', '')}",
        "stats": stats,
        "date_info": w_info,
        "legal_basis": [
            "Nghị định số 90/2020/NĐ-CP ngày 13/8/2020 của Chính phủ về đánh giá, xếp loại chất lượng viên chức;",
            "Nghị định số 48/2023/NĐ-CP ngày 17/7/2023 sửa đổi, bổ sung Nghị định số 90/2020/NĐ-CP;",
            "Thông tư số 20/2018/TT-BGDĐT ban hành chuẩn nghề nghiệp giáo viên cơ sở giáo dục phổ thông;",
            "Công văn số 5512/BGDĐT-GDTrH về xây dựng và tổ chức thực hiện kế hoạch giáo dục;",
            "Nghị định số 30/2020/NĐ-CP về công tác văn thư và chuẩn thể thức văn bản hành chính."
        ],
        "scale_standards": [
            {"name": "Tiêu chuẩn I", "max": 25.0, "desc": "Kỷ luật lao động, Ngày công & Tiến độ PPCT"},
            {"name": "Tiêu chuẩn II", "max": 30.0, "desc": "Chất lượng KHBD theo CV 5512, Thể thức NĐ 30 & Tích hợp NLS"},
            {"name": "Tiêu chuẩn III", "max": 25.0, "desc": "Giảng dạy, Dự giờ đồng nghiệp, Thao giảng & Đổi mới PPDH"},
            {"name": "Tiêu chuẩn IV", "max": 20.0, "desc": "Bồi dưỡng HSG, Kiêm nhiệm, Phối hợp & Đạo đức nghề nghiệp"}
        ],
        "total_teachers": len(teachers_kpi),
        "average_score": avg_score,
        "rank_counts": rank_counts,
        "teachers": teachers_kpi
    }


def calculate_all_teachers_kpi_month(month_num: int) -> dict:
    """
    Tính bảng điểm KPI Tháng (tổng hợp trung bình các tuần trong tháng) của 8 giáo viên.
    """
    cfg = MONTH_CYCLES.get(month_num, MONTH_CYCLES[9])
    weeks = cfg["weeks"]
    
    # Tính KPI từng tuần
    weekly_kpis = [calculate_all_teachers_kpi_week(w) for w in weeks]
    
    # Gom điểm theo giáo viên
    teacher_scores = {}
    for w_kpi in weekly_kpis:
        for t in w_kpi["teachers"]:
            tid = t["teacher_id"]
            if tid not in teacher_scores:
                teacher_scores[tid] = {
                    "teacher_id": tid,
                    "teacher_name": t["teacher_name"],
                    "role": t["role"],
                    "subjects": t["subjects"],
                    "tc1_sum": 0.0,
                    "tc2_sum": 0.0,
                    "tc3_sum": 0.0,
                    "tc4_sum": 0.0,
                    "total_sum": 0.0,
                    "pros_list": [],
                    "cons_list": [],
                    "weeks_count": 0
                }
            ts = teacher_scores[tid]
            ts["tc1_sum"] += t["tc1_score"]
            ts["tc2_sum"] += t["tc2_score"]
            ts["tc3_sum"] += t["tc3_score"]
            ts["tc4_sum"] += t["tc4_score"]
            ts["total_sum"] += t["total_score"]
            ts["weeks_count"] += 1
            if t["pros"] and t["pros"] not in ts["pros_list"]:
                ts["pros_list"].append(t["pros"])
            if t["cons"] and t["cons"] != "Không có khuyết điểm" and t["cons"] not in ts["cons_list"]:
                ts["cons_list"].append(t["cons"])

    teachers_month_kpi = []
    total_scores_sum = 0.0
    rank_counts = {"HTXSNV": 0, "HTTNV": 0, "HTNV": 0, "KHTNV": 0}

    for tid, ts in teacher_scores.items():
        w_cnt = ts["weeks_count"] or 1
        tc1 = round(ts["tc1_sum"] / w_cnt, 1)
        tc2 = round(ts["tc2_sum"] / w_cnt, 1)
        tc3 = round(ts["tc3_sum"] / w_cnt, 1)
        tc4 = round(ts["tc4_sum"] / w_cnt, 1)
        total = round(ts["total_sum"] / w_cnt, 1)

        if total >= 90.0 and len(ts["cons_list"]) == 0:
            classification = "Hoàn thành xuất sắc nhiệm vụ"
            rank_short = "HTXSNV"
            badge_color = "emerald"
        elif total >= 75.0:
            classification = "Hoàn thành tốt nhiệm vụ"
            rank_short = "HTTNV"
            badge_color = "blue"
        elif total >= 50.0:
            classification = "Hoàn thành nhiệm vụ"
            rank_short = "HTNV"
            badge_color = "amber"
        else:
            classification = "Không hoàn thành nhiệm vụ"
            rank_short = "KHTNV"
            badge_color = "rose"

        rank_counts[rank_short] = rank_counts.get(rank_short, 0) + 1
        total_scores_sum += total

        pro_str = "; ".join(ts["pros_list"][:2]) if ts["pros_list"] else "Thực hiện tốt kế hoạch giáo dục tháng"
        con_str = "; ".join(ts["cons_list"][:2]) if ts["cons_list"] else "Không có tồn tại đáng kể"

        teachers_month_kpi.append({
            "teacher_id": tid,
            "teacher_name": ts["teacher_name"],
            "role": ts["role"],
            "subjects": ts["subjects"],
            "tc1_score": tc1,
            "tc2_score": tc2,
            "tc3_score": tc3,
            "tc4_score": tc4,
            "total_score": total,
            "classification": classification,
            "rank_short": rank_short,
            "badge_color": badge_color,
            "pros": pro_str,
            "cons": con_str,
            "notes": f"✔ Ưu điểm: {pro_str}. ❌ Tồn tại: {con_str}."
        })

    teachers_month_kpi.sort(key=lambda x: x["total_score"], reverse=True)
    for idx, t in enumerate(teachers_month_kpi):
        t["stt"] = idx + 1

    avg_score = round(total_scores_sum / len(teachers_month_kpi), 1) if teachers_month_kpi else 0.0
    title = f"BẢNG ĐÁNH GIÁ KPI & XẾP LOẠI GIÁO VIÊN THÁNG {str(month_num).zfill(2)}"
    subtitle = f"(Tổng hợp từ Tuần {min(weeks)} đến Tuần {max(weeks)} – Năm học 2026 – 2027)"
    stats = _build_kpi_stats_dict(teachers_month_kpi, avg_score, rank_counts)

    return {
        "period_type": "month",
        "period_val": month_num,
        "title": title,
        "subtitle": subtitle,
        "period_label": f"Tháng {str(month_num).zfill(2)} (Tuần {min(weeks)} - {max(weeks)})",
        "stats": stats,
        "weeks_list": weeks,
        "total_teachers": len(teachers_month_kpi),
        "average_score": avg_score,
        "rank_counts": rank_counts,
        "teachers": teachers_month_kpi
    }


def calculate_all_teachers_kpi_semester(semester_num: int) -> dict:
    """
    Tính bảng điểm KPI Sơ kết Học kỳ I (Tuần 1-18) hoặc Học kỳ II (Tuần 19-35) của 8 giáo viên
    đạt chuẩn Nghị định số 90/2020/NĐ-CP và Thông tư số 20/2018/TT-BGDĐT.
    """
    weeks = list(range(1, 19)) if semester_num == 1 else list(range(19, 36))
    semester_name = "HỌC KỲ I" if semester_num == 1 else "HỌC KỲ II"

    weekly_kpis = [calculate_all_teachers_kpi_week(w) for w in weeks]

    teacher_scores = {}
    for w_kpi in weekly_kpis:
        for t in w_kpi["teachers"]:
            tid = t["teacher_id"]
            if tid not in teacher_scores:
                teacher_scores[tid] = {
                    "teacher_id": tid,
                    "teacher_name": t["teacher_name"],
                    "role": t["role"],
                    "subjects": t["subjects"],
                    "tc1_sum": 0.0,
                    "tc2_sum": 0.0,
                    "tc3_sum": 0.0,
                    "tc4_sum": 0.0,
                    "total_sum": 0.0,
                    "weeks_count": 0,
                    "rank_counts": {"HTXSNV": 0, "HTTNV": 0, "HTNV": 0, "KHTNV": 0}
                }
            ts = teacher_scores[tid]
            ts["tc1_sum"] += t["tc1_score"]
            ts["tc2_sum"] += t["tc2_score"]
            ts["tc3_sum"] += t["tc3_score"]
            ts["tc4_sum"] += t["tc4_score"]
            ts["total_sum"] += t["total_score"]
            ts["weeks_count"] += 1
            rk = t["rank_short"]
            ts["rank_counts"][rk] = ts["rank_counts"].get(rk, 0) + 1

    teachers_sem_kpi = []
    total_scores_sum = 0.0
    rank_counts = {"HTXSNV": 0, "HTTNV": 0, "HTNV": 0, "KHTNV": 0}

    for tid, ts in teacher_scores.items():
        w_cnt = ts["weeks_count"] or 1
        tc1 = round(ts["tc1_sum"] / w_cnt, 1)
        tc2 = round(ts["tc2_sum"] / w_cnt, 1)
        tc3 = round(ts["tc3_sum"] / w_cnt, 1)
        tc4 = round(ts["tc4_sum"] / w_cnt, 1)
        total = round(ts["total_sum"] / w_cnt, 1)

        rc = ts["rank_counts"]
        # Áp dụng chuẩn NĐ 90/2020/NĐ-CP & NĐ 48/2023/NĐ-CP
        if total >= 90.0 and rc.get("HTXSNV", 0) >= int(w_cnt * 0.4) and rc.get("HTNV", 0) <= 1 and rc.get("KHTNV", 0) == 0:
            classification = "Hoàn thành xuất sắc nhiệm vụ"
            rank_short = "HTXSNV"
            badge_color = "emerald"
            emulation_title = "Chiến sĩ thi đua cơ sở"
        elif total >= 75.0 and rc.get("KHTNV", 0) == 0:
            classification = "Hoàn thành tốt nhiệm vụ"
            rank_short = "HTTNV"
            badge_color = "blue"
            emulation_title = "Lao động tiên tiến"
        elif total >= 50.0:
            classification = "Hoàn thành nhiệm vụ"
            rank_short = "HTNV"
            badge_color = "amber"
            emulation_title = "Không bình xét"
        else:
            classification = "Không hoàn thành nhiệm vụ"
            rank_short = "KHTNV"
            badge_color = "rose"
            emulation_title = "Không bình xét"

        rank_counts[rank_short] = rank_counts.get(rank_short, 0) + 1
        total_scores_sum += total

        # Đánh giá Chuẩn nghề nghiệp giáo viên theo TT 20/2018/TT-BGDĐT
        if rank_short in ["HTXSNV", "HTTNV"]:
            tt20_rank = "Tốt"
        elif rank_short == "HTNV":
            tt20_rank = "Khá"
        else:
            tt20_rank = "Đạt"

        teachers_sem_kpi.append({
            "teacher_id": tid,
            "teacher_name": ts["teacher_name"],
            "role": ts["role"],
            "subjects": ts["subjects"],
            "tc1_score": tc1,
            "tc2_score": tc2,
            "tc3_score": tc3,
            "tc4_score": tc4,
            "total_score": total,
            "classification": classification,
            "rank_short": rank_short,
            "badge_color": badge_color,
            "tt20_rank": tt20_rank,
            "emulation_title": emulation_title,
            "rank_distribution": rc,
            "notes": f"Đạt chuẩn nghề nghiệp giáo viên loại {tt20_rank}. Đề xuất danh hiệu: {emulation_title}."
        })

    teachers_sem_kpi.sort(key=lambda x: x["total_score"], reverse=True)
    for idx, t in enumerate(teachers_sem_kpi):
        t["stt"] = idx + 1

    avg_score = round(total_scores_sum / len(teachers_sem_kpi), 1) if teachers_sem_kpi else 0.0

    title = f"BẢNG TỔNG HỢP KPI & ĐÁNH GIÁ XẾP LOẠI VIÊN CHỨC {semester_name}"
    subtitle = f"(Căn cứ Nghị định số 90/2020/NĐ-CP & Thông tư số 20/2018/TT-BGDĐT – Năm học 2026 – 2027)"
    stats = _build_kpi_stats_dict(teachers_sem_kpi, avg_score, rank_counts)

    return {
        "period_type": "semester",
        "period_val": semester_num,
        "title": title,
        "subtitle": subtitle,
        "period_label": f"{semester_name} (Tuần {min(weeks)} - {max(weeks)})",
        "stats": stats,
        "semester_name": semester_name,
        "weeks_range": f"Tuần {min(weeks)} đến Tuần {max(weeks)}",
        "total_teachers": len(teachers_sem_kpi),
        "average_score": avg_score,
        "rank_counts": rank_counts,
        "teachers": teachers_sem_kpi
    }


def calculate_all_teachers_kpi_year() -> dict:
    """
    Tính bảng điểm KPI Tổng kết Toàn diện CẢ NĂM HỌC (35 tuần) của 8 giáo viên
    đạt chuẩn Nghị định số 90/2020/NĐ-CP, Nghị định 48/2023/NĐ-CP và Thông tư 20/2018/TT-BGDĐT.
    """
    sem1 = calculate_all_teachers_kpi_semester(1)
    sem2 = calculate_all_teachers_kpi_semester(2)

    sem1_map = {t["teacher_id"]: t for t in sem1["teachers"]}
    sem2_map = {t["teacher_id"]: t for t in sem2["teachers"]}

    teachers_year_kpi = []
    total_scores_sum = 0.0
    rank_counts = {"HTXSNV": 0, "HTTNV": 0, "HTNV": 0, "KHTNV": 0}

    for t_meta in TEACHER_LIST:
        tid = t_meta["id"]
        s1 = sem1_map.get(tid, {})
        s2 = sem2_map.get(tid, {})

        # Trọng số tính điểm cả năm: HK1 (hệ số 1), HK2 (hệ số 2) hoặc trung bình cộng chuẩn
        s1_tot = s1.get("total_score", 85.0)
        s2_tot = s2.get("total_score", 85.0)
        year_score = round((s1_tot + s2_tot * 2) / 3, 1)

        tc1 = round((s1.get("tc1_score", 20.0) + s2.get("tc1_score", 20.0) * 2) / 3, 1)
        tc2 = round((s1.get("tc2_score", 25.0) + s2.get("tc2_score", 25.0) * 2) / 3, 1)
        tc3 = round((s1.get("tc3_score", 20.0) + s2.get("tc3_score", 20.0) * 2) / 3, 1)
        tc4 = round((s1.get("tc4_score", 15.0) + s2.get("tc4_score", 15.0) * 2) / 3, 1)

        # Xếp loại viên chức cả năm theo Nghị định 90/2020/NĐ-CP & NĐ 48/2023/NĐ-CP
        if year_score >= 90.0:
            classification = "Hoàn thành xuất sắc nhiệm vụ"
            rank_short = "HTXSNV"
            badge_color = "emerald"
            emulation_title = "Chiến sĩ thi đua cơ sở"
            reward_proposal = "Đề nghị UBND huyện tặng Giấy khen"
            tt20_rank = "Tốt"
        elif year_score >= 75.0:
            classification = "Hoàn thành tốt nhiệm vụ"
            rank_short = "HTTNV"
            badge_color = "blue"
            emulation_title = "Lao động tiên tiến"
            reward_proposal = "Đề nghị Hiệu trưởng khen thưởng"
            tt20_rank = "Tốt"
        elif year_score >= 50.0:
            classification = "Hoàn thành nhiệm vụ"
            rank_short = "HTNV"
            badge_color = "amber"
            emulation_title = "Không bình xét"
            reward_proposal = "Đạt chuẩn hoàn thành nhiệm vụ"
            tt20_rank = "Khá"
        else:
            classification = "Không hoàn thành nhiệm vụ"
            rank_short = "KHTNV"
            badge_color = "rose"
            emulation_title = "Không bình xét"
            reward_proposal = "Cần bồi dưỡng chuyên môn nghiệp vụ"
            tt20_rank = "Đạt"

        rank_counts[rank_short] = rank_counts.get(rank_short, 0) + 1
        total_scores_sum += year_score

        teachers_year_kpi.append({
            "teacher_id": tid,
            "teacher_name": t_meta["name"],
            "role": t_meta["role"],
            "subjects": ", ".join(t_meta["subjects"]),
            "sem1_score": s1_tot,
            "sem2_score": s2_tot,
            "tc1_score": tc1,
            "tc2_score": tc2,
            "tc3_score": tc3,
            "tc4_score": tc4,
            "total_score": year_score,
            "classification": classification,
            "rank_short": rank_short,
            "badge_color": badge_color,
            "tt20_rank": tt20_rank,
            "emulation_title": emulation_title,
            "reward_proposal": reward_proposal,
            "notes": f"Chuẩn nghề nghiệp: {tt20_rank}. Danh hiệu: {emulation_title}. {reward_proposal}."
        })

    teachers_year_kpi.sort(key=lambda x: x["total_score"], reverse=True)
    for idx, t in enumerate(teachers_year_kpi):
        t["stt"] = idx + 1

    avg_score = round(total_scores_sum / len(teachers_year_kpi), 1) if teachers_year_kpi else 0.0

    title = "BẢNG TỔNG HỢP KPI & ĐÁNH GIÁ XẾP LOẠI VIÊN CHỨC CẢ NĂM HỌC 2026 – 2027"
    subtitle = "(Đánh giá toàn diện 35 tuần theo Nghị định số 90/2020/NĐ-CP, NĐ 48/2023/NĐ-CP & TT 20/2018/TT-BGDĐT)"
    stats = _build_kpi_stats_dict(teachers_year_kpi, avg_score, rank_counts)

    return {
        "period_type": "year",
        "title": title,
        "subtitle": subtitle,
        "period_label": "Cả Năm Học 2026 - 2027 (35 Tuần)",
        "stats": stats,
        "total_weeks": 35,
        "total_teachers": len(teachers_year_kpi),
        "average_score": avg_score,
        "rank_counts": rank_counts,
        "teachers": teachers_year_kpi
    }


def export_kpi_report_word(period_type: str = "week", period_val: int = 1, output_path: str = None) -> str:
    """
    Xuất File Word Báo cáo Đánh giá KPI & Xếp loại Giáo viên/Viên chức
    chuẩn thể thức văn bản hành chính theo Nghị định số 30/2020/NĐ-CP.
    """
    if period_type == "week":
        data = calculate_all_teachers_kpi_week(period_val)
        file_prefix = f"BANG_KPI_TUAN_{period_val}"
    elif period_type == "month":
        data = calculate_all_teachers_kpi_month(period_val)
        file_prefix = f"BANG_KPI_THANG_{period_val}"
    elif period_type == "semester":
        data = calculate_all_teachers_kpi_semester(period_val)
        file_prefix = f"BANG_KPI_HOC_KY_{period_val}"
    else:
        data = calculate_all_teachers_kpi_year()
        file_prefix = "BANG_KPI_CA_NAM_HOC_2026_2027"

    if not output_path:
        output_path = os.path.join(CORRECTED_DIR, f"{file_prefix}.docx")

    doc = docx.Document()

    # Căn lề Landscape cho bảng nhiều cột hoặc Portrait chuẩn NĐ 30: Top 20mm, Bottom 20mm, Left 25mm, Right 15mm
    for section in doc.sections:
        section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
        section.page_width = Mm(297)
        section.page_height = Mm(210)
        section.top_margin = Mm(18)
        section.bottom_margin = Mm(18)
        section.left_margin = Mm(25)
        section.right_margin = Mm(15)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(2)

    def set_tnr(run, size_pt=None, bold=None, italic=None, color_rgb=None):
        run.font.name = 'Times New Roman'
        rPr = run._r.get_or_add_rPr()
        rF = rPr.find(qn('w:rFonts'))
        if rF is None:
            rF = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>')
            rPr.append(rF)
        else:
            rF.set(qn('w:ascii'), 'Times New Roman')
            rF.set(qn('w:hAnsi'), 'Times New Roman')
            rF.set(qn('w:cs'), 'Times New Roman')
            rF.set(qn('w:eastAsia'), 'Times New Roman')
        if size_pt is not None:
            run.font.size = Pt(size_pt)
        if bold is not None:
            run.font.bold = bold
        if italic is not None:
            run.font.italic = italic
        if color_rgb is not None:
            run.font.color.rgb = color_rgb
        return run

    def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def set_cell_bg(cell, color_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        tcPr.append(shd)

    def set_table_borders(table):
        tblPr = table._tbl.tblPr
        borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>
                <w:insideV w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>
            </w:tblBorders>
        ''')
        tblPr.append(borders)

    # 1. HEADER (2 CỘT)
    t0 = doc.add_table(rows=3, cols=2)
    t0.alignment = WD_TABLE_ALIGNMENT.CENTER
    t0.autofit = False
    t0.rows[0].cells[0].width = Mm(120)
    t0.rows[0].cells[1].width = Mm(135)
    t0.rows[1].cells[0].width = Mm(120)
    t0.rows[1].cells[1].width = Mm(135)
    t0.rows[2].cells[0].width = Mm(120)
    t0.rows[2].cells[1].width = Mm(135)

    p00 = t0.rows[0].cells[0].paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_tnr(p00.add_run("TRƯỜNG TH & THCS PHƯỚC HƯNG"), size_pt=11)

    p01 = t0.rows[0].cells[1].paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_tnr(p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"), size_pt=11.5, bold=True)

    p10 = t0.rows[1].cells[0].paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_tnr(p10.add_run("TỔ: TOÁN – KHTN – CÔNG NGHỆ"), size_pt=11.5, bold=True)

    p11 = t0.rows[1].cells[1].paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_tnr(p11.add_run("Độc lập – Tự do – Hạnh phúc"), size_pt=12, bold=True)

    p20 = t0.rows[2].cells[0].paragraphs[0]
    p20.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_tnr(p20.add_run("Số: ...... /TB-TCM"), size_pt=10.5, italic=True)

    now_d = datetime.now()
    p21 = t0.rows[2].cells[1].paragraphs[0]
    p21.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_tnr(p21.add_run(f"Nhơn Hội, ngày {now_d.day:02d} tháng {now_d.month:02d} năm {now_d.year}"), size_pt=10.5, italic=True)

    # ĐƯỜNG KẺ
    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line.paragraph_format.space_before = Pt(0)
    p_line.paragraph_format.space_after = Pt(4)
    set_tnr(p_line.add_run("_______________"), size_pt=9, bold=True)

    # TIÊU ĐỀ
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(2)
    set_tnr(p_title.add_run(data["title"] + "\n"), size_pt=13.5, bold=True, color_rgb=RGBColor(0, 32, 96))
    set_tnr(p_title.add_run(data.get("subtitle", "")), size_pt=11, italic=True)

    # CĂN CỨ PHÁP LÝ
    p_cc = doc.add_paragraph()
    p_cc.paragraph_format.space_before = Pt(4)
    p_cc.paragraph_format.space_after = Pt(2)
    set_tnr(p_cc.add_run("Căn cứ pháp lý đánh giá, xếp loại:"), size_pt=10, bold=True, italic=True)
    cc_lines = [
        "• Nghị định số 90/2020/NĐ-CP và Nghị định số 48/2023/NĐ-CP của Chính phủ về đánh giá, xếp loại chất lượng viên chức;",
        "• Thông tư số 20/2018/TT-BGDĐT ban hành quy định chuẩn nghề nghiệp giáo viên cơ sở giáo dục phổ thông;",
        "• Công văn số 5512/BGDĐT-GDTrH về xây dựng và tổ chức thực hiện kế hoạch giáo dục của nhà trường;",
        "• Quyết định 219/QĐ-TH&THCSPH, Quy chế chuyên môn 263 và Quy chế thi đua khen thưởng 265 của trường TH & THCS Phước Hưng."
    ]
    for ccl in cc_lines:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.left_indent = Inches(0.2)
        p_item.paragraph_format.space_after = Pt(1)
        set_tnr(p_item.add_run(ccl), size_pt=9.5, italic=True)

    # BẢNG ĐIỂM KPI 100 ĐIỂM
    p_table_title = doc.add_paragraph()
    p_table_title.paragraph_format.space_before = Pt(6)
    p_table_title.paragraph_format.space_after = Pt(2)
    set_tnr(p_table_title.add_run("BẢNG ĐIỂM ĐÁNH GIÁ KPI CHI TIẾT THEO THANG 100 ĐIỂM (CÔNG KHAI – MINH BẠCH):"), size_pt=10.5, bold=True)

    cols_count = 10
    table = doc.add_table(rows=len(data["teachers"]) + 1, cols=cols_count)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)

    headers = [
        ("STT", Mm(10)),
        ("Họ và tên giáo viên", Mm(42)),
        ("Môn phụ trách", Mm(28)),
        ("Tiêu chuẩn I\n(25.0đ)", Mm(24)),
        ("Tiêu chuẩn II\n(30.0đ)", Mm(24)),
        ("Tiêu chuẩn III\n(25.0đ)", Mm(24)),
        ("Tiêu chuẩn IV\n(20.0đ)", Mm(24)),
        ("Tổng điểm\n(100.0đ)", Mm(24)),
        ("Xếp loại chất lượng\n(NĐ 90/2020)", Mm(35)),
        ("Nhận xét & Đề xuất", Mm(38))
    ]

    for col_idx, (text, width) in enumerate(headers):
        cell = table.rows[0].cells[col_idx]
        cell.width = width
        set_cell_bg(cell, "1F4E79")
        set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_tnr(p.add_run(text), size_pt=9.5, bold=True, color_rgb=RGBColor(255, 255, 255))

    for row_idx, t in enumerate(data["teachers"]):
        row = table.rows[row_idx + 1]
        bg_col = "F9FAFB" if row_idx % 2 == 1 else "FFFFFF"

        values = [
            (str(t.get("stt", row_idx + 1)), WD_ALIGN_PARAGRAPH.CENTER, False),
            (t["teacher_name"] + (" (Tổ trưởng)" if t["teacher_id"] == "thang" else ""), WD_ALIGN_PARAGRAPH.LEFT, True),
            (t["subjects"], WD_ALIGN_PARAGRAPH.LEFT, False),
            (str(t["tc1_score"]), WD_ALIGN_PARAGRAPH.CENTER, False),
            (str(t["tc2_score"]), WD_ALIGN_PARAGRAPH.CENTER, False),
            (str(t["tc3_score"]), WD_ALIGN_PARAGRAPH.CENTER, False),
            (str(t["tc4_score"]), WD_ALIGN_PARAGRAPH.CENTER, False),
            (str(t["total_score"]), WD_ALIGN_PARAGRAPH.CENTER, True),
            (t["classification"], WD_ALIGN_PARAGRAPH.CENTER, True),
            (t["notes"], WD_ALIGN_PARAGRAPH.LEFT, False)
        ]

        for col_idx, (val, align, is_bold) in enumerate(values):
            cell = row.cells[col_idx]
            cell.width = headers[col_idx][1]
            set_cell_bg(cell, bg_col)
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            p = cell.paragraphs[0]
            p.alignment = align
            c_rgb = None
            if col_idx == 7:  # Tổng điểm
                c_rgb = RGBColor(0, 100, 0) if t["total_score"] >= 90 else (RGBColor(0, 50, 150) if t["total_score"] >= 75 else RGBColor(180, 0, 0))
            elif col_idx == 8:  # Xếp loại
                c_rgb = RGBColor(0, 100, 0) if "xuất sắc" in val.lower() else (RGBColor(0, 50, 150) if "tốt" in val.lower() else RGBColor(150, 75, 0))
            set_tnr(p.add_run(val), size_pt=9, bold=is_bold, color_rgb=c_rgb)

    # TỔNG KẾT & KÝ TÊN
    p_stat = doc.add_paragraph()
    p_stat.paragraph_format.space_before = Pt(8)
    p_stat.paragraph_format.space_after = Pt(4)
    r_stat = p_stat.add_run(
        f"Tổng hợp kết quả đánh giá ({data['total_teachers']} giáo viên): "
        f"Điểm trung bình toàn tổ: {data['average_score']}/100 điểm. "
        f"Hoàn thành xuất sắc nhiệm vụ: {data['rank_counts'].get('HTXSNV', 0)} GV ({round(data['rank_counts'].get('HTXSNV', 0)/data['total_teachers']*100, 1)}%); "
        f"Hoàn thành tốt nhiệm vụ: {data['rank_counts'].get('HTTNV', 0)} GV ({round(data['rank_counts'].get('HTTNV', 0)/data['total_teachers']*100, 1)}%); "
        f"Hoàn thành nhiệm vụ: {data['rank_counts'].get('HTNV', 0)} GV; "
        f"Không hoàn thành nhiệm vụ: {data['rank_counts'].get('KHTNV', 0)} GV."
    )
    set_tnr(r_stat, size_pt=10, bold=True, italic=True)

    # Chữ ký 2 cột
    t_sign = doc.add_table(rows=2, cols=2)
    t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_sign.autofit = False
    t_sign.rows[0].cells[0].width = Mm(125)
    t_sign.rows[0].cells[1].width = Mm(130)

    c0 = t_sign.rows[0].cells[0].paragraphs[0]
    c0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_tnr(c0.add_run("TỔ TRƯỞNG CHUYÊN MÔN\n"), size_pt=11, bold=True)
    set_tnr(c0.add_run("(Ký và ghi rõ họ tên)\n\n\n\n"), size_pt=9.5, italic=True)
    set_tnr(c0.add_run("Lê Văn Thắng"), size_pt=11.5, bold=True)

    c1 = t_sign.rows[0].cells[1].paragraphs[0]
    c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_tnr(c1.add_run("HIỆU TRƯỞNG PHÊ DUYỆT\n"), size_pt=11, bold=True)
    set_tnr(c1.add_run("(Ký, ghi rõ họ tên và đóng dấu)\n\n\n\n"), size_pt=9.5, italic=True)
    set_tnr(c1.add_run("Trần Quốc Việt"), size_pt=11.5, bold=True)

    doc.save(output_path)
    return output_path
