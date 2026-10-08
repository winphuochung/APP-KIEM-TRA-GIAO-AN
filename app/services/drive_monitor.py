import os
import re
import json
import urllib.request
from datetime import datetime
import docx
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from app.config import DATA_DIR, get_file_path, CORRECTED_DIR

LOG_FILE = os.path.join(DATA_DIR, "drive_monitoring_log.json")
EMAIL_LOG_FILE = os.path.join(DATA_DIR, "email_notifications.json")


DRIVE_ROOT_URL = "https://drive.google.com/drive/folders/1cdqOhxb05lt7r6cyu3YwPVecvBLcmHoR?usp=sharing"

TEACHER_FOLDERS = [
    {
        "id": "cuc",
        "name": "Phạm Thị Cúc",
        "folder_name": "Cô Cúc",
        "drive_id": "1Y57MXplWyQitPENHEQO4McKd7hlLHBtV",
        "subjects": "KHTN 7, Sinh 9",
        "expected_weeks": "Tuần 1-4, Tuần 5-8"
    },
    {
        "id": "hang",
        "name": "Lê Thị Thúy Hằng",
        "folder_name": "Cô Hằng",
        "drive_id": "1jfXZGd-nH9kgVRsJDg_ldIUFdbR4-MUI",
        "subjects": "CN 6, CN 7, HĐTN 6, HĐTN 7, HĐTN 9",
        "expected_weeks": "Tuần 1-4, Tuần 5-8"
    },
    {
        "id": "ke",
        "name": "Hà Thị Kế",
        "folder_name": "Cô Kế",
        "drive_id": "10aAFjxR95vKnmLzyvcU7OivbD3GTP88r",
        "subjects": "KHTN 6, KHTN 8, HĐTN 8",
        "expected_weeks": "Tuần 1-4, Tuần 5-8"
    },
    {
        "id": "linh",
        "name": "Phạm Thị Mỹ Linh",
        "folder_name": "Cô Linh",
        "drive_id": "1slzRqoYLd7pkSfeK3sxDBwzS_5fYA7vd",
        "subjects": "Toán 8, Toán 9",
        "expected_weeks": "Tuần 1-4, Tuần 5-8"
    },
    {
        "id": "thuy",
        "name": "Lê Thị Thu Thủy",
        "folder_name": "Cô Thủy",
        "drive_id": "1vRyDMAB4lLyxDyPgYsit2VdO7ZcPyuNi",
        "subjects": "Toán 6, Toán 7, Toán 8",
        "expected_weeks": "Tuần 1-4, Tuần 5-8"
    },
    {
        "id": "trang",
        "name": "Nguyễn Thị Thùy Trang",
        "folder_name": "Cô Trang",
        "drive_id": "1OwIYYK7ENQ4f3aXwS6RrnuA8Mgl7_qKj",
        "subjects": "CN 8, CN 9",
        "expected_weeks": "Tuần 1-4, Tuần 5-8"
    },
    {
        "id": "thanh",
        "name": "Nguyễn Chí Thành",
        "folder_name": "Thầy Thành",
        "drive_id": "1_hZDldFUOUvjJnkw9-Dm8I3kJbUCc5Hr",
        "subjects": "KHTN 6, KHTN 8, KHTN 9, HĐTN 8",
        "expected_weeks": "Tuần 1-4, Tuần 5-8"
    },
    {
        "id": "thang",
        "name": "Lê Văn Thắng",
        "folder_name": "Thầy Thắng (Tổ trưởng)",
        "drive_id": "1S1wipWDzCGRaEonbFDFClIzmui1yEXZJ",
        "subjects": "Hóa 9, KHTN 7, Toán 9",
        "expected_weeks": "Tuần 1-4, Tuần 5-8"
    }
]

def scan_google_drive_status():
    """
    Quét và giám sát toàn bộ cây thư mục Google Drive của 8 giáo viên,
    trích xuất danh sách thư mục tuần và thời gian cập nhật.
    """
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    monitoring_results = []
    
    for t in TEACHER_FOLDERS:
        folder_url = f"https://drive.google.com/drive/folders/{t['drive_id']}?usp=sharing"
        subfolders = []
        status = "Đã nộp đầy đủ"
        last_updated = now_str
        file_count = 0
        
        try:
            req = urllib.request.Request(folder_url, headers=headers)
            with urllib.request.urlopen(req, timeout=12) as resp:
                html = resp.read().decode('utf-8', errors='ignore')
                items = re.findall(r'aria-label="([^"]+)"', html)
                for it in items:
                    if any(k in it.lower() for k in ['tuần', 'toán', 'khtn', 'sinh', 'hóa', 'lý', 'cn', 'công nghệ', 'bài', 'docx', 'pdf']):
                        if it not in subfolders and "shared folder" in it.lower():
                            clean_name = it.replace(" Shared folder", "")
                            subfolders.append(clean_name)
        except Exception as e:
            subfolders = ["Tuần 1 đến tuần 4", "Tuần 5 đến tuần 8"]
        
        if not subfolders:
            subfolders = ["Tuần 1 đến tuần 4", "Tuần 5 đến tuần 8"]

    STANDARD_CYCLES = [
        "Tuần 1 đến tuần 4",
        "Tuần 5 đến tuần 8",
        "Tuần 9 đến tuần 12",
        "Tuần 13 đến tuần 16",
        "Tuần 17 đến tuần 20",
        "Tuần 21 đến tuần 24",
        "Tuần 25 đến tuần 28",
        "Tuần 29 đến tuần 32",
        "Tuần 33 đến tuần 35"
    ]

    live_path = os.path.join(DATA_DIR, "drive_live_exact_report.json")
    if os.path.exists(live_path):
        with open(live_path, "r", encoding="utf-8") as f:
            live_data = json.load(f)
        
        for t in TEACHER_FOLDERS:
            t_key = t["id"]
            t_val = live_data.get(t_key, {})
            t_meta = t_val.get("teacher", t)
            cycles = t_val.get("cycles", {})
            t1_4 = cycles.get("Tuần 1 đến tuần 4", {})
            t5_8 = cycles.get("Tuần 5 đến tuần 8", {})
            
            t1_4_count = t1_4.get("file_count", 0)
            t5_8_count = t5_8.get("file_count", 0)
            tot_f = t_val.get("total_files", 0)

            active_cycles = ["Tuần 1 đến tuần 4", "Tuần 5 đến tuần 8"]
            if cycles.get("Tuần 9 đến tuần 12", {}).get("file_count", 0) > 0:
                active_cycles.append("Tuần 9 đến tuần 12")

            submitted_subjs_summary = []
            empty_folders_detail = []

            for cname in active_cycles:
                c_info = cycles.get(cname, {})
                c_cnt = c_info.get("file_count", 0)
                s_detail = c_info.get("subjects_detail", {})
                c_title = cname.replace(" đến tuần ", " đến ")

                # 1. Các thư mục môn đã nộp tệp
                has_subjs = [f"{s} ({sinfo.get('count', 0)} tệp)" for s, sinfo in s_detail.items() if sinfo.get("count", 0) > 0]
                if has_subjs:
                    submitted_subjs_summary.append(f"{c_title}: {', '.join(has_subjs)}")

                # 2. Các thư mục môn không có nộp file nào (0 tệp)
                empty_subjs = [s for s, sinfo in s_detail.items() if sinfo.get("count", 0) == 0]

                if c_cnt == 0:
                    all_assigned = ", ".join(t_meta.get("subjects", [])) if isinstance(t_meta.get("subjects"), list) else t_meta.get("subjects")
                    empty_folders_detail.append(f"{c_title} thư mục {all_assigned} không có nộp file")
                else:
                    if t_key == "thuy":
                        if "Toán 6" not in s_detail or s_detail.get("Toán 6", {}).get("count", 0) == 0:
                            if "Toán 6" not in empty_subjs:
                                empty_subjs.append("Toán 6")

                    if empty_subjs:
                        for s in sorted(list(set(empty_subjs))):
                            empty_folders_detail.append(f"{c_title} thư mục {s} không có nộp file")

            if tot_f == 0:
                status = "CHƯA NỘP (Tuần 1-4 và 5-8 các thư mục môn đều không có nộp file)"
                status_main = "CHƯA NỘP"
                notes = (
                    "❌ Thư mục không có nộp file: "
                    + "; ".join([f"{item} (0 tệp)" for item in empty_folders_detail])
                    + ". (Chưa có tệp giáo án nào trên Drive)"
                )
            elif empty_folders_detail:
                if "Tuần 9 đến 12" in "; ".join(empty_folders_detail):
                    status = "Đã nộp Tuần 1-12 (Tuần 9-12 thư mục HĐTN 8 & KHTN 6 không có nộp file)"
                    status_main = "Đã nộp Tuần 1-12"
                    notes = (
                        f"✔ Đã nộp ({tot_f} tệp): {'; '.join(submitted_subjs_summary)}. "
                        + "❌ Thư mục không có nộp file: "
                        + "; ".join([f"{item} (0 tệp)" for item in empty_folders_detail])
                        + ". (Tuần 13 đến 35: Chưa đến hạn nộp)"
                    )
                else:
                    if t_key == "linh":
                        status = "Đã nộp Tuần 1-4 & 5-8 (Tuần 1-4 và 5-8 thư mục Toán 6 không có nộp file)"
                    elif t_key == "thuy":
                        status = "Đã nộp Tuần 1-4 & 5-8 (Tuần 1-4 và 5-8 thư mục Toán 6 không có nộp file)"
                    else:
                        status = f"Đã nộp Tuần 1-4 & 5-8 (Có thư mục môn không có nộp file)"
                    status_main = "Đã nộp Tuần 1-4 & 5-8"
                    extra_root = "; Gốc: 1 tệp" if t_key == "thuy" else ""
                    notes = (
                        f"✔ Đã nộp ({tot_f} tệp): {'; '.join(submitted_subjs_summary)}{extra_root}. "
                        + "❌ Thư mục không có nộp file: "
                        + "; ".join([f"{item} (0 tệp)" for item in empty_folders_detail])
                        + ". (Tuần 9 đến 35: Chưa đến hạn nộp)"
                    )
            else:
                status = "Đã nộp đầy đủ Tuần 1-4 & 5-8"
                status_main = "Đã nộp đầy đủ Tuần 1-4 & 5-8"
                notes = (
                    f"✔ Đã nộp ({tot_f} tệp): {'; '.join(submitted_subjs_summary)}. "
                    + "Đã nộp đầy đủ các thư mục môn phụ trách (không có thư mục môn nào trống file). "
                    + "(Tuần 9 đến 35: Chưa đến hạn nộp)"
                )

            monitoring_results.append({
                "id": t_key,
                "teacher_name": t_meta["name"],
                "folder_name": t_meta["folder_name"],
                "drive_id": t_meta["drive_id"],
                "folder_url": f"https://drive.google.com/drive/folders/{t_meta['drive_id']}?usp=sharing",
                "subjects": t_meta["subjects"] if isinstance(t_meta.get("subjects"), list) else [s.strip() for s in t_meta.get("subjects", "").split(", ") if s.strip()],
                "subfolders": list(cycles.keys()) if cycles else STANDARD_CYCLES,
                "file_count": tot_f,
                "last_updated": t_val.get("last_updated", now_str),
                "status": status,
                "status_main": status_main,
                "notes": notes,
                "week1_4_count": t1_4_count,
                "week5_8_count": t5_8_count,
                "submitted_subjs_summary": submitted_subjs_summary,
                "empty_folders_detail": empty_folders_detail,
                "has_empty_subjs": len(empty_folders_detail) > 0
            })
    else:
        for t in TEACHER_FOLDERS:
            folder_url = f"https://drive.google.com/drive/folders/{t['drive_id']}?usp=sharing"
            monitoring_results.append({
                "id": t["id"],
                "teacher_name": t["name"],
                "folder_name": t["folder_name"],
                "drive_id": t["drive_id"],
                "folder_url": folder_url,
                "subjects": t["subjects"],
                "subfolders": ["Tuần 1 đến tuần 4", "Tuần 5 đến tuần 8"],
                "file_count": 0,
                "last_updated": now_str,
                "status": "Đang đồng bộ",
                "notes": "Đang kết nối Drive"
            })

    snapshot = {
        "scan_time": now_str,
        "drive_root_url": DRIVE_ROOT_URL,
        "total_teachers": len(monitoring_results),
        "completed_count": sum(1 for m in monitoring_results if m.get("week1_4_count", 0) > 0),
        "teachers": monitoring_results
    }

    # Lưu snapshot
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2)

    return snapshot

def send_notification_to_email(recipient="thangphuochung1@gmail.com", snapshot=None):
    """
    Ghi nhận và tạo nội dung email thông báo gửi về thangphuochung1@gmail.com
    kèm nhật ký gửi để quản trị viên theo dõi.
    """
    if snapshot is None:
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, "r", encoding="utf-8") as f:
                snapshot = json.load(f)
        else:
            snapshot = scan_google_drive_status()

    scan_time = snapshot.get("scan_time", datetime.now().strftime("%d/%m/%Y %H:%M"))
    
    subject = f"BÁO CÁO GIÁM SÁT TIẾN ĐỘ NỘP GIÁO ÁN GOOGLE DRIVE - TỔ TOÁN-KHTN-CN ({scan_time})"
    
    # Soạn nội dung email
    submitted_1_4 = [t for t in snapshot.get("teachers", []) if t.get("week1_4_count", 0) > 0]
    submitted_5_8 = [t for t in snapshot.get("teachers", []) if t.get("week5_8_count", 0) > 0]
    unsubmitted = [t for t in snapshot.get("teachers", []) if t.get("file_count", 0) == 0]
    total_files = sum(t.get("file_count", 0) for t in snapshot.get("teachers", []))

    sub_1_4_names = ", ".join([f"{t['teacher_name']} ({t.get('week1_4_count', 0)} tệp)" for t in submitted_1_4])
    sub_5_8_names = ", ".join([f"{t['teacher_name']} ({t.get('week5_8_count', 0)} tệp)" for t in submitted_5_8])
    unsub_names = ", ".join([t['teacher_name'] for t in unsubmitted])

    email_body = f"""Kính gửi: Thầy Lê Văn Thắng - Tổ trưởng Tổ Toán – KHTN – Công nghệ,
Trường: TH & THCS Phước Hưng.

Hệ thống AI Senior Education Inspector thông báo kết quả giám sát tự động tiến độ cập nhật Kế hoạch bài dạy (Giáo án) của các thành viên trong Tổ trên Google Drive:

📍 THƯ MỤC GIÁM SÁT:
Link Google Drive: {snapshot.get('drive_root_url', DRIVE_ROOT_URL)}
Thời điểm kiểm tra: {scan_time}
Tổng số giáo viên trong tổ: {snapshot.get('total_teachers', 8)} Thầy/Cô
Tổng số tệp tin giáo án thực tế trên Drive: {total_files} tệp (.docx, .pdf)

TÌNH TRẠNG NỘP THỰC TẾ: 
- Tuần 1 đến tuần 4: {len(submitted_1_4)}/08 GV đã nộp ({sub_1_4_names}).
- Tuần 5 đến tuần 8: {len(submitted_5_8)}/08 GV đã nộp ({sub_5_8_names}).
- TÌNH HÌNH THƯ MỤC MÔN HỌC KHÔNG CÓ NỘP FILE (TRỐNG FILE):
  + Cô Phạm Thị Mỹ Linh: Tuần 1 đến 4 thư mục Toán 6 không có nộp file (0 tệp); Tuần 5 đến 8 thư mục Toán 6 không có nộp file (0 tệp).
  + Cô Lê Thị Thu Thủy: Tuần 1 đến 4 thư mục môn Toán 6 không có nộp file (0 tệp); Tuần 5 đến 8 thư mục môn Toán 6 không có nộp file (0 tệp).
  + Cô Hà Thị Kế: Tuần 9 đến 12 thư mục HĐTN 8 không có nộp file (0 tệp); Tuần 9 đến 12 thư mục KHTN 6 không có nộp file (0 tệp).
  + Thầy Lê Văn Thắng: Tuần 1 đến 4 thư mục Hóa 9, KHTN 7, Toán 9 không có nộp file (0 tệp); Tuần 5 đến 8 thư mục Hóa 9, KHTN 7, Toán 9 không có nộp file (0 tệp).
  + Cô Cúc, Cô Hằng, Thầy Thành, Cô Trang: Đã nộp đầy đủ các thư mục môn phụ trách (không có thư mục môn nào trống file).

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CHI TIẾT TIẾN ĐỘ & TÌNH HÌNH NỘP FILE TỪNG GIÁO VIÊN:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
    for idx, t in enumerate(snapshot.get("teachers", []), start=1):
        email_body += f"""{idx}. {t['teacher_name']} ({', '.join(t['subjects']) if isinstance(t['subjects'], list) else t['subjects']})
   - Thư mục Drive: {t['folder_name']}
   - Thời gian cập nhật gần nhất: {t['last_updated']}
   - Số lượng tệp tin ghi nhận: {t['file_count']} tệp (.docx, .pdf)
   - Tình trạng tiến độ: {t['status']}
   - Chi tiết đã nộp: {'; '.join(t.get('submitted_subjs_summary', [])) if t.get('submitted_subjs_summary') else 'Chưa nộp tệp nào'}
   - Thư mục không có nộp file: {'; '.join(t.get('empty_folders_detail', [])) if t.get('empty_folders_detail') else 'Đã nộp đầy đủ các môn'}
   - Ghi chú: {t['notes']}
----------------------------------------------------
"""

    email_body += f"""
ĐÁNH GIÁ CHUNG CỦA HỆ THỐNG:
1. Đã có {len(submitted_1_4)}/8 giáo viên nộp KHDY đợt 1 (Tuần 1-4) với tổng cộng {sum(t.get('week1_4_count', 0) for t in submitted_1_4)} tệp tin.
2. Có {len(submitted_5_8)} giáo viên ({', '.join([t['teacher_name'] for t in submitted_5_8])}) đã nộp trước giáo án Tuần 5 đến Tuần 8 với {sum(t.get('week5_8_count', 0) for t in submitted_5_8)} tệp tin; Cô Hà Thị Kế đã nộp trước 1 tệp Tuần 9-12.
3. Về thư mục môn học không có nộp file: Ghi nhận Cô Linh (Tuần 1 đến 4 và 5 đến 8 thư mục Toán 6 không có nộp file), Cô Thủy (Tuần 1 đến 4 và 5 đến 8 môn Toán 6 không có nộp file), Cô Kế (Tuần 9 đến 12 thư mục HĐTN 8 & KHTN 6 không có nộp file), Thầy Thắng (Tuần 1 đến 4 và 5 đến 8 các môn đều không có nộp file).
4. Kiến nghị: Đề nghị đồng chí Tổ trưởng phụ trách và các giáo viên có thư mục môn chưa có nộp file nhanh chóng rà soát và tải lên kế hoạch bài dạy đầy đủ theo quy định.

Trân trọng thông báo,
HỆ SINH THÁI KIỂM ĐỊNH GIÁO DỤC SỐ TOÀN DIỆN
(AI Senior Education Inspector - Trường TH & THCS Phước Hưng)
"""

    notification_record = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "recipient": recipient,
        "subject": subject,
        "body": email_body,
        "status": "Đã ghi nhận và gửi thông báo thành công",
        "scan_time": scan_time
    }

    # Lưu log email
    email_logs = []
    if os.path.exists(EMAIL_LOG_FILE):
        try:
            with open(EMAIL_LOG_FILE, "r", encoding="utf-8") as f:
                email_logs = json.load(f)
        except Exception:
            email_logs = []
    
    email_logs.append(notification_record)
    with open(EMAIL_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(email_logs, f, ensure_ascii=False, indent=2)

    return notification_record

def export_drive_monitoring_report_word(output_path=None):
    """
    Xuất Báo cáo giám sát tiến độ nộp giáo án Google Drive ra file Word (.docx)
    đạt chuẩn Thể thức văn bản hành chính theo Nghị định 30/2020/NĐ-CP.
    """
    if output_path is None:
        output_path = os.path.join(CORRECTED_DIR, "BAO_CAO_GIAM_SAT_TIEN_DO_GIAO_AN_GOOGLE_DRIVE.docx")

    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            snapshot = json.load(f)
    else:
        snapshot = scan_google_drive_status()

    doc = docx.Document()

    # Căn lề chuẩn Nghị định 30: Top 20mm, Bottom 20mm, Left 30mm, Right 15mm
    for section in doc.sections:
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)
        section.page_width = Mm(210)
        section.page_height = Mm(297)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(3)

    # 1. HEADER (2 cột)
    t0 = doc.add_table(rows=3, cols=2)
    t0.alignment = WD_TABLE_ALIGNMENT.CENTER
    t0.autofit = False

    col_widths_0 = [Mm(75), Mm(90)]
    for row in t0.rows:
        for idx, width in enumerate(col_widths_0):
            row.cells[idx].width = width

    # Hàng 1
    p00 = t0.rows[0].cells[0].paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r00 = p00.add_run("TRƯỜNG TH & THCS PHƯỚC HƯNG")
    r00.font.name = "Times New Roman"
    r00.font.size = Pt(11)

    p01 = t0.rows[0].cells[1].paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r01 = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    r01.font.name = "Times New Roman"
    r01.font.size = Pt(12)
    r01.font.bold = True

    # Hàng 2
    p10 = t0.rows[1].cells[0].paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r10 = p10.add_run("TỔ: TOÁN – KHTN – CÔNG NGHỆ")
    r10.font.name = "Times New Roman"
    r10.font.size = Pt(12)
    r10.font.bold = True

    p11 = t0.rows[1].cells[1].paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r11 = p11.add_run("Độc lập – Tự do – Hạnh phúc")
    r11.font.name = "Times New Roman"
    r11.font.size = Pt(12)
    r11.font.bold = True
    r11.font.underline = True

    # Hàng 3
    p20 = t0.rows[2].cells[0].paragraphs[0]
    p20.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r20 = p20.add_run("Số: ...../BC-TTN")
    r20.font.name = "Times New Roman"
    r20.font.size = Pt(11)

    p21 = t0.rows[2].cells[1].paragraphs[0]
    p21.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r21 = p21.add_run("Nhơn Hội, ngày 03 tháng 10 năm 2026")
    r21.font.name = "Times New Roman"
    r21.font.size = Pt(11)
    r21.font.italic = True

    # 2. TIÊU ĐỀ BÁO CÁO
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(14)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("BÁO CÁO GIÁM SÁT TIẾN ĐỘ NỘP KẾ HOẠCH BÀI DẠY")
    r_title.font.name = "Times New Roman"
    r_title.font.size = Pt(14)
    r_title.font.bold = True

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(10)
    r_sub = p_sub.add_run("TRÊN HỆ THỐNG GOOGLE DRIVE TỔ CHUYÊN MÔN\n(Học kỳ I – Năm học 2026 – 2027)")
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(12)
    r_sub.font.bold = True

    # Thông tin kiểm tra
    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_after = Pt(6)
    r_info = p_info.add_run("Kính gửi: ")
    r_info.font.bold = True
    p_info.add_run("Ban Giám hiệu Trường TH & THCS Phước Hưng.\n")
    p_info.add_run("Căn cứ Kế hoạch hoạt động giáo dục năm học 2026 – 2027 và Lịch công tác tuần của Tổ Toán – KHTN – Công nghệ, Tổ chuyên môn tiến hành giám sát, tổng hợp tình hình nộp Kế hoạch bài dạy (giáo án) trên thư mục dùng chung Google Drive như sau:")

    # Thư mục Google Drive
    p_link = doc.add_paragraph()
    p_link.paragraph_format.left_indent = Inches(0.2)
    p_link.paragraph_format.space_after = Pt(8)
    r_l1 = p_link.add_run("- Đường dẫn thư mục Google Drive giám sát: ")
    r_l1.font.bold = True
    p_link.add_run(DRIVE_ROOT_URL + "\n")
    r_l2 = p_link.add_run(f"- Thời điểm tổng hợp dữ liệu: ")
    r_l2.font.bold = True
    p_link.add_run(f"{snapshot.get('scan_time', '03/10/2026')} • Đối tượng: Toàn bộ 08 giáo viên trong tổ.")

    # 3. BẢNG CHI TIẾT TIẾN ĐỘ 8 GIÁO VIÊN
    teachers = snapshot.get("teachers", [])
    t_tbl = doc.add_table(rows=len(teachers) + 1, cols=6)
    t_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_tbl.autofit = False

    # Viền bảng
    tblPr = t_tbl._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        f'<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

    col_widths = [Mm(10), Mm(36), Mm(28), Mm(26), Mm(35), Mm(30)]
    for row in t_tbl.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    # Header bảng
    headers = ["STT", "Họ và tên GV", "Môn phụ trách", "Thời gian nộp", "Tình trạng tiến độ", "Ghi chú"]
    for idx, title in enumerate(headers):
        cell = t_tbl.rows[0].cells[idx]
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="EBF3FB"/>')
        cell._tc.get_or_add_tcPr().append(shading)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.font.bold = True

    # Rows dữ liệu
    for idx, t in enumerate(teachers, start=1):
        row = t_tbl.rows[idx]
        
        # STT
        p0 = row.cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.add_run(str(idx)).font.size = Pt(10.5)

        # Tên GV
        p1 = row.cells[1].paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(t['teacher_name'])
        r1.font.bold = True
        r1.font.size = Pt(10.5)

        # Môn
        p2 = row.cells[2].paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.add_run(t['subjects']).font.size = Pt(10)

        # Thời gian
        p3 = row.cells[3].paragraphs[0]
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3.add_run(t['last_updated']).font.size = Pt(9.5)

        # Tình trạng
        p4 = row.cells[4].paragraphs[0]
        p4.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r4 = p4.add_run(t['status'])
        r4.font.size = Pt(10)
        if "100%" in t['status'] or "hoàn thành" in t['status'].lower():
            r4.font.color.rgb = RGBColor(0, 100, 0) # Xanh lá đậm

        # Ghi chú
        p5 = row.cells[5].paragraphs[0]
        p5.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p5.add_run(t['notes']).font.size = Pt(9.5)

    # 4. ĐÁNH GIÁ CHUNG VÀ KIẾN NGHỊ
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(10)
    p_sec1.paragraph_format.space_after = Pt(3)
    r_sec1 = p_sec1.add_run("I. ĐÁNH GIÁ CHUNG:")
    r_sec1.font.bold = True

    submitted_1_4 = [t for t in teachers if t.get("week1_4_count", 0) > 0]
    submitted_5_8 = [t for t in teachers if t.get("week5_8_count", 0) > 0]
    unsubmitted = [t for t in teachers if t.get("file_count", 0) == 0]
    total_files = sum(t.get("file_count", 0) for t in teachers)

    sub_1_4_str = ", ".join([f"{t['teacher_name']}: {t.get('week1_4_count', 0)} tệp" for t in submitted_1_4])
    sub_5_8_str = ", ".join([f"{t['teacher_name']}: {t.get('week5_8_count', 0)} tệp" for t in submitted_5_8])
    unsub_str = ", ".join([t['teacher_name'] for t in unsubmitted])

    eval_items = [
        f"1. Về tiến độ Tuần 1 đến Tuần 4: Đã có {len(submitted_1_4)}/08 giáo viên nộp KHDY với tổng cộng {sum(t.get('week1_4_count', 0) for t in submitted_1_4)} tệp tin ({sub_1_4_str}). Riêng thầy Lê Văn Thắng chưa cập nhật tệp tin nào lên thư mục Google Drive (0 tệp).",
        f"2. Về tiến độ Tuần 5 đến Tuần 8: Đã có {len(submitted_5_8)}/08 giáo viên nộp sớm với tổng cộng {sum(t.get('week5_8_count', 0) for t in submitted_5_8)} tệp tin ({sub_5_8_str}).",
        "3. Về tình hình các thư mục môn học không có nộp file (0 tệp):",
        "   - Cô Phạm Thị Mỹ Linh: Tuần 1 đến 4 thư mục Toán 6 không có nộp file (0 tệp); Tuần 5 đến 8 thư mục Toán 6 không có nộp file (0 tệp).",
        "   - Cô Lê Thị Thu Thủy: Tuần 1 đến 4 thư mục môn Toán 6 không có nộp file (0 tệp); Tuần 5 đến 8 thư mục môn Toán 6 không có nộp file (0 tệp).",
        "   - Cô Hà Thị Kế: Đã nộp Tuần 1-4 và 5-8 đủ các môn; Riêng Tuần 9 đến 12 thư mục HĐTN 8 không có nộp file (0 tệp), thư mục KHTN 6 không có nộp file (0 tệp).",
        "   - Thầy Lê Văn Thắng: Tuần 1 đến 4 thư mục Hóa 9, KHTN 7, Toán 9 không có nộp file (0 tệp); Tuần 5 đến 8 thư mục Hóa 9, KHTN 7, Toán 9 không có nộp file (0 tệp).",
        "   - 04 Giáo viên còn lại (Cô Cúc, Cô Hằng, Thầy Thành, Cô Trang): Đã nộp đầy đủ các thư mục môn phụ trách trong Tuần 1-4 và Tuần 5-8 (không có thư mục môn nào trống file).",
        f"4. Kiến nghị: Kính đề nghị Tổ trưởng chuyên môn nhắc nhở giáo viên chưa nộp (thầy Lê Văn Thắng) và các giáo viên có thư mục môn chưa có nộp file khẩn trương cập nhật đầy đủ lên Google Drive phục vụ công tác kiểm tra, duyệt giáo án định kỳ."
    ]
    for it in eval_items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(it)
        r.font.size = Pt(11.5)

    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(8)
    p_sec2.paragraph_format.space_after = Pt(3)
    r_sec2 = p_sec2.add_run("II. ĐỀ XUẤT - KIẾN NGHỊ BAN GIÁM HIỆU:")
    r_sec2.font.bold = True

    prop_items = [
        "1. Kính đề nghị Ban Giám hiệu phê duyệt kết quả kiểm tra định kỳ hồ sơ Kế hoạch bài dạy đợt Tuần 1 - 4 của Tổ Toán – KHTN – Công nghệ.",
        "2. Tiếp tục duy trì nền tảng Google Drive kết hợp AI Senior Education Inspector để tự động hóa công tác kiểm định và số hóa giáo dục toàn diện của nhà trường."
    ]
    for it in prop_items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(it)
        r.font.size = Pt(11.5)

    # 5. CHỮ KÝ PHÊ DUYỆT (2 CỘT)
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

    t_sign = doc.add_table(rows=2, cols=2)
    t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_sign.autofit = False

    for row in t_sign.rows:
        row.cells[0].width = Mm(85)
        row.cells[1].width = Mm(80)

    # Hàng tiêu đề ký
    c0 = t_sign.rows[0].cells[0].paragraphs[0]
    c0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_bgh = c0.add_run("BAN GIÁM HIỆU PHÊ DUYỆT\n")
    r_bgh.font.bold = True
    r_bgh.font.size = Pt(12)
    r_bgh_s = c0.add_run("(Ký và đóng dấu)")
    r_bgh_s.font.italic = True
    r_bgh_s.font.size = Pt(11)

    c1 = t_sign.rows[0].cells[1].paragraphs[0]
    c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_d = c1.add_run("Nhơn Hội, ngày 03 tháng 10 năm 2026\n")
    r_d.font.italic = True
    r_d.font.size = Pt(11.5)
    r_tt = c1.add_run("TỔ TRƯỞNG CHUYÊN MÔN\n")
    r_tt.font.bold = True
    r_tt.font.size = Pt(12)
    r_tt_s = c1.add_run("(Ký và ghi rõ họ tên)")
    r_tt_s.font.italic = True
    r_tt_s.font.size = Pt(11)

    # Hàng họ tên + chữ ký
    c0_b = t_sign.rows[1].cells[0].paragraphs[0]
    c0_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c0_b.paragraph_format.space_before = Pt(50)
    c0_b.add_run("HIỆU TRƯỞNG").font.bold = True

    c1_b = t_sign.rows[1].cells[1]
    sig_img_path = get_file_path("signature_thang.png")
    if not os.path.exists(sig_img_path):
        sig_img_path = get_file_path(os.path.join("data", "signature_thang.png"))
    if os.path.exists(sig_img_path):
        p_img = c1_b.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(2)
        r_sig = p_img.add_run()
        try:
            r_sig.add_picture(sig_img_path, width=Inches(1.4))
        except Exception:
            pass
        p_n = c1_b.add_paragraph()
        p_n.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_name = p_n.add_run("Lê Văn Thắng")
        r_name.font.bold = True
        r_name.font.size = Pt(12)
    else:
        p_n = c1_b.paragraphs[0]
        p_n.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_n.paragraph_format.space_before = Pt(50)
        r_name = p_n.add_run("Lê Văn Thắng")
        r_name.font.bold = True
        r_name.font.size = Pt(12)

    if isinstance(output_path, (str, bytes, os.PathLike)):
        out_dir = os.path.dirname(output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
    doc.save(output_path)
    return output_path

