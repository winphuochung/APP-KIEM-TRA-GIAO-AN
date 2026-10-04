import urllib.request
import re
import json
import os
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

from app.config import DATA_DIR, resolve_data_file

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}

TEACHER_DRIVE_IDS = {
    "cuc": {"name": "Phạm Thị Cúc", "folder_name": "Cô Cúc", "drive_id": "1Y57MXplWyQitPENHEQO4McKd7hlLHBtV", "subjects": ["KHTN 7", "Sinh 9"]},
    "hang": {"name": "Lê Thị Thúy Hằng", "folder_name": "Cô Hằng", "drive_id": "1jfXZGd-nH9kgVRsJDg_ldIUFdbR4-MUI", "subjects": ["CN 6", "CN 7", "HĐTN 6", "HĐTN 7", "HĐTN 9"]},
    "ke": {"name": "Hà Thị Kế", "folder_name": "Cô Kế", "drive_id": "10aAFjxR95vKnmLzyvcU7OivbD3GTP88r", "subjects": ["KHTN 6", "KHTN 8", "HĐTN 8"]},
    "linh": {"name": "Phạm Thị Mỹ Linh", "folder_name": "Cô Linh", "drive_id": "1slzRqoYLd7pkSfeK3sxDBwzS_5fYA7vd", "subjects": ["Toán 8", "Toán 9"]},
    "thuy": {"name": "Lê Thị Thu Thủy", "folder_name": "Cô Thủy", "drive_id": "1vRyDMAB4lLyxDyPgYsit2VdO7ZcPyuNi", "subjects": ["Toán 6", "Toán 7", "Toán 8"]},
    "trang": {"name": "Nguyễn Thị Thùy Trang", "folder_name": "Cô Trang", "drive_id": "1OwIYYK7ENQ4f3aXwS6RrnuA8Mgl7_qKj", "subjects": ["CN 8", "CN 9"]},
    "thanh": {"name": "Nguyễn Chí Thành", "folder_name": "Thầy Thành", "drive_id": "1_hZDldFUOUvjJnkw9-Dm8I3kJbUCc5Hr", "subjects": ["KHTN 6", "KHTN 8", "KHTN 9", "HĐTN 8"]},
    "thang": {"name": "Lê Văn Thắng", "folder_name": "Thầy Thắng (Tổ trưởng)", "drive_id": "1S1wipWDzCGRaEonbFDFClIzmui1yEXZJ", "subjects": ["Hóa 9", "KHTN 7", "Toán 9"]}
}


def scrape_drive_folder_items(folder_id):
    """
    Quét trực tiếp nội dung thư mục công khai Google Drive qua Embedded View Engine.
    Trả về danh sách dict {id, name, is_folder}.
    """
    url = f"https://drive.google.com/embeddedfolderview?id={folder_id}#list"
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
            
            # Trích xuất danh sách các thẻ a chứa link và tiêu đề
            pattern = r'href="https://drive\.google\.com/(?:drive/folders/|file/d/|open\?id=)([a-zA-Z0-9_-]+)"[^>]*>(.*?)</a>'
            matches = re.findall(pattern, html, re.DOTALL)
            items = []
            
            # Trích xuất danh sách title div dành cho tệp tin
            title_divs = re.findall(r'<div[^>]*class="[^"]*title[^"]*"[^>]*>(.*?)</div>', html, re.DOTALL)
            clean_titles = [re.sub(r'<[^>]+>', '', t).strip() for t in title_divs if re.sub(r'<[^>]+>', '', t).strip() and re.sub(r'<[^>]+>', '', t).strip() != "TITLE"]

            for item_id, title_raw in matches:
                clean_title = re.sub(r'<[^>]+>', '', title_raw).strip()
                if clean_title:
                    is_folder = f"drive/folders/{item_id}" in html
                    items.append({"id": item_id, "name": clean_title, "is_folder": is_folder})

            for ct in clean_titles:
                if any(ct.endswith(ext) for ext in ['.docx', '.pdf', '.doc', '.xlsx']):
                    if not any(it["name"] == ct for it in items):
                        items.append({"id": "file_" + str(hash(ct)), "name": ct, "is_folder": False})

            return items
    except Exception:
        return []


def process_single_teacher_sync(t_id, meta, now_str):
    """Quét dữ liệu Google Drive của 1 giáo viên theo thời gian thực."""
    drive_id = meta["drive_id"]
    period_folders = scrape_drive_folder_items(drive_id)
    
    cycles_data = {}
    teacher_files_count = 0

    for pf in period_folders:
        p_name = pf["name"]
        p_id = pf["id"]
        
        sub_items = scrape_drive_folder_items(p_id) if pf["is_folder"] else []
        files_list = []
        subjs_detail = {}

        for item in sub_items:
            if item["is_folder"]:
                subj_name = item["name"]
                subj_files_items = scrape_drive_folder_items(item["id"])
                s_files = [sf["name"] for sf in subj_files_items if not sf["is_folder"]]
                subjs_detail[subj_name] = {
                    "folder_id": item["id"],
                    "count": len(s_files),
                    "files": s_files
                }
                files_list.extend(s_files)
            else:
                files_list.append(item["name"])

        if files_list and not subjs_detail:
            for fname in files_list:
                matched_subj = meta["subjects"][0]
                for sb in meta["subjects"]:
                    if sb.lower() in fname.lower():
                        matched_subj = sb
                        break

                s_entry = subjs_detail.setdefault(matched_subj, {
                    "folder_id": p_id,
                    "count": 0,
                    "files": []
                })
                if fname not in s_entry["files"]:
                    s_entry["files"].append(fname)
                s_entry["count"] = len(s_entry["files"])

        f_count = len(files_list)
        teacher_files_count += f_count

        cycles_data[p_name] = {
            "status": f"ĐÃ ĐỒNG BỘ AUTO DRIVE ({f_count} tệp)" if f_count > 0 else "CHƯA NỘP",
            "file_count": f_count,
            "subjects_detail": subjs_detail
        }

    return t_id, {
        "teacher": {
            "id": t_id,
            "name": meta["name"],
            "folder_name": meta["folder_name"],
            "drive_id": drive_id,
            "subjects": meta["subjects"],
            "subjects_str": ", ".join(meta["subjects"])
        },
        "total_files": teacher_files_count,
        "last_updated": now_str,
        "cycles": cycles_data
    }, teacher_files_count


def perform_zero_config_drive_sync():
    """
    Quét song song (Multithreaded) toàn bộ cây thư mục Google Drive 8 giáo viên
    không cần API Key hay Service Account và đồng bộ dữ liệu vào hệ thống.
    """
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    report_path = resolve_data_file("drive_live_exact_report.json")
    
    live_report = {}
    if os.path.exists(report_path):
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                live_report = json.load(f)
        except Exception:
            live_report = {}

    total_files_scanned = 0
    scanned_teachers_count = 0

    # Quét song song 8 giáo viên cùng lúc bằng ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = [
            executor.submit(process_single_teacher_sync, t_id, meta, now_str)
            for t_id, meta in TEACHER_DRIVE_IDS.items()
        ]
        
        for future in as_completed(futures):
            try:
                t_id, t_record, f_count = future.result()
                scanned_teachers_count += 1
                total_files_scanned += f_count
                
                # Cập nhật thông tin nếu có tệp tin quét được
                if t_record["total_files"] > 0:
                    live_report[t_id] = t_record
                else:
                    existing = live_report.get(t_id, {})
                    existing["last_updated"] = now_str
                    live_report[t_id] = existing
            except Exception:
                pass

    # Lưu lại kết quả vào drive_live_exact_report.json
    target_path = os.path.join(DATA_DIR, "drive_live_exact_report.json")
    with open(target_path, "w", encoding="utf-8") as f:
        json.dump(live_report, f, ensure_ascii=False, indent=2)

    return {
        "status": "success",
        "message": f"Đã tự động quét và đồng bộ trực tiếp dữ liệu Google Drive của {scanned_teachers_count} giáo viên!",
        "last_updated": now_str,
        "scanned_teachers": scanned_teachers_count,
        "total_files_scanned": total_files_scanned
    }
