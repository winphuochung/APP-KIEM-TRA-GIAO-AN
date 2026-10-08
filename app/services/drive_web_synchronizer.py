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


def get_all_files_recursive(folder_id, max_depth=3):
    """Lấy tất cả các tệp đệ quy trong 1 thư mục bất kể số cấp lồng nhau."""
    if max_depth <= 0:
        return []
    items = scrape_drive_folder_items(folder_id)
    files = []
    for it in items:
        if it.get("is_folder", False):
            files.extend(get_all_files_recursive(it["id"], max_depth - 1))
        else:
            if any(it.get("name", "").endswith(ext) for ext in ['.docx', '.pdf', '.doc', '.xlsx']):
                files.append(it["name"])
    return files


def process_single_teacher_sync(t_id, meta, now_str):
    """Quét dữ liệu Google Drive của 1 giáo viên theo thời gian thực (đệ quy sâu)."""
    drive_id = meta["drive_id"]
    period_folders = scrape_drive_folder_items(drive_id)
    
    cycles_data = {}
    teacher_files_count = 0

    root_files = [pf["name"] for pf in period_folders if not pf.get("is_folder", False) and any(pf.get("name", "").endswith(ext) for ext in ['.docx', '.pdf', '.doc', '.xlsx'])]

    for pf in period_folders:
        if not pf.get("is_folder", False):
            continue
        p_name = pf["name"]
        p_id = pf["id"]
        
        sub_items = scrape_drive_folder_items(p_id)
        files_list = []
        subjs_detail = {}

        for item in sub_items:
            if item["is_folder"]:
                subj_name = item["name"]
                s_files = get_all_files_recursive(item["id"], max_depth=3)
                subjs_detail[subj_name] = {
                    "folder_id": item["id"],
                    "count": len(s_files),
                    "files": s_files
                }
                files_list.extend(s_files)
            else:
                if any(item["name"].endswith(ext) for ext in ['.docx', '.pdf', '.doc', '.xlsx']):
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
            "status": f"ĐÃ NỘP ({f_count} tệp)" if f_count > 0 else "CHƯA NỘP",
            "file_count": f_count,
            "subjects_detail": subjs_detail
        }

    if root_files:
        teacher_files_count += len(root_files)

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
        "cycles": cycles_data,
        "root_files": root_files
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
                live_report[t_id] = t_record
            except Exception:
                pass

    # 1. Lưu lại kết quả vào drive_live_exact_report.json
    target_path = os.path.join(DATA_DIR, "drive_live_exact_report.json")
    with open(target_path, "w", encoding="utf-8") as f:
        json.dump(live_report, f, ensure_ascii=False, indent=2)

    # 2. Cập nhật drive_full_cycles_report.json
    try:
        from app.services.cycle_report_generator import generate_fallback_9_cycles_matrix, generate_9_cycles_monitoring_word
        full_cycles = generate_fallback_9_cycles_matrix()
        full_cycles_path = os.path.join(DATA_DIR, "drive_full_cycles_report.json")
        with open(full_cycles_path, "w", encoding="utf-8") as f:
            json.dump(full_cycles, f, ensure_ascii=False, indent=2)
        generate_9_cycles_monitoring_word()
    except Exception as e:
        print(f"Error updating full cycles report: {e}")

    # 3. Cập nhật drive_monitoring_log.json
    try:
        from app.services.drive_monitor import scan_google_drive_status, export_drive_monitoring_report_word
        scan_google_drive_status()
        export_drive_monitoring_report_word()
    except Exception as e:
        print(f"Error updating monitoring log: {e}")

    return {
        "status": "success",
        "message": f"Đã tự động quét và đồng bộ trực tiếp dữ liệu Google Drive của {scanned_teachers_count} giáo viên!",
        "last_updated": now_str,
        "scanned_teachers": scanned_teachers_count,
        "total_files_scanned": total_files_scanned
    }

