import os
import sys
import json
import logging
from datetime import datetime

# Ensure user site-packages are accessible
site_pkg = r"C:\Users\ADMIN\AppData\Roaming\Python\Python313\site-packages"
if os.path.exists(site_pkg) and site_pkg not in sys.path:
    sys.path.append(site_pkg)

try:
    from supabase import create_client, Client
except ImportError:
    Client = None
    create_client = None

logger = logging.getLogger("supabase_client")

# Helper to load .env manually if python-dotenv is not installed
def load_env_file():
    env_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), ".env")
    if not os.path.exists(env_path):
        env_path = r"D:\APP-KIEM-TRA-GIAO-AN\.env"
    
    if os.path.exists(env_path):
        try:
            with open(env_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        os.environ.setdefault(k.strip(), v.strip())
        except Exception as e:
            logger.warning(f"Error loading .env file: {e}")

load_env_file()

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://creselczgrzorustszzc.supabase.co")
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")

_supabase_client = None

def get_supabase_client() -> Client:
    global _supabase_client
    if _supabase_client is not None:
        return _supabase_client
    
    if not create_client or not SUPABASE_URL or not (SUPABASE_SERVICE_ROLE_KEY or SUPABASE_ANON_KEY):
        return None

    key = SUPABASE_SERVICE_ROLE_KEY if SUPABASE_SERVICE_ROLE_KEY else SUPABASE_ANON_KEY
    try:
        _supabase_client = create_client(SUPABASE_URL, key)
        return _supabase_client
    except Exception as e:
        logger.error(f"Failed to initialize Supabase client: {e}")
        return None

def get_supabase_status():
    """Kiểm tra trạng thái kết nối và tình trạng các bảng trong Supabase Cloud"""
    client = get_supabase_client()
    if not client:
        return {
            "status": "error",
            "connected": False,
            "url": SUPABASE_URL,
            "message": "Không thể khởi tạo thư viện Supabase client",
            "tables": {}
        }
    
    tables_to_check = ["notebook_records", "teachers", "drive_logs", "inspection_records"]
    table_status = {}
    connected = True
    
    for tbl in tables_to_check:
        try:
            res = client.table(tbl).select("id", count="exact").limit(1).execute()
            count = res.count if hasattr(res, 'count') and res.count is not None else len(res.data)
            table_status[tbl] = {"exists": True, "count": count}
        except Exception as e:
            err_msg = str(e)
            if "PGRST205" in err_msg or "Could not find the table" in err_msg:
                table_status[tbl] = {"exists": False, "count": 0, "error": "Bảng chưa được tạo trong Supabase SQL"}
            else:
                table_status[tbl] = {"exists": False, "count": 0, "error": err_msg}

    all_exist = all(t["exists"] for t in table_status.values())
    
    return {
        "status": "connected" if connected else "error",
        "connected": connected,
        "url": SUPABASE_URL,
        "tables_ready": all_exist,
        "message": "Đã kết nối thành công tới Supabase Cloud!" if all_exist else "Kết nối thành công nhưng cần khởi tạo bảng SQL trong Supabase",
        "tables": table_status,
        "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

def sync_notebook_to_supabase(notebook_data: dict = None):
    """Đồng bộ Sổ tay tổ trưởng (so_tay_to_truong.json) lên Supabase table `notebook_records`"""
    client = get_supabase_client()
    if not client:
        return {"status": "error", "message": "Supabase client chưa sẵn sàng"}
    
    data_file = r"D:\APP-KIEM-TRA-GIAO-AN\data\so_tay_to_truong.json"
    if notebook_data is None:
        if not os.path.exists(data_file):
            return {"status": "error", "message": "Không tìm thấy file dữ liệu so_tay_to_truong.json"}
        try:
            with open(data_file, "r", encoding="utf-8") as f:
                notebook_data = json.load(f)
        except Exception as e:
            return {"status": "error", "message": f"Lỗi đọc file JSON: {e}"}
            
    rows_to_upsert = []
    for week_str, records in notebook_data.items():
        try:
            week_num = int(str(week_str).replace("week_", "").replace("tuan_", "").strip())
        except ValueError:
            continue
            
        for rec in records:
            rows_to_upsert.append({
                "week": week_num,
                "teacher_id": rec.get("teacher_id", ""),
                "teacher_name": rec.get("teacher_name", ""),
                "so_tiet_ngay": rec.get("so_tiet_ngay", ""),
                "so_tiet_tuan": rec.get("so_tiet_tuan", 0),
                "giao_an_drive": rec.get("giao_an_drive", "Chưa nộp"),
                "dddh": rec.get("dddh", 0),
                "cntt": rec.get("cntt", 0),
                "nhan_xet": rec.get("nhan_xet", ""),
                "xep_loai": rec.get("xep_loai", "Tốt"),
                "updated_at": datetime.now().isoformat()
            })
            
    if not rows_to_upsert:
        return {"status": "success", "synced_count": 0, "message": "Không có bản ghi nào cần đồng bộ"}

    try:
        res = client.table("notebook_records").upsert(rows_to_upsert, on_conflict="week,teacher_id").execute()
        return {
            "status": "success",
            "synced_count": len(rows_to_upsert),
            "message": f"Đã đồng bộ {len(rows_to_upsert)} bản ghi Sổ tay lên Supabase Cloud"
        }
    except Exception as e:
        logger.error(f"Error upserting notebook_records to Supabase: {e}")
        return {"status": "error", "message": f"Lỗi đồng bộ Supabase: {e}"}

def fetch_notebook_from_supabase(week: int = None):
    """Lấy dữ liệu Sổ tay từ Supabase Cloud `notebook_records`"""
    client = get_supabase_client()
    if not client:
        return None

    try:
        query = client.table("notebook_records").select("*")
        if week is not None:
            query = query.eq("week", week)
        res = query.execute()
        return res.data
    except Exception as e:
        logger.error(f"Error fetching notebook_records from Supabase: {e}")
        return None

def sync_teachers_to_supabase(teachers_list: list = None):
    """Đồng bộ Danh sách giáo viên lên Supabase table `teachers`"""
    client = get_supabase_client()
    if not client:
        return {"status": "error", "message": "Supabase client chưa sẵn sàng"}

    if teachers_list is None:
        from app.services.notebook_monthly_report import TEACHER_LIST
        teachers_list = TEACHER_LIST

    rows = []
    for t in teachers_list:
        rows.append({
            "id": t.get("id"),
            "name": t.get("name"),
            "role": t.get("role", "Giáo viên"),
            "subject_str": t.get("subject_str", ""),
            "grades": t.get("grades", ""),
            "classes": t.get("classes", []),
            "homeroom": t.get("homeroom", ""),
            "default_cntt": t.get("default_cntt", 0),
            "updated_at": datetime.now().isoformat()
        })

    try:
        res = client.table("teachers").upsert(rows, on_conflict="id").execute()
        return {
            "status": "success",
            "synced_count": len(rows),
            "message": f"Đã đồng bộ {len(rows)} giáo viên lên Supabase Cloud"
        }
    except Exception as e:
        logger.error(f"Error upserting teachers to Supabase: {e}")
        return {"status": "error", "message": f"Lỗi đồng bộ Giáo viên: {e}"}

def sync_drive_logs_to_supabase(log_data: dict = None):
    """Đồng bộ log giám sát Google Drive lên Supabase table `drive_logs`"""
    client = get_supabase_client()
    if not client:
        return {"status": "error", "message": "Supabase client chưa sẵn sàng"}

    data_file = r"D:\APP-KIEM-TRA-GIAO-AN\data\drive_monitoring_log.json"
    if log_data is None and os.path.exists(data_file):
        try:
            with open(data_file, "r", encoding="utf-8") as f:
                log_data = json.load(f)
        except Exception:
            pass

    if not log_data:
        return {"status": "warning", "message": "Chưa có dữ liệu log Drive để đồng bộ"}

    row = {
        "scanned_at": log_data.get("scanned_at", datetime.now().isoformat()),
        "total_teachers": len(log_data.get("teachers", [])),
        "teachers_data": log_data.get("teachers", [])
    }

    try:
        res = client.table("drive_logs").insert(row).execute()
        return {
            "status": "success",
            "message": "Đã lưu nhật ký giám sát Drive lên Supabase Cloud"
        }
    except Exception as e:
        logger.error(f"Error inserting drive_logs to Supabase: {e}")
        return {"status": "error", "message": f"Lỗi đồng bộ Drive log: {e}"}

def sync_all_to_supabase():
    """Đồng bộ toàn bộ dữ liệu hệ thống (Sổ tay, Giáo viên, Nhật ký Drive) lên Supabase Cloud"""
    status = get_supabase_status()
    if not status.get("connected"):
        return {"status": "error", "message": "Không thể kết nối Supabase Cloud"}

    results = {
        "teachers": sync_teachers_to_supabase(),
        "notebook": sync_notebook_to_supabase(),
        "drive_logs": sync_drive_logs_to_supabase(),
        "synced_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    success_count = sum(1 for r in results.values() if isinstance(r, dict) and r.get("status") == "success")
    results["overall_status"] = "success" if success_count > 0 else "warning"
    results["summary_message"] = f"Hoàn tất đồng bộ dữ liệu tới Supabase Cloud. ({success_count}/3 tác vụ thành công)"
    return results
