import os
import json
import time
import urllib.request
import urllib.parse
from datetime import datetime
import jwt

from app.config import DATA_DIR, BASE_DIR, resolve_data_file

SERVICE_ACCOUNT_FILE = resolve_data_file("service_account_credentials.json")
API_KEY_FILE = resolve_data_file("google_drive_api_key.txt")
TEMPLATE_SA_FILE = os.path.join(DATA_DIR, "service_account_credentials.json.template")

DRIVE_ROOT_ID = "1cdqOhxb05lt7r6cyu3YwPVecvBLcmHoR"

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


def get_api_key():
    """Lấy Google Drive API Key từ biến môi trường hoặc tệp lưu trữ local."""
    api_key = os.environ.get("GOOGLE_DRIVE_API_KEY")
    if not api_key and os.path.exists(API_KEY_FILE):
        try:
            with open(API_KEY_FILE, "r", encoding="utf-8") as f:
                api_key = f.read().strip()
        except Exception:
            pass
    return api_key if api_key else None


def set_api_key(api_key_str):
    """Lưu Google Drive API Key vào tệp cấu hình local."""
    target_path = os.path.join(DATA_DIR, "google_drive_api_key.txt")
    with open(target_path, "w", encoding="utf-8") as f:
        f.write(api_key_str.strip())
    os.environ["GOOGLE_DRIVE_API_KEY"] = api_key_str.strip()
    return target_path


def load_service_account_info():
    """Đọc thông tin Google Service Account JSON."""
    sa_json_env = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if sa_json_env:
        try:
            return json.loads(sa_json_env)
        except Exception:
            pass
            
    if os.path.exists(SERVICE_ACCOUNT_FILE):
        try:
            with open(SERVICE_ACCOUNT_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict) and "private_key" in data and "client_email" in data:
                    return data
        except Exception:
            pass
    return None


def get_service_account_access_token():
    """
    Tự động tạo OAuth2 Access Token từ Service Account credentials (tệp .json)
    sử dụng PyJWT để ký RS256 token và trao đổi tại https://oauth2.googleapis.com/token.
    """
    sa_info = load_service_account_info()
    if not sa_info:
        return None, "Chưa tìm thấy tệp Service Account Credentials (service_account_credentials.json)."

    client_email = sa_info.get("client_email")
    private_key = sa_info.get("private_key")
    if not client_email or not private_key:
        return None, "Tệp Service Account thiếu client_email hoặc private_key."

    now = int(time.time())
    payload = {
        "iss": client_email,
        "scope": "https://www.googleapis.com/auth/drive.readonly",
        "aud": "https://oauth2.googleapis.com/token",
        "iat": now,
        "exp": now + 3600
    }

    try:
        signed_jwt = jwt.encode(payload, private_key, algorithm="RS256")
        token_url = "https://oauth2.googleapis.com/token"
        req_data = urllib.parse.urlencode({
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "assertion": signed_jwt
        }).encode("utf-8")

        req = urllib.request.Request(token_url, data=req_data, headers={"Content-Type": "application/x-www-form-urlencoded"})
        with urllib.request.urlopen(req, timeout=12) as resp:
            body = json.loads(resp.read().decode("utf-8"))
            access_token = body.get("access_token")
            return access_token, "Thành công"
    except Exception as e:
        return None, f"Lỗi tạo Access Token từ Service Account: {str(e)}"


def get_drive_auth_status():
    """Kiểm tra tình trạng kết nối Google Drive API (Service Account hoặc API Key)."""
    sa_info = load_service_account_info()
    api_key = get_api_key()
    
    status = {
        "service_account_configured": bool(sa_info),
        "service_account_email": sa_info.get("client_email") if sa_info else None,
        "api_key_configured": bool(api_key),
        "active_mode": "None"
    }

    if sa_info:
        token, msg = get_service_account_access_token()
        if token:
            status["active_mode"] = "Service Account (OAuth2 Full Read)"
            status["token_status"] = "Valid Token Active"
        else:
            status["token_status"] = f"Error: {msg}"

    elif api_key:
        status["active_mode"] = "API Key (Public Drive Read)"

    return status


def fetch_files_from_google_drive(folder_id):
    """
    Truy vấn trực tiếp danh sách tệp tin bên trong thư mục Google Drive qua API v3.
    """
    access_token, token_err = get_service_account_access_token()
    api_key = get_api_key()

    if not access_token and not api_key:
        return None, "Cần cấu hình Service Account Credentials hoặc API Key để truy vấn Google Drive API."

    query = urllib.parse.quote(f"'{folder_id}' in parents and trashed = false")
    fields = urllib.parse.quote("files(id, name, mimeType, modifiedTime, webViewLink, size)")
    url = f"https://www.googleapis.com/drive/v3/files?q={query}&fields={fields}&pageSize=100"

    headers = {}
    if access_token:
        headers["Authorization"] = f"Bearer {access_token}"
    elif api_key:
        url += f"&key={api_key}"

    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            files = data.get("files", [])
            return files, None
    except Exception as e:
        return None, f"Lỗi kết nối Google Drive API v3: {str(e)}"


def sync_google_drive_api_to_local_cache():
    """
    Quét trực tiếp qua Google Drive API v3 cho tất cả 8 Giáo viên,
    tự động phân loại file theo các chu kỳ tuần và cập nhật dữ liệu JSON.
    """
    auth_status = get_drive_auth_status()
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")

    report_path = resolve_data_file("drive_live_exact_report.json")
    live_report = {}
    if os.path.exists(report_path):
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                live_report = json.load(f)
        except Exception:
            live_report = {}

    api_synced_count = 0
    errors = []

    for t_id, meta in TEACHER_DRIVE_IDS.items():
        drive_id = meta["drive_id"]
        remote_files, err = fetch_files_from_google_drive(drive_id)

        if err:
            errors.append(f"{meta['name']}: {err}")
            continue

        api_synced_count += 1
        t_record = live_report.setdefault(t_id, {
            "teacher": {
                "id": t_id,
                "name": meta["name"],
                "folder_name": meta["folder_name"],
                "drive_id": drive_id,
                "subjects": meta["subjects"],
                "subjects_str": ", ".join(meta["subjects"])
            },
            "total_files": 0,
            "last_updated": now_str,
            "cycles": {}
        })

        t_record["last_updated"] = now_str
        
        # Phân loại tệp tin theo các chu kỳ
        for rf in remote_files:
            fname = rf.get("name", "")
            # nếu là folder con (ví dụ: thư mục "Tuần 1 đến tuần 4")
            if rf.get("mimeType") == "application/vnd.google-apps.folder":
                sub_files, _ = fetch_files_from_google_drive(rf.get("id"))
                if sub_files:
                    period_name = fname
                    cycle_obj = t_record["cycles"].setdefault(period_name, {
                        "status": f"ĐÃ CẬP NHẬT LIVE DRIVE ({len(sub_files)} tệp)",
                        "file_count": len(sub_files),
                        "subjects_detail": {}
                    })
                    cycle_obj["file_count"] = len(sub_files)
                    cycle_obj["status"] = f"ĐÃ CẬP NHẬT LIVE DRIVE ({len(sub_files)} tệp)"

                    for sf in sub_files:
                        sf_name = sf.get("name", "")
                        matched_subj = meta["subjects"][0]
                        for sb in meta["subjects"]:
                            if sb.lower() in sf_name.lower():
                                matched_subj = sb
                                break

                        s_dict = cycle_obj["subjects_detail"].setdefault(matched_subj, {
                            "folder_id": rf.get("id"),
                            "count": 0,
                            "files": []
                        })
                        if sf_name not in s_dict["files"]:
                            s_dict["files"].append(sf_name)
                        s_dict["count"] = len(s_dict["files"])

        # Tính lại tổng tệp
        tot = sum(c.get("file_count", 0) for c in t_record["cycles"].values())
        t_record["total_files"] = tot if tot > 0 else t_record.get("total_files", 0)

    # Lưu lại file drive_live_exact_report.json
    target_live_path = os.path.join(DATA_DIR, "drive_live_exact_report.json")
    with open(target_live_path, "w", encoding="utf-8") as f:
        json.dump(live_report, f, ensure_ascii=False, indent=2)

    return {
        "status": "success" if api_synced_count > 0 else "warning",
        "api_synced_teachers": api_synced_count,
        "auth_status": auth_status,
        "last_updated": now_str,
        "errors": errors
    }


def create_service_account_template_if_missing():
    """Tạo mẫu file service_account_credentials.json.template để hỗ trợ người dùng tạo Service Account."""
    template_content = {
        "type": "service_account",
        "project_id": "ten-du-an-google-cloud-cua-ban",
        "private_key_id": "key-id-1234567890",
        "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC...\n-----END PRIVATE KEY-----\n",
        "client_email": "ten-service-account@ten-du-an.iam.gserviceaccount.com",
        "client_id": "123456789012345678901",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/ten-service-account%40ten-du-an.iam.gserviceaccount.com"
    }

    if not os.path.exists(TEMPLATE_SA_FILE):
        with open(TEMPLATE_SA_FILE, "w", encoding="utf-8") as f:
            json.dump(template_content, f, ensure_ascii=False, indent=2)

create_service_account_template_if_missing()
