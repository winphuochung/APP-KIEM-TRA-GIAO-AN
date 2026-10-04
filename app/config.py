import os
import sys
import tempfile

# Base directory of the project
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Ensure user site-packages are loaded if available (for local Windows setup)
site_pkg = r"C:\Users\ADMIN\AppData\Roaming\Python\Python313\site-packages"
if os.path.exists(site_pkg) and site_pkg not in sys.path:
    sys.path.append(site_pkg)

# Get writable directory (falls back to system temp on read-only environments like Vercel Serverless)
def get_writable_dir(*subdirs):
    target_dir = os.path.join(BASE_DIR, *subdirs)
    try:
        os.makedirs(target_dir, exist_ok=True)
        test_file = os.path.join(target_dir, ".write_test")
        with open(test_file, "w") as f:
            f.write("test")
        os.remove(test_file)
        return target_dir
    except Exception:
        tmp_target = os.path.join(tempfile.gettempdir(), "app_kiem_tra_giao_an", *subdirs)
        os.makedirs(tmp_target, exist_ok=True)
        return tmp_target

# Safe helper to locate input files or templates
def get_file_path(relative_path):
    # Thử tìm từ BASE_DIR
    p1 = os.path.join(BASE_DIR, relative_path)
    if os.path.exists(p1):
        return p1
    # Thử tìm từ current directory
    p2 = os.path.abspath(relative_path)
    if os.path.exists(p2):
        return p2
    return p1

# Data directory
DATA_DIR = get_writable_dir("data")
UPLOAD_DIR = get_writable_dir("data", "uploads")
CORRECTED_DIR = get_writable_dir("data", "corrected")

def resolve_data_file(filename):
    """Tìm đường dẫn đọc file dữ liệu bằng cách rà soát DATA_DIR, get_file_path và BASE_DIR/data."""
    candidates = [
        os.path.join(DATA_DIR, filename),
        get_file_path(os.path.join("data", filename)),
        os.path.join(BASE_DIR, "data", filename)
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return candidates[0]

