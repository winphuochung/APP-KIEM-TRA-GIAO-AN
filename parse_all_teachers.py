import re
import json

teacher_folders = {
    "cuc": "Cô Cúc",
    "hang": "Cô Hằng",
    "ke": "Cô Kế",
    "linh": "Cô Linh",
    "thuy": "Cô Thủy",
    "trang": "Cô Trang",
    "thanh": "Thầy Thành",
    "thang": "Thầy Thắng",
}

results = {}

for key, name in teacher_folders.items():
    with open(f"drive_{key}.html", "r", encoding="utf-8") as f:
        html = f.read()

    # Extract all aria-labels
    labels = re.findall(r'aria-label="([^"]+)"', html)
    
    # Filter labels to find folders, files, dates
    items = []
    for l in labels:
        if any(ignore in l for ignore in [
            "Tìm hiểu thêm", "Loại bỏ", "Trình đơn", "Quay lại", "Đóng", "Drive", 
            "Các ứng dụng", "Đăng nhập", "Danh sách", "Sắp xếp", "Điều khiển", 
            "Chiều sắp", "Thư mục", "circular", "Gói thành viên"
        ]):
            continue
        items.append(l)

    results[key] = {
        "teacher": name,
        "items": items
    }

with open("all_teachers_parsed.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("Saved all_teachers_parsed.json")
