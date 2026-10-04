import json
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Dữ liệu mô hình năng lực sư phạm
TEACHER_COMPETENCY_DATA = {
    "teacher_name": "Phạm Thị Cúc",
    "role": "Giáo viên môn KHTN, Tổ phó chuyên môn",
    "school": "Trường TH & THCS Phước Hưng",
    "cycle": "Giai đoạn Tuần 01 - 04, Năm học 2026 - 2027",
    "metrics": {
        "pedagogy_5512": 9.2,         # Thiết kế tiến trình theo CV 5512
        "bloom_assessment": 9.4,      # Phân hóa Bloom & Tiêu chí Rubrics
        "digital_integration": 8.8,   # Khai thác CNTT, Video, Học liệu số
        "format_nd30": 7.5,           # Kỷ luật thể thức văn thư NĐ 30
        "stem_real_world": 8.0        # Mức độ gắn kết thực tiễn & STEM
    },
    "kpi_indices": {
        "teaching_innovation_index": "88/100 (Mức Xuất sắc)",  # TII
        "curriculum_adaptation_rate": "92% (Thích ứng cao)",   # CAR
        "interaction_density": "68% (Đạt mức tương tác tích cực)",
        "plagiarism_rate": "< 2.5% (Tự biên soạn độc lập, tính nguyên bản cao)"
    },
    "recommendations_taphuan": [
        {
            "module_code": "Mô-đun 08_KHTN_THCS",
            "title": "Ứng dụng công nghệ thông tin và chuyển đổi số trong dạy học môn KHTN",
            "url": "https://taphuan.nxbgd.vn/tap-huan?grade=-1",
            "rationale": "Khắc phục triệt để lỗi thể thức văn bản hành chính theo NĐ 30/2020/NĐ-CP, đồng bộ hóa định dạng đám mây .docx trên Google Drive."
        },
        {
            "module_code": "Mô-đun 02_KHTN_THCS",
            "title": "Kỹ thuật dạy học tích cực và giáo dục STEM trong môn Khoa học tự nhiên",
            "url": "https://taphuan.nxbgd.vn/tap-huan?grade=-1",
            "rationale": "Tăng cường thiết kế các nhiệm vụ thực hành/trải nghiệm thực tế ở hoạt động Vận dụng thay vì giao bài tập lý thuyết về nhà (áp dụng cho KHTN 7 Bài 1, 2, 3)."
        },
        {
            "module_code": "Mô-đun 03_KHTN_THCS",
            "title": "Kiểm tra, đánh giá học sinh theo hướng phát triển phẩm chất, năng lực môn KHTN",
            "url": "https://taphuan.nxbgd.vn/tap-huan?grade=-1",
            "rationale": "Tiếp tục hoàn thiện và chuyển giao mô hình Rubrics 3 mức độ (đã làm rất tốt ở Bài 36 KHTN 9) cho toàn bộ các bài học thuộc phân môn Hóa học và Vật lý."
        }
    ]
}

def generate_competency_markdown(data):
    md = f"""# 🧭 BẢN ĐỒ NĂNG LỰC GIÁO VIÊN (TEACHER COMPETENCY MAP)
**Họ và tên:** {data['teacher_name']} | **Vị trí:** {data['role']}  
**Đơn vị:** {data['school']} | **Chu kỳ:** {data['cycle']}  

---

## 1. 📊 CHỈ SỐ QUẢN TRỊ HIỆU SUẤT & SÁNG TẠO (BI EXECUTIVE KPI)
*   🚀 **Chỉ số Sáng tạo Giảng dạy (Teaching Innovation Index - TII):** `{data['kpi_indices']['teaching_innovation_index']}`
*   🎯 **Tỉ lệ Thích ứng Chương trình mới (Curriculum Adaptation Rate - CAR):** `{data['kpi_indices']['curriculum_adaptation_rate']}`
*   🤝 **Mật độ Tương tác Bình quân (Average Interaction Density):** `{data['kpi_indices']['interaction_density']}`
*   🛡️ **Kiểm soát Tính nguyên bản & Đạo văn (Originality / Plagiarism):** `{data['kpi_indices']['plagiarism_rate']}`

---

## 2. 🕸️ MA TRẬN 5 CHIỀU KÍCH NĂNG LỰC SƯ PHẠM (RADAR METRICS)
| Chiều kích Năng lực | Điểm đánh giá (1-10) | Nhận định chuyên sâu của Chuyên gia Kiểm định |
|:---|:---:|:---|
| **1. Cấu trúc 5512 & Tiến trình dạy học** | **{data['metrics']['pedagogy_5512']}/10** | Thiết kế chuỗi 4 hoạt động bài bản. Phân chia rõ ràng 4 bước chuyển giao - thực hiện - báo cáo - nhận định. Cần khắc sâu hơn phần nhận xét chốt kiến thức chuẩn hóa ở bước 4. |
| **2. Phân hóa Bloom & Đánh giá năng lực** | **{data['metrics']['bloom_assessment']}/10** | Xuất sắc. Tiên phong áp dụng bảng tiêu chí Rubrics phân cấp 3 mức độ ở môn KHTN 9. Ma trận câu hỏi cân đối giữa Nhận biết, Thông hiểu và Vận dụng cao. |
| **3. Ứng dụng CNTT & Học liệu số** | **{data['metrics']['digital_integration']}/10** | Tích hợp video tư liệu trực quan, hình ảnh mô phỏng mô hình nguyên tử tốt. Khuyến khích tích hợp thêm các công cụ tương tác số (Quizizz, Mentimeter, Padlet). |
| **4. Kỷ luật Thể thức Văn thư (NĐ 30)** | **{data['metrics']['format_nd30']}/10** | Cần chấn chỉnh: Bài 3 và Bài 4 KHTN 7 căn lề trái hẹp (<30mm), một số tiêu đề còn lỗi chính tả ('BẢNG TUẦN TOÀN', 'NGUYỀN TỐ') và lưu tệp đuôi .doc cũ. |
| **5. Đổi mới phương pháp & STEM/Thực tiễn** | **{data['metrics']['stem_real_world']}/10** | Khá tốt. Đã có ý thức gắn kết di truyền học với thực tiễn đời sống. Cần nâng cấp các hoạt động vận dụng của KHTN 7 thành các dự án mini STEM trải nghiệm tại nhà. |

---

## 3. 🎯 LỘ TRÌNH PHÁT TRIỂN CHUYÊN MÔN CÁ NHÂN HÓA (PERSONALIZED CPD)
Dựa trên phân tích tự động từ hồ sơ giáo án, hệ thống khuyến nghị giáo viên hoàn thành các khóa bồi dưỡng trọng tâm trên cổng thông tin Tập huấn Bộ Giáo dục & Đào tạo:
"""
    for rec in data['recommendations_taphuan']:
        md += f"""
### 📌 {rec['module_code']}: {rec['title']}
*   **Địa chỉ truy cập:** [{rec['url']}]({rec['url']})
*   **Mục tiêu bồi dưỡng:** {rec['rationale']}
"""
    return md

if __name__ == "__main__":
    content = generate_competency_markdown(TEACHER_COMPETENCY_DATA)
    out_file = r"D:\APP-KIEM-TRA-GIAO-AN\BAN_DO_NANG_LUC_GIAO_VIEN_CO_CUC.md"
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Đã tạo lập Bản đồ Năng lực Giáo viên tại: {out_file}")
