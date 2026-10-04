import os
import re
from datetime import datetime, timedelta
import docx
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

try:
    import pypdf
except ImportError:
    pypdf = None

try:
    import openpyxl
except ImportError:
    openpyxl = None


def set_tnr(run, size_pt=13, bold=False, italic=False, color_rgb=None):
    """
    Đặt font chuẩn Times New Roman trực tiếp vào thuộc tính XML run properties (w:rFonts),
    ngăn ngừa triệt để lỗi nhảy font sang Calibri hay Arial khi mở trên Word máy tính.
    """
    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(
        f'<w:rFonts {nsdecls("w")} '
        f'w:ascii="Times New Roman" '
        f'w:hAnsi="Times New Roman" '
        f'w:cs="Times New Roman" '
        f'w:eastAsia="Times New Roman"/>'
    )
    rPr.append(rFonts)


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_table_borders(table, color="A0AEC0", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


# =========================================================================
# DỮ LIỆU ĐỐI SOÁT CHUẨN 35 TUẦN TỪ KẾ HOẠCH TỔ TOÁN - KHTN - CÔNG NGHỆ
# Trích xuất từ Kế hoạch GD tổ: Mục 4 (SHCM), Mục B (STEM), Mục 1 (Bồi dưỡng HSG),
# Mục 3 (Kiểm tra định kỳ & Kế hoạch dạy học theo tuần)
# =========================================================================
SECTION_4_ACTIVITIES = {
    1: {
        "month": 9,
        "shcm": "Họp tổ chuyên môn: Thống nhất chỉ tiêu thi đua năm học 2026 - 2027, triển khai các văn bản hướng dẫn thực hiện nhiệm vụ năm học đầu năm.",
        "shcm_performer": "Cả tổ",
        "observations": [],
        "stem": None,
        "notes": "Họp đầu năm thống nhất quy chế chuyên môn và chỉ tiêu năm học"
    },
    2: {
        "month": 9,
        "shcm": "Sinh hoạt Tổ chuyên môn theo hướng NCBH (Bước 1: Toán 9, KHTN 8, Công Nghệ 8). Thảo luận phân công GV thiết kế bài dạy minh họa.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "KHTN 9 (Sinh)", "lesson": "Khái quát di truyền", "teacher": "Cô Cúc"}
        ],
        "stem": None,
        "notes": "Thực hiện SHCM nghiên cứu bài học bước 1 và dự giờ Cô Cúc"
    },
    3: {
        "month": 9,
        "shcm": "Họp sơ kết tổ chuyên môn tháng 09/2026. Đánh giá việc soạn giảng và thực hiện chương trình tháng 9.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "KHTN 9 (Sinh)", "lesson": "Các qui luật di truyền của Mendel", "teacher": "Cô Cúc"}
        ],
        "stem": {"subject": "KHTN 7", "topic": "Mô hình cấu tạo nguyên tử", "teacher": "Thầy Thắng"},
        "notes": "Sơ kết tổ tháng 9, dự giờ Cô Cúc và thực hiện chuyên đề STEM của Thầy Thắng"
    },
    4: {
        "month": 10,
        "shcm": "Sinh hoạt Tổ chuyên môn theo hướng NCBH (Bước 2: Dạy minh họa và dự giờ phân tích bài học).",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "Toán 9", "lesson": "Phương trình quy về phương trình bậc nhất một ẩn", "teacher": "Cô Linh"}
        ],
        "stem": None,
        "notes": "SHCM theo hướng NCBH bước 2 kết hợp dự giờ Cô Linh"
    },
    5: {
        "month": 10,
        "shcm": "Sinh hoạt chuyên môn tổ: Rà soát tiến độ chương trình tuần 5; đối chiếu việc soạn bài theo CV 5512 các môn Toán, KHTN, Công nghệ; chuẩn bị kiến thức trọng tâm tháng 10.",
        "shcm_performer": "Toàn thể GV tổ",
        "observations": [],
        "stem": None,
        "notes": "Tăng cường bồi dưỡng học sinh giỏi và rà soát kiểm tra hồ sơ bài dạy trên Drive"
    },
    6: {
        "month": 10,
        "shcm": "Sinh hoạt chuyên môn định kỳ: Thống nhất nội dung ôn tập và chuẩn bị ma trận, bản đặc tả đề kiểm tra giữa học kỳ 1.",
        "shcm_performer": "GV Toán, KHTN, CN",
        "observations": [],
        "stem": None,
        "notes": "Chuẩn bị đề kiểm tra giữa kỳ 1 theo quy định ma trận - đặc tả"
    },
    7: {
        "month": 10,
        "shcm": "Sinh hoạt chuyên môn định kỳ: Trao đổi đổi mới phương pháp giảng dạy phân hóa và ứng dụng CNTT trong kiểm tra đánh giá.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "Toán 7", "lesson": "Làm quen với số thập phân vô hạn tuần hoàn", "teacher": "Cô Thủy"},
            {"subject": "Toán 9", "lesson": "Tỉ số lượng giác của góc nhọn", "teacher": "Cô Linh"},
            {"subject": "Công nghệ 7", "lesson": "Nhân giống vô tính cây trồng", "teacher": "Cô Hằng"}
        ],
        "stem": None,
        "notes": "Tổ chức dự giờ 03 tiết chuyên môn: Cô Thủy, Cô Linh, Cô Hằng"
    },
    8: {
        "month": 10,
        "shcm": "Họp sơ kết tổ chuyên môn tháng 10/2026. Đánh giá công tác kiểm tra giữa kỳ và nền nếp chuyên môn.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "Công nghệ 8", "lesson": "Bản vẽ nhà", "teacher": "Cô Trang"}
        ],
        "stem": None,
        "notes": "Dự giờ Cô Trang và sơ kết công tác chuyên môn tổ tháng 10"
    },
    9: {
        "month": 11,
        "shcm": "Sinh hoạt chuyên môn theo hướng NCBH (Bước 3: Toán 9, KHTN 8, Công Nghệ 8). Thảo luận và rút kinh nghiệm bài dạy minh họa.",
        "shcm_performer": "Cả tổ",
        "observations": [],
        "stem": None,
        "notes": "Thực hiện SHCM nghiên cứu bài học bước 3"
    },
    10: {
        "month": 11,
        "shcm": "Sinh hoạt chuyên môn: Đẩy mạnh phong trào thi đua dạy tốt - học tốt chào mừng ngày Nhà giáo Việt Nam 20/11.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "KHTN 9 (Lí)", "lesson": "Công và công suất", "teacher": "Thầy Thành"}
        ],
        "stem": None,
        "notes": "Dự giờ Thầy Thành môn KHTN 9 thi đua chào mừng 20/11"
    },
    11: {
        "month": 11,
        "shcm": "Sinh hoạt chuyên môn tổ: Đánh giá nền nếp dạy học và trao đổi việc sử dụng đồ dùng, phòng thực hành bộ môn.",
        "shcm_performer": "Toàn thể GV tổ",
        "observations": [
            {"subject": "KHTN 8 (Sinh)", "lesson": "Hô hấp ở người", "teacher": "Cô Kế"}
        ],
        "stem": None,
        "notes": "Dự giờ Cô Kế môn KHTN 8 bài Hô hấp ở người"
    },
    12: {
        "month": 12,
        "shcm": "Họp sơ kết tổ chuyên môn tháng 11/2026. Đánh giá phong trào thi đua 20/11 và chuẩn bị kế hoạch tháng 12.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "KHTN 6", "lesson": "Sự lớn lên và sinh sản của tế bào", "teacher": "Thầy Thành"},
            {"subject": "Công nghệ 7", "lesson": "Giới thiệu về rừng", "teacher": "Cô Hằng"}
        ],
        "stem": None,
        "notes": "Sơ kết tổ tháng 11 và dự giờ 02 tiết: Thầy Thành, Cô Hằng"
    },
    13: {
        "month": 12,
        "shcm": "Sinh hoạt chuyên môn theo hướng NCBH (Bước 4: Áp dụng bài học vào thực tiễn giảng dạy).",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "Toán 7", "lesson": "Tam giác cân. Đường trung trực của đoạn thẳng", "teacher": "Cô Thủy"},
            {"subject": "KHTN 6", "lesson": "Cơ thể sinh vật", "teacher": "Cô Kế"}
        ],
        "stem": {"subject": "Công nghệ 6", "topic": "Dự án bữa ăn yêu thương", "teacher": "Cô Hằng"},
        "notes": "SHCM NCBH bước 4, dự giờ Cô Thủy, Cô Kế và dự án STEM của Cô Hằng"
    },
    14: {
        "month": 12,
        "shcm": "Sinh hoạt chuyên môn: Chuẩn bị nội dung ôn tập kiểm tra cuối học kỳ 1 và rà soát tiến độ chương trình.",
        "shcm_performer": "Toàn thể GV",
        "observations": [],
        "stem": {"subject": "Công nghệ 9", "topic": "TH: Lắp đặt mạng điện trong nhà", "teacher": "Cô Trang"},
        "notes": "Thực hiện chuyên đề thực hành STEM Công nghệ 9 do Cô Trang phụ trách"
    },
    15: {
        "month": 12,
        "shcm": "Sinh hoạt chuyên môn: Thống nhất lịch kiểm tra và phân công coi, chấm kiểm tra cuối học kỳ 1.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "KHTN 9", "lesson": "Đột biến gene", "teacher": "Cô Cúc"},
            {"subject": "Công nghệ 9", "lesson": "TH: Lắp đặt mạng điện trong nhà", "teacher": "Cô Trang"},
            {"subject": "Công nghệ 6", "lesson": "Trang phục trong đời sống", "teacher": "Cô Hằng"}
        ],
        "stem": None,
        "notes": "Dự giờ 03 tiết chuyên môn: Cô Cúc, Cô Trang, Cô Hằng"
    },
    16: {
        "month": 12,
        "shcm": "Họp sơ kết tổ chuyên môn tháng 12/2026. Kiểm tra việc hoàn thành chương trình HK1 và hồ sơ sổ sách.",
        "shcm_performer": "Cả tổ",
        "observations": [],
        "stem": None,
        "notes": "Sơ kết tổ tháng 12, rà soát tiến độ hoàn thành chương trình HK1"
    },
    17: {
        "month": 12,
        "shcm": "Tổ chức kiểm tra cuối học kỳ 1 theo kế hoạch của trường. Tiến hành chấm bài, vào điểm nghiêm túc.",
        "shcm_performer": "Toàn thể GV",
        "observations": [],
        "stem": None,
        "notes": "Kiểm tra đánh giá cuối học kỳ 1"
    },
    18: {
        "month": 1,
        "shcm": "Họp sơ kết học kỳ 1. Đánh giá chất lượng bộ môn và triển khai kế hoạch chuyên môn học kỳ 2.",
        "shcm_performer": "Toàn thể GV tổ",
        "observations": [],
        "stem": None,
        "notes": "Sơ kết học kỳ 1, hoàn thành điểm số và học bạ điện tử"
    },
    19: {
        "month": 1,
        "shcm": "Sinh hoạt Tổ chuyên môn theo hướng NCBH HK2 (Bước 1: Công Nghệ 6, KHTN 9, Toán 7).",
        "shcm_performer": "Cả tổ",
        "observations": [],
        "stem": None,
        "notes": "Khởi động SHCM nghiên cứu bài học học kỳ 2 bước 1"
    },
    20: {
        "month": 1,
        "shcm": "Sinh hoạt chuyên môn định kỳ: Rà soát bài giảng số và thiết bị dạy học các khối lớp.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "KHTN 6", "lesson": "Nấm", "teacher": "Cô Kế"},
            {"subject": "Công nghệ 6", "lesson": "Sử dụng và bảo quản trang phục", "teacher": "Cô Hằng"},
            {"subject": "Công nghệ 9", "lesson": "Nghề nghiệp trong lĩnh vực kĩ thuật và công nghệ", "teacher": "Cô Trang"}
        ],
        "stem": None,
        "notes": "Tổ chức dự giờ 03 tiết: Cô Kế, Cô Hằng, Cô Trang"
    },
    21: {
        "month": 1,
        "shcm": "Họp sơ kết tổ chuyên môn tháng 01/2027. Đánh giá nền nếp dạy học trước kỳ nghỉ Tết Nguyên đán.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "KHTN 9 (Sinh)", "lesson": "NST và bộ NST", "teacher": "Cô Cúc"},
            {"subject": "Toán 9", "lesson": "Định lí Viete và ứng dụng", "teacher": "Cô Linh"}
        ],
        "stem": None,
        "notes": "Sơ kết tổ tháng 01 và dự giờ Cô Cúc, Cô Linh"
    },
    22: {
        "month": 2,
        "shcm": "Sinh hoạt chuyên môn theo hướng NCBH (Bước 2: Thảo luận xây dựng tiến trình bài dạy minh họa).",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "Toán 6", "lesson": "Phép cộng và phép trừ phân thức đại số", "teacher": "Cô Thủy"}
        ],
        "stem": None,
        "notes": "SHCM NCBH bước 2 và dự giờ Cô Thủy môn Toán 6"
    },
    23: {
        "month": 2,
        "shcm": "Họp sơ kết tổ chuyên môn tháng 02/2027. Đánh giá tình hình dạy học sau kỳ nghỉ Tết Nguyên đán.",
        "shcm_performer": "Cả tổ",
        "observations": [],
        "stem": {"subject": "Toán 8", "topic": "Mô hình ứng dụng thực tế hai tam giác đồng dạng", "teacher": "Cô Thủy"},
        "notes": "Sơ kết tổ tháng 2 và thực hiện chuyên đề STEM Toán 8 của Cô Thủy"
    },
    24: {
        "month": 3,
        "shcm": "Sinh hoạt chuyên môn theo hướng NCBH (Bước 3: Dạy thực nghiệm và trao đổi rút kinh nghiệm).",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "KHTN 6", "lesson": "Đa dạng sinh học", "teacher": "Thầy Thành"}
        ],
        "stem": None,
        "notes": "SHCM NCBH bước 3 và dự giờ Thầy Thành môn KHTN 6"
    },
    25: {
        "month": 3,
        "shcm": "Sinh hoạt chuyên môn: Đổi mới hoạt động trải nghiệm, hướng nghiệp và giáo dục STEM.",
        "shcm_performer": "Toàn thể GV tổ",
        "observations": [
            {"subject": "HĐTN 8", "lesson": "Thiết kế Infographic: 'Ứng phó với thiên tai'", "teacher": "Cô Kế"}
        ],
        "stem": None,
        "notes": "Dự giờ Cô Kế môn HĐTN 8 chuyên đề thiết kế Infographic"
    },
    26: {
        "month": 3,
        "shcm": "Sinh hoạt chuyên môn: Rà soát chương trình và chuẩn bị ma trận đề kiểm tra giữa học kỳ 2.",
        "shcm_performer": "GV Toán, KHTN, CN",
        "observations": [
            {"subject": "KHTN 7", "lesson": "Cảm ứng ở sinh vật và tập tính ở động vật", "teacher": "Cô Cúc"},
            {"subject": "Toán 7", "lesson": "Phép cộng và phép trừ đa thức một biến", "teacher": "Cô Thủy"}
        ],
        "stem": None,
        "notes": "Dự giờ 02 tiết: Cô Cúc (KHTN 7) và Cô Thủy (Toán 7)"
    },
    27: {
        "month": 3,
        "shcm": "Họp sơ kết tổ chuyên môn tháng 03/2027. Đánh giá kết quả kiểm tra giữa kỳ 2 và công tác bồi dưỡng HSG.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "Toán 9", "lesson": "Bản tần số và biểu đồ tần số", "teacher": "Cô Linh"}
        ],
        "stem": {"subject": "Công nghệ 8", "topic": "Dự án hệ thống nhà thông minh sử dụng cảm biến ánh sáng", "teacher": "Cô Trang"},
        "notes": "Sơ kết tổ tháng 3, dự giờ Cô Linh và chuyên đề STEM của Cô Trang"
    },
    28: {
        "month": 4,
        "shcm": "Sinh hoạt chuyên môn theo hướng NCBH (Bước 4: Tổng kết và nhân rộng mô hình bài dạy).",
        "shcm_performer": "Cả tổ",
        "observations": [],
        "stem": None,
        "notes": "Hoàn thành chu trình SHCM nghiên cứu bài học bước 4"
    },
    29: {
        "month": 4,
        "shcm": "Sinh hoạt chuyên môn định kỳ: Rà soát việc sử dụng đồ dùng, phòng thực hành bộ môn.",
        "shcm_performer": "Toàn thể GV tổ",
        "observations": [
            {"subject": "Công nghệ 8", "lesson": "Mạch điện điều khiển sử dụng mô đun cảm biến", "teacher": "Cô Trang"}
        ],
        "stem": None,
        "notes": "Dự giờ Cô Trang môn Công nghệ 8"
    },
    30: {
        "month": 4,
        "shcm": "Sinh hoạt chuyên môn: Đẩy mạnh các hoạt động trải nghiệm sáng tạo STEM cuối năm học.",
        "shcm_performer": "Cả tổ",
        "observations": [],
        "stem": {"subject": "KHTN 8", "topic": "Dự án các kiểu hệ sinh thái nhân tạo", "teacher": "Cô Kế"},
        "notes": "Thực hiện chuyên đề STEM KHTN 8 do Cô Kế chủ trì"
    },
    31: {
        "month": 4,
        "shcm": "Sinh hoạt chuyên môn: Đánh giá hiệu quả các chuyên đề STEM và phương pháp dạy học dự án.",
        "shcm_performer": "Cả tổ",
        "observations": [],
        "stem": {"subject": "KHTN 7 & 8", "topic": "STEM Mầm Xanh Sinh Trưởng (Cô Cúc) & Đèn kéo quân (Thầy Thành)", "teacher": "Cô Cúc, Thầy Thành"},
        "notes": "Tổ chức 02 chuyên đề STEM: Cô Cúc (KHTN 7) và Thầy Thành (KHTN 8)"
    },
    32: {
        "month": 4,
        "shcm": "Họp sơ kết tổ chuyên môn tháng 04/2027. Thống nhất kế hoạch ôn tập và kiểm tra cuối năm học.",
        "shcm_performer": "Cả tổ",
        "observations": [
            {"subject": "KHTN 8", "lesson": "Sự nở vì nhiệt", "teacher": "Thầy Thành"}
        ],
        "stem": {"subject": "Toán 9", "topic": "Chuyên đề STEM: Mũ sinh nhật", "teacher": "Cô Linh"},
        "notes": "Sơ kết tổ tháng 4, dự giờ Thầy Thành và chuyên đề STEM Toán 9 của Cô Linh"
    },
    33: {
        "month": 5,
        "shcm": "Sinh hoạt chuyên môn: Hoàn thành ôn tập và chuẩn bị cơ sở vật chất cho kỳ kiểm tra cuối năm.",
        "shcm_performer": "Toàn thể GV tổ",
        "observations": [],
        "stem": None,
        "notes": "Rà soát hoàn thành chương trình môn học các khối lớp"
    },
    34: {
        "month": 5,
        "shcm": "Tổ chức kiểm tra cuối năm học. Hoàn thiện chấm bài, vào điểm hệ thống và kiểm kê thiết bị dạy học.",
        "shcm_performer": "Toàn thể GV",
        "observations": [],
        "stem": None,
        "notes": "Kiểm tra đánh giá cuối năm, nhập điểm và kiểm kê thiết bị bộ môn"
    },
    35: {
        "month": 5,
        "shcm": "Họp tổng kết hoạt động tổ chuyên môn cả năm học 2026 - 2027. Đánh giá xếp loại thi đua viên chức.",
        "shcm_performer": "Cả tổ",
        "observations": [],
        "stem": None,
        "notes": "Tổng kết tổ chuyên môn cả năm, bình xét thi đua năm học 2026 - 2027"
    }
}

# =========================================================================
# BÀI DẠY TRỌNG TÂM CÁC KHỐI LỚP THEO TUẦN (MỤC 3 TRONG KẾ HOẠCH GD TỔ)
# =========================================================================
WEEKLY_LESSONS = {
    1: "Toán 6 (Tập hợp các số tự nhiên); KHTN 6 (Mở đầu KHTN); KHTN 7 (Bài 1. Mở đầu KHTN); KHTN 9 (Khái quát di truyền).",
    2: "Toán 6 (Bài 6. Lũy thừa với số mũ tự nhiên); KHTN 7 (Bài 2. Nguyên tử); KHTN 8 (Bài 13. Khối lượng riêng); KHTN 9 (Di truyền Mendel).",
    3: "Toán 6 (Lũy thừa với số mũ tự nhiên); KHTN 7 (Bài 3. Nguyên tố hóa học); KHTN 8 (Áp suất chất lỏng); KHTN 9 (Lai một cặp tính trạng).",
    4: "Toán 9 (Bài phương trình quy về PT bậc nhất một ẩn); KHTN 7 (Nguyên tố hóa học); KHTN 8 (Khối lượng riêng); Công nghệ 7 (Đất trồng).",
    5: "Toán 8 (Tiết 17-18: Bài 5. Phép chia đa thức cho đơn thức; Tiết 19: Bài 13. Hình chữ nhật; Tiết 20: Bài 14. Hình thoi và hình vuông); Toán 7 (Tiết 19: Bài 11. Định lí và chứng minh định lí); KHTN 7 (Trao đổi nước và dinh dưỡng); KHTN 8 (Áp suất chất khí).",
    6: "Toán 9 (Tiết 21-22: Bài 6. Bất phương trình bậc nhất một ẩn); Toán 7 (Tiết 22: Bài 5. Làm quen với số thập phân vô hạn tuần hoàn); KHTN 8 (Tiết 4: Bài 31. Hệ vận động ở người).",
    7: "Toán 7 (Tiết 25: Số thập phân vô hạn; Tiết 26: Số vô tỉ - Căn bậc hai; Tiết 27: Tổng các góc tam giác; Tiết 28: Hai tam giác bằng nhau); KHTN 9 (Tiết 13: Phi kim; Tiết 7: Cơ năng); Công nghệ 7 (Nhân giống vô tính).",
    8: "Công nghệ 8 (Bài 4. Bản vẽ nhà); Công nghệ 7 (Tiết 8: Bài 5. Bảo quản chế biến thực phẩm & Bài 6: Dự án trồng rau an toàn); KHTN 6 (Tế bào).",
    9: "Toán 9 (Hệ hai phương trình bậc nhất hai ẩn); KHTN 8 (Chất - Biến đổi hóa học); KHTN 9 (Định luật bảo toàn khối lượng); Công nghệ 8 (Bản vẽ kỹ thuật).",
    10: "Toán 9 (Giải toán bằng cách lập hệ phương trình); KHTN 9 (Tiết 10: Bài 39. Tái bản DNA & phiên mã RNA; Tiết Lý: Công và công suất); Công nghệ 9 (Mạng điện trong nhà).",
    11: "Toán 8 (Tiết 41: Bài 8. Tổng và hiệu hai lập phương; Tiết 42: Luyện tập); KHTN 8 (Tiết 13: Bài 34. Hệ hô hấp; Tiết 14: Bài 18. Moment lực; Tiết 17: Bài 6. Tính theo PTHH); Công nghệ 7 (Tiết 11: Bảo quản thực phẩm).",
    12: "KHTN 6 (Sự lớn lên và sinh sản của tế bào); KHTN 8 (Tiết 15: Bài 35. Hệ bài tiết ở người); Công nghệ 7 (Giới thiệu về rừng).",
    13: "Toán 7 (Tam giác cân - Đường trung trực); KHTN 6 (Cơ thể sinh vật); Công nghệ 6 (Dự án Bữa ăn yêu thương - STEM); KHTN 9 (Nhiễm sắc thể).",
    14: "Toán 9 (Tiết 55: Bài 14. Cung và dây của một đường tròn); KHTN 7 (Tiết 58-60: Bài 18. Nam châm); Công nghệ 9 (Bài 6. Thực hành: Lắp đặt mạng điện trong nhà).",
    15: "KHTN 6 (Tiết 62-63: Bài 27. Vi khuẩn); KHTN 9 (Đột biến gene); Công nghệ 8 (Tiết 15: Bài 8. Gia công cơ khí bằng tay); Công nghệ 6 (Trang phục trong đời sống).",
    16: "Toán 8 (Ôn tập học kỳ 1); Công nghệ 9 (Tiết 32: Bài 7. Ngành nghề lắp đặt mạng điện); KHTN 7 (Ôn tập chủ đề Ánh sáng và Từ trường).",
    17: "Ôn tập tổng hợp và thực hiện Kiểm tra đánh giá cuối học kỳ 1 các môn Toán, KHTN, Công nghệ, HĐTN theo lịch trường.",
    18: "Hoàn thành kiểm tra bù, chấm bài, nhập điểm CSDL ngành; sơ kết học kỳ 1 và triển khai bài dạy đầu học kỳ 2.",
    19: "KHTN 8 (Tiết 22: Bài 38. Hệ nội tiết; Tiết 29: Bài 9. Base - Thang pH); Công nghệ 8 (Tiết 19: Gia công cơ khí bằng tay); Toán 7 (Biểu thức đại số).",
    20: "Toán 7 (Tiết 79: Bài 32. Đường vuông góc - đường xiên; Tiết 80: Bài 33. Quan hệ ba cạnh tam giác); KHTN 7 (Tiết 79-80: Bài 23. Quang hợp); KHTN 9 (Tiết 20: NST; Tiết 39-40: Tinh bột và cellulose); KHTN 6 (Bài 30. Nấm).",
    21: "Toán 7 (Tiết 83-84: Luyện tập chung tam giác); Toán 9 (Định lí Viete và ứng dụng); KHTN 9 (Tiết 21-22: Bài 8. Thấu kính; NST và bộ NST).",
    22: "Toán 6 (Tiết 88: Bài 32. Điểm và đường thẳng; Phép cộng trừ phân thức); KHTN 8 (Tiết 33: Bài 10. Oxide); KHTN 7 (Tập tính động vật).",
    23: "Toán 7 (Tiết 91-92: Bài 35. Sự đồng quy trong tam giác); Toán 9 (Tiết 91: Bài 27. Góc nội tiếp); KHTN 8 (Tiết 26-27: Bài 40. Sinh sản ở người); STEM Toán 8 (Tam giác đồng dạng).",
    24: "Toán 9 (Tiết 93: Bài 28. Đường tròn ngoại tiếp - nội tiếp); KHTN 6 (Đa dạng sinh học); KHTN 8 (Tiết 32: Bài 26. Năng lượng nhiệt và nội năng).",
    25: "Toán 8 (Tiết 98: Bài 25. Phương trình bậc nhất một ẩn); Công nghệ 7 (Tiết 25-26: Bài 13. Thực hành lập kế hoạch nuôi vật nuôi); HĐTN 8 (Infographic Thiên tai).",
    26: "KHTN 7 (Cảm ứng ở sinh vật); Toán 7 (Phép cộng trừ đa thức một biến); Chuẩn bị kiểm tra giữa kỳ 2 môn Toán, KHTN.",
    27: "Toán 9 (Bảng tần số và biểu đồ tần số); Công nghệ 8 (Dự án Nhà thông minh cảm biến ánh sáng - STEM); KHTN 8 (Nhiệt năng).",
    28: "Toán 8 (Định lí Thalès trong tam giác); KHTN 9 (Hợp chất hữu cơ); Công nghệ 9 (Thực hành mạch điện); Ôn tập giữa kỳ 2.",
    29: "Công nghệ 8 (Mạch điện điều khiển dùng mô đun cảm biến); KHTN 7 (Sinh trưởng và phát triển ở sinh vật); Toán 9 (Độ dài cung tròn, diện tích hình quạt).",
    30: "Toán 8 (Tiết 117: Hàm số - đồ thị; Tiết 118, 121: Hàm số bậc nhất; Tiết 119-120: Xác suất biến cố; Tiết 122: Hệ số góc); KHTN 6 (Tiết 122-123: Biểu đồ cột); STEM KHTN 8 (Hệ sinh thái nhân tạo).",
    31: "KHTN 7 (Tiết 125-127: Bài 40. Sinh sản hữu tính ở sinh vật); STEM KHTN 7 (Mầm Xanh Sinh Trưởng); STEM KHTN 8 (Đèn kéo quân).",
    32: "Toán 9 (Tiết 128-129: Bài 32. Hình cầu); KHTN 8 (Sự nở vì nhiệt); Công nghệ 7 (Tiết 32: Bài 14. An toàn tiết kiệm điện năng); STEM Toán 9 (Mũ sinh nhật).",
    33: "Toán, KHTN, Công nghệ các khối 6, 7, 8, 9: Hoàn thành nội dung bài dạy theo phân phối chương trình; ôn tập tổng hợp cuối năm học.",
    34: "Tổ chức Kiểm tra đánh giá cuối năm học; chấm kiểm tra, vào điểm học bạ điện tử và kiểm kê thiết bị dạy học.",
    35: "Tổng kết chuyên môn năm học 2026 - 2027; hoàn thiện hồ sơ thi đua tổ và nộp báo cáo tổng kết về Ban Giám Hiệu."
}

# =========================================================================
# BỒI DƯỠNG HỌC SINH GIỎI THEO TUẦN (MỤC 1 TRANG 12-14)
# =========================================================================
WEEKLY_HSG = {
    1: "KHTN 9 (Hóa: Nguyên tử - Nguyên tố hóa học, tiết 1-2); KHTN 9 (Sinh: Di truyền và biến dị); Toán 8, 9 (Chuyên đề số học).",
    2: "KHTN 9 (Hóa: Nguyên tử - Nguyên tố hóa học, tiết 3-4); KHTN 9 (Sinh: Các qui luật của Mendel); Toán 8, 9 (Đại số nâng cao).",
    3: "KHTN 9 (Hóa: Tách chất, tiết 5-6); KHTN 9 (Sinh: Nhiễm sắc thể và nguyên phân); Toán 8, 9 (Hình học nâng cao).",
    4: "KHTN 9 (Hóa: Tách chất, tiết 7-8); KHTN 9 (Sinh: Giảm phân và thụ tinh); Toán 8, 9 (Phương trình nâng cao).",
    5: "KHTN 9 (Hóa: Cấu tạo bảng tuần hoàn, tiết 9-10); KHTN 9 (Sinh: Di truyền liên kết); Toán 8, 9 (Bất đẳng thức và cực trị đại số).",
    6: "KHTN 9 (Hóa: Liên kết hóa học, tiết 11-12); KHTN 9 (Sinh: Môi trường và các nhân tố sinh thái); Toán 8, 9 (Phương trình vô tỉ).",
    7: "KHTN 9 (Hóa: Hiệu suất phản ứng, tiết 13-14); KHTN 9 (Sinh: Ảnh hưởng ánh sáng và nhiệt độ); Toán 8, 9 (Hệ thức lượng).",
    8: "KHTN 9 (Hóa: So sánh số mol, tiết 15-16); KHTN 9 (Sinh: Quần thể sinh vật); Toán 8, 9 (Tứ giác nội tiếp).",
    9: "KHTN 9 (Hóa: Định luật bảo toàn khối lượng, tiết 17-18); KHTN 9 (Sinh: Quần xã sinh vật); Toán 8, 9 (Số chính phương).",
    10: "KHTN 9 (Hóa: Mol và tỉ khối chất khí, tiết 19-20); KHTN 9 (Sinh: Hệ sinh thái); Toán 8, 9 (Phương trình nghiệm nguyên).",
    11: "KHTN 9 (Hóa: Xác định tên Kim loại và CTHH, tiết 21-22); KHTN 9 (Sinh: Hệ sinh thái); Toán 8, 9 (Đa thức và nghiệm).",
    12: "KHTN 9 (Hóa: Hỗn hợp kim loại, tiết 23-24); KHTN 9 (Sinh: Lý thuyết hệ sinh thái); Toán 8, 9 (Đường tròn nâng cao).",
    13: "KHTN 9 (Hóa: Bài tập dung dịch, tiết 25-26); KHTN 9 (Sinh: Bài tập thí nghiệm Mendel); Toán 8, 9 (Toán thực tế).",
    14: "KHTN 9 (Hóa: Kim loại mạnh và muối kim loại yếu, tiết 27-28); KHTN 9 (Sinh: Bài tập di truyền); Toán 8, 9 (Tam giác đồng dạng nâng cao).",
    15: "KHTN 9 (Hóa: Dạng toán quy về 100, tiết 29-30); KHTN 9 (Sinh: Bài tập di truyền liên kết); Toán 8, 9 (Tổ hợp và logic).",
    16: "KHTN 9 (Hóa: Nồng độ dung dịch trước phản ứng, tiết 31-32); KHTN 9 (Sinh: Bài tập về Nhiễm sắc thể); Toán 8, 9 (Giải đề HSG trường).",
    17: "KHTN 9 (Hóa: Tổng hợp bài tập kim loại - phi kim, tiết 33-34); KHTN 9 (Sinh: Hệ sinh thái & thực hành); Toán 8, 9 (Rà soát kiến thức).",
    18: "Giải các bộ đề thi học sinh giỏi cấp trường và thị xã; rèn kỹ năng trình bày bài thi (Toán, KHTN Hóa, KHTN Sinh)."
}

# =========================================================================
# LỊCH KIỂM TRA ĐỊNH KỲ (MỤC 3 TRANG 8-12)
# =========================================================================
WEEKLY_EXAMS = {
    8: "Kiểm tra Giữa kì I môn KHTN 6 (90 phút) & HĐTN 9 (45 phút)",
    9: "Kiểm tra Giữa kì I môn KHTN 9 (90 phút)",
    10: "Kiểm tra Giữa kì I môn Công nghệ 9 (45 phút) & Toán 9 (90 phút)",
    17: "Kiểm tra Cuối kì I môn Toán 9 (90 phút), KHTN 9 (90 phút), Công nghệ 9 (45 phút)",
    18: "Kiểm tra Cuối kì I môn HĐTN 9 (45 phút) & hoàn thành kiểm tra bù",
    26: "Kiểm tra Giữa kì II môn HĐTN 9 (45 phút) & Toán 8 (90 phút)",
    27: "Kiểm tra Giữa kì II môn KHTN 9 (90 phút)",
    28: "Kiểm tra Giữa kì II môn Công nghệ 9 (45 phút) & Toán 9 (90 phút)",
    34: "Kiểm tra Cuối kì II môn KHTN 9 (90 phút), Công nghệ 9 (45 phút), Toán 9 (90 phút)",
    35: "Kiểm tra Cuối kì II môn HĐTN 9 & hoàn thành tổng kết điểm số"
}


def extract_text_from_file(file_path):
    """Trích xuất toàn bộ text thô từ các định dạng file: docx, pdf, xlsx, txt"""
    ext = os.path.splitext(file_path)[1].lower()
    text_lines = []

    if ext == ".docx":
        doc = docx.Document(file_path)
        for p in doc.paragraphs:
            if p.text.strip():
                text_lines.append(p.text.strip())
        for tbl in doc.tables:
            for row in tbl.rows:
                row_txt = " | ".join([c.text.strip() for c in row.cells if c.text.strip()])
                if row_txt:
                    text_lines.append(row_txt)

    elif ext == ".pdf" and pypdf is not None:
        reader = pypdf.PdfReader(file_path)
        for page in reader.pages:
            t = page.extract_text()
            if t:
                for line in t.splitlines():
                    if line.strip():
                        text_lines.append(line.strip())

    elif ext in [".xlsx", ".xlsm"] and openpyxl is not None:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                r_txt = " | ".join([str(c).strip() for c in row if c is not None and str(c).strip()])
                if r_txt:
                    text_lines.append(r_txt)

    elif ext in [".txt", ".csv"]:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                if line.strip():
                    text_lines.append(line.strip())

    return text_lines


def parse_plan_to_weeks(text_lines):
    """
    Phân tích nội dung kế hoạch thành các khối thông tin theo tuần (Tuần 1, Tuần 2, ... Tuần 35)
    hoặc trích xuất các hoạt động trọng tâm.
    """
    weeks_data = {}
    current_week = None
    week_pattern = re.compile(r'(?:tuần|tu?n)\s*(\d{1,2})', re.IGNORECASE)

    general_tasks = []

    for line in text_lines:
        match = week_pattern.search(line)
        if match:
            w_num = int(match.group(1))
            current_week = f"Tuần {w_num}"
            if current_week not in weeks_data:
                weeks_data[current_week] = []
            weeks_data[current_week].append(line)
        elif current_week:
            weeks_data[current_week].append(line)
        else:
            if len(line) > 5:
                general_tasks.append(line)

    if not weeks_data:
        for i in range(1, 36):
            weeks_data[f"Tuần {i}"] = [f"Thực hiện giảng dạy theo tiến độ chương trình Tuần {i}"]

    return {
        "weeks": list(weeks_data.keys()),
        "weeks_details": weeks_data,
        "general_tasks": general_tasks[:15]
    }


def compute_week_dates(week_num, school_start_date=None):
    """
    Tính ngày bắt đầu (Thứ Hai) và kết thúc (Thứ Bảy) của tuần dựa trên ngày khai giảng 07/09/2026.
    """
    if not school_start_date:
        base_monday = datetime(2026, 9, 7)
    else:
        base_monday = school_start_date

    target_monday = base_monday + timedelta(weeks=(week_num - 1))
    target_saturday = target_monday + timedelta(days=5)

    return target_monday, target_saturday


def build_schedule_events(week_num, target_monday, sec4_info, lesson_focus, hsg_focus, exam_focus):
    """
    Xây dựng bảng lịch công tác tuần chi tiết từng ngày (Thứ 2 đến Thứ 7, Sáng và Chiều)
    kết hợp toàn bộ dữ liệu thực tế từ Kế hoạch giáo dục tổ:
    - Mục 4. Sinh hoạt chuyên môn & NCBH & Sơ kết tổ
    - Mục B. Chuyên đề STEM
    - Mục 1. Bồi dưỡng học sinh giỏi
    - Mục 3. Bài dạy trọng tâm & Kiểm tra đánh giá định kỳ
    """
    def format_date_str(dt):
        return dt.strftime("%d/%m/%Y")

    d2 = target_monday
    d3 = target_monday + timedelta(days=1)
    d4 = target_monday + timedelta(days=2)
    d5 = target_monday + timedelta(days=3)
    d6 = target_monday + timedelta(days=4)
    d7 = target_monday + timedelta(days=5)

    shcm_text = sec4_info.get("shcm", f"Sinh hoạt Tổ chuyên môn định kỳ tuần {week_num}: Rà soát tiến độ chương trình và nền nếp chuyên môn.")
    shcm_performer = sec4_info.get("shcm_performer", "Toàn thể giáo viên tổ")
    observations = sec4_info.get("observations", [])
    stem = sec4_info.get("stem", None)

    # 1. THỨ BA SÁNG: GIẢNG DẠY CHÍNH KHÓA & TRỌNG TÂM BÀI DẠY
    tue_morning_content = (
        f"- Giảng dạy chính khóa theo TKB (Toán, KHTN, Công nghệ, HĐTN).\n"
        f"- Trọng tâm bài dạy tuần {week_num}: {lesson_focus}\n"
        f"- Chuẩn bị đồ dùng dạy học, thiết bị thí nghiệm cho các tiết học tại phòng bộ môn."
    )

    # 2. THỨ BA CHIỀU: BỒI DƯỠNG HỌC SINH GIỎI
    hsg_text = hsg_focus if hsg_focus else "Bồi dưỡng đội tuyển HSG khối 8, 9 các môn Toán, KHTN theo chuyên đề nâng cao."
    tue_afternoon_content = (
        f"- Bồi dưỡng Đội tuyển học sinh giỏi theo kế hoạch tổ:\n  + {hsg_text}\n"
        f"- Kiểm tra, cập nhật điểm số và nhận xét học sinh trên Cổng thông tin Giáo dục điện tử."
    )

    # 3. THỨ TƯ SÁNG: DỰ GIỜ THỨ NHẤT (NẾU CÓ) HOẶC GIẢNG DẠY DỰ GIỜ ĐỒNG NGHIỆP
    if observations and len(observations) >= 1:
        obs1 = observations[0]
        wed_morning_content = (
            f"- Giảng dạy chính khóa theo TKB.\n"
            f"- Tổ chức dự giờ chuyên môn cấp tổ: Môn {obs1.get('subject', '')} - Bài '{obs1.get('lesson', '')}' do {obs1.get('teacher', '')} giảng dạy.\n"
            f"- Các thành viên trong tổ tham dự, quan sát và ghi chép phiếu đánh giá tiết dạy theo định hướng phát triển năng lực."
        )
        wed_morning_attendees = f"Tổ chuyên môn & GV {obs1.get('subject', '')}"
        wed_morning_location = "Phòng học bộ môn / Lớp học"
        wed_morning_leader = f"{obs1.get('teacher', '')} (Dạy) & TTCM (Chủ trì dự)"
        wed_morning_note = "Ghi phiếu dự giờ theo mẫu quy định"
    else:
        wed_morning_content = (
            f"- Thực hiện giảng dạy chính khóa theo TKB các khối 6, 7, 8, 9.\n"
            f"- Dự giờ đồng nghiệp học hỏi kinh nghiệm (1-2 tiết/GV trong tổ).\n"
            f"- Tăng cường ứng dụng công nghệ thông tin, bài giảng số và khai thác thiết bị phòng bộ môn."
        )
        wed_morning_attendees = "GV trong tổ"
        wed_morning_location = "Các lớp học"
        wed_morning_leader = "Các tổ viên"
        wed_morning_note = "Đúng giờ, đúng phân phối"

    # THỨ NĂM: CHUYÊN ĐỀ STEM / DỰ GIỜ & SINH HOẠT CHUYÊN MÔN
    if stem:
        thu_content = (
            f"- Giảng dạy chính khóa theo TKB các lớp.\n"
            f"- Triển khai Chuyên đề STEM cấp tổ: '{stem.get('topic', '')}' - Môn {stem.get('subject', '')} do {stem.get('teacher', '')} chủ trì.\n"
            f"- Sinh hoạt chuyên môn tổ: {shcm_text}\n"
            f"- Kiểm tra, đánh giá việc tích hợp năng lực số (NLS), dinh dưỡng học đường và quốc phòng an ninh trong giáo án."
        )
        thu_attendees = "Toàn thể GV tổ & HS"
        thu_location = "Phòng STEM / Phòng họp CM"
        thu_leader = f"{stem.get('teacher', 'GV')} & TTCM"
        thu_note = "Đánh giá SP STEM, biên bản SHCM"
    elif len(observations) >= 2:
        obs2 = observations[1]
        thu_content = (
            f"- Giảng dạy chính khóa theo TKB các lớp.\n"
            f"- Tổ chức dự giờ chuyên môn: Môn {obs2.get('subject', '')} - Bài '{obs2.get('lesson', '')}' do {obs2.get('teacher', '')} giảng dạy.\n"
            f"- Sinh hoạt chuyên môn tổ: {shcm_text}\n"
            f"- Kiểm tra, đánh giá việc tích hợp năng lực số (NLS), dinh dưỡng học đường và quốc phòng an ninh trong giáo án."
        )
        thu_attendees = "Toàn thể GV tổ"
        thu_location = "Lớp học / Phòng họp CM"
        thu_leader = f"{obs2.get('teacher', '')} (Dạy) & TTCM (Chủ trì)"
        thu_note = "Ghi phiếu dự giờ, biên bản họp"
    else:
        thu_content = (
            f"- Giảng dạy theo thời khóa biểu các khối 6, 7, 8, 9.\n"
            f"- Sinh hoạt chuyên môn tổ: {shcm_text}\n"
            f"- Kiểm tra, đánh giá việc tích hợp năng lực số (NLS), dinh dưỡng học đường và quốc phòng an ninh trong giáo án.\n"
            f"- Phổ biến, quán triệt các văn bản chỉ đạo chuyên môn mới từ Ban Giám Hiệu."
        )
        thu_attendees = "Toàn thể GV tổ"
        thu_location = "Phòng họp Chuyên môn"
        thu_leader = "Tổ trưởng chuyên môn"
        thu_note = "Ghi biên bản họp đầy đủ"

    # THỨ SÁU: KIỂM TRA ĐỊNH KỲ / DỰ GIỜ & RÀ SOÁT TIẾN ĐỘ
    if exam_focus:
        fri_content = (
            f"- Giảng dạy theo TKB và tổ chức Kiểm tra định kỳ theo kế hoạch: {exam_focus}.\n"
            f"- Phân công giáo viên coi kiểm tra nghiêm túc, đúng quy chế.\n"
            f"- Bàn giao, kiểm kê thiết bị thí nghiệm và đồ dùng dạy học đã sử dụng trong tuần.\n"
            f"- Hoàn thiện hồ sơ sổ sách, sổ kế hoạch bài dạy, sổ điểm cá nhân."
        )
        fri_attendees = "GV bộ môn & HS"
        fri_location = "Lớp học / Phòng thi"
        fri_leader = "GV bộ môn & TTCM"
        fri_note = "Coi kiểm tra đúng quy chế"
    elif len(observations) >= 3:
        obs3 = observations[2]
        fri_content = (
            f"- Giảng dạy các tiết theo TKB.\n"
            f"- Tổ chức dự giờ chuyên môn: Môn {obs3.get('subject', '')} - Bài '{obs3.get('lesson', '')}' do {obs3.get('teacher', '')} giảng dạy.\n"
            f"- Bàn giao, kiểm kê thiết bị thí nghiệm và đồ dùng dạy học đã sử dụng trong tuần.\n"
            f"- Rà soát tiến độ chương trình tuần {week_num}, đối chiếu với phân phối chương trình của tổ."
        )
        fri_attendees = "GV trong tổ"
        fri_location = "Lớp học / Phòng thiết bị"
        fri_leader = f"{obs3.get('teacher', '')} (Dạy) & Tổ phó"
        fri_note = "Ghi phiếu dự giờ, kiểm kê TB"
    else:
        fri_content = (
            f"- Giảng dạy các tiết cuối tuần theo TKB.\n"
            f"- Rà soát tiến độ chương trình môn Toán, KHTN, Công nghệ, đối chiếu với phân phối chương trình của tổ.\n"
            f"- Bàn giao, kiểm kê thiết bị thí nghiệm và đồ dùng dạy học đã sử dụng trong tuần.\n"
            f"- Hoàn thiện hồ sơ sổ sách, sổ kế hoạch bài dạy, sổ điểm cá nhân."
        )
        fri_attendees = "GV bộ môn & CB Thiết bị"
        fri_location = "Lớp học / Phòng thiết bị"
        fri_leader = "GV bộ môn & Tổ phó"
        fri_note = "Báo cáo dạy bù nếu chậm tiết"

    # THỨ TƯ: GIẢNG DẠY, DỰ GIỜ & KIỂM TRA HỒ SƠ GIÁO ÁN DRIVE
    if observations and len(observations) >= 1:
        obs1 = observations[0]
        wed_content = (
            f"- Giảng dạy chính khóa theo TKB các khối 6, 7, 8, 9.\n"
            f"- Tổ chức dự giờ chuyên môn cấp tổ: Môn {obs1.get('subject', '')} - Bài '{obs1.get('lesson', '')}' do {obs1.get('teacher', '')} giảng dạy.\n"
            f"- Kiểm tra định kỳ kế hoạch bài dạy (giáo án) trên Google Drive theo Công văn 5512/BGDĐT và Thể thức Nghị định 30/2020/NĐ-CP.\n"
            f"- Rà soát việc tích hợp năng lực số (NLS), dinh dưỡng học đường và quốc phòng an ninh trong bài soạn."
        )
        wed_attendees = f"Tổ chuyên môn & GV {obs1.get('subject', '')}"
        wed_location = "Phòng học bộ môn / VP Tổ"
        wed_leader = f"{obs1.get('teacher', '')} (Dạy) & TTCM (Chủ trì dự)"
        wed_note = "Ghi phiếu dự giờ, đối chiếu giáo án"
    else:
        wed_content = (
            f"- Thực hiện giảng dạy chính khóa theo TKB các khối 6, 7, 8, 9.\n"
            f"- Dự giờ đồng nghiệp học hỏi kinh nghiệm (1-2 tiết/GV trong tổ).\n"
            f"- Kiểm tra định kỳ kế hoạch bài dạy (giáo án) trên Google Drive theo Công văn 5512/BGDĐT và Thể thức Nghị định 30/2020/NĐ-CP.\n"
            f"- Rà soát việc tích hợp năng lực số (NLS), dinh dưỡng học đường và quốc phòng an ninh trong bài soạn."
        )
        wed_attendees = "GV trong tổ"
        wed_location = "Các lớp học / VP Tổ"
        wed_leader = "Tổ trưởng & các GV"
        wed_note = "Ghi phiếu dự giờ, ký duyệt giáo án"

    events = [
        # THỨ HAI
        {
            "day": f"Thứ Hai\n({format_date_str(d2)})",
            "content": (
                f"- Chào cờ đầu tuần, sinh hoạt toàn trường.\n"
                f"- Thực hiện giảng dạy tuần {week_num} theo Thời khóa biểu nhà trường.\n"
                f"- Phân công nhiệm vụ chuyên môn và triển khai kế hoạch hoạt động tuần {week_num}.\n"
                f"- Hướng dẫn học sinh tự học, phụ đạo củng cố kiến thức cho học sinh chưa đạt về KQHT môn Toán, KHTN."
            ),
            "attendees": "Toàn thể GV tổ, HS",
            "location": "Sân trường / Phòng học",
            "leader": "Tổ trưởng & GVBM",
            "note": "Đúng giờ, trang phục chuẩn mực"
        },

        # THỨ BA
        {
            "day": f"Thứ Ba\n({format_date_str(d3)})",
            "content": (
                f"- Giảng dạy chính khóa theo TKB (Toán, KHTN, Công nghệ, HĐTN).\n"
                f"- Trọng tâm bài dạy tuần {week_num}: {lesson_focus}\n"
                f"- Bồi dưỡng Đội tuyển học sinh giỏi theo kế hoạch tổ:\n  + {hsg_text}\n"
                f"- Kiểm tra, cập nhật điểm số và nhận xét học sinh trên Cổng thông tin Giáo dục điện tử."
            ),
            "attendees": "GV Toán, KHTN, CN, HSG",
            "location": "Lớp học / Phòng bồi dưỡng HSG",
            "leader": "GV bồi dưỡng & GVBM",
            "note": "Ký sổ mượn TB, ghi chép HSG"
        },

        # THỨ TƯ
        {
            "day": f"Thứ Tư\n({format_date_str(d4)})",
            "content": wed_content,
            "attendees": wed_attendees,
            "location": wed_location,
            "leader": wed_leader,
            "note": wed_note
        },

        # THỨ NĂM
        {
            "day": f"Thứ Năm\n({format_date_str(d5)})",
            "content": thu_content,
            "attendees": thu_attendees,
            "location": thu_location,
            "leader": thu_leader,
            "note": thu_note
        },

        # THỨ SÁU
        {
            "day": f"Thứ Sáu\n({format_date_str(d6)})",
            "content": fri_content,
            "attendees": fri_attendees,
            "location": fri_location,
            "leader": fri_leader,
            "note": fri_note
        },

        # THỨ BẢY
        {
            "day": f"Thứ Bảy\n({format_date_str(d7)})",
            "content": (
                f"- Giảng dạy các tiết học theo TKB.\n"
                f"- Tham gia sinh hoạt lớp cuối tuần, đánh giá thi đua và nền nếp học tập của học sinh.\n"
                f"- Tổ trưởng tổng kết tình hình thực hiện công tác tuần {week_num}, lập dự kiến công tác tuần {week_num + 1} báo cáo BGH.\n"
                f"- Giáo viên hoàn thiện giáo án, kế hoạch bài dạy tuần tiếp theo đưa lên Google Drive đúng hạn quy định."
            ),
            "attendees": "GVCN, GVBM, HS",
            "location": "Lớp học / VP Tổ",
            "leader": "Tổ trưởng & GVCN",
            "note": "Gửi báo cáo trước 11h00, nộp giáo án đúng hạn"
        }
    ]

    return events


def build_focus_tasks(week_num, sec4_info, lesson_focus, hsg_focus, exam_focus):
    """
    Xây dựng danh sách công tác trọng tâm dựa trực tiếp trên Kế hoạch giáo dục tổ:
    Mục 4 (Sinh hoạt chuyên môn), Dự giờ, STEM, Bồi dưỡng HSG, Kiểm tra định kỳ.
    Đánh số liên tục 1, 2, 3, 4, 5, 6...
    """
    tasks = []
    idx = 1

    # 1. Nền nếp giảng dạy
    tasks.append(f"{idx}. Tiếp tục duy trì nghiêm túc nền nếp giảng dạy tuần {week_num} theo kế hoạch thời khóa biểu nhà trường.")
    idx += 1

    # 2. Sinh hoạt chuyên môn Mục 4
    shcm = sec4_info.get("shcm", "")
    if shcm:
        tasks.append(f"{idx}. Thực hiện nội dung Sinh hoạt chuyên môn: {shcm}")
        idx += 1

    # 3. Dự giờ chuyên môn
    obs = sec4_info.get("observations", [])
    if obs:
        obs_details = ", ".join([f"môn {o['subject']} - Bài '{o['lesson']}' ({o['teacher']})" for o in obs])
        tasks.append(f"{idx}. Tổ chức dự giờ chuyên môn theo kế hoạch tổ: {obs_details}.")
        idx += 1

    # 4. Chuyên đề STEM
    stem = sec4_info.get("stem", None)
    if stem:
        tasks.append(f"{idx}. Triển khai Chuyên đề STEM cấp tổ: '{stem.get('topic', '')}' - Môn {stem.get('subject', '')} do {stem.get('teacher', '')} phụ trách.")
        idx += 1

    # 5. Kiểm tra định kỳ hoặc Bồi dưỡng HSG
    if exam_focus:
        tasks.append(f"{idx}. Tổ chức Kiểm tra đánh giá định kỳ: {exam_focus}.")
        idx += 1
    elif hsg_focus:
        tasks.append(f"{idx}. Tăng cường bồi dưỡng học sinh giỏi theo kế hoạch: {hsg_focus}.")
        idx += 1

    # 6. Kiểm tra giáo án trên Drive
    tasks.append(f"{idx}. Kiểm tra hồ sơ chuyên môn, kế hoạch bài dạy của giáo viên trên Google Drive theo chuẩn Công văn 5512/BGDĐT và Thể thức Nghị định 30/2020/NĐ-CP.")
    idx += 1

    # 7. Thiết bị dạy học và phụ đạo
    tasks.append(f"{idx}. Tăng cường khai thác thiết bị dạy học, phòng thực hành KHTN và phụ đạo học sinh có học lực còn hạn chế.")

    return tasks


def generate_weekly_schedule_data(plan_path_or_text=None, week_num=5, school_name="TRƯỜNG TH & THCS PHƯỚC HƯNG", dept_name="TỔ TOÁN – KHTN – CÔNG NGHỆ", leader_name="Lê Văn Thắng", approver_name="Ban Giám Hiệu"):
    """
    Tạo dữ liệu JSON hoàn chỉnh cho Lịch công tác tuần, kết hợp đầy đủ:
    1. Mục 4. Sinh hoạt chuyên môn & NCBH & Sơ kết tổ
    2. Mục B. Kế hoạch chuyên đề STEM
    3. Mục 1. Kế hoạch bồi dưỡng học sinh giỏi
    4. Mục 3. Bài dạy trọng tâm & Kiểm tra đánh giá định kỳ
    """
    text_lines = []
    if plan_path_or_text and os.path.exists(plan_path_or_text):
        text_lines = extract_text_from_file(plan_path_or_text)
    elif isinstance(plan_path_or_text, list):
        text_lines = plan_path_or_text
    elif isinstance(plan_path_or_text, str) and plan_path_or_text.strip():
        text_lines = [l.strip() for l in plan_path_or_text.splitlines() if l.strip()]

    target_monday, target_saturday = compute_week_dates(week_num)
    start_str = target_monday.strftime("%d/%m/%Y")
    end_str = target_saturday.strftime("%d/%m/%Y")

    # Lấy thông tin đối soát chuẩn từ kế hoạch tổ
    sec4_info = SECTION_4_ACTIVITIES.get(week_num, {})
    lesson_focus = WEEKLY_LESSONS.get(week_num, f"Thực hiện giảng dạy tuần {week_num} theo phân phối chương trình các môn Toán, KHTN, Công nghệ, HĐTN.")
    hsg_focus = WEEKLY_HSG.get(week_num, "")
    exam_focus = WEEKLY_EXAMS.get(week_num, "")

    events = build_schedule_events(week_num, target_monday, sec4_info, lesson_focus, hsg_focus, exam_focus)
    focus_tasks = build_focus_tasks(week_num, sec4_info, lesson_focus, hsg_focus, exam_focus)

    return {
        "school_name": school_name,
        "dept_name": dept_name,
        "week_num": week_num,
        "date_range": f"Từ ngày {start_str} đến ngày {end_str}",
        "start_date": start_str,
        "end_date": end_str,
        "school_year": "2026 – 2027",
        "leader_name": leader_name,
        "approver_name": approver_name,
        "focus_tasks": focus_tasks,
        "events": events,
        "sec4_info": sec4_info,
        "lesson_focus": lesson_focus,
        "hsg_focus": hsg_focus,
        "exam_focus": exam_focus
    }


def export_schedule_to_word(schedule_data, output_path):
    """
    Xuất lịch công tác tuần ra file Word .docx đạt chuẩn Thể thức văn bản hành chính theo Nghị định 30/2020/NĐ-CP:
    - Bỏ mục 'I. Đánh giá công tác tuần trước' theo đúng yêu cầu người dùng.
    - Mục I: LỊCH CÔNG TÁC CHI TIẾT TRONG TUẦN (Thứ Hai đến Thứ Bảy).
    - Đổi chữ III thành II: Mục II: MỘT SỐ CÔNG TÁC TRỌNG TÂM CẦN LƯU Ý TRONG TUẦN.
    - 100% Font Times New Roman 13-14pt, chuẩn XML w:rFonts.
    - Căn lề chuẩn: Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm.
    """
    doc = docx.Document()

    # Căn lề chuẩn Nghị định 30
    for section in doc.sections:
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)
        section.page_width = Mm(210)
        section.page_height = Mm(297)

    # Style mặc định
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(3)

    # 1. HEADER (Bảng 2 cột: Bên trái Cơ quan/Tổ, Bên phải Quốc hiệu/Tiêu ngữ)
    t0 = doc.add_table(rows=2, cols=2)
    t0.alignment = WD_TABLE_ALIGNMENT.CENTER
    t0.autofit = False

    col_widths_0 = [Mm(75), Mm(90)]
    for row in t0.rows:
        for idx, width in enumerate(col_widths_0):
            row.cells[idx].width = width

    # Hàng 1
    c00 = t0.rows[0].cells[0].paragraphs[0]
    c00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = c00.add_run(schedule_data.get("school_name", "TRƯỜNG TH & THCS PHƯỚC HƯNG").upper())
    set_tnr(r, size_pt=11, bold=False)

    c01 = t0.rows[0].cells[1].paragraphs[0]
    c01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = c01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_tnr(r, size_pt=12, bold=True)

    # Hàng 2
    c10 = t0.rows[1].cells[0].paragraphs[0]
    c10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = c10.add_run(schedule_data.get("dept_name", "TỔ TOÁN – KHTN – CÔNG NGHỆ").upper())
    set_tnr(r, size_pt=12, bold=True)

    c11 = t0.rows[1].cells[1].paragraphs[0]
    c11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = c11.add_run("Độc lập - Tự do - Hạnh phúc")
    set_tnr(r, size_pt=13, bold=True)
    r.font.underline = True

    # 2. TIÊU ĐỀ CHÍNH
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(14)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run(f"LỊCH CÔNG TÁC TUẦN {schedule_data.get('week_num', '')}")
    set_tnr(r_title, size_pt=15, bold=True)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    r_sub = p_sub.add_run(f"({schedule_data.get('date_range', '')} - Năm học {schedule_data.get('school_year', '2026 – 2027')})")
    set_tnr(r_sub, size_pt=12, italic=True)

    # =========================================================================
    # 3. MỤC I: LỊCH CÔNG TÁC CHI TIẾT TRONG TUẦN (ĐÃ BỎ ĐÁNH GIÁ TUẦN TRƯỚC)
    # =========================================================================
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(8)
    p_sec1.paragraph_format.space_after = Pt(4)
    r_sec1 = p_sec1.add_run("I. LỊCH CÔNG TÁC CHI TIẾT TRONG TUẦN (THỨ HAI ĐẾN THỨ BẢY):")
    set_tnr(r_sec1, size_pt=13, bold=True)

    events = schedule_data.get("events", [])
    num_rows = len(events) + 1
    tbl = doc.add_table(rows=num_rows, cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    set_table_borders(tbl, color="A0AEC0", sz="4", val="single")

    # Kích thước 5 cột (Tổng cộng 165mm vừa vặn trang A4 căn lề chuẩn NĐ 30: 30-15)
    col_widths = [Mm(26), Mm(82), Mm(22), Mm(20), Mm(15)]
    for row in tbl.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    # Header Bảng (Đã bỏ cột Buổi theo yêu cầu)
    headers = ["Thứ / Ngày", "Nội dung công việc", "Thành phần", "Địa điểm & Phụ trách", "Ghi chú"]
    hdr_row = tbl.rows[0]
    for idx, text in enumerate(headers):
        cell = hdr_row.cells[idx]
        set_cell_background(cell, "EBF3FB")
        set_cell_margins(cell, top=120, bottom=120, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        set_tnr(r, size_pt=11, bold=True)

    # Điền các dòng sự kiện (5 cột)
    for i, ev in enumerate(events):
        row = tbl.rows[i + 1]
        
        # 1. Thứ / Ngày
        c0 = row.cells[0]
        set_cell_margins(c0, top=100, bottom=100, left=80, right=80)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(ev.get("day", ""))
        set_tnr(r0, size_pt=10.5, bold=True)

        # 2. Nội dung công việc
        c1 = row.cells[1]
        set_cell_margins(c1, top=100, bottom=100, left=100, right=100)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(ev.get("content", ""))
        set_tnr(r1, size_pt=11)

        # 3. Thành phần
        c2 = row.cells[2]
        set_cell_margins(c2, top=100, bottom=100, left=80, right=80)
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(ev.get("attendees", ""))
        set_tnr(r2, size_pt=10.5)

        # 4. Địa điểm & Phụ trách
        c3 = row.cells[3]
        set_cell_margins(c3, top=100, bottom=100, left=80, right=80)
        p3 = c3.paragraphs[0]
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lead_str = ev.get("leader", "")
        loc_str = ev.get("location", "")
        if lead_str and loc_str and lead_str != loc_str:
            combined = f"{loc_str}\n({lead_str})"
        elif lead_str:
            combined = lead_str
        else:
            combined = loc_str
        r3 = p3.add_run(combined)
        set_tnr(r3, size_pt=10.5)

        # 5. Ghi chú
        c4 = row.cells[4]
        set_cell_margins(c4, top=100, bottom=100, left=60, right=60)
        p4 = c4.paragraphs[0]
        p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r4 = p4.add_run(ev.get("note", ""))
        set_tnr(r4, size_pt=10, italic=True)

    # =========================================================================
    # 4. MỤC II: MỘT SỐ CÔNG TÁC TRỌNG TÂM CẦN LƯU Ý (CHỮ III ĐỔI THÀNH II)
    # =========================================================================
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(12)
    p_sec2.paragraph_format.space_after = Pt(4)
    r_sec2 = p_sec2.add_run("II. MỘT SỐ CÔNG TÁC TRỌNG TÂM CẦN LƯU Ý TRONG TUẦN:")
    set_tnr(r_sec2, size_pt=13, bold=True)

    for task in schedule_data.get("focus_tasks", []):
        p_t = doc.add_paragraph()
        p_t.paragraph_format.left_indent = Inches(0.2)
        p_t.paragraph_format.space_after = Pt(2)
        r_t = p_t.add_run(task)
        set_tnr(r_t, size_pt=12)

    # 5. PHẦN KÝ TÊN VÀ PHÊ DUYỆT (2 cột)
    t_sign = doc.add_table(rows=2, cols=2)
    t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_sign.autofit = False

    for row in t_sign.rows:
        for idx, width in enumerate(col_widths_0):
            row.cells[idx].width = width

    # Hàng ngày tháng & chức danh
    cs0 = t_sign.rows[0].cells[0].paragraphs[0]
    cs0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_bgh = cs0.add_run("BAN GIÁM HIỆU PHÊ DUYỆT\n")
    set_tnr(r_bgh, size_pt=12, bold=True)
    r_bgh_sub = cs0.add_run("(Ký và ghi rõ họ tên)")
    set_tnr(r_bgh_sub, size_pt=11, italic=True)

    cs1 = t_sign.rows[0].cells[1].paragraphs[0]
    cs1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_date = cs1.add_run(f"Phước Hưng, ngày {schedule_data.get('start_date', '')}\n")
    set_tnr(r_date, size_pt=12, italic=True)
    r_tt = cs1.add_run("TỔ TRƯỞNG CHUYÊN MÔN\n")
    set_tnr(r_tt, size_pt=12, bold=True)
    r_tt_sub = cs1.add_run("(Ký và ghi rõ họ tên)")
    set_tnr(r_tt_sub, size_pt=11, italic=True)

    # Khoảng trống ký tên
    cs0_name = t_sign.rows[1].cells[0].paragraphs[0]
    cs0_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cs0_name.paragraph_format.space_before = Pt(50)
    r_app = cs0_name.add_run(schedule_data.get("approver_name", "Ban Giám Hiệu"))
    set_tnr(r_app, size_pt=12, bold=True)

    cs1_name = t_sign.rows[1].cells[1].paragraphs[0]
    cs1_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cs1_name.paragraph_format.space_before = Pt(50)
    r_lead = cs1_name.add_run(schedule_data.get("leader_name", "Lê Văn Thắng"))
    set_tnr(r_lead, size_pt=12, bold=True)

    doc.save(output_path)
    return output_path
