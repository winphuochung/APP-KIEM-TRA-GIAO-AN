import os
import docx
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders_custom(table):
    tblPr = table._tbl.tblPr
    # Viền ngoài đôi (double) như trong ảnh, viền trong đơn (single)
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="double" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="double" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:left w:val="double" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:right w:val="double" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def generate_weekly_plan_exact_template(output_path, include_signature=True):
    doc = docx.Document()

    # Căn lề chuẩn A4
    for section in doc.sections:
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(25)
        section.right_margin = Mm(20)
        section.page_width = Mm(210)
        section.page_height = Mm(297)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(2)

    # 1. BẢNG HEADER 2 CỘT
    t0 = doc.add_table(rows=3, cols=2)
    t0.alignment = WD_TABLE_ALIGNMENT.CENTER
    t0.autofit = False

    col_widths_0 = [Mm(80), Mm(85)]
    for row in t0.rows:
        for idx, width in enumerate(col_widths_0):
            row.cells[idx].width = width

    # Hàng 1
    p00 = t0.rows[0].cells[0].paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r00 = p00.add_run("TRƯỜNG TH & THCS PHƯỚC HƯNG")
    r00.font.name = "Times New Roman"
    r00.font.size = Pt(12)

    p01 = t0.rows[0].cells[1].paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r01 = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    r01.font.name = "Times New Roman"
    r01.font.size = Pt(12)
    r01.font.bold = True

    # Hàng 2
    p10 = t0.rows[1].cells[0].paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r10 = p10.add_run("TỔ TOÁN – KHTN - CN")
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
    r20 = p20.add_run("Số: ...../KH-TTN")
    r20.font.name = "Times New Roman"
    r20.font.size = Pt(11)

    p21 = t0.rows[2].cells[1].paragraphs[0]
    p21.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r21 = p21.add_run("Nhơn Hội, ngày 27 tháng 09  năm 2025")
    r21.font.name = "Times New Roman"
    r21.font.size = Pt(11)
    r21.font.italic = True

    # 2. TIÊU ĐỀ KẾ HOẠCH
    p_t1 = doc.add_paragraph()
    p_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t1.paragraph_format.space_before = Pt(12)
    p_t1.paragraph_format.space_after = Pt(2)
    r_t1 = p_t1.add_run("KẾ HOẠCH HOẠT ĐỘNG TUẦN 04")
    r_t1.font.name = "Times New Roman"
    r_t1.font.size = Pt(14)
    r_t1.font.bold = True

    p_t2 = doc.add_paragraph()
    p_t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t2.paragraph_format.space_after = Pt(4)
    r_t2 = p_t2.add_run("NĂM HỌC: 2026 - 2027")
    r_t2.font.name = "Times New Roman"
    r_t2.font.size = Pt(13)
    r_t2.font.bold = True

    # Icon / biểu tượng sách mở trang trí như ảnh: 🙞 📖 🙜
    p_icon = doc.add_paragraph()
    p_icon.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_icon.paragraph_format.space_after = Pt(12)
    r_icon = p_icon.add_run("❦  🕮  ❧")
    r_icon.font.name = "Times New Roman"
    r_icon.font.size = Pt(13)

    # 3. BẢNG KẾ HOẠCH HOẠT ĐỘNG (3 CỘT)
    table_data = [
        {
            "time": "29/09/2025\nThứ hai",
            "content": "- Chào cờ đầu tuần.\n- Nộp Báo cáo sơ kết tháng 09",
            "person": "- GVCN\n- TT→PHT"
        },
        {
            "time": "30/09/2025\nThứ ba",
            "content": "- Duyệt kế hoạch cá nhân và kế hoạch tổ bộ môn,\nKế hoạch chủ nhiệm\n- Nhập danh sách dự thi HSG (Sinh, Hóa)",
            "person": "- GVBM→TT\n\n- GVBM"
        },
        {
            "time": "01/10/2025\nThứ tư",
            "content": "",
            "person": ""
        },
        {
            "time": "02/10/2025\nThứ năm",
            "content": "- Nộp Kế hoạch bài dạy tuần 5 đến tuần 8",
            "person": "- GVBM→TT"
        },
        {
            "time": "03/10/2025\nThứ sáu",
            "content": "- Họp sơ kết tổ tháng 09 lúc 7h30",
            "person": "- Cả tổ"
        },
        {
            "time": "04/10/2025\nThứ bảy",
            "content": "- Hội nghị công nhân viên chức đầu năm lúc 13h30",
            "person": "- Cả tổ"
        }
    ]

    t_main = doc.add_table(rows=len(table_data) + 1, cols=3)
    t_main.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_main.autofit = False
    set_table_borders_custom(t_main)

    col_widths = [Mm(32), Mm(92), Mm(41)]
    for row in t_main.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    # Header hàng
    headers = ["Thời gian", "Nội dung hoạt động", "Người thực hiện"]
    hdr_row = t_main.rows[0]
    for idx, title in enumerate(headers):
        cell = hdr_row.cells[idx]
        set_cell_margins(cell, top=120, bottom=120, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.font.bold = True

    # Các hàng dữ liệu
    for i, item in enumerate(table_data):
        row = t_main.rows[i + 1]
        
        # Cột 1: Thời gian (canh giữa)
        c0 = row.cells[0]
        set_cell_margins(c0, top=100, bottom=100, left=60, right=60)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for line in item["time"].split("\n"):
            r = p0.add_run(line + "\n")
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
        # Bỏ dấu xuống dòng thừa cuối
        if p0.runs:
            p0.runs[-1].text = p0.runs[-1].text.rstrip("\n")

        # Cột 2: Nội dung hoạt động (canh trái)
        c1 = row.cells[1]
        set_cell_margins(c1, top=100, bottom=100, left=80, right=80)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for line in item["content"].split("\n"):
            r = p1.add_run(line + "\n")
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
        if p1.runs:
            p1.runs[-1].text = p1.runs[-1].text.rstrip("\n")

        # Cột 3: Người thực hiện (canh trái)
        c2 = row.cells[2]
        set_cell_margins(c2, top=100, bottom=100, left=80, right=80)
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for line in item["person"].split("\n"):
            r = p2.add_run(line + "\n")
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
        if p2.runs:
            p2.runs[-1].text = p2.runs[-1].text.rstrip("\n")

    # 4. PHẦN DƯỚI BẢNG: LƯU Ý & CHỮ KÝ TỔ TRƯỞNG
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(8)
    p_space.paragraph_format.space_after = Pt(2)

    t_bottom = doc.add_table(rows=1, cols=2)
    t_bottom.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_bottom.autofit = False

    t_bottom.rows[0].cells[0].width = Mm(85)
    t_bottom.rows[0].cells[1].width = Mm(80)

    # Ô bên trái: Lưu ý
    c_left = t_bottom.rows[0].cells[0]
    p_luuy = c_left.paragraphs[0]
    r_luuy = p_luuy.add_run("Lưu ý:")
    r_luuy.font.name = "Times New Roman"
    r_luuy.font.size = Pt(12)
    r_luuy.font.bold = True

    # Ô bên phải: Chữ ký Tổ trưởng
    c_right = t_bottom.rows[0].cells[1]
    p_sign = c_right.paragraphs[0]
    p_sign.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sign_title = p_sign.add_run("Tổ trưởng\n")
    r_sign_title.font.name = "Times New Roman"
    r_sign_title.font.size = Pt(13)
    r_sign_title.font.bold = True

    from app.config import get_file_path
    sig_img_path = get_file_path("signature_thang.png")
    if not os.path.exists(sig_img_path):
        sig_img_path = get_file_path(os.path.join("data", "signature_thang.png"))

    if include_signature and os.path.exists(sig_img_path):
        p_sig_img = c_right.add_paragraph()
        p_sig_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sig_img.paragraph_format.space_before = Pt(0)
        p_sig_img.paragraph_format.space_after = Pt(2)
        r_img = p_sig_img.add_run()
        r_img.add_picture(sig_img_path, width=Inches(1.5))
    else:
        # Nếu không chèn ảnh, để khoảng trắng ký tay
        p_sign.paragraph_format.space_after = Pt(45)

    p_name = c_right.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(2)
    r_name = p_name.add_run("Lê Văn Thắng")
    r_name.font.name = "Times New Roman"
    r_name.font.size = Pt(13)
    r_name.font.bold = True

    doc.save(output_path)
    return output_path
