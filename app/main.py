import os
import re
import shutil
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from app.services.inspector import inspect_lesson_plan
from app.services.corrector import correct_document
from app.services.report_generator import generate_inspection_report
from app.services.competency import TEACHERS_DATA, get_heatmap_matrix, get_bi_overview
from app.services.schedule_generator import (
    generate_weekly_schedule_data,
    export_schedule_to_word,
    extract_text_from_file,
    parse_plan_to_weeks
)

app = FastAPI(title="Hệ sinh thái Kiểm định Giáo dục Số Toàn diện", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = r"D:\APP-KIEM-TRA-GIAO-AN\data\uploads"
CORRECTED_DIR = r"D:\APP-KIEM-TRA-GIAO-AN\data\corrected"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(CORRECTED_DIR, exist_ok=True)

@app.get("/api/teachers")
def list_teachers():
    return {"teachers": TEACHERS_DATA}

@app.get("/api/heatmap")
def heatmap_data():
    return get_heatmap_matrix()

@app.get("/api/bi-metrics")
def bi_overview():
    return get_bi_overview()

@app.get("/api/competency/{teacher_id}")
def teacher_competency(teacher_id: str):
    for t in TEACHERS_DATA:
        if t["id"] == teacher_id:
            return t
    raise HTTPException(status_code=404, detail="Không tìm thấy giáo viên")

@app.post("/api/inspect-sample")
def inspect_sample_file(sample_index: int = Form(0), teacher_name: str = Form("Phạm Thị Cúc")):
    sample_files = [
        r"D:\APP-KIEM-TRA-GIAO-AN\data\Co_Cuc\Cô Cúc\KHTN 7\KHTN 7 tuần 1-4\Tuần 1-2 Tiết 1-5  BÀI 1 MỞ ĐẦU KHTN.docx",
        r"D:\APP-KIEM-TRA-GIAO-AN\data\Co_Cuc\Cô Cúc\KHTN 7\KHTN 7 tuần 1-4\Tuần 2-3  Tiết 6-9  Bài 2  Nguyên tử - KHTN7 - CTST.docx",
        r"D:\APP-KIEM-TRA-GIAO-AN\data\Co_Cuc\Cô Cúc\KHTN 7\KHTN 7 tuần 1-4\Tuần 3-4, Tiết 10-13  -CHỦ ĐỀ 1 - BÀI 3- NTHH.docx",
        r"D:\APP-KIEM-TRA-GIAO-AN\data\Co_Cuc\Cô Cúc\KHTN 7\KHTN 7 tuần 1-4\Tuần 4-5  -Tiết 14-18  Bài 4- SƠ LƯỢC BẢNG TUẦN TOÀN CÁC NTHH-KHTN 7-CTST-ST.docx",
        r"D:\APP-KIEM-TRA-GIAO-AN\data\Co_Cuc\Cô Cúc\KHTN 9\KHTN 9 Tuần 1-4\Tuần 2 - Tiết 1-2   Bài 36-Khái quát về di truyền học-Sinh9- KNTT.docx",
        r"D:\APP-KIEM-TRA-GIAO-AN\data\Co_Cuc\Cô Cúc\KHTN 9\KHTN 9 Tuần 1-4\Tuần 3,4,5,6 - Tiết 3,4,5,6  BÀI 37- khtn 9- kntt.docx"
    ]
    if sample_index < 0 or sample_index >= len(sample_files):
        sample_index = 0
    fpath = sample_files[sample_index]
    result = inspect_lesson_plan(fpath)
    result["file_path"] = fpath
    result["teacher_name"] = teacher_name
    return result

@app.post("/api/inspect-upload")
async def inspect_uploaded_file(file: UploadFile = File(...), teacher_name: str = Form("Phạm Thị Cúc")):
    target_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(target_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    result = inspect_lesson_plan(target_path)
    result["file_path"] = target_path
    result["teacher_name"] = teacher_name
    return result

@app.post("/api/auto-correct")
def auto_correct_file(file_path: str = Form(...)):
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Không tìm thấy tệp nguồn")
    fname = os.path.basename(file_path)
    out_name = fname.replace(".docx", "_ChuanTheThuc_ND30.docx")
    out_path = os.path.join(CORRECTED_DIR, out_name)
    result = correct_document(file_path, out_path)
    return {
        "status": "success",
        "download_url": f"/api/download-file?path={out_path}",
        "filename": out_name,
        "details": result
    }

@app.get("/api/download-file")
def download_file(path: str):
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Tệp không tồn tại")
    return FileResponse(path, filename=os.path.basename(path), media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.get("/api/export-report")
def export_report_endpoint(teacher: str = "Phạm Thị Cúc", inspector: str = "Lê Văn Thắng", period: str = "Tuần 1 đến tuần 4"):
    out_path = r"D:\APP-KIEM-TRA-GIAO-AN\BIEN_BAN_KIEM_TRA_GIAO_AN.docx"
    generate_inspection_report(teacher_name=teacher, inspector_name=inspector, period_str=period, output_path=out_path)
    return FileResponse(out_path, filename=os.path.basename(out_path), media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# =========================================================================
# TÍNH NĂNG: LỊCH CÔNG TÁC TUẦN TỪ KẾ HOẠCH HOẠT ĐỘNG CỦA TỔ CHUYÊN MÔN
# =========================================================================

@app.get("/api/schedule/sample-plan")
def get_sample_plan_schedule(week: int = 5):
    """Sử dụng file kế hoạch hoạt động tổ có sẵn trong hệ thống"""
    sample_pdf = r"D:\APP-KIEM-TRA-GIAO-AN\2026-2027 Kế hoạch GD tổ Toán-KHTN-C.Nghệ.pdf"
    if not os.path.exists(sample_pdf):
        # Thử tìm file pdf bất kỳ trong thư mục
        for f in os.listdir(r"D:\APP-KIEM-TRA-GIAO-AN"):
            if f.endswith(".pdf") and "K" in f:
                sample_pdf = os.path.join(r"D:\APP-KIEM-TRA-GIAO-AN", f)
                break

    data = generate_weekly_schedule_data(
        plan_path_or_text=sample_pdf if os.path.exists(sample_pdf) else None,
        week_num=week,
        school_name="TRƯỜNG TH & THCS PHƯỚC HƯNG",
        dept_name="TỔ TOÁN – KHTN – CÔNG NGHỆ",
        leader_name="Lê Văn Thắng",
        approver_name="Ban Giám Hiệu"
    )
    data["source_file"] = os.path.basename(sample_pdf) if os.path.exists(sample_pdf) else "Kế hoạch chuyên môn tổ mặc định"
    return data

@app.post("/api/schedule/upload-plan")
async def upload_activity_plan(
    file: UploadFile = File(...),
    week: int = Form(5),
    school_name: str = Form("TRƯỜNG TH & THCS PHƯỚC HƯNG"),
    dept_name: str = Form("TỔ TOÁN – KHTN – CÔNG NGHỆ"),
    leader_name: str = Form("Lê Văn Thắng")
):
    """Nhận file kế hoạch của tổ (.docx, .pdf, .xlsx, .txt), trích xuất và tạo lịch công tác tuần"""
    target_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(target_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    data = generate_weekly_schedule_data(
        plan_path_or_text=target_path,
        week_num=week,
        school_name=school_name,
        dept_name=dept_name,
        leader_name=leader_name,
        approver_name="Ban Giám Hiệu"
    )
    data["source_file"] = file.filename
    data["file_path"] = target_path
    return data

@app.post("/api/schedule/export-docx")
async def export_schedule_docx(request: Request):
    """Xuất lịch công tác tuần ra file Word .docx chuẩn Thể thức Nghị định 30/2020/NĐ-CP"""
    payload = await request.json()
    week_num = payload.get("week_num", 1)
    dept_slug = re.sub(r'[^a-zA-Z0-9]', '_', str(payload.get("dept_name", "TO_CHUYEN_MON")))[:20]
    out_filename = f"LICH_CONG_TAC_TUAN_{week_num}_{dept_slug}.docx"
    out_path = os.path.join(CORRECTED_DIR, out_filename)
    
    export_schedule_to_word(payload, out_path)
    
    return FileResponse(
        out_path,
        filename=out_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/api/schedule/download-template-docx")
def download_template_docx():
    """Tải file Word Kế hoạch tuần mẫu chuẩn của Tổ chuyên môn"""
    fpath = r"D:\APP-KIEM-TRA-GIAO-AN\KE_HOACH_HOAT_DONG_TUAN_04_TO_TOAN_KHTN_CN.docx"
    if not os.path.exists(fpath):
        from app.services.exact_template_generator import generate_weekly_plan_exact_template
        generate_weekly_plan_exact_template(fpath, include_signature=True)
    return FileResponse(
        fpath,
        filename="KE_HOACH_HOAT_DONG_TUAN_04_TO_TOAN_KHTN_CN.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# =========================================================================
# TÍNH NĂNG: GIÁM SÁT GOOGLE DRIVE VÀ THÔNG BÁO GMAIL
# =========================================================================
from app.services.drive_monitor import (
    scan_google_drive_status,
    send_notification_to_email,
    export_drive_monitoring_report_word
)

@app.get("/api/drive/status")
def get_drive_monitoring_status(refresh: bool = False):
    """Quét và trả về tình trạng cập nhật giáo án trên Google Drive của 8 GV"""
    log_path = r"D:\APP-KIEM-TRA-GIAO-AN\data\drive_monitoring_log.json"
    if not refresh and os.path.exists(log_path):
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    try:
        return scan_google_drive_status()
    except Exception as e:
        if os.path.exists(log_path):
            with open(log_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {"error": str(e), "teachers": []}

@app.post("/api/drive/notify-email")
def trigger_email_notification(recipient: str = "thangphuochung1@gmail.com"):
    """Ghi nhận và gửi email thông báo tiến độ nộp giáo án tới Gmail"""
    result = send_notification_to_email(recipient=recipient)
    return result

@app.get("/api/drive/export-report-docx")
def download_drive_report_docx():
    """Xuất Báo cáo giám sát tiến độ giáo án Google Drive ra file Word .docx chuẩn NĐ 30"""
    out_docx = export_drive_monitoring_report_word()
    return FileResponse(
        out_docx,
        filename="BAO_CAO_GIAM_SAT_TIEN_DO_GIAO_AN_GOOGLE_DRIVE.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/api/drive/detailed-matrix")
def get_detailed_matrix():
    """Lấy dữ liệu chi tiết 8 giáo viên qua 9 chu kỳ tuần (Tuần 1 đến tuần 35)"""
    mat_path = r"D:\APP-KIEM-TRA-GIAO-AN\data\drive_full_cycles_report.json"
    try:
        if os.path.exists(mat_path):
            with open(mat_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}
    except Exception as e:
        return {"error": str(e)}

@app.get("/api/drive/export-9cycles-docx")
def download_9cycles_docx():
    """Xuất Báo cáo chi tiết 9 chu kỳ tuần ra file Word .docx chuẩn NĐ 30"""
    from app.services.cycle_report_generator import generate_9_cycles_monitoring_word
    fpath = r"D:\APP-KIEM-TRA-GIAO-AN\BAO_CAO_CHI_TIET_TIEN_DO_9_CHU_KY_TUAN_GOOGLE_DRIVE.docx"
    if not os.path.exists(fpath):
        generate_9_cycles_monitoring_word(fpath)
    return FileResponse(
        fpath,
        filename="BAO_CAO_CHI_TIET_TIEN_DO_9_CHU_KY_TUAN_GOOGLE_DRIVE.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/api/inspection/data")
def get_inspection_data(teacher_id: str = "thanh", period: str = "Tuần 1 đến tuần 4", subject: str = "all"):
    """Lấy dữ liệu kiểm tra giáo án đối chiếu PPCT và SGK gộp tất cả các môn của giáo viên"""
    from app.services.inspection_engine import get_teacher_inspected_data
    data = get_teacher_inspected_data(teacher_id, period)
    if not data:
        raise HTTPException(status_code=404, detail="Không tìm thấy giáo viên")
    if subject and subject != "all":
        data["lessons"] = [l for l in data["lessons"] if l["subject"] == subject]
    return data

@app.get("/api/inspection/download-docx")
def download_teacher_inspection_docx(teacher_id: str = "thanh", period: str = "Tuần 1 đến tuần 4"):
    """Xuất file Word Biên bản kiểm tra giáo án gộp tất cả các môn của giáo viên theo giai đoạn tuần chuẩn NĐ 30"""
    from app.services.inspection_engine import export_teacher_inspection_report_word
    out_docx = export_teacher_inspection_report_word(teacher_id, period)
    fname = os.path.basename(out_docx)
    return FileResponse(
        out_docx,
        filename=fname,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/api/inspection/drive-files")
def get_inspection_drive_files(teacher_id: str = "thanh", period: str = "Tuần 1 đến tuần 4", subject: str = "all"):
    """Lấy danh sách các tệp giáo án thực tế trên Google Drive theo giáo viên, chu kỳ và môn học"""
    from app.services.inspection_engine import TEACHERS_INFO, LIVE_REPORT_PATH
    if not os.path.exists(LIVE_REPORT_PATH):
        return {"files": []}
    with open(LIVE_REPORT_PATH, "r", encoding="utf-8") as f:
        live = json.load(f)
    t_data = live.get(teacher_id, {})
    cycle = t_data.get("cycles", {}).get(period, {})
    subjs_detail = cycle.get("subjects_detail", {})
    files = []
    if subject == "all":
        for sname, sinfo in subjs_detail.items():
            for fn in sinfo.get("files", []):
                files.append({"subject": sname, "name": fn, "display": f"[{sname}] {os.path.basename(fn)}"})
    else:
        sinfo = subjs_detail.get(subject, {})
        for fn in sinfo.get("files", []):
            files.append({"subject": subject, "name": fn, "display": os.path.basename(fn)})
    return {"files": files, "teacher": TEACHERS_INFO.get(teacher_id)}

# =========================================================================
# TÍNH NĂNG: SỔ TAY TỔ TRƯỞNG & BÁO CÁO SƠ KẾT THÁNG TỔ CHUYÊN MÔN
# =========================================================================
from app.services.notebook_monthly_report import (
    get_notebook_week,
    save_notebook_week,
    sync_notebook_week_from_drive,
    generate_monthly_report_data,
    export_monthly_report_word
)

@app.get("/api/notebook/week")
def api_get_notebook_week(week: int = 1):
    """Lấy danh sách ghi chép các giáo viên trong tuần đó từ Sổ tay tổ trưởng"""
    return get_notebook_week(week)

@app.post("/api/notebook/week")
async def api_save_notebook_week(request: Request):
    """Lưu cập nhật sổ tay ghi chép tuần của tổ trưởng"""
    payload = await request.json()
    week_num = payload.get("week", 1)
    records = payload.get("records", [])
    return save_notebook_week(week_num, records)

@app.post("/api/notebook/sync-drive")
def api_sync_notebook_drive(week: int = 1):
    """Tự động đồng bộ tình trạng giáo án trên Drive vào Sổ tay tuần"""
    return sync_notebook_week_from_drive(week)

@app.get("/api/monthly-report/data")
def api_get_monthly_report_data(month: int = 9):
    """Lấy dữ liệu Báo cáo sơ kết tháng được tổng hợp trực tiếp từ Sổ tay các tuần"""
    return generate_monthly_report_data(month)

@app.get("/api/monthly-report/export-docx")
def api_export_monthly_report_docx(month: int = 9):
    """Xuất file Word Báo cáo sơ kết hoạt động tổ chuyên môn theo tháng chuẩn mẫu .doc cũ và NĐ 30"""
    report_data = generate_monthly_report_data(month)
    out_docx = export_monthly_report_word(report_data)
    fname = f"BAO_CAO_SO_KET_HDCM_THANG_{month:02d}.docx"
    return FileResponse(
        out_docx,
        filename=fname,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# =========================================================================
# TÍNH NĂNG: KẾT NỐI VÀ ĐỒNG BỘ CƠ SỞ DỮ LIỆU SUPABASE CLOUD
# =========================================================================
from app.services.supabase_client import get_supabase_status, sync_all_to_supabase

@app.get("/api/supabase/status")
def api_supabase_status():
    """Lấy trạng thái kết nối và các bảng dữ liệu Supabase Cloud"""
    return get_supabase_status()

@app.post("/api/supabase/sync-all")
def api_supabase_sync_all():
    """Thực hiện đồng bộ 2 chiều giữa dữ liệu Local và Supabase Cloud"""
    return sync_all_to_supabase()

@app.get("/api/supabase/schema-sql")
def api_supabase_schema_sql():
    """Lấy câu lệnh SQL để khởi tạo bảng dữ liệu trên Supabase SQL Editor"""
    sql_file = r"D:\APP-KIEM-TRA-GIAO-AN\supabase_schema.sql"
    if os.path.exists(sql_file):
        with open(sql_file, "r", encoding="utf-8") as f:
            return {"sql": f.read()}
    return {"sql": "-- File schema.sql không tồn tại"}

# Phục vụ giao diện tĩnh
static_path = os.path.join(os.path.dirname(__file__), "static")
app.mount("/", StaticFiles(directory=static_path, html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="127.0.0.1", port=8000, reload=True)
