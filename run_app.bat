@echo off
chcp 65001 > nul
echo =========================================================================
echo    HỆ SINH THÁI KIỂM ĐỊNH GIÁO DỤC SỐ TOÀN DIỆN - AI SENIOR INSPECTOR
echo =========================================================================
echo Đang khởi động Web App máy chủ tại http://127.0.0.1:8000 ...
start "" http://127.0.0.1:8000
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000 --reload
pause
