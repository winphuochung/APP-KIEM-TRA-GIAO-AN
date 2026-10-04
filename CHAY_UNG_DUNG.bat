@echo off
chcp 65001 > nul
echo ========================================================
echo   HE SINH THAI KIEM DINH GIAO DUC SO TOAN DIEN
echo   Dang khoi chay may chu ung dung...
echo ========================================================
cd /d "D:\APP-KIEM-TRA-GIAO-AN"
start http://127.0.0.1:8000
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000 --reload
pause
