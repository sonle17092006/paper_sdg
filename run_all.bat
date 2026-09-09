@echo off
chcp 65001 > nul
echo ======================================================================
echo   KHỞI ĐỘNG CHẠY TOÀN BỘ PIPELINE PHÂN TÍCH SDG ^& TẨY XANH
echo ======================================================================
python run_all.py %*
pause
