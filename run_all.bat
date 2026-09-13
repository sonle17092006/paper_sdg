@echo off
chcp 65001 > nul
echo ======================================================================
echo   KHOI DONG PIPELINE PHAN TICH SDG ^& TAY XANH (7 CONG TY: 2020-2025)
echo ======================================================================
python run_all.py %*
pause
