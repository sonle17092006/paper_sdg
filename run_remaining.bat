@echo off
chcp 65001 > nul
echo ======================================================================
echo   CHAY TIEP CAC BUOC CON LAI: SBERT + SENTIMENT + BIEU DO + NOTEBOOK
echo   (Da co san 95.804 cau tu 42 bao cao PDF, khong can extract lai)
echo ======================================================================
python run_all.py --skip-models --skip-extract %*
pause
