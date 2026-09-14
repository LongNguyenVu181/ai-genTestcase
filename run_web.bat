@echo off
title Running Streamlit App...
cd /d "%~dp0"

echo ===================================================
echo   KICH HOAT MOI TRUONG AO VENV
echo ===================================================
call .venv\Scripts\activate.bat

echo.
echo ===================================================
echo   DANG KIEM TRA VA CAI DAT THU VIEN...
echo ===================================================
pip install streamlit openai pypdf python-dotenv pandas openpyxl

echo.
echo ===================================================
echo   DANG KHOI CHAY STREAMLIT APP...
echo ===================================================
streamlit run multiAgent.py

pause